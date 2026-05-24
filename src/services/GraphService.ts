import { SPHttpClient, SPHttpClientResponse, MSGraphClientV3 } from '@microsoft/sp-http';

export interface IGraphUser {
  id: string;               // UPN (from AccountName) — matches what Manager field stores
  displayName: string;
  mail: string;
  jobTitle: string;
  mobilePhone: string;
  businessPhones: string[];
  department: string;
  officeLocation: string;
  userPrincipalName: string;
  accountEnabled?: boolean;  // false = disabled / blocked sign-in
  userType?: string;         // 'Member' | 'Guest' | undefined
}

export type PresenceAvailability =
  'Available' | 'Busy' | 'DoNotDisturb' | 'BeRightBack' | 'Away' | 'Offline' | 'Unknown';

export interface IOrgNode {
  user: IGraphUser;
  directReports: IOrgNode[];
  isExpanded: boolean;
  childrenLoaded: boolean;
  level: number;
}

const PEOPLE_SOURCE = 'b09a7990-05ea-4af9-81ef-edfab16c4e31';
const SELECT_PROPS   = 'AccountName,DisplayName,PreferredName,JobTitle,Department,' +
                       'WorkEmail,WorkPhone,MobilePhone,OfficeNumber,PictureURL,Manager';
const BATCH_SIZE = 500;

export class GraphService {
  private _client:         SPHttpClient;
  private _graphClient:    MSGraphClientV3 | undefined;
  private _webUrl:         string;
  private _photoCache:     Map<string, string | null> = new Map();
  private _allUsersCache:  IGraphUser[] | null = null;
  private _childrenMap:    Map<string, IGraphUser[]> = new Map();
  private _managerMap:        Map<string, string> = new Map();   // userId → managerId (resolved)
  private _pendingManagerIds: Map<string, string> = new Map();   // userId → raw managerId (pre-resolution)
  private _loadingPromise: Promise<IGraphUser[]> | null = null;
  private _upnToObjectId:   Map<string, string> = new Map();     // upn → AAD object id
  private _presenceCache:   Map<string, PresenceAvailability> = new Map();
  private _presenceExpiry = 0;
  private readonly _PRESENCE_TTL = 60_000;

  constructor(client: SPHttpClient, webUrl: string, graphClient?: MSGraphClientV3) {
    this._client      = client;
    this._webUrl      = webUrl.replace(/\/$/, '');
    this._graphClient = graphClient;
  }

  /* ── Public API ──────────────────────────────────────────────────── */

  public getAllUsers(): Promise<IGraphUser[]> {
    if (this._allUsersCache) return Promise.resolve(this._allUsersCache);
    if (this._loadingPromise) return this._loadingPromise;

    this._loadingPromise = this._fetchAllUsers()
      .then(users => this._supplementUnlicensedReports(users))
      .then(users => {
        this._allUsersCache = users;
        this._buildMaps(users);
        this._loadingPromise = null;
        return users;
      })
      .catch(err => {
        this._loadingPromise = null;
        throw err;
      });

    return this._loadingPromise;
  }

  public async getUserPhoto(userId: string): Promise<string | null> {
    const key = userId.toLowerCase();
    if (this._photoCache.has(key)) return this._photoCache.get(key) ?? null;
    await this.getAllUsers(); // photo cache is populated during user load
    return this._photoCache.get(key) ?? null;
  }

  public async getDirectReports(userId: string): Promise<IGraphUser[]> {
    if (!this._allUsersCache) await this.getAllUsers();
    return this._childrenMap.get(userId.toLowerCase()) || [];
  }

  public async hasDirectReports(userId: string): Promise<boolean> {
    if (!this._allUsersCache) await this.getAllUsers();
    const kids = this._childrenMap.get(userId.toLowerCase());
    return !!(kids && kids.length > 0);
  }

  public async getManagerChain(userId: string, levels: number): Promise<IGraphUser[]> {
    const allUsers = await this.getAllUsers();
    const byId: { [id: string]: IGraphUser } = {};
    allUsers.forEach(u => { byId[u.id] = u; });

    const chain: IGraphUser[] = [];
    let curId = userId.toLowerCase();
    for (let i = 0; i < levels; i++) {
      const mgrId = this._managerMap.get(curId);
      if (!mgrId) break;
      const mgrUser = byId[mgrId];
      if (!mgrUser) break;
      chain.unshift(mgrUser);
      curId = mgrId;
    }
    return chain;
  }

  public async findUser(identifier: string): Promise<IGraphUser | null> {
    if (!identifier) return null;
    const users = await this.getAllUsers();
    const q = identifier.toLowerCase();
    return (
      users.find(u => u.id === q) ||
      users.find(u => (u.mail || '').toLowerCase() === q) ||
      users.find(u => (u.userPrincipalName || '').toLowerCase() === q) ||
      users.find(u => (u.displayName || '').toLowerCase().startsWith(q)) ||
      null
    );
  }

  public async getPresence(): Promise<Map<string, PresenceAvailability>> {
    if (!this._graphClient) return new Map();
    if (Date.now() < this._presenceExpiry) return this._presenceCache;

    if (this._upnToObjectId.size === 0) await this._fetchObjectIdMap();

    const objectIds = Array.from(this._upnToObjectId.values());
    if (objectIds.length === 0) return this._presenceCache;

    // Build reverse map so we can key results back by UPN
    const reverseMap = new Map<string, string>();
    this._upnToObjectId.forEach((objId, upn) => reverseMap.set(objId, upn));

    this._presenceCache.clear();
    const CHUNK = 650;
    for (let i = 0; i < objectIds.length; i += CHUNK) {
      try {
        const response = await this._graphClient
          .api('/communications/getPresencesByUserId')
          .version('v1.0')
          .post({ ids: objectIds.slice(i, i + CHUNK) });
        const presences: Array<{ id: string; availability: string }> = response?.value || [];
        for (const p of presences) {
          const upn = reverseMap.get(p.id);
          if (upn) this._presenceCache.set(upn, this._normalizeAvailability(p.availability));
        }
      } catch {
        // Presence.Read.All not yet approved or unavailable — return empty
        break;
      }
    }

    this._presenceExpiry = Date.now() + this._PRESENCE_TTL;
    return this._presenceCache;
  }

  public async buildOrgTree(rootUserId: string, levelsBelow: number): Promise<IOrgNode> {
    const allUsers = await this.getAllUsers();
    const user = allUsers.find(u => u.id === rootUserId.toLowerCase())
      || allUsers.find(u => (u.mail || '').toLowerCase() === rootUserId.toLowerCase());
    if (!user) throw new Error(`User not found: ${rootUserId}`);

    const root: IOrgNode = {
      user, directReports: [], isExpanded: true, childrenLoaded: false, level: 0
    };
    this._loadChildren(root, levelsBelow);
    return root;
  }

  /* ── Private helpers ─────────────────────────────────────────────── */

  private _loadChildren(node: IOrgNode, remaining: number): void {
    const reports = this._childrenMap.get(node.user.id) || [];
    node.childrenLoaded = true;
    node.directReports = reports.map(u => {
      const hasKids = (this._childrenMap.get(u.id) || []).length > 0;
      return {
        user: u,
        directReports: [],
        isExpanded: remaining > 1,
        childrenLoaded: remaining <= 1 ? !hasKids : false,
        level: node.level + 1
      };
    });
    if (remaining > 1) {
      node.directReports.forEach(child => this._loadChildren(child, remaining - 1));
    }
  }

  private async _fetchObjectIdMap(): Promise<void> {
    if (!this._graphClient) return;
    let url: string | null = '/users?$select=id,userPrincipalName&$top=999';
    while (url) {
      try {
        const req = this._graphClient.api(url);
        if (!url.startsWith('https://')) req.version('v1.0');
        const response = await req.get();
        const users: Array<{ id: string; userPrincipalName: string }> = response?.value || [];
        for (const u of users) {
          if (u.id && u.userPrincipalName) {
            this._upnToObjectId.set(u.userPrincipalName.toLowerCase(), u.id);
          }
        }
        url = response?.['@odata.nextLink'] || null;
      } catch {
        break;
      }
    }
  }

  private _normalizeAvailability(raw: string): PresenceAvailability {
    const map: { [k: string]: PresenceAvailability } = {
      Available: 'Available', AvailableIdle: 'Available',
      Busy: 'Busy', BusyIdle: 'Busy',
      DoNotDisturb: 'DoNotDisturb',
      BeRightBack: 'BeRightBack',
      Away: 'Away',
      Offline: 'Offline',
      PresenceUnknown: 'Unknown',
    };
    return map[raw] || 'Unknown';
  }

  // Queries Graph API directReports for each SP-found manager and merges in any
  // unlicensed users that SP Search omitted. Silently skips if no graph client.
  private async _supplementUnlicensedReports(users: IGraphUser[]): Promise<IGraphUser[]> {
    if (!this._graphClient) return users;

    const knownIds = new Set<string>();
    for (const u of users) {
      knownIds.add(u.id);
      if (u.mail) knownIds.add(u.mail.toLowerCase());
    }

    // _pendingManagerIds is already populated by _rowToUser during _fetchAllUsers
    const managerIds = Array.from(new Set<string>(this._pendingManagerIds.values()));
    const newUsers: IGraphUser[] = [];

    for (const managerId of managerIds) {
      try {
        const response = await this._graphClient
          .api(`/users/${encodeURIComponent(managerId)}/directReports`)
          .version('v1.0')
          .select('id,displayName,mail,jobTitle,department,officeLocation,mobilePhone,businessPhones,userPrincipalName')
          .top(999)
          .get();

        const reports: Array<{
          id: string; displayName: string; mail: string; jobTitle: string;
          department: string; officeLocation: string; mobilePhone: string;
          businessPhones: string[]; userPrincipalName: string;
        }> = response?.value || [];

        for (const rep of reports) {
          const upn  = (rep.userPrincipalName || '').toLowerCase();
          const mail = (rep.mail || '').toLowerCase();
          const id   = upn || mail || rep.id?.toLowerCase() || '';
          if (!id || knownIds.has(upn) || knownIds.has(mail)) continue;

          knownIds.add(id);
          if (mail) knownIds.add(mail);
          this._photoCache.set(id, null);
          this._pendingManagerIds.set(id, managerId);
          if (rep.id) this._upnToObjectId.set(id, rep.id);

          newUsers.push({
            id,
            displayName:       rep.displayName    || '',
            mail:              rep.mail            || '',
            jobTitle:          rep.jobTitle        || '',
            mobilePhone:       rep.mobilePhone     || '',
            businessPhones:    rep.businessPhones  || [],
            department:        rep.department      || '',
            officeLocation:    rep.officeLocation  || '',
            userPrincipalName: upn,
          });
        }
      } catch {
        // Manager not found in Graph (external user, deleted account, etc.) — skip
      }
    }

    return newUsers.length > 0
      ? [...users, ...newUsers].sort((a, b) => a.displayName.localeCompare(b.displayName))
      : users;
  }

  private async _fetchAllUsers(): Promise<IGraphUser[]> {
    const users: IGraphUser[] = [];
    let startRow  = 0;
    let totalRows = Infinity;

    while (users.length < totalRows) {
      const url =
        `${this._webUrl}/_api/search/query` +
        `?querytext='*'` +
        `&sourceid='${PEOPLE_SOURCE}'` +
        `&selectproperties='${SELECT_PROPS}'` +
        `&rowlimit=${BATCH_SIZE}` +
        `&startrow=${startRow}` +
        `&trimduplicates=false`;

      const resp: SPHttpClientResponse = await this._client.get(
        url,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );

      if (!resp.ok) throw new Error(`People search failed: ${resp.status}`);
      const json = await resp.json();

      const results = json?.PrimaryQueryResult?.RelevantResults;
      if (!results) break;

      totalRows = results.TotalRows || 0;
      const rows: Array<{ Cells: Array<{ Key: string; Value: string }> }> =
        results.Table?.Rows || [];

      for (const row of rows) {
        const user = this._rowToUser(row.Cells);
        if (user) users.push(user);
      }

      startRow += BATCH_SIZE;
      if (rows.length < BATCH_SIZE) break;
    }

    return users.sort((a, b) => (a.displayName || '').localeCompare(b.displayName || ''));
  }

  private _buildMaps(users: IGraphUser[]): void {
    const canonicalId: { [key: string]: string } = {};
    for (const user of users) {
      canonicalId[user.id] = user.id;
      if (user.mail) canonicalId[user.mail.toLowerCase()] = user.id;
      canonicalId[user.displayName.toLowerCase()] = user.id;
    }

    this._childrenMap.clear();
    this._managerMap.clear();
    for (const user of users) {
      const rawMgr = this._pendingManagerIds.get(user.id);
      if (!rawMgr) continue;
      const mgrId = canonicalId[rawMgr] || rawMgr;
      this._managerMap.set(user.id, mgrId);
      if (!this._childrenMap.has(mgrId)) this._childrenMap.set(mgrId, []);
      (this._childrenMap.get(mgrId) as IGraphUser[]).push(user);
    }
    this._pendingManagerIds.clear();
  }

  private _rowToUser(cells: Array<{ Key: string; Value: string }>): IGraphUser | null {
    const p: { [key: string]: string } = {};
    for (const cell of cells) p[cell.Key] = cell.Value || '';

    const accountName = p['AccountName'] || '';
    const displayName = p['PreferredName'] || p['DisplayName'] || '';
    if (!accountName || !displayName) return null;

    // Strip claims prefix from AccountName: i:0#.f|membership|user@domain → user@domain
    const upn = (accountName.includes('|')
      ? (accountName.split('|').pop() || '')
      : accountName).toLowerCase();

    // Email: use WorkEmail if set, fall back to UPN
    const email = (p['WorkEmail'] || upn).toLowerCase();

    // Use UPN as id — same format Manager field uses, so parent-child matching
    // works correctly even in tenants where UPN ≠ WorkEmail.
    const id = upn || email;
    if (!id || !displayName) return null;

    // Manager: extract from claims format, bare email, or store raw (display name fallback)
    const managerRaw = p['Manager'] || '';
    let managerId = '';
    if (managerRaw.includes('|')) {
      managerId = (managerRaw.split('|').pop() || '').toLowerCase();
    } else if (managerRaw.trim()) {
      // bare email or display name — store as-is for _buildMaps to resolve
      managerId = managerRaw.trim().toLowerCase();
    }

    // Photo: only return a URL when the user actually has a profile picture
    const hasPicture = !!(p['PictureURL'] && p['PictureURL'].trim());
    const photoUrl = hasPicture
      ? `${this._webUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(email)}`
      : null;
    this._photoCache.set(id, photoUrl);

    const user: IGraphUser = {
      id,
      displayName,
      mail:              p['WorkEmail'] || upn,
      jobTitle:          p['JobTitle']   || '',
      mobilePhone:       p['MobilePhone'] || '',
      businessPhones:    p['WorkPhone'] ? [p['WorkPhone']] : [],
      department:        p['Department'] || '',
      officeLocation:    p['OfficeNumber'] || '',
      userPrincipalName: upn,
    };
    if (managerId) this._pendingManagerIds.set(id, managerId);
    return user;
  }
}
