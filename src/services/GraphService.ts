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
  country: string;           // AAD profile country — free text, empty on the Search data source
  userPrincipalName: string;
  accountEnabled?: boolean;  // false = disabled / blocked sign-in
  userType?: string;         // 'Member' | 'Guest' | undefined
  employeeType?: string;     // 'Employee' | 'Contractor' | 'PTE' | undefined (tenant-defined)
  dottedManagerId?: string;  // secondary "dotted line" manager (resolved user id)
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

export type DataSource = 'auto' | 'graph' | 'search';

export interface IUserFilterOptions {
  tenantDomain?: string;       // only show users whose email domain matches (e.g. 'contoso.com')
  excludedPatterns?: string[]; // lower-case substrings — hide any user whose name/UPN/mail contains one
  hideGuestUsers?: boolean;    // hide userType === 'Guest'
  hideDisabledAccounts?: boolean; // hide accountEnabled === false
  hideNoJobTitle?: boolean;    // hide users with no jobTitle
  hideNoDepartment?: boolean;  // hide users with no department
}

export class GraphService {
  private _client:         SPHttpClient;
  private _graphClient:    MSGraphClientV3 | undefined;
  private _webUrl:         string;
  private _dataSource:     DataSource;
  private _filterOptions:  IUserFilterOptions;
  private _photoCache:     Map<string, string | null> = new Map();
  private _allUsersCache:  IGraphUser[] | null = null;
  private _childrenMap:    Map<string, IGraphUser[]> = new Map();
  private _managerMap:        Map<string, string> = new Map();   // userId → managerId (resolved)
  private _pendingManagerIds: Map<string, string> = new Map();   // userId → raw managerId (pre-resolution)
  private _dottedLineAttribute: string;                          // AAD extension attribute holding the dotted-line manager
  private _pendingDottedIds:  Map<string, string> = new Map();   // userId → raw dotted managerId (pre-resolution)
  protected _dottedReportsMap: Map<string, IGraphUser[]> = new Map(); // managerId → dotted-line reports
  private _loadingPromise: Promise<IGraphUser[]> | null = null;
  private _upnToObjectId:   Map<string, string> = new Map();     // upn → AAD object id
  private _presenceCache:   Map<string, PresenceAvailability> = new Map();
  private _presenceExpiry = 0;
  private _presenceFailedUntil = 0;                              // back-off after a presence permission failure
  private readonly _PRESENCE_TTL = 60_000;
  private readonly _PRESENCE_RETRY = 15 * 60_000;
  private _prefilterUsers: IGraphUser[] | null = null;           // full fetched list, before user filters

  constructor(
    client: SPHttpClient,
    webUrl: string,
    graphClient?: MSGraphClientV3,
    dataSource: DataSource = 'auto',
    filterOptions: IUserFilterOptions = {},
    dottedLineAttribute = ''
  ) {
    this._client        = client;
    this._webUrl        = webUrl.replace(/\/$/, '');
    this._graphClient   = graphClient;
    this._dataSource    = dataSource;
    this._filterOptions = filterOptions;
    this._dottedLineAttribute = dottedLineAttribute.trim();
  }

  /* ── Public API ──────────────────────────────────────────────────── */

  public getAllUsers(): Promise<IGraphUser[]> {
    if (this._allUsersCache) return Promise.resolve(this._allUsersCache);
    if (this._loadingPromise) return this._loadingPromise;

    this._loadingPromise = this._loadUsers()
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

  private async _loadUsers(): Promise<IGraphUser[]> {
    const useGraph = this._dataSource === 'graph' ||
                     (this._dataSource === 'auto' && !!this._graphClient);

    if (this._dataSource === 'graph' && !this._graphClient) {
      throw new Error(
        'Graph API data source selected but Microsoft Graph permissions have not been granted. ' +
        'Please approve the app permissions in the SharePoint App Catalog, or switch the Data Source setting to "SharePoint Search".'
      );
    }

    if (useGraph) {
      try {
        const users = await this._fetchAllUsersFromGraph();
        if (users.length === 0) throw new Error('Graph API returned 0 users');
        this._prefilterUsers = users;
        return this._applyUserFilters(users);
      } catch (err) {
        if (this._dataSource === 'graph') throw err;
        // 'auto' mode: fall back to SharePoint Search
        console.warn('[SmartOrgChart] Graph API unavailable, falling back to SharePoint Search:', err);
        this._pendingManagerIds.clear();
      }
    }

    // SharePoint Search path (primary or fallback)
    const users = await this._supplementUnlicensedReports(await this._fetchAllUsers());
    this._prefilterUsers = users;
    return this._applyUserFilters(users);
  }

  protected _applyUserFilters(users: IGraphUser[]): IGraphUser[] {
    const { tenantDomain, excludedPatterns, hideGuestUsers, hideDisabledAccounts,
            hideNoJobTitle, hideNoDepartment } = this._filterOptions;
    const hasFilters = tenantDomain || (excludedPatterns && excludedPatterns.length > 0) ||
                       hideGuestUsers || hideDisabledAccounts || hideNoJobTitle || hideNoDepartment;
    if (!hasFilters) return users;

    return users.filter(user => {
      if (hideDisabledAccounts && user.accountEnabled === false) return false;
      if (hideGuestUsers && user.userType === 'Guest') return false;
      if (hideNoJobTitle && !user.jobTitle) return false;
      if (hideNoDepartment && !user.department) return false;
      if (tenantDomain) {
        const emailDomain = ((user.mail || user.id || '').split('@')[1] || '').toLowerCase();
        if (emailDomain && emailDomain !== tenantDomain) return false;
      }
      if (excludedPatterns && excludedPatterns.length > 0) {
        const haystack = [
          user.displayName || '',
          user.id || '',
          user.mail || '',
          user.userPrincipalName || '',
        ].join('\0').toLowerCase();
        if (excludedPatterns.some(p => haystack.includes(p))) return false;
      }
      return true;
    });
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

  public async getDottedLineReports(userId: string): Promise<IGraphUser[]> {
    if (!this._allUsersCache) await this.getAllUsers();
    return this._dottedReportsMap.get(userId.toLowerCase()) || [];
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

  // Fetches presence for the given user IDs only (pass the users currently on
  // screen — fetching the whole tenant every poll does not scale). Omitting
  // userIds falls back to all known users.
  public async getPresence(userIds?: string[]): Promise<Map<string, PresenceAvailability>> {
    if (!this._graphClient) return new Map();
    if (Date.now() < this._presenceFailedUntil) return this._presenceCache;

    if (this._upnToObjectId.size === 0) await this._fetchObjectIdMap();

    const wanted = userIds
      ? Array.from(new Set(userIds.map(id => id.toLowerCase())))
      : Array.from(this._upnToObjectId.keys());

    // Within the TTL, only fetch IDs we have not resolved yet
    const fresh = Date.now() < this._presenceExpiry;
    const toFetch = fresh ? wanted.filter(upn => !this._presenceCache.has(upn)) : wanted;
    if (toFetch.length === 0) return this._presenceCache;

    // Build reverse map so we can key results back by UPN
    const objectIds: string[] = [];
    const reverseMap = new Map<string, string>();
    for (const upn of toFetch) {
      const objId = this._upnToObjectId.get(upn);
      if (objId) { objectIds.push(objId); reverseMap.set(objId, upn); }
    }
    if (objectIds.length === 0) return this._presenceCache;

    const CHUNK = 650;
    for (let i = 0; i < objectIds.length; i += CHUNK) {
      try {
        const response = await this._graphClient
          .api('/communications/getPresencesByUserId')
          .version('v1.0')
          .post({ ids: objectIds.slice(i, i + CHUNK) });
        const presences: Array<{ id: string; availability: string }> = response?.value || [];
        this._presenceFailedUntil = 0;
        for (const p of presences) {
          const upn = reverseMap.get(p.id);
          if (upn) this._presenceCache.set(upn, this._normalizeAvailability(p.availability));
        }
      } catch {
        // Presence.Read.All not yet approved or unavailable — back off so the
        // 60-second polls don't hammer Graph with a permanently failing call
        this._presenceFailedUntil = Date.now() + this._PRESENCE_RETRY;
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

    const processManager = async (managerId: string): Promise<void> => {
      try {
        const response = await (this._graphClient as MSGraphClientV3)
          .api(`/users/${encodeURIComponent(managerId)}/directReports`)
          .version('v1.0')
          .select('id,displayName,mail,jobTitle,department,officeLocation,country,mobilePhone,businessPhones,userPrincipalName,accountEnabled,userType,employeeType')
          .top(999)
          .get();

        const reports: Array<{
          id: string; displayName: string; mail: string; jobTitle: string;
          department: string; officeLocation: string; country: string; mobilePhone: string;
          businessPhones: string[]; userPrincipalName: string;
          accountEnabled?: boolean; userType?: string; employeeType?: string;
        }> = response?.value || [];

        // The loop below runs synchronously after the await, so concurrent
        // managers sharing a report can't both pass the knownIds check
        for (const rep of reports) {
          const upn  = (rep.userPrincipalName || '').toLowerCase();
          const mail = (rep.mail || '').toLowerCase();
          const id   = upn || mail || rep.id?.toLowerCase() || '';
          if (!id || knownIds.has(id) || knownIds.has(upn) || knownIds.has(mail)) continue;

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
            country:           rep.country         || '',
            userPrincipalName: upn,
            accountEnabled:    rep.accountEnabled,
            userType:          rep.userType,
            employeeType:      rep.employeeType,
          });
        }
      } catch {
        // Manager not found in Graph (external user, deleted account, etc.) — skip
      }
    };

    // One request per manager gets slow in large orgs — run a small pool in parallel
    const CONCURRENCY = 8;
    for (let i = 0; i < managerIds.length; i += CONCURRENCY) {
      await Promise.all(managerIds.slice(i, i + CONCURRENCY).map(processManager));
    }

    return newUsers.length > 0
      ? [...users, ...newUsers].sort((a, b) => a.displayName.localeCompare(b.displayName))
      : users;
  }

  private async _fetchAllUsersFromGraph(): Promise<IGraphUser[]> {
    if (!this._graphClient) throw new Error('Graph client not available');

    const users: IGraphUser[] = [];
    let SELECT = 'id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,country,mobilePhone,businessPhones,accountEnabled,userType,employeeType';
    if (this._dottedLineAttribute) SELECT += ',onPremisesExtensionAttributes';
    let url: string | null =
      `/users?$select=${SELECT}&$expand=manager($select=id,userPrincipalName,mail)&$top=999`;

    while (url) {
      const req = this._graphClient.api(url);
      if (!url.startsWith('https://')) req.version('v1.0');
      const response = await req.get();
      const items: Array<{
        id: string;
        displayName: string;
        mail: string;
        userPrincipalName: string;
        jobTitle: string;
        department: string;
        officeLocation: string;
        country: string;
        mobilePhone: string;
        businessPhones: string[];
        accountEnabled?: boolean;
        userType?: string;
        employeeType?: string;
        manager?: { id: string; userPrincipalName: string; mail: string };
        onPremisesExtensionAttributes?: { [key: string]: string | null };
      }> = response?.value || [];

      for (const item of items) {
        const upn  = (item.userPrincipalName || '').toLowerCase();
        const mail = (item.mail || '').toLowerCase();
        const id   = upn || mail;
        if (!id || !item.displayName) continue;

        // Pre-populate photo cache with SharePoint profile photo URL
        const photoEmail = mail || upn;
        const photoUrl = photoEmail
          ? `${this._webUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(photoEmail)}`
          : null;
        this._photoCache.set(id, photoUrl);

        // Store AAD object ID for presence lookups
        if (item.id) this._upnToObjectId.set(id, item.id);

        // Manager relationship — prefer UPN, fall back to mail
        const mgrUpn  = (item.manager?.userPrincipalName || '').toLowerCase();
        const mgrMail = (item.manager?.mail || '').toLowerCase();
        const mgrId   = mgrUpn || mgrMail;
        if (mgrId) this._pendingManagerIds.set(id, mgrId);

        // Dotted-line manager from the configured extension attribute (email or UPN)
        if (this._dottedLineAttribute) {
          const dotted = (item.onPremisesExtensionAttributes?.[this._dottedLineAttribute] || '').trim().toLowerCase();
          if (dotted) this._pendingDottedIds.set(id, dotted);
        }

        users.push({
          id,
          displayName:       item.displayName,
          mail:              mail,
          jobTitle:          item.jobTitle          || '',
          mobilePhone:       item.mobilePhone        || '',
          businessPhones:    item.businessPhones     || [],
          department:        item.department         || '',
          officeLocation:    item.officeLocation     || '',
          country:           item.country            || '',
          userPrincipalName: upn,
          accountEnabled:    item.accountEnabled,
          userType:          item.userType,
          employeeType:      item.employeeType,
        });
      }

      url = response?.['@odata.nextLink'] || null;
    }

    return users.sort((a, b) => (a.displayName || '').localeCompare(b.displayName || ''));
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
    // Resolve manager references against ALL fetched users (before the admin
    // user filters), so a report whose manager is hidden by a filter can be
    // bridged to the nearest visible ancestor instead of orphaning the branch.
    const allUsers = this._prefilterUsers && this._prefilterUsers.length >= users.length
      ? this._prefilterUsers
      : users;

    // Display name is a last-resort manager lookup (SP Search sometimes stores
    // only the manager's name). Skip names shared by multiple users — guessing
    // would silently attach reports to the wrong person.
    const nameCount: { [key: string]: number } = {};
    for (const user of allUsers) {
      const dn = user.displayName.toLowerCase();
      nameCount[dn] = (nameCount[dn] || 0) + 1;
    }

    const canonicalId: { [key: string]: string } = {};
    for (const user of allUsers) {
      canonicalId[user.id] = user.id;
      if (user.mail) canonicalId[user.mail.toLowerCase()] = user.id;
      const dn = user.displayName.toLowerCase();
      if (nameCount[dn] === 1) canonicalId[dn] = user.id;
    }

    // Raw manager edges across all fetched users. Self-managed accounts
    // (common for CEOs in Azure AD) are dropped here — without this guard a
    // user renders as their own child and Expand All never terminates.
    const rawManagerOf = new Map<string, string>();
    for (const user of allUsers) {
      const rawMgr = this._pendingManagerIds.get(user.id);
      if (!rawMgr) continue;
      const mgrId = canonicalId[rawMgr] || rawMgr;
      if (mgrId !== user.id) rawManagerOf.set(user.id, mgrId);
    }

    const visibleIds = new Set(users.map(u => u.id));

    this._childrenMap.clear();
    this._managerMap.clear();
    for (const user of users) {
      // Walk up through hidden managers to the nearest visible one; the
      // visited set stops manager cycles (A→B→A) from looping forever
      let mgrId = rawManagerOf.get(user.id);
      const visited = new Set<string>([user.id]);
      while (mgrId && !visibleIds.has(mgrId) && !visited.has(mgrId)) {
        visited.add(mgrId);
        mgrId = rawManagerOf.get(mgrId);
      }
      if (!mgrId || !visibleIds.has(mgrId) || mgrId === user.id) continue;
      this._managerMap.set(user.id, mgrId);
      if (!this._childrenMap.has(mgrId)) this._childrenMap.set(mgrId, []);
      (this._childrenMap.get(mgrId) as IGraphUser[]).push(user);
    }
    this._pendingManagerIds.clear();

    this._dottedReportsMap.clear();
    for (const user of users) {
      const rawDotted = this._pendingDottedIds.get(user.id);
      if (!rawDotted) continue;
      const dottedId = canonicalId[rawDotted];
      if (!dottedId || dottedId === user.id || !visibleIds.has(dottedId)) continue; // unresolvable, hidden, or self-reference
      user.dottedManagerId = dottedId;
      if (!this._dottedReportsMap.has(dottedId)) this._dottedReportsMap.set(dottedId, []);
      (this._dottedReportsMap.get(dottedId) as IGraphUser[]).push(user);
    }
    this._pendingDottedIds.clear();
    this._prefilterUsers = null;
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
      // The people search source exposes no Country managed property, and no
      // EmployeeType at all — both stay empty on this data source. Same limitation
      // as userType / accountEnabled above.
      country:           '',
      userPrincipalName: upn,
    };
    if (managerId) this._pendingManagerIds.set(id, managerId);
    return user;
  }
}
