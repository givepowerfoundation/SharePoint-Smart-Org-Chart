import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { IGraphUser, PresenceAvailability } from '../../../../services/GraphService';
import { IEmployeeDirectoryProps } from './IEmployeeDirectoryProps';
import { exportDirectoryToExcel } from '../../../../services/PdfExportService';
import { PRESENCE_COLOR, getInitials } from '../personUtils';
import { getCountryColor } from '../countryUtils';
import styles from './EmployeeDirectory.module.scss';

const ALPHABET = ['All', ...'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('')];

type ViewMode = 'card' | 'list';

const LS_VIEWMODE_KEY = 'smartOrgChart_dirViewMode';

function viewModeKey(instanceId: string): string {
  return instanceId ? `${LS_VIEWMODE_KEY}_${instanceId}` : LS_VIEWMODE_KEY;
}

function readViewMode(instanceId: string): ViewMode {
  try {
    const v = localStorage.getItem(viewModeKey(instanceId));
    if (v === 'card' || v === 'list') return v;
  } catch { /* ignore */ }
  return 'card';
}

interface IEmployeeDirectoryState {
  users: IGraphUser[];
  isLoading: boolean;
  error: string | null;
  selectedLetter: string;
  searchQuery: string;
  photos: { [id: string]: string | null };
  presenceMap: Map<string, PresenceAvailability>;
  currentPage: number;
  viewMode: ViewMode;
  selectedDepartment: string;
  selectedOffice: string;
  selectedCountry: string;
  selectedEmployeeType: string;
}

export class EmployeeDirectory extends React.Component<IEmployeeDirectoryProps, IEmployeeDirectoryState> {
  private _photoQueue: string[] = [];
  private _photoQueueSet: Set<string> = new Set();
  private _processingPhotos = false;
  private _mounted = false;
  private _presenceInterval: number | null = null;
  private _presenceDebounce: number | null = null;

  constructor(props: IEmployeeDirectoryProps) {
    super(props);
    this.state = {
      users: [], isLoading: true, error: null,
      selectedLetter: 'All', searchQuery: '',
      photos: {}, presenceMap: new Map(), currentPage: 1,
      viewMode: readViewMode(props.instanceId),
      selectedDepartment: '', selectedOffice: '',
      selectedCountry: '', selectedEmployeeType: ''
    };
  }

  public async componentDidMount(): Promise<void> {
    this._mounted = true;
    await this._loadUsers();
    this._refreshPresence();
    this._presenceInterval = window.setInterval(() => { this._refreshPresence(); }, 60_000);
  }

  public componentDidUpdate(prevProps: IEmployeeDirectoryProps, prevState: IEmployeeDirectoryState): void {
    if (prevProps.alphabetFilterField !== this.props.alphabetFilterField) {
      this.setState({ selectedLetter: 'All', currentPage: 1 });
      return;
    }
    if (
      prevState.users !== this.state.users ||
      prevState.selectedLetter !== this.state.selectedLetter ||
      prevState.searchQuery !== this.state.searchQuery ||
      prevState.currentPage !== this.state.currentPage ||
      prevState.selectedDepartment !== this.state.selectedDepartment ||
      prevState.selectedOffice !== this.state.selectedOffice
    ) {
      const paged = this._getCurrentPageUsers();
      this._enqueuePhotos(paged.map(u => u.id));
      // Presence for the newly visible page — debounced so every search
      // keystroke doesn't fire a Graph call for the transient page
      if (this._presenceDebounce !== null) window.clearTimeout(this._presenceDebounce);
      this._presenceDebounce = window.setTimeout(() => {
        this._presenceDebounce = null;
        this._refreshPresence().catch(() => { /* ignore */ });
      }, 500);
    }
  }

  public componentWillUnmount(): void {
    this._mounted = false;
    if (this._presenceInterval !== null) {
      window.clearInterval(this._presenceInterval);
      this._presenceInterval = null;
    }
    if (this._presenceDebounce !== null) {
      window.clearTimeout(this._presenceDebounce);
      this._presenceDebounce = null;
    }
  }

  private async _refreshPresence(): Promise<void> {
    if (!this._mounted) return;
    // Only request presence for the page of users currently on screen
    const paged = this._getCurrentPageUsers();
    if (paged.length === 0) return;
    const presenceMap = await this.props.graphService.getPresence(paged.map(u => u.id));
    if (this._mounted) this.setState({ presenceMap });
  }

  private _getCurrentPageUsers(): IGraphUser[] {
    const filtered = this._getFilteredUsers();
    const { pageSize } = this.props;
    const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
    const safePage = Math.min(this.state.currentPage, totalPages);
    return filtered.slice((safePage - 1) * pageSize, safePage * pageSize);
  }

  public exportExcel(): void {
    const { showEmail, showPhone, showDepartment, showOffice } = this.props;
    exportDirectoryToExcel(this._getFilteredUsers(), { showEmail, showPhone, showDepartment, showOffice });
  }

  private async _loadUsers(): Promise<void> {
    try {
      // GraphService/MockGraphService both return the list pre-sorted by display name
      const users = await this.props.graphService.getAllUsers();
      if (this._mounted) this.setState({ users, isLoading: false });
    } catch (err) {
      const detail = err instanceof Error && err.message
        ? err.message
        : 'Ensure the web part has User.Read.All permission.';
      if (this._mounted) this.setState({ isLoading: false, error: `Failed to load employees. ${detail}` });
    }
  }

  private _enqueuePhotos(ids: string[]): void {
    const toLoad = ids.filter(id => !(id in this.state.photos) && !this._photoQueueSet.has(id));
    toLoad.forEach(id => { this._photoQueue.push(id); this._photoQueueSet.add(id); });
    if (!this._processingPhotos) this._drainPhotoQueue();
  }

  private async _drainPhotoQueue(): Promise<void> {
    this._processingPhotos = true;
    // Batch setState calls — one render per 10 photos instead of one per photo
    let batch: { [id: string]: string | null } = {};
    const flush = (): void => {
      const toApply = batch;
      batch = {};
      if (this._mounted && Object.keys(toApply).length > 0) {
        this.setState(prev => ({ photos: { ...prev.photos, ...toApply } }));
      }
    };
    while (this._photoQueue.length > 0 && this._mounted) {
      const id = this._photoQueue.shift();
      if (!id) break;
      this._photoQueueSet.delete(id);
      if (id in this.state.photos || id in batch) continue;
      batch[id] = await this.props.graphService.getUserPhoto(id);
      if (Object.keys(batch).length >= 10) flush();
    }
    flush();
    this._processingPhotos = false;
  }

  private _getFirstName(dn: string): string { return (dn || '').split(' ')[0] || ''; }
  private _getLastName(dn: string): string {
    const p = (dn || '').split(' ').filter(x => x);
    return p.length > 1 ? p[p.length - 1] : p[0] || '';
  }
  private _getPhone(user: IGraphUser): string {
    return user.mobilePhone || (user.businessPhones && user.businessPhones[0]) || '';
  }

  private _getDepartments(): IDropdownOption[] {
    const seen: { [k: string]: boolean } = {};
    const depts: string[] = [];
    this.state.users.forEach(u => { if (u.department && !seen[u.department]) { seen[u.department] = true; depts.push(u.department); } });
    depts.sort();
    return [{ key: '', text: 'All Departments' }, ...depts.map(d => ({ key: d, text: d }))];
  }

  private _getOffices(): IDropdownOption[] {
    const seen: { [k: string]: boolean } = {};
    const offices: string[] = [];
    this.state.users.forEach(u => { if (u.officeLocation && !seen[u.officeLocation]) { seen[u.officeLocation] = true; offices.push(u.officeLocation); } });
    offices.sort();
    return [{ key: '', text: 'All Offices' }, ...offices.map(o => ({ key: o, text: o }))];
  }

  private _getCountries(): IDropdownOption[] {
    const seen: { [k: string]: boolean } = {};
    const countries: string[] = [];
    this.state.users.forEach(u => { if (u.country && !seen[u.country]) { seen[u.country] = true; countries.push(u.country); } });
    countries.sort();
    return [{ key: '', text: 'All Countries' }, ...countries.map(c => ({ key: c, text: c }))];
  }

  private _getEmployeeTypes(): IDropdownOption[] {
    const seen: { [k: string]: boolean } = {};
    const types: string[] = [];
    this.state.users.forEach(u => { if (u.employeeType && !seen[u.employeeType]) { seen[u.employeeType] = true; types.push(u.employeeType); } });
    types.sort();
    return [{ key: '', text: 'All Employee Types' }, ...types.map(t => ({ key: t, text: t }))];
  }

  private _getFilteredUsers(): IGraphUser[] {
    const { users, selectedLetter, searchQuery, selectedDepartment, selectedOffice,
            selectedCountry, selectedEmployeeType } = this.state;
    const { alphabetFilterField } = this.props;
    let result = users;

    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      result = result.filter(u =>
        (u.displayName || '').toLowerCase().includes(q) ||
        (u.mail || '').toLowerCase().includes(q) ||
        (u.jobTitle || '').toLowerCase().includes(q) ||
        (u.department || '').toLowerCase().includes(q)
      );
    } else if (selectedLetter !== 'All') {
      result = result.filter(u => {
        const name = alphabetFilterField === 'firstName'
          ? this._getFirstName(u.displayName)
          : this._getLastName(u.displayName);
        return name.toUpperCase().startsWith(selectedLetter);
      });
    }

    if (selectedDepartment) result = result.filter(u => u.department === selectedDepartment);
    if (selectedOffice) result = result.filter(u => u.officeLocation === selectedOffice);
    if (selectedCountry) result = result.filter(u => u.country === selectedCountry);
    if (selectedEmployeeType) result = result.filter(u => u.employeeType === selectedEmployeeType);

    return result;
  }

  private _selectLetter = (letter: string): void => {
    this.setState({ selectedLetter: letter, searchQuery: '', currentPage: 1 });
  }

  private _onSearch = (value: string): void => {
    this.setState({ searchQuery: value || '', selectedLetter: 'All', currentPage: 1 });
  }

  private _clearFilters = (): void => {
    this.setState({
      selectedDepartment: '', selectedOffice: '', selectedCountry: '', selectedEmployeeType: '',
      selectedLetter: 'All', searchQuery: '', currentPage: 1
    });
  }

  private _setViewMode = (mode: ViewMode): void => {
    try { localStorage.setItem(viewModeKey(this.props.instanceId), mode); } catch { /* ignore */ }
    this.setState({ viewMode: mode });
  }

  /* ── Render helpers ── */

  // Country and employee-type pills, shared by the card grid and the list view.
  // The country pill is tinted from its configured colour; employee type stays neutral.
  private _renderFacetBadges(user: IGraphUser): React.ReactElement | null {
    if (!user.country && !user.employeeType) return null;
    const countryColor = user.country ? getCountryColor(user.country, this.props.countryColors) : '';
    return (
      <>
        {user.country && (
          <span
            className={`${styles.statusBadge} ${styles.statusCountry}`}
            style={{ background: `${countryColor}1a`, color: countryColor }}
            title={user.country}
          >
            {user.country}
          </span>
        )}
        {user.employeeType && (
          <span className={`${styles.statusBadge} ${styles.statusEmployeeType}`}>
            {user.employeeType}
          </span>
        )}
      </>
    );
  }

  private _renderCardGrid(paged: IGraphUser[]): React.ReactElement {
    const { cardSize, showEmail, showPhone, showDepartment, showOffice } = this.props;
    const { photos, presenceMap } = this.state;

    return (
      <div className={`${styles.grid} ${styles[`size_${cardSize}`]}`}>
        {paged.map(user => {
          const phone = this._getPhone(user);
          const photoUrl = photos[user.id] || null;
          const hasDetails =
            (showDepartment && !!user.department) ||
            (showOffice && !!user.officeLocation) ||
            (showEmail && !!user.mail) ||
            (showPhone && !!phone);
          return (
            <div key={user.id} className={styles.card}>
              <div className={styles.cardHeader}>
                <div className={styles.avatarWrap}>
                  {photoUrl
                    ? <img src={photoUrl} alt={user.displayName} className={styles.avatar} />
                    : <div className={styles.initials}>{getInitials(user.displayName)}</div>
                  }
                  {(() => {
                    const s = presenceMap.get(user.id);
                    return s && s !== 'Unknown'
                      ? <span className={styles.presenceDot} style={{ background: PRESENCE_COLOR[s] }} />
                      : null;
                  })()}
                </div>
                <div className={styles.cardBody}>
                  <div className={styles.userName}>{user.displayName}</div>
                  {user.jobTitle && <div className={styles.jobTitle}>{user.jobTitle}</div>}
                  <div className={styles.statusBadges}>
                    {this._renderFacetBadges(user)}
                    {user.accountEnabled === false && (
                      <span className={`${styles.statusBadge} ${styles.statusDisabled}`}>Disabled</span>
                    )}
                    {user.userType === 'Guest' && (
                      <span className={`${styles.statusBadge} ${styles.statusGuest}`}>Guest</span>
                    )}
                  </div>
                </div>
                {user.mail && (
                  <a
                    href={`https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(user.mail)}`}
                    target="_blank" rel="noopener noreferrer"
                    className={styles.chatBtn} title="Chat in Teams"
                    onClick={e => e.stopPropagation()}
                  >
                    <Icon iconName="Chat" />
                  </a>
                )}
              </div>
              {hasDetails && (
                <div className={styles.cardDetails}>
                  {showDepartment && user.department && (
                    <div className={styles.detail}>
                      <Icon iconName="Work" className={styles.detailIcon} />
                      <span>{user.department}</span>
                    </div>
                  )}
                  {showOffice && user.officeLocation && (
                    <div className={styles.detail}>
                      <Icon iconName="POI" className={styles.detailIcon} />
                      <span>{user.officeLocation}</span>
                    </div>
                  )}
                  {showEmail && user.mail && (
                    <div className={styles.detail}>
                      <Icon iconName="Mail" className={styles.detailIcon} />
                      <a href={`mailto:${user.mail}`} className={styles.link}>{user.mail}</a>
                    </div>
                  )}
                  {showPhone && phone && (
                    <div className={styles.detail}>
                      <Icon iconName="Phone" className={styles.detailIcon} />
                      <a href={`tel:${phone}`} className={styles.link}>{phone}</a>
                    </div>
                  )}
                </div>
              )}
            </div>
          );
        })}
      </div>
    );
  }

  private _renderListView(paged: IGraphUser[]): React.ReactElement {
    const { showEmail, showPhone, showDepartment, showOffice } = this.props;
    const { photos, presenceMap } = this.state;

    return (
      <table className={styles.listTable}>
        <thead>
          <tr className={styles.listHead}>
            <th scope="col" className={styles.listTh}>Name</th>
            <th scope="col" className={styles.listTh}>Job Title</th>
            {showDepartment && <th scope="col" className={styles.listTh}>Department</th>}
            {showOffice && <th scope="col" className={styles.listTh}>Office</th>}
            {showEmail && <th scope="col" className={styles.listTh}>Email</th>}
            {showPhone && <th scope="col" className={styles.listTh}>Phone</th>}
            <th scope="col" className={styles.listTh} />
          </tr>
        </thead>
        <tbody>
          {paged.map(user => {
            const phone = this._getPhone(user);
            const photoUrl = photos[user.id] || null;
            return (
              <tr key={user.id} className={styles.listRow}>
                <td className={styles.listTd}>
                  <div className={styles.listNameCell}>
                    <div className={styles.listAvatarWrap}>
                      {photoUrl
                        ? <img src={photoUrl} alt={user.displayName} className={styles.listAvatar} />
                        : <div className={styles.listInitials}>{getInitials(user.displayName)}</div>
                      }
                      {(() => {
                        const s = presenceMap.get(user.id);
                        return s && s !== 'Unknown'
                          ? <span className={styles.presenceDot} style={{ background: PRESENCE_COLOR[s], width: 10, height: 10 }} />
                          : null;
                      })()}
                    </div>
                    <span className={styles.listName}>{user.displayName}</span>
                    {this._renderFacetBadges(user)}
                    {user.accountEnabled === false && (
                      <span className={`${styles.statusBadge} ${styles.statusDisabled}`}>Disabled</span>
                    )}
                    {user.userType === 'Guest' && (
                      <span className={`${styles.statusBadge} ${styles.statusGuest}`}>Guest</span>
                    )}
                  </div>
                </td>
                <td className={styles.listTd}><span className={styles.listCell}>{user.jobTitle || '—'}</span></td>
                {showDepartment && <td className={styles.listTd}><span className={styles.listCell}>{user.department || '—'}</span></td>}
                {showOffice && <td className={styles.listTd}><span className={styles.listCell}>{user.officeLocation || '—'}</span></td>}
                {showEmail && (
                  <td className={styles.listTd}>
                    {user.mail
                      ? <a href={`mailto:${user.mail}`} className={styles.listLink}>{user.mail}</a>
                      : <span className={styles.listCell}>—</span>}
                  </td>
                )}
                {showPhone && (
                  <td className={styles.listTd}>
                    {phone
                      ? <a href={`tel:${phone}`} className={styles.listLink}>{phone}</a>
                      : <span className={styles.listCell}>—</span>}
                  </td>
                )}
                <td className={styles.listTd}>
                  {user.mail && (
                    <a
                      href={`https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(user.mail)}`}
                      target="_blank" rel="noopener noreferrer"
                      className={styles.listChatBtn} title="Chat in Teams"
                    >
                      <Icon iconName="Chat" />
                    </a>
                  )}
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    );
  }

  public render(): React.ReactElement {
    const { isLoading, error, selectedLetter, currentPage, searchQuery,
            viewMode, selectedDepartment, selectedOffice,
            selectedCountry, selectedEmployeeType } = this.state;
    const { pageSize } = this.props;

    if (isLoading) return (
      <div className={styles.centered}><Spinner size={SpinnerSize.large} label="Loading employees..." /></div>
    );

    if (error) return (
      <div className={styles.errorMsg}><Icon iconName="Warning" /><span>{error}</span></div>
    );

    const deptOptions = this._getDepartments();
    const officeOptions = this._getOffices();
    const countryOptions = this._getCountries();
    const employeeTypeOptions = this._getEmployeeTypes();
    const filtered = this._getFilteredUsers();
    const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
    const safePage = Math.min(currentPage, totalPages);
    const paged = filtered.slice((safePage - 1) * pageSize, safePage * pageSize);

    const activeFilters = (selectedDepartment ? 1 : 0) + (selectedOffice ? 1 : 0) +
                          (selectedCountry ? 1 : 0) + (selectedEmployeeType ? 1 : 0);

    const THEME_CLASS: Partial<Record<string, string>> = {
      minimal: styles.themeMinimal,
      corporate: styles.themeCorporate,
      dark: styles.themeDark,
    };
    const themeClass = THEME_CLASS[this.props.theme] || '';

    return (
      <div className={[styles.container, themeClass].filter(Boolean).join(' ')}>

        {/* ── Top bar: search + filters + view toggle ── */}
        <div className={styles.toolbar}>
          <SearchBox
            placeholder="Search by name, title, email, or department..."
            value={searchQuery}
            onChange={(_, v) => this._onSearch(v || '')}
            className={styles.searchBox}
            underlined
          />

          {deptOptions.length > 2 && (
            <Dropdown
              placeholder="Department"
              selectedKey={selectedDepartment}
              options={deptOptions}
              onChange={(_, o) => o && this.setState({ selectedDepartment: o.key as string, currentPage: 1 })}
              className={styles.filterDropdown}
            />
          )}

          {officeOptions.length > 2 && (
            <Dropdown
              placeholder="Office"
              selectedKey={selectedOffice}
              options={officeOptions}
              onChange={(_, o) => o && this.setState({ selectedOffice: o.key as string, currentPage: 1 })}
              className={styles.filterDropdown}
            />
          )}

          {countryOptions.length > 2 && (
            <Dropdown
              placeholder="Country"
              selectedKey={selectedCountry}
              options={countryOptions}
              onChange={(_, o) => o && this.setState({ selectedCountry: o.key as string, currentPage: 1 })}
              className={styles.filterDropdown}
            />
          )}

          {employeeTypeOptions.length > 2 && (
            <Dropdown
              placeholder="Employee Type"
              selectedKey={selectedEmployeeType}
              options={employeeTypeOptions}
              onChange={(_, o) => o && this.setState({ selectedEmployeeType: o.key as string, currentPage: 1 })}
              className={styles.filterDropdown}
            />
          )}

          {activeFilters > 0 && (
            <button className={styles.clearBtn} onClick={this._clearFilters} title="Clear all filters">
              <Icon iconName="Cancel" /> Clear
            </button>
          )}

          <button
            className={styles.exportExcelBtn}
            onClick={() => this.exportExcel()}
            title="Export all filtered results to Excel"
          >
            <Icon iconName="Download" />
            <span>Export</span>
          </button>

          <div className={styles.viewToggle}>
            <button
              className={`${styles.viewBtn} ${viewMode === 'card' ? styles.viewBtnActive : ''}`}
              onClick={() => this._setViewMode('card')}
              title="Card view"
            >
              <Icon iconName="GridViewMedium" />
            </button>
            <button
              className={`${styles.viewBtn} ${viewMode === 'list' ? styles.viewBtnActive : ''}`}
              onClick={() => this._setViewMode('list')}
              title="List view"
            >
              <Icon iconName="BulletedList" />
            </button>
          </div>
        </div>

        {/* ── Alphabet bar ── */}
        <div className={styles.alphabetBar} role="toolbar" aria-label="Alphabet filter">
          {ALPHABET.map(letter => (
            <button
              key={letter}
              className={`${styles.letterBtn} ${selectedLetter === letter && !searchQuery ? styles.active : ''}`}
              onClick={() => this._selectLetter(letter)}
              aria-pressed={selectedLetter === letter && !searchQuery}
              title={letter === 'All' ? 'Show all' : `Filter by ${letter}`}
            >
              {letter}
            </button>
          ))}
        </div>

        {/* ── Result count + top pagination ── */}
        <div className={styles.resultMeta}>
          <span>
            {filtered.length} {filtered.length === 1 ? 'person' : 'people'}
            {totalPages > 1 && ` · page ${safePage} of ${totalPages}`}
            {activeFilters > 0 && <span className={styles.filterBadge}>{activeFilters} filter{activeFilters > 1 ? 's' : ''} active</span>}
          </span>
          {totalPages > 1 && (
            <div className={styles.paginationInline}>
              <button
                className={styles.pageBtn}
                onClick={() => this.setState({ currentPage: Math.max(1, safePage - 1) })}
                disabled={safePage === 1}
              >
                <Icon iconName="ChevronLeft" /> Prev
              </button>
              <span className={styles.pageInfo}>{safePage} / {totalPages}</span>
              <button
                className={styles.pageBtn}
                onClick={() => this.setState({ currentPage: Math.min(totalPages, safePage + 1) })}
                disabled={safePage === totalPages}
              >
                Next <Icon iconName="ChevronRight" />
              </button>
            </div>
          )}
        </div>

        {/* ── Content ── */}
        {viewMode === 'card' ? this._renderCardGrid(paged) : this._renderListView(paged)}

        {filtered.length === 0 && (
          <div className={styles.noResults}>
            <Icon iconName="SearchIssue" />
            <span>No people found</span>
            {activeFilters > 0 && (
              <button className={styles.clearBtn} onClick={this._clearFilters}>Clear filters</button>
            )}
          </div>
        )}

        {/* ── Bottom pagination ── */}
        {totalPages > 1 && (
          <div className={styles.pagination}>
            <button
              className={styles.pageBtn}
              onClick={() => this.setState({ currentPage: Math.max(1, safePage - 1) })}
              disabled={safePage === 1}
            >
              <Icon iconName="ChevronLeft" /> Prev
            </button>
            <span className={styles.pageInfo}>{safePage} / {totalPages}</span>
            <button
              className={styles.pageBtn}
              onClick={() => this.setState({ currentPage: Math.min(totalPages, safePage + 1) })}
              disabled={safePage === totalPages}
            >
              Next <Icon iconName="ChevronRight" />
            </button>
          </div>
        )}
      </div>
    );
  }
}
