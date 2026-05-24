import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { IGraphUser, PresenceAvailability } from '../../../../services/GraphService';
import { IEmployeeDirectoryProps } from './IEmployeeDirectoryProps';
import { exportDirectoryToPdf } from '../../../../services/PdfExportService';
import styles from './EmployeeDirectory.module.scss';

const ALPHABET = ['All', ...'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('')];

const PRESENCE_COLOUR: Record<PresenceAvailability, string> = {
  Available:    '#6BB700',
  Busy:         '#C50F1F',
  DoNotDisturb: '#C50F1F',
  BeRightBack:  '#FFAA44',
  Away:         '#FFAA44',
  Offline:      '#8A8886',
  Unknown:      '#8A8886',
};

type ViewMode = 'card' | 'list';

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
}

export class EmployeeDirectory extends React.Component<IEmployeeDirectoryProps, IEmployeeDirectoryState> {
  private _photoQueue: string[] = [];
  private _photoQueueSet: Set<string> = new Set();
  private _processingPhotos = false;
  private _mounted = false;
  private _presenceInterval: number | null = null;

  constructor(props: IEmployeeDirectoryProps) {
    super(props);
    this.state = {
      users: [], isLoading: true, error: null,
      selectedLetter: 'All', searchQuery: '',
      photos: {}, presenceMap: new Map(), currentPage: 1,
      viewMode: 'card',
      selectedDepartment: '', selectedOffice: ''
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
      const filtered = this._getFilteredUsers();
      const { pageSize } = this.props;
      const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
      const safePage = Math.min(this.state.currentPage, totalPages);
      const paged = filtered.slice((safePage - 1) * pageSize, safePage * pageSize);
      this._enqueuePhotos(paged.map(u => u.id));
    }
  }

  public componentWillUnmount(): void {
    this._mounted = false;
    if (this._presenceInterval !== null) {
      window.clearInterval(this._presenceInterval);
      this._presenceInterval = null;
    }
  }

  private async _refreshPresence(): Promise<void> {
    if (!this._mounted) return;
    const presenceMap = await this.props.graphService.getPresence();
    if (this._mounted) this.setState({ presenceMap });
  }

  public exportPdf(): void {
    const { showEmail, showPhone, showDepartment, showOffice } = this.props;
    exportDirectoryToPdf(this._getFilteredUsers(), { showEmail, showPhone, showDepartment, showOffice });
  }

  private async _loadUsers(): Promise<void> {
    try {
      const users = await this.props.graphService.getAllUsers();
      const sorted = [...users].sort((a, b) => (a.displayName || '').localeCompare(b.displayName || ''));
      if (this._mounted) this.setState({ users: sorted, isLoading: false });
    } catch {
      if (this._mounted) this.setState({ isLoading: false, error: 'Failed to load employees. Ensure the web part has User.Read.All permission.' });
    }
  }

  private _enqueuePhotos(ids: string[]): void {
    const toLoad = ids.filter(id => !(id in this.state.photos) && !this._photoQueueSet.has(id));
    toLoad.forEach(id => { this._photoQueue.push(id); this._photoQueueSet.add(id); });
    if (!this._processingPhotos) this._drainPhotoQueue();
  }

  private async _drainPhotoQueue(): Promise<void> {
    this._processingPhotos = true;
    while (this._photoQueue.length > 0 && this._mounted) {
      const id = this._photoQueue.shift();
      if (!id) break;
      this._photoQueueSet.delete(id);
      if (id in this.state.photos) continue;
      const url = await this.props.graphService.getUserPhoto(id);
      if (this._mounted) this.setState(prev => ({ photos: { ...prev.photos, [id]: url } }));
    }
    this._processingPhotos = false;
  }

  private _getInitials(displayName: string): string {
    const parts = (displayName || '').split(' ').filter(p => p.length > 0);
    if (parts.length === 0) return '?';
    if (parts.length === 1) return parts[0][0].toUpperCase();
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
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

  private _getFilteredUsers(): IGraphUser[] {
    const { users, selectedLetter, searchQuery, selectedDepartment, selectedOffice } = this.state;
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

    return result;
  }

  private _selectLetter = (letter: string): void => {
    this.setState({ selectedLetter: letter, searchQuery: '', currentPage: 1 });
  }

  private _onSearch = (value: string): void => {
    this.setState({ searchQuery: value || '', selectedLetter: 'All', currentPage: 1 });
  }

  private _clearFilters = (): void => {
    this.setState({ selectedDepartment: '', selectedOffice: '', selectedLetter: 'All', searchQuery: '', currentPage: 1 });
  }

  /* ── Render helpers ── */

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
                    : <div className={styles.initials}>{this._getInitials(user.displayName)}</div>
                  }
                  {(() => {
                    const s = presenceMap.get(user.id);
                    return s && s !== 'Unknown'
                      ? <span className={styles.presenceDot} style={{ background: PRESENCE_COLOUR[s] }} />
                      : null;
                  })()}
                </div>
                <div className={styles.cardBody}>
                  <div className={styles.userName}>{user.displayName}</div>
                  {user.jobTitle && <div className={styles.jobTitle}>{user.jobTitle}</div>}
                  <div className={styles.statusBadges}>
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
            <th className={styles.listTh}>Name</th>
            <th className={styles.listTh}>Job Title</th>
            {showDepartment && <th className={styles.listTh}>Department</th>}
            {showOffice && <th className={styles.listTh}>Office</th>}
            {showEmail && <th className={styles.listTh}>Email</th>}
            {showPhone && <th className={styles.listTh}>Phone</th>}
            <th className={styles.listTh} />
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
                        : <div className={styles.listInitials}>{this._getInitials(user.displayName)}</div>
                      }
                      {(() => {
                        const s = presenceMap.get(user.id);
                        return s && s !== 'Unknown'
                          ? <span className={styles.presenceDot} style={{ background: PRESENCE_COLOUR[s], width: 8, height: 8 }} />
                          : null;
                      })()}
                    </div>
                    <span className={styles.listName}>{user.displayName}</span>
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
            viewMode, selectedDepartment, selectedOffice } = this.state;
    const { pageSize } = this.props;

    if (isLoading) return (
      <div className={styles.centered}><Spinner size={SpinnerSize.large} label="Loading employees..." /></div>
    );

    if (error) return (
      <div className={styles.errorMsg}><Icon iconName="Warning" /><span>{error}</span></div>
    );

    const deptOptions = this._getDepartments();
    const officeOptions = this._getOffices();
    const filtered = this._getFilteredUsers();
    const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
    const safePage = Math.min(currentPage, totalPages);
    const paged = filtered.slice((safePage - 1) * pageSize, safePage * pageSize);

    const activeFilters = (selectedDepartment ? 1 : 0) + (selectedOffice ? 1 : 0);

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

          {activeFilters > 0 && (
            <button className={styles.clearBtn} onClick={this._clearFilters} title="Clear all filters">
              <Icon iconName="Cancel" /> Clear
            </button>
          )}

          <div className={styles.viewToggle}>
            <button
              className={`${styles.viewBtn} ${viewMode === 'card' ? styles.viewBtnActive : ''}`}
              onClick={() => this.setState({ viewMode: 'card' })}
              title="Card view"
            >
              <Icon iconName="GridViewMedium" />
            </button>
            <button
              className={`${styles.viewBtn} ${viewMode === 'list' ? styles.viewBtnActive : ''}`}
              onClick={() => this.setState({ viewMode: 'list' })}
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
