import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import { TextField } from '@fluentui/react/lib/TextField';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { IGraphUser, IOrgNode, PresenceAvailability } from '../../../../services/GraphService';
import { exportOrgChartToPdf, exportOrgChartToPng } from '../../../../services/PdfExportService';
import { IOrgChartProps, IOrgChartState } from './IOrgChartProps';
import styles from './OrgChart.module.scss';

/* ── Tree mutation helpers ───────────────── */

function cloneTree(node: IOrgNode): IOrgNode {
  return { ...node, directReports: node.directReports.map(cloneTree) };
}

function setNodeExpanded(root: IOrgNode, targetId: string, expanded: boolean): IOrgNode {
  if (root.user.id === targetId) return { ...root, isExpanded: expanded, directReports: root.directReports.map(cloneTree) };
  return { ...root, directReports: root.directReports.map(c => setNodeExpanded(c, targetId, expanded)) };
}

function injectChildren(root: IOrgNode, targetId: string, children: IOrgNode[]): IOrgNode {
  if (root.user.id === targetId) return { ...root, directReports: children, childrenLoaded: true, isExpanded: true };
  return { ...root, directReports: root.directReports.map(c => injectChildren(c, targetId, children)) };
}

function collapseAll(node: IOrgNode): IOrgNode {
  return { ...node, isExpanded: false, directReports: node.directReports.map(collapseAll) };
}

function expandLoaded(node: IOrgNode): IOrgNode {
  return { ...node, isExpanded: node.directReports.length > 0, directReports: node.directReports.map(expandLoaded) };
}

function markNodeLoaded(root: IOrgNode, targetId: string): IOrgNode {
  if (root.user.id === targetId) return { ...root, childrenLoaded: true, directReports: [] };
  return { ...root, directReports: root.directReports.map(c => markNodeLoaded(c, targetId)) };
}

function countSearchMatches(node: IOrgNode, q: string, isVisible: (u: IGraphUser) => boolean): number {
  if (!isVisible(node.user)) return 0;
  const self = matchesQuery(node, q) ? 1 : 0;
  return self + node.directReports.reduce((s, c) => s + countSearchMatches(c, q, isVisible), 0);
}

function matchesQuery(node: IOrgNode, lowerQ: string): boolean {
  if (!lowerQ) return false;
  const { displayName, jobTitle, department, mail } = node.user;
  return (
    (displayName || '').toLowerCase().includes(lowerQ) ||
    (jobTitle || '').toLowerCase().includes(lowerQ) ||
    (department || '').toLowerCase().includes(lowerQ) ||
    (mail || '').toLowerCase().includes(lowerQ)
  );
}

function matchUserQuery(user: IGraphUser, lowerQ: string): boolean {
  if (!lowerQ) return false;
  return (
    (user.displayName || '').toLowerCase().includes(lowerQ) ||
    (user.jobTitle || '').toLowerCase().includes(lowerQ) ||
    (user.department || '').toLowerCase().includes(lowerQ) ||
    (user.mail || '').toLowerCase().includes(lowerQ)
  );
}

function countTreeUsers(node: IOrgNode): { members: number; guests: number; disabled: number } {
  const c = { members: 0, guests: 0, disabled: 0 };
  const visit = (n: IOrgNode) => {
    const u = n.user;
    if (u.accountEnabled === false) c.disabled++;
    else if (u.userType === 'Guest') c.guests++;
    else c.members++;
    n.directReports.forEach(visit);
  };
  visit(node);
  return c;
}


function getInitials(displayName: string): string {
  const parts = (displayName || '').split(' ').filter(p => p.length > 0);
  if (parts.length === 0) return '?';
  if (parts.length === 1) return parts[0][0].toUpperCase();
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

function getUniqueDepts(node: IOrgNode): Map<string, number> {
  const map = new Map<string, number>();
  const visit = (n: IOrgNode) => {
    const dept = n.user.department || '';
    if (dept) map.set(dept, (map.get(dept) || 0) + 1);
    n.directReports.forEach(visit);
  };
  visit(node);
  return map;
}

function computeStats(node: IOrgNode, isVisible: (u: IGraphUser) => boolean): {
  total: number; members: number; guests: number;
  depts: number; avgSpan: number;
} {
  let total = 0, members = 0, guests = 0, managerCount = 0, totalReports = 0;
  const deptSet = new Set<string>();
  const visit = (n: IOrgNode) => {
    if (!isVisible(n.user)) return;
    total++;
    if (n.user.department) deptSet.add(n.user.department);
    if (n.user.userType === 'Guest') guests++;
    else members++;
    const visReports = n.directReports.filter(c => isVisible(c.user));
    if (visReports.length > 0) { managerCount++; totalReports += visReports.length; }
    n.directReports.forEach(visit);
  };
  visit(node);
  return { total, members, guests, depts: deptSet.size, avgSpan: managerCount ? Math.round(totalReports / managerCount * 10) / 10 : 0 };
}

/* ── Department color palettes & themes ──── */

import { OrgChartTheme } from '../ISmartOrgChartProps';
export type { OrgChartTheme };

function getSiteColor(theme: OrgChartTheme): string {
  if (theme === 'corporate') return '#0052a5';
  try {
    const t = (window as any).__themeState__?.theme;
    if (!t) return theme === 'dark' ? '#71afe5' : '#0078d4';
    return theme === 'dark' ? (t.themeTertiary ?? '#71afe5') : (t.themePrimary ?? '#0078d4');
  } catch {
    return theme === 'dark' ? '#71afe5' : '#0078d4';
  }
}

const THEME_CONTAINER_CLASS: Record<OrgChartTheme, string> = {
  modern:    '',
  minimal:   styles.themeMinimal,
  corporate: styles.themeCorporate,
  dark:      styles.themeDark,
};

/* ── Presence helpers ────────────────────── */

const PRESENCE_COLOUR: Record<PresenceAvailability, string> = {
  Available: '#6BB700', Busy: '#C50F1F', DoNotDisturb: '#C50F1F',
  BeRightBack: '#FFAA44', Away: '#FFAA44', Offline: '#8A8886', Unknown: '#8A8886',
};

const PRESENCE_LABEL: Record<PresenceAvailability, string> = {
  Available: 'Available', Busy: 'Busy', DoNotDisturb: 'Do Not Disturb',
  BeRightBack: 'Be Right Back', Away: 'Away', Offline: 'Offline', Unknown: '',
};

const PresenceDot: React.FC<{ status: PresenceAvailability | undefined }> = ({ status }) => {
  if (!status || status === 'Unknown') return null;
  return <span className={styles.presenceDot} style={{ background: PRESENCE_COLOUR[status] }} />;
};

/* ── Person Card (Outlook-style popup) ───── */

interface IPersonCardProps {
  user: IGraphUser;
  photo: string | null;
  presence: PresenceAvailability | undefined;
  theme: OrgChartTheme;
  managerChain: IGraphUser[];
  onClose: () => void;
  onFocus: (user: IGraphUser) => void;
}

const PersonCard: React.FC<IPersonCardProps> = ({ user, photo, presence, theme, managerChain, onClose, onFocus }) => {
  const deptColor = getSiteColor(theme);
  const isDark = theme === 'dark';
  const initials = getInitials(user.displayName);
  const phone = user.mobilePhone || (user.businessPhones && user.businessPhones[0]) || '';
  const hasPhone = !!phone;
  const isDisabled = user.accountEnabled === false;
  const isGuest = user.userType === 'Guest';

  const cardBg    = isDark ? '#242740' : '#ffffff';
  const textColor = isDark ? '#f0f0f0' : '#1a1a2e';
  const subColor  = isDark ? '#a0a8c0' : '#555';
  const fieldBg   = isDark ? '#1e2138' : '#f8f9fb';
  const borderClr = isDark ? '#3a3d5c' : '#e8ecf0';
  const chainBg   = isDark ? '#1a1c2e' : '#f2f4f8';

  return (
    <div className={styles.personCardOverlay} onClick={onClose}>
      <div
        className={styles.personCard}
        style={{ background: cardBg, borderColor: borderClr }}
        onClick={e => e.stopPropagation()}
      >
        {/* Coloured header band */}
        <div className={styles.personCardHeader} style={{ background: deptColor }}>
          <button className={styles.personCardClose} onClick={onClose} title="Close">
            <Icon iconName="Cancel" />
          </button>
          {photo
            ? <img src={photo} alt={user.displayName} className={styles.personCardPhoto} />
            : <div className={styles.personCardInitials}>{initials}</div>
          }
          {presence && presence !== 'Unknown' && (
            <div className={styles.personCardPresence}>
              <span className={styles.personCardPresenceDot} style={{ background: PRESENCE_COLOUR[presence] }} />
              <span>{PRESENCE_LABEL[presence]}</span>
            </div>
          )}
        </div>

        {/* Body */}
        <div className={styles.personCardBody}>
          <div className={styles.personCardName} style={{ color: textColor }}>{user.displayName}</div>
          {user.jobTitle && (
            <div className={styles.personCardTitle} style={{ color: deptColor }}>{user.jobTitle}</div>
          )}

          {/* Badges */}
          <div className={styles.personCardBadges}>
            {user.department && (
              <span className={styles.personCardDeptBadge} style={{ background: `${deptColor}1a`, color: deptColor }}>
                {user.department}
              </span>
            )}
            {isDisabled && <span className={styles.personCardStatusBadge} style={{ background: '#fde7e9', color: '#c50f1f' }}>Disabled</span>}
            {isGuest   && <span className={styles.personCardStatusBadge} style={{ background: '#fff4ce', color: '#835c00' }}>Guest</span>}
          </div>

          {/* Info fields */}
          <div className={styles.personCardFields} style={{ background: fieldBg, borderColor: borderClr }}>
            {user.mail && (
              <div className={styles.personCardField}>
                <Icon iconName="Mail" className={styles.personCardFieldIcon} style={{ color: deptColor }} />
                <a href={`mailto:${user.mail}`} className={styles.personCardFieldLink} style={{ color: deptColor }}>{user.mail}</a>
              </div>
            )}
            {user.businessPhones && user.businessPhones[0] && (
              <div className={styles.personCardField}>
                <Icon iconName="Phone" className={styles.personCardFieldIcon} style={{ color: subColor }} />
                <a href={`tel:${user.businessPhones[0]}`} className={styles.personCardFieldText} style={{ color: subColor }}>{user.businessPhones[0]}</a>
              </div>
            )}
            {user.mobilePhone && (
              <div className={styles.personCardField}>
                <Icon iconName="CellPhone" className={styles.personCardFieldIcon} style={{ color: subColor }} />
                <a href={`tel:${user.mobilePhone}`} className={styles.personCardFieldText} style={{ color: subColor }}>{user.mobilePhone}</a>
              </div>
            )}
            {user.officeLocation && (
              <div className={styles.personCardField}>
                <Icon iconName="POI" className={styles.personCardFieldIcon} style={{ color: subColor }} />
                <span className={styles.personCardFieldText} style={{ color: subColor }}>{user.officeLocation}</span>
              </div>
            )}
          </div>

          {/* Reporting chain */}
          {managerChain.length > 0 && (
            <div className={styles.personCardChain} style={{ background: chainBg, borderColor: borderClr }}>
              <div className={styles.personCardChainLabel} style={{ color: subColor }}>Reports to</div>
              <div className={styles.personCardChainItems}>
                {managerChain.map((mgr, i) => (
                  <React.Fragment key={mgr.id}>
                    {i > 0 && <Icon iconName="ChevronRight" className={styles.personCardChainSep} style={{ color: subColor }} />}
                    <button
                      className={styles.personCardChainChip}
                      onClick={() => { onClose(); onFocus(mgr); }}
                      title={`Focus on ${mgr.displayName}`}
                    >
                      <span
                        className={styles.personCardChainInitials}
                        style={{ background: getSiteColor(theme) }}
                      >
                        {getInitials(mgr.displayName)}
                      </span>
                      <span className={styles.personCardChainName} style={{ color: textColor }}>
                        {mgr.displayName.split(' ')[0]}
                      </span>
                    </button>
                  </React.Fragment>
                ))}
              </div>
            </div>
          )}

          {/* Action buttons */}
          {user.mail && (
            <div className={styles.personCardActions}>
              <a
                href={`https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(user.mail)}`}
                target="_blank" rel="noopener noreferrer"
                className={styles.personCardAction}
                style={{ background: deptColor, color: '#fff' }}
              >
                <Icon iconName="Chat" />&nbsp;Chat
              </a>
              <a
                href={`mailto:${user.mail}`}
                className={styles.personCardAction}
                style={{ background: isDark ? '#3a3d5c' : '#eef0f4', color: isDark ? '#e0e0f0' : '#333' }}
              >
                <Icon iconName="Mail" />&nbsp;Email
              </a>
              {hasPhone && (
                <a
                  href={`tel:${phone}`}
                  className={styles.personCardAction}
                  style={{ background: isDark ? '#3a3d5c' : '#eef0f4', color: isDark ? '#e0e0f0' : '#333' }}
                >
                  <Icon iconName="Phone" />&nbsp;Call
                </a>
              )}
              <button
                className={styles.personCardAction}
                style={{ background: isDark ? '#3a3d5c' : '#eef0f4', color: isDark ? '#e0e0f0' : '#333', border: 'none', cursor: 'pointer' }}
                onClick={() => { onClose(); onFocus(user); }}
                title="Focus org chart on this person"
              >
                <Icon iconName="Org" />&nbsp;Focus
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

/* ── Filter dropdown ─────────────────────── */

interface IFilterCounts { members: number; guests: number; disabled: number; }

interface IFilterPanelProps {
  filterMembers: boolean;
  filterGuests: boolean;
  counts: IFilterCounts;
  onToggle: (key: 'members' | 'guests') => void;
}

const FilterPanel: React.FC<IFilterPanelProps> = ({ filterMembers, filterGuests, counts, onToggle }) => {
  const items: Array<{ key: 'members' | 'guests'; label: string; count: number; checked: boolean }> = [
    { key: 'members',  label: 'Regular members',   count: counts.members,  checked: filterMembers  },
    { key: 'guests',   label: 'Guest users',        count: counts.guests,   checked: filterGuests   },
  ];
  return (
    <div className={styles.filterPanel}>
      <div className={styles.filterPanelTitle}>Show in chart</div>
      {items.map(({ key, label, count, checked }) => count >= 0 && (
        <label key={key} className={styles.filterItem}>
          <input type="checkbox" checked={checked} onChange={() => onToggle(key)} className={styles.filterCheckbox} />
          <span className={styles.filterLabel}>{label}</span>
          <span className={styles.filterCount}>{count}</span>
        </label>
      ))}
    </div>
  );
};

/* ── Node card ───────────────────────────── */

interface IOrgNodeCardProps {
  node: IOrgNode;
  photos: { [id: string]: string | null };
  presenceMap: Map<string, PresenceAvailability>;
  showDepartment: boolean;
  showOffice: boolean;
  isExpanding: boolean;
  searchQuery: string;
  theme: OrgChartTheme;
  directReportCount: number;
  managerUser?: IGraphUser;
  compactCards: boolean;
  onToggle: (node: IOrgNode) => void;
  onCardClick: (user: IGraphUser) => void;
  onFocus: (user: IGraphUser) => void;
}

const OrgNodeCard: React.FC<IOrgNodeCardProps> = ({
  node, photos, presenceMap, showDepartment, showOffice, isExpanding,
  searchQuery, theme, directReportCount, managerUser,
  compactCards, onToggle, onCardClick, onFocus
}) => {
  const { user } = node;
  const photo     = photos[user.id];
  const deptColor = getSiteColor(theme);
  const isDark    = theme === 'dark';
  const initials  = getInitials(user.displayName);
  const isRoot    = node.level === 0;

  const hasReports    = node.directReports.length > 0 || !node.childrenLoaded;
  const isHighlighted = searchQuery.trim() ? matchesQuery(node, searchQuery.toLowerCase()) : false;
  const isDisabled    = user.accountEnabled === false;
  const isGuest       = user.userType === 'Guest';

  const levelClass   = isRoot ? styles.rootCard : node.level === 1 ? styles.level1Card : '';
  const compactClass = compactCards ? styles.compactCard : '';
  const classes      = [styles.nodeCard, levelClass, compactClass, isHighlighted ? styles.highlightedCard : ''].filter(Boolean).join(' ');

  const borderStyle = theme === 'minimal'
    ? { borderLeft: `3px solid ${deptColor}`, borderTop: '1px solid #d8d8d8', opacity: isDisabled ? 0.55 : 1 }
    : { borderTopColor: deptColor, opacity: isDisabled ? 0.55 : 1 };

  const countLabel = directReportCount;

  return (
    <div
      className={classes}
      style={{ ...borderStyle, cursor: 'pointer', position: 'relative' }}
      title={`View ${user.displayName}'s profile`}
      onClick={() => onCardClick(user)}
    >
      {/* Focus button — shown on hover */}
      <button
        className={styles.focusBtn}
        onClick={e => { e.stopPropagation(); onFocus(user); }}
        title={`Focus org chart on ${user.displayName}`}
      >
        <Icon iconName="FitPage" />
      </button>

      <div className={styles.nodeAvatar}>
        {photo
          ? <img src={photo} alt={user.displayName} className={styles.photo} style={{ opacity: 0, transition: 'opacity 0.35s ease' }} onLoad={e => { (e.currentTarget as HTMLImageElement).style.opacity = '1'; }} />
          : <div className={styles.initials} style={{ background: deptColor }}>{initials}</div>
        }
        <PresenceDot status={presenceMap.get(user.id)} />
      </div>
      <div className={styles.nodeName} title={user.displayName}>{user.displayName}</div>
      {user.jobTitle && (
        <div className={styles.nodeTitle} title={user.jobTitle} style={{ color: deptColor }}>
          {user.jobTitle}
        </div>
      )}
      {showDepartment && user.department && !isGuest && !isDisabled && (
        <div className={styles.nodeDept}>{user.department}</div>
      )}
      {showOffice && user.officeLocation && !isGuest && !isDisabled && (
        <div className={styles.nodeOffice}>
          <Icon iconName="POI" className={styles.nodeOfficeIcon} />
          {user.officeLocation}
        </div>
      )}
      {!showDepartment && managerUser && !isGuest && !isDisabled && (
        <div className={styles.managerLine} style={{ color: isDark ? '#8090b0' : '#999' }}>
          ↑ {managerUser.displayName.split(' ')[0]}
        </div>
      )}
      {(isGuest || isDisabled) && (
        <div
          className={styles.nodeDept}
          style={{ background: isDisabled ? '#fde7e9' : '#fff4ce', color: isDisabled ? '#c50f1f' : '#835c00' }}
        >
          {isDisabled ? 'Disabled' : 'Guest'}
        </div>
      )}
      {hasReports && (
        <button
          className={styles.expandBtn}
          onClick={e => { e.stopPropagation(); onToggle(node); }}
          title={node.isExpanded ? 'Collapse' : 'Expand'}
          aria-expanded={node.isExpanded}
        >
          {isExpanding
            ? <Icon iconName="ProgressRingDots" className={styles.spinning} />
            : (
              <>
                <Icon iconName={node.isExpanded ? 'ChevronUp' : 'ChevronDown'} />
                {countLabel > 0 && <span className={styles.reportCount}>{countLabel}</span>}
              </>
            )
          }
        </button>
      )}
    </div>
  );
};

/* ── Recursive tree ──────────────────────── */

interface IOrgTreeProps {
  node: IOrgNode;
  photos: { [id: string]: string | null };
  presenceMap: Map<string, PresenceAvailability>;
  showDepartment: boolean;
  showOffice: boolean;
  chartLayout: ChartLayout;
  expandingNodes: Set<string>;
  searchQuery: string;
  theme: OrgChartTheme;
  isVisible: (user: IGraphUser) => boolean;
  parentUser?: IGraphUser;
  compactCards: boolean;
  onToggle: (node: IOrgNode) => void;
  onCardClick: (user: IGraphUser) => void;
  onFocus: (user: IGraphUser) => void;
}

const OrgTree: React.FC<IOrgTreeProps> = ({
  node, photos, presenceMap, showDepartment, showOffice, chartLayout,
  expandingNodes, searchQuery, theme, isVisible, parentUser, compactCards,
  onToggle, onCardClick, onFocus
}) => {
  if (!isVisible(node.user)) return null;

  const visibleReports     = node.directReports.filter(c => isVisible(c.user));
  const hasVisibleChildren = node.isExpanded && visibleReports.length > 0;

  return (
    <div className={`${styles.nodeWrapper} ${hasVisibleChildren ? styles.hasChildren : ''}`}>
      <OrgNodeCard
        node={node}
        photos={photos}
        presenceMap={presenceMap}
        showDepartment={showDepartment}
        showOffice={showOffice}
        isExpanding={expandingNodes.has(node.user.id)}
        searchQuery={searchQuery}
        theme={theme}
        directReportCount={visibleReports.length}
        managerUser={parentUser}
        compactCards={compactCards}
        onToggle={onToggle}
        onCardClick={onCardClick}
        onFocus={onFocus}
      />
      {hasVisibleChildren && (
        <div className={styles.children}>
          {visibleReports.map(child => (
            <OrgTree
              key={child.user.id}
              node={child}
              photos={photos}
              presenceMap={presenceMap}
              showDepartment={showDepartment}
              showOffice={showOffice}
              chartLayout={chartLayout}
              expandingNodes={expandingNodes}
              searchQuery={searchQuery}
              theme={theme}
              isVisible={isVisible}
              parentUser={node.user}
              compactCards={compactCards}
              onToggle={onToggle}
              onCardClick={onCardClick}
              onFocus={onFocus}
            />
          ))}
        </div>
      )}
    </div>
  );
};

/* ── No-config form ──────────────────────── */

const NoConfigForm: React.FC<{ onLoad: (id: string) => void }> = ({ onLoad }) => {
  const [val, setVal] = React.useState('');
  return (
    <div className={styles.noConfig}>
      <Icon iconName="Org" className={styles.noConfigIcon} />
      <div className={styles.noConfigTitle}>Set Up the Org Chart</div>
      <div className={styles.noConfigSubtitle}>Enter the top-level person&apos;s email or UPN to get started.</div>
      <div className={styles.noConfigForm}>
        <TextField placeholder="ceo@company.com" value={val} onChange={(_, v) => setVal(v || '')} className={styles.noConfigInput} />
        <PrimaryButton text="Load" onClick={() => val.trim() && onLoad(val.trim())} disabled={!val.trim()} />
      </div>
      <div className={styles.noConfigHint}>You can also set this permanently in the web part settings (admin).</div>
    </div>
  );
};

/* ── Main OrgChart component ─────────────── */

type ChartLayout = 'drill' | 'vertical' | 'horizontal';

interface IOrgChartLocalState extends IOrgChartState {
  searchQuery: string;
  presenceMap: Map<string, PresenceAvailability>;
  zoomLevel: number;
  selectedUser: IGraphUser | null;
  showFilters: boolean;
  isDragging: boolean;
  filterMembers: boolean;
  filterGuests: boolean;
  // Focus / navigation (full-tree mode)
  focusedUser: IGraphUser | null;
  ancestorChain: IGraphUser[];
  allUsers: IGraphUser[];
  showSearchResults: boolean;
  personCardManagerChain: IGraphUser[];
  // Layout
  chartLayout: ChartLayout;
  // Find Me feedback
  findMeError: string;
  // Tier 3
  filterDepartments: Set<string>;
  showDeptFilter: boolean;
  showStats: boolean;
  // Drill-down mode
  drillPath: IGraphUser[];
  drillReports: IGraphUser[];
  drillLoadingId: string | null;
  drillReportCounts: Map<string, number>;
  showLayoutPicker: boolean;
}

/* ── Chart state persistence ─────────────── */

const LS_CHART_KEY = 'smartOrgChart_chartState';

interface IChartStoredState {
  chartLayout?: ChartLayout;
  showStats?: boolean;
  filterMembers?: boolean;
  filterGuests?: boolean;
  filterDepartments?: string[];
  zoomLevel?: number;
  focusEmail?: string | null;
}

function loadChartState(): IChartStoredState {
  try {
    const s = localStorage.getItem(LS_CHART_KEY);
    if (s) return JSON.parse(s) as IChartStoredState;
  } catch { /* ignore */ }
  return {};
}

function saveChartState(s: IChartStoredState): void {
  try { localStorage.setItem(LS_CHART_KEY, JSON.stringify(s)); } catch { /* ignore */ }
}

const LAYOUT_CYCLE: ChartLayout[] = ['drill', 'vertical', 'horizontal'];

const LAYOUT_ICON: Record<ChartLayout, string> = {
  drill:      'Org',
  vertical:   'Down',
  horizontal: 'Forward',
};

const LAYOUT_TITLE: Record<ChartLayout, string> = {
  drill:      'Drill-down',
  vertical:   'Top-down tree',
  horizontal: 'Left-right tree',
};

export class OrgChart extends React.Component<IOrgChartProps, IOrgChartLocalState> {
  private _mounted          = false;
  private _pendingFocusEmail: string | null = null;
  private _presenceInterval: number | null = null;
  private _scrollRef        = React.createRef<HTMLDivElement>();
  private _drillViewRef     = React.createRef<HTMLDivElement>();
  private _searchRef        = React.createRef<HTMLDivElement>();
  private _isPanning        = false;
  private _requestedReportCounts = new Set<string>();
  private _panStartX        = 0;
  private _panStartY        = 0;
  private _scrollStartX     = 0;
  private _scrollStartY     = 0;
  private _panDistance      = 0;
  private _lastPanEndTime   = 0;

  constructor(props: IOrgChartProps) {
    super(props);
    const stored = loadChartState();
    this._pendingFocusEmail = stored.focusEmail ?? null;
    this.state = {
      rootNode: null, isLoading: false, error: null,
      photos: {}, expandingNodes: new Set(), searchQuery: '',
      presenceMap: new Map(), zoomLevel: stored.zoomLevel ?? 1,
      selectedUser: null,
      showFilters: false,
      filterMembers: stored.filterMembers ?? true,
      filterGuests: stored.filterGuests ?? true,
      isDragging: false,
      focusedUser: null, ancestorChain: [], allUsers: [],
      showSearchResults: false, personCardManagerChain: [],
      chartLayout: stored.chartLayout ?? (props.defaultLayout || 'drill'),
      findMeError: '',
      filterDepartments: new Set(stored.filterDepartments ?? []), showDeptFilter: false,
      showStats: stored.showStats ?? false,
      drillPath: [], drillReports: [], drillLoadingId: null,
      drillReportCounts: new Map(),
      showLayoutPicker: false,
    };
  }

  public async componentDidMount(): Promise<void> {
    this._mounted = true;
    if (this.props.graphService && this.props.topLevelUser) await this._loadTree();
    this._refreshPresence();
    this._presenceInterval = window.setInterval(() => { this._refreshPresence(); }, 60_000);

    if (this._scrollRef.current) {
      this._scrollRef.current.addEventListener('touchmove', this._handleTouchMoveDirect, { passive: false });
    }
    document.addEventListener('mousedown', this._handleOutsideClick);
  }

  public async componentDidUpdate(prev: IOrgChartProps, prevState: IOrgChartLocalState): Promise<void> {
    if (
      prev.topLevelUser !== this.props.topLevelUser ||
      prev.levelsBelow  !== this.props.levelsBelow  ||
      (!prev.graphService && this.props.graphService)
    ) {
      this._requestedReportCounts.clear();
      this.setState({
        rootNode: null, error: null, searchQuery: '',
        focusedUser: null, ancestorChain: [],
        drillPath: [], drillReports: [], drillLoadingId: null,
        drillReportCounts: new Map(),
      });
      if (this.props.graphService && this.props.topLevelUser) await this._loadTree();
    }

    const { isLoading, chartLayout, showStats, filterMembers, filterGuests,
            filterDepartments, zoomLevel, drillPath, focusedUser } = this.state;
    if (!isLoading && (
      prevState.chartLayout       !== chartLayout       ||
      prevState.showStats         !== showStats         ||
      prevState.filterMembers     !== filterMembers     ||
      prevState.filterGuests      !== filterGuests      ||
      prevState.filterDepartments !== filterDepartments ||
      prevState.zoomLevel         !== zoomLevel         ||
      prevState.drillPath         !== drillPath         ||
      prevState.focusedUser       !== focusedUser
    )) {
      const focusEmail = chartLayout === 'drill'
        ? (drillPath.length > 0 ? drillPath[drillPath.length - 1].mail ?? null : null)
        : (focusedUser?.mail ?? null);
      saveChartState({
        chartLayout, showStats, filterMembers, filterGuests,
        filterDepartments: Array.from(filterDepartments),
        zoomLevel,
        focusEmail,
      });
    }
  }

  public componentWillUnmount(): void {
    this._mounted = false;
    if (this._presenceInterval !== null) { window.clearInterval(this._presenceInterval); this._presenceInterval = null; }
    if (this._scrollRef.current) {
      this._scrollRef.current.removeEventListener('touchmove', this._handleTouchMoveDirect);
    }
    document.removeEventListener('mousedown', this._handleOutsideClick);
  }

  private _handleOutsideClick = (e: MouseEvent): void => {
    if (this._searchRef.current && !this._searchRef.current.contains(e.target as Node)) {
      if (this.state.showSearchResults) this.setState({ showSearchResults: false });
    }
  }

  private async _refreshPresence(): Promise<void> {
    const { graphService } = this.props;
    if (!graphService || !this._mounted) return;
    const presenceMap = await graphService.getPresence();
    if (this._mounted) this.setState({ presenceMap });
  }

  public exportPdf(): void {
    if (this.state.rootNode) exportOrgChartToPdf(this.state.rootNode);
  }

  public exportPng(): void {
    const el = this._drillViewRef.current ?? this._scrollRef.current;
    if (el) exportOrgChartToPng(el);
  }

  private async _loadTree(): Promise<void> {
    const { graphService, topLevelUser, levelsBelow } = this.props;
    if (!graphService) return;
    this.setState({ isLoading: true, error: null });
    try {
      const [rootUser, allUsers] = await Promise.all([
        graphService.findUser(topLevelUser),
        graphService.getAllUsers().catch(() => [] as IGraphUser[]),
      ]);
      if (!rootUser) {
        if (this._mounted) this.setState({ isLoading: false, error: `User "${topLevelUser}" not found. Check the UPN or email in Settings.` });
        return;
      }
      const rawRoot = await graphService.buildOrgTree(rootUser.id, levelsBelow);
      const rootNode = expandLoaded(rawRoot);
      const drillPath = [rootUser];
      const drillReports = rootNode.directReports.map(n => n.user);
      if (this._mounted) {
        this.setState({ rootNode, isLoading: false, allUsers, drillPath, drillReports, drillLoadingId: null }, () => {
          this._autoFitZoom();
        });
        this._loadPhotosForTree(rootNode);
        this._checkFrontierNodes(rootNode);
        this._loadDrillReportCounts(drillReports);

        // Restore last navigation position
        if (this._pendingFocusEmail) {
          const emailToRestore = this._pendingFocusEmail;
          this._pendingFocusEmail = null;
          if (emailToRestore.toLowerCase() !== (rootUser.mail || '').toLowerCase()) {
            const foundUser = allUsers.find(u => (u.mail || '').toLowerCase() === emailToRestore.toLowerCase())
              || await graphService.findUser(emailToRestore).catch(() => null);
            if (foundUser && this._mounted) await this._handleFocusUser(foundUser);
          }
        }
      }
    } catch {
      if (this._mounted) this.setState({ isLoading: false, error: 'Failed to load org chart. Check permissions.' });
    }
  }

  private _autoFitZoom(): void {
    requestAnimationFrame(() => {
      const container = this._scrollRef.current;
      if (!container || !this._mounted) return;
      const treeWrapper = container.firstElementChild as HTMLElement;
      if (!treeWrapper) return;
      const naturalW = treeWrapper.offsetWidth;
      const naturalH = treeWrapper.offsetHeight;
      if (!naturalW || !naturalH) return;
      const cW = container.clientWidth;
      const cH = container.clientHeight;
      if (naturalW <= cW && naturalH <= cH) return;
      const scaleX = cW / naturalW;
      const scaleY = cH / naturalH;
      const fitZoom = Math.min(scaleX, scaleY) * 0.90;
      const newZoom = Math.max(0.4, Math.min(0.95, fitZoom));
      if (newZoom < this.state.zoomLevel) {
        this.setState({ zoomLevel: newZoom });
      }
    });
  }

  private _loadPhotosForTree(node: IOrgNode): void {
    const ids: string[] = [];
    const collect = (n: IOrgNode) => { ids.push(n.user.id); n.directReports.forEach(collect); };
    collect(node);
    this._loadPhotos(ids);
  }

  private async _loadPhotos(ids: string[]): Promise<void> {
    const { graphService } = this.props;
    if (!graphService) return;
    for (const id of ids) {
      if (!this._mounted || id in this.state.photos) continue;
      const url = await graphService.getUserPhoto(id);
      if (this._mounted) this.setState(prev => ({ photos: { ...prev.photos, [id]: url } }));
    }
  }

  private async _checkFrontierNodes(rootNode: IOrgNode): Promise<void> {
    const { graphService } = this.props;
    if (!graphService) return;
    const frontier: IOrgNode[] = [];
    const collect = (n: IOrgNode) => { if (!n.childrenLoaded) { frontier.push(n); return; } n.directReports.forEach(collect); };
    collect(rootNode);
    const batchSize = 10;
    for (let i = 0; i < frontier.length; i += batchSize) {
      if (!this._mounted) return;
      await Promise.all(frontier.slice(i, i + batchSize).map(async node => {
        const hasReports = await graphService.hasDirectReports(node.user.id);
        if (!hasReports && this._mounted) {
          this.setState(prev => ({ rootNode: prev.rootNode ? markNodeLoaded(prev.rootNode, node.user.id) : null }));
        }
      }));
    }
  }

  /* ── Toggle expand (full-tree mode) ── */

  private _handleToggle = async (node: IOrgNode): Promise<void> => {
    const { rootNode, expandingNodes } = this.state;
    if (!rootNode) return;
    if (node.isExpanded) { this.setState({ rootNode: setNodeExpanded(rootNode, node.user.id, false) }); return; }
    if (!node.childrenLoaded && this.props.graphService) {
      const s1 = new Set(expandingNodes); s1.add(node.user.id);
      this.setState({ expandingNodes: s1 });
      try {
        const reports  = await this.props.graphService.getDirectReports(node.user.id);
        const children = reports.map(u => ({ user: u, directReports: [], isExpanded: false, childrenLoaded: false, level: node.level + 1 }));
        if (this._mounted && this.state.rootNode) {
          const s2      = new Set(this.state.expandingNodes); s2.delete(node.user.id);
          const newRoot = injectChildren(this.state.rootNode, node.user.id, children);
          this.setState({ rootNode: newRoot, expandingNodes: s2 });
          this._loadPhotos(children.map(c => c.user.id));
          this._checkFrontierNodes(newRoot);
        }
      } catch {
        if (this._mounted) { const s2 = new Set(this.state.expandingNodes); s2.delete(node.user.id); this.setState({ expandingNodes: s2 }); }
      }
    } else {
      this.setState({ rootNode: setNodeExpanded(rootNode, node.user.id, true) });
    }
  }

  private _handleCollapseAll = (): void => { if (this.state.rootNode) this.setState({ rootNode: collapseAll(this.state.rootNode) }); }

  // Expand All: first expand already-loaded nodes for immediate feedback,
  // then BFS-load every unloaded frontier level until the full tree is in memory.
  private _handleExpandLoaded = async (): Promise<void> => {
    if (!this.state.rootNode) return;
    let root = expandLoaded(this.state.rootNode);
    this.setState({ rootNode: root });

    const gs = this.props.graphService;
    if (!gs) return;

    while (this._mounted) {
      const frontier = this._getUnloadedFrontier(root);
      if (frontier.length === 0) break;

      const BATCH = 8;
      for (let i = 0; i < frontier.length; i += BATCH) {
        if (!this._mounted) return;
        const slice = frontier.slice(i, i + BATCH);
        const results = await Promise.all(
          slice.map(n => gs.getDirectReports(n.user.id)
            .then(r  => ({ node: n, reports: r }))
            .catch(() => ({ node: n, reports: [] as IGraphUser[] }))
          )
        );
        const newIds: string[] = [];
        for (const { node: n, reports } of results) {
          const children: IOrgNode[] = reports.map(u => ({
            user: u, directReports: [], isExpanded: false, childrenLoaded: false, level: n.level + 1,
          }));
          root = injectChildren(root, n.user.id, children);
          newIds.push(...reports.map(u => u.id));
        }
        root = expandLoaded(root);
        if (this._mounted) {
          this.setState({ rootNode: root });
          if (newIds.length) this._loadPhotos(newIds);
        }
      }
    }
  }

  private _getUnloadedFrontier = (node: IOrgNode): IOrgNode[] => {
    const result: IOrgNode[] = [];
    const visit = (n: IOrgNode): void => {
      if (!n.childrenLoaded) result.push(n);
      else n.directReports.forEach(visit);
    };
    visit(node);
    return result;
  }

  /* ── Card click → profile popup ── */

  private _handleCardClick = (user: IGraphUser): void => {
    if (Date.now() - this._lastPanEndTime < 150) return;
    this.setState({ selectedUser: user, showFilters: false, personCardManagerChain: [] });
    if (this.props.graphService) {
      this.props.graphService.getManagerChain(user.id, 8).then(chain => {
        if (this._mounted) this.setState({ personCardManagerChain: chain });
      }).catch(() => { /* ignore */ });
    }
  }

  /* ── Drill-down handlers ── */

  private _handleDrillInto = async (user: IGraphUser): Promise<void> => {
    const { graphService } = this.props;
    if (!graphService) return;
    this.setState({ drillLoadingId: user.id });
    try {
      const reports = await graphService.getDirectReports(user.id);
      if (!this._mounted) return;
      if (reports.length === 0) {
        // No reports — open profile popup instead of drilling into a dead end
        this.setState({ drillLoadingId: null, selectedUser: user, personCardManagerChain: [] });
        graphService.getManagerChain(user.id, 8).then(chain => {
          if (this._mounted) this.setState({ personCardManagerChain: chain });
        }).catch(_err => { /* ignore */ });
        return;
      }
      this.setState(prev => ({
        drillPath: [...prev.drillPath, user],
        drillReports: reports,
        drillLoadingId: null,
      }));
      this._loadPhotos([user.id, ...reports.map(u => u.id)]);
      this._loadDrillReportCounts(reports);
    } catch {
      if (this._mounted) this.setState({ drillLoadingId: null });
    }
  }

  private _handleDrillNavigate = async (index: number): Promise<void> => {
    const { graphService } = this.props;
    const { drillPath } = this.state;
    if (!graphService || index >= drillPath.length) return;
    if (index === drillPath.length - 1) return; // already at this level
    const targetUser = drillPath[index];
    this.setState({ drillLoadingId: targetUser.id });
    try {
      const reports = await graphService.getDirectReports(targetUser.id);
      if (!this._mounted) return;
      this.setState({
        drillPath: drillPath.slice(0, index + 1),
        drillReports: reports,
        drillLoadingId: null,
      });
      this._loadPhotos(reports.map(u => u.id));
      this._loadDrillReportCounts(reports);
    } catch {
      if (this._mounted) this.setState({ drillLoadingId: null });
    }
  }

  /* ── Focus on person (full-tree mode or drill mode) ── */

  private _handleFocusUser = async (user: IGraphUser): Promise<void> => {
    const { graphService, levelsBelow } = this.props;
    if (!graphService) return;
    const { chartLayout } = this.state;

    this.setState({ isLoading: true, error: null, searchQuery: '', showSearchResults: false });

    if (chartLayout === 'drill') {
      try {
        const [reports, managerChain] = await Promise.all([
          graphService.getDirectReports(user.id),
          graphService.getManagerChain(user.id, 10),
        ]);
        if (!this._mounted) return;
        this.setState({
          drillPath: [...managerChain, user],
          drillReports: reports,
          isLoading: false,
          drillLoadingId: null,
        });
        this._loadPhotos([user.id, ...managerChain.map(u2 => u2.id), ...reports.map(u2 => u2.id)]);
        this._loadDrillReportCounts(reports);
      } catch {
        if (this._mounted) this.setState({ isLoading: false });
      }
      return;
    }

    // Full-tree mode
    try {
      const [rootNode, ancestorChain] = await Promise.all([
        graphService.buildOrgTree(user.id, levelsBelow),
        graphService.getManagerChain(user.id, 10),
      ]);
      if (this._mounted) {
        const expanded = expandLoaded(rootNode);
        this.setState({ rootNode: expanded, focusedUser: user, ancestorChain, isLoading: false }, () => {
          if (this._scrollRef.current) {
            this._scrollRef.current.scrollLeft = 0;
            this._scrollRef.current.scrollTop  = 0;
          }
          this._autoFitZoom();
        });
        this._loadPhotosForTree(expanded);
        this._checkFrontierNodes(expanded);
      }
    } catch {
      if (this._mounted) this.setState({ isLoading: false, error: 'Failed to load org chart for this person.' });
    }
  }

  private _handleReturnToRoot = (): void => {
    const { chartLayout, rootNode } = this.state;
    if (chartLayout === 'drill' && rootNode) {
      const rootReports = rootNode.directReports.map(n => n.user);
      this.setState({
        drillPath: [rootNode.user],
        drillReports: rootReports,
        focusedUser: null,
        ancestorChain: [],
        drillLoadingId: null,
      });
      this._loadDrillReportCounts(rootReports);
      return;
    }
    this.setState({ focusedUser: null, ancestorChain: [], rootNode: null, error: null });
    this._loadTree();
  }

  /* ── Background-fetch direct report counts for drill cards ── */

  private _loadDrillReportCounts(users: IGraphUser[]): void {
    const { graphService } = this.props;
    if (!graphService) return;
    users.forEach(user => {
      if (this._requestedReportCounts.has(user.id)) return;
      this._requestedReportCounts.add(user.id);
      graphService.getDirectReports(user.id).then(reports => {
        if (!this._mounted) return;
        this.setState(prev => {
          const next = new Map(prev.drillReportCounts);
          next.set(user.id, reports.length);
          return { drillReportCounts: next };
        });
      }).catch(() => {
        this._requestedReportCounts.delete(user.id);
      });
    });
  }

  /* ── Find Me ── */

  private _handleFindMe = async (): Promise<void> => {
    const { graphService, currentUserEmail } = this.props;
    if (!graphService || !currentUserEmail) return;
    const user = await graphService.findUser(currentUserEmail).catch(() => null);
    if (user) {
      await this._handleFocusUser(user);
    } else {
      this.setState({ findMeError: 'Your account was not found in this org.' });
      setTimeout(() => { if (this._mounted) this.setState({ findMeError: '' }); }, 3000);
    }
  }

  /* ── Layout picker ── */

  private _setLayout = (layout: ChartLayout): void => {
    const { rootNode, drillPath } = this.state;
    const nextLayout = layout;
    const updates: Partial<IOrgChartLocalState> = { chartLayout: nextLayout };
    if (nextLayout === 'drill' && drillPath.length === 0 && rootNode) {
      const rootReports = rootNode.directReports.map(n => n.user);
      updates.drillPath = [rootNode.user];
      updates.drillReports = rootReports;
      this._loadDrillReportCounts(rootReports);
    }
    this.setState(updates as IOrgChartLocalState);
  }

  /* ── Drag-to-pan (mouse) ── */

  private _handlePanStart = (e: React.MouseEvent<HTMLDivElement>): void => {
    if ((e.target as HTMLElement).closest('button, a, input')) return;
    this._isPanning    = true;
    this._panDistance  = 0;
    this._panStartX    = e.clientX;
    this._panStartY    = e.clientY;
    this._scrollStartX = this._scrollRef.current?.scrollLeft ?? 0;
    this._scrollStartY = this._scrollRef.current?.scrollTop  ?? 0;
    this.setState({ isDragging: true });
    e.preventDefault();
  }

  private _handlePanMove = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (!this._isPanning || !this._scrollRef.current) return;
    const dx = e.clientX - this._panStartX;
    const dy = e.clientY - this._panStartY;
    this._panDistance = Math.sqrt(dx * dx + dy * dy);
    this._scrollRef.current.scrollLeft = this._scrollStartX - dx;
    this._scrollRef.current.scrollTop  = this._scrollStartY - dy;
  }

  private _handlePanEnd = (): void => {
    if (!this._isPanning) return;
    if (this._panDistance > 8) this._lastPanEndTime = Date.now();
    this._isPanning = false;
    this.setState({ isDragging: false });
  }

  /* ── Drag-to-pan (touch) ── */

  private _handleTouchStart = (e: React.TouchEvent<HTMLDivElement>): void => {
    if ((e.target as HTMLElement).closest('button, a, input')) return;
    const touch = e.touches[0];
    this._isPanning    = true;
    this._panDistance  = 0;
    this._panStartX    = touch.clientX;
    this._panStartY    = touch.clientY;
    this._scrollStartX = this._scrollRef.current?.scrollLeft ?? 0;
    this._scrollStartY = this._scrollRef.current?.scrollTop  ?? 0;
    this.setState({ isDragging: true });
  }

  private _handleTouchMoveDirect = (e: TouchEvent): void => {
    if (!this._isPanning || !this._scrollRef.current) return;
    e.preventDefault();
    const touch = e.touches[0];
    const dx = touch.clientX - this._panStartX;
    const dy = touch.clientY - this._panStartY;
    this._panDistance = Math.sqrt(dx * dx + dy * dy);
    this._scrollRef.current.scrollLeft = this._scrollStartX - dx;
    this._scrollRef.current.scrollTop  = this._scrollStartY - dy;
  }

  private _handleTouchEnd = (): void => {
    if (!this._isPanning) return;
    if (this._panDistance > 8) this._lastPanEndTime = Date.now();
    this._isPanning = false;
    this.setState({ isDragging: false });
  }

  /* ── Filter ── */

  private _buildIsVisible = (): (user: IGraphUser) => boolean => {
    const { filterMembers, filterGuests, filterDepartments } = this.state;
    return (user: IGraphUser) => {
      if (user.userType === 'Guest'     && !filterGuests)   return false;
      if (user.accountEnabled !== false && user.userType !== 'Guest' && !filterMembers) return false;
      if (filterDepartments.size > 0 && !filterDepartments.has(user.department || '')) return false;
      return true;
    };
  }

  /* ── Drill view renderer ── */

  private _renderDrillView(): React.ReactElement {
    const {
      drillPath, drillReports, drillLoadingId,
      photos, presenceMap, drillReportCounts,
    } = this.state;
    const { theme, showDepartment, compactCards } = this.props;
    const isVisible = this._buildIsVisible();
    const currentUser = drillPath.length > 0 ? drillPath[drillPath.length - 1] : null;
    const visibleReports = drillReports.filter(u => isVisible(u));
    const isDark = theme === 'dark';
    const headerBg    = isDark ? '#1e2138' : '#ffffff';
    const headerBorder = isDark ? '#3a3d5c' : '#e8ecf0';
    const nameColor   = isDark ? '#e8ecff' : '#1a1a2e';
    const deptColor2  = isDark ? '#9098b8' : '#5a6472';

    return (
      <div className={styles.drillView} ref={this._drillViewRef}>

        {/* Breadcrumb nav — only shown when drilled deeper than root */}
        {drillPath.length > 1 && (
          <div className={styles.drillNav}>
            <button
              className={styles.drillNavHomeBtn}
              onClick={this._handleReturnToRoot}
              title="Back to top"
            >
              <Icon iconName="Home" />
            </button>
            <Icon iconName="ChevronRight" className={styles.drillNavChevron} />
            {drillPath.slice(0, -1).map((person, i) => (
              <React.Fragment key={person.id}>
                <button
                  className={styles.drillNavItem}
                  onClick={() => this._handleDrillNavigate(i)}
                  title={`Go back to ${person.displayName}`}
                >
                  <span
                    className={styles.drillNavInitials}
                    style={{ background: getSiteColor(theme) }}
                  >
                    {getInitials(person.displayName)}
                  </span>
                  <span className={styles.drillNavName}>{person.displayName.split(' ')[0]}</span>
                </button>
                <Icon iconName="ChevronRight" className={styles.drillNavChevron} />
              </React.Fragment>
            ))}
            {currentUser && (
              <span className={styles.drillNavCurrent}>
                <span
                  className={styles.drillNavInitials}
                  style={{ background: getSiteColor(theme) }}
                >
                  {getInitials(currentUser.displayName)}
                </span>
                <span className={styles.drillNavName}>{currentUser.displayName}</span>
              </span>
            )}
          </div>
        )}

        {/* Current person header */}
        {currentUser && (
          <div
            className={styles.drillCurrentHeader}
            style={{
              background: headerBg,
              borderBottom: `3px solid ${getSiteColor(theme)}`,
              borderTop: `1px solid ${headerBorder}`,
            }}
          >
            <div className={styles.drillCurrentAvatar} style={{ position: 'relative' }}>
              {photos[currentUser.id]
                ? <img
                    src={photos[currentUser.id] as string}
                    alt={currentUser.displayName}
                    className={styles.drillCurrentAvatarImg}
                    style={{ opacity: 0, transition: 'opacity 0.35s ease' }}
                    onLoad={e => { (e.currentTarget as HTMLImageElement).style.opacity = '1'; }}
                  />
                : <div
                    className={styles.drillCurrentAvatarInitials}
                    style={{ background: getSiteColor(theme) }}
                  >
                    {getInitials(currentUser.displayName)}
                  </div>
              }
              <PresenceDot status={presenceMap.get(currentUser.id)} />
            </div>
            <div className={styles.drillCurrentInfo}>
              <div className={styles.drillCurrentName} style={{ color: nameColor }}>
                {currentUser.displayName}
              </div>
              {currentUser.jobTitle && (
                <div
                  className={styles.drillCurrentTitle}
                  style={{ color: getSiteColor(theme) }}
                >
                  {currentUser.jobTitle}
                </div>
              )}
              {showDepartment && currentUser.department && (
                <div className={styles.drillCurrentDept} style={{ color: deptColor2 }}>
                  {currentUser.department}
                </div>
              )}
            </div>
            <button
              className={styles.drillCurrentProfileBtn}
              onClick={() => this._handleCardClick(currentUser)}
              title="View profile"
            >
              <Icon iconName="Contact" />
              <span>Profile</span>
            </button>
          </div>
        )}

        {/* Direct reports grid */}
        <div className={styles.drillBody}>
          {drillLoadingId && !drillReports.find(u => u.id === drillLoadingId) ? (
            <div className={styles.drillSpinner}>
              <Spinner size={SpinnerSize.medium} label="Loading..." />
            </div>
          ) : visibleReports.length > 0 ? (
            <>
              <div className={styles.drillSectionTitle}>
                Direct Reports &nbsp;
                <span className={styles.drillSectionCount}>{visibleReports.length}</span>
              </div>
              <div className={`${styles.drillReportsGrid} ${compactCards ? styles.compactMode : ''}`}>
                {visibleReports.map(report => {
                  const count = drillReportCounts.get(report.id);
                  const countKnown = count !== undefined;
                  const fakeNode: IOrgNode = {
                    user: report,
                    directReports: [],
                    isExpanded: false,
                    childrenLoaded: countKnown && count === 0,
                    level: 1,
                  };
                  return (
                    <OrgNodeCard
                      key={report.id}
                      node={fakeNode}
                      photos={photos}
                      presenceMap={presenceMap}
                      showDepartment={showDepartment}
                      showOffice={this.props.showOffice}
                      isExpanding={drillLoadingId === report.id}
                      searchQuery=""
                      theme={theme}
                      directReportCount={count ?? 0}
                      compactCards={compactCards}
                      onToggle={node => this._handleDrillInto(node.user)}
                      onCardClick={this._handleDrillInto}
                      onFocus={this._handleCardClick}
                    />
                  );
                })}
              </div>
            </>
          ) : (
            <div className={styles.drillNoReports}>
              No direct reports
            </div>
          )}
        </div>
      </div>
    );
  }

  /* ── No-config handler ── */

  private _renderNoConfig(): React.ReactElement {
    return (
      <NoConfigForm onLoad={identifier => {
        const gs = this.props.graphService;
        if (!gs) return;
        this.setState({ isLoading: true, error: null });
        gs.findUser(identifier).then(user => {
          if (!user) { if (this._mounted) this.setState({ isLoading: false, error: `User "${identifier}" not found.` }); return; }
          gs.buildOrgTree(user.id, this.props.levelsBelow).then(rootNode => {
            if (this._mounted) { this.setState({ rootNode, isLoading: false }); this._loadPhotosForTree(rootNode); }
          }).catch(() => { if (this._mounted) this.setState({ isLoading: false, error: 'Failed to load org chart.' }); });
        }).catch(() => { if (this._mounted) this.setState({ isLoading: false, error: 'Failed to search for user.' }); });
      }} />
    );
  }

  /* ── Render ── */

  public render(): React.ReactElement {
    const {
      rootNode, isLoading, error, photos, expandingNodes, searchQuery,
      presenceMap, zoomLevel, selectedUser, showFilters,
      filterMembers, filterGuests, isDragging,
      focusedUser, ancestorChain, allUsers, showSearchResults,
      personCardManagerChain, chartLayout, findMeError,
      filterDepartments, showDeptFilter, showStats, showLayoutPicker,
    } = this.state;
    const { showDepartment, theme, currentUserEmail, compactCards,
      enableFindMe, enableLayoutToggle, enableStats, enableDeptFilter, enableUserFilter } = this.props;

    if (isLoading) return (
      <div className={styles.centered}><Spinner size={SpinnerSize.large} label="Building org chart..." /></div>
    );

    if (error) return (
      <div className={styles.errorState}>
        <Icon iconName="Warning" className={styles.errorIcon} />
        <div className={styles.errorText}>{error}</div>
        <DefaultButton text="Retry" onClick={() => this._loadTree()} />
      </div>
    );

    if (!rootNode) return this._renderNoConfig();

    const isVisible    = this._buildIsVisible();
    const treeCounts   = countTreeUsers(rootNode);
    const lowerQ       = searchQuery.trim().toLowerCase();
    const matchCount   = lowerQ ? countSearchMatches(rootNode, lowerQ, isVisible) : 0;
    const activeFilters = (!filterMembers ? 1 : 0) + (!filterGuests ? 1 : 0);
    const uniqueDepts  = getUniqueDepts(rootNode);
    const stats        = showStats ? computeStats(rootNode, isVisible) : null;
    const isDrillMode  = chartLayout === 'drill';

    const searchResults = lowerQ && allUsers.length > 0
      ? allUsers.filter(u => matchUserQuery(u, lowerQ)).slice(0, 8)
      : [];

    const containerClasses = [
      styles.container,
      THEME_CONTAINER_CLASS[theme],
    ].filter(Boolean).join(' ');

    const treeScrollClasses = [
      styles.treeScroll,
      isDragging ? styles.treeScrollPanning : '',
      chartLayout === 'horizontal' ? styles.layoutHorizontal : '',
      compactCards ? styles.compactMode : '',
    ].filter(Boolean).join(' ');

    return (
      <div className={containerClasses}>

        {/* ── Toolbar ── */}
        <div className={styles.chartToolbar}>

          {/* Search with results dropdown */}
          <div className={styles.searchWrapper} ref={this._searchRef}>
            <SearchBox
              placeholder="Search people..."
              value={searchQuery}
              onChange={(_, v) => this.setState({ searchQuery: v || '', showSearchResults: !!(v && v.trim()), showFilters: false })}
              onFocus={() => { if (searchQuery.trim()) this.setState({ showSearchResults: true }); }}
              className={styles.chartSearch}
              underlined
            />
            {lowerQ && !showSearchResults && !isDrillMode && (
              <span className={styles.chartSearchHit}>{matchCount} {matchCount === 1 ? 'match' : 'matches'} in tree</span>
            )}
            {showSearchResults && searchResults.length > 0 && (
              <div className={styles.searchResults}>
                {searchResults.map(u => {
                  const color = getSiteColor(theme);
                  return (
                    <button
                      key={u.id}
                      className={styles.searchResult}
                      onClick={() => { this._handleFocusUser(u); this.setState({ searchQuery: '' }); }}
                    >
                      <span className={styles.searchResultInitials} style={{ background: color }}>
                        {getInitials(u.displayName)}
                      </span>
                      <span className={styles.searchResultInfo}>
                        <span className={styles.searchResultName}>{u.displayName}</span>
                        <span className={styles.searchResultMeta}>{[u.jobTitle, u.department].filter(Boolean).join(' · ')}</span>
                      </span>
                    </button>
                  );
                })}
              </div>
            )}
          </div>

          {/* Expand/Collapse — only in full-tree mode */}
          {!isDrillMode && (
            <div className={styles.chartActions}>
              <button className={styles.chartActionBtn} onClick={this._handleExpandLoaded}>
                <Icon iconName="ExploreContent" /> Expand All
              </button>
              <button className={styles.chartActionBtn} onClick={this._handleCollapseAll}>
                <Icon iconName="CollapseContent" /> Collapse All
              </button>
            </div>
          )}

          {/* Find Me */}
          {currentUserEmail && enableFindMe && (
            <button
              className={styles.iconToolBtn}
              onClick={this._handleFindMe}
              title="Find me in the org chart"
            >
              <Icon iconName="Contact" />
            </button>
          )}

          {/* Layout picker */}
          {enableLayoutToggle && (
            <div className={styles.toolbarPopupAnchor}>
              <button
                className={`${styles.chartActionBtn} ${showLayoutPicker ? styles.iconToolBtnActive : ''}`}
                onClick={() => this.setState(p => ({ showLayoutPicker: !p.showLayoutPicker, showFilters: false, showDeptFilter: false }))}
                title="Switch view layout"
              >
                <Icon iconName="ViewAll" />
                <span>View</span>
              </button>
              {showLayoutPicker && (
                <div className={styles.filterPanel} style={{ minWidth: 210 }}>
                  <div className={styles.filterPanelTitle}>View layout</div>
                  {LAYOUT_CYCLE.map(layout => (
                    <button
                      key={layout}
                      className={styles.filterItem}
                      style={{
                        border: 'none',
                        background: chartLayout === layout ? '#e8f4fd' : 'transparent',
                        cursor: 'pointer',
                        width: '100%',
                        textAlign: 'left',
                        display: 'flex',
                        alignItems: 'center',
                        gap: 8,
                        padding: '8px 10px',
                        borderRadius: 4,
                        fontWeight: chartLayout === layout ? 600 : 400,
                        color: chartLayout === layout ? '#0078d4' : 'inherit',
                      }}
                      onClick={() => { this._setLayout(layout); this.setState({ showLayoutPicker: false }); }}
                    >
                      <Icon iconName={LAYOUT_ICON[layout]} />
                      <span style={{ flex: 1 }}>{LAYOUT_TITLE[layout]}</span>
                      {chartLayout === layout && <Icon iconName="CheckMark" />}
                    </button>
                  ))}
                </div>
              )}
            </div>
          )}

          {/* Stats toggle */}
          {enableStats && (
            <button
              className={`${styles.iconToolBtn} ${showStats ? styles.iconToolBtnActive : ''}`}
              onClick={() => this.setState(p => ({ showStats: !p.showStats }))}
              title="Org stats summary"
            >
              <Icon iconName="BarChartVertical" />
            </button>
          )}

          {/* Dept filter button */}
          {enableDeptFilter && <div className={styles.toolbarPopupAnchor}>
            <button
              className={`${styles.iconToolBtn} ${filterDepartments.size > 0 ? styles.iconToolBtnActive : ''}`}
              onClick={() => this.setState(p => ({ showDeptFilter: !p.showDeptFilter, showFilters: false }))}
              title="Filter by department"
            >
              <Icon iconName="DeveloperTools" />
              {filterDepartments.size > 0 && <span className={styles.toolBtnBadge}>{filterDepartments.size}</span>}
            </button>
            {showDeptFilter && (
              <div className={styles.filterPanel} style={{ minWidth: 220 }}>
                <div className={styles.filterPanelTitle}>Filter by department</div>
                {Array.from(uniqueDepts.entries()).sort((a, b) => a[0].localeCompare(b[0])).map(([dept, count]) => (
                  <label key={dept} className={styles.filterItem}>
                    <input
                      type="checkbox"
                      className={styles.filterCheckbox}
                      checked={filterDepartments.has(dept)}
                      onChange={() => {
                        const next = new Set(filterDepartments);
                        next.has(dept) ? next.delete(dept) : next.add(dept);
                        this.setState({ filterDepartments: next });
                      }}
                    />
                    <span className={styles.filterLabel}>{dept}</span>
                    <span className={styles.filterCount}>{count}</span>
                  </label>
                ))}
                {filterDepartments.size > 0 && (
                  <button
                    className={styles.filterItem}
                    style={{ border: 'none', background: 'none', cursor: 'pointer', color: '#0078d4', fontWeight: 600, fontSize: 12 }}
                    onClick={() => this.setState({ filterDepartments: new Set() })}
                  >
                    Clear all filters
                  </button>
                )}
              </div>
            )}
          </div>}

          {/* Filter button */}
          {enableUserFilter && (
            <div className={styles.toolbarPopupAnchor}>
              <button
                className={`${styles.iconToolBtn} ${activeFilters > 0 ? styles.iconToolBtnActive : ''}`}
                onClick={() => this.setState(p => ({ showFilters: !p.showFilters }))}
                title="Filter user types"
              >
                <Icon iconName="Filter" />
                {activeFilters > 0 && <span className={styles.toolBtnBadge}>{activeFilters}</span>}
              </button>
              {showFilters && (
                <FilterPanel
                  filterMembers={filterMembers}
                  filterGuests={filterGuests}
                  counts={treeCounts}
                  onToggle={key => this.setState(p => ({
                    filterMembers: key === 'members' ? !p.filterMembers : p.filterMembers,
                    filterGuests:  key === 'guests'  ? !p.filterGuests  : p.filterGuests,
                  }))}
                />
              )}
            </div>
          )}

          {/* Zoom — only in full-tree mode */}
          {!isDrillMode && (
            <div className={styles.zoomControls}>
              <button className={styles.zoomBtn} onClick={() => this.setState({ zoomLevel: Math.max(0.4, zoomLevel - 0.1) })} title="Zoom out" disabled={zoomLevel <= 0.4}>
                <Icon iconName="Remove" />
              </button>
              <span className={styles.zoomLabel}>{Math.round(zoomLevel * 100)}%</span>
              <button className={styles.zoomBtn} onClick={() => this.setState({ zoomLevel: Math.min(1.5, zoomLevel + 0.1) })} title="Zoom in" disabled={zoomLevel >= 1.5}>
                <Icon iconName="Add" />
              </button>
              <button className={styles.zoomBtn} onClick={() => this.setState({ zoomLevel: 1 })} title="Reset zoom" disabled={zoomLevel === 1}>
                <Icon iconName="Refresh" />
              </button>
            </div>
          )}
        </div>

        {/* ── Find Me error toast ── */}
        {findMeError && (
          <div className={styles.findMeToast}>{findMeError}</div>
        )}

        {/* ── Stats bar ── */}
        {stats && (
          <div className={styles.statsBar}>
            <div className={styles.statItem}><span className={styles.statValue}>{stats.total}</span><span className={styles.statLabel}>People</span></div>
            <div className={styles.statItem}><span className={styles.statValue}>{stats.members}</span><span className={styles.statLabel}>Members</span></div>
            {stats.guests > 0 && <div className={styles.statItem}><span className={styles.statValue}>{stats.guests}</span><span className={styles.statLabel}>Guests</span></div>}
            <div className={styles.statItem}><span className={styles.statValue}>{stats.depts}</span><span className={styles.statLabel}>Depts</span></div>
            <div className={styles.statItem}><span className={styles.statValue}>{stats.avgSpan}</span><span className={styles.statLabel}>Avg span</span></div>
          </div>
        )}

        {/* ── Ancestor strip (full-tree mode only) ── */}
        {!isDrillMode && focusedUser && (
          <div className={styles.ancestorStrip}>
            <button className={styles.ancestorReturnBtn} onClick={this._handleReturnToRoot} title="Back to full org chart">
              <Icon iconName="Home" /> Full org
            </button>
            <Icon iconName="ChevronRight" className={styles.ancestorChevron} />
            {ancestorChain.map((ancestor, i) => (
              <React.Fragment key={ancestor.id}>
                <button
                  className={styles.ancestorLink}
                  onClick={() => this._handleFocusUser(ancestor)}
                  title={`Focus on ${ancestor.displayName}`}
                >
                  <span className={styles.ancestorInitials} style={{ background: getSiteColor(theme) }}>
                    {getInitials(ancestor.displayName)}
                  </span>
                  <span className={styles.ancestorName}>{ancestor.displayName.split(' ')[0]}</span>
                </button>
                {i < ancestorChain.length && <Icon iconName="ChevronRight" className={styles.ancestorChevron} />}
              </React.Fragment>
            ))}
            <span className={styles.ancestorCurrent}>
              <span className={styles.ancestorInitials} style={{ background: getSiteColor(theme) }}>
                {getInitials(focusedUser.displayName)}
              </span>
              <span className={styles.ancestorName}>{focusedUser.displayName}</span>
            </span>
          </div>
        )}

        {/* ── DRILL-DOWN VIEW ── */}
        {isDrillMode && this._renderDrillView()}

        {/* ── FULL TREE VIEW ── */}
        {!isDrillMode && (
          <div
            ref={this._scrollRef}
            className={treeScrollClasses}
            style={{ position: 'relative' }}
            onMouseDown={this._handlePanStart}
            onMouseMove={this._handlePanMove}
            onMouseUp={this._handlePanEnd}
            onMouseLeave={this._handlePanEnd}
            onTouchStart={this._handleTouchStart}
            onTouchEnd={this._handleTouchEnd}
          >
            <div style={{ transformOrigin: 'top center', transform: `scale(${zoomLevel})`, transition: 'transform 0.15s ease' }}>
              <OrgTree
                node={rootNode}
                photos={photos}
                presenceMap={presenceMap}
                showDepartment={showDepartment}
                showOffice={this.props.showOffice}
                chartLayout={chartLayout}
                expandingNodes={expandingNodes}
                searchQuery={lowerQ}
                theme={theme}
                isVisible={isVisible}
                compactCards={compactCards}
                onToggle={this._handleToggle}
                onCardClick={this._handleCardClick}
                onFocus={this._handleFocusUser}
              />
            </div>
          </div>
        )}

        {/* ── Popups backdrop ── */}
        {(showFilters || showDeptFilter || showLayoutPicker) && (
          <div
            className={styles.popupBackdrop}
            onClick={() => this.setState({ showFilters: false, showDeptFilter: false, showLayoutPicker: false })}
          />
        )}

        {/* ── Person card ── */}
        {selectedUser && (
          <PersonCard
            user={selectedUser}
            photo={photos[selectedUser.id] ?? null}
            presence={presenceMap.get(selectedUser.id)}
            theme={theme}
            managerChain={personCardManagerChain}
            onClose={() => this.setState({ selectedUser: null, personCardManagerChain: [] })}
            onFocus={this._handleFocusUser}
          />
        )}
      </div>
    );
  }
}
