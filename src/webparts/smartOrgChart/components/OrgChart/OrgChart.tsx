import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import { TextField } from '@fluentui/react/lib/TextField';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { IGraphUser, IOrgNode, PresenceAvailability } from '../../../../services/GraphService';
import { exportOrgChartToPdf, exportOrgChartToCsv } from '../../../../services/PdfExportService';
import { IOrgChartProps, IOrgChartState } from './IOrgChartProps';
import { PRESENCE_COLOR, PRESENCE_LABEL, getInitials } from '../personUtils';
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

// Expands every ancestor of a node matching the query so search hits inside
// collapsed branches become visible. Only already-loaded nodes are affected.
function expandToMatches(node: IOrgNode, lowerQ: string): { node: IOrgNode; hasMatch: boolean } {
  const children = node.directReports.map(c => expandToMatches(c, lowerQ));
  const childMatch = children.some(c => c.hasMatch);
  return {
    node: { ...node, directReports: children.map(c => c.node), isExpanded: node.isExpanded || childMatch },
    hasMatch: childMatch || matchesQuery(node, lowerQ),
  };
}

// Returns true if this node or any descendant passes the visibility filter.
// Used so ancestor nodes stay visible when a dept filter is active.
function subtreeHasVisible(node: IOrgNode, isVisible: (u: IGraphUser) => boolean): boolean {
  if (isVisible(node.user)) return true;
  return node.directReports.some(c => subtreeHasVisible(c, isVisible));
}

// Prunes users hidden by the active filters (mirrors what OrgTree renders)
function filterTreeForExport(node: IOrgNode, isVisible: (u: IGraphUser) => boolean): IOrgNode | null {
  if (!isVisible(node.user)) return null;
  const directReports = node.directReports
    .map(c => filterTreeForExport(c, isVisible))
    .filter((c): c is IOrgNode => c !== null);
  return { ...node, directReports };
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

function computeStats(users: IGraphUser[]): {
  total: number; members: number; guests: number; depts: number;
} {
  let members = 0, guests = 0;
  const deptSet = new Set<string>();
  for (const u of users) {
    if (u.department) deptSet.add(u.department);
    if (u.userType === 'Guest') guests++;
    else members++;
  }
  return { total: users.length, members, guests, depts: deptSet.size };
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

const PresenceDot: React.FC<{ status: PresenceAvailability | undefined }> = ({ status }) => {
  if (!status || status === 'Unknown') return null;
  return <span className={styles.presenceDot} style={{ background: PRESENCE_COLOR[status] }} />;
};

/* ── Person Card (Outlook-style popup) ───── */

interface IPersonCardProps {
  user: IGraphUser;
  photo: string | null;
  presence: PresenceAvailability | undefined;
  theme: OrgChartTheme;
  managerChain: IGraphUser[];
  dottedManager: IGraphUser | null;
  dottedReports: IGraphUser[];
  onClose: () => void;
  onFocus: (user: IGraphUser) => void;
}

const PersonCard: React.FC<IPersonCardProps> = ({
  user, photo, presence, theme, managerChain, dottedManager, dottedReports, onClose, onFocus
}) => {
  React.useEffect(() => {
    const onKey = (e: KeyboardEvent): void => { if (e.key === 'Escape') onClose(); };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [onClose]);

  // Transient "Copied!" feedback for the copy buttons
  const mountedRef = React.useRef(true);
  React.useEffect(() => () => { mountedRef.current = false; }, []);
  const [copied, setCopied] = React.useState('');
  const copy = (text: string, label: string): void => {
    const done = (): void => {
      if (!mountedRef.current) return;
      setCopied(label);
      window.setTimeout(() => { if (mountedRef.current) setCopied(''); }, 2000);
    };
    try { navigator.clipboard.writeText(text).then(done).catch(done); } catch { done(); }
  };

  const deptColor = getSiteColor(theme);
  const isDark = theme === 'dark';
  const initials = getInitials(user.displayName);
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
              <span className={styles.personCardPresenceDot} style={{ background: PRESENCE_COLOR[presence] }} />
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
                <button
                  onClick={() => copy(user.mail, 'email')}
                  title="Copy email address"
                  style={{ background: 'none', border: 'none', cursor: 'pointer', color: subColor, padding: '2px 4px' }}
                >
                  <Icon iconName={copied === 'email' ? 'CheckMark' : 'Copy'} />
                </button>
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

          {/* Dotted-line relationships */}
          {(dottedManager || dottedReports.length > 0) && (
            <div className={styles.personCardChain} style={{ background: chainBg, borderColor: borderClr }}>
              <div className={styles.personCardChainLabel} style={{ color: subColor }}>Dotted line</div>
              <div className={styles.personCardChainItems}>
                {dottedManager && (
                  <button
                    className={styles.personCardChainChip}
                    onClick={() => { onClose(); onFocus(dottedManager); }}
                    title={`Dotted-line manager: ${dottedManager.displayName}`}
                  >
                    <span className={styles.personCardChainInitials} style={{ background: getSiteColor(theme) }}>
                      {getInitials(dottedManager.displayName)}
                    </span>
                    <span className={styles.personCardChainName} style={{ color: textColor }}>
                      ↑ {dottedManager.displayName.split(' ')[0]}
                    </span>
                  </button>
                )}
                {dottedReports.map(rep => (
                  <button
                    key={rep.id}
                    className={styles.personCardChainChip}
                    onClick={() => { onClose(); onFocus(rep); }}
                    title={`Dotted-line report: ${rep.displayName}`}
                  >
                    <span className={styles.personCardChainInitials} style={{ background: getSiteColor(theme) }}>
                      {getInitials(rep.displayName)}
                    </span>
                    <span className={styles.personCardChainName} style={{ color: textColor }}>
                      ↓ {rep.displayName.split(' ')[0]}
                    </span>
                  </button>
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
      onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onCardClick(user); } }}
      role="button"
      tabIndex={0}
      aria-label={`View ${user.displayName}'s profile`}
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
  if (!subtreeHasVisible(node, isVisible)) return null;

  const visibleReports     = node.directReports.filter(c => subtreeHasVisible(c, isVisible));
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
  personCardDottedManager: IGraphUser | null;
  personCardDottedReports: IGraphUser[];
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
  rootPickerQuery: string;
  rootPickerResults: IGraphUser[];
  runtimeRootUser: IGraphUser | null;
}

/* ── Chart state persistence ─────────────── */

const LS_CHART_KEY = 'smartOrgChart_chartState';

interface IChartStoredState {
  chartLayout?: ChartLayout;
  showStats?: boolean;
  filterMembers?: boolean;
  filterGuests?: boolean;
  filterDepartments?: string[];
  focusEmail?: string | null;
}

// Storage is scoped per web part instance — all SharePoint sites share one
// origin, so a bare key would leak state between instances on different pages.
// The un-scoped legacy key is read as a migration fallback.
function chartStateKey(instanceId: string): string {
  return instanceId ? `${LS_CHART_KEY}_${instanceId}` : LS_CHART_KEY;
}

function loadChartState(instanceId: string): IChartStoredState {
  try {
    const s = localStorage.getItem(chartStateKey(instanceId)) ?? localStorage.getItem(LS_CHART_KEY);
    if (s) return JSON.parse(s) as IChartStoredState;
  } catch { /* ignore */ }
  return {};
}

function saveChartState(instanceId: string, s: IChartStoredState): void {
  try { localStorage.setItem(chartStateKey(instanceId), JSON.stringify(s)); } catch { /* ignore */ }
}

/* ── Deep links (?socFocus=email) ────────── */

const FOCUS_URL_PARAM = 'socFocus';

function readUrlFocus(): string | null {
  try { return new URLSearchParams(window.location.search).get(FOCUS_URL_PARAM); } catch { return null; }
}


function updateUrlFocus(email: string | null): void {
  try {
    const url = new URL(window.location.href);
    if ((email || null) === url.searchParams.get(FOCUS_URL_PARAM)) return;
    if (email) url.searchParams.set(FOCUS_URL_PARAM, email);
    else url.searchParams.delete(FOCUS_URL_PARAM);
    window.history.replaceState(null, '', url.toString());
  } catch { /* ignore */ }
}

const LAYOUT_CYCLE: ChartLayout[] = ['drill', 'vertical', 'horizontal'];

const LAYOUT_ICON: Record<ChartLayout, string> = {
  drill:      'Org',
  vertical:   'Down',
  horizontal: 'Forward',
};

const LAYOUT_TITLE: Record<ChartLayout, string> = {
  drill:      'Drill-Down',
  vertical:   'Top Down',
  horizontal: 'Left to Right',
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
  private _rootPickerRef    = React.createRef<HTMLDivElement>();

  constructor(props: IOrgChartProps) {
    super(props);
    const stored = loadChartState(props.instanceId);
    // A shared deep link takes precedence over the user's own saved position
    this._pendingFocusEmail = readUrlFocus() || stored.focusEmail || null;
    // Stored preferences only apply while the admin has the matching control
    // enabled — otherwise users could be stuck in a state they can't change.
    this.state = {
      rootNode: null, isLoading: false, error: null,
      photos: {}, expandingNodes: new Set(), searchQuery: '',
      presenceMap: new Map(), zoomLevel: props.defaultZoom > 0 ? props.defaultZoom : 1,
      selectedUser: null,
      showFilters: false,
      filterMembers: props.enableUserFilter ? (stored.filterMembers ?? true) : true,
      filterGuests: props.enableUserFilter ? (stored.filterGuests ?? true) : true,
      isDragging: false,
      focusedUser: null, ancestorChain: [], allUsers: [],
      showSearchResults: false, personCardManagerChain: [],
      personCardDottedManager: null, personCardDottedReports: [],
      chartLayout: props.enableLayoutToggle
        ? (stored.chartLayout ?? (props.defaultLayout || 'drill'))
        : (props.defaultLayout || 'drill'),
      findMeError: '',
      filterDepartments: props.enableDeptFilter ? new Set(stored.filterDepartments ?? []) : new Set(),
      showDeptFilter: false,
      showStats: props.enableStats ? (stored.showStats ?? false) : false,
      drillPath: [], drillReports: [], drillLoadingId: null,
      drillReportCounts: new Map(),
      showLayoutPicker: false,
      rootPickerQuery: '', rootPickerResults: [], runtimeRootUser: null,
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

    if (
      prevState.rootNode          !== this.state.rootNode          ||
      prevState.filterDepartments !== this.state.filterDepartments ||
      prevState.filterMembers     !== this.state.filterMembers     ||
      prevState.filterGuests      !== this.state.filterGuests      ||
      prevState.zoomLevel         !== this.state.zoomLevel         ||
      prevState.chartLayout       !== this.state.chartLayout
    ) {
      this._fixConnectorLines();
    }

    const { isLoading, chartLayout, showStats, filterMembers, filterGuests,
            filterDepartments, drillPath, focusedUser } = this.state;
    if (!isLoading && (
      prevState.chartLayout       !== chartLayout       ||
      prevState.showStats         !== showStats         ||
      prevState.filterMembers     !== filterMembers     ||
      prevState.filterGuests      !== filterGuests      ||
      prevState.filterDepartments !== filterDepartments ||
      prevState.drillPath         !== drillPath         ||
      prevState.focusedUser       !== focusedUser
    )) {
      const focusEmail = chartLayout === 'drill'
        ? (drillPath.length > 0 ? drillPath[drillPath.length - 1].mail ?? null : null)
        : (focusedUser?.mail ?? null);
      saveChartState(this.props.instanceId, {
        chartLayout, showStats, filterMembers, filterGuests,
        filterDepartments: Array.from(filterDepartments),
        focusEmail,
      });
      updateUrlFocus(focusEmail);
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
    if (this._rootPickerRef.current && !this._rootPickerRef.current.contains(e.target as Node)) {
      if (this.state.rootPickerResults.length > 0) this.setState({ rootPickerResults: [] });
    }
  }

  private async _refreshPresence(): Promise<void> {
    const { graphService } = this.props;
    if (!graphService || !this._mounted) return;
    // Only request presence for users actually on screen
    const ids = new Set<string>();
    const collect = (n: IOrgNode): void => { ids.add(n.user.id); n.directReports.forEach(collect); };
    if (this.state.rootNode) collect(this.state.rootNode);
    this.state.drillPath.forEach(u => ids.add(u.id));
    this.state.drillReports.forEach(u => ids.add(u.id));
    if (this.state.selectedUser) ids.add(this.state.selectedUser.id);
    if (ids.size === 0) return;
    const presenceMap = await graphService.getPresence(Array.from(ids));
    if (this._mounted) this.setState({ presenceMap });
  }

  // Builds the tree that matches what's on screen: the current drill level in
  // drill mode, otherwise the loaded tree with filtered-out users pruned.
  private _getExportTree(): { node: IOrgNode; note?: string } | null {
    const { rootNode, chartLayout, drillPath, drillReports } = this.state;
    const isVisible = this._buildIsVisible();

    if (chartLayout === 'drill' && drillPath.length > 0) {
      const current = drillPath[drillPath.length - 1];
      return {
        node: {
          user: current,
          directReports: drillReports.filter(isVisible).map(u => ({
            user: u, directReports: [], isExpanded: false, childrenLoaded: true, level: 1,
          })),
          isExpanded: true, childrenLoaded: true, level: 0,
        },
        note: 'Current drill-down level (one level of direct reports)',
      };
    }

    if (!rootNode) return null;
    const filtered = filterTreeForExport(rootNode, isVisible);
    if (!filtered) return null;
    const note = this._getUnloadedFrontier(rootNode).length > 0
      ? 'Includes loaded levels only — use Expand All before exporting to include deeper levels'
      : undefined;
    return { node: filtered, note };
  }

  public exportPdf(): void {
    const tree = this._getExportTree();
    if (tree) exportOrgChartToPdf(tree.node, tree.note);
  }

  private _exportCsv = (): void => {
    const tree = this._getExportTree();
    if (tree) exportOrgChartToCsv(tree.node);
  }

  private _onSearchChange = (value: string): void => {
    const q = value.trim().toLowerCase();
    this.setState(prev => {
      const updates: Partial<IOrgChartLocalState> = {
        searchQuery: value,
        showSearchResults: !!q,
        showFilters: false,
      };
      // Reveal matches hiding inside collapsed branches
      if (q && prev.chartLayout !== 'drill' && prev.rootNode) {
        updates.rootNode = expandToMatches(prev.rootNode, q).node;
      }
      return updates as IOrgChartLocalState;
    });
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
    } catch (err) {
      const detail = err instanceof Error && err.message ? ` ${err.message}` : ' Check permissions.';
      if (this._mounted) this.setState({ isLoading: false, error: `Failed to load org chart.${detail}` });
    }
  }

  private _fixConnectorLines(): void {
    requestAnimationFrame(() => {
      const container = this._scrollRef.current;
      if (!container || !this._mounted) return;
      const zoom = this.state.zoomLevel || 1;
      const allChildren = container.querySelectorAll(`.${styles.children}`) as NodeListOf<HTMLElement>;
      allChildren.forEach(childrenDiv => {
        const first = childrenDiv.firstElementChild as HTMLElement;
        const last  = childrenDiv.lastElementChild  as HTMLElement;
        if (!first || !last) return;
        const left  = first.getBoundingClientRect().width / (2 * zoom);
        const right = last.getBoundingClientRect().width  / (2 * zoom);
        childrenDiv.style.setProperty('--conn-left',  `${left}px`);
        childrenDiv.style.setProperty('--conn-right', `${right}px`);
      });
    });
  }

  private _autoFitZoom(): void {
    if (this.props.defaultZoom > 0) return; // fixed zoom configured — don't override
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
      const scaleX = cW / naturalW;
      const scaleY = cH / naturalH;
      const fitZoom = Math.min(scaleX, scaleY) * 0.90;
      // Shrink to fit, or grow back toward 100% when a smaller subtree is shown
      const newZoom = Math.max(0.25, Math.min(1, fitZoom));
      if (Math.abs(newZoom - this.state.zoomLevel) > 0.01) {
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
    // Batch setState calls — one render per 10 photos instead of one per photo
    let batch: { [id: string]: string | null } = {};
    const flush = (): void => {
      const toApply = batch;
      batch = {};
      if (this._mounted && Object.keys(toApply).length > 0) {
        this.setState(prev => ({ photos: { ...prev.photos, ...toApply } }));
      }
    };
    for (const id of ids) {
      if (!this._mounted) return;
      if (id in this.state.photos || id in batch) continue;
      batch[id] = await graphService.getUserPhoto(id);
      if (Object.keys(batch).length >= 10) flush();
    }
    flush();
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
    if (this._mounted) this._autoFitZoom();
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
    this.setState({
      selectedUser: user, showFilters: false, personCardManagerChain: [],
      personCardDottedManager: null, personCardDottedReports: [],
    });
    const gs = this.props.graphService;
    if (!gs) return;
    gs.getManagerChain(user.id, 8).then(chain => {
      if (this._mounted) this.setState({ personCardManagerChain: chain });
    }).catch(() => { /* ignore */ });
    gs.getDottedLineReports(user.id).then(reports => {
      if (this._mounted && reports.length > 0) this.setState({ personCardDottedReports: reports });
    }).catch(() => { /* ignore */ });
    if (user.dottedManagerId) {
      gs.findUser(user.dottedManagerId).then(mgr => {
        if (this._mounted && mgr) this.setState({ personCardDottedManager: mgr });
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
        this.setState({ drillLoadingId: null });
        this._handleCardClick(user);
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
    const { graphService, levelsBelow, levelsAbove } = this.props;
    if (!graphService) return;
    const { chartLayout } = this.state;

    this.setState({ isLoading: true, error: null, searchQuery: '', showSearchResults: false });

    if (chartLayout === 'drill') {
      try {
        const [reports, managerChain] = await Promise.all([
          graphService.getDirectReports(user.id),
          graphService.getManagerChain(user.id, levelsAbove),
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
        graphService.getManagerChain(user.id, levelsAbove),
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

  /* ── Root picker ── */

  private _onRootPickerChange = (query: string): void => {
    const { allUsers } = this.state;
    if (!query.trim()) {
      this.setState({ rootPickerQuery: query, rootPickerResults: [] });
      return;
    }
    const q = query.toLowerCase();
    const results = allUsers.filter(u =>
      (u.displayName || '').toLowerCase().includes(q) ||
      (u.mail || '').toLowerCase().includes(q)
    ).slice(0, 8);
    this.setState({ rootPickerQuery: query, rootPickerResults: results });
  }

  private _onRootPickerSelect = async (user: IGraphUser): Promise<void> => {
    const { graphService, levelsBelow } = this.props;
    if (!graphService) return;
    this.setState({ rootPickerQuery: '', rootPickerResults: [], runtimeRootUser: user, isLoading: true });
    try {
      const rawRoot = await graphService.buildOrgTree(user.id, levelsBelow);
      const rootNode = expandLoaded(rawRoot);
      const drillPath = [user];
      const drillReports = rootNode.directReports.map(n => n.user);
      if (this._mounted) {
        this._requestedReportCounts.clear();
        this.setState({
          rootNode, drillPath, drillReports, drillLoadingId: null,
          isLoading: false, focusedUser: null, ancestorChain: [], error: null,
        }, () => { this._autoFitZoom(); });
        this._loadPhotosForTree(rootNode);
        this._checkFrontierNodes(rootNode);
        this._loadDrillReportCounts(drillReports);
      }
    } catch {
      if (this._mounted) this.setState({ isLoading: false, error: 'Failed to load org chart for this person.' });
    }
  }

  private _resetRoot = (): void => {
    this.setState({ runtimeRootUser: null, rootPickerQuery: '', rootPickerResults: [] });
    this._loadTree();
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
      <NoConfigForm onLoad={async (identifier) => {
        const gs = this.props.graphService;
        if (!gs) return;
        this.setState({ isLoading: true, error: null });
        try {
          const user = await gs.findUser(identifier);
          if (!user) {
            if (this._mounted) this.setState({ isLoading: false, error: `User "${identifier}" not found.` });
            return;
          }
          const [rawRoot, allUsers] = await Promise.all([
            gs.buildOrgTree(user.id, this.props.levelsBelow),
            gs.getAllUsers().catch(() => [] as IGraphUser[]),
          ]);
          if (this._mounted) {
            const rootNode = expandLoaded(rawRoot);
            const drillPath = [user];
            const drillReports = rootNode.directReports.map(n => n.user);
            this.setState({ rootNode, allUsers, drillPath, drillReports, isLoading: false });
            this._loadPhotosForTree(rootNode);
            this._loadDrillReportCounts(drillReports);
          }
        } catch {
          if (this._mounted) this.setState({ isLoading: false, error: 'Failed to load org chart.' });
        }
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
      rootPickerQuery, rootPickerResults, runtimeRootUser,
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
    const stats        = showStats ? computeStats(allUsers) : null;
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
              onChange={(_, v) => this._onSearchChange(v || '')}
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

          {/* Export PDF / CSV */}
          <button className={styles.iconToolBtn} onClick={() => this.exportPdf()} title="Download as PDF">
            <Icon iconName="PDF" />
          </button>
          <button className={styles.iconToolBtn} onClick={this._exportCsv} title="Download as CSV spreadsheet">
            <Icon iconName="ExcelDocument" />
          </button>

          {/* Root picker — "View from person…" */}
          <div className={styles.rootPickerWrapper} ref={this._rootPickerRef}>
            {runtimeRootUser ? (
              <div className={styles.rootPickerActive}>
                <Icon iconName="Org" className={styles.rootPickerIcon} />
                <span className={styles.rootPickerActiveName}>{runtimeRootUser.displayName}</span>
                <button
                  className={styles.rootPickerReset}
                  onClick={this._resetRoot}
                  title="Reset to default root"
                >
                  <Icon iconName="Cancel" />
                </button>
              </div>
            ) : (
              <>
                <div className={styles.rootPickerInputWrap}>
                  <Icon iconName="Org" className={styles.rootPickerIcon} />
                  <input
                    type="text"
                    className={styles.rootPickerInput}
                    placeholder="View from person…"
                    value={rootPickerQuery}
                    onChange={e => this._onRootPickerChange(e.target.value)}
                  />
                </div>
                {rootPickerResults.length > 0 && (
                  <div className={styles.rootPickerDropdown}>
                    {rootPickerResults.map(u => (
                      <button
                        key={u.id}
                        className={styles.rootPickerOption}
                        onClick={() => this._onRootPickerSelect(u)}
                      >
                        <span className={styles.rootPickerOptionInitials} style={{ background: getSiteColor(theme) }}>
                          {getInitials(u.displayName)}
                        </span>
                        <span className={styles.rootPickerOptionInfo}>
                          <span className={styles.rootPickerOptionName}>{u.displayName}</span>
                          {u.jobTitle && <span className={styles.rootPickerOptionMeta}>{u.jobTitle}</span>}
                        </span>
                      </button>
                    ))}
                  </div>
                )}
              </>
            )}
          </div>

          {/* Zoom — only in full-tree mode */}
          {!isDrillMode && (
            <div className={styles.zoomControls}>
              <button className={styles.zoomBtn} onClick={() => this.setState({ zoomLevel: Math.max(0.25, zoomLevel - 0.1) })} title="Zoom out" disabled={zoomLevel <= 0.25}>
                <Icon iconName="Remove" />
              </button>
              <span className={styles.zoomLabel}>{Math.round(zoomLevel * 100)}%</span>
              <button className={styles.zoomBtn} onClick={() => { const next = Math.min(1.5, zoomLevel + 0.1); this.setState({ zoomLevel: zoomLevel < 1 && next > 1 ? 1 : next }); }} title="Zoom in" disabled={zoomLevel >= 1.5}>
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
            <div style={{ zoom: zoomLevel, display: 'inline-block', minWidth: '100%' }}>
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
            dottedManager={this.state.personCardDottedManager}
            dottedReports={this.state.personCardDottedReports}
            onClose={() => this.setState({
              selectedUser: null, personCardManagerChain: [],
              personCardDottedManager: null, personCardDottedReports: [],
            })}
            onFocus={this._handleFocusUser}
          />
        )}
      </div>
    );
  }
}
