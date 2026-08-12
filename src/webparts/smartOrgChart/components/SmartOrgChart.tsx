import * as React from 'react';
import { IconButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { GraphService, IUserFilterOptions } from '../../../services/GraphService';
import { MockGraphService, MockCompanySize } from '../../../services/MockGraphService';
import { ISmartOrgChartProps, IUserSettings } from './ISmartOrgChartProps';
import { EmployeeDirectory } from './EmployeeDirectory/EmployeeDirectory';
import { OrgChart } from './OrgChart/OrgChart';
import { SettingsPanel } from './SettingsPanel/SettingsPanel';
import { parseCountryColors } from './countryUtils';
import styles from './SmartOrgChart.module.scss';

const VIEW_META = {
  directory: { label: 'Employee Directory', icon: 'People', toggleIcon: 'Org',  toggleTitle: 'Switch to Org Chart' },
  orgchart:  { label: 'Org Chart',          icon: 'Org',    toggleIcon: 'People', toggleTitle: 'Switch to Employee Directory' },
};

const LS_KEY      = 'smartOrgChart_userSettings';
const LS_MOCK_KEY = 'smartOrgChart_mockSize';
const LS_VIEW_KEY = 'smartOrgChart_currentView';

// All SharePoint sites in a tenant share one origin, so bare keys would be
// shared by every web part instance on every page. Scope them by instance ID,
// reading the legacy un-scoped key as a migration fallback.
function scopedKey(base: string, instanceId: string): string {
  return instanceId ? `${base}_${instanceId}` : base;
}

function lsGet(base: string, instanceId: string): string | null {
  try {
    return localStorage.getItem(scopedKey(base, instanceId)) ?? localStorage.getItem(base);
  } catch {
    return null;
  }
}

function lsSet(base: string, instanceId: string, value: string): void {
  try { localStorage.setItem(scopedKey(base, instanceId), value); } catch { /* ignore */ }
}

interface ISmartOrgChartState {
  currentView: 'directory' | 'orgchart';
  isSettingsOpen: boolean;
  graphService: GraphService | null;
  userSettings: IUserSettings;
  mockSize: MockCompanySize;
  serviceGen: number;
}

function loadUserSettings(defaultFontScale: number, instanceId: string): IUserSettings {
  try {
    const stored = lsGet(LS_KEY, instanceId);
    if (stored) return { ...buildDefaultSettings(defaultFontScale), ...JSON.parse(stored) };
  } catch {
    // ignore
  }
  return buildDefaultSettings(defaultFontScale);
}

function buildDefaultSettings(defaultFontScale = 1): IUserSettings {
  return {
    alphabetFilterField: 'firstName',
    cardSize: 'medium',
    showEmail: true,
    showPhone: true,
    showDepartment: true,
    showOffice: true,
    levelsAbove: 1,
    compactCards: false,
    fontScale: defaultFontScale,
  };
}

function readMockSize(instanceId: string): MockCompanySize {
  const v = lsGet(LS_MOCK_KEY, instanceId);
  if (v === '500')  return 500;
  if (v === '1000') return 1000;
  return 150;
}

function readCurrentView(fallback: 'directory' | 'orgchart', instanceId: string): 'directory' | 'orgchart' {
  const v = lsGet(LS_VIEW_KEY, instanceId);
  if (v === 'directory' || v === 'orgchart') return v;
  return fallback;
}

export class SmartOrgChart extends React.Component<ISmartOrgChartProps, ISmartOrgChartState> {
  private _instanceId: string;

  // Parsed country colour map, re-derived only when the raw property changes. Both child
  // views take the Map, so parsing once here keeps it out of their render paths.
  private _countryColorsRaw = '';
  private _countryColors: Map<string, string> = new Map();

  constructor(props: ISmartOrgChartProps) {
    super(props);
    this._instanceId   = props.context?.instanceId || '';
    const userSettings = loadUserSettings(props.defaultFontScale || 1, this._instanceId);
    const mockSize     = readMockSize(this._instanceId);
    this.state = {
      currentView: readCurrentView(props.defaultView || 'directory', this._instanceId),
      isSettingsOpen: false,
      graphService: null,
      userSettings,
      mockSize,
      serviceGen: 0,
    };
  }

  public async componentDidMount(): Promise<void> {
    await this._initGraphService();
  }

  private _getCountryColors(): Map<string, string> {
    const raw = this.props.countryColors || '';
    if (raw !== this._countryColorsRaw) {
      this._countryColorsRaw = raw;
      this._countryColors    = parseCountryColors(raw);
    }
    return this._countryColors;
  }

  public async componentDidUpdate(prev: ISmartOrgChartProps): Promise<void> {
    if (
      prev.useDemoData           !== this.props.useDemoData ||
      prev.dataSource            !== this.props.dataSource  ||
      prev.excludedAccounts      !== this.props.excludedAccounts ||
      prev.hideDisabledAccounts  !== this.props.hideDisabledAccounts ||
      prev.hideGuestUsers        !== this.props.hideGuestUsers ||
      prev.restrictToTenantDomain !== this.props.restrictToTenantDomain ||
      prev.hideNoJobTitle        !== this.props.hideNoJobTitle ||
      prev.hideNoDepartment      !== this.props.hideNoDepartment ||
      prev.dottedLineAttribute   !== this.props.dottedLineAttribute
    ) {
      await this._initGraphService();
    }
  }

  private _isDemoMode(): boolean {
    return window.location.hostname === 'localhost' || !!this.props.useDemoData;
  }

  private _buildFilterOptions(tenantDomain?: string): IUserFilterOptions {
    return {
      tenantDomain,
      excludedPatterns: (this.props.excludedAccounts || '')
        .split(',')
        .map(s => s.trim().toLowerCase())
        .filter(s => s.length > 0),
      hideGuestUsers:       this.props.hideGuestUsers       || false,
      hideDisabledAccounts: this.props.hideDisabledAccounts || false,
      hideNoJobTitle:       this.props.hideNoJobTitle       || false,
      hideNoDepartment:     this.props.hideNoDepartment     || false,
    };
  }

  private async _initGraphService(): Promise<void> {
    if (this._isDemoMode()) {
      // No tenantDomain here — demo users are @contoso.com, so restricting to
      // the real tenant's domain would hide everyone
      const filterOptions = this._buildFilterOptions();
      this.setState(prev => ({
        graphService: new MockGraphService(prev.mockSize, filterOptions) as unknown as GraphService,
        serviceGen: prev.serviceGen + 1,
      }));
      return;
    }
    const { spHttpClient, msGraphClientFactory, pageContext } = this.props.context;
    let graphClient: MSGraphClientV3 | undefined;
    try {
      graphClient = await msGraphClientFactory.getClient('3');
    } catch {
      // Graph client unavailable — fall back to SP Search only
    }

    // Derive tenant domain from the current user's email when the restriction is enabled
    let tenantDomain: string | undefined;
    if (this.props.restrictToTenantDomain) {
      const userEmail = (pageContext.user?.email || '').toLowerCase();
      const atIdx = userEmail.lastIndexOf('@');
      if (atIdx > 0) tenantDomain = userEmail.substring(atIdx + 1);
    }

    const filterOptions = this._buildFilterOptions(tenantDomain);

    const service = new GraphService(
      spHttpClient,
      pageContext.web.absoluteUrl,
      graphClient,
      this.props.dataSource || 'auto',
      filterOptions,
      this.props.dottedLineAttribute || ''
    );
    this.setState(prev => ({ graphService: service, serviceGen: prev.serviceGen + 1 }));
  }

  private _setMockSize = (size: MockCompanySize): void => {
    lsSet(LS_MOCK_KEY, this._instanceId, String(size));
    this.setState(prev => ({
      mockSize: size,
      graphService: new MockGraphService(size, this._buildFilterOptions()) as unknown as GraphService,
      serviceGen: prev.serviceGen + 1,
    }));
  }

  private _toggleView = (): void => {
    this.setState(prev => {
      const newView: 'directory' | 'orgchart' = prev.currentView === 'directory' ? 'orgchart' : 'directory';
      lsSet(LS_VIEW_KEY, this._instanceId, newView);
      return { currentView: newView };
    });
  }

  private _openSettings = (): void => {
    this.setState({ isSettingsOpen: true });
  }

  private _closeSettings = (): void => {
    this.setState({ isSettingsOpen: false });
  }

  private _saveSettings = (settings: IUserSettings): void => {
    lsSet(LS_KEY, this._instanceId, JSON.stringify(settings));
    this.setState({ userSettings: settings, isSettingsOpen: false });
  }

  public render(): React.ReactElement<ISmartOrgChartProps> {
    const { currentView, isSettingsOpen, graphService, userSettings, mockSize, serviceGen } = this.state;
    const { theme, defaultLayout, logoUrl, companyName } = this.props;
    const meta        = VIEW_META[currentView];
    const resolvedLogoUrl = (() => {
      if (!logoUrl) return '';
      if (logoUrl.startsWith('http://') || logoUrl.startsWith('https://')) return logoUrl;
      // Relative paths can't be served from the local dev server
      if (window.location.hostname === 'localhost') return '';
      if (logoUrl.startsWith('/')) return `${window.location.origin}${logoUrl}`;
      // Site-relative (no leading slash) — resolve against the SharePoint site URL
      const siteUrl = (this.props.context?.pageContext?.web?.absoluteUrl || '').replace(/\/$/, '');
      return siteUrl ? `${siteUrl}/${logoUrl}` : '';
    })();

    return (
      <div className={styles.container} style={{ '--soc-font-scale': userSettings.fontScale || 1 } as React.CSSProperties}>
        <div className={styles.header}>
          <div className={styles.brandArea}>
            {resolvedLogoUrl && (
              <img
                key={resolvedLogoUrl}
                src={resolvedLogoUrl}
                alt="Logo"
                className={styles.logo}
                onError={e => { (e.currentTarget as HTMLImageElement).style.display = 'none'; }}
              />
            )}
            <div className={styles.viewTitle}>
              {companyName && (
                <>
                  <span className={styles.companyName}>{companyName}</span>
                  <span className={styles.brandDivider}>·</span>
                </>
              )}
              <Icon iconName={meta.icon} className={styles.viewIcon} />
              <span>{meta.label}</span>
            </div>
          </div>

          <div className={styles.headerActions}>
            <IconButton
              iconProps={{ iconName: meta.toggleIcon }}
              title={meta.toggleTitle}
              ariaLabel={meta.toggleTitle}
              onClick={this._toggleView}
              className={styles.actionBtn}
            />
            <IconButton
              iconProps={{ iconName: 'Settings' }}
              title="Preferences"
              ariaLabel="Open preferences panel"
              onClick={this._openSettings}
              className={styles.actionBtn}
            />
          </div>
        </div>

        <div className={styles.content}>
          {currentView === 'directory' && graphService && (
            <EmployeeDirectory
              key={`dir-${serviceGen}`}
              graphService={graphService}
              instanceId={this._instanceId}
              alphabetFilterField={userSettings.alphabetFilterField}
              cardSize={userSettings.cardSize}
              showEmail={userSettings.showEmail}
              showPhone={userSettings.showPhone}
              showDepartment={userSettings.showDepartment}
              showOffice={userSettings.showOffice}
              pageSize={this.props.pageSize}
              theme={theme}
              countryColors={this._getCountryColors()}
            />
          )}

          {currentView === 'orgchart' && graphService && (
            <OrgChart
              key={`org-${serviceGen}`}
              graphService={graphService}
              instanceId={this._instanceId}
              topLevelUser={this.props.topLevelUser}
              levelsBelow={this.props.levelsBelow}
              levelsAbove={userSettings.levelsAbove}
              showDepartment={userSettings.showDepartment}
              showOffice={userSettings.showOffice}
              theme={theme}
              currentUserEmail={this.props.context?.pageContext?.user?.email || ''}
              compactCards={userSettings.compactCards}
              defaultLayout={defaultLayout || 'drill'}
              enableFindMe={this.props.enableFindMe !== false}
              enableLayoutToggle={this.props.enableLayoutToggle !== false}
              enableStats={this.props.enableStats !== false}
              enableDeptFilter={this.props.enableDeptFilter !== false}
              enableCountryFilter={this.props.enableCountryFilter !== false}
              enableUserFilter={this.props.enableUserFilter !== false}
              countryColors={this._getCountryColors()}
              defaultZoom={this.props.defaultZoom ?? 0}
            />
          )}
        </div>

        <SettingsPanel
          isOpen={isSettingsOpen}
          settings={userSettings}
          onDismiss={this._closeSettings}
          onSave={this._saveSettings}
          mockSize={this._isDemoMode() ? mockSize : undefined}
          onMockSizeChange={this._isDemoMode() ? this._setMockSize : undefined}
        />
      </div>
    );
  }
}
