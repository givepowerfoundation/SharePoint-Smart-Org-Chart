import * as React from 'react';
import { IconButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { GraphService } from '../../../services/GraphService';
import { MockGraphService, MockCompanySize } from '../../../services/MockGraphService';
import { ISmartOrgChartProps, IUserSettings } from './ISmartOrgChartProps';
import { EmployeeDirectory } from './EmployeeDirectory/EmployeeDirectory';
import { OrgChart } from './OrgChart/OrgChart';
import { SettingsPanel } from './SettingsPanel/SettingsPanel';
import styles from './SmartOrgChart.module.scss';

const VIEW_META = {
  directory: { label: 'Employee Directory', icon: 'People', toggleIcon: 'Org',  toggleTitle: 'Switch to Org Chart' },
  orgchart:  { label: 'Org Chart',          icon: 'Org',    toggleIcon: 'People', toggleTitle: 'Switch to Employee Directory' },
};

const LS_KEY      = 'smartOrgChart_userSettings';
const LS_MOCK_KEY = 'smartOrgChart_mockSize';
const LS_VIEW_KEY = 'smartOrgChart_currentView';

const MOCK_SIZES: MockCompanySize[] = [150, 500, 1000];

interface ISmartOrgChartState {
  currentView: 'directory' | 'orgchart';
  isSettingsOpen: boolean;
  graphService: GraphService | null;
  userSettings: IUserSettings;
  mockSize: MockCompanySize;
}

function loadUserSettings(): IUserSettings {
  try {
    const stored = localStorage.getItem(LS_KEY);
    if (stored) return { ...buildDefaultSettings(), ...JSON.parse(stored) };
  } catch {
    // ignore
  }
  return buildDefaultSettings();
}

function buildDefaultSettings(): IUserSettings {
  return {
    alphabetFilterField: 'firstName',
    cardSize: 'medium',
    showEmail: true,
    showPhone: true,
    showDepartment: true,
    showOffice: true,
    levelsAbove: 1,
    compactCards: false,
  };
}

function readMockSize(): MockCompanySize {
  try {
    const v = localStorage.getItem(LS_MOCK_KEY);
    if (v === '500')  return 500;
    if (v === '1000') return 1000;
  } catch { /* ignore */ }
  return 150;
}

function readCurrentView(fallback: 'directory' | 'orgchart'): 'directory' | 'orgchart' {
  try {
    const v = localStorage.getItem(LS_VIEW_KEY);
    if (v === 'directory' || v === 'orgchart') return v;
  } catch { /* ignore */ }
  return fallback;
}

export class SmartOrgChart extends React.Component<ISmartOrgChartProps, ISmartOrgChartState> {
  private _directoryRef = React.createRef<EmployeeDirectory>();
  private _orgChartRef  = React.createRef<OrgChart>();

  constructor(props: ISmartOrgChartProps) {
    super(props);
    const userSettings = loadUserSettings();
    const mockSize     = readMockSize();
    this.state = {
      currentView: readCurrentView(props.defaultView || 'directory'),
      isSettingsOpen: false,
      graphService: null,
      userSettings,
      mockSize,
    };
  }

  public async componentDidMount(): Promise<void> {
    await this._initGraphService();
  }

  public async componentDidUpdate(prev: ISmartOrgChartProps): Promise<void> {
    if (prev.useDemoData !== this.props.useDemoData) {
      await this._initGraphService();
    }
  }

  private _isDemoMode(): boolean {
    return window.location.hostname === 'localhost' || !!this.props.useDemoData;
  }

  private async _initGraphService(): Promise<void> {
    if (this._isDemoMode()) {
      this.setState({ graphService: new MockGraphService(this.state.mockSize) as unknown as GraphService });
      return;
    }
    const { spHttpClient, msGraphClientFactory, pageContext } = this.props.context;
    let graphClient: MSGraphClientV3 | undefined;
    try {
      graphClient = await msGraphClientFactory.getClient('3');
    } catch {
      // Graph client unavailable — fall back to SP Search only
    }
    this.setState({ graphService: new GraphService(spHttpClient, pageContext.web.absoluteUrl, graphClient) });
  }

  private _setMockSize = (size: MockCompanySize): void => {
    try { localStorage.setItem(LS_MOCK_KEY, String(size)); } catch { /* ignore */ }
    this.setState({
      mockSize: size,
      graphService: new MockGraphService(size) as unknown as GraphService,
    });
  }

  private _toggleView = (): void => {
    this.setState(prev => {
      const newView: 'directory' | 'orgchart' = prev.currentView === 'directory' ? 'orgchart' : 'directory';
      try { localStorage.setItem(LS_VIEW_KEY, newView); } catch { /* ignore */ }
      return { currentView: newView };
    });
  }

  private _exportPdf = (): void => {
    if (this.state.currentView === 'directory') {
      this._directoryRef.current?.exportPdf();
    } else {
      this._orgChartRef.current?.exportPdf();
    }
  }

  private _exportPng = (): void => {
    this._orgChartRef.current?.exportPng();
  }

  private _openSettings = (): void => {
    this.setState({ isSettingsOpen: true });
  }

  private _closeSettings = (): void => {
    this.setState({ isSettingsOpen: false });
  }

  private _saveSettings = (settings: IUserSettings): void => {
    try {
      localStorage.setItem(LS_KEY, JSON.stringify(settings));
    } catch {
      // ignore storage errors
    }
    this.setState({ userSettings: settings, isSettingsOpen: false });
  }

  public render(): React.ReactElement<ISmartOrgChartProps> {
    const { currentView, isSettingsOpen, graphService, userSettings, mockSize } = this.state;
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
      <div className={styles.container}>
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
              iconProps={{ iconName: 'Download' }}
              title="Export to PDF"
              ariaLabel="Export to PDF"
              onClick={this._exportPdf}
              className={styles.actionBtn}
            />
            {currentView === 'orgchart' && (
              <IconButton
                iconProps={{ iconName: 'Photo2' }}
                title="Export to PNG image"
                ariaLabel="Export org chart to PNG"
                onClick={this._exportPng}
                className={styles.actionBtn}
              />
            )}
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
              ref={this._directoryRef}
              graphService={graphService}
              alphabetFilterField={userSettings.alphabetFilterField}
              cardSize={userSettings.cardSize}
              showEmail={userSettings.showEmail}
              showPhone={userSettings.showPhone}
              showDepartment={userSettings.showDepartment}
              showOffice={userSettings.showOffice}
              pageSize={this.props.pageSize}
              theme={theme}
            />
          )}

          {currentView === 'orgchart' && graphService && (
            <OrgChart
              ref={this._orgChartRef}
              graphService={graphService}
              topLevelUser={this.props.topLevelUser}
              levelsBelow={this.props.levelsBelow}
              levelsAbove={userSettings.levelsAbove}
              showEmail={userSettings.showEmail}
              showPhone={userSettings.showPhone}
              showDepartment={userSettings.showDepartment}
              theme={theme}
              currentUserEmail={this.props.context?.pageContext?.user?.email || ''}
              compactCards={userSettings.compactCards}
              defaultLayout={defaultLayout || 'drill'}
              enableFindMe={this.props.enableFindMe !== false}
              enableLayoutToggle={this.props.enableLayoutToggle !== false}
              enableStats={this.props.enableStats !== false}
              enableDeptFilter={this.props.enableDeptFilter !== false}
              enableUserFilter={this.props.enableUserFilter !== false}
            />
          )}
        </div>

        {this.props.useDemoData && !this.props.hideDemoBanner && (
          <div className={styles.demoBanner}>
            <span className={styles.demoBannerLabel}>Demo data</span>
            {MOCK_SIZES.map(s => (
              <button
                key={s}
                className={`${styles.demoBannerBtn}${mockSize === s ? ` ${styles.demoBannerBtnActive}` : ''}`}
                onClick={() => this._setMockSize(s)}
              >
                {s.toLocaleString()} people
              </button>
            ))}
          </div>
        )}

        <SettingsPanel
          isOpen={isSettingsOpen}
          settings={userSettings}
          onDismiss={this._closeSettings}
          onSave={this._saveSettings}
        />
      </div>
    );
  }
}
