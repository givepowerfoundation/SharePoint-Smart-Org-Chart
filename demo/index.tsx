import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { SmartOrgChart } from '../src/webparts/smartOrgChart/components/SmartOrgChart';
import { MockCompanySize } from '../src/services/MockGraphService';
import { OrgChartTheme } from '../src/webparts/smartOrgChart/components/ISmartOrgChartProps';

const p = new URLSearchParams(window.location.search);

const view     = (p.get('view')   || 'directory') as 'directory' | 'orgchart';
const layout   = (p.get('layout') || 'drill')     as 'drill' | 'vertical' | 'horizontal';
const theme    = (p.get('theme')  || 'modern')    as OrgChartTheme;
const sizeRaw  = p.get('mockSize');
const mockSize: MockCompanySize = sizeRaw === '500' ? 500 : sizeRaw === '1000' ? 1000 : 150;

// Persist the requested mock size so SmartOrgChart's readMockSize() picks it up.
try { localStorage.setItem('smartOrgChart_mockSize', String(mockSize)); } catch { /* ignore */ }
// Force the requested view so readCurrentView() doesn't restore a stale value.
try { localStorage.setItem('smartOrgChart_currentView', view); } catch { /* ignore */ }
// Force the requested layout so OrgChart doesn't restore a stale chartLayout from localStorage.
try {
  const chartState = JSON.parse(localStorage.getItem('smartOrgChart_chartState') || '{}');
  chartState.chartLayout = layout;
  // Clear saved navigation so tree layouts always start from the root.
  delete chartState.focusEmail;
  localStorage.setItem('smartOrgChart_chartState', JSON.stringify(chartState));
} catch { /* ignore */ }

const mockContext = {
  pageContext: {
    user: { email: 'demo@contoso.com' },
    web: { absoluteUrl: 'https://contoso.sharepoint.com/sites/demo' },
  },
  spHttpClient: {},
  msGraphClientFactory: {},
} as any;

ReactDOM.render(
  <SmartOrgChart
    context={mockContext}
    defaultView={view}
    topLevelUser="a.chen@contoso.com"
    levelsBelow={3}
    pageSize={50}
    useDemoData={true}
    hideDemoBanner={true}
    companyName="Contoso"
    logoUrl=""
    theme={theme}
    defaultLayout={layout}
    enableFindMe={true}
    enableLayoutToggle={true}
    enableStats={true}
    enableDeptFilter={true}
    enableUserFilter={true}
    onSettingsSaved={() => { /* no-op in demo */ }}
  />,
  document.getElementById('root'),
);
