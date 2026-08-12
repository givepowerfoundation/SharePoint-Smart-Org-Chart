import { GraphService, IOrgNode } from '../../../../services/GraphService';
import { OrgChartTheme } from '../ISmartOrgChartProps';

export interface IOrgChartProps {
  graphService: GraphService | null;
  topLevelUser: string;
  levelsBelow: number;
  levelsAbove: number;
  showDepartment: boolean;
  showOffice: boolean;
  theme: OrgChartTheme;
  currentUserEmail: string;
  compactCards: boolean;
  defaultLayout: 'drill' | 'vertical' | 'horizontal';
  enableFindMe: boolean;
  enableLayoutToggle: boolean;
  enableStats: boolean;
  enableDeptFilter: boolean;
  enableCountryFilter: boolean;
  /** Employee-type filter (property key is enableUserFilter for upgrade compatibility) */
  enableUserFilter: boolean;
  /** Country name (lower-cased) → hex colour, parsed from the countryColors property */
  countryColors: Map<string, string>;
  defaultZoom: number;
  /** Web part instance ID — scopes localStorage state so instances on different pages don't clash */
  instanceId: string;
}

export interface IOrgChartState {
  rootNode: IOrgNode | null;
  isLoading: boolean;
  error: string | null;
  photos: { [userId: string]: string | null };
  expandingNodes: Set<string>;
  searchQuery: string;
}
