import { GraphService, IOrgNode } from '../../../../services/GraphService';
import { OrgChartTheme } from '../ISmartOrgChartProps';

export interface IOrgChartProps {
  graphService: GraphService | null;
  topLevelUser: string;
  levelsBelow: number;
  levelsAbove: number;
  showEmail: boolean;
  showPhone: boolean;
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
  enableUserFilter: boolean;
  defaultZoom: number;
}

export interface IOrgChartState {
  rootNode: IOrgNode | null;
  isLoading: boolean;
  error: string | null;
  photos: { [userId: string]: string | null };
  expandingNodes: Set<string>;
  searchQuery: string;
}
