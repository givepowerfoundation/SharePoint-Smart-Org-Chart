import { GraphService } from '../../../../services/GraphService';
import { OrgChartTheme } from '../ISmartOrgChartProps';

export interface IEmployeeDirectoryProps {
  graphService: GraphService;
  alphabetFilterField: 'firstName' | 'lastName';
  cardSize: 'small' | 'medium' | 'large';
  showEmail: boolean;
  showPhone: boolean;
  showDepartment: boolean;
  showOffice: boolean;
  pageSize: number;
  theme: OrgChartTheme;
  /** Country name (lower-cased) → hex colour, parsed from the countryColors property */
  countryColors: Map<string, string>;
  /** Web part instance ID — scopes localStorage state so instances on different pages don't clash */
  instanceId: string;
}
