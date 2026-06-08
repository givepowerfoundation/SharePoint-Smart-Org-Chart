import { IUserSettings } from '../ISmartOrgChartProps';

export interface ISettingsPanelProps {
  isOpen: boolean;
  settings: IUserSettings;
  onDismiss: () => void;
  onSave: (settings: IUserSettings) => void;
  mockSize?: number;
  onMockSizeChange?: (size: number) => void;
}
