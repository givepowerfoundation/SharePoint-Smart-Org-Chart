// Shared person-rendering helpers used by both the Employee Directory and the Org Chart.
import { PresenceAvailability } from '../../../services/GraphService';

export const PRESENCE_COLOR: Record<PresenceAvailability, string> = {
  Available:    '#6BB700',
  Busy:         '#C50F1F',
  DoNotDisturb: '#C50F1F',
  BeRightBack:  '#FFAA44',
  Away:         '#FFAA44',
  Offline:      '#8A8886',
  Unknown:      '#8A8886',
};

export const PRESENCE_LABEL: Record<PresenceAvailability, string> = {
  Available: 'Available', Busy: 'Busy', DoNotDisturb: 'Do Not Disturb',
  BeRightBack: 'Be Right Back', Away: 'Away', Offline: 'Offline', Unknown: '',
};

export function getInitials(displayName: string): string {
  const parts = (displayName || '').split(' ').filter(p => p.length > 0);
  if (parts.length === 0) return '?';
  if (parts.length === 1) return parts[0][0].toUpperCase();
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

