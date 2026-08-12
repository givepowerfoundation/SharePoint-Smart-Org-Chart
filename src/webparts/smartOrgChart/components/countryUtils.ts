// Country pill colouring and employee-type helpers.
//
// Country colours come from the `countryColors` web-part property, one mapping per line
// (`Kenya=#0b7d5a`). Anything not mapped falls back to a deterministic colour derived from
// the country name, so an unconfigured tenant still gets stable, distinguishable pills.

/** Bucket label for users with no employeeType set. Shared by both views so the
 *  filter and the pill agree on what "unset" is called. */
export const EMPLOYEE_TYPE_UNSET = 'Not set';

/** Expands #abc to #aabbcc and lower-cases. Returns '' if not a valid hex colour.
 *  Always returning 6 digits matters: callers append '1a' for the tint background. */
function normalizeHex(value: string): string {
  const hex = value.trim().replace(/^#/, '');
  if (/^[0-9a-fA-F]{3}$/.test(hex)) {
    const [r, g, b] = hex.toLowerCase().split('');
    return `#${r}${r}${g}${g}${b}${b}`;
  }
  if (/^[0-9a-fA-F]{6}$/.test(hex)) return `#${hex.toLowerCase()}`;
  return '';
}

/**
 * Parses the `countryColors` property into a lookup map.
 * One mapping per line, `Country=#hex` or `Country: #hex`. The leading `#` is optional
 * and 3-digit hex is accepted. Blank lines, comment lines (`//`, `#` alone) and lines
 * with an unparseable colour are skipped rather than throwing — an admin typo in the
 * property pane should degrade to the fallback colour, not break the web part.
 * Keys are trimmed and lower-cased for case-insensitive lookup.
 */
export function parseCountryColors(raw: string | undefined): Map<string, string> {
  const map = new Map<string, string>();
  if (!raw) return map;

  for (const line of raw.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('//')) continue;

    const sepIndex = trimmed.search(/[=:]/);
    if (sepIndex <= 0) continue;

    const name  = trimmed.slice(0, sepIndex).trim().toLowerCase();
    const color = normalizeHex(trimmed.slice(sepIndex + 1));
    if (name && color) map.set(name, color);
  }
  return map;
}

function hslToHex(h: number, s: number, l: number): string {
  const a = s * Math.min(l, 1 - l);
  const channel = (n: number): string => {
    const k = (n + h / 30) % 12;
    const v = l - a * Math.max(-1, Math.min(k - 3, Math.min(9 - k, 1)));
    // Manual pad — the project targets ES5, so String.padStart is unavailable
    const hex = Math.round(v * 255).toString(16);
    return hex.length === 1 ? `0${hex}` : hex;
  };
  return `#${channel(0)}${channel(8)}${channel(4)}`;
}

/**
 * Colour for a country pill: the admin-configured value when present, otherwise a
 * deterministic hue derived from the name. Saturation and lightness are fixed at values
 * that stay legible as pill text on the light tint background used by the card pills.
 * Returns a 6-digit hex string.
 */
export function getCountryColor(country: string, colorMap: Map<string, string>): string {
  const key = (country || '').trim().toLowerCase();
  const configured = colorMap.get(key);
  if (configured) return configured;

  let hash = 0;
  for (let i = 0; i < key.length; i++) hash = (Math.imul(31, hash) + key.charCodeAt(i)) | 0;
  return hslToHex(Math.abs(hash) % 360, 0.55, 0.38);
}
