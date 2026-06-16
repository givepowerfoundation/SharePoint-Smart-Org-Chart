import html2canvas from 'html2canvas';
import { IGraphUser, IOrgNode } from './GraphService';

/* ── Shared download / CSV helpers ── */

function downloadBlob(content: string, filename: string, mime: string): void {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Prefix cells that Excel would interpret as formulas (CSV injection mitigation)
function csvCell(c: string): string {
  const safe = /^[=+\-@\t\r]/.test(c) ? `'${c}` : c;
  return `"${safe.replace(/"/g, '""')}"`;
}

function csvRow(cells: string[]): string {
  return cells.map(csvCell).join(',');
}

// UTF-8 BOM ensures Excel opens the file with correct character encoding
const CSV_BOM = '﻿';

export interface IDirectoryExportOptions {
  showEmail: boolean;
  showPhone: boolean;
  showDepartment: boolean;
  showOffice: boolean;
}

const BASE_STYLES = `
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Segoe UI, Arial, sans-serif; font-size: 11pt; color: #333; }
  h1 { font-size: 18pt; color: #0078d4; margin-bottom: 4px; }
  .subtitle { font-size: 9pt; color: #666; margin-bottom: 20px; }
  @page { margin: 15mm; }
  @media print {
    body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  }
`;

function openPrintWindow(html: string): void {
  const win = window.open('', '_blank', 'width=900,height=700');
  if (!win) { alert('Please allow pop-ups to export PDF.'); return; }
  win.document.write(html);
  win.document.close();
  win.focus();
  win.onload = () => win.print();
}

export function exportDirectoryToPdf(users: IGraphUser[], opts: IDirectoryExportOptions): void {
  if (users.length === 0) return;

  const cols: { label: string; get: (u: IGraphUser) => string }[] = [
    { label: 'Name',      get: u => u.displayName || '' },
    { label: 'Job Title', get: u => u.jobTitle || '' },
  ];
  if (opts.showDepartment) cols.push({ label: 'Department', get: u => u.department || '' });
  if (opts.showOffice)     cols.push({ label: 'Office',     get: u => u.officeLocation || '' });
  if (opts.showEmail)      cols.push({ label: 'Email',      get: u => u.mail || '' });
  if (opts.showPhone)      cols.push({ label: 'Phone',      get: u => u.mobilePhone || (u.businessPhones && u.businessPhones[0]) || '' });

  const headers = cols.map(c => `<th>${c.label}</th>`).join('');
  const rows = users.map((u, i) => {
    const cells = cols.map(c => `<td>${escHtml(c.get(u))}</td>`).join('');
    return `<tr class="${i % 2 === 0 ? 'even' : 'odd'}">${cells}</tr>`;
  }).join('');

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Employee Directory</title>
<style>
${BASE_STYLES}
table { width: 100%; border-collapse: collapse; }
th { background: #0078d4; color: #fff; text-align: left; padding: 6px 8px; font-size: 9pt; }
td { padding: 5px 8px; font-size: 9pt; border-bottom: 1px solid #eee; vertical-align: top; }
tr.odd td { background: #f7f9fc; }
</style></head><body>
<h1>Employee Directory</h1>
<div class="subtitle">${users.length} employee${users.length !== 1 ? 's' : ''} &nbsp;·&nbsp; ${new Date().toLocaleDateString()}</div>
<table><thead><tr>${headers}</tr></thead><tbody>${rows}</tbody></table>
</body></html>`;

  openPrintWindow(html);
}

export function exportDirectoryToExcel(users: IGraphUser[], opts: IDirectoryExportOptions): void {
  if (users.length === 0) return;

  const cols: { label: string; get: (u: IGraphUser) => string }[] = [
    { label: 'Name',      get: u => u.displayName || '' },
    { label: 'Job Title', get: u => u.jobTitle || '' },
  ];
  if (opts.showDepartment) cols.push({ label: 'Department', get: u => u.department || '' });
  if (opts.showOffice)     cols.push({ label: 'Office',     get: u => u.officeLocation || '' });
  if (opts.showEmail)      cols.push({ label: 'Email',      get: u => u.mail || '' });
  if (opts.showPhone)      cols.push({ label: 'Phone',      get: u => u.mobilePhone || (u.businessPhones && u.businessPhones[0]) || '' });

  const lines = [
    csvRow(cols.map(c => c.label)),
    ...users.map(u => csvRow(cols.map(c => c.get(u)))),
  ];

  downloadBlob(
    CSV_BOM + lines.join('\r\n'),
    `employee-directory-${new Date().toISOString().slice(0, 10)}.csv`,
    'text/csv;charset=utf-8;'
  );
}

export function exportOrgChartToCsv(rootNode: IOrgNode): void {
  const rows: string[] = [
    csvRow(['Name', 'Job Title', 'Department', 'Office', 'Email', 'Phone', 'Manager', 'Level']),
  ];
  const visit = (n: IOrgNode, managerName: string, depth: number): void => {
    const u = n.user;
    rows.push(csvRow([
      u.displayName || '',
      u.jobTitle || '',
      u.department || '',
      u.officeLocation || '',
      u.mail || '',
      u.mobilePhone || (u.businessPhones && u.businessPhones[0]) || '',
      managerName,
      String(depth),
    ]));
    n.directReports.forEach(c => visit(c, u.displayName || '', depth + 1));
  };
  visit(rootNode, '', 0);

  downloadBlob(
    CSV_BOM + rows.join('\r\n'),
    `org-chart-${new Date().toISOString().slice(0, 10)}.csv`,
    'text/csv;charset=utf-8;'
  );
}

export async function exportOrgChartToPng(treeEl: HTMLElement): Promise<void> {
  try {
    const canvas = await html2canvas(treeEl, {
      scale: 2,
      useCORS: true,
      allowTaint: true,
      backgroundColor: '#ffffff',
      logging: false,
    });
    const url = canvas.toDataURL('image/png');
    const a = document.createElement('a');
    a.href = url;
    a.download = `org-chart-${new Date().toISOString().slice(0, 10)}.png`;
    a.click();
  } catch {
    alert('Could not capture the org chart image. Try reducing zoom first.');
  }
}

export function exportVCard(user: IGraphUser): void {
  // Escape vCard special characters (backslash, comma, semicolon)
  const esc = (s: string): string => s.replace(/\\/g, '\\\\').replace(/[,;]/g, m => `\\${m}`);
  const parts = (user.displayName || '').split(' ').filter(p => p);
  const first = parts[0] || '';
  const last  = parts.length > 1 ? parts[parts.length - 1] : '';

  const lines = ['BEGIN:VCARD', 'VERSION:3.0', `FN:${esc(user.displayName || '')}`, `N:${esc(last)};${esc(first)};;;`];
  if (user.jobTitle)       lines.push(`TITLE:${esc(user.jobTitle)}`);
  if (user.department)     lines.push(`ORG:${esc(user.department)}`);
  if (user.mail)           lines.push(`EMAIL;TYPE=WORK:${user.mail}`);
  if (user.businessPhones && user.businessPhones[0]) lines.push(`TEL;TYPE=WORK,VOICE:${user.businessPhones[0]}`);
  if (user.mobilePhone)    lines.push(`TEL;TYPE=CELL:${user.mobilePhone}`);
  if (user.officeLocation) lines.push(`ADR;TYPE=WORK:;;${esc(user.officeLocation)};;;;`);
  lines.push('END:VCARD');

  const safeName = (user.displayName || 'contact').replace(/[^\w\- ]/g, '').trim() || 'contact';
  downloadBlob(lines.join('\r\n'), `${safeName}.vcf`, 'text/vcard;charset=utf-8;');
}

export function exportOrgChartToPdf(rootNode: IOrgNode, note?: string): void {
  const lines: string[] = [];

  const renderNode = (node: IOrgNode, depth: number): void => {
    const indent = depth * 20;
    const isRoot = depth === 0;
    const weight = isRoot ? 'bold' : depth === 1 ? '600' : 'normal';
    const size = isRoot ? '13pt' : depth === 1 ? '11pt' : '10pt';
    const color = isRoot ? '#0078d4' : '#222';

    let meta = '';
    if (node.user.jobTitle || node.user.department) {
      const parts = [node.user.jobTitle, node.user.department].filter(Boolean);
      meta = `<div style="font-size:8.5pt;color:#666;margin-top:2px;">${escHtml(parts.join(' · '))}</div>`;
    }

    lines.push(
      `<div style="margin-left:${indent}px;margin-bottom:${isRoot ? 10 : 5}px;padding-left:${depth > 0 ? 10 : 0}px;border-left:${depth > 0 ? '2px solid #e0e0e0' : 'none'};">` +
      `<span style="font-weight:${weight};font-size:${size};color:${color};">${escHtml(node.user.displayName)}</span>` +
      meta +
      `</div>`
    );

    node.directReports.forEach(child => renderNode(child, depth + 1));
  };

  renderNode(rootNode, 0);

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Org Chart</title>
<style>${BASE_STYLES}</style></head><body>
<h1>Organization Chart</h1>
<div class="subtitle">Root: ${escHtml(rootNode.user.displayName)} &nbsp;·&nbsp; ${new Date().toLocaleDateString()}${note ? ` &nbsp;·&nbsp; ${escHtml(note)}` : ''}</div>
${lines.join('')}
</body></html>`;

  openPrintWindow(html);
}

function escHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
