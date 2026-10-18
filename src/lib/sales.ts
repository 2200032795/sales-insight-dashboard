export type Row = Record<string, string>;
export type Role = 'date' | 'revenue' | 'profit' | 'order' | 'quantity' | 'product' | 'category' | 'status';
export type Mapping = Record<Role, string>;
export type Dataset = { columns: string[]; rows: Row[]; fileName: string };

export const ROLE_META: { role: Role; label: string; hints: string[] }[] = [
  { role: 'date', label: 'Order date', hints: ['date', 'orderdate', 'created', 'time'] },
  { role: 'revenue', label: 'Revenue', hints: ['revenue', 'salesamount', 'salesvalue', 'salesinr', 'sales', 'totalrevenue', 'totalamount', 'amount'] },
  { role: 'profit', label: 'Profit', hints: ['profit', 'margin'] },
  { role: 'order', label: 'Order ID', hints: ['orderid', 'transactionid', 'invoiceid', 'ordernumber', 'transaction', 'invoice'] },
  { role: 'quantity', label: 'Quantity', hints: ['quantity', 'qty', 'units'] },
  { role: 'product', label: 'Product', hints: ['productname', 'product', 'item'] },
  { role: 'category', label: 'Category', hints: ['category', 'department', 'type'] },
  { role: 'status', label: 'Order status', hints: ['status', 'state'] },
];

export function inferMapping(columns: string[]): Mapping {
  const result = {} as Mapping;
  for (const { role, hints } of ROLE_META) {
    const normalizedHints = hints.map(hint => hint.toLowerCase().replace(/[^a-z0-9]/g, ''));
    const candidates = columns.map((column, index) => {
      const normalized = column.toLowerCase().replace(/[^a-z0-9]/g, '');
      const matchedHint = normalizedHints
        .filter(hint => normalized === hint || normalized.includes(hint))
        .sort((a, b) => b.length - a.length)[0];
      if (role === 'revenue' && matchedHint === 'sales'
        && normalized !== 'sales'
        && !/^sales(?:inr|amount|value|revenue|total|net|gross)/.test(normalized)) {
        return { column, index, score: 0 };
      }
      if (!matchedHint) return { column, index, score: 0 };
      const exact = normalized === matchedHint;
      const startsWith = normalized.startsWith(matchedHint);
      const score = (exact ? 100 : 0) + matchedHint.length * 2 + (startsWith ? 1 : 0);
      return { column, index, score };
    });
    const best = candidates.sort((a, b) => b.score - a.score || a.index - b.index)[0];
    result[role] = best?.score ? best.column : '';
  }
  return result;
}

export function toNumber(value: string | undefined): number | null {
  if (value == null || value.trim() === '') return null;
  const normalized = value.replace(/[₹,\s]/g, '').replace(/^\((.*)\)$/, '-$1');
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

export const currency = (value: number) => new Intl.NumberFormat('en-IN', {
  style: 'currency', currency: 'INR', maximumFractionDigits: 0,
}).format(value);
export const numberFormat = (value: number) => new Intl.NumberFormat('en-IN', { maximumFractionDigits: 1 }).format(value);

export function parseDate(value: string | undefined): Date | null {
  if (!value?.trim()) return null;
  const normalized = value.trim();
  const isoDate = /^(\d{4})-(\d{2})-(\d{2})(?:[Tt ].*)?$/.exec(normalized);
  if (isoDate) {
    const year = Number(isoDate[1]);
    const month = Number(isoDate[2]);
    const day = Number(isoDate[3]);
    const date = new Date(normalized);
    const validatedDate = new Date(Date.UTC(year, month - 1, day));
    if (validatedDate.getUTCFullYear() !== year || validatedDate.getUTCMonth() + 1 !== month
      || validatedDate.getUTCDate() !== day) return null;
    return Number.isNaN(date.getTime()) ? null : date;
  }
  if (!/^\d{1,2}[/. -]\d{1,2}[/. -]\d{2,4}$/.test(normalized)) return null;
  const date = new Date(normalized);
  return Number.isNaN(date.getTime()) ? null : date;
}
export function dateLabel(value: string) {
  const date = parseDate(value);
  return date ? new Intl.DateTimeFormat('en-IN', { day: '2-digit', month: 'short', year: 'numeric' }).format(date) : value || '—';
}

export function csvString(rows: Row[], columns: string[]) {
  const quote = (v: string) => `"${String(v ?? '').replace(/"/g, '""')}"`;
  return [columns.map(quote).join(','), ...rows.map(row => columns.map(col => quote(row[col] ?? '')).join(','))].join('\r\n');
}

export function qualitySummary(dataset: Dataset) {
  const missing = Object.fromEntries(dataset.columns.map(col => [col, dataset.rows.filter(row => !row[col]?.trim()).length]));
  const signatures = new Set<string>();
  let duplicates = 0;
  for (const row of dataset.rows) {
    const signature = JSON.stringify(dataset.columns.map(col => row[col] ?? ''));
    if (signatures.has(signature)) duplicates += 1;
    signatures.add(signature);
  }
  const types = Object.fromEntries(dataset.columns.map(col => {
    const values = dataset.rows.map(row => row[col]).filter(value => value?.trim());
    if (values.length && values.every(value => toNumber(value) !== null)) return [col, 'Number'];
    if (values.length && values.every(value => parseDate(value) !== null)) return [col, 'Date'];
    return [col, 'Text'];
  }));
  return { missing, duplicates, types };
}
