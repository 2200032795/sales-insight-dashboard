import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import { Link, Route, Switch, useLocation } from 'wouter';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import {
  Activity, AlertCircle, ArrowDown, ArrowUp, BarChart3, Check, CircleHelp,
  Download, FileSpreadsheet, Filter, LayoutDashboard, Menu, Package, RotateCcw,
  Search, Settings2, ShieldCheck, SlidersHorizontal, Tag, Upload, X,
} from 'lucide-react';
import {
  Area, AreaChart, Bar, BarChart, CartesianGrid, Cell, Pie, PieChart, ResponsiveContainer,
  Tooltip, XAxis, YAxis,
} from 'recharts';
import type { Dataset, Mapping, Role, Row } from './lib/sales';
import { currency, dateLabel, inferMapping, numberFormat, parseDate, qualitySummary, ROLE_META, toNumber } from './lib/sales';

const nav = [
  { href: '/', label: 'Dashboard', icon: LayoutDashboard },
  { href: '/analysis', label: 'Sales Analysis', icon: Activity },
  { href: '/products', label: 'Products', icon: Package },
  { href: '/categories', label: 'Categories', icon: Tag },
  { href: '/quality', label: 'Data Quality', icon: ShieldCheck },
  { href: '/settings', label: 'Settings', icon: Settings2 },
];
const chartColors = ['#513B2F', '#879475', '#D6B58C', '#A56854', '#769093', '#C9986B'];
type Filters = { from: string; to: string; category: string; product: string; status: string };
const defaultFilters: Filters = { from: '', to: '', category: '', product: '', status: '' };

function parseFile(file: File): Promise<Dataset> {
  return new Promise((resolve, reject) => {
    const ext = file.name.split('.').pop()?.toLowerCase();
    const finish = (headers: string[], values: unknown[][]) => {
      const columns = headers.map(value => String(value ?? '').trim()).filter(Boolean);
      if (!columns.length) return reject(new Error('No column headings were found in this file.'));
      const rows = values.filter(row => row.some(value => String(value ?? '').trim() !== '')).map(valuesRow => {
        const row: Row = {};
        columns.forEach((column, index) => { row[column] = String(valuesRow[index] ?? '').trim(); });
        return row;
      });
      if (!rows.length) return reject(new Error('This file has headings but no data rows.'));
      resolve({ columns, rows, fileName: file.name });
    };
    if (ext === 'csv') {
      Papa.parse<string[]>(file, {
        skipEmptyLines: 'greedy',
        complete: result => {
          if (result.errors.length && result.data.length === 0) return reject(new Error(result.errors[0].message));
          const [head, ...body] = result.data;
          finish((head ?? []).map(String), body);
        },
        error: error => reject(error),
      });
    } else if (ext === 'xlsx' || ext === 'xls') {
      const reader = new FileReader();
      reader.onload = () => {
        try {
          const workbook = XLSX.read(reader.result, { type: 'array', cellDates: true });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          if (!sheet) return reject(new Error('No worksheet was found.'));
          const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, raw: false, defval: '' });
          const [header, ...body] = matrix;
          finish((header ?? []).map(value => String(value)), body);
        } catch { reject(new Error('The spreadsheet could not be read. Please check the file and try again.')); }
      };
      reader.onerror = () => reject(new Error('The file could not be read.'));
      reader.readAsArrayBuffer(file);
    } else reject(new Error('Choose a CSV, XLS, or XLSX file.'));
  });
}

function App() {
  const [dataset, setDataset] = useState<Dataset | null>(null);
  const [mapping, setMapping] = useState<Mapping>({} as Mapping);
  const [filters, setFilters] = useState<Filters>(defaultFilters);
  const [loading, setLoading] = useState(true);
  const [fileError, setFileError] = useState('');
  const [fileBusy, setFileBusy] = useState(false);
  const [mobileNav, setMobileNav] = useState(false);
  const uploadRef = useRef<HTMLInputElement>(null);
  const [location] = useLocation();

  useEffect(() => {
    fetch(`${import.meta.env.BASE_URL}sales-data.csv`)
      .then(response => {
        if (!response.ok) throw new Error('The starter dataset could not be loaded.');
        return response.text();
      })
      .then(text => {
        Papa.parse<string[]>(text, { skipEmptyLines: 'greedy', complete: result => {
          const [header, ...body] = result.data;
          const columns = (header ?? []).map(value => String(value).trim());
          const rows = body.filter(row => row.some(value => String(value ?? '').trim())).map(line => Object.fromEntries(columns.map((column, index) => [column, String(line[index] ?? '').trim()])));
          if (!rows.length) { setFileError('The starter dataset is empty. Replace it with a valid file.'); return; }
          const next = { columns, rows, fileName: 'sales-data.csv' };
          setDataset(next); setMapping(inferMapping(columns));
        }, error: (error: Error) => setFileError(error.message), });
      })
      .catch(error => setFileError(error.message || 'Unable to load the starter dataset.'))
      .finally(() => setLoading(false));
  }, []);

  const replaceDataset = async (file?: File) => {
    if (!file) return;
    setFileError(''); setFileBusy(true);
    try {
      const next = await parseFile(file);
      setDataset(next); setMapping(inferMapping(next.columns)); setFilters(defaultFilters);
    } catch (error) { setFileError(error instanceof Error ? error.message : 'Could not parse this file.'); }
    finally { setFileBusy(false); if (uploadRef.current) uploadRef.current.value = ''; }
  };
  const filteredRows = useMemo(() => {
    if (!dataset) return [];
    return dataset.rows.filter(row => {
      const date = mapping.date ? parseDate(row[mapping.date]) : null;
      if (filters.from && date && date < new Date(`${filters.from}T00:00:00`)) return false;
      if (filters.to && date && date > new Date(`${filters.to}T23:59:59`)) return false;
      if (filters.from && !date) return false;
      if (filters.to && !date) return false;
      if (filters.category && row[mapping.category] !== filters.category) return false;
      if (filters.product && row[mapping.product] !== filters.product) return false;
      if (filters.status && row[mapping.status] !== filters.status) return false;
      return true;
    });
  }, [dataset, mapping, filters]);
  const categories = useMemo(() => dataset && mapping.category ? [...new Set(dataset.rows.map(row => row[mapping.category]).filter(Boolean))].sort() : [], [dataset, mapping.category]);
  const products = useMemo(() => dataset && mapping.product ? [...new Set(dataset.rows.map(row => row[mapping.product]).filter(Boolean))].sort() : [], [dataset, mapping.product]);
  const statuses = useMemo(() => dataset && mapping.status ? [...new Set(dataset.rows.map(row => row[mapping.status]).filter(Boolean))].sort() : [], [dataset, mapping.status]);

  const exportRows = (rows = filteredRows) => {
    if (!dataset) return;
    const blob = new Blob(['\ufeff' + [dataset.columns.map(escapeCsv).join(','), ...rows.map(row => dataset.columns.map(column => escapeCsv(row[column] ?? '')).join(','))].join('\r\n')], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a'); anchor.href = url; anchor.download = `${dataset.fileName.replace(/\.[^.]+$/, '')}-filtered.csv`; anchor.click(); URL.revokeObjectURL(url);
  };
  const shellContext: DashboardContext = {
    dataset, mapping, setMapping, filteredRows, allRows: dataset?.rows ?? [], filters, setFilters, categories, products, statuses,
    loading, fileError, fileBusy, chooseFile: () => uploadRef.current?.click(), replaceDataset, exportRows,
  };

  return <div className="min-h-[100dvh] bg-[#F7F2E8]">
    <input ref={uploadRef} className="hidden" type="file" accept=".csv,.xls,.xlsx" onChange={event => void replaceDataset(event.target.files?.[0])} data-testid="input-dataset-file" />
    <aside className={`fixed inset-y-0 left-0 z-40 flex w-[248px] flex-col border-r border-[#513B2F]/10 bg-[#F3EDE1] px-5 py-6 transition-transform duration-300 md:translate-x-0 ${mobileNav ? 'translate-x-0' : '-translate-x-full md:translate-x-0'}`}>
      <div className="mb-10 flex items-center gap-3 px-1">
        <div className="grid h-10 w-10 place-items-center rounded-xl bg-[#513B2F] text-[#F7F2E8]"><BarChart3 size={20}/></div>
        <div><div className="font-display text-[19px] leading-tight">Fieldnotes</div><div className="mt-1 text-[10px] uppercase tracking-[.2em] text-[#806D5C]">Sales intelligence</div></div>
        <button onClick={() => setMobileNav(false)} aria-label="Close navigation" className="ml-auto md:hidden"><X size={18}/></button>
      </div>
      <div className="mb-3 px-3 text-[10px] font-bold uppercase tracking-[.18em] text-[#9A8874]">Workspace</div>
      <nav className="space-y-1" aria-label="Main navigation">{nav.map(item => {
        const Icon = item.icon; const active = location === item.href;
        return <Link key={item.href} href={item.href} onClick={() => setMobileNav(false)} data-testid={`link-nav-${item.label.toLowerCase().replaceAll(' ','-')}`} className={`group flex items-center gap-3 rounded-xl px-3 py-3 text-[13px] transition-all duration-200 ${active ? 'bg-[#513B2F] text-[#F7F2E8] shadow-md shadow-[#513B2F]/10' : 'text-[#6D594B] hover:translate-x-0.5 hover:bg-[#E9E0D1]'}`}>
          <Icon size={17} strokeWidth={1.8}/><span className="flex-1">{item.label}</span>{active && <span className="h-1.5 w-1.5 rounded-full bg-[#D6B58C]"/>}
        </Link>;
      })}</nav>
      <div className="mt-auto rounded-2xl border border-[#513B2F]/10 bg-[#EAE2D3] p-4">
        <div className="mb-2 flex items-center gap-2 text-[11px] font-semibold text-[#513B2F]"><ShieldCheck size={14}/> Private workspace</div>
        <p className="text-[11px] leading-relaxed text-[#806D5C]">Your dataset stays in this browser. No live system connection.</p>
      </div>
      <div className="mt-4 flex items-center gap-2 px-1 text-[11px] text-[#988574]"><span className="h-2 w-2 rounded-full bg-[#879475]"/> Local dataset · {dataset?.rows.length ?? '—'} rows</div>
    </aside>
    {mobileNav && <button onClick={() => setMobileNav(false)} aria-label="Dismiss navigation" className="fixed inset-0 z-30 bg-[#37291f]/30 md:hidden"/>}
    <main className="min-h-[100dvh] md:ml-[248px]">
      <header className="sticky top-0 z-20 flex h-[68px] items-center justify-between border-b border-[#513B2F]/10 bg-[#F7F2E8]/90 px-5 backdrop-blur-md sm:px-8 lg:px-10">
        <div className="flex items-center gap-3">
          <button onClick={() => setMobileNav(true)} aria-label="Open navigation" className="rounded-lg p-2 hover:bg-[#EAE2D3] md:hidden"><Menu size={19}/></button>
          <div className="hidden text-[11px] text-[#927F6D] sm:block">Portfolio workspace <span className="mx-2 text-[#D6B58C]">/</span> <span className="text-[#513B2F]">{nav.find(item => item.href === location)?.label ?? 'Dashboard'}</span></div>
          <div className="text-[12px] text-[#513B2F] sm:hidden">{nav.find(item => item.href === location)?.label ?? 'Dashboard'}</div>
        </div>
        <div className="flex items-center gap-3">
          <span className="hidden items-center gap-1.5 rounded-full border border-[#879475]/35 bg-[#E7E9DF] px-3 py-1.5 text-[10px] font-semibold text-[#53624B] sm:flex"><span className="h-1.5 w-1.5 rounded-full bg-[#879475]"/> Local data</span>
          <button onClick={() => uploadRef.current?.click()} disabled={fileBusy} className="flex items-center gap-2 rounded-lg bg-[#513B2F] px-3.5 py-2.5 text-[11px] font-semibold text-[#F7F2E8] transition hover:bg-[#684D3D] disabled:opacity-60" data-testid="button-replace-dataset"><Upload size={14}/><span className="hidden sm:inline">{fileBusy ? 'Reading file…' : 'Replace dataset'}</span><span className="sm:hidden">Upload</span></button>
        </div>
      </header>
      <div className="mx-auto max-w-[1500px] px-4 py-7 sm:px-7 sm:py-9 lg:px-10">
        {fileError && <div role="alert" className="mb-5 flex items-center gap-3 rounded-xl border border-[#B76E60]/25 bg-[#F6E8E2] px-4 py-3 text-[13px] text-[#82493E]"><AlertCircle size={17}/><span className="flex-1">{fileError}</span><button onClick={() => setFileError('')} aria-label="Dismiss error"><X size={16}/></button></div>}
        {loading ? <LoadingView/> : !dataset ? <MissingDataset onUpload={() => uploadRef.current?.click()} error={fileError}/> :
          <><FilterBar context={shellContext}/><Switch>
            <Route path="/"><DashboardPage context={shellContext}/></Route>
            <Route path="/analysis"><AnalysisPage context={shellContext}/></Route>
            <Route path="/products"><ProductsPage context={shellContext}/></Route>
            <Route path="/categories"><CategoriesPage context={shellContext}/></Route>
            <Route path="/quality"><QualityPage context={shellContext}/></Route>
            <Route path="/settings"><SettingsPage context={shellContext}/></Route>
            <Route><DashboardPage context={shellContext}/></Route>
          </Switch></>}
      </div>
    </main>
  </div>;
}

type DashboardContext = {
  dataset: Dataset | null; mapping: Mapping; setMapping: (mapping: Mapping) => void;
  filteredRows: Row[]; allRows: Row[]; filters: Filters; setFilters: (filters: Filters) => void;
  categories: string[]; products: string[]; statuses: string[]; loading: boolean;
  fileError: string; fileBusy: boolean; chooseFile: () => void;
  replaceDataset: (file?: File) => Promise<void>; exportRows: (rows?: Row[]) => void;
};
function escapeCsv(value: string) { return `"${String(value).replace(/"/g, '""')}"`; }
function LoadingView() { return <div className="space-y-6 page-enter"><div className="h-7 w-48 animate-pulse rounded bg-[#E6DDCD]"/><div className="grid grid-cols-2 gap-4 lg:grid-cols-4">{[1,2,3,4].map(i=><div key={i} className="panel h-32 animate-pulse bg-[#F4EDE0]"/> )}</div><div className="panel h-[330px] animate-pulse bg-[#F4EDE0]"/></div>; }
function MissingDataset({ onUpload, error }: { onUpload: () => void; error: string }) { return <div className="panel mx-auto mt-16 max-w-xl p-10 text-center"><div className="mx-auto mb-4 grid h-14 w-14 place-items-center rounded-2xl bg-[#EAE2D3] text-[#513B2F]"><FileSpreadsheet/></div><h2 className="font-display text-2xl">A dataset is needed</h2><p className="mx-auto mt-2 max-w-sm text-sm leading-6 text-[#806D5C]">{error || 'Load a CSV or Excel workbook to begin exploring your sales records.'}</p><button onClick={onUpload} className="mt-6 rounded-lg bg-[#513B2F] px-5 py-3 text-sm text-[#F7F2E8]">Choose a file</button></div>; }

function FilterBar({ context }: { context: DashboardContext }) {
  const { filters, setFilters, categories, products, statuses, mapping, filteredRows, exportRows } = context;
  const set = (key: keyof Filters, value: string) => setFilters({ ...filters, [key]: value });
  const active = Object.values(filters).some(Boolean);
  return <div className="mb-8 flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
    <div className="flex items-center gap-2 text-[11px] font-semibold uppercase tracking-[.14em] text-[#927F6D]"><Filter size={14}/> Filter view <span className="font-normal normal-case tracking-normal text-[#AD9C88]">{filteredRows.length} of {context.allRows.length} rows</span></div>
    <div className="flex flex-wrap items-center gap-2">
      {mapping.date && <><label className="sr-only" htmlFor="date-from">Date from</label><input id="date-from" aria-label="Date from" type="date" value={filters.from} onChange={e=>set('from',e.target.value)} className="filter-control" data-testid="input-date-from"/><span className="text-[11px] text-[#A18F7C]">to</span><label className="sr-only" htmlFor="date-to">Date to</label><input id="date-to" aria-label="Date to" type="date" value={filters.to} onChange={e=>set('to',e.target.value)} className="filter-control" data-testid="input-date-to"/></>}
      {mapping.category && <select aria-label="Category filter" value={filters.category} onChange={e=>set('category',e.target.value)} className="filter-control" data-testid="select-category-filter"><option value="">All categories</option>{categories.map(v=><option key={v}>{v}</option>)}</select>}
      {mapping.product && <select aria-label="Product filter" value={filters.product} onChange={e=>set('product',e.target.value)} className="filter-control max-w-40" data-testid="select-product-filter"><option value="">All products</option>{products.map(v=><option key={v}>{v}</option>)}</select>}
      {mapping.status && <select aria-label="Status scope" value={filters.status} onChange={e=>set('status',e.target.value)} className="filter-control" data-testid="select-status-filter"><option value="">All statuses</option>{statuses.map(v=><option key={v}>{v}</option>)}</select>}
      {active && <button onClick={()=>setFilters(defaultFilters)} className="flex items-center gap-1.5 rounded-lg px-2.5 py-2 text-[11px] text-[#806D5C] hover:bg-[#EAE2D3]" data-testid="button-clear-filters"><RotateCcw size={13}/> Clear</button>}
      <button onClick={()=>exportRows()} className="flex items-center gap-2 rounded-lg border border-[#513B2F]/15 bg-[#FBF8F1] px-3 py-2 text-[11px] font-semibold text-[#513B2F] transition hover:bg-[#EAE2D3]" data-testid="button-export-filtered"><Download size={14}/> Export</button>
    </div>
  </div>;
}

type Metric = { revenue: number; profit: number; quantity: number; orders: number; aov: number; completed: number };
function metrics(rows: Row[], mapping: Mapping): Metric {
  const revenue = rows.reduce((sum,row)=>sum+(toNumber(mapping.revenue ? row[mapping.revenue] : undefined) ?? 0),0);
  const profit = rows.reduce((sum,row)=>sum+(toNumber(mapping.profit ? row[mapping.profit] : undefined) ?? 0),0);
  const quantity = rows.reduce((sum,row)=>sum+(toNumber(mapping.quantity ? row[mapping.quantity] : undefined) ?? 0),0);
  const orders = mapping.order ? new Set(rows.map(row=>row[mapping.order]).filter(Boolean)).size : rows.length;
  const completed = mapping.status ? rows.filter(row=>row[mapping.status]?.toLowerCase()==='completed').length : rows.length;
  return { revenue, profit, quantity, orders, aov: orders ? revenue/orders : 0, completed };
}
function hasValue(mapping: Mapping, role: Role) { return Boolean(mapping[role]); }
function PageHeading({ eyebrow, title, description, action }: { eyebrow: string; title: string; description: string; action?: ReactNode }) {
  return <div className="mb-7 flex flex-wrap items-end justify-between gap-4 page-enter"><div><div className="mb-2 text-[10px] font-bold uppercase tracking-[.2em] text-[#927F6D]">{eyebrow}</div><h1 className="font-display text-[32px] leading-tight tracking-[-.025em] sm:text-[38px]">{title}</h1><p className="mt-2 max-w-2xl text-[13px] leading-6 text-[#806D5C]">{description}</p></div>{action}</div>;
}
function Kpi({ title, value, caption, icon: Icon, tone = 'brown' }: { title: string; value: string; caption: string; icon: typeof Activity; tone?: string }) {
  return <div className="panel group relative overflow-hidden p-5 transition duration-200 hover:-translate-y-0.5 hover:shadow-lg hover:shadow-[#513B2F]/[.07]"><div className="absolute -right-4 -top-6 h-24 w-24 rounded-full bg-[#E9E1D3]/55 transition-transform duration-500 group-hover:scale-125"/><div className="relative flex items-start justify-between"><div><div className="text-[11px] font-medium text-[#806D5C]">{title}</div><div className={`mt-3 font-display text-[27px] leading-none tracking-tight ${tone==='sage'?'text-[#657257]':'text-[#513B2F]'}`}>{value}</div><div className="mt-3 text-[10px] leading-4 text-[#9B8876]">{caption}</div></div><span className="grid h-9 w-9 place-items-center rounded-xl bg-[#F2EBDD] text-[#826B57]"><Icon size={17}/></span></div></div>;
}
function ScopeNote({ mapping }: { mapping: Mapping }) { return <div className="flex items-start gap-2 rounded-xl border border-[#D6B58C]/35 bg-[#F3E9D8]/65 px-4 py-3 text-[11px] leading-5 text-[#765F4B]"><CircleHelp size={15} className="mt-0.5 shrink-0"/>Totals use every record in the current filter scope, including pending, returned and cancelled statuses. {mapping.status && 'Use the status filter to inspect a specific status; no status is excluded by default.'}</div>; }

function ChartPanel({ title, subtitle, children, className = '' }: { title: string; subtitle: string; children: ReactNode; className?: string }) {
  return <section className={`panel p-5 sm:p-6 ${className}`}><div className="mb-4"><h2 className="font-display text-[19px]">{title}</h2><p className="mt-1 text-[11px] text-[#978574]">{subtitle}</p></div>{children}</section>;
}
function TrendChart({ rows, mapping, group = 'month' }: { rows: Row[]; mapping: Mapping; group?: 'month'|'day' }) {
  const data = useMemo(() => {
    if (!mapping.date || !mapping.revenue) return [];
    const groups = new Map<string, { revenue: number; profit: number; date: Date }>();
    rows.forEach(row => {
      const date = parseDate(row[mapping.date]); if (!date) return;
      const key = group === 'month' ? `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}` : `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}`;
      const current = groups.get(key) ?? { revenue: 0, profit: 0, date };
      current.revenue += toNumber(row[mapping.revenue]) ?? 0;
      current.profit += mapping.profit ? toNumber(row[mapping.profit]) ?? 0 : 0;
      groups.set(key,current);
    });
    return [...groups.entries()].sort(([a],[b])=>a.localeCompare(b)).map(([key,value])=>({ key, label: group==='month' ? new Intl.DateTimeFormat('en',{month:'short',year:'2-digit'}).format(value.date) : dateLabel(key).replace(/, \d{4}$/,''), ...value }));
  }, [rows,mapping,group]);
  if (!mapping.date || !mapping.revenue) return <NoChart reason={!mapping.date ? 'Map a date column to view trends.' : 'Map a revenue column to view sales trends.'}/>;
  if (!data.length) return <NoChart reason="No valid date and revenue values in this filter scope." />;
  return <ResponsiveContainer width="100%" height={260}><AreaChart data={data} margin={{top:8,right:8,bottom:0,left:0}}><defs><linearGradient id={`revenueFill-${group}`} x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#879475" stopOpacity={.37}/><stop offset="100%" stopColor="#879475" stopOpacity={.02}/></linearGradient></defs><CartesianGrid vertical={false} stroke="#E8E0D2" strokeDasharray="3 5"/><XAxis dataKey="label" tick={{fontSize:10,fill:'#947F6B'}} axisLine={false} tickLine={false}/><YAxis tickFormatter={v=>`₹${Number(v)>=1000?`${(Number(v)/1000).toFixed(0)}k`:v}`} tick={{fontSize:10,fill:'#947F6B'}} axisLine={false} tickLine={false} width={48}/><Tooltip formatter={(value: number)=>currency(value)} contentStyle={{borderRadius:12,border:'1px solid #e6ddcd',background:'#fffdf8',fontSize:12}}/><Area type="monotone" dataKey="revenue" name="Revenue" stroke="#71805F" strokeWidth={2.5} fill={`url(#revenueFill-${group})`} activeDot={{r:4,fill:'#513B2F'}} isAnimationActive={false}/>{mapping.profit&&<Area type="monotone" dataKey="profit" name="Profit" stroke="#C2915F" strokeWidth={2} fill="transparent" activeDot={{r:3}} isAnimationActive={false}/>}</AreaChart></ResponsiveContainer>;
}
function NoChart({ reason }: { reason: string }) { return <div className="grid h-[250px] place-items-center rounded-xl border border-dashed border-[#D6CBB9] bg-[#F8F4EC] px-6 text-center text-[12px] text-[#8C7967]"><div><SlidersHorizontal size={22} className="mx-auto mb-2 opacity-55"/>{reason}</div></div>; }
function CategoryDonut({ rows, mapping }: { rows: Row[]; mapping: Mapping }) {
  const data = useMemo(() => {
    if(!mapping.category) return [];
    const values = new Map<string,number>();
    rows.forEach(row => { const key=row[mapping.category]?.trim(); if(key) values.set(key,(values.get(key)??0)+(mapping.revenue?toNumber(row[mapping.revenue])??0:1)); });
    return [...values.entries()].map(([name,value])=>({name,value})).sort((a,b)=>b.value-a.value);
  },[rows,mapping]);
  if(!mapping.category) return <NoChart reason="Map a category column to compare category share."/>;
  if(!data.length) return <NoChart reason="No category values in this filter scope."/>;
  return <div className="grid items-center gap-2 sm:grid-cols-[1fr_1fr]"><ResponsiveContainer width="100%" height={210}><PieChart><Pie data={data} dataKey="value" nameKey="name" innerRadius={56} outerRadius={82} paddingAngle={3} cornerRadius={3} stroke="none" isAnimationActive={false}>{data.map((entry,index)=><Cell key={entry.name} fill={chartColors[index%chartColors.length]}/>)}</Pie><Tooltip formatter={(value:number)=>mapping.revenue?currency(value):numberFormat(value)} contentStyle={{borderRadius:12,border:'1px solid #e6ddcd',background:'#fffdf8',fontSize:12}}/></PieChart></ResponsiveContainer><div className="space-y-3">{data.map((entry,index)=><div key={entry.name} className="flex items-start gap-2.5 text-[11px]"><span className="mt-1 h-2.5 w-2.5 shrink-0 rounded-full" style={{background:chartColors[index%chartColors.length]}}/><span className="min-w-0 flex-1 text-[#806D5C]">{entry.name}</span><span className="font-semibold text-[#513B2F]">{mapping.revenue?currency(entry.value):numberFormat(entry.value)}</span></div>)}</div></div>;
}

function DataTable({ rows, columns, mapping, searchable = true, initialSort, empty = 'No records match the current view.' }: { rows: Row[]; columns: string[]; mapping: Mapping; searchable?: boolean; initialSort?: string; empty?: string }) {
  const [query,setQuery]=useState('');
  const [sort,setSort]=useState(initialSort??columns[0]??'');
  const [direction,setDirection]=useState<'asc'|'desc'>('desc');
  const filtered=useMemo(()=>rows.filter(row=>columns.some(col=>(row[col]??'').toLowerCase().includes(query.toLowerCase()))).slice().sort((a,b)=>{
    const av=a[sort]??'',bv=b[sort]??''; const an=toNumber(av),bn=toNumber(bv);
    const compare=an!==null&&bn!==null?an-bn:av.localeCompare(bv,undefined,{numeric:true,sensitivity:'base'});
    return direction==='asc'?compare:-compare;
  }),[rows,columns,query,sort,direction]);
  const toggle=(column:string)=>{if(sort===column)setDirection(direction==='asc'?'desc':'asc');else{setSort(column);setDirection('desc');}};
  return <div>
    {searchable&&<div className="mb-4 flex flex-wrap items-center justify-between gap-3"><div className="relative w-full max-w-xs"><Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-[#9C8977]"/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Search records…" aria-label="Search records" className="w-full rounded-lg border border-[#513B2F]/15 bg-[#FCFAF5] py-2.5 pl-9 pr-3 text-[12px] outline-none placeholder:text-[#AE9E8C] focus:border-[#879475]" data-testid="input-table-search"/></div><span className="text-[11px] text-[#927F6D]">{filtered.length} record{filtered.length===1?'':'s'}</span></div>}
    {filtered.length===0 ? <div className="rounded-xl border border-dashed border-[#D6CBB9] p-10 text-center text-[13px] text-[#927F6D]">{empty}</div> :
    <div className="overflow-x-auto rounded-xl border border-[#513B2F]/10"><table className="w-full min-w-[650px] border-collapse text-left"><thead className="bg-[#F2EBDD]"><tr>{columns.map(column=><th key={column} className="whitespace-nowrap px-4 py-3 text-[10px] font-bold uppercase tracking-[.12em] text-[#806D5C]"><button onClick={()=>toggle(column)} className="inline-flex items-center gap-1.5 hover:text-[#513B2F]" data-testid={`sort-${column}`}>{column}{sort===column&&(direction==='asc'?<ArrowUp size={11}/>:<ArrowDown size={11}/>)}</button></th>)}</tr></thead><tbody>{filtered.map((row,index)=><tr key={`${index}-${row[mapping.order]??''}`} className="border-t border-[#513B2F]/[.07] transition hover:bg-[#F7F3EA]">{columns.map(column=><td key={column} className="whitespace-nowrap px-4 py-3 text-[12px] text-[#584438]">{column===mapping.date?dateLabel(row[column]):column===mapping.revenue||column===mapping.profit?formatIfNumber(row[column]):column===mapping.status?<StatusPill value={row[column]}/>:row[column]||<span className="text-[#B0A08D]">—</span>}</td>)}</tr>)}</tbody></table></div>}
  </div>;
}
function formatIfNumber(value:string) {const number=toNumber(value);return number===null?value:currency(number);}
function StatusPill({value}:{value:string}) {const status=value?.toLowerCase(); const cls=status==='completed'?'bg-[#E7EBDD] text-[#5A694D]':status==='pending'?'bg-[#F3EAD7] text-[#927243]':status==='returned'?'bg-[#F2E4DF] text-[#965E50]':status==='cancelled'?'bg-[#EAE4E0] text-[#75645B]':'bg-[#ECE7DD] text-[#75645B]';return <span className={`rounded-full px-2.5 py-1 text-[10px] font-semibold capitalize ${cls}`}>{value||'Unknown'}</span>;}

function DashboardPage({context:c}:{context:DashboardContext}) {
  const m=metrics(c.filteredRows,c.mapping);
  const trend=c.filteredRows.length&&c.mapping.date&&c.mapping.revenue?c.filteredRows.filter(row=>parseDate(row[c.mapping.date])).length:0;
  const topProducts=useMemo(()=>aggregateProducts(c.filteredRows,c.mapping).slice(0,5),[c.filteredRows,c.mapping]);
  const columns=[c.mapping.order,c.mapping.date,c.mapping.product,c.mapping.status,c.mapping.revenue].filter(Boolean);
  return <div className="page-enter">
    <PageHeading eyebrow="Portfolio pulse · dataset overview" title="Sales, in perspective." description={`A clear read of ${c.dataset?.fileName ?? 'your dataset'} — every figure traces back to your uploaded rows.`} action={<div className="rounded-xl border border-[#513B2F]/10 bg-[#F1EBDD] px-3.5 py-2 text-right"><div className="text-[10px] uppercase tracking-[.12em] text-[#978574]">Dataset window</div><div className="mt-0.5 text-[12px] font-semibold">{dateRange(c.allRows,c.mapping)}</div></div>}/>
    <div className="mb-4 grid grid-cols-1 gap-3 sm:grid-cols-2 xl:grid-cols-4">
      {hasValue(c.mapping,'revenue')&&<Kpi title="Gross revenue" value={currency(m.revenue)} caption="Sum of mapped revenue · all statuses" icon={BarChart3}/>}
      {hasValue(c.mapping,'order')&&<Kpi title="Distinct orders" value={numberFormat(m.orders)} caption="Unique mapped order IDs · all statuses" icon={FileSpreadsheet}/>}
      {hasValue(c.mapping,'profit')&&<Kpi title="Recorded profit" value={currency(m.profit)} caption="Sum of mapped profit · as provided" icon={Activity} tone="sage"/>}
      {hasValue(c.mapping,'revenue')&&hasValue(c.mapping,'order')&&<Kpi title="Revenue / order" value={currency(m.aov)} caption="Gross revenue ÷ distinct orders" icon={SlidersHorizontal}/>}
      {!hasValue(c.mapping,'revenue')&&!hasValue(c.mapping,'order')&&!hasValue(c.mapping,'profit')&&<div className="panel col-span-full p-6 text-sm text-[#806D5C]">Map a revenue, order ID, or profit field in Settings to see the available KPIs.</div>}
    </div>
    <div className="mb-5"><ScopeNote mapping={c.mapping}/></div>
    {c.mapping.status&&<div className="mb-5 flex flex-wrap items-center gap-2.5 rounded-xl border border-[#513B2F]/10 bg-[#FBF8F1] px-4 py-3"><span className="mr-1 text-[10px] font-bold uppercase tracking-[.12em] text-[#927F6D]">Status mix</span>{statusCounts(c.allRows,c.mapping).map(entry=><span key={entry.status} className="flex items-center gap-2 rounded-full bg-[#F2EBDD] px-2.5 py-1"><StatusPill value={entry.status}/><span className="text-[10px] text-[#806D5C]">{entry.count}</span></span>)}<span className="ml-auto text-[10px] text-[#978574]">Counts across all {c.allRows.length} rows</span></div>}
    <div className="mb-5 grid grid-cols-1 gap-4 xl:grid-cols-[1.65fr_1fr]">
      <ChartPanel title="Revenue over time" subtitle={`${c.mapping.date? 'Monthly view of mapped revenue':'Date mapping required'} · ${trend} dated rows`}>
        <TrendChart rows={c.filteredRows} mapping={c.mapping}/>
      </ChartPanel>
      <ChartPanel title="Category share" subtitle={c.mapping.revenue?'Revenue contribution by category':'Record count by category'}>
        <CategoryDonut rows={c.filteredRows} mapping={c.mapping}/>
      </ChartPanel>
    </div>
    <div className="grid grid-cols-1 gap-4 xl:grid-cols-[.85fr_1.5fr]">
      <section className="panel p-5 sm:p-6"><div className="mb-5 flex items-start justify-between"><div><h2 className="font-display text-[19px]">Products to watch</h2><p className="mt-1 text-[11px] text-[#978574]">Ranked by mapped revenue</p></div><Link href="/products" className="text-[11px] font-semibold text-[#72805F] hover:underline">All products →</Link></div>
        {!c.mapping.product?<NoChart reason="Map a product column to rank products."/>:topProducts.length===0?<NoChart reason="No product values in this filter scope."/>:<div className="space-y-4">{topProducts.map((item,index)=><div key={item.name} className="flex items-center gap-3"><span className="grid h-7 w-7 shrink-0 place-items-center rounded-lg bg-[#F1EBDD] font-display text-[13px] text-[#806D5C]">{String(index+1).padStart(2,'0')}</span><div className="min-w-0 flex-1"><div className="truncate text-[12px] font-semibold">{item.name}</div><div className="mt-1 text-[10px] text-[#978574]">{item.orders} rows{item.quantity!==null?` · ${numberFormat(item.quantity)} units`:''}</div></div><div className="text-[12px] font-semibold">{c.mapping.revenue?currency(item.revenue):'—'}</div></div>)}</div>}
      </section>
      <section className="panel p-5 sm:p-6"><div className="mb-5 flex flex-wrap items-end justify-between gap-3"><div><h2 className="font-display text-[19px]">Recent records</h2><p className="mt-1 text-[11px] text-[#978574]">Most recent dated rows in this filtered view</p></div><button onClick={()=>c.exportRows()} className="flex items-center gap-1.5 text-[11px] font-semibold text-[#72805F] hover:underline"><Download size={13}/> Export view</button></div>
        {columns.length? <DataTable rows={sortRecent(c.filteredRows,c.mapping)} columns={columns} mapping={c.mapping} searchable={false}/>:<NoChart reason="Map columns in Settings to preview records."/>}
      </section>
    </div>
  </div>;
}
function dateRange(rows:Row[],mapping:Mapping) {if(!mapping.date)return 'No date mapped';const dates=rows.map(row=>parseDate(row[mapping.date])).filter((d):d is Date=>Boolean(d)).sort((a,b)=>a.getTime()-b.getTime());return dates.length?`${new Intl.DateTimeFormat('en',{month:'short',year:'numeric'}).format(dates[0])} — ${new Intl.DateTimeFormat('en',{month:'short',year:'numeric'}).format(dates[dates.length-1])}`:'No valid dates';}
function statusCounts(rows:Row[],mapping:Mapping) {const counts=new Map<string,number>();rows.forEach(row=>{const value=row[mapping.status]||'Blank';counts.set(value,(counts.get(value)??0)+1)});return [...counts.entries()].map(([status,count])=>({status,count})).sort((a,b)=>b.count-a.count);}
function aggregateProducts(rows:Row[],mapping:Mapping) {
  if(!mapping.product)return [];
  const map=new Map<string,{name:string;revenue:number;profit:number;quantity:number|null;orders:number}>();
  rows.forEach(row=>{const name=row[mapping.product]?.trim();if(!name)return;const item=map.get(name)??{name,revenue:0,profit:0,quantity:mapping.quantity?0:null,orders:0};item.revenue+=mapping.revenue?toNumber(row[mapping.revenue])??0:0;item.profit+=mapping.profit?toNumber(row[mapping.profit])??0:0;if(mapping.quantity)item.quantity!+=toNumber(row[mapping.quantity])??0;item.orders++;map.set(name,item)});
  return [...map.values()].sort((a,b)=>b.revenue-a.revenue);
}
function sortRecent(rows:Row[],mapping:Mapping) {if(!mapping.date)return rows.slice().reverse();return rows.slice().sort((a,b)=>(parseDate(b[mapping.date])?.getTime()??0)-(parseDate(a[mapping.date])?.getTime()??0)).slice(0,8);}

function AnalysisPage({context:c}:{context:DashboardContext}) {
  const [granularity,setGranularity]=useState<'month'|'day'>('month');
  const rows=c.filteredRows; const m=metrics(rows,c.mapping);
  const comp=useMemo(()=>[
    {name:'Revenue',value:m.revenue,color:'#513B2F'},
    ...(c.mapping.profit?[{name:'Profit',value:m.profit,color:'#879475'}]:[]),
  ],[m.revenue,m.profit,c.mapping.profit]);
  const comparable=!!(c.mapping.revenue||c.mapping.profit);
  return <div className="page-enter">
    <PageHeading eyebrow="Explore the record" title="Sales analysis" description="Follow the shape of revenue, compare recorded profit, and see how category performance shifts with your filters."/>
    <div className="mb-5 grid grid-cols-1 gap-4 xl:grid-cols-[1.55fr_1fr]">
      <ChartPanel title="Sales trend" subtitle={`${granularity==='month'?'Monthly':'Daily'} totals from the currently filtered records`} >
        <div className="mb-3 flex justify-end"><div className="rounded-lg bg-[#F0E9DC] p-1">{(['month','day'] as const).map(value=><button key={value} onClick={()=>setGranularity(value)} className={`rounded-md px-3 py-1.5 text-[10px] font-semibold capitalize ${granularity===value?'bg-[#FBF8F1] text-[#513B2F] shadow-sm':'text-[#907E6C]'}`}>{value}</button>)}</div></div>
        <TrendChart rows={rows} mapping={c.mapping} group={granularity}/>
      </ChartPanel>
      <ChartPanel title="Metric comparison" subtitle="Totals shown in INR from mapped columns">
        {!comparable?<NoChart reason="Map revenue and/or profit to compare available totals."/>:<div className="space-y-5 pt-3">{comp.map((entry,index)=><div key={entry.name}><div className="mb-2 flex justify-between text-[12px]"><span className="text-[#806D5C]">{entry.name}</span><span className="font-semibold">{currency(entry.value)}</span></div><div className="h-3 overflow-hidden rounded-full bg-[#EEE7DA]"><div className="h-full rounded-full transition-all duration-500" style={{width:`${Math.max(1,entry.value/Math.max(...comp.map(item=>item.value),1)*100)}%`,background:entry.color}}/></div></div>)}
        <div className="mt-7 border-t border-[#513B2F]/10 pt-4"><div className="text-[10px] uppercase tracking-[.12em] text-[#978574]">Filtered scope</div><div className="mt-2 flex items-baseline gap-2"><span className="font-display text-2xl">{numberFormat(rows.length)}</span><span className="text-[11px] text-[#927F6D]">records · {numberFormat(m.orders)} distinct orders where mapped</span></div></div></div>}
      </ChartPanel>
    </div>
    <div className="mb-5 grid grid-cols-1 gap-4 lg:grid-cols-[1fr_1fr]">
      <ChartPanel title="Category comparison" subtitle="Revenue by category in the current filter view"><CategoryBar rows={rows} mapping={c.mapping}/></ChartPanel>
      <ChartPanel title="Category composition" subtitle="Relative contribution to the selected measure"><CategoryDonut rows={rows} mapping={c.mapping}/></ChartPanel>
    </div>
    <div className="panel p-5 sm:p-6"><div className="mb-4"><h2 className="font-display text-[19px]">Status ledger</h2><p className="mt-1 text-[11px] text-[#978574]">All statuses remain included unless you choose a status filter above.</p></div><StatusLedger rows={rows} mapping={c.mapping}/></div>
  </div>;
}
function CategoryBar({rows,mapping}:{rows:Row[];mapping:Mapping}) {
  const data=useMemo(()=>{if(!mapping.category)return [];const values=new Map<string,number>();rows.forEach(row=>{const name=row[mapping.category];if(name)values.set(name,(values.get(name)??0)+(mapping.revenue?toNumber(row[mapping.revenue])??0:1))});return [...values.entries()].map(([name,value])=>({name,value})).sort((a,b)=>b.value-a.value)},[rows,mapping]);
  if(!mapping.category)return <NoChart reason="Map a category field to compare category totals."/>;
  return <ResponsiveContainer width="100%" height={240}><BarChart data={data} margin={{top:8,right:8,left:0,bottom:4}}><CartesianGrid vertical={false} stroke="#E8E0D2" strokeDasharray="3 5"/><XAxis dataKey="name" tick={{fontSize:10,fill:'#806D5C'}} axisLine={false} tickLine={false}/><YAxis tickFormatter={v=>mapping.revenue?`₹${Number(v)>=1000?`${(Number(v)/1000).toFixed(0)}k`:v}`:String(v)} tick={{fontSize:10,fill:'#947F6B'}} axisLine={false} tickLine={false} width={45}/><Tooltip formatter={(value:number)=>mapping.revenue?currency(value):numberFormat(value)} contentStyle={{borderRadius:12,border:'1px solid #e6ddcd',background:'#fffdf8',fontSize:12}}/><Bar dataKey="value" name={mapping.revenue?'Revenue':'Records'} radius={[6,6,0,0]} isAnimationActive={false}>{data.map((entry,index)=><Cell key={entry.name} fill={chartColors[index%chartColors.length]}/>)}</Bar></BarChart></ResponsiveContainer>;
}
function StatusLedger({rows,mapping}:{rows:Row[];mapping:Mapping}) {
  if(!mapping.status)return <NoChart reason="Map an order-status field to inspect the status mix."/>;
  const counts=statusCounts(rows,mapping);
  return <div className="grid grid-cols-2 gap-3 md:grid-cols-4">{counts.map(entry=><div key={entry.status} className="rounded-xl bg-[#F4EFE5] p-4"><StatusPill value={entry.status}/><div className="mt-3 font-display text-2xl">{entry.count}</div><div className="mt-1 text-[10px] text-[#978574]">{rows.length?((entry.count/rows.length)*100).toFixed(1):0}% of filtered rows</div></div>)}</div>;
}

function ProductsPage({context:c}:{context:DashboardContext}) {
  const productData=useMemo(()=>aggregateProducts(c.filteredRows,c.mapping),[c.filteredRows,c.mapping]);
  const columns=[...(c.mapping.product?[c.mapping.product]:[]),...(c.mapping.revenue?[`Total ${c.mapping.revenue}`]:[]),...(c.mapping.profit?[`Total ${c.mapping.profit}`]:[]),...(c.mapping.quantity?[`Total ${c.mapping.quantity}`]:[]),'Rows'];
  const data=useMemo(()=>productData.map(item=>({[c.mapping.product||'Product']:item.name,...(c.mapping.revenue?{[`Total ${c.mapping.revenue}`]:String(item.revenue)}:{}),...(c.mapping.profit?{[`Total ${c.mapping.profit}`]:String(item.profit)}:{}),...(c.mapping.quantity?{[`Total ${c.mapping.quantity}`]:String(item.quantity)}:{}),Rows:String(item.orders)})),[productData,c.mapping]);
  const mappingForTable={...c.mapping,revenue:c.mapping.revenue?`Total ${c.mapping.revenue}`:'',profit:c.mapping.profit?`Total ${c.mapping.profit}`:'',product:c.mapping.product||'Product'};
  return <div className="page-enter">
    <PageHeading eyebrow="Product performance" title="Products" description="Find what carries the most value. Search and sort within the current filtered dataset; revenue and profit are mapped sums."/>
    <div className="mb-5 grid grid-cols-1 gap-3 sm:grid-cols-3"><Kpi title="Products in view" value={numberFormat(productData.length)} caption="Distinct mapped product names" icon={Package}/><Kpi title="Product revenue" value={c.mapping.revenue?currency(metrics(c.filteredRows,c.mapping).revenue):'Not mapped'} caption="Mapped revenue across current rows" icon={BarChart3}/><Kpi title="Units recorded" value={c.mapping.quantity?numberFormat(metrics(c.filteredRows,c.mapping).quantity):'Not mapped'} caption="Sum of mapped quantity" icon={FileSpreadsheet} tone="sage"/></div>
    <div className="panel p-5 sm:p-6"><div className="mb-5 flex flex-wrap items-center justify-between gap-2"><div><h2 className="font-display text-[20px]">Ranked product ledger</h2><p className="mt-1 text-[11px] text-[#978574]">Sorted by highest revenue · {productData.length} unique products</p></div><button onClick={()=>c.exportRows()} className="flex items-center gap-2 rounded-lg border border-[#513B2F]/15 px-3 py-2 text-[11px] hover:bg-[#F2EBDD]"><Download size={14}/> Export filtered rows</button></div>
      {!c.mapping.product?<NoChart reason="Map a product column in Settings to create the product ranking."/>:<DataTable rows={data} columns={columns} mapping={mappingForTable} initialSort={c.mapping.revenue?`Total ${c.mapping.revenue}`:'Rows'}/>}
    </div>
  </div>;
}

function CategoriesPage({context:c}:{context:DashboardContext}) {
  const [selected,setSelected]=useState('');
  const categoryRows=useMemo(()=>{if(!c.mapping.category)return [];return c.filteredRows.filter(row=>!selected||row[c.mapping.category]===selected)},[c.filteredRows,c.mapping.category,selected]);
  const stats=useMemo(()=>{if(!c.mapping.category)return [];const groups=new Map<string,Row[]>();c.filteredRows.forEach(row=>{const name=row[c.mapping.category];if(name)groups.set(name,[...(groups.get(name)??[]),row])});return [...groups.entries()].map(([name,rows])=>({name,rows,revenue:metrics(rows,c.mapping).revenue,profit:metrics(rows,c.mapping).profit,orders:metrics(rows,c.mapping).orders,quantity:metrics(rows,c.mapping).quantity})).sort((a,b)=>b.revenue-a.revenue)},[c.filteredRows,c.mapping]);
  return <div className="page-enter">
    <PageHeading eyebrow="Mix & contribution" title="Categories" description="A category-by-category view of the values in your dataset. Choose a category to inspect its records and compare contribution."/>
    {!c.mapping.category?<div className="panel p-8"><NoChart reason="Map a category field in Settings to view category analysis."/></div>:<>
      <div className="mb-5 grid grid-cols-1 gap-4 lg:grid-cols-[1fr_1fr]">
        <ChartPanel title="Revenue by category" subtitle={c.mapping.revenue?'Ranked by mapped revenue total':'Record count by category'}><CategoryBar rows={c.filteredRows} mapping={c.mapping}/></ChartPanel>
        <ChartPanel title="Share of the whole" subtitle={c.mapping.revenue?'Contribution to total revenue':'Share of records'}><CategoryDonut rows={c.filteredRows} mapping={c.mapping}/></ChartPanel>
      </div>
      <div className="mb-5 grid grid-cols-1 gap-3 sm:grid-cols-2 xl:grid-cols-3">{stats.map((item,index)=><button key={item.name} onClick={()=>setSelected(selected===item.name?'':item.name)} className={`panel group p-5 text-left transition hover:-translate-y-0.5 hover:shadow-md ${selected===item.name?'ring-2 ring-[#879475]/50':''}`}><div className="flex items-start justify-between"><span className="grid h-9 w-9 place-items-center rounded-xl text-white" style={{background:chartColors[index%chartColors.length]}}><Tag size={16}/></span><span className="text-[10px] text-[#978574]">{item.rows.length} rows</span></div><div className="mt-4 font-display text-[19px]">{item.name}</div><div className="mt-3 grid grid-cols-2 gap-3"><div><div className="text-[10px] text-[#978574]">Revenue</div><div className="mt-1 text-[13px] font-semibold">{c.mapping.revenue?currency(item.revenue):'—'}</div></div><div><div className="text-[10px] text-[#978574]">Profit</div><div className="mt-1 text-[13px] font-semibold">{c.mapping.profit?currency(item.profit):'—'}</div></div></div><div className="mt-3 border-t border-[#513B2F]/10 pt-3 text-[10px] text-[#978574]">{item.orders} rows/orders as represented{c.mapping.quantity?` · ${numberFormat(item.quantity)} units`:''}</div></button>)}</div>
      <div className="panel p-5 sm:p-6"><div className="mb-4 flex items-center justify-between"><div><h2 className="font-display text-[19px]">{selected||'Category records'}</h2><p className="mt-1 text-[11px] text-[#978574]">{selected?'Selected category · current global filters':'Select a category card above or explore all current rows'}</p></div>{selected&&<button onClick={()=>setSelected('')} className="text-[11px] text-[#72805F] hover:underline">Clear selection</button>}</div><DataTable rows={categoryRows} columns={[c.mapping.category,c.mapping.product,c.mapping.revenue,c.mapping.profit,c.mapping.status].filter(Boolean)} mapping={c.mapping}/></div>
    </>}
  </div>;
}

function QualityPage({context:c}:{context:DashboardContext}) {
  const data=c.dataset!;
  const quality=useMemo(()=>qualitySummary(data),[data]);
  const invalid=useMemo(()=>c.mapping ? Object.entries(c.mapping).flatMap(([role,column])=>{
    if(!column)return [];
    if(role==='revenue'||role==='profit'||role==='quantity')return c.allRows.flatMap((row,index)=>row[column]?.trim()&&toNumber(row[column])===null?[{role,column,index}]:[]);
    if(role==='date')return c.allRows.flatMap((row,index)=>row[column]?.trim()&&!parseDate(row[column])?[{role,column,index}]:[]);
    return [];
  }):[],[data,c.mapping]);
  const missingTotal=Object.values(quality.missing).reduce((sum,value)=>sum+value,0);
  const duplicateRows=quality.duplicates;
  return <div className="page-enter">
    <PageHeading eyebrow="Trust the source" title="Data quality" description="An honest inventory of the uploaded file. Missing cells, repeated rows, and values that cannot be interpreted are reported—not silently repaired."/>
    <div className="mb-5 grid grid-cols-2 gap-3 xl:grid-cols-4"><Kpi title="Rows loaded" value={numberFormat(data.rows.length)} caption={`Source file · ${data.fileName}`} icon={FileSpreadsheet}/><Kpi title="Columns" value={numberFormat(data.columns.length)} caption="Headers found in source" icon={SlidersHorizontal}/><Kpi title="Missing cells" value={numberFormat(missingTotal)} caption={missingTotal?'Blank values preserved':'No blank cells detected'} icon={AlertCircle} tone={missingTotal?'brown':'sage'}/><Kpi title="Duplicate rows" value={numberFormat(duplicateRows)} caption={duplicateRows?'Exact full-row repeats':'No exact duplicates detected'} icon={Check} tone="sage"/></div>
    <div className="mb-5 grid grid-cols-1 gap-4 xl:grid-cols-[1.2fr_.8fr]">
        <section className="panel p-5 sm:p-6"><div className="mb-4"><h2 className="font-display text-[20px]">Column inventory</h2><p className="mt-1 text-[11px] text-[#978574]">Detected from the current dataset · all {data.rows.length} rows</p></div><div className="overflow-x-auto rounded-xl border border-[#513B2F]/10"><table className="w-full min-w-[580px] text-left"><thead className="bg-[#F2EBDD]"><tr>{['Column','Inferred type','Missing','Mapped role'].map(v=><th key={v} className="px-4 py-3 text-[10px] font-bold uppercase tracking-[.12em] text-[#806D5C]">{v}</th>)}</tr></thead><tbody>{data.columns.map(column=><tr key={column} className="border-t border-[#513B2F]/[.07]"><td className="px-4 py-3 text-[12px] font-semibold">{column}</td><td className="px-4 py-3 text-[12px] text-[#806D5C]">{quality.types[column]}</td><td className="px-4 py-3 text-[12px]">{quality.missing[column] ? <span className="text-[#A55E4D]">{quality.missing[column]}</span>:'0'}</td><td className="px-4 py-3 text-[11px] text-[#806D5C]">{ROLE_META.find(item=>c.mapping[item.role]===column)?.label??<span className="text-[#B1A18E]">Not mapped</span>}</td></tr>)}</tbody></table></div></section>
      <div className="space-y-4">
        <section className="panel p-5"><div className="mb-3 flex items-center gap-2"><div className={`grid h-8 w-8 place-items-center rounded-lg ${missingTotal?'bg-[#F2E4DF] text-[#965E50]':'bg-[#E7EBDD] text-[#5A694D]'}`}>{missingTotal?<AlertCircle size={15}/>:<Check size={15}/>}</div><h2 className="font-display text-[18px]">Completeness</h2></div><p className="text-[12px] leading-5 text-[#806D5C]">{missingTotal?`${missingTotal} blank cells across ${Object.entries(quality.missing).filter(([,value])=>value).length} columns. Values are left as-is.`:'No missing values were found across the loaded rows and columns.'}</p></section>
        <section className="panel p-5"><div className="mb-3 flex items-center gap-2"><div className={`grid h-8 w-8 place-items-center rounded-lg ${duplicateRows?'bg-[#F3EAD7] text-[#927243]':'bg-[#E7EBDD] text-[#5A694D]'}`}>{duplicateRows?<AlertCircle size={15}/>:<Check size={15}/>}</div><h2 className="font-display text-[18px]">Duplicate rows</h2></div><p className="text-[12px] leading-5 text-[#806D5C]">{duplicateRows?`${duplicateRows} exact duplicate row(s) detected. None have been removed.`:'No exact full-row duplicates detected.'}</p></section>
        <section className="panel p-5"><div className="mb-3 flex items-center gap-2"><div className={`grid h-8 w-8 place-items-center rounded-lg ${invalid.length?'bg-[#F3EAD7] text-[#927243]':'bg-[#E7EBDD] text-[#5A694D]'}`}>{invalid.length?<AlertCircle size={15}/>:<Check size={15}/>}</div><h2 className="font-display text-[18px]">Mapped values</h2></div><p className="text-[12px] leading-5 text-[#806D5C]">{invalid.length?`${invalid.length} non-empty value(s) in mapped numeric/date fields cannot be interpreted. They remain visible in the source preview; calculations skip unparseable values.`:'No invalid values found in mapped numeric and date fields.'}</p></section>
      </div>
    </div>
    <section className="panel p-5 sm:p-6"><div className="mb-4 flex flex-wrap items-end justify-between gap-2"><div><h2 className="font-display text-[20px]">Source preview</h2><p className="mt-1 text-[11px] text-[#978574]">First 12 source rows · values shown without correction</p></div><span className="rounded-full bg-[#F1EBDD] px-3 py-1.5 text-[10px] text-[#806D5C]">{data.rows.length} rows × {data.columns.length} columns</span></div><div className="overflow-x-auto rounded-xl border border-[#513B2F]/10"><table className="w-full min-w-[900px] text-left"><thead className="bg-[#F2EBDD]"><tr>{data.columns.map(column=><th key={column} className="px-3 py-3 text-[10px] font-bold text-[#806D5C]">{column}</th>)}</tr></thead><tbody>{data.rows.slice(0,12).map((row,index)=><tr key={index} className="border-t border-[#513B2F]/[.07]">{data.columns.map(column=><td key={column} className="max-w-48 truncate px-3 py-2.5 text-[11px] text-[#584438]">{row[column]||<span className="text-[#A55E4D]">Blank</span>}</td>)}</tr>)}</tbody></table></div></section>
  </div>;
}

function SettingsPage({context:c}:{context:DashboardContext}) {
  const data=c.dataset!;
  const update=(role:Role,value:string)=>c.setMapping({...c.mapping,[role]:value});
  return <div className="page-enter">
    <PageHeading eyebrow="Shape your view" title="Settings" description="Choose which source columns play each analytical role. Mapping changes immediately update the metrics, charts, filters, and tables across the workspace."/>
    <div className="grid grid-cols-1 gap-5 xl:grid-cols-[1.2fr_.8fr]">
      <section className="panel p-5 sm:p-7"><div className="mb-6 flex items-start justify-between gap-3"><div><h2 className="font-display text-[20px]">Column mapping</h2><p className="mt-1 text-[12px] text-[#978574]">Auto-detected suggestions · editable any time</p></div><span className="flex items-center gap-1.5 rounded-full bg-[#E7EBDD] px-3 py-1.5 text-[10px] font-semibold text-[#5A694D]"><Check size={12}/> Live mapping</span></div>
        <div className="space-y-3">{ROLE_META.map(({role,label})=><div key={role} className="grid grid-cols-1 items-center gap-2 rounded-xl bg-[#F6F1E7] p-3 sm:grid-cols-[.75fr_1.25fr]"><div><div className="text-[12px] font-semibold">{label}</div><div className="mt-0.5 text-[10px] text-[#978574]">{role==='revenue'||role==='profit'?'Used for INR totals':role==='order'?'Distinct IDs for order count':role==='date'?'Trend and date filters':'Available for breakdowns'}</div></div><select value={c.mapping[role]??''} onChange={event=>update(role,event.target.value)} className="w-full rounded-lg border border-[#513B2F]/15 bg-[#FFFCF5] px-3 py-2.5 text-[12px] outline-none focus:border-[#879475]" aria-label={`${label} column mapping`} data-testid={`select-map-${role}`}><option value="">Not mapped</option>{data.columns.map(column=><option key={column} value={column}>{column} · {qualitySummary(data).types[column]}</option>)}</select></div>)}</div>
        <div className="mt-5 flex items-start gap-2 rounded-xl border border-[#D6B58C]/35 bg-[#F3E9D8]/65 px-4 py-3 text-[11px] leading-5 text-[#765F4B]"><CircleHelp size={15} className="mt-0.5 shrink-0"/>No data is altered by mapping. An unmapped role hides dependent calculations rather than filling in a guess.</div>
      </section>
      <div className="space-y-5">
        <section className="panel p-5 sm:p-6"><div className="mb-4 grid h-10 w-10 place-items-center rounded-xl bg-[#EAE2D3] text-[#513B2F]"><FileSpreadsheet size={19}/></div><h2 className="font-display text-[20px]">Current dataset</h2><div className="mt-4 space-y-3 text-[12px]"><InfoLine label="File" value={data.fileName}/><InfoLine label="Rows" value={numberFormat(data.rows.length)}/><InfoLine label="Columns" value={numberFormat(data.columns.length)}/><InfoLine label="Processing" value="This browser only"/></div><button onClick={c.chooseFile} className="mt-5 flex w-full items-center justify-center gap-2 rounded-lg bg-[#513B2F] px-4 py-3 text-[12px] font-semibold text-[#F7F2E8] transition hover:bg-[#684D3D]" data-testid="button-upload-new"><Upload size={15}/>Replace with another file</button><p className="mt-3 text-[10px] leading-5 text-[#978574]">Supported formats: CSV, XLS, XLSX. The replacement stays in browser memory for this session; nothing is uploaded to a server.</p></section>
        <section className="panel p-5 sm:p-6"><h2 className="font-display text-[19px]">Reading preferences</h2><p className="mt-1 text-[11px] text-[#978574]">A few notes on how this workspace interprets your source.</p><div className="mt-4 space-y-3"><Preference icon={Check} title="Currency" detail="INR · Indian numbering conventions"/><Preference icon={Check} title="Order count" detail="Distinct IDs when an Order ID role is mapped"/><Preference icon={Check} title="Status scope" detail="All statuses included unless filtered"/><Preference icon={Check} title="Client-side" detail="No backend, integrations, or polling"/></div></section>
      </div>
    </div>
  </div>;
}
function InfoLine({label,value}:{label:string;value:string}) {return <div className="flex justify-between gap-3 border-b border-[#513B2F]/[.08] pb-2"><span className="text-[#927F6D]">{label}</span><span className="max-w-[65%] truncate text-right font-medium">{value}</span></div>;}
function Preference({icon:Icon,title,detail}:{icon:typeof Check;title:string;detail:string}) {return <div className="flex items-start gap-3 rounded-lg bg-[#F6F1E7] p-3"><span className="mt-0.5 text-[#71805F]"><Icon size={14}/></span><div><div className="text-[11px] font-semibold">{title}</div><div className="mt-0.5 text-[10px] text-[#978574]">{detail}</div></div></div>;}

export default App;
