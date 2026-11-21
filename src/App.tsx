import { useEffect, useMemo, useRef, useState } from 'react';
import {
  ArrowDownToLine, ArrowRight, Building2, Check, ChevronDown, ChevronRight,
  CircleAlert, ClipboardList, FileSpreadsheet, Gauge, Landmark, Plus,
  RotateCcw, SlidersHorizontal, Trash2, Upload, Wallet,
} from 'lucide-react';
import {
  calculateAllocations, centsToEuros, createDefaultCinemas, DEFAULT_PRIORITY_FACTORS,
  type Cinema, type HistoricalCategory, type PriorityFactor,
} from './lib/budget-model';
import {
  detectImportMapping, readWorkbook, summarizeImport,
  type ImportMapping, type ImportSummary, type WorkbookData, type WorkbookSheet,
} from './lib/spreadsheet-import';

type SavedPlan = {
  cinemas: Cinema[];
  factors: PriorityFactor[];
  categories: HistoricalCategory[];
  totalBudget: number;
  categoryGroups: Record<string, string>;
  importedAt: string | null;
};

const STORAGE_KEY = 'the-space-budget-plan-v1';
const FACTOR_COLORS = ['#347663', '#d47b4c', '#578da0', '#c4a044', '#a56b7e', '#718269'];
const FACTOR_LABELS: Record<string, string> = {
  attendance: 'Impatto sul pubblico',
  areaSqm: 'Superficie dell’edificio',
  age: 'Età dell’edificio',
  renovation: 'Storico ristrutturazioni',
  locationType: 'Esposizione della sede',
  roofStatus: 'Stato della copertura',
  maintenanceResponsibility: 'Responsabilità della manutenzione',
};
const formatEUR = (value: number) => new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', maximumFractionDigits: 2 }).format(value || 0);
const formatNumber = (value: number, maximumFractionDigits = 0) => new Intl.NumberFormat('it-IT', { maximumFractionDigits }).format(value || 0);
const formatPercent = (value: number, maximumFractionDigits = 1) => new Intl.NumberFormat('it-IT', { style: 'percent', maximumFractionDigits }).format(value || 0);
const normalizedName = (value: string) => value.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
const mostRecentCompletedYear = (years: number[]) => {
  const available = [...new Set(years)].sort((a, b) => b - a);
  return available.find((year) => year < new Date().getFullYear()) ?? available[0] ?? new Date().getFullYear();
};

const LEGACY_FACTOR_LABELS: Record<string, string> = {
  'Audience impact': 'Impatto sul pubblico',
  'Building footprint': 'Superficie dell’edificio',
  'Building lifecycle': 'Età dell’edificio',
  'Renovation history': 'Storico ristrutturazioni',
  'Location exposure': 'Esposizione della sede',
  'Roof condition': 'Stato della copertura',
  'Maintenance responsibility': 'Responsabilità della manutenzione',
};

function initialPlan(): SavedPlan {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw) as Partial<SavedPlan>;
      const factors = Array.isArray(parsed.factors)
        ? parsed.factors.map((factor) => ({ ...factor, label: LEGACY_FACTOR_LABELS[factor.label] ?? factor.label }))
        : DEFAULT_PRIORITY_FACTORS.map((factor) => ({ ...factor, label: FACTOR_LABELS[factor.id] ?? factor.label }));
      return {
        cinemas: Array.isArray(parsed.cinemas) && parsed.cinemas.length
          ? parsed.cinemas
          : createDefaultCinemas(),
        factors,
        categories: Array.isArray(parsed.categories) ? parsed.categories : [],
        totalBudget: Number(parsed.totalBudget) || 0,
        categoryGroups: parsed.categoryGroups && typeof parsed.categoryGroups === 'object' ? parsed.categoryGroups : {},
        importedAt: parsed.importedAt ?? null,
      };
    }
  } catch { /* continue with a clean plan when stored data is unreadable */ }
  return {
    cinemas: createDefaultCinemas(),
    factors: DEFAULT_PRIORITY_FACTORS.map((factor) => ({ ...factor, label: FACTOR_LABELS[factor.id] ?? factor.label })),
    categories: [],
    totalBudget: 0,
    categoryGroups: {},
    importedAt: null,
  };
}

function hasRequiredMapping(mapping: ImportMapping) {
  return Boolean(mapping.cinemaField && mapping.categoryField && mapping.amountField && mapping.yearField);
}

function defaultSummaryForSheet(sheet: WorkbookSheet, mapping: ImportMapping, categoryGroups: Record<string, string>) {
  if (!hasRequiredMapping(mapping)) return null;
  const available = summarizeImport(sheet, mapping, new Date().getFullYear(), categoryGroups);
  const year = mostRecentCompletedYear(available.years);
  return year === available.selectedYear ? available : summarizeImport(sheet, mapping, year, categoryGroups);
}

function App() {
  const [plan, setPlan] = useState<SavedPlan>(initialPlan);
  const [workbook, setWorkbook] = useState<WorkbookData | null>(null);
  const [sheetName, setSheetName] = useState('');
  const [mapping, setMapping] = useState<ImportMapping>({ cinemaField: '', categoryField: '', amountField: '', yearField: '' });
  const [selectedYear, setSelectedYear] = useState(0);
  const [fileName, setFileName] = useState('');
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const [busy, setBusy] = useState(false);
  const [selectedCinema, setSelectedCinema] = useState<string | null>(null);
  const [categoryOpen, setCategoryOpen] = useState(true);
  const [newFactor, setNewFactor] = useState('');
  const [search, setSearch] = useState('');
  const fileRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(plan)); }
    catch { setError('Impossibile salvare i dati nel browser. Esporta il piano per conservarne una copia.'); }
  }, [plan]);

  const activeSheet: WorkbookSheet | undefined = workbook?.sheets.find((sheet) => sheet.name === sheetName);
  const canSummarize = Boolean(activeSheet && hasRequiredMapping(mapping));
  const summary: ImportSummary | null = useMemo(() => {
    if (!activeSheet || !hasRequiredMapping(mapping)) return null;
    return summarizeImport(activeSheet, mapping, selectedYear || new Date().getFullYear(), plan.categoryGroups);
  }, [activeSheet, mapping, selectedYear, plan.categoryGroups]);
  useEffect(() => {
    if (summary?.years.length && !summary.years.includes(selectedYear)) {
      setSelectedYear(mostRecentCompletedYear(summary.years));
    }
  }, [summary, selectedYear]);
  const sourceCategories = useMemo(() => {
    if (!activeSheet || !mapping.categoryField) return [];
    return [...new Set(activeSheet.rows
      .filter((row) => Number(String(row[mapping.yearField] ?? '').match(/(?:19|20)\d{2}/)?.[0]) === selectedYear)
      .map((row) => String(row[mapping.categoryField] ?? '').trim())
      .filter(Boolean))].sort((a, b) => a.localeCompare(b));
  }, [activeSheet, mapping.categoryField, mapping.yearField, selectedYear]);
  const years = summary?.years ?? [];
  const expenseSourceLabel = mapping.categoryField || 'Causale';
  const allocations = useMemo(
    () => calculateAllocations(plan.cinemas, plan.factors, plan.totalBudget, plan.categories),
    [plan.cinemas, plan.factors, plan.totalBudget, plan.categories],
  );
  const allocationById = useMemo(() => new Map(allocations.map((allocation) => [allocation.cinemaId, allocation])), [allocations]);
  const ranked = useMemo(() => plan.cinemas.map((cinema) => ({
    ...cinema,
    result: allocationById.get(cinema.id),
  })).sort((a, b) => (b.result?.amountCents ?? 0) - (a.result?.amountCents ?? 0)), [plan.cinemas, allocationById]);
  const visibleCinemas = ranked.filter((cinema) => cinema.name.toLowerCase().includes(search.toLowerCase()));
  const selected = ranked.find((cinema) => cinema.id === selectedCinema);
  const totalWeight = plan.factors.filter((factor) => factor.enabled).reduce((sum, factor) => sum + Math.max(0, factor.weight), 0);

  const updatePlan = (changes: Partial<SavedPlan>) => setPlan((current) => ({ ...current, ...changes }));
  const updateCinema = (id: string, changes: Partial<Cinema>) => updatePlan({ cinemas: plan.cinemas.map((cinema) => cinema.id === id ? { ...cinema, ...changes } : cinema) });
  const handleWorkbook = async (file?: File) => {
    if (!file) return;
    setError(''); setNotice(''); setBusy(true); setFileName(file.name);
    try {
      const loaded = await readWorkbook(file);
      if (!loaded.sheets.length) throw new Error('Il file non contiene fogli di lavoro.');
      const sheet = loaded.sheets[0];
      const detected = detectImportMapping(sheet.headers);
      const nextMapping: ImportMapping = {
        cinemaField: detected.cinemaField ?? '',
        categoryField: detected.categoryField ?? '',
        amountField: detected.amountField ?? '',
        yearField: detected.yearField ?? '',
      };
      const firstSummary = defaultSummaryForSheet(sheet, nextMapping, plan.categoryGroups);
      setWorkbook(loaded); setSheetName(sheet.name); setMapping(nextMapping);
      setSelectedYear(firstSummary?.selectedYear ?? new Date().getFullYear());
       if (!hasRequiredMapping(nextMapping)) setNotice('Controlla l’associazione delle colonne qui sotto. La colonna Causale è proposta come fonte iniziale delle categorie di spesa.');
      else setNotice(`File pronto · ${formatNumber(sheet.rows.length)} righe nel foglio «${sheet.name}». Rivedi i raggruppamenti prima dell’importazione.`);
    } catch (cause) {
      setWorkbook(null);
      setError(cause instanceof Error && cause.message === 'Il file non contiene fogli di lavoro.'
        ? cause.message
        : 'Impossibile leggere il file. Scegli un file .xls, .xlsx o .csv e riprova.');
    } finally { setBusy(false); }
  };
  const changeSheet = (name: string) => {
    if (!workbook) return;
    const sheet = workbook.sheets.find((candidate) => candidate.name === name);
    if (!sheet) return;
    const detected = detectImportMapping(sheet.headers);
    const nextMapping: ImportMapping = {
      cinemaField: detected.cinemaField ?? '',
      categoryField: detected.categoryField ?? '',
      amountField: detected.amountField ?? '',
      yearField: detected.yearField ?? '',
    };
    const firstSummary = defaultSummaryForSheet(sheet, nextMapping, plan.categoryGroups);
    setSheetName(name); setMapping(nextMapping); setSelectedYear(firstSummary?.selectedYear ?? new Date().getFullYear());
  };
  const importWorkbook = () => {
    if (!activeSheet || !hasRequiredMapping(mapping)) {
      setError('Prima di importare, associa le colonne cinema, Causale/categoria, importo e anno.');
      return;
    }
    const result = summarizeImport(activeSheet, mapping, selectedYear, plan.categoryGroups);
    if (!result.cinemas.length) { setError('Non sono stati trovati nomi di cinema. Controlla il foglio e l’associazione delle colonne.'); return; }
    if (!result.years.length) { setError('Nella colonna associata all’anno non sono stati trovati valori validi.'); return; }
    if (result.transactionCount === 0) {
      setError(`Nessuna riga di spesa corrisponde al ${formatNumber(selectedYear)}. Controlla l’anno e le colonne associate prima di importare.`);
      return;
    }
    const existingByName = new Map(plan.cinemas.map((cinema) => [normalizedName(cinema.name), cinema]));
    const cinemas = result.cinemas.map((cinema) => {
      const existing = existingByName.get(normalizedName(cinema.name));
      return existing ? { ...existing, name: cinema.name } : cinema;
    });
    updatePlan({ cinemas, categories: result.categories, factors: plan.factors, categoryGroups: plan.categoryGroups, importedAt: new Date().toISOString() });
    setSelectedCinema(null); setError('');
    setNotice(`Importati ${formatNumber(cinemas.length)} cinema e ${formatNumber(result.categories.length)} categorie di spesa consuntiva per il ${formatNumber(result.selectedYear)}.${result.excludedCenters.length ? ` Escluse ${formatNumber(result.excludedCenters.length)} righe relative alla sede centrale.` : ''}`);
  };
  const updateFactor = (id: string, changes: Partial<PriorityFactor>) => updatePlan({ factors: plan.factors.map((factor) => factor.id === id ? { ...factor, ...changes } : factor) });
  const addFactor = () => {
    const label = newFactor.trim();
    if (!label) return;
    const id = `priority-${Date.now()}`;
    const factor: PriorityFactor = { id, label, weight: 1, enabled: true, kind: 'manual', direction: 'higher' };
    updatePlan({
      factors: [...plan.factors, factor],
      cinemas: plan.cinemas.map((cinema) => ({ ...cinema, factorScores: { ...cinema.factorScores, [id]: 3 } })),
    });
    setNewFactor('');
  };
  const removeFactor = (id: string) => updatePlan({ factors: plan.factors.filter((factor) => factor.id !== id) });
  const addOdeon = () => {
    if (plan.cinemas.some((cinema) => normalizedName(cinema.name).includes('odeonmilano'))) {
      setNotice('TSC Odeon Milano è già incluso nel piano.'); return;
    }
    const cinema: Cinema = {
      id: 'tsc-odeon-milano', name: 'TSC Odeon Milano', attendance: null, areaSqm: null,
      openingYear: null, renovated: null, locationType: null, roofStatus: null,
      maintenanceResponsibility: null,
      factorScores: {},
    };
    updatePlan({ cinemas: [...plan.cinemas, cinema] });
    setSelectedCinema(cinema.id);
  };
  const exportCsv = () => {
    if (!ranked.length) return;
    const headers = [
      'Cinema', 'Quota assegnata EUR', 'Coefficiente', 'Quota sul budget %', 'Presenze', 'Superficie mq', 'Anno di apertura',
      'Ristrutturato', 'Tipologia sede', 'Stato copertura', 'Responsabilità manutenzione',
      ...plan.factors.map((factor) => `Punteggio: ${factor.label}`),
      ...plan.categories.map((category) => `${category.name} EUR`),
    ];
    const rows = ranked.map((cinema) => [
      cinema.name,
      formatNumber(centsToEuros(cinema.result?.amountCents ?? 0), 2),
      formatNumber(cinema.result?.coefficient ?? 0, 3),
      formatNumber((cinema.result?.share ?? 0) * 100, 2),
      cinema.attendance === null ? '' : formatNumber(cinema.attendance),
      cinema.areaSqm === null ? '' : formatNumber(cinema.areaSqm),
      cinema.openingYear ?? '',
      cinema.renovated === null ? '' : cinema.renovated ? 'Sì' : 'No',
      cinema.locationType === 'standalone' ? 'Sede indipendente' : cinema.locationType === 'mall' ? 'Centro commerciale' : '',
      cinema.roofStatus === 'new' ? 'Nuova / rifatta' : cinema.roofStatus === 'old' ? 'Datata / da verificare' : '',
      cinema.maintenanceResponsibility === 'operator' ? 'Operatore' : cinema.maintenanceResponsibility === 'property' ? 'Proprietà' : cinema.maintenanceResponsibility === 'shared' ? 'Condivisa' : '',
      ...plan.factors.map((factor) => cinema.factorScores[factor.id] ?? ''),
      ...(cinema.result?.categories ?? []).map((category) => formatNumber(centsToEuros(category.amountCents), 2)),
    ]);
    const csv = [headers, ...rows].map((row) => row.map((value) => `"${String(value).replace(/"/g, '""')}"`).join(';')).join('\r\n');
    const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    const anchor = document.createElement('a');
    anchor.href = url; anchor.download = 'the-space-ripartizione-budget.csv'; anchor.click(); URL.revokeObjectURL(url);
  };
  const resetPlan = () => {
    if (!window.confirm('Vuoi cancellare i dati importati e iniziare un nuovo piano? L’operazione non può essere annullata.')) return;
    updatePlan({
      cinemas: createDefaultCinemas(), categories: [],
      factors: DEFAULT_PRIORITY_FACTORS.map((factor) => ({ ...factor, label: FACTOR_LABELS[factor.id] ?? factor.label })),
      totalBudget: 0, categoryGroups: {}, importedAt: null,
    });
    setWorkbook(null); setFileName(''); setError(''); setNotice('Piano azzerato.');
  };

  return (
    <div className="workspace" lang="it">
      <header className="topbar">
        <div className="topbar-inner mx-auto flex max-w-[1500px] items-center justify-between px-8 py-4">
          <div className="flex items-center gap-3">
            <div className="flex h-9 w-9 items-center justify-center rounded-lg border border-white/20 bg-white/10"><Landmark size={18} /></div>
            <div><div className="text-[15px] font-semibold tracking-tight">THE SPACE <span className="mx-1 text-white/35">/</span> Manutenzione</div><div className="mt-0.5 font-data text-[9px] uppercase tracking-[.16em] text-white/55">Pianificazione degli investimenti</div></div>
          </div>
          <div className="flex items-center gap-2">
            <span className="desktop-only mr-2 flex items-center gap-2 text-xs text-white/65"><span className="h-1.5 w-1.5 rounded-full bg-[#8ec9a2]" />Salvataggio locale nel browser</span>
            <button className="btn btn-quiet btn-small" type="button" onClick={resetPlan} data-testid="button-reset-plan"><RotateCcw size={13} /> Nuovo piano</button>
            <button className="btn btn-primary btn-small" type="button" onClick={exportCsv} disabled={!plan.cinemas.length} data-testid="button-export-csv"><ArrowDownToLine size={14} /> Esporta CSV</button>
          </div>
        </div>
      </header>

      <main className="workspace-main mx-auto max-w-[1500px] px-8 pb-16 pt-8">
        <section className="fade-in mb-7 flex flex-col justify-between gap-5 lg:flex-row lg:items-end">
          <div>
            <div className="eyebrow mb-2">Manutenzione straordinaria · Piano {new Date().getFullYear()}</div>
            <h1 className="font-display text-[38px] leading-none tracking-[-.035em] text-[#25353d] sm:text-[48px]">Un piano, scelte trasparenti.</h1>
            <p className="mt-3 max-w-[650px] text-[14px] leading-6 text-[#69736f]">Ripartisci il budget annuale tra i cinema partendo dai dati del tuo file. Definisci le priorità, verifica l’effetto di ogni scelta e conserva il piano su questo dispositivo.</p>
          </div>
          <div className="flex gap-2 text-xs">
            <a href="#source" className="btn btn-quiet no-underline"><FileSpreadsheet size={14} /> Dati di partenza</a>
            <a href="#priorities" className="btn btn-quiet no-underline"><SlidersHorizontal size={14} /> Priorità</a>
            <a href="#allocation" className="btn btn-quiet no-underline"><Wallet size={14} /> Ripartizione</a>
          </div>
        </section>

        <section className="mb-6 grid gap-3 sm:grid-cols-3">
          <div className="panel flex items-center gap-4 p-4"><div className="flex h-10 w-10 items-center justify-center rounded-lg bg-[#e1eee8] text-[#286653]"><Building2 size={18} /></div><div><div className="eyebrow">Portafoglio</div><div className="mt-1 text-[20px] font-semibold tracking-tight">{formatNumber(plan.cinemas.length)}<span className="ml-1 text-sm font-normal text-[#818780]">cinema</span></div></div></div>
          <div className="panel flex items-center gap-4 p-4"><div className="flex h-10 w-10 items-center justify-center rounded-lg bg-[#f6e8dc] text-[#bb6841]"><Wallet size={18} /></div><div className="min-w-0 flex-1"><div className="eyebrow">Budget annuale · EUR</div><input className="mt-1 w-full max-w-[220px] border-0 bg-transparent p-0 font-data text-[20px] font-medium tracking-tight outline-none focus:ring-0" type="number" min="0" step="1000" value={plan.totalBudget || ''} placeholder="Inserisci il budget" onChange={(event) => updatePlan({ totalBudget: Math.max(0, Number(event.target.value) || 0) })} data-testid="input-total-budget" /></div></div>
          <div className="panel flex items-center gap-4 p-4"><div className="flex h-10 w-10 items-center justify-center rounded-lg bg-[#e7edf0] text-[#477b8c]"><Gauge size={18} /></div><div><div className="eyebrow">Categorie storiche</div><div className="mt-1 text-[20px] font-semibold tracking-tight">{formatNumber(plan.categories.length)}<span className="ml-1 text-sm font-normal text-[#818780]">da Causale</span></div></div></div>
        </section>

        {(error || notice) && <div className={`mb-5 flex items-start gap-2 rounded-lg border px-4 py-3 text-[13px] ${error ? 'border-[#e8b9aa] bg-[#fbefeb] text-[#8e3f2e]' : 'border-[#bad5c5] bg-[#eaf3ed] text-[#2c634e]'}`} role="status" data-testid="status-import-message">
          {error ? <CircleAlert size={16} className="mt-0.5 shrink-0" /> : <Check size={16} className="mt-0.5 shrink-0" />}<span>{error || notice}</span><button className="ml-auto text-current opacity-70" onClick={() => { setError(''); setNotice(''); }} aria-label="Chiudi messaggio">×</button>
        </div>}

        <section id="source" className="panel fade-in mb-6 p-5 sm:p-6">
          <div className="mb-4 flex flex-wrap items-start justify-between gap-3">
            <div className="flex items-start gap-3"><div className="flex h-9 w-9 items-center justify-center rounded-lg bg-[#e7eee8] text-[#286653]"><FileSpreadsheet size={17} /></div><div><div className="eyebrow">01 — File di partenza</div><h2 className="mt-1 text-[18px] font-semibold tracking-tight">Importa i dati del portafoglio</h2><p className="mt-1 text-[12px] text-[#7d837d]">.xlsx, .xls o .csv · elaborato sul dispositivo, senza caricamento online</p></div></div>
            {plan.importedAt && <span className="rounded-full bg-[#e6f0e8] px-3 py-1.5 text-[11px] font-medium text-[#286653]">Ultima importazione salvata</span>}
          </div>
          <div className="grid gap-4 lg:grid-cols-[minmax(250px,1fr)_1.1fr]">
            <div onDragOver={(event) => event.preventDefault()} onDrop={(event) => { event.preventDefault(); void handleWorkbook(event.dataTransfer.files[0]); }} className="flex min-h-[108px] items-center justify-between gap-3 rounded-lg border border-dashed border-[#c9c6ba] bg-[#f8f6f0] px-4 py-3">
              <div className="flex min-w-0 items-center gap-3"><div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-lg bg-white text-[#62746d]"><Upload size={17} /></div><div className="min-w-0"><div className="truncate text-sm font-semibold">{busy ? 'Lettura del file in corso…' : fileName || 'Scegli un file'}</div><div className="mt-1 text-[11px] text-[#7d837d]">Trascina qui il file oppure cercalo sul dispositivo</div></div></div>
              <button type="button" className="btn btn-quiet shrink-0" onClick={() => fileRef.current?.click()} disabled={busy} data-testid="button-choose-workbook">Sfoglia <ArrowRight size={13} /></button>
              <input ref={fileRef} type="file" accept=".xls,.xlsx,.csv" className="hidden" onChange={(event) => void handleWorkbook(event.target.files?.[0])} data-testid="input-workbook-file" />
            </div>
             {workbook ? <div className="grid content-center gap-3 sm:grid-cols-2">
              <label className="text-[11px] font-medium text-[#77817d]">FOGLIO<select className="input mt-1.5 py-2 text-[13px]" value={sheetName} onChange={(event) => changeSheet(event.target.value)} data-testid="select-workbook-sheet">{workbook.sheets.map((sheet) => <option key={sheet.name} value={sheet.name}>{sheet.name}</option>)}</select></label>
              <label className="text-[11px] font-medium text-[#77817d]">ANNO DEI CONSUNTIVI<select className="input mt-1.5 py-2 text-[13px]" value={selectedYear || ''} onChange={(event) => setSelectedYear(Number(event.target.value))} data-testid="select-import-year"><option value="">Scegli un anno</option>{years.map((year) => <option key={year} value={year}>{formatNumber(year)}</option>)}</select></label>
             </div> : <div className="flex items-center gap-2 rounded-lg bg-[#f4f1e9] px-4 py-3 text-[12px] leading-5 text-[#79817b]"><ClipboardList size={16} className="shrink-0" />Importa il file della contabilità. La colonna Causale è la fonte iniziale per le sottocategorie di spesa; non sono presenti dati dimostrativi.</div>}
          </div>
          {workbook && activeSheet && <div className="mt-4 border-t border-[#e9e5dc] pt-4">
            <div className="mb-3 flex flex-wrap items-center justify-between gap-3"><div><div className="text-[13px] font-semibold">Associa le colonne del file</div><div className="mt-1 text-[11px] text-[#7d837d]">I campi vengono riconosciuti quando possibile. Se necessario, seleziona le colonne corrette.</div></div><div className="flex items-center gap-2"><span className="text-[11px] text-[#7d837d]">{formatNumber(activeSheet.rows.length)} righe nel file</span><button type="button" className="btn btn-primary btn-small" disabled={!canSummarize || !years.length} onClick={importWorkbook} data-testid="button-import-workbook"><Check size={13} /> Importa l’anno selezionato</button></div></div>
            <div className="grid gap-2 sm:grid-cols-2 lg:grid-cols-4">
              {([
                ['cinemaField', 'Cinema / centro di costo *'], ['categoryField', `${mapping.categoryField || 'Causale'} / categoria *`], ['amountField', 'Importo consuntivo *'], ['yearField', 'Anno di competenza *'],
              ] as [keyof ImportMapping, string][]).map(([key, label]) => <label key={key} className="text-[10px] font-medium uppercase tracking-wider text-[#77817d]">{label}<select className="input mt-1.5 py-2 text-[12px] normal-case tracking-normal" value={mapping[key]} onChange={(event) => setMapping((current) => ({ ...current, [key]: event.target.value }))} data-testid={`select-map-${key}`}><option value="">— scegli una colonna —</option>{activeSheet.headers.map((header) => <option key={header} value={header}>{header}</option>)}</select></label>)}
            </div>
            {canSummarize && summary && <div className="mt-4 grid gap-4 lg:grid-cols-[1fr_1.1fr]">
              <div className="rounded-lg border border-[#e9e5dc] bg-[#fdfcf8] p-4">
                <div className="mb-2 flex items-center justify-between"><div><div className="text-[12px] font-semibold">Raggruppa i valori di {expenseSourceLabel}</div><div className="mt-1 text-[10px] text-[#7d837d]">Modifica una descrizione per riunire più voci nella stessa categoria.</div></div><span className="font-data text-[10px] text-[#78817b]">{formatNumber(sourceCategories.length)} voci di origine</span></div>
                {sourceCategories.length ? <div className="max-h-48 space-y-2 overflow-auto pr-1">{sourceCategories.map((source) => <label key={source} className="grid items-center gap-2 sm:grid-cols-[1fr_20px_1fr]"><span className="truncate text-[11px] text-[#58635d]" title={source}>{source}</span><ArrowRight size={12} className="hidden text-[#9da098] sm:block" /><input className="input py-1.5 text-[11px]" value={plan.categoryGroups[source] ?? source} aria-label={`Raggruppa ${source} in`} onChange={(event) => updatePlan({ categoryGroups: { ...plan.categoryGroups, [source]: event.target.value } })} data-testid={`input-category-group-${source}`} /></label>)}</div> : <p className="py-3 text-[11px] text-[#7d837d]">In questo foglio non sono presenti valori in {expenseSourceLabel}.</p>}
              </div>
              <div className="rounded-lg border border-[#e9e5dc] bg-[#fdfcf8] p-4">
                <div className="mb-3 flex items-center justify-between"><div><div className="text-[12px] font-semibold">Anteprima dell’importazione</div><div className="mt-1 text-[10px] text-[#7d837d]">Consuntivi {selectedYear ? `del ${formatNumber(selectedYear)}` : 'dell’anno selezionato'} · sede centrale esclusa</div></div><span className="rounded-full bg-[#e7eee8] px-2.5 py-1 text-[10px] text-[#286653]">{formatNumber(summary.cinemas.length)} cinema</span></div>
                <div className="grid grid-cols-3 gap-2">
                  <div className="rounded-md bg-[#f2efe7] p-2.5"><div className="eyebrow">Spesa consuntiva</div><div className="mt-1 font-data text-[13px]">{formatEUR(summary.totalSpend)}</div></div>
                  <div className="rounded-md bg-[#f2efe7] p-2.5"><div className="eyebrow">Registrazioni</div><div className="mt-1 font-data text-[13px]">{formatNumber(summary.transactionCount)}</div></div>
                  <div className="rounded-md bg-[#f2efe7] p-2.5"><div className="eyebrow">Categorie</div><div className="mt-1 font-data text-[13px]">{formatNumber(summary.categories.length)}</div></div>
                </div>
                {summary.skippedAmountRows > 0 && <p className="mt-2 text-[10px] text-[#a05c3d]">{formatNumber(summary.skippedAmountRows)} righe senza un importo leggibile, escluse dai totali per categoria.</p>}
                {summary.excludedCenters.length > 0 && <p className="mt-2 text-[10px] text-[#6c776f]">Righe amministrative escluse: {summary.excludedCenters.join(', ')}</p>}
                <div className="mt-3 max-h-28 overflow-auto border-t border-[#ebe7df] pt-2">{summary.categories.slice(0, 5).map((category) => <div key={category.id} className="flex justify-between gap-3 py-1 text-[11px]"><span className="truncate">{category.name}</span><span className="shrink-0 font-data">{formatEUR(category.previousYearSpend)}</span></div>)}</div>
              </div>
            </div>}
            <div className="mt-4 table-wrap max-h-40 rounded-lg border border-[#e9e5dc]">
              <table className="w-full min-w-[520px] border-collapse text-left text-[11px]"><thead className="sticky top-0 bg-[#f1eee6] text-[#77817d]"><tr>{activeSheet.headers.map((header) => <th key={header} className="px-3 py-2 font-medium">{header}</th>)}</tr></thead><tbody>{activeSheet.rows.slice(0, 4).map((row, index) => <tr key={index} className="border-t border-[#eeeae1]">{activeSheet.headers.map((header) => <td key={header} className="max-w-48 truncate px-3 py-2 text-[#47534f]">{String(row[header] ?? '—')}</td>)}</tr>)}</tbody></table>
            </div>
          </div>}
        </section>

         <section id="priorities" className="panel fade-in fade-delay-1 mb-6 p-5 sm:p-6">
           <div className="mb-5 flex flex-wrap items-start justify-between gap-3"><div className="flex items-start gap-3"><div className="flex h-9 w-9 items-center justify-center rounded-lg bg-[#f6e8dc] text-[#bb6841]"><SlidersHorizontal size={17} /></div><div><div className="eyebrow">02 — Modello di valutazione</div><h2 className="mt-1 text-[18px] font-semibold tracking-tight">Definisci le priorità</h2><p className="mt-1 text-[12px] text-[#7d837d]">I pesi indicano l’importanza relativa. Il coefficiente è la media ponderata dei punteggi (1–5); la quota del cinema è proporzionale al suo coefficiente. Puoi modificare i punteggi per ogni sede.</p></div></div><div className="rounded-md bg-[#f1eee6] px-3 py-2 font-data text-[11px] text-[#637069]">{totalWeight ? 'Pesi attivi, normalizzati automaticamente' : 'Attiva un fattore per calcolare'}</div></div>
          <div className="grid gap-2">
            {plan.factors.map((factor, index) => <div key={factor.id} className={`grid items-center gap-3 rounded-lg border border-[#ece8df] bg-[#fdfcf8] p-3 sm:grid-cols-[minmax(145px,1fr)_120px_110px_32px] ${!factor.enabled ? 'opacity-55' : ''}`}>
              <label><span className="eyebrow">Fattore prioritario</span><input className="input mt-1.5 py-2 text-[13px] font-medium" value={factor.label} onChange={(event) => updateFactor(factor.id, { label: event.target.value })} data-testid={`input-factor-label-${factor.id}`} /></label>
              <label><span className="eyebrow">Peso</span><input className="input mt-1.5 py-2 text-right font-data text-[13px]" type="number" min="0" step="0.25" value={factor.weight} onChange={(event) => updateFactor(factor.id, { weight: Math.max(0, Number(event.target.value) || 0) })} data-testid={`input-factor-weight-${factor.id}`} /></label>
              <label className="flex items-center gap-2 self-end rounded-md border border-[#ded9ce] bg-[#fffefa] px-3 py-2.5 text-[12px]"><input type="checkbox" checked={factor.enabled} onChange={(event) => updateFactor(factor.id, { enabled: event.target.checked })} data-testid={`checkbox-factor-enabled-${factor.id}`} /> Attivo</label>
              <button type="button" className="flex h-8 w-8 items-center justify-center rounded-md text-[#9a8177] hover:bg-[#f5e6e1] hover:text-[#a24836]" aria-label={`Rimuovi ${factor.label}`} onClick={() => removeFactor(factor.id)} data-testid={`button-remove-factor-${factor.id}`}><Trash2 size={14} /></button>
              <div className="h-1 overflow-hidden rounded-full bg-[#eeeae1] sm:col-span-4"><div className="h-full rounded-full" style={{ width: `${factor.enabled && totalWeight ? Math.max(1, factor.weight / totalWeight * 100) : 0}%`, backgroundColor: FACTOR_COLORS[index % FACTOR_COLORS.length] }} /></div>
            </div>)}
            {!plan.factors.length && <div className="rounded-lg border border-dashed border-[#d5d0c5] p-5 text-center text-sm text-[#77817d]">Aggiungi almeno un fattore prioritario per calcolare il piano.</div>}
          </div>
          <form className="mt-3 flex gap-2" onSubmit={(event) => { event.preventDefault(); addFactor(); }}><input className="input max-w-[300px] py-2 text-[12px]" value={newFactor} onChange={(event) => setNewFactor(event.target.value)} placeholder="Aggiungi un fattore…" aria-label="Nome del nuovo fattore" data-testid="input-new-factor" /><button type="submit" className="btn btn-quiet btn-small" data-testid="button-add-factor"><Plus size={14} /> Aggiungi fattore</button></form>
        </section>

        <section id="allocation" className="panel fade-in fade-delay-2 mb-6 overflow-hidden">
          <div className="flex flex-wrap items-center justify-between gap-3 p-5 sm:px-6">
            <div className="flex items-start gap-3"><div className="flex h-9 w-9 items-center justify-center rounded-lg bg-[#e7edf0] text-[#477b8c]"><Wallet size={17} /></div><div><div className="eyebrow">03 — Ripartizione del portafoglio</div><h2 className="mt-1 text-[18px] font-semibold tracking-tight">Esamina la quota di ogni cinema</h2><p className="mt-1 text-[12px] text-[#7d837d]">La quota dipende dai coefficienti ponderati e dalle spese consuntive dell’anno precedente.</p></div></div>
            <div className="flex flex-wrap items-center gap-2"><input className="input w-[180px] py-2 text-[12px]" value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Cerca un cinema…" aria-label="Cerca cinema" data-testid="input-search-cinemas" /><button type="button" className="btn btn-quiet btn-small" onClick={addOdeon} disabled={plan.cinemas.some((cinema) => normalizedName(cinema.name).includes('odeonmilano'))} data-testid="button-add-odeon"><Plus size={13} /> Aggiungi TSC Odeon Milano</button></div>
          </div>
          {!plan.cinemas.length ? <div className="mx-5 mb-5 flex flex-col items-center rounded-xl border border-dashed border-[#d6d1c5] bg-[#f8f6f0] px-6 py-10 text-center sm:mx-6"><div className="mb-3 flex h-12 w-12 items-center justify-center rounded-full bg-[#e7eee8] text-[#286653]"><Building2 size={21} /></div><h3 className="font-display text-[22px]">Il portafoglio parte dai tuoi dati.</h3><p className="mt-2 max-w-sm text-[13px] leading-5 text-[#77817d]">Importa il file delle strutture qui sopra. Qui troverai i cinema, le spese dell’anno precedente e le quote per categoria: nessun dato viene inventato.</p><button className="btn btn-primary mt-4" onClick={() => document.getElementById('source')?.scrollIntoView({ behavior: 'smooth' })} type="button" data-testid="button-go-to-import"><Upload size={14} /> Importa il file</button></div> : <>
            <div className="table-wrap border-t border-[#e8e4dc]"><table className="budget-table">
              <thead><tr><th>Cinema / sede</th><th>Presenze</th><th>Superficie · m²</th><th>Punteggi priorità · 1–5</th><th>Coefficiente</th><th>Quota assegnata · EUR</th><th /></tr></thead>
              <tbody>{visibleCinemas.map((cinema) => <tr key={cinema.id} data-testid={`row-cinema-${cinema.id}`}>
                <td><button type="button" className="text-left text-[13px] font-semibold text-[#2d3e43] hover:text-[#286653]" onClick={() => setSelectedCinema(selectedCinema === cinema.id ? null : cinema.id)} data-testid={`button-edit-cinema-${cinema.id}`}>{cinema.name}</button><div className="mt-1 text-[10px] text-[#92958e]">{cinema.locationType === 'mall' ? 'Centro commerciale' : cinema.locationType === 'standalone' ? 'Sede indipendente' : 'Sede non indicata'} · {cinema.maintenanceResponsibility === 'operator' ? 'manutenzione a carico dell’operatore' : cinema.maintenanceResponsibility === 'property' ? 'manutenzione a carico della proprietà' : cinema.maintenanceResponsibility === 'shared' ? 'manutenzione condivisa' : 'responsabilità non indicata'}</div></td>
                <td><input className="score-input w-[84px]" aria-label={`Presenze di ${cinema.name}`} type="number" min="0" value={cinema.attendance ?? ''} placeholder="—" onChange={(event) => updateCinema(cinema.id, { attendance: event.target.value === '' ? null : Number(event.target.value) })} data-testid={`input-attendance-${cinema.id}`} /></td>
                <td><input className="score-input w-[76px]" aria-label={`Superficie in metri quadrati di ${cinema.name}`} type="number" min="0" value={cinema.areaSqm ?? ''} placeholder="—" onChange={(event) => updateCinema(cinema.id, { areaSqm: event.target.value === '' ? null : Number(event.target.value) })} data-testid={`input-area-${cinema.id}`} /></td>
                <td><div className="flex flex-wrap gap-1.5">{plan.factors.map((factor, factorIndex) => <label key={factor.id} className="flex items-center gap-1 rounded-md px-1.5 py-1" style={{ backgroundColor: `${FACTOR_COLORS[factorIndex % FACTOR_COLORS.length]}12` }} title={`${factor.label} · lascia vuoto per usare il punteggio automatico`}><span className="h-1.5 w-1.5 rounded-full" style={{ backgroundColor: FACTOR_COLORS[factorIndex % FACTOR_COLORS.length] }} /><input className="score-input !w-[43px] !border-transparent !bg-transparent !px-1 !py-0.5" type="number" min="1" max="5" step="1" placeholder="—" aria-label={`Punteggio personalizzato di ${cinema.name} per ${factor.label}`} value={cinema.factorScores[factor.id] ?? ''} onChange={(event) => {
                  const nextScores = { ...cinema.factorScores };
                  if (event.target.value === '') delete nextScores[factor.id];
                  else nextScores[factor.id] = Math.max(1, Math.min(5, Number(event.target.value) || 1));
                  updateCinema(cinema.id, { factorScores: nextScores });
                }} data-testid={`input-score-${cinema.id}-${factor.id}`} /></label>)}</div></td>
                <td className="font-data text-[12px] text-[#637069]">{cinema.result?.coefficient.toLocaleString('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) ?? '—'}</td>
                <td><div className="font-data text-[14px] font-medium text-[#286653]">{formatEUR(centsToEuros(cinema.result?.amountCents ?? 0))}</div><div className="mt-1 text-[10px] text-[#858b84]">{plan.totalBudget > 0 ? `${formatPercent(cinema.result?.share ?? 0)} del budget` : 'Indica il budget sopra'}</div></td>
                <td><button type="button" className="flex h-8 w-8 items-center justify-center rounded-md text-[#90958d] hover:bg-[#f5e6e1] hover:text-[#a24836]" aria-label={`Rimuovi ${cinema.name}`} onClick={() => { if (window.confirm(`Vuoi rimuovere ${cinema.name} dal piano?`)) updatePlan({ cinemas: plan.cinemas.filter((item) => item.id !== cinema.id) }); }} data-testid={`button-remove-cinema-${cinema.id}`}><Trash2 size={14} /></button></td>
              </tr>)}</tbody>
              <tfoot><tr className="border-t border-[#d8d3c7] bg-[#f8f6f0]"><td colSpan={5} className="px-3 py-3 text-right text-[11px] font-medium uppercase tracking-wider text-[#78817b]">Totale ripartito</td><td className="px-3 py-3 font-data text-[14px] font-medium text-[#286653]">{formatEUR(allocations.reduce((sum, item) => sum + centsToEuros(item.amountCents), 0))}</td><td /></tr></tfoot>
            </table></div>
            {selected && <div className="border-t border-[#e8e4dc] bg-[#f8f6f0] p-5 sm:px-6" data-testid={`panel-cinema-details-${selected.id}`}>
              <div className="mb-4 flex items-center justify-between"><div><div className="eyebrow">Dati della sede e dettaglio della quota</div><h3 className="mt-1 text-[15px] font-semibold">{selected.name}</h3></div><button className="flex items-center gap-1 text-xs text-[#69736f]" onClick={() => setSelectedCinema(null)} type="button" data-testid="button-close-cinema-details"><ChevronDown size={14} /> Chiudi</button></div>
              <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-6">
                <label className="text-[10px] font-medium uppercase tracking-wider text-[#77817d]">Nome del cinema<input className="input mt-1.5 py-2 text-[12px] normal-case tracking-normal" value={selected.name} onChange={(event) => updateCinema(selected.id, { name: event.target.value })} data-testid={`input-cinema-name-${selected.id}`} /></label>
                <label className="text-[10px] font-medium uppercase tracking-wider text-[#77817d]">Anno di apertura<input className="input mt-1.5 py-2 text-[12px] normal-case tracking-normal" type="number" min="1800" max="2100" value={selected.openingYear ?? ''} placeholder="Non indicato" onChange={(event) => updateCinema(selected.id, { openingYear: event.target.value === '' ? null : Number(event.target.value) })} data-testid={`input-opening-year-${selected.id}`} /></label>
                <label className="text-[10px] font-medium uppercase tracking-wider text-[#77817d]">Tipologia della sede<select className="input mt-1.5 py-2 text-[12px] normal-case tracking-normal" value={selected.locationType ?? ''} onChange={(event) => updateCinema(selected.id, { locationType: (event.target.value || null) as Cinema['locationType'] })} data-testid={`select-location-type-${selected.id}`}><option value="">Non indicata</option><option value="standalone">Sede indipendente</option><option value="mall">Centro commerciale</option></select></label>
                <label className="text-[10px] font-medium uppercase tracking-wider text-[#77817d]">Stato della copertura<select className="input mt-1.5 py-2 text-[12px] normal-case tracking-normal" value={selected.roofStatus ?? ''} onChange={(event) => updateCinema(selected.id, { roofStatus: (event.target.value || null) as Cinema['roofStatus'] })} data-testid={`select-roof-status-${selected.id}`}><option value="">Non indicato</option><option value="new">Nuova / rifatta</option><option value="old">Datata / da verificare</option></select></label>
                <label className="text-[10px] font-medium uppercase tracking-wider text-[#77817d]">Manutenzione a carico di<select className="input mt-1.5 py-2 text-[12px] normal-case tracking-normal" value={selected.maintenanceResponsibility ?? ''} onChange={(event) => updateCinema(selected.id, { maintenanceResponsibility: (event.target.value || null) as Cinema['maintenanceResponsibility'] })} data-testid={`select-maintenance-${selected.id}`}><option value="">Non indicata</option><option value="operator">Operatore</option><option value="property">Proprietà</option><option value="shared">Condivisa</option></select></label>
                 <label className="text-[10px] font-medium uppercase tracking-wider text-[#77817d]">Ristrutturato<select className="input mt-1.5 py-2 text-[12px] normal-case tracking-normal" value={selected.renovated === null ? '' : selected.renovated ? 'yes' : 'no'} onChange={(event) => updateCinema(selected.id, { renovated: event.target.value === '' ? null : event.target.value === 'yes' })} data-testid={`select-renovated-${selected.id}`}><option value="">Non indicato</option><option value="yes">Sì</option><option value="no">No</option></select></label>
              </div>
              {selected.result?.categories.length ? <div className="mt-5 border-t border-[#e4dfd5] pt-4"><div className="mb-2 flex items-center justify-between"><div className="eyebrow">Ripartizione per categoria · consuntivi Causale</div><div className="font-data text-[11px] text-[#286653]">Totale cinema {formatEUR(centsToEuros(selected.result.amountCents))}</div></div><div className="grid gap-2 sm:grid-cols-2 lg:grid-cols-3">{selected.result.categories.map((category) => <div key={category.categoryId} className="flex items-center justify-between gap-2 rounded-md border border-[#e9e5dc] bg-[#fffefa] px-3 py-2"><div className="min-w-0"><div className="truncate text-[11px] font-medium">{category.categoryName}</div><div className="mt-0.5 text-[10px] text-[#8a8e86]">{formatPercent(category.share)} del consuntivo precedente</div></div><div className="shrink-0 font-data text-[12px] text-[#286653]">{formatEUR(centsToEuros(category.amountCents))}</div></div>)}</div></div> : <p className="mt-4 border-t border-[#e4dfd5] pt-4 text-[11px] text-[#818780]">Per questo piano non sono disponibili importi consuntivi per categoria.</p>}
            </div>}
          </>}
        </section>

        <section className="panel overflow-hidden">
          <button className="flex w-full items-center justify-between gap-3 p-5 text-left sm:px-6" type="button" onClick={() => setCategoryOpen(!categoryOpen)} aria-expanded={categoryOpen} data-testid="button-toggle-categories">
            <div className="flex items-start gap-3"><div className="flex h-9 w-9 items-center justify-center rounded-lg bg-[#f4edda] text-[#a3812e]"><ClipboardList size={17} /></div><div><div className="eyebrow">04 — Dettaglio delle spese</div><h2 className="mt-1 text-[18px] font-semibold tracking-tight">Quote effettive per causale dell’anno precedente</h2><p className="mt-1 text-[12px] text-[#7d837d]">La quota di ciascun cinema viene suddivisa secondo il peso dei consuntivi raggruppati.</p></div></div>{categoryOpen ? <ChevronDown size={17} className="text-[#7d837d]" /> : <ChevronRight size={17} className="text-[#7d837d]" />}
          </button>
          {categoryOpen && <div className="border-t border-[#e8e4dc]">
            {!plan.categories.length ? <div className="px-6 py-8 text-center"><div className="font-display text-[19px] text-[#35454a]">Nessuna categoria storica disponibile.</div><p className="mx-auto mt-2 max-w-lg text-[12px] leading-5 text-[#7d837d]">Importa il file contabile dell’anno precedente e associa la colonna Causale agli importi consuntivi. In assenza di dati effettivi, non mostriamo quote ipotetiche.</p></div>
              : <><div className="flex flex-wrap items-center justify-between gap-2 bg-[#f8f6f0] px-5 py-3 text-[11px] text-[#707a73] sm:px-6"><span>{formatNumber(plan.categories.length)} categorie consuntive · {plan.importedAt ? `importate il ${new Date(plan.importedAt).toLocaleDateString('it-IT')}` : 'dal file'}</span><span className="font-data">Spesa consuntiva totale {formatEUR(plan.categories.reduce((sum, category) => sum + category.previousYearSpend, 0))}</span></div>
                <div className="table-wrap"><table className="budget-table min-w-[720px]"><thead><tr><th>Categoria di spesa</th><th>Consuntivo precedente · EUR</th><th>Incidenza sul consuntivo</th><th>Quota di budget · EUR</th></tr></thead><tbody>{plan.categories.map((category) => {
                  const totalSpend = plan.categories.reduce((sum, item) => sum + item.previousYearSpend, 0);
                  const share = totalSpend > 0 ? category.previousYearSpend / totalSpend : 0;
                   const allocatedCents = allocations.reduce((sum, allocation) => sum + (allocation.categories.find((item) => item.categoryId === category.id)?.amountCents ?? 0), 0);
                   return <tr key={category.id}><td className="text-[13px] font-medium">{category.name}</td><td className="font-data text-[12px]">{formatEUR(category.previousYearSpend)}</td><td><div className="flex items-center gap-3"><div className="h-1.5 w-28 overflow-hidden rounded-full bg-[#eeeae1]"><div className="h-full rounded-full bg-[#c09b4a]" style={{ width: `${share * 100}%` }} /></div><span className="font-data text-[11px]">{formatPercent(share)}</span></div></td><td className="font-data text-[13px] font-medium text-[#286653]">{formatEUR(centsToEuros(allocatedCents))}</td></tr>;
                })}</tbody></table></div>
                <div className="border-t border-[#e8e4dc] px-5 py-4 text-[11px] leading-5 text-[#858b84] sm:px-6">Apri una riga del portafoglio per vedere la ripartizione del singolo cinema. Gli stessi dettagli sono inclusi nel file CSV.</div>
              </>}
          </div>}
        </section>
        <footer className="mt-6 flex flex-wrap items-center justify-between gap-2 px-1 text-[10px] text-[#8a8e86]"><span>THE SPACE · Piano investimenti e manutenzione</span><span>I dati restano in questo browser · {plan.importedAt ? `Ultima importazione: ${new Date(plan.importedAt).toLocaleDateString('it-IT')}` : 'Nessun file importato'}</span></footer>
      </main>
    </div>
  );
}

export default App;
