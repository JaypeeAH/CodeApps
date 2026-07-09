import React, { useState, useEffect, useMemo, useCallback } from 'react';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { YAMLvaultService } from './generated/services/YAMLvaultService';
import { YAMLcategoriesService } from './generated/services/YAMLcategoriesService';
import type { YAMLvaultRead } from './generated/models/YAMLvaultModel';
import type { YAMLcategoriesRead } from './generated/models/YAMLcategoriesModel';
import './App.css';

// ─── Types ────────────────────────────────────────────────────────────────────
interface FormData {
  Title: string;
  Category: string;
  Subcategory: string;
  YAMLcode: string;
  Description: string;
  PreviewUrl: string;
}
interface Toast { id: number; type: 'success' | 'error'; message: string; }

// ─── Helpers ──────────────────────────────────────────────────────────────────
const PALETTE = [
  { bg: '#EEF2FF', text: '#4338CA', dot: '#6366F1' },
  { bg: '#F0FDF4', text: '#166534', dot: '#22C55E' },
  { bg: '#FFF7ED', text: '#9A3412', dot: '#F97316' },
  { bg: '#FDF4FF', text: '#7E22CE', dot: '#A855F7' },
  { bg: '#ECFEFF', text: '#155E75', dot: '#06B6D4' },
  { bg: '#FFF1F2', text: '#9F1239', dot: '#F43F5E' },
  { bg: '#F0F9FF', text: '#075985', dot: '#0EA5E9' },
  { bg: '#FEFCE8', text: '#854D0E', dot: '#EAB308' },
];
const colorCache: Record<string, typeof PALETTE[0]> = {};
function getCategoryColor(cat: string) {
  if (!colorCache[cat]) {
    colorCache[cat] = PALETTE[[...cat].reduce((a, c) => a + c.charCodeAt(0), 0) % PALETTE.length];
  }
  return colorCache[cat];
}

const GRADIENTS = [
  'linear-gradient(135deg,#667eea 0%,#764ba2 100%)',
  'linear-gradient(135deg,#f093fb 0%,#f5576c 100%)',
  'linear-gradient(135deg,#4facfe 0%,#00f2fe 100%)',
  'linear-gradient(135deg,#43e97b 0%,#38f9d7 100%)',
  'linear-gradient(135deg,#fa709a 0%,#fee140 100%)',
  'linear-gradient(135deg,#a18cd1 0%,#fbc2eb 100%)',
  'linear-gradient(135deg,#fda085 0%,#f6d365 100%)',
  'linear-gradient(135deg,#84fab0 0%,#8fd3f4 100%)',
];
function getGradient(title: string) {
  return GRADIENTS[[...(title || '')].reduce((a, c) => a + c.charCodeAt(0), 0) % GRADIENTS.length];
}

const AVATAR_COLORS = ['#6366F1','#8B5CF6','#EC4899','#10B981','#F59E0B','#3B82F6','#EF4444','#14B8A6'];
function getAvatarColor(name?: string) {
  return AVATAR_COLORS[[...(name || '?')].reduce((a, c) => a + c.charCodeAt(0), 0) % AVATAR_COLORS.length];
}
function getInitials(name?: string) {
  if (!name) return '?';
  return name.split(' ').map(n => n[0]).join('').toUpperCase().slice(0, 2);
}

function getPreviewUrl(rec: YAMLvaultRead): string | null {
  // {Thumbnail} = SharePoint connector's thumbnail field (maps to the "Preview" Thumbnail column)
  const t = rec['{Thumbnail}'];
  if (t?.Large)  return t.Large;
  if (t?.Medium) return t.Medium;
  if (t?.Small)  return t.Small;
  // Fallback: try 'Preview' column returned directly (Thumbnail / Image column)
  const p = (rec as any).Preview;
  if (typeof p === 'string' && p) return p;
  if (p && typeof p === 'object') {
    if (p.Large || p.Medium || p.Small) return p.Large || p.Medium || p.Small;
    if (p.serverUrl && p.serverRelativeUrl) return p.serverUrl + p.serverRelativeUrl;
  }
  // Fallback: plain-text PreviewUrl column
  const pu = (rec as any).PreviewUrl;
  if (typeof pu === 'string' && pu) return pu;
  return null;
}

let toastId = 0;

// ─── Authenticated image fetch (APIM proxy URLs need credentials; CSP only allows data:) ───
function useFetchedImage(url: string | null): string | null {
  const [dataUrl, setDataUrl] = useState<string | null>(null);
  useEffect(() => {
    if (!url) { setDataUrl(null); return; }
    let cancelled = false;
    setDataUrl(null);
    fetch(url, { credentials: 'omit' })
      .then(r => { if (!r.ok) throw new Error('fetch failed'); return r.blob(); })
      .then(blob => new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result as string);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
      }))
      .then(data => { if (!cancelled) setDataUrl(data); })
      .catch(() => { if (!cancelled) setDataUrl(null); });
    return () => { cancelled = true; };
  }, [url]);
  return dataUrl;
}

// ─── Image sub-components ──────────────────────────────────────────────────────
const CardPreview: React.FC<{ previewUrl: string | null; title: string; yamlCode?: string }> = ({ previewUrl, title, yamlCode }) => {
  const src = useFetchedImage(previewUrl);
  if (!src) {
    return (
      <div className="card-preview" style={{ background: getGradient(title) }}>
        <div className="preview-code"><code>{yamlCode?.split('\n').slice(0, 4).join('\n') || '# YAML Component'}</code></div>
      </div>
    );
  }
  return (
    <div className="card-preview card-preview--img">
      <img src={src} alt="" className="card-img" />
    </div>
  );
};

const DetailPreview: React.FC<{ previewUrl: string | null; title: string; yamlCode?: string }> = ({ previewUrl, title, yamlCode }) => {
  const src = useFetchedImage(previewUrl);
  if (!src) {
    return (
      <div className="detail-preview" style={{ background: getGradient(title) }}>
        <div className="preview-code preview-code--lg"><code>{yamlCode?.split('\n').slice(0, 10).join('\n') || '# No preview available'}</code></div>
      </div>
    );
  }
  return (
    <div className="detail-preview detail-preview--img">
      <img src={src} alt="" className="detail-img" />
    </div>
  );
};

// ─── SVG Icons (inline, no extra dep) ─────────────────────────────────────────────
const IconPlus    = () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>;
const IconSearch  = () => <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>;
const IconReset   = () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/></svg>;
const IconCopy    = () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>;
const IconCheck   = () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="20 6 9 17 4 12"/></svg>;
const IconEdit    = () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>;
const IconTrash   = () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/><path d="M9 6V4h6v2"/></svg>;
const IconClose   = () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>;
const IconHome    = () => <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>;
const IconAdmin   = () => <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="3"/><path d="M19.07 4.93a10 10 0 0 1 0 14.14M4.93 4.93a10 10 0 0 0 0 14.14"/></svg>;

// ─── App ──────────────────────────────────────────────────────────────────────
const App: React.FC = () => {
  const [records, setRecords]           = useState<YAMLvaultRead[]>([]);
  const [catRecords, setCatRecords]     = useState<YAMLcategoriesRead[]>([]);
  const [loading, setLoading]           = useState(true);
  const [error, setError]               = useState<string | null>(null);
  const [search, setSearch]             = useState('');
  const [selCat, setSelCat]             = useState('');
  const [selSub, setSelSub]             = useState('');
  const [selected, setSelected]         = useState<YAMLvaultRead | null>(null);
  const [isFormOpen, setIsFormOpen]     = useState(false);
  const [isEditing, setIsEditing]       = useState(false);
  const [isDeleteOpen, setIsDeleteOpen] = useState(false);
  const [isSaving, setIsSaving]         = useState(false);
  const [isDeleting, setIsDeleting]     = useState(false);
  const [copied, setCopied]             = useState(false);
  const [toasts, setToasts]             = useState<Toast[]>([]);
  const [cols, setCols]                 = useState(3);
  const [form, setForm]                 = useState<FormData>({ Title: '', Category: '', Subcategory: '', YAMLcode: '', Description: '', PreviewUrl: '' });

  // ─── Data ──────────────────────────────────────────────────────────────────
  const loadRecords = useCallback(async () => {
    try {
      setLoading(true); setError(null);
      const r = await YAMLvaultService.getAll();
      if (r.data) {
        setRecords(r.data);
        if (r.data.length > 0) {
          const first = r.data[0];
          console.log('[YAMLVault] {Thumbnail}:', JSON.stringify(first['{Thumbnail}']));
          console.log('[YAMLVault] Preview col:', JSON.stringify((first as any).Preview));
          console.log('[YAMLVault] getPreviewUrl:', getPreviewUrl(first));
        }
      }
    } catch { setError('Failed to load components. Please try again.'); }
    finally { setLoading(false); }
  }, []);

  useEffect(() => {
    loadRecords();
    YAMLcategoriesService.getAll().then(r => { if (r.data) setCatRecords(r.data); }).catch(() => {});
  }, [loadRecords]);

  const categories = useMemo(() =>
    [...new Set(catRecords.map(c => c.YAMLcategory).filter(Boolean) as string[])].sort()
  , [catRecords]);

  const subcategories = useMemo(() =>
    [...new Set(catRecords.filter(c => !selCat || c.YAMLcategory === selCat).map(c => c.YAMLsubcategory).filter(Boolean) as string[])].sort()
  , [catRecords, selCat]);

  const formSubs = useMemo(() =>
    [...new Set(catRecords.filter(c => !form.Category || c.YAMLcategory === form.Category).map(c => c.YAMLsubcategory).filter(Boolean) as string[])].sort()
  , [catRecords, form.Category]);

  const filtered = useMemo(() => records.filter(r => {
    const q = search.toLowerCase();
    return (!q || r.Title?.toLowerCase().includes(q) || r.Description?.toLowerCase().includes(q) || r.Category?.toLowerCase().includes(q))
        && (!selCat || r.Category === selCat)
        && (!selSub || r.Subcategory === selSub);
  }), [records, search, selCat, selSub]);

  // ─── Toasts ────────────────────────────────────────────────────────────────
  const toast = useCallback((type: Toast['type'], message: string) => {
    const id = ++toastId;
    setToasts(p => [...p, { id, type, message }]);
    setTimeout(() => setToasts(p => p.filter(t => t.id !== id)), 4000);
  }, []);

  // ─── Actions ───────────────────────────────────────────────────────────────
  const copyYAML = useCallback(async () => {
    if (!selected?.YAMLcode) return;
    try {
      await navigator.clipboard.writeText(selected.YAMLcode);
      setCopied(true); setTimeout(() => setCopied(false), 2000);
      toast('success', 'YAML code copied to clipboard!');
    } catch { toast('error', 'Failed to copy to clipboard.'); }
  }, [selected, toast]);

  const openAdd = useCallback(() => {
    setForm({ Title: '', Category: '', Subcategory: '', YAMLcode: '', Description: '', PreviewUrl: '' });
    setIsEditing(false); setIsFormOpen(true);
  }, []);

  const openEdit = useCallback(() => {
    if (!selected) return;
    setForm({ Title: selected.Title || '', Category: selected.Category || '', Subcategory: selected.Subcategory || '', YAMLcode: selected.YAMLcode || '', Description: selected.Description || '', PreviewUrl: (selected as any).PreviewUrl || '' });
    setIsEditing(true); setIsFormOpen(true);
  }, [selected]);

  const save = useCallback(async () => {
    if (!form.Title.trim()) { toast('error', 'Title is required.'); return; }
    try {
      setIsSaving(true);
      const payload = { Title: form.Title, Category: form.Category || undefined, Subcategory: form.Subcategory || undefined, YAMLcode: form.YAMLcode || undefined, Description: form.Description || undefined, PreviewUrl: form.PreviewUrl || undefined };
      if (isEditing && selected?.ID) {
        const r = await YAMLvaultService.update(String(selected.ID), payload);
        if (r.data) { setSelected(r.data); setRecords(p => p.map(x => x.ID === selected.ID ? r.data! : x)); toast('success', 'Component updated!'); }
      } else {
        const r = await YAMLvaultService.create(payload);
        if (r.data) { setRecords(p => [...p, r.data!]); toast('success', 'Component added!'); }
      }
      setIsFormOpen(false);
    } catch { toast('error', 'Failed to save. Please try again.'); }
    finally { setIsSaving(false); }
  }, [form, isEditing, selected, toast]);

  const del = useCallback(async () => {
    if (!selected?.ID) return;
    try {
      setIsDeleting(true);
      await YAMLvaultService.delete(String(selected.ID));
      setRecords(p => p.filter(r => r.ID !== selected.ID));
      setSelected(null); setIsDeleteOpen(false);
      toast('success', 'Component deleted.');
    } catch { toast('error', 'Failed to delete.'); }
    finally { setIsDeleting(false); }
  }, [selected, toast]);

  const reset = () => { setSearch(''); setSelCat(''); setSelSub(''); };

  // ─── Render ────────────────────────────────────────────────────────────────
  return (
    <FluentProvider theme={webLightTheme}>
      <div className="app">

        {/* ── Top Nav ────────────────────────────────────────────────────── */}
        <header className="topnav">
          <div className="topnav-left">
            <div className="topnav-logo">
              <span className="logo-icon">⟨/⟩</span>
              <span className="logo-text">YAML<span className="logo-accent">Vault</span></span>
            </div>
          </div>
          <nav className="topnav-right">
            <a href="#" className="nav-link nav-link--active"><IconHome />Home</a>
            <a href="#" className="nav-link"><IconAdmin />Admin</a>
            <div className="avatar avatar--nav" style={{ background: '#6366F1' }}>JH</div>
          </nav>
        </header>

        <main className={`main${selected ? ' main--split' : ''}`}>

          {/* ── List Panel ─────────────────────────────────────────────── */}
          <div className="list-panel">
            {/* Filter bar */}
            <div className="filter-bar">
              <button className="btn btn--primary" onClick={openAdd}><IconPlus />Add</button>
              <div className="search-wrapper">
                <span className="search-icon"><IconSearch /></span>
                <input className="search-input" placeholder="Search components..." value={search} onChange={e => setSearch(e.target.value)} />
              </div>
              <select className="select-filter" value={selCat} onChange={e => { setSelCat(e.target.value); setSelSub(''); }}>
                <option value="">All</option>
                {categories.map(c => <option key={c} value={c}>{c}</option>)}
              </select>
              <select className="select-filter" value={selSub} onChange={e => setSelSub(e.target.value)}>
                <option value="">All</option>
                {subcategories.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
              <button className="btn btn--ghost" onClick={reset}><IconReset />Reset</button>
              <div className="cols-control">
                {[2, 3, 4, 5].map(n => (
                  <button key={n} className={`cols-btn${cols === n ? ' cols-btn--active' : ''}`} onClick={() => setCols(n)} title={`${n} columns`}>
                    {n}
                  </button>
                ))}
              </div>
            </div>

            {/* Count */}
            <div className="results-header">
              <span className="results-text">
                {loading ? 'Loading…' : `${filtered.length} component${filtered.length !== 1 ? 's' : ''}`}
              </span>
            </div>

            {/* Grid */}
            {loading ? (
              <div className="loading-grid">
                {Array.from({ length: 8 }).map((_, i) => <div key={i} className="card-skeleton" />)}
              </div>
            ) : error ? (
              <div className="empty-state">
                <div className="empty-icon">⚠️</div>
                <p className="empty-title">Failed to load</p>
                <p className="empty-sub">{error}</p>
                <button className="btn btn--primary" onClick={loadRecords}>Retry</button>
              </div>
            ) : filtered.length === 0 ? (
              <div className="empty-state">
                <div className="empty-icon">📦</div>
                <p className="empty-title">No components found</p>
                <p className="empty-sub">Try adjusting your filters or add a new component.</p>
                <button className="btn btn--primary" onClick={openAdd}><IconPlus />Add Component</button>
              </div>
            ) : (
              <div className="card-grid" style={{ gridTemplateColumns: `repeat(${cols}, minmax(0, 1fr))` }}>
                {filtered.map(rec => {
                  const cc = getCategoryColor(rec.Category || 'Other');
                  const owner = rec.Owner?.DisplayName || (rec['Author#Claims'] as string | undefined)?.split('|').pop() || 'Unknown';
                  const isActive = selected?.ID === rec.ID;
                  const previewUrl = getPreviewUrl(rec);
                  return (
                    <div key={rec.ID} className={`component-card${isActive ? ' component-card--selected' : ''}`} onClick={() => setSelected(isActive ? null : rec)}>
                      <div className="card-header">
                        <div className="avatar avatar--sm" style={{ background: getAvatarColor(owner) }}>{getInitials(owner)}</div>
                        <div className="card-meta">
                          <span className="card-title">{rec.Title}</span>
                          <span className="card-sub">{[rec.Category, rec.Subcategory].filter(Boolean).join(' / ')}</span>
                        </div>
                      </div>
                      <CardPreview key={`${rec.ID}-${rec.Modified}`} previewUrl={previewUrl} title={rec.Title || ''} yamlCode={rec.YAMLcode} />
                      <div className="card-footer">
                        {rec.Category && <span className="badge" style={{ background: cc.bg, color: cc.text }}><span className="badge-dot" style={{ background: cc.dot }} />{rec.Category}</span>}
                        {rec.Subcategory && <span className="badge badge--outline">{rec.Subcategory}</span>}
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>

          {/* ── Detail Panel ───────────────────────────────────────────── */}
          {selected && (
            <div className="detail-panel">
              <div className="detail-header">
                <div className="detail-title-group">
                  <div className="avatar" style={{ background: getAvatarColor(selected.Owner?.DisplayName) }}>{getInitials(selected.Owner?.DisplayName)}</div>
                  <div>
                    <h2 className="detail-title">{selected.Title}</h2>
                    <p className="detail-sub">{[selected.Category, selected.Subcategory].filter(Boolean).join(' / ') || 'Uncategorized'}</p>
                  </div>
                </div>
                <div className="detail-actions">
                  <button className="btn btn--sm btn--ghost" onClick={copyYAML} title="Copy YAML code">
                    {copied ? <><IconCheck />Copied!</> : <><IconCopy />Copy</>}
                  </button>
                  <button className="btn btn--sm btn--secondary" onClick={openEdit}><IconEdit />Edit</button>
                  <button className="btn btn--sm btn--danger-ghost" onClick={() => setIsDeleteOpen(true)}><IconTrash />Delete</button>
                  <button className="btn btn--sm btn--ghost btn--icon" onClick={() => setSelected(null)} title="Close"><IconClose /></button>
                </div>
              </div>

              <div className="detail-body">
                <div className="detail-section">
                  <h3 className="section-label">Preview</h3>
                  <DetailPreview
                    key={`${selected.ID}-${selected.Modified}`}
                    previewUrl={getPreviewUrl(selected)}
                    title={selected.Title || ''}
                    yamlCode={selected.YAMLcode}
                  />
                </div>

                {selected.Description && (
                  <div className="detail-section">
                    <h3 className="section-label">Description</h3>
                    <div className="detail-description">
                      <pre className="description-pre">{selected.Description}</pre>
                    </div>
                  </div>
                )}

                <div className="detail-section">
                  <div className="section-header-row">
                    <h3 className="section-label">YAML Code</h3>
                    <button className="btn btn--xs btn--ghost" onClick={copyYAML}>{copied ? '✓ Copied' : 'Copy'}</button>
                  </div>
                  <div className="code-block">
                    <pre className="code-pre">{selected.YAMLcode || '# No YAML code'}</pre>
                  </div>
                </div>

                <div className="detail-section">
                  <h3 className="section-label">Details</h3>
                  <div className="meta-grid">
                    {selected.Category    && <div className="meta-item"><span className="meta-label">Category</span><span className="meta-value">{selected.Category}</span></div>}
                    {selected.Subcategory && <div className="meta-item"><span className="meta-label">Subcategory</span><span className="meta-value">{selected.Subcategory}</span></div>}
                    {selected.Owner?.DisplayName && <div className="meta-item"><span className="meta-label">Owner</span><span className="meta-value">{selected.Owner.DisplayName}</span></div>}
                    {selected.Modified    && <div className="meta-item"><span className="meta-label">Modified</span><span className="meta-value">{new Date(selected.Modified).toLocaleDateString()}</span></div>}
                  </div>
                </div>
              </div>
            </div>
          )}
        </main>

        {/* ── Add / Edit Modal ────────────────────────────────────────────── */}
        {isFormOpen && (
          <div className="modal-overlay" onClick={e => e.target === e.currentTarget && setIsFormOpen(false)}>
            <div className="modal">
              <div className="modal-header">
                <h2 className="modal-title">{isEditing ? 'Edit Component' : 'Add Component'}</h2>
                <button className="btn btn--sm btn--ghost btn--icon" onClick={() => setIsFormOpen(false)}><IconClose /></button>
              </div>
              <div className="modal-body">
                <div className="form-row form-row--2">
                  <div className="form-group">
                    <label className="form-label">Title <span className="required">*</span></label>
                    <input className="form-input" placeholder="Component name" value={form.Title} onChange={e => setForm(p => ({ ...p, Title: e.target.value }))} />
                  </div>
                  <div className="form-group">
                    <label className="form-label">Category</label>
                    <select className="form-select" value={form.Category} onChange={e => setForm(p => ({ ...p, Category: e.target.value, Subcategory: '' }))}>
                      <option value="">Select category</option>
                      {categories.map(c => <option key={c} value={c}>{c}</option>)}
                      {form.Category && !categories.includes(form.Category) && <option value={form.Category}>{form.Category}</option>}
                    </select>
                  </div>
                </div>
                <div className="form-group">
                  <label className="form-label">Subcategory</label>
                  <select className="form-select" value={form.Subcategory} onChange={e => setForm(p => ({ ...p, Subcategory: e.target.value }))}>
                    <option value="">Select subcategory</option>
                    {formSubs.map(s => <option key={s} value={s}>{s}</option>)}
                    {form.Subcategory && !formSubs.includes(form.Subcategory) && <option value={form.Subcategory}>{form.Subcategory}</option>}
                  </select>
                </div>
                <div className="form-group">
                  <label className="form-label">Description</label>
                  <textarea className="form-textarea" rows={3} placeholder="Describe what this component does…" value={form.Description} onChange={e => setForm(p => ({ ...p, Description: e.target.value }))} />
                </div>
                <div className="form-group">
                  <label className="form-label">Preview Image URL</label>
                  <input className="form-input" placeholder="https://… (paste any public image URL)" value={form.PreviewUrl} onChange={e => setForm(p => ({ ...p, PreviewUrl: e.target.value }))} />
                  {form.PreviewUrl && (
                    <div className="form-img-preview">
                      <img src={form.PreviewUrl} alt="Preview" className="form-img" onError={e => { (e.currentTarget as HTMLImageElement).style.display = 'none'; }} />
                    </div>
                  )}
                </div>
                <div className="form-group">
                  <label className="form-label">YAML Code</label>
                  <textarea className="form-textarea form-textarea--code" rows={12} placeholder="# Paste your YAML code here…" value={form.YAMLcode} onChange={e => setForm(p => ({ ...p, YAMLcode: e.target.value }))} />
                </div>
              </div>
              <div className="modal-footer">
                <button className="btn btn--ghost" onClick={() => setIsFormOpen(false)}>Cancel</button>
                <button className="btn btn--primary" onClick={save} disabled={isSaving}>
                  {isSaving ? 'Saving…' : isEditing ? 'Save Changes' : 'Add Component'}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* ── Delete Confirm ──────────────────────────────────────────────── */}
        {isDeleteOpen && (
          <div className="modal-overlay" onClick={e => e.target === e.currentTarget && setIsDeleteOpen(false)}>
            <div className="modal modal--sm">
              <div className="modal-header">
                <h2 className="modal-title">Delete Component</h2>
              </div>
              <div className="modal-body">
                <p className="delete-message">Are you sure you want to delete <strong>{selected?.Title}</strong>? This action cannot be undone.</p>
              </div>
              <div className="modal-footer">
                <button className="btn btn--ghost" onClick={() => setIsDeleteOpen(false)}>Cancel</button>
                <button className="btn btn--danger" onClick={del} disabled={isDeleting}>{isDeleting ? 'Deleting…' : 'Delete'}</button>
              </div>
            </div>
          </div>
        )}

        {/* ── Toasts ──────────────────────────────────────────────────────── */}
        <div className="toast-container">
          {toasts.map(t => (
            <div key={t.id} className={`toast toast--${t.type}`}>
              <span className="toast-icon">{t.type === 'success' ? '✓' : '⚠'}</span>
              <span>{t.message}</span>
            </div>
          ))}
        </div>

      </div>
    </FluentProvider>
  );
};

export default App;
