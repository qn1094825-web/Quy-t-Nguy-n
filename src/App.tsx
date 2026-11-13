import { useEffect, useMemo, useState } from 'react';
import {
  CheckCircle2,
  Clock,
  FileCheck2,
  FolderOpen,
  Loader2,
  Menu,
  Plus,
  Search,
  Trash2,
  X,
  AlertCircle,
} from 'lucide-react';
import { api } from './api';
import type { CaseStatus, CustomerCase, WordTemplate } from './shared/types';
import { CaseWorkflowView } from './components/CaseWorkflowView';
import { TemplateLibraryView } from './components/TemplateLibraryView';

const STATUS_META: Record<CaseStatus, { label: string; colorClass: string }> = {
  draft: { label: 'Mới tạo', colorClass: 'text-slate-500' },
  extracted: { label: 'Chờ kiểm tra', colorClass: 'text-amber-700' },
  verified: { label: 'Đã xác nhận', colorClass: 'text-blue-700' },
  exported: { label: 'Đã xuất Word', colorClass: 'text-emerald-700' },
};

function StatusIcon({ status }: { status: CaseStatus }) {
  if (status === 'exported') return <CheckCircle2 className="w-3.5 h-3.5 text-emerald-700 shrink-0" />;
  if (status === 'verified') return <FileCheck2 className="w-3.5 h-3.5 text-blue-700 shrink-0" />;
  if (status === 'extracted') return <Clock className="w-3.5 h-3.5 text-amber-700 shrink-0" />;
  return <FolderOpen className="w-3.5 h-3.5 text-slate-400 shrink-0" />;
}

export default function App() {
  const [view, setView] = useState<'cases' | 'templates'>('cases');
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const [templates, setTemplates] = useState<WordTemplate[]>([]);
  const [cases, setCases] = useState<CustomerCase[]>([]);
  const [activeCaseId, setActiveCaseId] = useState<string | null>(null);

  const [caseSearch, setCaseSearch] = useState('');
  const [creatingCase, setCreatingCase] = useState(false);
  const [mobileSidebarOpen, setMobileSidebarOpen] = useState(false);

  useEffect(() => {
    let cancelled = false;
    async function bootstrap() {
      setLoading(true);
      setError(null);
      try {
        const [tplRes, caseRes] = await Promise.all([api.listTemplates(), api.listCases()]);
        if (cancelled) return;
        setTemplates(tplRes.templates);
        setCases(caseRes.cases);
        if (caseRes.cases.length > 0) {
          setActiveCaseId(caseRes.cases[0].id);
        }
      } catch (err) {
        if (!cancelled) setError((err as Error).message);
      } finally {
        if (!cancelled) setLoading(false);
      }
    }
    bootstrap();
    return () => {
      cancelled = true;
    };
  }, []);

  const activeCase = useMemo(
    () => cases.find((c) => c.id === activeCaseId) ?? cases[0] ?? null,
    [cases, activeCaseId]
  );

  const filteredCases = useMemo(() => {
    const q = caseSearch.trim().toLowerCase();
    if (!q) return cases;
    return cases.filter(
      (c) =>
        c.title.toLowerCase().includes(q) ||
        c.customerName.toLowerCase().includes(q)
    );
  }, [cases, caseSearch]);

  async function handleCreateNewCase() {
    setCreatingCase(true);
    setError(null);
    try {
      const res = await api.createCase({
        title: `Hồ sơ khách hàng mới #${cases.length + 1}`,
        customerName: 'Khách hàng mới',
      });
      setCases(res.cases);
      setActiveCaseId(res.caseItem.id);
      setView('cases');
      setMobileSidebarOpen(false);
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setCreatingCase(false);
    }
  }

  async function handleDeleteCase(caseId: string) {
    setError(null);
    try {
      const res = await api.deleteCase(caseId);
      setCases(res.cases);
      if (activeCaseId === caseId) {
        setActiveCaseId(res.cases[0]?.id ?? null);
      }
    } catch (err) {
      setError((err as Error).message);
    }
  }

  function handleCaseUpdated(updated: CustomerCase) {
    setCases((prev) => prev.map((c) => (c.id === updated.id ? updated : c)));
  }

  const sidebarContent = (
    <div className="flex flex-col h-full bg-white border-r border-slate-200">
      <div className="p-4 border-b border-slate-200 space-y-3">
        <div className="flex items-center justify-between">
          <h2 className="text-xs font-bold tracking-wide text-slate-700 uppercase">
            Hồ sơ khách hàng ({cases.length})
          </h2>
          <button
            type="button"
            onClick={() => setMobileSidebarOpen(false)}
            className="lg:hidden p-1 text-slate-500 hover:text-slate-900 rounded"
            aria-label="Đóng danh sách"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="relative">
          <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
          <input
            type="text"
            value={caseSearch}
            onChange={(e) => setCaseSearch(e.target.value)}
            placeholder="Tìm khách hàng..."
            className="w-full pl-8 pr-3 py-1.5 text-xs border border-slate-200 rounded-lg bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-600"
          />
        </div>
      </div>

      <div className="flex-1 overflow-y-auto divide-y divide-slate-100">
        {filteredCases.length === 0 ? (
          <div className="p-6 text-center text-xs text-slate-500">
            Không có hồ sơ phù hợp.
          </div>
        ) : (
          filteredCases.map((c) => {
            const isActive = view === 'cases' && activeCase?.id === c.id;
            const meta = STATUS_META[c.status];
            return (
              <div
                key={c.id}
                className={`group flex items-start justify-between gap-2 px-4 py-3 cursor-pointer transition-colors ${
                  isActive ? 'bg-blue-50/70' : 'hover:bg-slate-50'
                }`}
                onClick={() => {
                  setActiveCaseId(c.id);
                  setView('cases');
                  setMobileSidebarOpen(false);
                }}
              >
                <div className="min-w-0 flex-1">
                  <div
                    className={`text-sm font-semibold truncate ${
                      isActive ? 'text-blue-800' : 'text-slate-900'
                    }`}
                  >
                    {c.title}
                  </div>
                  <div className="flex items-center gap-1.5 text-xs mt-1 tabular-nums">
                    <StatusIcon status={c.status} />
                    <span className={meta.colorClass}>{meta.label}</span>
                    <span className="text-slate-300" aria-hidden="true">
                      ·
                    </span>
                    <span className="text-slate-500">{c.documents.length} tệp</span>
                  </div>
                </div>
                <button
                  type="button"
                  onClick={(e) => {
                    e.stopPropagation();
                    handleDeleteCase(c.id);
                  }}
                  className="p-1.5 text-slate-400 opacity-70 lg:opacity-0 group-hover:opacity-100 hover:text-red-600 hover:bg-red-50 rounded transition-all shrink-0"
                  title="Xóa hồ sơ này"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                </button>
              </div>
            );
          })
        )}
      </div>
    </div>
  );

  return (
    <div className="min-h-screen flex flex-col bg-slate-50 text-slate-900">
      {/* Top Bar Contract: 3 zones */}
      <header className="sticky top-0 z-30 bg-white border-b border-slate-200 px-4 sm:px-6 h-14 flex items-center justify-between gap-8">
        {/* Zone 1: Brand Title */}
        <div className="flex items-center gap-3 shrink-0">
          <button
            type="button"
            onClick={() => setMobileSidebarOpen(true)}
            className="lg:hidden p-1.5 text-slate-600 hover:text-slate-900 hover:bg-slate-100 rounded-md"
            aria-label="Mở danh sách hồ sơ"
          >
            <Menu className="w-5 h-5" />
          </button>
          <button
            type="button"
            onClick={() => setView('cases')}
            className="flex items-center gap-2.5 text-left whitespace-nowrap shrink-0"
          >
            <span className="w-7 h-7 rounded-md bg-blue-700 text-white font-bold text-xs flex items-center justify-center tracking-tight">
              HS
            </span>
            <span className="text-base font-bold tracking-tight text-slate-900">
              HỒ SƠ AI
            </span>
          </button>
        </div>

        {/* Zone 2: Concise single-line nav links */}
        <nav className="flex items-center gap-6 text-sm font-medium">
          <button
            type="button"
            onClick={() => setView('cases')}
            className={`py-4 border-b-2 transition-colors whitespace-nowrap shrink-0 ${
              view === 'cases'
                ? 'border-blue-700 text-blue-700 font-semibold'
                : 'border-transparent text-slate-600 hover:text-slate-900'
            }`}
          >
            Xử lý hồ sơ
          </button>
          <button
            type="button"
            onClick={() => setView('templates')}
            className={`py-4 border-b-2 transition-colors whitespace-nowrap shrink-0 ${
              view === 'templates'
                ? 'border-blue-700 text-blue-700 font-semibold'
                : 'border-transparent text-slate-600 hover:text-slate-900'
            }`}
          >
            Kho mẫu Word
          </button>
        </nav>

        {/* Zone 3: 1 Primary CTA */}
        <div className="flex items-center shrink-0">
          <button
            type="button"
            disabled={creatingCase}
            onClick={handleCreateNewCase}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-60 rounded-lg transition-colors whitespace-nowrap shrink-0"
          >
            {creatingCase ? (
              <Loader2 className="w-3.5 h-3.5 animate-spin" />
            ) : (
              <Plus className="w-3.5 h-3.5" />
            )}
            <span className="hidden sm:inline">Hồ sơ mới</span>
            <span className="sm:hidden">Mới</span>
          </button>
        </div>
      </header>

      {/* Main Workspace */}
      <div className="flex-1 flex min-h-0">
        {/* Desktop Sidebar */}
        <aside className="hidden lg:block w-72 shrink-0">{sidebarContent}</aside>

        {/* Mobile Drawer Sidebar */}
        {mobileSidebarOpen && (
          <div className="fixed inset-0 z-40 lg:hidden flex">
            <div
              className="fixed inset-0 bg-slate-900/40"
              onClick={() => setMobileSidebarOpen(false)}
            />
            <div className="relative z-10 w-72 max-w-[85vw] h-full">{sidebarContent}</div>
          </div>
        )}

        {/* Content Viewport */}
        <main className="flex-1 min-w-0 p-4 sm:p-6 lg:p-8">
          {error && (
            <div className="max-w-6xl mx-auto mb-5 flex items-center gap-2 px-4 py-3 text-sm text-red-700 bg-red-50 border border-red-200 rounded-lg">
              <AlertCircle className="w-4 h-4 shrink-0" />
              <span>{error}</span>
            </div>
          )}

          {loading ? (
            <div className="flex flex-col items-center justify-center py-24 gap-3">
              <Loader2 className="w-7 h-7 text-blue-700 animate-spin" />
              <p className="text-sm text-slate-600">Đang tải dữ liệu HỒ SƠ AI...</p>
            </div>
          ) : view === 'templates' ? (
            <TemplateLibraryView
              templates={templates}
              onTemplatesChange={setTemplates}
            />
          ) : activeCase ? (
            <CaseWorkflowView
              key={activeCase.id}
              caseItem={activeCase}
              templates={templates}
              onCaseUpdated={handleCaseUpdated}
              onTemplatesUpdated={setTemplates}
            />
          ) : (
            <div className="max-w-lg mx-auto bg-white border border-slate-200 rounded-xl p-10 text-center space-y-4 mt-10">
              <div className="w-12 h-12 rounded-xl bg-blue-50 text-blue-700 flex items-center justify-center mx-auto">
                <FolderOpen className="w-6 h-6" />
              </div>
              <div className="space-y-1">
                <h2 className="text-base font-semibold text-slate-900">
                  Chưa có hồ sơ khách hàng nào
                </h2>
                <p className="text-sm text-slate-500">
                  Tạo hồ sơ khách hàng đầu tiên để chọn mẫu Word, tải CCCD / Sổ đỏ và điền tự động.
                </p>
              </div>
              <button
                type="button"
                onClick={handleCreateNewCase}
                className="inline-flex items-center gap-2 px-4 py-2 text-sm font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-lg"
              >
                <Plus className="w-4 h-4" />
                Tạo hồ sơ mới
              </button>
            </div>
          )}
        </main>
      </div>
    </div>
  );
}
