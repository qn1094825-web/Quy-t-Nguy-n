import { useMemo, useRef, useState } from 'react';
import {
  AlertCircle,
  Download,
  FileText,
  Loader2,
  MapPin,
  RotateCcw,
  Search,
  Sparkles,
  Trash2,
  Upload,
} from 'lucide-react';
import { api } from '../api';
import type { WordTemplate } from '../shared/types';
import { TemplateMapperModal } from './TemplateMapperModal';

interface TemplateLibraryViewProps {
  templates: WordTemplate[];
  onTemplatesChange: (next: WordTemplate[]) => void;
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

export function TemplateLibraryView({
  templates,
  onTemplatesChange,
}: TemplateLibraryViewProps) {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const reuploadInputRef = useRef<HTMLInputElement | null>(null);
  const [reuploadTargetId, setReuploadTargetId] = useState<string | null>(null);
  const [uploading, setUploading] = useState(false);
  const [restoring, setRestoring] = useState(false);
  const [detectingId, setDetectingId] = useState<string | null>(null);
  const [categoryInput, setCategoryInput] = useState('Hợp đồng & Giao dịch');
  const [searchTerm, setSearchTerm] = useState('');
  const [error, setError] = useState<string | null>(null);
  const [mapperTemplateId, setMapperTemplateId] = useState<string | null>(null);
  const [dragOver, setDragOver] = useState(false);

  const filtered = useMemo(() => {
    const q = searchTerm.trim().toLowerCase();
    if (!q) return templates;
    return templates.filter(
      (t) =>
        t.name.toLowerCase().includes(q) ||
        t.category.toLowerCase().includes(q) ||
        t.originalFileName.toLowerCase().includes(q)
    );
  }, [templates, searchTerm]);

  async function handleFilesUpload(fileList: FileList | null) {
    if (!fileList || fileList.length === 0) return;
    setUploading(true);
    setError(null);
    try {
      const files = Array.from(fileList);
      const res = await api.uploadTemplates(files, categoryInput);
      onTemplatesChange(res.templates);
      if (fileInputRef.current) fileInputRef.current.value = '';
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setUploading(false);
    }
  }

  async function handleDelete(tpl: WordTemplate) {
    setError(null);
    try {
      const res = await api.deleteTemplate(tpl.id);
      onTemplatesChange(res.templates);
    } catch (err) {
      setError((err as Error).message);
    }
  }

  async function handleRestoreSamples() {
    setRestoring(true);
    setError(null);
    try {
      const res = await api.restoreSampleTemplates();
      onTemplatesChange(res.templates);
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setRestoring(false);
    }
  }

  async function handleAutoDetect(tpl: WordTemplate) {
    setDetectingId(tpl.id);
    setError(null);
    try {
      const res = await api.autoDetectTemplate(tpl.id);
      onTemplatesChange(res.templates);
      setMapperTemplateId(tpl.id);
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setDetectingId(null);
    }
  }

  async function handleReuploadFile(fileList: FileList | null) {
    if (!fileList || fileList.length === 0 || !reuploadTargetId) return;
    setUploading(true);
    setError(null);
    try {
      const res = await api.reuploadTemplateFile(reuploadTargetId, fileList[0]);
      onTemplatesChange(res.templates);
      if (reuploadInputRef.current) reuploadInputRef.current.value = '';
      setReuploadTargetId(null);
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setUploading(false);
    }
  }

  return (
    <div className="max-w-6xl mx-auto space-y-6">
      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">
            Kho mẫu Word (.docx)
          </h1>
          <p className="text-sm text-slate-600 mt-1">
            Tải lên một hoặc nhiều mẫu văn bản, lưu cấu hình vị trí điền để dùng lại cho mọi hồ sơ khách hàng.
          </p>
        </div>
        <button
          type="button"
          onClick={handleRestoreSamples}
          disabled={restoring}
          className="inline-flex items-center gap-2 px-3.5 py-2 text-xs font-semibold text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-lg transition-colors whitespace-nowrap shrink-0 self-start sm:self-auto"
        >
          {restoring ? (
            <Loader2 className="w-3.5 h-3.5 animate-spin" />
          ) : (
            <RotateCcw className="w-3.5 h-3.5" />
          )}
          Khôi phục 3 mẫu chuẩn
        </button>
      </div>

      {error && (
        <div className="flex items-center gap-2 px-4 py-3 text-sm text-red-700 bg-red-50 border border-red-200 rounded-lg">
          <AlertCircle className="w-4 h-4 shrink-0" />
          <span>{error}</span>
        </div>
      )}

      {/* Upload Box */}
      <div
        onDragOver={(e) => {
          e.preventDefault();
          setDragOver(true);
        }}
        onDragLeave={() => setDragOver(false)}
        onDrop={(e) => {
          e.preventDefault();
          setDragOver(false);
          handleFilesUpload(e.dataTransfer.files);
        }}
        className={`border-2 border-dashed rounded-xl p-6 bg-white transition-colors ${
          dragOver ? 'border-blue-600 bg-blue-50/50' : 'border-slate-300'
        }`}
      >
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="flex items-start gap-3.5">
            <div className="w-10 h-10 rounded-lg bg-blue-50 border border-blue-100 flex items-center justify-center text-blue-700 shrink-0">
              <Upload className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-sm font-semibold text-slate-900">
                Tải mẫu Word mới lên hệ thống
              </h2>
              <p className="text-xs text-slate-600 mt-1 leading-relaxed">
                Chọn hoặc kéo thả nhiều file <span className="font-mono">.docx</span>. Giữ nguyên hoàn toàn định dạng, bảng biểu, font chữ gốc khi xuất kết quả.
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2.5 shrink-0">
            <input
              type="text"
              value={categoryInput}
              onChange={(e) => setCategoryInput(e.target.value)}
              placeholder="Phân loại mẫu..."
              className="px-3 py-2 text-xs border border-slate-300 rounded-lg bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-600 w-44"
            />
            <input
              ref={fileInputRef}
              type="file"
              multiple
              accept=".docx"
              onChange={(e) => handleFilesUpload(e.target.files)}
              className="hidden"
            />
            <button
              type="button"
              disabled={uploading}
              onClick={() => fileInputRef.current?.click()}
              className="inline-flex items-center gap-2 px-4 py-2 text-sm font-semibold text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-60 rounded-lg transition-colors whitespace-nowrap shrink-0"
            >
              {uploading ? (
                <Loader2 className="w-4 h-4 animate-spin" />
              ) : (
                <Upload className="w-4 h-4" />
              )}
              Chọn file .docx
            </button>
          </div>
        </div>
      </div>

      {/* Template List */}
      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 px-5 py-3.5 border-b border-slate-200">
          <div className="text-sm font-semibold text-slate-900 tabular-nums">
            Mẫu đã lưu ({templates.length})
          </div>
          <div className="relative w-full sm:w-72">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              placeholder="Tìm tên mẫu hoặc phân loại..."
              className="w-full pl-9 pr-3 py-1.5 text-sm border border-slate-200 rounded-lg bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-600"
            />
          </div>
        </div>

        {filtered.length === 0 ? (
          <div className="p-10 text-center text-sm text-slate-500">
            Không có mẫu Word nào phù hợp. Hãy tải lên file .docx hoặc khôi phục mẫu chuẩn.
          </div>
        ) : (
          <div className="divide-y divide-slate-200">
            {filtered.map((tpl) => (
              <div
                key={tpl.id}
                className="px-5 py-4 flex flex-col md:flex-row md:items-center justify-between gap-4 hover:bg-slate-50/80 transition-colors"
              >
                <div className="flex items-start gap-3.5 min-w-0">
                  <div className="w-9 h-9 rounded-lg bg-blue-50 text-blue-700 flex items-center justify-center shrink-0 mt-0.5">
                    <FileText className="w-4 h-4" />
                  </div>
                  <div className="min-w-0">
                    <h3 className="text-sm font-semibold text-slate-900 truncate">
                      {tpl.name}
                    </h3>
                    <div className="flex flex-wrap items-center gap-2 text-xs text-slate-500 mt-1 tabular-nums">
                      <span>{tpl.category}</span>
                      <span aria-hidden="true">·</span>
                      {tpl.fileMissing ? (
                        <span className="text-red-700 font-semibold">
                          Thiếu file .docx gốc trên máy chủ
                        </span>
                      ) : (
                        <span className={tpl.mappings.length === 0 ? 'text-amber-700 font-medium' : ''}>
                          {tpl.mappings.length > 0
                            ? `${tpl.mappings.length} vị trí điền`
                            : 'Chưa có vị trí điền'}
                        </span>
                      )}
                      <span aria-hidden="true">·</span>
                      <span>{formatBytes(tpl.sizeBytes)}</span>
                      <span aria-hidden="true">·</span>
                      <span className="truncate max-w-[220px]">{tpl.originalFileName}</span>
                    </div>
                  </div>
                </div>

                <div className="flex flex-wrap items-center gap-2 shrink-0">
                  {tpl.fileMissing ? (
                    <button
                      type="button"
                      disabled={uploading}
                      onClick={() => {
                        setReuploadTargetId(tpl.id);
                        reuploadInputRef.current?.click();
                      }}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-red-600 hover:bg-red-700 disabled:opacity-60 rounded-md transition-colors whitespace-nowrap"
                    >
                      <Upload className="w-3.5 h-3.5" />
                      Tải lại file .docx
                    </button>
                  ) : (
                    <>
                      <button
                        type="button"
                        disabled={detectingId === tpl.id}
                        onClick={() => handleAutoDetect(tpl)}
                        className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-blue-700 bg-blue-50 hover:bg-blue-100 disabled:opacity-60 rounded-md transition-colors whitespace-nowrap"
                      >
                        {detectingId === tpl.id ? (
                          <Loader2 className="w-3.5 h-3.5 animate-spin" />
                        ) : (
                          <Sparkles className="w-3.5 h-3.5" />
                        )}
                        AI nhận diện mẫu cũ
                      </button>
                      <button
                        type="button"
                        onClick={() => setMapperTemplateId(tpl.id)}
                        className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-slate-700 bg-white border border-slate-200 hover:bg-slate-50 rounded-md transition-colors whitespace-nowrap"
                      >
                        <MapPin className="w-3.5 h-3.5 text-blue-700" />
                        Cấu hình vị trí
                      </button>
                      <a
                        href={api.getTemplateDownloadUrl(tpl.id)}
                        download={tpl.originalFileName}
                        className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-200 hover:bg-slate-50 rounded-md transition-colors whitespace-nowrap"
                      >
                        <Download className="w-3.5 h-3.5" />
                        Tải gốc
                      </a>
                    </>
                  )}
                  <button
                    type="button"
                    onClick={() => handleDelete(tpl)}
                    className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium text-red-600 hover:bg-red-50 rounded-md transition-colors whitespace-nowrap"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                    Xóa
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      <input
        ref={reuploadInputRef}
        type="file"
        accept=".docx"
        onChange={(e) => handleReuploadFile(e.target.files)}
        className="hidden"
      />

      {mapperTemplateId && (
        <TemplateMapperModal
          templateId={mapperTemplateId}
          onClose={() => setMapperTemplateId(null)}
          onSaved={(updated) => onTemplatesChange(updated)}
        />
      )}
    </div>
  );
}
