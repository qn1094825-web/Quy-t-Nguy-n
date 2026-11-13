import { useEffect, useRef, useState } from 'react';
import { renderAsync } from 'docx-preview';
import { Download, Loader2, ZoomIn, ZoomOut, AlertCircle } from 'lucide-react';

interface DocxPreviewPaneProps {
  fileUrl: string;
  fileName: string;
  refreshToken: string;
}

export function DocxPreviewPane({ fileUrl, fileName, refreshToken }: DocxPreviewPaneProps) {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [zoom, setZoom] = useState(1);

  useEffect(() => {
    let cancelled = false;

    async function loadAndRender() {
      if (!hostRef.current) return;
      setLoading(true);
      setError(null);

      try {
        const res = await fetch(`${fileUrl}?v=${encodeURIComponent(refreshToken)}`);
        if (!res.ok) {
          throw new Error('Không thể tải bản xem trước của file Word.');
        }
        const blob = await res.blob();
        if (cancelled || !hostRef.current) return;

        hostRef.current.innerHTML = '';
        await renderAsync(blob, hostRef.current, undefined, {
          className: 'docx',
          inWrapper: true,
          ignoreWidth: false,
          ignoreHeight: false,
          ignoreFonts: false,
          breakPages: true,
          useBase64URL: true,
        });
      } catch (err) {
        if (!cancelled) {
          setError((err as Error).message || 'Lỗi hiển thị xem trước Word.');
        }
      } finally {
        if (!cancelled) {
          setLoading(false);
        }
      }
    }

    loadAndRender();
    return () => {
      cancelled = true;
    };
  }, [fileUrl, refreshToken]);

  return (
    <div className="flex flex-col border border-slate-200 rounded-lg bg-white overflow-hidden">
      <div className="flex flex-wrap items-center justify-between gap-3 px-4 py-2.5 border-b border-slate-200 bg-slate-50">
        <div className="min-w-0">
          <p className="text-sm font-semibold text-slate-900 truncate">{fileName}</p>
        </div>
        <div className="flex items-center gap-2 shrink-0">
          <div className="flex items-center border border-slate-200 rounded-md bg-white">
            <button
              type="button"
              onClick={() => setZoom((z) => Math.max(0.65, +(z - 0.1).toFixed(2)))}
              className="p-1.5 text-slate-600 hover:text-slate-900 hover:bg-slate-50 transition-colors"
              title="Thu nhỏ"
            >
              <ZoomOut className="w-4 h-4" />
            </button>
            <span className="px-2 text-xs font-mono tabular-nums text-slate-700">
              {Math.round(zoom * 100)}%
            </span>
            <button
              type="button"
              onClick={() => setZoom((z) => Math.min(1.3, +(z + 0.1).toFixed(2)))}
              className="p-1.5 text-slate-600 hover:text-slate-900 hover:bg-slate-50 transition-colors"
              title="Phóng to"
            >
              <ZoomIn className="w-4 h-4" />
            </button>
          </div>
          <a
            href={fileUrl}
            download={fileName}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-md transition-colors whitespace-nowrap shrink-0"
          >
            <Download className="w-3.5 h-3.5" />
            Tải file Word
          </a>
        </div>
      </div>

      <div className="relative bg-slate-100 overflow-auto max-h-[680px] min-h-[420px]">
        {loading && (
          <div className="absolute inset-0 z-10 flex flex-col items-center justify-center bg-white/80 gap-2">
            <Loader2 className="w-6 h-6 text-blue-700 animate-spin" />
            <p className="text-sm text-slate-600">Đang dựng bản xem trước Word...</p>
          </div>
        )}
        {error && (
          <div className="p-8 flex flex-col items-center justify-center text-center gap-2">
            <AlertCircle className="w-6 h-6 text-red-600" />
            <p className="text-sm text-red-700">{error}</p>
          </div>
        )}
        <div
          style={{
            transform: `scale(${zoom})`,
            transformOrigin: 'top center',
            width: zoom !== 1 ? `${100 / zoom}%` : '100%',
          }}
        >
          <div ref={hostRef} className="docx-preview-host w-full" />
        </div>
      </div>
    </div>
  );
}
