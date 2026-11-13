import { useEffect, useMemo, useState } from 'react';
import {
  AlertCircle,
  Check,
  Loader2,
  MapPin,
  Plus,
  Search,
  Sparkles,
  Trash2,
  X,
} from 'lucide-react';
import { api } from '../api';
import {
  STANDARD_FIELD_CATALOG,
  normalizeKeyToVietnameseLabel,
  type DocxParagraphInspect,
  type TemplateFieldMapping,
  type WordTemplate,
} from '../shared/types';

interface TemplateMapperModalProps {
  templateId: string;
  onClose: () => void;
  onSaved: (updatedTemplates: WordTemplate[]) => void;
}

export function TemplateMapperModal({
  templateId,
  onClose,
  onSaved,
}: TemplateMapperModalProps) {
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [detecting, setDetecting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [infoMsg, setInfoMsg] = useState<string | null>(null);

  const [name, setName] = useState('');
  const [category, setCategory] = useState('');
  const [paragraphs, setParagraphs] = useState<DocxParagraphInspect[]>([]);
  const [mappings, setMappings] = useState<TemplateFieldMapping[]>([]);

  const [searchTerm, setSearchTerm] = useState('');
  const [selectedPara, setSelectedPara] = useState<DocxParagraphInspect | null>(null);
  const [targetText, setTargetText] = useState('');
  const [replaceScope, setReplaceScope] = useState<
    'text_replace' | 'anchor_replace' | 'insert_after' | 'cell_fill' | 'erase_text'
  >('text_replace');
  const [occurrenceIdx, setOccurrenceIdx] = useState(0);
  const [selectedFieldKey, setSelectedFieldKey] = useState(STANDARD_FIELD_CATALOG[0].key);
  const [customKey, setCustomKey] = useState('');
  const [customLabel, setCustomLabel] = useState('');

  const isEmptyCellPara = (p: DocxParagraphInspect | null) => {
    if (!p) return false;
    const t = p.text.trim();
    return p.locationLabel.includes('Ô bảng') && (t === '' || t === '.' || t === '-');
  };

  useEffect(() => {
    let cancelled = false;
    async function load() {
      setLoading(true);
      setError(null);
      try {
        const res = await api.inspectTemplate(templateId);
        if (cancelled) return;
        setName(res.template.name);
        setCategory(res.template.category);
        setMappings(res.template.mappings);
        setParagraphs(res.paragraphs);
        if (res.paragraphs.length > 0) {
          setSelectedPara(res.paragraphs[0]);
        }
      } catch (err) {
        if (!cancelled) setError((err as Error).message);
      } finally {
        if (!cancelled) setLoading(false);
      }
    }
    load();
    return () => {
      cancelled = true;
    };
  }, [templateId]);

  const filteredParagraphs = useMemo(() => {
    const q = searchTerm.trim().toLowerCase();
    if (!q) return paragraphs;
    return paragraphs.filter(
      (p) =>
        p.text.toLowerCase().includes(q) ||
        p.locationLabel.toLowerCase().includes(q)
    );
  }, [paragraphs, searchTerm]);

  const globalMatchCount = useMemo(() => {
    const clean = targetText.trim().toLowerCase();
    if (!clean) return 0;
    let count = 0;
    for (const p of paragraphs) {
      const hay = p.text.replace(/\u00A0/g, ' ').toLowerCase();
      let idx = 0;
      while ((idx = hay.indexOf(clean, idx)) !== -1) {
        count++;
        idx += clean.length;
      }
    }
    return count;
  }, [paragraphs, targetText]);

  function handleSelectParagraph(p: DocxParagraphInspect) {
    setSelectedPara(p);
    setOccurrenceIdx(0);
    if (isEmptyCellPara(p)) {
      setTargetText('');
      setReplaceScope('cell_fill');
    } else if (p.detectedBlanks.length > 0) {
      setTargetText(p.detectedBlanks[0]);
      setReplaceScope('anchor_replace');
    } else {
      setTargetText('');
      setReplaceScope('text_replace');
    }
  }

  async function handleAutoDetectWithAI() {
    setDetecting(true);
    setError(null);
    setInfoMsg(null);
    try {
      const res = await api.autoDetectTemplate(templateId);
      setMappings(res.template.mappings);
      onSaved(res.templates);
      setInfoMsg(
        `AI đã nhận diện ${res.detectedCount} vị trí thông tin cần thay thế từ mẫu Word.`
      );
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setDetecting(false);
    }
  }

  function handleAddMapping() {
    const cleanTarget = targetText.trim();

    if (replaceScope !== 'cell_fill') {
      if (!cleanTarget) {
        setError('Vui lòng nhập chuỗi chữ cũ hoặc chọn chỗ trống cần thay thế.');
        return;
      }

      if (replaceScope === 'anchor_replace' || replaceScope === 'insert_after') {
        if (!selectedPara) return;
        const normPara = selectedPara.text.replace(/\u00A0/g, ' ');
        if (!normPara.includes(cleanTarget)) {
          setError('Đoạn ký tự này không xuất hiện trong đoạn văn bản đang chọn.');
          return;
        }
      } else if (globalMatchCount === 0) {
        setError('Không tìm thấy chuỗi chữ này trong toàn bộ file Word.');
        return;
      }
    } else if (!selectedPara) {
      return;
    }

    if (replaceScope === 'erase_text') {
      if (globalMatchCount === 0) {
        setError('Không tìm thấy chuỗi chữ này trong toàn bộ file Word.');
        return;
      }
      const eraseMapping: TemplateFieldMapping = {
        id: `map_erase_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
        fieldKey: 'old_info',
        label: 'Xóa bỏ thông tin cũ',
        group: 'custom',
        mode: 'erase_text',
        targetText: cleanTarget,
        matchCase: false,
      };
      setMappings((prev) => [...prev, eraseMapping]);
      setError(null);
      setTargetText('');
      return;
    }

    const useCustom = selectedFieldKey === '__custom__';
    const finalKey = useCustom
      ? customKey
          .trim()
          .toLowerCase()
          .replace(/\s+/g, '_')
          .replace(/[^a-z0-9_]/g, '')
      : selectedFieldKey;

    if (!finalKey) {
      setError('Vui lòng nhập mã trường dữ liệu hợp lệ (vd: so_dien_thoai).');
      return;
    }

    const meta = useCustom
      ? {
          label: customLabel.trim() || normalizeKeyToVietnameseLabel(finalKey).label,
          group: 'custom' as const,
        }
      : normalizeKeyToVietnameseLabel(finalKey);

    let newMapping: TemplateFieldMapping;
    if (replaceScope === 'cell_fill') {
      newMapping = {
        id: `map_cell_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
        fieldKey: finalKey,
        label: meta.label,
        group: meta.group,
        mode: 'cell_fill',
        partName: selectedPara!.partName,
        paragraphIndex: selectedPara!.paragraphIndex,
        targetText: '',
        sampleContext: '[Ô bảng trống]',
      };
    } else if (replaceScope === 'text_replace') {
      newMapping = {
        id: `map_text_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
        fieldKey: finalKey,
        label: meta.label,
        group: meta.group,
        mode: 'text_replace',
        targetText: cleanTarget,
        matchCase: false,
      };
    } else {
      newMapping = {
        id: `map_${replaceScope}_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
        fieldKey: finalKey,
        label: meta.label,
        group: meta.group,
        mode: replaceScope,
        partName: selectedPara!.partName,
        paragraphIndex: selectedPara!.paragraphIndex,
        targetText: cleanTarget,
        occurrenceInParagraph: occurrenceIdx,
        sampleContext: selectedPara!.text.slice(0, 120),
      };
    }

    setMappings((prev) => [...prev, newMapping]);
    setError(null);
    setTargetText('');
    setCustomKey('');
    setCustomLabel('');
  }

  function handleChangeMappingKey(mappingId: string, newFieldKey: string) {
    const meta = normalizeKeyToVietnameseLabel(newFieldKey);
    setMappings((prev) =>
      prev.map((m) =>
        m.id === mappingId
          ? { ...m, fieldKey: newFieldKey, label: meta.label, group: meta.group }
          : m
      )
    );
  }

  function handleUpdateMappingTarget(mappingId: string, nextTarget: string) {
    setMappings((prev) =>
      prev.map((m) =>
        m.id === mappingId ? { ...m, targetText: nextTarget } : m
      )
    );
  }

  function handleRemoveMapping(mappingId: string) {
    setMappings((prev) => prev.filter((m) => m.id !== mappingId));
  }

  async function handleSave() {
    setSaving(true);
    setError(null);
    try {
      const res = await api.updateTemplate(templateId, {
        name,
        category,
        mappings,
      });
      onSaved(res.templates);
      onClose();
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setSaving(false);
    }
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-3 sm:p-6">
      <div className="bg-white border border-slate-200 rounded-xl w-full max-w-6xl max-h-[92vh] flex flex-col overflow-hidden shadow-xl">
        {/* Header */}
        <div className="flex flex-wrap items-center justify-between gap-4 px-5 py-4 border-b border-slate-200 bg-white">
          <div className="min-w-0">
            <h2 className="text-lg font-semibold text-slate-900 truncate">
              Cấu hình vị trí điền mẫu Word
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Hỗ trợ thẻ {'{{ten_truong}}'}, chỗ trống (… / ___), hoặc tự động thay tên &amp; thông tin khách cũ trong mẫu cũ.
            </p>
          </div>
          <div className="flex items-center gap-2 shrink-0">
            <button
              type="button"
              disabled={loading || detecting}
              onClick={handleAutoDetectWithAI}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-blue-700 bg-blue-50 hover:bg-blue-100 disabled:opacity-60 rounded-lg transition-colors whitespace-nowrap"
            >
              {detecting ? (
                <Loader2 className="w-3.5 h-3.5 animate-spin" />
              ) : (
                <Sparkles className="w-3.5 h-3.5" />
              )}
              AI nhận diện vị trí mẫu cũ
            </button>
            <button
              type="button"
              onClick={onClose}
              className="p-2 text-slate-500 hover:text-slate-900 hover:bg-slate-100 rounded-lg transition-colors"
              aria-label="Đóng"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {loading ? (
          <div className="flex-1 flex flex-col items-center justify-center p-12 gap-2">
            <Loader2 className="w-6 h-6 text-blue-700 animate-spin" />
            <p className="text-sm text-slate-600">Đang phân tích cấu trúc file Word...</p>
          </div>
        ) : (
          <div className="flex-1 overflow-y-auto p-5 space-y-5">
            {error && (
              <div className="flex items-center gap-2 px-4 py-3 text-sm text-red-700 bg-red-50 border border-red-200 rounded-lg">
                <AlertCircle className="w-4 h-4 shrink-0" />
                <span>{error}</span>
              </div>
            )}

            {infoMsg && (
              <div className="flex items-center gap-2 px-4 py-3 text-sm text-emerald-800 bg-emerald-50 border border-emerald-200 rounded-lg">
                <Check className="w-4 h-4 shrink-0 text-emerald-700" />
                <span>{infoMsg}</span>
              </div>
            )}

            {/* Template metadata */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1.5">
                  Tên hiển thị mẫu Word
                </label>
                <input
                  type="text"
                  value={name}
                  onChange={(e) => setName(e.target.value)}
                  className="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1.5">
                  Phân loại hồ sơ
                </label>
                <input
                  type="text"
                  value={category}
                  onChange={(e) => setCategory(e.target.value)}
                  placeholder="Ví dụ: Hợp đồng chuyển nhượng, Di sản thừa kế..."
                  className="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600"
                />
              </div>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-12 gap-5">
              {/* Left: Paragraphs from Word document */}
              <div className="lg:col-span-6 flex flex-col border border-slate-200 rounded-lg overflow-hidden bg-slate-50">
                <div className="p-3 bg-white border-b border-slate-200 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-semibold text-slate-800">
                      Nội dung trong mẫu Word ({paragraphs.length} đoạn)
                    </span>
                    <span className="text-xs text-slate-500">Chọn đoạn để gắn vị trí điền</span>
                  </div>
                  <div className="relative">
                    <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                    <input
                      type="text"
                      value={searchTerm}
                      onChange={(e) => setSearchTerm(e.target.value)}
                      placeholder="Tìm dòng văn bản (vd: CCCD, thửa đất, tên cũ...)..."
                      className="w-full pl-9 pr-3 py-1.5 text-sm border border-slate-200 rounded-md bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-600"
                    />
                  </div>
                </div>

                <div className="max-h-64 lg:max-h-80 overflow-y-auto divide-y divide-slate-200">
                  {filteredParagraphs.map((p) => {
                    const isSelected =
                      selectedPara?.partName === p.partName &&
                      selectedPara?.paragraphIndex === p.paragraphIndex;
                    return (
                      <button
                        key={`${p.partName}_${p.paragraphIndex}`}
                        type="button"
                        onClick={() => handleSelectParagraph(p)}
                        className={`w-full text-left px-3.5 py-2.5 transition-colors ${
                          isSelected
                            ? 'bg-blue-50 text-slate-900'
                            : 'bg-white hover:bg-slate-50 text-slate-700'
                        }`}
                      >
                        <div className="flex items-center justify-between gap-2 text-xs text-slate-500 mb-0.5">
                          <span className="font-medium text-blue-700">{p.locationLabel}</span>
                          <span>
                            {p.detectedPlaceholders.length > 0 &&
                              `${p.detectedPlaceholders.length} thẻ`}
                            {p.detectedPlaceholders.length > 0 &&
                              p.detectedBlanks.length > 0 &&
                              ' · '}
                            {p.detectedBlanks.length > 0 &&
                              `${p.detectedBlanks.length} chỗ trống`}
                          </span>
                        </div>
                        {isEmptyCellPara(p) ? (
                          <p className="text-xs italic text-slate-400">(Ô bảng trống)</p>
                        ) : (
                          <p className="text-xs leading-relaxed line-clamp-2">{p.text}</p>
                        )}
                      </button>
                    );
                  })}
                </div>

                {/* Add replacement mapping panel */}
                {selectedPara && (
                  <div className="p-3.5 bg-white border-t border-slate-200 space-y-3">
                    <div className="flex items-center justify-between gap-2">
                      <div className="flex items-center gap-1.5 text-xs font-semibold text-slate-900">
                        <MapPin className="w-3.5 h-3.5 text-blue-700" />
                        Thêm quy tắc thay thế từ {selectedPara.locationLabel}
                      </div>
                      {targetText.trim() && replaceScope !== 'cell_fill' && (
                        <span className="text-[11px] font-medium text-blue-700 tabular-nums">
                          Khớp {globalMatchCount} chỗ trong mẫu
                        </span>
                      )}
                    </div>

                    <p className="text-xs text-slate-600 bg-slate-50 p-2 rounded border border-slate-200 leading-relaxed select-all">
                      {isEmptyCellPara(selectedPara)
                        ? '(Ô bảng trống — chọn trường dữ liệu bên dưới để điền vào ô này)'
                        : selectedPara.text}
                    </p>

                    {selectedPara.detectedBlanks.length > 0 && (
                      <div className="flex flex-wrap items-center gap-1.5">
                        <span className="text-xs text-slate-500">Chỗ trống phát hiện:</span>
                        {selectedPara.detectedBlanks.map((b, idx) => (
                          <button
                            key={`${b}_${idx}`}
                            type="button"
                            onClick={() => {
                              setTargetText(b);
                              setOccurrenceIdx(idx);
                              setReplaceScope('anchor_replace');
                            }}
                            className={`px-2 py-0.5 text-xs font-mono border rounded transition-colors ${
                              targetText === b && occurrenceIdx === idx
                                ? 'border-blue-600 bg-blue-50 text-blue-800'
                                : 'border-slate-300 bg-white text-slate-700 hover:bg-slate-50'
                            }`}
                          >
                            {b} (#{idx + 1})
                          </button>
                        ))}
                      </div>
                    )}

                    <div className="flex flex-wrap items-center gap-3 text-xs">
                      {isEmptyCellPara(selectedPara) && (
                        <label className="inline-flex items-center gap-1.5 cursor-pointer text-blue-700 font-medium">
                          <input
                            type="radio"
                            name="replaceScope"
                            checked={replaceScope === 'cell_fill'}
                            onChange={() => setReplaceScope('cell_fill')}
                            className="text-blue-700 focus:ring-blue-600"
                          />
                          <span>Điền vào ô bảng trống</span>
                        </label>
                      )}
                      <label className="inline-flex items-center gap-1.5 cursor-pointer text-slate-700">
                        <input
                          type="radio"
                          name="replaceScope"
                          checked={replaceScope === 'text_replace'}
                          onChange={() => setReplaceScope('text_replace')}
                          className="text-blue-700 focus:ring-blue-600"
                        />
                        <span>Thay mọi nơi có chuỗi này</span>
                      </label>
                      <label className="inline-flex items-center gap-1.5 cursor-pointer text-slate-700">
                        <input
                          type="radio"
                          name="replaceScope"
                          checked={replaceScope === 'anchor_replace'}
                          onChange={() => setReplaceScope('anchor_replace')}
                          className="text-blue-700 focus:ring-blue-600"
                        />
                        <span>Chỉ thay tại đoạn này</span>
                      </label>
                      <label className="inline-flex items-center gap-1.5 cursor-pointer text-slate-700">
                        <input
                          type="radio"
                          name="replaceScope"
                          checked={replaceScope === 'insert_after'}
                          onChange={() => setReplaceScope('insert_after')}
                          className="text-blue-700 focus:ring-blue-600"
                        />
                        <span>Điền ngay sau nhãn này</span>
                      </label>
                      <label className="inline-flex items-center gap-1.5 cursor-pointer text-red-700">
                        <input
                          type="radio"
                          name="replaceScope"
                          checked={replaceScope === 'erase_text'}
                          onChange={() => setReplaceScope('erase_text')}
                          className="text-red-600 focus:ring-red-600"
                        />
                        <span>Xóa sạch (bỏ thông tin cũ này)</span>
                      </label>
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                      {replaceScope !== 'cell_fill' && (
                        <div className={replaceScope === 'erase_text' ? 'sm:col-span-2' : ''}>
                          <label className="block text-xs text-slate-600 mb-1">
                            {replaceScope === 'insert_after'
                              ? 'Nhãn đứng ngay trước chỗ trống (vd: Ngày lập:)'
                              : 'Chuỗi chữ cũ hoặc chỗ trống cần thay/xóa'}
                          </label>
                          <input
                            type="text"
                            value={targetText}
                            onChange={(e) => setTargetText(e.target.value)}
                            placeholder={
                              replaceScope === 'insert_after'
                                ? 'Nhập nhãn kết thúc bằng dấu :'
                                : 'Copy tên/CCCD/đoạn cũ từ trên dán vào đây...'
                            }
                            className="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600"
                          />
                        </div>
                      )}
                      {replaceScope !== 'erase_text' && (
                        <div className={replaceScope === 'cell_fill' ? 'sm:col-span-2' : ''}>
                          <label className="block text-xs text-slate-600 mb-1">
                            Thay bằng dữ liệu mới từ trường
                          </label>
                          <select
                            value={selectedFieldKey}
                            onChange={(e) => setSelectedFieldKey(e.target.value)}
                            className="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-blue-600"
                          >
                            {STANDARD_FIELD_CATALOG.map((f) => (
                              <option key={f.key} value={f.key}>
                                {f.label} ({f.key})
                              </option>
                            ))}
                            <option value="__custom__">+ Tự định nghĩa mã trường mới...</option>
                          </select>
                        </div>
                      )}
                    </div>

                    {selectedFieldKey === '__custom__' && (
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                        <input
                          type="text"
                          value={customKey}
                          onChange={(e) => setCustomKey(e.target.value)}
                          placeholder="Mã trường (vd: so_dien_thoai)"
                          className="px-2.5 py-1.5 text-xs border border-slate-300 rounded-md"
                        />
                        <input
                          type="text"
                          value={customLabel}
                          onChange={(e) => setCustomLabel(e.target.value)}
                          placeholder="Tên hiển thị (vd: Số điện thoại)"
                          className="px-2.5 py-1.5 text-xs border border-slate-300 rounded-md"
                        />
                      </div>
                    )}

                    <button
                      type="button"
                      onClick={handleAddMapping}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-md transition-colors"
                    >
                      <Plus className="w-3.5 h-3.5" />
                      Thêm vị trí điền này
                    </button>
                  </div>
                )}
              </div>

              {/* Right: Saved mappings */}
              <div className="lg:col-span-6 flex flex-col border border-slate-200 rounded-lg overflow-hidden bg-white">
                <div className="px-4 py-3 border-b border-slate-200 bg-slate-50 flex items-center justify-between">
                  <span className="text-xs font-semibold text-slate-900">
                    Danh sách vị trí sẽ tự động điền ({mappings.length})
                  </span>
                  <span className="text-xs text-slate-500">Lưu vĩnh viễn theo mẫu</span>
                </div>

                {mappings.length === 0 ? (
                  <div className="flex-1 flex flex-col items-center justify-center p-8 text-center text-sm text-slate-500 space-y-3">
                    <p>
                      Chưa có vị trí điền nào. Hãy bấm <strong>&ldquo;AI nhận diện vị trí mẫu cũ&rdquo;</strong> ở góc trên để tự động quét tên, CCCD, địa chỉ cũ trong file Word, hoặc chọn thủ công từng đoạn bên trái.
                    </p>
                  </div>
                ) : (
                  <div className="flex-1 max-h-[500px] overflow-y-auto divide-y divide-slate-200">
                    {mappings.map((m) => (
                      <div key={m.id} className="p-3 flex items-start justify-between gap-3 hover:bg-slate-50">
                        <div className="min-w-0 flex-1 space-y-1.5">
                          <div className="flex flex-wrap items-center gap-2 text-xs text-slate-500">
                            <span
                              className={`font-medium ${
                                m.mode === 'erase_text' ? 'text-red-700' : 'text-blue-700'
                              }`}
                            >
                              {m.mode === 'placeholder'
                                ? 'Thẻ tự động'
                                : m.mode === 'erase_text'
                                  ? 'Xóa sạch chữ cũ'
                                  : m.mode === 'text_replace'
                                    ? 'Thay toàn bộ văn bản'
                                    : m.mode === 'insert_after'
                                      ? `Điền sau nhãn · Đoạn #${(m.paragraphIndex ?? 0) + 1}`
                                      : m.mode === 'cell_fill'
                                        ? `Điền ô bảng #${(m.paragraphIndex ?? 0) + 1}`
                                        : `Đoạn #${(m.paragraphIndex ?? 0) + 1}`}
                            </span>
                            <span aria-hidden="true">·</span>
                            {m.mode === 'placeholder' ? (
                              <span className="font-mono text-slate-700">{m.placeholderTag}</span>
                            ) : m.mode === 'cell_fill' ? (
                              <span className="italic text-slate-500">[Ô bảng trống]</span>
                            ) : (
                              <input
                                type="text"
                                value={m.targetText ?? ''}
                                onChange={(e) => handleUpdateMappingTarget(m.id, e.target.value)}
                                className="px-2 py-0.5 text-xs font-mono text-slate-800 border border-slate-200 rounded bg-white min-w-[160px] flex-1"
                              />
                            )}
                          </div>
                          {m.mode === 'erase_text' ? (
                            <p className="text-[11px] text-slate-500">
                              Sẽ tự động xóa bỏ chuỗi này khi xuất file Word.
                            </p>
                          ) : (
                            <select
                              value={
                                STANDARD_FIELD_CATALOG.some((f) => f.key === m.fieldKey)
                                  ? m.fieldKey
                                  : '__keep_custom__'
                              }
                              onChange={(e) => {
                                if (e.target.value !== '__keep_custom__') {
                                  handleChangeMappingKey(m.id, e.target.value);
                                }
                              }}
                              className="w-full px-2.5 py-1.5 text-xs font-medium text-slate-900 border border-slate-300 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-blue-600"
                            >
                              {!STANDARD_FIELD_CATALOG.some((f) => f.key === m.fieldKey) && (
                                <option value="__keep_custom__">
                                  {m.label} ({m.fieldKey})
                                </option>
                              )}
                              {STANDARD_FIELD_CATALOG.map((f) => (
                                <option key={f.key} value={f.key}>
                                  {f.label} ({f.key})
                                </option>
                              ))}
                            </select>
                          )}
                        </div>
                        <button
                          type="button"
                          onClick={() => handleRemoveMapping(m.id)}
                          className="p-1.5 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded transition-colors shrink-0"
                          title="Xóa vị trí này"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {/* Footer */}
        <div className="flex items-center justify-end gap-3 px-5 py-3.5 border-t border-slate-200 bg-slate-50">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-2 text-sm font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-lg transition-colors"
          >
            Hủy
          </button>
          <button
            type="button"
            disabled={loading || saving}
            onClick={handleSave}
            className="inline-flex items-center gap-2 px-4 py-2 text-sm font-semibold text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-60 rounded-lg transition-colors"
          >
            {saving ? <Loader2 className="w-4 h-4 animate-spin" /> : <Check className="w-4 h-4" />}
            Lưu cấu hình vị trí
          </button>
        </div>
      </div>
    </div>
  );
}
