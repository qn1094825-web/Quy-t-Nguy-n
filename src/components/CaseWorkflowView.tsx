import { useEffect, useMemo, useRef, useState } from 'react';
import {
  AlertCircle,
  ArrowLeft,
  ArrowRight,
  Check,
  CheckCircle2,
  Download,
  Eye,
  FileArchive,
  FileCheck2,
  FileSpreadsheet,
  FileText,
  FolderPlus,
  Image as ImageIcon,
  Loader2,
  MapPin,
  Pencil,
  Plus,
  Search,
  Sparkles,
  Trash2,
  Upload,
} from 'lucide-react';
import { api } from '../api';
import {
  FIELD_GROUP_LABELS,
  STANDARD_FIELD_CATALOG,
  normalizeKeyToVietnameseLabel,
  type CaseFillOptions,
  type CustomerCase,
  type DocumentTypeHint,
  type ExtractedFieldValue,
  type FieldGroup,
  type WordTemplate,
} from '../shared/types';
import { DocxPreviewPane } from './DocxPreviewPane';
import { TemplateMapperModal } from './TemplateMapperModal';

interface CaseWorkflowViewProps {
  caseItem: CustomerCase;
  templates: WordTemplate[];
  onCaseUpdated: (updated: CustomerCase) => void;
  onTemplatesUpdated: (updated: WordTemplate[]) => void;
}

const DOC_TYPE_OPTIONS: Array<{ value: DocumentTypeHint; label: string }> = [
  { value: 'cccd', label: 'CCCD / CMND' },
  { value: 'so_do', label: 'Sổ đỏ / Giấy chứng nhận QSDĐ' },
  { value: 'hop_dong_chuyen_nhuong', label: 'Hợp đồng chuyển nhượng' },
  { value: 'hop_dong_di_san', label: 'Hợp đồng phân chia di sản' },
  { value: 'pdf', label: 'File PDF hồ sơ' },
  { value: 'other', label: 'Giấy tờ khác' },
];

function getInitialStep(c: CustomerCase): 1 | 2 | 3 | 4 {
  if (c.generatedOutputs.length > 0 && c.status === 'exported') return 4;
  if (Object.keys(c.extractedFields).length > 0) return 3;
  if (c.selectedTemplateIds.length > 0) return 2;
  return 1;
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

export function CaseWorkflowView({
  caseItem,
  templates,
  onCaseUpdated,
  onTemplatesUpdated,
}: CaseWorkflowViewProps) {
  const [step, setStep] = useState<1 | 2 | 3 | 4>(() => getInitialStep(caseItem));
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);

  // Header inline rename
  const [editingMeta, setEditingMeta] = useState(false);
  const [draftTitle, setDraftTitle] = useState(caseItem.title);
  const [draftCustomerName, setDraftCustomerName] = useState(caseItem.customerName);

  // Step 1: Inline template upload & mapper
  const inlineTplInputRef = useRef<HTMLInputElement | null>(null);
  const reuploadTplInputRef = useRef<HTMLInputElement | null>(null);
  const [reuploadTplId, setReuploadTplId] = useState<string | null>(null);
  const [uploadingTpl, setUploadingTpl] = useState(false);
  const [mapperTplId, setMapperTplId] = useState<string | null>(null);

  // Step 2: Document upload & AI extraction
  const docInputRef = useRef<HTMLInputElement | null>(null);
  const [docTypeHint, setDocTypeHint] = useState<DocumentTypeHint>('cccd');
  const [uploadingDocs, setUploadingDocs] = useState(false);
  const [analyzing, setAnalyzing] = useState(false);
  const [dragOverDocs, setDragOverDocs] = useState(false);

  // Step 3: Verification form & side-by-side doc viewer
  const [localFields, setLocalFields] = useState<Record<string, ExtractedFieldValue>>(
    caseItem.extractedFields || {}
  );
  const [activeDocId, setActiveDocId] = useState<string | null>(
    caseItem.documents[0]?.id ?? null
  );
  const [fieldSearch, setFieldSearch] = useState('');
  const [onlyTemplateFields, setOnlyTemplateFields] = useState(false);
  const [customFieldKey, setCustomFieldKey] = useState('');
  const [customFieldLabel, setCustomFieldLabel] = useState('');
  const [customFieldValue, setCustomFieldValue] = useState('');
  const [savingAndFilling, setSavingAndFilling] = useState(false);
  const [fillOptions, setFillOptions] = useState<CaseFillOptions>(() => ({
    eraseUnfilledOldText: caseItem.fillOptions?.eraseUnfilledOldText ?? true,
    autoInsertSpousePartyA: caseItem.fillOptions?.autoInsertSpousePartyA ?? false,
    autoInsertSpousePartyB: caseItem.fillOptions?.autoInsertSpousePartyB ?? false,
    customSpouseTextPartyA: caseItem.fillOptions?.customSpouseTextPartyA ?? '',
    customSpouseTextPartyB: caseItem.fillOptions?.customSpouseTextPartyB ?? '',
  }));

  // Step 4: Preview & quick edit
  const [activeOutputFile, setActiveOutputFile] = useState<string | null>(
    caseItem.generatedOutputs[0]?.outputFileName ?? null
  );
  const [previewRefreshToken, setPreviewRefreshToken] = useState<string>(() =>
    String(Date.now())
  );
  const [showQuickEdit, setShowQuickEdit] = useState(false);

  // Reset isolated local state whenever user switches to a different customer dossier
  useEffect(() => {
    setStep(getInitialStep(caseItem));
    setError(null);
    setNotice(null);
    setEditingMeta(false);
    setDraftTitle(caseItem.title);
    setDraftCustomerName(caseItem.customerName);
    setLocalFields(caseItem.extractedFields || {});
    setFillOptions({
      eraseUnfilledOldText: caseItem.fillOptions?.eraseUnfilledOldText ?? true,
      autoInsertSpousePartyA: caseItem.fillOptions?.autoInsertSpousePartyA ?? false,
      autoInsertSpousePartyB: caseItem.fillOptions?.autoInsertSpousePartyB ?? false,
      customSpouseTextPartyA: caseItem.fillOptions?.customSpouseTextPartyA ?? '',
      customSpouseTextPartyB: caseItem.fillOptions?.customSpouseTextPartyB ?? '',
    });
    setActiveDocId(caseItem.documents[0]?.id ?? null);
    setActiveOutputFile(caseItem.generatedOutputs[0]?.outputFileName ?? null);
    setPreviewRefreshToken(caseItem.updatedAt || String(Date.now()));
    setShowQuickEdit(false);
  }, [caseItem.id]);

  // Keep synced when server updates caseItem (e.g. after AI extract or upload)
  useEffect(() => {
    setDraftTitle(caseItem.title);
    setDraftCustomerName(caseItem.customerName);
    setLocalFields(caseItem.extractedFields || {});
    if (caseItem.fillOptions) {
      setFillOptions({
        eraseUnfilledOldText: caseItem.fillOptions.eraseUnfilledOldText ?? true,
        autoInsertSpousePartyA: caseItem.fillOptions.autoInsertSpousePartyA ?? false,
        autoInsertSpousePartyB: caseItem.fillOptions.autoInsertSpousePartyB ?? false,
        customSpouseTextPartyA: caseItem.fillOptions.customSpouseTextPartyA ?? '',
        customSpouseTextPartyB: caseItem.fillOptions.customSpouseTextPartyB ?? '',
      });
    }
    if (caseItem.documents.length > 0 && !caseItem.documents.some((d) => d.id === activeDocId)) {
      setActiveDocId(caseItem.documents[0].id);
    }
    if (
      caseItem.generatedOutputs.length > 0 &&
      !caseItem.generatedOutputs.some((o) => o.outputFileName === activeOutputFile)
    ) {
      setActiveOutputFile(caseItem.generatedOutputs[0].outputFileName);
    }
  }, [caseItem]);

  const selectedTemplates = useMemo(
    () => templates.filter((t) => caseItem.selectedTemplateIds.includes(t.id)),
    [templates, caseItem.selectedTemplateIds]
  );

  const requiredTemplateFieldKeys = useMemo(() => {
    const keys = new Set<string>();
    for (const tpl of selectedTemplates) {
      for (const m of tpl.mappings) {
        keys.add(m.fieldKey);
      }
    }
    return keys;
  }, [selectedTemplates]);

  // Build ordered editable fields merging standard catalog, template-required keys, and extracted fields
  const mergedDisplayFields = useMemo(() => {
    const map = new Map<string, ExtractedFieldValue>();

    for (const def of STANDARD_FIELD_CATALOG) {
      const existing = localFields[def.key];
      map.set(def.key, {
        key: def.key,
        label: existing?.label || def.label,
        group: existing?.group || def.group,
        value: existing?.value || '',
        sourceNote: existing?.sourceNote || '',
      });
    }

    for (const tpl of selectedTemplates) {
      for (const m of tpl.mappings) {
        if (!map.has(m.fieldKey)) {
          const existing = localFields[m.fieldKey];
          map.set(m.fieldKey, {
            key: m.fieldKey,
            label: existing?.label || m.label,
            group: existing?.group || m.group,
            value: existing?.value || '',
            sourceNote: existing?.sourceNote || '',
          });
        }
      }
    }

    for (const [k, v] of Object.entries(localFields)) {
      if (!map.has(k)) {
        map.set(k, v);
      }
    }

    return Array.from(map.values());
  }, [localFields, selectedTemplates]);

  const filteredFields = useMemo(() => {
    const q = fieldSearch.trim().toLowerCase();
    return mergedDisplayFields.filter((f) => {
      if (onlyTemplateFields && requiredTemplateFieldKeys.size > 0 && !requiredTemplateFieldKeys.has(f.key)) {
        return false;
      }
      if (!q) return true;
      return (
        f.label.toLowerCase().includes(q) ||
        f.key.toLowerCase().includes(q) ||
        f.value.toLowerCase().includes(q)
      );
    });
  }, [mergedDisplayFields, fieldSearch, onlyTemplateFields, requiredTemplateFieldKeys]);

  const groupedFields = useMemo(() => {
    const groups: Record<FieldGroup, ExtractedFieldValue[]> = {
      general: [],
      party_a: [],
      party_b: [],
      property: [],
      custom: [],
    };
    for (const f of filteredFields) {
      const g = groups[f.group] ? f.group : 'custom';
      groups[g].push(f);
    }
    return groups;
  }, [filteredFields]);

  const filledRequiredCount = useMemo(() => {
    if (requiredTemplateFieldKeys.size === 0) return 0;
    let count = 0;
    for (const k of requiredTemplateFieldKeys) {
      if ((localFields[k]?.value ?? '').trim() !== '') count++;
    }
    return count;
  }, [requiredTemplateFieldKeys, localFields]);

  async function handleSaveCaseMeta() {
    setError(null);
    try {
      const res = await api.updateCase(caseItem.id, {
        title: draftTitle.trim() || caseItem.title,
        customerName: draftCustomerName.trim() || caseItem.customerName,
      });
      onCaseUpdated(res.caseItem);
      setEditingMeta(false);
    } catch (err) {
      setError((err as Error).message);
    }
  }

  async function handleToggleTemplate(tplId: string) {
    setError(null);
    const exists = caseItem.selectedTemplateIds.includes(tplId);
    const nextIds = exists
      ? caseItem.selectedTemplateIds.filter((id) => id !== tplId)
      : [...caseItem.selectedTemplateIds, tplId];

    try {
      const res = await api.updateCase(caseItem.id, {
        selectedTemplateIds: nextIds,
      });
      onCaseUpdated(res.caseItem);
    } catch (err) {
      setError((err as Error).message);
    }
  }

  async function handleInlineTemplateUpload(fileList: FileList | null) {
    if (!fileList || fileList.length === 0) return;
    setUploadingTpl(true);
    setError(null);
    try {
      const res = await api.uploadTemplates(Array.from(fileList), 'Mẫu văn bản');
      onTemplatesUpdated(res.templates);
      const newIds = res.created.map((t) => t.id);
      const mergedIds = Array.from(new Set([...caseItem.selectedTemplateIds, ...newIds]));
      const updatedCase = await api.updateCase(caseItem.id, {
        selectedTemplateIds: mergedIds,
      });
      onCaseUpdated(updatedCase.caseItem);
      if (inlineTplInputRef.current) inlineTplInputRef.current.value = '';
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setUploadingTpl(false);
    }
  }

  async function handleReuploadMissingTemplate(fileList: FileList | null) {
    if (!fileList || fileList.length === 0 || !reuploadTplId) return;
    setUploadingTpl(true);
    setError(null);
    setNotice(null);
    try {
      const res = await api.reuploadTemplateFile(reuploadTplId, fileList[0]);
      onTemplatesUpdated(res.templates);
      if (reuploadTplInputRef.current) reuploadTplInputRef.current.value = '';
      setNotice(`Đã khôi phục file .docx cho mẫu "${res.template.name}".`);
      setReuploadTplId(null);
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setUploadingTpl(false);
    }
  }

  async function handleUploadDocs(fileList: FileList | null) {
    if (!fileList || fileList.length === 0) return;
    setUploadingDocs(true);
    setError(null);
    try {
      const res = await api.uploadCaseDocuments(caseItem.id, Array.from(fileList), docTypeHint);
      onCaseUpdated(res.caseItem);
      if (docInputRef.current) docInputRef.current.value = '';
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setUploadingDocs(false);
    }
  }

  async function handleDeleteDoc(docId: string) {
    setError(null);
    try {
      const res = await api.deleteCaseDocument(caseItem.id, docId);
      onCaseUpdated(res.caseItem);
    } catch (err) {
      setError((err as Error).message);
    }
  }

  async function handleAnalyzeDocuments() {
    if (caseItem.documents.length === 0) {
      setError('Vui lòng tải lên ít nhất 1 ảnh hoặc file PDF trước khi phân tích.');
      return;
    }
    setAnalyzing(true);
    setError(null);
    setNotice(null);
    try {
      const res = await api.analyzeCase(caseItem.id);
      onCaseUpdated(res.caseItem);
      setLocalFields(res.caseItem.extractedFields || {});
      setNotice('AI đã hoàn tất đọc dữ liệu từ giấy tờ. Vui lòng kiểm tra và chỉnh sửa nếu cần.');
      setStep(3);
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setAnalyzing(false);
    }
  }

  function handleFieldValueChange(field: ExtractedFieldValue, nextVal: string) {
    setLocalFields((prev) => ({
      ...prev,
      [field.key]: {
        ...field,
        value: nextVal,
      },
    }));
  }

  function handleAddCustomField() {
    const cleanKey = customFieldKey
      .trim()
      .toLowerCase()
      .replace(/\s+/g, '_')
      .replace(/[^a-z0-9_]/g, '');
    if (!cleanKey) {
      setError('Vui lòng nhập mã trường (vd: so_dien_thoai).');
      return;
    }
    const meta = normalizeKeyToVietnameseLabel(cleanKey);
    const label = customFieldLabel.trim() || meta.label;
    setLocalFields((prev) => ({
      ...prev,
      [cleanKey]: {
        key: cleanKey,
        label,
        group: 'custom',
        value: customFieldValue,
        sourceNote: 'Nhập thủ công',
      },
    }));
    setCustomFieldKey('');
    setCustomFieldLabel('');
    setCustomFieldValue('');
    setError(null);
  }

  async function handleConfirmAndFill(goToStep4 = true) {
    if (caseItem.selectedTemplateIds.length === 0) {
      setError('Vui lòng chọn ít nhất 1 mẫu Word ở Bước 1 trước khi xuất kết quả.');
      return;
    }

    setSavingAndFilling(true);
    setError(null);
    setNotice(null);
    try {
      const nextCustomerName =
        localFields.ben_mua_ho_ten?.value?.trim() ||
        localFields.ben_ban_ho_ten?.value?.trim() ||
        caseItem.customerName;

      await api.updateCase(caseItem.id, {
        customerName: nextCustomerName,
        extractedFields: localFields,
        fillOptions,
        isVerified: true,
        status: 'verified',
      });

      const fillRes = await api.fillCaseTemplates(caseItem.id);
      if (fillRes.templates) {
        onTemplatesUpdated(fillRes.templates);
      }
      onCaseUpdated(fillRes.caseItem);
      setPreviewRefreshToken(String(Date.now()));
      if (fillRes.outputs.length === 0) {
        setError('Không có file Word nào được tạo. Vui lòng kiểm tra lại mẫu Word đã chọn.');
        return;
      }
      if (!activeOutputFile || !fillRes.outputs.some((o) => o.outputFileName === activeOutputFile)) {
        setActiveOutputFile(fillRes.outputs[0].outputFileName);
      }
      setNotice(`Đã điền dữ liệu vào ${fillRes.outputs.length} mẫu Word thành công.`);
      if (goToStep4) setStep(4);
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setSavingAndFilling(false);
    }
  }

  async function handleAutoDetectTemplateAndRefill(templateId: string, shouldRefill: boolean) {
    setSavingAndFilling(true);
    setError(null);
    setNotice(null);
    try {
      const detectRes = await api.autoDetectTemplate(templateId);
      onTemplatesUpdated(detectRes.templates);
      if (shouldRefill) {
        const fillRes = await api.fillCaseTemplates(caseItem.id);
        if (fillRes.templates) onTemplatesUpdated(fillRes.templates);
        onCaseUpdated(fillRes.caseItem);
        setPreviewRefreshToken(String(Date.now()));
        setNotice(
          `AI đã nhận diện ${detectRes.detectedCount} vị trí trong mẫu cũ và điền lại dữ liệu mới.`
        );
      } else {
        setNotice(`AI đã nhận diện ${detectRes.detectedCount} vị trí cần điền trong mẫu Word.`);
        setMapperTplId(templateId);
      }
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setSavingAndFilling(false);
    }
  }

  const activeDoc = caseItem.documents.find((d) => d.id === activeDocId) ?? caseItem.documents[0];
  const activeOutput =
    caseItem.generatedOutputs.find((o) => o.outputFileName === activeOutputFile) ??
    caseItem.generatedOutputs[0];

  const steps = [
    { num: 1 as const, label: '1. Chọn mẫu Word', sub: `${caseItem.selectedTemplateIds.length} mẫu` },
    { num: 2 as const, label: '2. Tải hồ sơ & AI đọc', sub: `${caseItem.documents.length} tài liệu` },
    {
      num: 3 as const,
      label: '3. Kiểm tra dữ liệu',
      sub: caseItem.isVerified ? 'Đã xác nhận' : 'Chờ kiểm tra',
    },
    {
      num: 4 as const,
      label: '4. Xem trước & Xuất Word',
      sub: `${caseItem.generatedOutputs.length} file`,
    },
  ];

  return (
    <div className="max-w-7xl mx-auto space-y-5">
      {/* Case Title Header */}
      <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-5">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          {editingMeta ? (
            <div className="flex flex-wrap items-center gap-2.5 flex-1">
              <input
                type="text"
                value={draftTitle}
                onChange={(e) => setDraftTitle(e.target.value)}
                placeholder="Tên hồ sơ..."
                className="px-3 py-1.5 text-sm font-semibold border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600 min-w-[220px]"
              />
              <input
                type="text"
                value={draftCustomerName}
                onChange={(e) => setDraftCustomerName(e.target.value)}
                placeholder="Tên khách hàng..."
                className="px-3 py-1.5 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600 min-w-[180px]"
              />
              <button
                type="button"
                onClick={handleSaveCaseMeta}
                className="px-3 py-1.5 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-lg"
              >
                Lưu
              </button>
              <button
                type="button"
                onClick={() => setEditingMeta(false)}
                className="px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-100 rounded-lg"
              >
                Hủy
              </button>
            </div>
          ) : (
            <div className="min-w-0">
              <div className="flex items-center gap-2">
                <h1 className="text-xl font-bold tracking-tight text-slate-900 truncate">
                  {caseItem.title}
                </h1>
                <button
                  type="button"
                  onClick={() => setEditingMeta(true)}
                  className="p-1 text-slate-400 hover:text-blue-700 rounded transition-colors"
                  title="Đổi tên hồ sơ"
                >
                  <Pencil className="w-3.5 h-3.5" />
                </button>
              </div>
              <div className="flex flex-wrap items-center gap-2 text-xs text-slate-500 mt-1 tabular-nums">
                <span>Khách hàng: {caseItem.customerName}</span>
                <span aria-hidden="true">·</span>
                <span>{caseItem.selectedTemplateIds.length} mẫu Word đã chọn</span>
                <span aria-hidden="true">·</span>
                <span>{caseItem.documents.length} giấy tờ tải lên</span>
                <span aria-hidden="true">·</span>
                <span className={caseItem.isVerified ? 'text-emerald-700 font-medium' : 'text-amber-700 font-medium'}>
                  {caseItem.isVerified ? 'Đã xác nhận dữ liệu' : 'Chưa xác nhận'}
                </span>
              </div>
            </div>
          )}

          {/* Step Switcher */}
          <nav className="grid grid-cols-2 sm:grid-cols-4 gap-1.5 bg-slate-100 p-1 rounded-lg shrink-0">
            {steps.map((s) => {
              const active = step === s.num;
              return (
                <button
                  key={s.num}
                  type="button"
                  onClick={() => setStep(s.num)}
                  className={`px-3 py-2 rounded-md text-left transition-colors ${
                    active
                      ? 'bg-white text-blue-700 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  <div className="text-xs font-semibold whitespace-nowrap truncate">{s.label}</div>
                  <div className="text-[11px] text-slate-500 tabular-nums whitespace-nowrap truncate">
                    {s.sub}
                  </div>
                </button>
              );
            })}
          </nav>
        </div>
      </div>

      {error && (
        <div className="flex items-center gap-2 px-4 py-3 text-sm text-red-700 bg-red-50 border border-red-200 rounded-lg">
          <AlertCircle className="w-4 h-4 shrink-0" />
          <span>{error}</span>
        </div>
      )}

      {notice && (
        <div className="flex items-center gap-2 px-4 py-3 text-sm text-emerald-800 bg-emerald-50 border border-emerald-200 rounded-lg">
          <CheckCircle2 className="w-4 h-4 shrink-0 text-emerald-700" />
          <span>{notice}</span>
        </div>
      )}

      {/* ===================================================================== */}
      {/* STEP 1: CHỌN MẪU WORD                                                */}
      {/* ===================================================================== */}
      {step === 1 && (
        <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 px-5 py-4 border-b border-slate-200">
            <div>
              <h2 className="text-base font-semibold text-slate-900">
                Chọn một hoặc nhiều mẫu Word để điền cho hồ sơ này
              </h2>
              <p className="text-xs text-slate-500 mt-0.5">
                AI sẽ tự đọc giấy tờ và điền đồng thời vào tất cả các mẫu bạn đã chọn, giữ nguyên bố cục và định dạng gốc.
              </p>
            </div>
            <div className="flex items-center gap-2 shrink-0">
              <input
                ref={inlineTplInputRef}
                type="file"
                multiple
                accept=".docx"
                onChange={(e) => handleInlineTemplateUpload(e.target.files)}
                className="hidden"
              />
              <button
                type="button"
                disabled={uploadingTpl}
                onClick={() => inlineTplInputRef.current?.click()}
                className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-blue-700 bg-blue-50 hover:bg-blue-100 rounded-lg transition-colors whitespace-nowrap"
              >
                {uploadingTpl ? (
                  <Loader2 className="w-3.5 h-3.5 animate-spin" />
                ) : (
                  <Upload className="w-3.5 h-3.5" />
                )}
                Tải thêm mẫu .docx
              </button>
            </div>
          </div>

          {templates.length === 0 ? (
            <div className="p-10 text-center text-sm text-slate-500">
              Chưa có mẫu Word nào trong kho. Vui lòng tải file .docx lên để bắt đầu.
            </div>
          ) : (
            <div className="divide-y divide-slate-200">
              {templates.map((tpl) => {
                const checked = caseItem.selectedTemplateIds.includes(tpl.id);
                return (
                  <div
                    key={tpl.id}
                    className={`px-5 py-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3 transition-colors ${
                      checked ? 'bg-blue-50/40' : 'hover:bg-slate-50'
                    }`}
                  >
                    <label className="flex items-start gap-3.5 cursor-pointer flex-1 min-w-0">
                      <input
                        type="checkbox"
                        checked={checked}
                        disabled={Boolean(tpl.fileMissing && !checked)}
                        onChange={() => handleToggleTemplate(tpl.id)}
                        className="mt-1 w-4 h-4 rounded border-slate-300 text-blue-700 focus:ring-blue-600 disabled:opacity-40"
                      />
                      <div className="min-w-0">
                        <div className="text-sm font-semibold text-slate-900 truncate">
                          {tpl.name}
                        </div>
                        <div className="flex flex-wrap items-center gap-2 text-xs text-slate-500 mt-0.5 tabular-nums">
                          <span>{tpl.category}</span>
                          <span aria-hidden="true">·</span>
                          {tpl.fileMissing ? (
                            <span className="text-red-700 font-semibold">
                              Thiếu file .docx gốc — cần tải lại
                            </span>
                          ) : (
                            <span className={tpl.mappings.length === 0 ? 'text-amber-700 font-medium' : ''}>
                              {tpl.mappings.length > 0
                                ? `${tpl.mappings.length} vị trí điền`
                                : 'Chưa nhận diện vị trí điền'}
                            </span>
                          )}
                          <span aria-hidden="true">·</span>
                          <span className="truncate max-w-[200px]">{tpl.originalFileName}</span>
                        </div>
                      </div>
                    </label>

                    <div className="flex flex-wrap items-center gap-2 self-start sm:self-center shrink-0">
                      {tpl.fileMissing ? (
                        <button
                          type="button"
                          disabled={uploadingTpl}
                          onClick={() => {
                            setReuploadTplId(tpl.id);
                            reuploadTplInputRef.current?.click();
                          }}
                          className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-red-600 hover:bg-red-700 disabled:opacity-60 rounded-md transition-colors"
                        >
                          <Upload className="w-3.5 h-3.5" />
                          Tải lại .docx
                        </button>
                      ) : (
                        <>
                          <button
                            type="button"
                            disabled={savingAndFilling}
                            onClick={() => handleAutoDetectTemplateAndRefill(tpl.id, false)}
                            className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-blue-700 bg-blue-50 hover:bg-blue-100 disabled:opacity-60 rounded-md transition-colors"
                          >
                            <Sparkles className="w-3.5 h-3.5" />
                            AI nhận diện mẫu cũ
                          </button>
                          <button
                            type="button"
                            onClick={() => setMapperTplId(tpl.id)}
                            className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-200 hover:bg-slate-50 rounded-md transition-colors"
                          >
                            <MapPin className="w-3.5 h-3.5 text-blue-700" />
                            Cấu hình vị trí
                          </button>
                        </>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          <input
            ref={reuploadTplInputRef}
            type="file"
            accept=".docx"
            onChange={(e) => handleReuploadMissingTemplate(e.target.files)}
            className="hidden"
          />

          <div className="px-5 py-4 border-t border-slate-200 bg-slate-50 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div className="text-xs text-slate-600 tabular-nums">
              Đã chọn <strong className="text-slate-900">{caseItem.selectedTemplateIds.length}</strong> mẫu Word
              {requiredTemplateFieldKeys.size > 0 &&
                ` · Tổng cộng ${requiredTemplateFieldKeys.size} trường thông tin cần điền`}
            </div>
            <button
              type="button"
              onClick={() => setStep(2)}
              className="inline-flex items-center justify-center gap-2 px-4 py-2 text-sm font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-lg transition-colors whitespace-nowrap"
            >
              Tiếp tục: Tải giấy tờ hồ sơ
              <ArrowRight className="w-4 h-4" />
            </button>
          </div>
        </div>
      )}

      {/* ===================================================================== */}
      {/* STEP 2: TẢI HỒ SƠ & AI PHÂN TÍCH                                     */}
      {/* ===================================================================== */}
      {step === 2 && (
        <div className="space-y-5">
          <div
            onDragOver={(e) => {
              e.preventDefault();
              setDragOverDocs(true);
            }}
            onDragLeave={() => setDragOverDocs(false)}
            onDrop={(e) => {
              e.preventDefault();
              setDragOverDocs(false);
              handleUploadDocs(e.dataTransfer.files);
            }}
            className={`bg-white border-2 border-dashed rounded-xl p-6 transition-colors ${
              dragOverDocs ? 'border-blue-600 bg-blue-50/40' : 'border-slate-300'
            }`}
          >
            <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-5">
              <div className="flex items-start gap-3.5">
                <div className="w-11 h-11 rounded-lg bg-blue-50 border border-blue-100 flex items-center justify-center text-blue-700 shrink-0">
                  <FolderPlus className="w-5 h-5" />
                </div>
                <div>
                  <h2 className="text-base font-semibold text-slate-900">
                    Tải giấy tờ khách hàng cho hồ sơ này
                  </h2>
                  <p className="text-xs text-slate-600 mt-1 leading-relaxed">
                    Hỗ trợ ảnh CCCD (mặt trước, mặt sau), Sổ đỏ / Giấy chứng nhận, hợp đồng chuyển nhượng, văn bản phân chia di sản hoặc file PDF. Có thể chọn nhiều file cùng lúc.
                  </p>
                </div>
              </div>

              <div className="flex flex-wrap items-center gap-2.5 shrink-0">
                <select
                  value={docTypeHint}
                  onChange={(e) => setDocTypeHint(e.target.value as DocumentTypeHint)}
                  className="px-3 py-2 text-xs font-medium text-slate-800 border border-slate-300 rounded-lg bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-600"
                >
                  {DOC_TYPE_OPTIONS.map((opt) => (
                    <option key={opt.value} value={opt.value}>
                      {opt.label}
                    </option>
                  ))}
                </select>

                <input
                  ref={docInputRef}
                  type="file"
                  multiple
                  accept="image/*,.pdf,application/pdf"
                  onChange={(e) => handleUploadDocs(e.target.files)}
                  className="hidden"
                />
                <button
                  type="button"
                  disabled={uploadingDocs}
                  onClick={() => docInputRef.current?.click()}
                  className="inline-flex items-center gap-2 px-4 py-2 text-sm font-semibold text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-60 rounded-lg transition-colors whitespace-nowrap"
                >
                  {uploadingDocs ? (
                    <Loader2 className="w-4 h-4 animate-spin" />
                  ) : (
                    <Upload className="w-4 h-4" />
                  )}
                  Chọn ảnh / PDF
                </button>
              </div>
            </div>
          </div>

          <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
            <div className="px-5 py-3.5 border-b border-slate-200 flex items-center justify-between">
              <span className="text-sm font-semibold text-slate-900 tabular-nums">
                Tài liệu trong hồ sơ ({caseItem.documents.length})
              </span>
              <span className="text-xs text-slate-500">
                Dữ liệu tách biệt hoàn toàn với các khách hàng khác
              </span>
            </div>

            {caseItem.documents.length === 0 ? (
              <div className="p-10 text-center text-sm text-slate-500">
                Chưa có giấy tờ nào được tải lên. Hãy chọn ảnh CCCD, Sổ đỏ hoặc file PDF ở khung trên.
              </div>
            ) : (
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4 p-5">
                {caseItem.documents.map((doc) => {
                  const isPdf =
                    doc.mimeType === 'application/pdf' ||
                    doc.originalName.toLowerCase().endsWith('.pdf');
                  const docUrl = api.getCaseDocumentUrl(caseItem.id, doc.id);
                  const hintLabel =
                    DOC_TYPE_OPTIONS.find((o) => o.value === doc.docTypeHint)?.label ||
                    'Giấy tờ';

                  return (
                    <div
                      key={doc.id}
                      className="border border-slate-200 rounded-lg overflow-hidden bg-white flex flex-col"
                    >
                      <div className="h-40 bg-slate-100 flex items-center justify-center overflow-hidden border-b border-slate-200">
                        {isPdf ? (
                          <div className="flex flex-col items-center gap-1.5 text-slate-600">
                            <FileText className="w-8 h-8 text-blue-700" />
                            <span className="text-xs font-medium">Tài liệu PDF</span>
                          </div>
                        ) : (
                          <img
                            src={docUrl}
                            alt={doc.originalName}
                            referrerPolicy="no-referrer"
                            className="w-full h-full object-cover"
                          />
                        )}
                      </div>
                      <div className="p-3 flex items-start justify-between gap-2">
                        <div className="min-w-0">
                          <p className="text-xs font-semibold text-slate-900 truncate">
                            {doc.originalName}
                          </p>
                          <p className="text-[11px] text-slate-500 mt-0.5 tabular-nums">
                            {hintLabel} · {formatBytes(doc.sizeBytes)}
                          </p>
                        </div>
                        <button
                          type="button"
                          onClick={() => handleDeleteDoc(doc.id)}
                          className="p-1.5 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded transition-colors shrink-0"
                          title="Xóa tài liệu"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}

            <div className="px-5 py-4 border-t border-slate-200 bg-slate-50 flex flex-wrap items-center justify-between gap-3">
              <button
                type="button"
                onClick={() => setStep(1)}
                className="inline-flex items-center gap-1.5 px-3.5 py-2 text-sm font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-lg transition-colors"
              >
                <ArrowLeft className="w-4 h-4" />
                Quay lại chọn mẫu
              </button>

              <div className="flex flex-wrap items-center gap-2.5">
                {Object.keys(caseItem.extractedFields).length > 0 && (
                  <button
                    type="button"
                    onClick={() => setStep(3)}
                    className="px-3.5 py-2 text-sm font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-lg transition-colors"
                  >
                    Xem dữ liệu đã đọc
                  </button>
                )}
                <button
                  type="button"
                  disabled={analyzing || caseItem.documents.length === 0}
                  onClick={handleAnalyzeDocuments}
                  className="inline-flex items-center gap-2 px-5 py-2.5 text-sm font-semibold text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-60 rounded-lg transition-colors whitespace-nowrap"
                >
                  {analyzing ? (
                    <>
                      <Loader2 className="w-4 h-4 animate-spin" />
                      AI đang đọc &amp; trích xuất...
                    </>
                  ) : (
                    <>
                      <Sparkles className="w-4 h-4" />
                      AI Phân tích hồ sơ
                    </>
                  )}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ===================================================================== */}
      {/* STEP 3: KIỂM TRA & XÁC NHẬN DỮ LIỆU                                  */}
      {/* ===================================================================== */}
      {step === 3 && (
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 items-start">
          {/* Left: Document Reference Viewer */}
          <div className="lg:col-span-5 bg-white border border-slate-200 rounded-xl overflow-hidden lg:sticky lg:top-20">
            <div className="px-4 py-3 border-b border-slate-200 bg-slate-50 flex items-center justify-between gap-2">
              <div className="flex items-center gap-2 text-xs font-semibold text-slate-900">
                <ImageIcon className="w-4 h-4 text-blue-700" />
                Tài liệu gốc đối chiếu ({caseItem.documents.length})
              </div>
              {caseItem.documents.length > 1 && (
                <select
                  value={activeDoc?.id || ''}
                  onChange={(e) => setActiveDocId(e.target.value)}
                  className="text-xs border border-slate-300 rounded-md px-2 py-1 bg-white max-w-[200px] truncate"
                >
                  {caseItem.documents.map((d, i) => (
                    <option key={d.id} value={d.id}>
                      #{i + 1}: {d.originalName}
                    </option>
                  ))}
                </select>
              )}
            </div>

            {!activeDoc ? (
              <div className="p-8 text-center text-xs text-slate-500">
                Chưa có tài liệu nào. Bạn vẫn có thể nhập liệu thủ công bên cạnh hoặc quay lại Bước 2 để tải ảnh/PDF.
              </div>
            ) : (
              <div className="p-3 bg-slate-100">
                {activeDoc.mimeType === 'application/pdf' ||
                activeDoc.originalName.toLowerCase().endsWith('.pdf') ? (
                  <iframe
                    src={api.getCaseDocumentUrl(caseItem.id, activeDoc.id)}
                    title={activeDoc.originalName}
                    className="w-full h-[460px] rounded border border-slate-200 bg-white"
                  />
                ) : (
                  <img
                    src={api.getCaseDocumentUrl(caseItem.id, activeDoc.id)}
                    alt={activeDoc.originalName}
                    referrerPolicy="no-referrer"
                    className="w-full max-h-[520px] object-contain rounded border border-slate-200 bg-white"
                  />
                )}
              </div>
            )}

            {caseItem.aiSummary && (
              <div className="px-4 py-3 border-t border-slate-200 bg-blue-50/40 text-xs text-slate-700 leading-relaxed">
                <span className="font-semibold text-blue-900">Tóm tắt AI: </span>
                {caseItem.aiSummary}
              </div>
            )}
          </div>

          {/* Right: Structured Verification Form */}
          <div className="lg:col-span-7 bg-white border border-slate-200 rounded-xl overflow-hidden">
            <div className="px-5 py-4 border-b border-slate-200 space-y-3">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                <div>
                  <h2 className="text-base font-semibold text-slate-900">
                    Kiểm tra &amp; Xác nhận thông tin trích xuất
                  </h2>
                  <p className="text-xs text-slate-500 mt-0.5 tabular-nums">
                    {requiredTemplateFieldKeys.size > 0
                      ? `Đã có dữ liệu ${filledRequiredCount}/${requiredTemplateFieldKeys.size} trường dùng trong các mẫu Word đã chọn.`
                      : 'Chỉnh sửa trực tiếp vào từng ô nếu AI đọc thiếu dấu hoặc cần bổ sung.'}
                  </p>
                </div>
                <button
                  type="button"
                  onClick={() => setStep(2)}
                  className="text-xs font-semibold text-blue-700 hover:underline self-start sm:self-center shrink-0"
                >
                  + Tải thêm giấy tờ / Phân tích lại
                </button>
              </div>

              <div className="flex flex-col sm:flex-row gap-2.5">
                <div className="relative flex-1">
                  <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                  <input
                    type="text"
                    value={fieldSearch}
                    onChange={(e) => setFieldSearch(e.target.value)}
                    placeholder="Tìm trường thông tin (vd: CCCD, thửa đất, giá)..."
                    className="w-full pl-9 pr-3 py-1.5 text-sm border border-slate-200 rounded-lg bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-600"
                  />
                </div>
                {requiredTemplateFieldKeys.size > 0 && (
                  <label className="inline-flex items-center gap-2 px-3 py-1.5 text-xs font-medium text-slate-700 border border-slate-200 rounded-lg bg-slate-50 cursor-pointer shrink-0">
                    <input
                      type="checkbox"
                      checked={onlyTemplateFields}
                      onChange={(e) => setOnlyTemplateFields(e.target.checked)}
                      className="rounded border-slate-300 text-blue-700 focus:ring-blue-600"
                    />
                    Chỉ hiện trường có trong mẫu ({requiredTemplateFieldKeys.size})
                  </label>
                )}
              </div>
            </div>

            <div className="p-5 space-y-6">
              {(['party_a', 'party_b', 'property', 'general', 'custom'] as FieldGroup[]).map(
                (groupKey) => {
                  const items = groupedFields[groupKey];
                  if (items.length === 0) return null;

                  return (
                    <section key={groupKey} className="space-y-3">
                      <h3 className="text-xs font-bold text-blue-800 border-b border-slate-200 pb-1.5">
                        {FIELD_GROUP_LABELS[groupKey]}
                      </h3>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                        {items.map((field) => {
                          const usedInTpl = requiredTemplateFieldKeys.has(field.key);
                          return (
                            <div key={field.key} className="space-y-1">
                              <div className="flex items-center justify-between gap-2">
                                <label className="text-xs font-semibold text-slate-800 truncate">
                                  {field.label}
                                </label>
                                {usedInTpl && (
                                  <span className="text-[11px] font-medium text-blue-700 shrink-0">
                                    Có trong mẫu
                                  </span>
                                )}
                              </div>
                              <input
                                type="text"
                                value={field.value}
                                onChange={(e) => handleFieldValueChange(field, e.target.value)}
                                placeholder={`Nhập ${field.label.toLowerCase()}...`}
                                className={`w-full px-3 py-2 text-sm border rounded-lg tabular-nums focus:outline-none focus:ring-2 focus:ring-blue-600 ${
                                  field.value.trim()
                                    ? 'border-slate-300 bg-white text-slate-900'
                                    : usedInTpl
                                      ? 'border-amber-300 bg-amber-50/40 text-slate-900'
                                      : 'border-slate-200 bg-slate-50 text-slate-900'
                                }`}
                              />
                              <div className="flex items-center justify-between text-[11px] text-slate-400">
                                <span className="font-mono">{`{{${field.key}}}`}</span>
                                {field.sourceNote && (
                                  <span className="truncate max-w-[170px] text-slate-500">
                                    Nguồn: {field.sourceNote}
                                  </span>
                                )}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </section>
                  );
                }
              )}

              {/* Replacement & Optional Extra Rules */}
              <div className="pt-4 border-t border-slate-200 space-y-3">
                <div>
                  <p className="text-xs font-bold text-blue-800">
                    Tùy chọn điền mẫu Word (Giữ nguyên form gốc của mẫu)
                  </p>
                  <p className="text-[11px] text-slate-500 mt-0.5">
                    Hệ thống chỉ thay thế các giá trị thông tin cần thiết tại vị trí có sẵn trên mẫu, giữ nguyên cấu trúc bảng biểu, câu chữ và dấu chấm trống của mẫu mặc định.
                  </p>
                </div>

                <label className="flex items-start gap-2.5 text-xs text-slate-700 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={fillOptions.eraseUnfilledOldText}
                    onChange={(e) =>
                      setFillOptions((prev) => ({
                        ...prev,
                        eraseUnfilledOldText: e.target.checked,
                      }))
                    }
                    className="mt-0.5 rounded border-slate-300 text-blue-700 focus:ring-blue-600"
                  />
                  <span>
                    <strong className="text-slate-900">Xóa sạch thông tin khách cũ nếu hồ sơ mới để trống</strong>
                    <span className="block text-slate-500 mt-0.5">
                      Khi dùng lại file Word cũ của khách hàng trước (vd: có tên vợ/chồng cũ mà khách mới chỉ có 1 người), tự động xóa giá trị cũ để không bị sót. Không xóa dấu chấm (……) hay thay đổi form.
                    </span>
                  </span>
                </label>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-3 pt-1">
                  <div className="border border-slate-200 rounded-lg p-3 space-y-2 bg-slate-50/70">
                    <label className="flex items-center gap-2 text-xs font-semibold text-slate-800 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={fillOptions.autoInsertSpousePartyB}
                        onChange={(e) =>
                          setFillOptions((prev) => ({
                            ...prev,
                            autoInsertSpousePartyB: e.target.checked,
                          }))
                        }
                        className="rounded border-slate-300 text-blue-700 focus:ring-blue-600"
                      />
                      Tự động chèn thêm Vợ/Chồng Bên nhận (Bên B)
                    </label>
                    <p className="text-[11px] text-slate-500">
                      Khi mẫu Word chỉ có chỗ cho 1 người nhưng hồ sơ có thêm vợ/chồng, tự động thêm dòng thông tin vợ/chồng ngay bên dưới (giữ nguyên font chữ mẫu).
                    </p>
                    {fillOptions.autoInsertSpousePartyB && (
                      <textarea
                        rows={2}
                        value={fillOptions.customSpouseTextPartyB ?? ''}
                        onChange={(e) =>
                          setFillOptions((prev) => ({
                            ...prev,
                            customSpouseTextPartyB: e.target.value,
                          }))
                        }
                        placeholder="Để trống để tự ghép từ các ô Vợ/Chồng Bên B ở trên, hoặc tự nhập câu chữ tùy ý..."
                        className="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-blue-600"
                      />
                    )}
                  </div>

                  <div className="border border-slate-200 rounded-lg p-3 space-y-2 bg-slate-50/70">
                    <label className="flex items-center gap-2 text-xs font-semibold text-slate-800 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={fillOptions.autoInsertSpousePartyA}
                        onChange={(e) =>
                          setFillOptions((prev) => ({
                            ...prev,
                            autoInsertSpousePartyA: e.target.checked,
                          }))
                        }
                        className="rounded border-slate-300 text-blue-700 focus:ring-blue-600"
                      />
                      Tự động chèn thêm Vợ/Chồng Bên bán (Bên A)
                    </label>
                    <p className="text-[11px] text-slate-500">
                      Tự động thêm dòng thông tin vợ/chồng đồng sở hữu Bên bán ngay bên dưới đoạn Bên A nếu mẫu chưa có sẵn.
                    </p>
                    {fillOptions.autoInsertSpousePartyA && (
                      <textarea
                        rows={2}
                        value={fillOptions.customSpouseTextPartyA ?? ''}
                        onChange={(e) =>
                          setFillOptions((prev) => ({
                            ...prev,
                            customSpouseTextPartyA: e.target.value,
                          }))
                        }
                        placeholder="Để trống để tự ghép từ các ô Vợ/Chồng Bên A ở trên, hoặc tự nhập câu chữ tùy ý..."
                        className="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-blue-600"
                      />
                    )}
                  </div>
                </div>
              </div>

              {/* Add custom field */}
              <div className="pt-3 border-t border-slate-200 space-y-2.5">
                <p className="text-xs font-semibold text-slate-800">
                  Thêm trường dữ liệu bổ sung (nếu mẫu Word có thẻ riêng)
                </p>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                  <input
                    type="text"
                    value={customFieldKey}
                    onChange={(e) => setCustomFieldKey(e.target.value)}
                    placeholder="Mã thẻ (vd: so_dien_thoai)"
                    className="px-3 py-1.5 text-xs border border-slate-300 rounded-md"
                  />
                  <input
                    type="text"
                    value={customFieldLabel}
                    onChange={(e) => setCustomFieldLabel(e.target.value)}
                    placeholder="Tên hiển thị (vd: Số điện thoại)"
                    className="px-3 py-1.5 text-xs border border-slate-300 rounded-md"
                  />
                  <div className="flex gap-2">
                    <input
                      type="text"
                      value={customFieldValue}
                      onChange={(e) => setCustomFieldValue(e.target.value)}
                      placeholder="Giá trị..."
                      className="flex-1 px-3 py-1.5 text-xs border border-slate-300 rounded-md"
                    />
                    <button
                      type="button"
                      onClick={handleAddCustomField}
                      className="px-3 py-1.5 text-xs font-semibold text-blue-700 bg-blue-50 hover:bg-blue-100 rounded-md shrink-0 inline-flex items-center gap-1"
                    >
                      <Plus className="w-3.5 h-3.5" />
                      Thêm
                    </button>
                  </div>
                </div>
              </div>
            </div>

            <div className="px-5 py-4 border-t border-slate-200 bg-slate-50 flex flex-wrap items-center justify-between gap-3">
              <button
                type="button"
                onClick={() => setStep(2)}
                className="inline-flex items-center gap-1.5 px-3.5 py-2 text-sm font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-lg transition-colors"
              >
                <ArrowLeft className="w-4 h-4" />
                Quay lại
              </button>

              <button
                type="button"
                disabled={savingAndFilling}
                onClick={() => handleConfirmAndFill(true)}
                className="inline-flex items-center gap-2 px-5 py-2.5 text-sm font-semibold text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-60 rounded-lg transition-colors whitespace-nowrap"
              >
                {savingAndFilling ? (
                  <>
                    <Loader2 className="w-4 h-4 animate-spin" />
                    Đang điền vào mẫu Word...
                  </>
                ) : (
                  <>
                    <FileCheck2 className="w-4 h-4" />
                    Xác nhận dữ liệu &amp; Điền mẫu Word
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ===================================================================== */}
      {/* STEP 4: XEM TRƯỚC & XUẤT FILE WORD                                   */}
      {/* ===================================================================== */}
      {step === 4 && (
        <div className="space-y-5">
          {/* Output files bar & batch download */}
          <div className="bg-white border border-slate-200 rounded-xl p-5 space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div>
                <h2 className="text-base font-semibold text-slate-900">
                  Kết quả file Word đã điền thông tin
                </h2>
                <p className="text-xs text-slate-500 mt-0.5">
                  Kiểm tra bản xem trước bên dưới. Bạn có thể chỉnh sửa nhanh thông tin hoặc tải từng file / tải toàn bộ dưới dạng ZIP.
                </p>
              </div>
              <div className="flex flex-wrap items-center gap-2 shrink-0">
                <button
                  type="button"
                  onClick={() => setShowQuickEdit((v) => !v)}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-lg transition-colors whitespace-nowrap"
                >
                  <Pencil className="w-3.5 h-3.5" />
                  {showQuickEdit ? 'Ẩn chỉnh sửa nhanh' : 'Chỉnh sửa nhanh dữ liệu'}
                </button>
                {caseItem.generatedOutputs.length > 0 && (
                  <a
                    href={api.getOutputsZipUrl(caseItem.id)}
                    className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-lg transition-colors whitespace-nowrap"
                  >
                    <FileArchive className="w-4 h-4" />
                    Tải tất cả ({caseItem.generatedOutputs.length} file .zip)
                  </a>
                )}
              </div>
            </div>

            {caseItem.generatedOutputs.length === 0 ? (
              <div className="p-8 text-center border border-dashed border-slate-300 rounded-lg space-y-3">
                <p className="text-sm text-slate-600">
                  Chưa có file Word kết quả nào được tạo cho hồ sơ này.
                </p>
                <button
                  type="button"
                  onClick={() => handleConfirmAndFill(true)}
                  disabled={savingAndFilling}
                  className="inline-flex items-center gap-2 px-4 py-2 text-sm font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-lg"
                >
                  {savingAndFilling ? (
                    <Loader2 className="w-4 h-4 animate-spin" />
                  ) : (
                    <FileSpreadsheet className="w-4 h-4" />
                  )}
                  Tiến hành điền mẫu Word ngay
                </button>
              </div>
            ) : (
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                {caseItem.generatedOutputs.map((out) => {
                  const isActive = activeOutput?.outputFileName === out.outputFileName;
                  return (
                    <div
                      key={out.outputFileName}
                      className={`border rounded-lg p-3.5 flex flex-col justify-between gap-3 transition-colors ${
                        isActive
                          ? 'border-blue-600 bg-blue-50/40'
                          : 'border-slate-200 bg-white hover:border-slate-300'
                      }`}
                    >
                      <div className="min-w-0">
                        <p className="text-sm font-semibold text-slate-900 truncate">
                          {out.templateName}
                        </p>
                        <p className="text-xs text-slate-500 mt-1 tabular-nums">
                          Đã điền {out.filledCount} vị trí
                          {out.missingKeys.length > 0 && (
                            <span className="text-amber-700">
                              {' '}
                              · Còn trống {out.missingKeys.length} trường
                            </span>
                          )}
                        </p>
                      </div>
                      <div className="flex items-center gap-2">
                        <button
                          type="button"
                          onClick={() => setActiveOutputFile(out.outputFileName)}
                          className={`flex-1 inline-flex items-center justify-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-md transition-colors ${
                            isActive
                              ? 'bg-blue-700 text-white'
                              : 'bg-slate-100 text-slate-800 hover:bg-slate-200'
                          }`}
                        >
                          <Eye className="w-3.5 h-3.5" />
                          Xem trước
                        </button>
                        <a
                          href={api.getOutputFileUrl(caseItem.id, out.outputFileName)}
                          download={out.outputFileName}
                          className="inline-flex items-center justify-center gap-1 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-md transition-colors"
                        >
                          <Download className="w-3.5 h-3.5" />
                          Tải .docx
                        </a>
                      </div>
                      {out.filledCount === 0 && (
                        <div className="pt-2 border-t border-amber-200 flex flex-wrap items-center gap-2">
                          <button
                            type="button"
                            disabled={savingAndFilling}
                            onClick={() => handleAutoDetectTemplateAndRefill(out.templateId, true)}
                            className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-semibold text-blue-700 bg-blue-50 hover:bg-blue-100 rounded"
                          >
                            <Sparkles className="w-3 h-3" />
                            AI nhận diện tên/thông tin cũ &amp; Điền lại
                          </button>
                          <button
                            type="button"
                            onClick={() => setMapperTplId(out.templateId)}
                            className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded"
                          >
                            <MapPin className="w-3 h-3" />
                            Cấu hình thủ công
                          </button>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}

            {activeOutput && activeOutput.filledCount === 0 && (
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 px-4 py-3 text-xs text-amber-900 bg-amber-50 border border-amber-200 rounded-lg">
                <span>
                  Mẫu <strong>{activeOutput.templateName}</strong> chưa khớp được vị trí nào (có thể đây là file Word cũ có sẵn tên người cũ thay vì thẻ {'{{ho_ten}}'}). Bấm nút bên cạnh để AI tự tìm tên/CCCD/địa chỉ cũ và thay thế tự động.
                </span>
                <button
                  type="button"
                  disabled={savingAndFilling}
                  onClick={() => handleAutoDetectTemplateAndRefill(activeOutput.templateId, true)}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-md shrink-0 whitespace-nowrap"
                >
                  {savingAndFilling ? (
                    <Loader2 className="w-3.5 h-3.5 animate-spin" />
                  ) : (
                    <Sparkles className="w-3.5 h-3.5" />
                  )}
                  AI nhận diện &amp; Điền lại ngay
                </button>
              </div>
            )}
          </div>

          {/* Quick inline edit drawer in Step 4 */}
          {showQuickEdit && (
            <div className="bg-white border border-slate-200 rounded-xl p-5 space-y-4">
              <div className="flex flex-wrap items-center justify-between gap-3">
                <h3 className="text-sm font-semibold text-slate-900">
                  Chỉnh sửa nhanh thông tin &amp; cập nhật lại toàn bộ file Word
                </h3>
                <button
                  type="button"
                  disabled={savingAndFilling}
                  onClick={() => handleConfirmAndFill(false)}
                  className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-60 rounded-lg"
                >
                  {savingAndFilling ? (
                    <Loader2 className="w-3.5 h-3.5 animate-spin" />
                  ) : (
                    <Check className="w-3.5 h-3.5" />
                  )}
                  Cập nhật &amp; Điền lại
                </button>
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3 max-h-80 overflow-y-auto pr-1">
                {mergedDisplayFields
                  .filter(
                    (f) =>
                      requiredTemplateFieldKeys.size === 0 ||
                      requiredTemplateFieldKeys.has(f.key) ||
                      f.value.trim() !== ''
                  )
                  .map((field) => (
                    <div key={field.key} className="space-y-1">
                      <label className="block text-xs font-medium text-slate-700 truncate">
                        {field.label}
                      </label>
                      <input
                        type="text"
                        value={field.value}
                        onChange={(e) => handleFieldValueChange(field, e.target.value)}
                        className="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded-md tabular-nums focus:outline-none focus:ring-2 focus:ring-blue-600"
                      />
                    </div>
                  ))}
              </div>
            </div>
          )}

          {/* Live DOCX Preview */}
          {activeOutput && (
            <DocxPreviewPane
              fileUrl={api.getOutputFileUrl(caseItem.id, activeOutput.outputFileName)}
              fileName={activeOutput.outputFileName}
              refreshToken={previewRefreshToken}
            />
          )}

          <div className="flex items-center justify-between">
            <button
              type="button"
              onClick={() => setStep(3)}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 text-sm font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-lg transition-colors"
            >
              <ArrowLeft className="w-4 h-4" />
              Quay lại kiểm tra dữ liệu
            </button>
          </div>
        </div>
      )}

      {mapperTplId && (
        <TemplateMapperModal
          templateId={mapperTplId}
          onClose={() => setMapperTplId(null)}
          onSaved={(updated) => onTemplatesUpdated(updated)}
        />
      )}
    </div>
  );
}
