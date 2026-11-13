import type {
  CaseFillOptions,
  CustomerCase,
  DocxParagraphInspect,
  DocumentTypeHint,
  ExtractedFieldValue,
  GeneratedOutput,
  TemplateFieldMapping,
  WordTemplate,
} from './shared/types';

async function requestJson<T>(url: string, init?: RequestInit): Promise<T> {
  const res = await fetch(url, init);
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error((data as { error?: string }).error || `Lỗi kết nối (${res.status})`);
  }
  return data as T;
}

export const api = {
  // Templates
  listTemplates: () => requestJson<{ templates: WordTemplate[] }>('/api/templates'),

  uploadTemplates: (files: File[], category: string) => {
    const fd = new FormData();
    fd.append('category', category);
    for (const f of files) {
      fd.append('files', f);
    }
    return requestJson<{ created: WordTemplate[]; templates: WordTemplate[] }>(
      '/api/templates/upload',
      { method: 'POST', body: fd }
    );
  },

  reuploadTemplateFile: (templateId: string, file: File) => {
    const fd = new FormData();
    fd.append('file', file);
    return requestJson<{ template: WordTemplate; templates: WordTemplate[] }>(
      `/api/templates/${encodeURIComponent(templateId)}/reupload`,
      { method: 'POST', body: fd }
    );
  },

  restoreSampleTemplates: () =>
    requestJson<{ templates: WordTemplate[] }>('/api/templates/restore-samples', {
      method: 'POST',
    }),

  inspectTemplate: (templateId: string) =>
    requestJson<{ template: WordTemplate; paragraphs: DocxParagraphInspect[] }>(
      `/api/templates/${encodeURIComponent(templateId)}/inspect`
    ),

  autoDetectTemplate: (templateId: string) =>
    requestJson<{ template: WordTemplate; templates: WordTemplate[]; detectedCount: number }>(
      `/api/templates/${encodeURIComponent(templateId)}/auto-detect`,
      { method: 'POST' }
    ),

  updateTemplate: (
    templateId: string,
    payload: { name?: string; category?: string; mappings?: TemplateFieldMapping[] }
  ) =>
    requestJson<{ template: WordTemplate; templates: WordTemplate[] }>(
      `/api/templates/${encodeURIComponent(templateId)}`,
      {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      }
    ),

  deleteTemplate: (templateId: string) =>
    requestJson<{ templates: WordTemplate[] }>(
      `/api/templates/${encodeURIComponent(templateId)}`,
      { method: 'DELETE' }
    ),

  getTemplateDownloadUrl: (templateId: string) =>
    `/api/templates/${encodeURIComponent(templateId)}/download`,

  // Cases (Hồ sơ khách hàng)
  listCases: () => requestJson<{ cases: CustomerCase[] }>('/api/cases'),

  createCase: (payload?: {
    title?: string;
    customerName?: string;
    selectedTemplateIds?: string[];
  }) =>
    requestJson<{ caseItem: CustomerCase; cases: CustomerCase[] }>('/api/cases', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload ?? {}),
    }),

  updateCase: (
    caseId: string,
    patch: {
      title?: string;
      customerName?: string;
      selectedTemplateIds?: string[];
      extractedFields?: Record<string, ExtractedFieldValue>;
      fillOptions?: CaseFillOptions;
      isVerified?: boolean;
      status?: CustomerCase['status'];
    }
  ) =>
    requestJson<{ caseItem: CustomerCase; cases: CustomerCase[] }>(
      `/api/cases/${encodeURIComponent(caseId)}`,
      {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(patch),
      }
    ),

  deleteCase: (caseId: string) =>
    requestJson<{ cases: CustomerCase[] }>(`/api/cases/${encodeURIComponent(caseId)}`, {
      method: 'DELETE',
    }),

  uploadCaseDocuments: (caseId: string, files: File[], docTypeHint: DocumentTypeHint) => {
    const fd = new FormData();
    fd.append('docTypeHint', docTypeHint);
    for (const f of files) {
      fd.append('files', f);
    }
    return requestJson<{ caseItem: CustomerCase }>(
      `/api/cases/${encodeURIComponent(caseId)}/documents`,
      { method: 'POST', body: fd }
    );
  },

  deleteCaseDocument: (caseId: string, docId: string) =>
    requestJson<{ caseItem: CustomerCase }>(
      `/api/cases/${encodeURIComponent(caseId)}/documents/${encodeURIComponent(docId)}`,
      { method: 'DELETE' }
    ),

  getCaseDocumentUrl: (caseId: string, docId: string) =>
    `/api/cases/${encodeURIComponent(caseId)}/documents/${encodeURIComponent(docId)}/file`,

  analyzeCase: (caseId: string) =>
    requestJson<{ caseItem: CustomerCase }>(
      `/api/cases/${encodeURIComponent(caseId)}/analyze`,
      { method: 'POST' }
    ),

  fillCaseTemplates: (caseId: string) =>
    requestJson<{ caseItem: CustomerCase; outputs: GeneratedOutput[]; templates?: WordTemplate[] }>(
      `/api/cases/${encodeURIComponent(caseId)}/fill`,
      { method: 'POST' }
    ),

  getOutputFileUrl: (caseId: string, outputFileName: string) =>
    `/api/cases/${encodeURIComponent(caseId)}/outputs/${encodeURIComponent(outputFileName)}`,

  getOutputsZipUrl: (caseId: string) =>
    `/api/cases/${encodeURIComponent(caseId)}/outputs-zip`,
};
