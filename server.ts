import 'dotenv/config';
import fs from 'node:fs';
import path from 'node:path';
import express from 'express';
import multer from 'multer';
import JSZip from 'jszip';
import {
  deleteCaseDirectory,
  ensureCaseDirs,
  getCaseDocFilePath,
  getCaseOutputFilePath,
  getTemplateFilePath,
  initStorage,
  readCases,
  readTemplates,
  restoreStarterTemplates,
  writeCases,
  writeTemplates,
} from './server/storage.ts';
import { fillDocx, inspectDocx } from './server/docxEngine.ts';
import {
  detectTemplateMappingsWithAI,
  extractCustomerDocuments,
} from './server/aiExtractor.ts';
import {
  deriveGranularFields,
  isLegacyNumericBracketMapping,
  type CustomerCase,
  type DocumentTypeHint,
  type FieldDefinition,
  type GeneratedOutput,
  type TemplateFieldMapping,
  type UploadedDoc,
  type WordTemplate,
} from './src/shared/types.ts';

const PORT = 3000;

function decodeUploadFilename(raw: string): string {
  try {
    // Multer often interprets UTF-8 filenames as latin1
    const decoded = Buffer.from(raw, 'latin1').toString('utf8');
    if (decoded.includes('\uFFFD')) return raw;
    return decoded;
  } catch {
    return raw;
  }
}

function sanitizeFileBase(name: string): string {
  return name
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .replace(/[^a-zA-Z0-9._-]+/g, '_')
    .replace(/^_+|_+$/g, '') || 'file';
}

function setContentDisposition(res: express.Response, fileName: string) {
  const asciiFallback = sanitizeFileBase(fileName);
  const encoded = encodeURIComponent(fileName);
  res.setHeader(
    'Content-Disposition',
    `attachment; filename="${asciiFallback}"; filename*=UTF-8''${encoded}`
  );
}

async function startServer() {
  initStorage();

  const app = express();
  app.use(express.json({ limit: '25mb' }));

  const uploadMemory = multer({
    storage: multer.memoryStorage(),
    limits: { fileSize: 30 * 1024 * 1024 },
  });

  // ---------------------------------------------------------------------------
  // TEMPLATES API
  // ---------------------------------------------------------------------------

  app.get('/api/templates', (_req, res) => {
    try {
      const templates = readTemplates();
      res.json({ templates });
    } catch (err) {
      res.status(500).json({ error: (err as Error).message });
    }
  });

  app.post('/api/templates/restore-samples', (_req, res) => {
    try {
      const templates = restoreStarterTemplates();
      res.json({ templates });
    } catch (err) {
      res.status(500).json({ error: (err as Error).message });
    }
  });

  app.post('/api/templates/upload', uploadMemory.array('files', 20), async (req, res) => {
    try {
      const files = (req.files as Express.Multer.File[]) ?? [];
      if (files.length === 0) {
        res.status(400).json({ error: 'Vui lòng chọn ít nhất một file Word (.docx).' });
        return;
      }

      const category = typeof req.body.category === 'string' && req.body.category.trim()
        ? req.body.category.trim()
        : 'Mẫu văn bản';

      const templates = readTemplates();
      const created: WordTemplate[] = [];
      const now = new Date().toISOString();

      for (const file of files) {
        const originalName = decodeUploadFilename(file.originalname);
        if (!originalName.toLowerCase().endsWith('.docx')) {
          res.status(400).json({
            error: `File "${originalName}" không đúng định dạng .docx. Vui lòng chuyển sang .docx trước khi tải lên.`,
          });
          return;
        }

        const inspected = inspectDocx(file.buffer);
        let initialMappings = inspected.defaultMappings;

        // If uploaded Word file is an existing old document (no {{tags}}), automatically detect old client fields & blanks via AI
        if (initialMappings.length === 0 && inspected.paragraphs.length > 0) {
          try {
            initialMappings = await detectTemplateMappingsWithAI(inspected.paragraphs);
          } catch (aiErr) {
            console.warn('Auto-detect template mappings skipped:', (aiErr as Error).message);
          }
        }

        const id = `tpl_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
        const diskName = `${id}_${sanitizeFileBase(originalName)}`;
        fs.writeFileSync(getTemplateFilePath(diskName), file.buffer);

        const baseName = originalName.replace(/\.docx$/i, '');
        const newTpl: WordTemplate = {
          id,
          name: baseName,
          originalFileName: originalName,
          category,
          fileNameOnDisk: diskName,
          sizeBytes: file.size,
          placeholdersDetected: inspected.placeholders,
          mappings: initialMappings,
          createdAt: now,
          updatedAt: now,
        };
        templates.unshift(newTpl);
        created.push(newTpl);
      }

      writeTemplates(templates);
      res.json({ created, templates: readTemplates() });
    } catch (err) {
      res.status(500).json({ error: `Không thể đọc file Word: ${(err as Error).message}` });
    }
  });

  app.post('/api/templates/:id/reupload', uploadMemory.single('file'), async (req, res) => {
    try {
      const file = req.file as Express.Multer.File | undefined;
      if (!file) {
        res.status(400).json({ error: 'Vui lòng chọn file Word (.docx).' });
        return;
      }
      const originalName = decodeUploadFilename(file.originalname);
      if (!originalName.toLowerCase().endsWith('.docx')) {
        res.status(400).json({ error: 'File không đúng định dạng .docx.' });
        return;
      }

      const templates = readTemplates();
      const idx = templates.findIndex((t) => t.id === req.params.id);
      if (idx === -1) {
        res.status(404).json({ error: 'Không tìm thấy mẫu Word.' });
        return;
      }

      const tpl = templates[idx];
      const inspected = inspectDocx(file.buffer);
      let mappings = tpl.mappings;
      if (mappings.length === 0 || mappings.every(isLegacyNumericBracketMapping)) {
        mappings = inspected.defaultMappings;
        if (mappings.length === 0 && inspected.paragraphs.length > 0) {
          try {
            mappings = await detectTemplateMappingsWithAI(inspected.paragraphs);
          } catch (aiErr) {
            console.warn('Auto-detect on reupload skipped:', (aiErr as Error).message);
          }
        }
      }

      const diskName = tpl.fileNameOnDisk || `${tpl.id}_${sanitizeFileBase(originalName)}`;
      fs.writeFileSync(getTemplateFilePath(diskName), file.buffer);

      templates[idx] = {
        ...tpl,
        originalFileName: originalName,
        fileNameOnDisk: diskName,
        sizeBytes: file.size,
        placeholdersDetected: inspected.placeholders,
        mappings,
        fileMissing: false,
        updatedAt: new Date().toISOString(),
      };
      writeTemplates(templates);
      const refreshed = readTemplates();
      res.json({
        template: refreshed.find((t) => t.id === tpl.id) ?? templates[idx],
        templates: refreshed,
      });
    } catch (err) {
      res.status(500).json({ error: `Không thể tải lại file Word: ${(err as Error).message}` });
    }
  });

  app.post('/api/templates/:id/auto-detect', async (req, res) => {
    try {
      const templates = readTemplates();
      const idx = templates.findIndex((t) => t.id === req.params.id);
      if (idx === -1) {
        res.status(404).json({ error: 'Không tìm thấy mẫu Word.' });
        return;
      }
      const tpl = templates[idx];
      const absPath = getTemplateFilePath(tpl.fileNameOnDisk);
      if (!fs.existsSync(absPath)) {
        res.status(404).json({ error: 'File mẫu gốc không còn trên ổ đĩa.' });
        return;
      }

      const buf = fs.readFileSync(absPath);
      const inspected = inspectDocx(buf);
      const aiMappings = await detectTemplateMappingsWithAI(inspected.paragraphs);

      // Merge placeholder mappings + newly detected AI mappings without duplicates
      const merged = inspected.defaultMappings.filter((m) => !isLegacyNumericBracketMapping(m));
      for (const m of aiMappings) {
        const dup = merged.some(
          (ex) =>
            ex.mode === m.mode &&
            ex.fieldKey === m.fieldKey &&
            (ex.targetText ?? '') === (m.targetText ?? '') &&
            (ex.paragraphIndex ?? -1) === (m.paragraphIndex ?? -1)
        );
        if (!dup) merged.push(m);
      }

      const updated: WordTemplate = {
        ...tpl,
        placeholdersDetected: inspected.placeholders,
        mappings: merged,
        updatedAt: new Date().toISOString(),
      };
      templates[idx] = updated;
      writeTemplates(templates);

      res.json({ template: updated, templates, detectedCount: merged.length });
    } catch (err) {
      res.status(500).json({ error: `Lỗi AI nhận diện mẫu: ${(err as Error).message}` });
    }
  });

  app.get('/api/templates/:id/inspect', (req, res) => {
    try {
      const templates = readTemplates();
      const tpl = templates.find((t) => t.id === req.params.id);
      if (!tpl) {
        res.status(404).json({ error: 'Không tìm thấy mẫu Word.' });
        return;
      }
      const absPath = getTemplateFilePath(tpl.fileNameOnDisk);
      if (!fs.existsSync(absPath)) {
        res.status(404).json({ error: 'File mẫu gốc không còn trên ổ đĩa.' });
        return;
      }
      const buf = fs.readFileSync(absPath);
      const inspected = inspectDocx(buf);
      res.json({
        template: tpl,
        paragraphs: inspected.paragraphs,
      });
    } catch (err) {
      res.status(500).json({ error: (err as Error).message });
    }
  });

  app.put('/api/templates/:id', (req, res) => {
    try {
      const templates = readTemplates();
      const idx = templates.findIndex((t) => t.id === req.params.id);
      if (idx === -1) {
        res.status(404).json({ error: 'Không tìm thấy mẫu Word.' });
        return;
      }

      const { name, category, mappings } = req.body as {
        name?: string;
        category?: string;
        mappings?: TemplateFieldMapping[];
      };

      const updated: WordTemplate = {
        ...templates[idx],
        name: typeof name === 'string' && name.trim() ? name.trim() : templates[idx].name,
        category: typeof category === 'string' && category.trim() ? category.trim() : templates[idx].category,
        mappings: Array.isArray(mappings) ? mappings : templates[idx].mappings,
        updatedAt: new Date().toISOString(),
      };

      templates[idx] = updated;
      writeTemplates(templates);
      res.json({ template: updated, templates });
    } catch (err) {
      res.status(500).json({ error: (err as Error).message });
    }
  });

  app.get('/api/templates/:id/download', (req, res) => {
    try {
      const templates = readTemplates();
      const tpl = templates.find((t) => t.id === req.params.id);
      if (!tpl) {
        res.status(404).json({ error: 'Không tìm thấy mẫu Word.' });
        return;
      }
      const absPath = getTemplateFilePath(tpl.fileNameOnDisk);
      if (!fs.existsSync(absPath)) {
        res.status(404).json({ error: 'File mẫu không tồn tại.' });
        return;
      }
      setContentDisposition(res, tpl.originalFileName);
      res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      );
      res.send(fs.readFileSync(absPath));
    } catch (err) {
      res.status(500).json({ error: (err as Error).message });
    }
  });

  app.delete('/api/templates/:id', (req, res) => {
    try {
      const templates = readTemplates();
      const tpl = templates.find((t) => t.id === req.params.id);
      if (!tpl) {
        res.status(404).json({ error: 'Không tìm thấy mẫu Word.' });
        return;
      }
      const absPath = getTemplateFilePath(tpl.fileNameOnDisk);
      if (fs.existsSync(absPath)) fs.unlinkSync(absPath);

      const remaining = templates.filter((t) => t.id !== req.params.id);
      writeTemplates(remaining);
      res.json({ templates: remaining });
    } catch (err) {
      res.status(500).json({ error: (err as Error).message });
    }
  });

  // ---------------------------------------------------------------------------
  // CUSTOMER CASES API (STRICTLY ISOLATED PER CASE)
  // ---------------------------------------------------------------------------

  app.get('/api/cases', (_req, res) => {
    try {
      const cases = readCases();
      res.json({ cases });
    } catch (err) {
      res.status(500).json({ error: (err as Error).message });
    }
  });

  app.post('/api/cases', (req, res) => {
    try {
      const cases = readCases();
      const body = req.body as { title?: string; customerName?: string; selectedTemplateIds?: string[] };
      const now = new Date().toISOString();
      const id = `case_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
      ensureCaseDirs(id);

      const newCase: CustomerCase = {
        id,
        title: (body.title || '').trim() || `Hồ sơ khách hàng #${cases.length + 1}`,
        customerName: (body.customerName || '').trim() || 'Khách hàng mới',
        status: 'draft',
        selectedTemplateIds: Array.isArray(body.selectedTemplateIds) ? body.selectedTemplateIds : [],
        documents: [],
        extractedFields: {},
        isVerified: false,
        generatedOutputs: [],
        createdAt: now,
        updatedAt: now,
      };

      cases.unshift(newCase);
      writeCases(cases);
      res.json({ caseItem: newCase, cases });
    } catch (err) {
      res.status(500).json({ error: (err as Error).message });
    }
  });

  app.put('/api/cases/:id', (req, res) => {
    try {
      const cases = readCases();
      const idx = cases.findIndex((c) => c.id === req.params.id);
      if (idx === -1) {
        res.status(404).json({ error: 'Không tìm thấy hồ sơ khách hàng.' });
        return;
      }

      const patch = req.body as Partial<CustomerCase>;
      const prev = cases[idx];
      const updated: CustomerCase = {
        ...prev,
        title: typeof patch.title === 'string' ? patch.title : prev.title,
        customerName: typeof patch.customerName === 'string' ? patch.customerName : prev.customerName,
        status: patch.status ?? prev.status,
        selectedTemplateIds: Array.isArray(patch.selectedTemplateIds)
          ? patch.selectedTemplateIds
          : prev.selectedTemplateIds,
        extractedFields:
          patch.extractedFields && typeof patch.extractedFields === 'object'
            ? patch.extractedFields
            : prev.extractedFields,
        fillOptions:
          patch.fillOptions && typeof patch.fillOptions === 'object'
            ? patch.fillOptions
            : prev.fillOptions,
        isVerified: typeof patch.isVerified === 'boolean' ? patch.isVerified : prev.isVerified,
        updatedAt: new Date().toISOString(),
      };

      cases[idx] = updated;
      writeCases(cases);
      res.json({ caseItem: updated, cases });
    } catch (err) {
      res.status(500).json({ error: (err as Error).message });
    }
  });

  app.delete('/api/cases/:id', (req, res) => {
    try {
      const cases = readCases();
      const exists = cases.some((c) => c.id === req.params.id);
      if (!exists) {
        res.status(404).json({ error: 'Không tìm thấy hồ sơ khách hàng.' });
        return;
      }
      deleteCaseDirectory(req.params.id);
      const remaining = cases.filter((c) => c.id !== req.params.id);
      writeCases(remaining);
      res.json({ cases: remaining });
    } catch (err) {
      res.status(500).json({ error: (err as Error).message });
    }
  });

  app.post('/api/cases/:id/documents', uploadMemory.array('files', 30), (req, res) => {
    try {
      const cases = readCases();
      const idx = cases.findIndex((c) => c.id === req.params.id);
      if (idx === -1) {
        res.status(404).json({ error: 'Không tìm thấy hồ sơ khách hàng.' });
        return;
      }

      const files = (req.files as Express.Multer.File[]) ?? [];
      if (files.length === 0) {
        res.status(400).json({ error: 'Vui lòng chọn ít nhất một ảnh hoặc file PDF.' });
        return;
      }

      const docTypeHint = (
        typeof req.body.docTypeHint === 'string' ? req.body.docTypeHint : 'other'
      ) as DocumentTypeHint;

      ensureCaseDirs(req.params.id);
      const now = new Date().toISOString();
      const addedDocs: UploadedDoc[] = [];

      for (const file of files) {
        const originalName = decodeUploadFilename(file.originalname);
        const docId = `doc_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
        const diskName = `${docId}_${sanitizeFileBase(originalName)}`;
        const absPath = getCaseDocFilePath(req.params.id, diskName);
        fs.writeFileSync(absPath, file.buffer);

        let effectiveHint = docTypeHint;
        if (file.mimetype === 'application/pdf' || originalName.toLowerCase().endsWith('.pdf')) {
          effectiveHint = docTypeHint === 'other' ? 'pdf' : docTypeHint;
        }

        const doc: UploadedDoc = {
          id: docId,
          originalName,
          mimeType: file.mimetype || 'application/octet-stream',
          sizeBytes: file.size,
          docTypeHint: effectiveHint,
          fileNameOnDisk: diskName,
          uploadedAt: now,
        };
        addedDocs.push(doc);
      }

      const updated: CustomerCase = {
        ...cases[idx],
        documents: [...cases[idx].documents, ...addedDocs],
        updatedAt: now,
      };
      cases[idx] = updated;
      writeCases(cases);

      res.json({ addedDocs, caseItem: updated });
    } catch (err) {
      res.status(500).json({ error: (err as Error).message });
    }
  });

  app.get('/api/cases/:id/documents/:docId/file', (req, res) => {
    try {
      const cases = readCases();
      const caseItem = cases.find((c) => c.id === req.params.id);
      if (!caseItem) {
        res.status(404).json({ error: 'Không tìm thấy hồ sơ.' });
        return;
      }
      const doc = caseItem.documents.find((d) => d.id === req.params.docId);
      if (!doc) {
        res.status(404).json({ error: 'Không tìm thấy tài liệu.' });
        return;
      }
      const absPath = getCaseDocFilePath(caseItem.id, doc.fileNameOnDisk);
      if (!fs.existsSync(absPath)) {
        res.status(404).json({ error: 'File tài liệu không tồn tại.' });
        return;
      }
      res.setHeader('Content-Type', doc.mimeType);
      res.send(fs.readFileSync(absPath));
    } catch (err) {
      res.status(500).json({ error: (err as Error).message });
    }
  });

  app.delete('/api/cases/:id/documents/:docId', (req, res) => {
    try {
      const cases = readCases();
      const idx = cases.findIndex((c) => c.id === req.params.id);
      if (idx === -1) {
        res.status(404).json({ error: 'Không tìm thấy hồ sơ.' });
        return;
      }
      const targetDoc = cases[idx].documents.find((d) => d.id === req.params.docId);
      if (targetDoc) {
        const absPath = getCaseDocFilePath(req.params.id, targetDoc.fileNameOnDisk);
        if (fs.existsSync(absPath)) fs.unlinkSync(absPath);
      }
      cases[idx] = {
        ...cases[idx],
        documents: cases[idx].documents.filter((d) => d.id !== req.params.docId),
        updatedAt: new Date().toISOString(),
      };
      writeCases(cases);
      res.json({ caseItem: cases[idx] });
    } catch (err) {
      res.status(500).json({ error: (err as Error).message });
    }
  });

  // AI EXTRACTION FOR THIS SPECIFIC CASE
  app.post('/api/cases/:id/analyze', async (req, res) => {
    try {
      const cases = readCases();
      const idx = cases.findIndex((c) => c.id === req.params.id);
      if (idx === -1) {
        res.status(404).json({ error: 'Không tìm thấy hồ sơ khách hàng.' });
        return;
      }

      const caseItem = cases[idx];
      if (caseItem.documents.length === 0) {
        res.status(400).json({ error: 'Vui lòng tải lên ít nhất 1 giấy tờ (CCCD, Sổ đỏ, hợp đồng hoặc PDF) trước khi phân tích.' });
        return;
      }

      const templates = readTemplates();
      const selectedTemplates = templates.filter((t) => caseItem.selectedTemplateIds.includes(t.id));

      // Collect all target field keys from selected templates
      const targetFieldMap = new Map<string, FieldDefinition>();
      for (const tpl of selectedTemplates) {
        for (const m of tpl.mappings) {
          if (!targetFieldMap.has(m.fieldKey)) {
            targetFieldMap.set(m.fieldKey, {
              key: m.fieldKey,
              label: m.label,
              group: m.group,
            });
          }
        }
      }

      const docsWithPaths = caseItem.documents
        .map((doc) => ({
          doc,
          absolutePath: getCaseDocFilePath(caseItem.id, doc.fileNameOnDisk),
        }))
        .filter((d) => fs.existsSync(d.absolutePath));

      const result = await extractCustomerDocuments(
        docsWithPaths,
        Array.from(targetFieldMap.values())
      );

      // Preserve any custom manual fields already typed by user if AI returned empty for them
      const mergedFields = { ...result.fields };
      for (const [k, existingVal] of Object.entries(caseItem.extractedFields || {})) {
        if (!mergedFields[k]) {
          mergedFields[k] = existingVal;
        } else if (!mergedFields[k].value && existingVal.value) {
          mergedFields[k] = existingVal;
        }
      }
      const enrichedFields = deriveGranularFields(mergedFields);

      const suggestedName =
        result.suggestedCustomerName ||
        enrichedFields.ben_mua_ho_ten?.value ||
        enrichedFields.ben_ban_ho_ten?.value ||
        caseItem.customerName;

      const updated: CustomerCase = {
        ...caseItem,
        customerName: suggestedName || caseItem.customerName,
        title:
          caseItem.title.startsWith('Hồ sơ khách hàng mới') && suggestedName
            ? `Hồ sơ ${suggestedName}`
            : caseItem.title,
        extractedFields: enrichedFields,
        aiSummary: result.summary,
        status: 'extracted',
        isVerified: false,
        updatedAt: new Date().toISOString(),
      };

      cases[idx] = updated;
      writeCases(cases);

      res.json({ caseItem: updated });
    } catch (err) {
      console.error('AI analyze error:', err);
      res.status(500).json({
        error: `Lỗi khi AI phân tích hồ sơ: ${(err as Error).message}`,
      });
    }
  });

  // FILL SELECTED WORD TEMPLATES FOR THIS CASE (NO AI REWRITING, PURE XML PLACEHOLDER/ANCHOR REPLACEMENT)
  app.post('/api/cases/:id/fill', async (req, res) => {
    try {
      const cases = readCases();
      const idx = cases.findIndex((c) => c.id === req.params.id);
      if (idx === -1) {
        res.status(404).json({ error: 'Không tìm thấy hồ sơ khách hàng.' });
        return;
      }

      const caseItem = cases[idx];
      if (caseItem.selectedTemplateIds.length === 0) {
        res.status(400).json({ error: 'Vui lòng chọn ít nhất một mẫu Word để điền thông tin.' });
        return;
      }

      const templates = readTemplates();
      const selectedTemplates = templates.filter((t) => caseItem.selectedTemplateIds.includes(t.id));
      if (selectedTemplates.length === 0) {
        res.status(400).json({ error: 'Các mẫu Word đã chọn không còn tồn tại.' });
        return;
      }

      ensureCaseDirs(caseItem.id);
      const now = new Date().toISOString();
      const outputs: GeneratedOutput[] = [];
      const customerSlug = sanitizeFileBase(caseItem.customerName || 'Khach_hang');
      const effectiveFields = deriveGranularFields(caseItem.extractedFields || {});
      let templatesChanged = false;

      const missingFiles = selectedTemplates.filter(
        (tpl) => !fs.existsSync(getTemplateFilePath(tpl.fileNameOnDisk))
      );
      if (missingFiles.length > 0) {
        const names = missingFiles.map((t) => `"${t.name}"`).join(', ');
        res.status(400).json({
          error: `File .docx gốc của mẫu ${names} hiện không còn trên máy chủ. Vui lòng bấm "Tải lại .docx" cho mẫu đó ở Bước 1 (hoặc Kho mẫu Word), hoặc bỏ chọn và chọn mẫu khác.`,
        });
        return;
      }

      for (const tpl of selectedTemplates) {
        const tplAbsPath = getTemplateFilePath(tpl.fileNameOnDisk);
        if (!fs.existsSync(tplAbsPath)) continue;

        const templateBuf = fs.readFileSync(tplAbsPath);

        // If an uploaded old Word document has 0 configured mappings (or only stale [01]..[26] tags), auto-detect via AI before filling
        if (tpl.mappings.length === 0 || tpl.mappings.every(isLegacyNumericBracketMapping)) {
          tpl.mappings = [];
          try {
            const inspected = inspectDocx(templateBuf);
            const aiMappings = await detectTemplateMappingsWithAI(inspected.paragraphs);
            if (aiMappings.length > 0) {
              tpl.mappings = aiMappings;
              tpl.placeholdersDetected = inspected.placeholders;
              tpl.updatedAt = now;
              templatesChanged = true;
            }
          } catch (aiErr) {
            console.warn('Auto-detect during fill skipped:', (aiErr as Error).message);
          }
        }

        const filled = fillDocx(templateBuf, tpl.mappings, effectiveFields, {
          eraseUnfilledOldText: caseItem.fillOptions?.eraseUnfilledOldText ?? true,
          autoInsertSpousePartyA: caseItem.fillOptions?.autoInsertSpousePartyA ?? false,
          autoInsertSpousePartyB: caseItem.fillOptions?.autoInsertSpousePartyB ?? false,
          customSpouseTextPartyA: caseItem.fillOptions?.customSpouseTextPartyA,
          customSpouseTextPartyB: caseItem.fillOptions?.customSpouseTextPartyB,
        });

        const baseTplSlug = sanitizeFileBase(tpl.name);
        const outputFileName = `${baseTplSlug}_${customerSlug}.docx`;
        const outputPath = getCaseOutputFilePath(caseItem.id, outputFileName);
        fs.writeFileSync(outputPath, filled.outputBuffer);

        outputs.push({
          templateId: tpl.id,
          templateName: tpl.name,
          outputFileName,
          generatedAt: now,
          filledCount: filled.filledCount,
          totalMappings: filled.totalMappings,
          missingKeys: filled.missingKeys,
        });
      }

      if (outputs.length === 0) {
        res.status(400).json({
          error: 'Không tạo được file Word nào. Vui lòng kiểm tra lại các mẫu đã chọn.',
        });
        return;
      }

      if (templatesChanged) {
        writeTemplates(templates);
      }

      const updated: CustomerCase = {
        ...caseItem,
        extractedFields: effectiveFields,
        isVerified: true,
        status: 'exported',
        generatedOutputs: outputs,
        updatedAt: now,
      };

      cases[idx] = updated;
      writeCases(cases);

      res.json({ caseItem: updated, outputs, templates: readTemplates() });
    } catch (err) {
      console.error('Fill docx error:', err);
      res.status(500).json({ error: `Không thể điền file Word: ${(err as Error).message}` });
    }
  });

  app.get('/api/cases/:id/outputs/:fileName', (req, res) => {
    try {
      const cases = readCases();
      const caseItem = cases.find((c) => c.id === req.params.id);
      if (!caseItem) {
        res.status(404).json({ error: 'Không tìm thấy hồ sơ.' });
        return;
      }
      const absPath = getCaseOutputFilePath(caseItem.id, req.params.fileName);
      if (!fs.existsSync(absPath)) {
        res.status(404).json({ error: 'File kết quả chưa được tạo.' });
        return;
      }
      setContentDisposition(res, path.basename(req.params.fileName));
      res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      );
      res.send(fs.readFileSync(absPath));
    } catch (err) {
      res.status(500).json({ error: (err as Error).message });
    }
  });

  app.get('/api/cases/:id/outputs-zip', async (req, res) => {
    try {
      const cases = readCases();
      const caseItem = cases.find((c) => c.id === req.params.id);
      if (!caseItem) {
        res.status(404).json({ error: 'Không tìm thấy hồ sơ.' });
        return;
      }
      if (caseItem.generatedOutputs.length === 0) {
        res.status(400).json({ error: 'Chưa có file Word nào được xuất trong hồ sơ này.' });
        return;
      }

      const zip = new JSZip();
      for (const out of caseItem.generatedOutputs) {
        const absPath = getCaseOutputFilePath(caseItem.id, out.outputFileName);
        if (fs.existsSync(absPath)) {
          zip.file(out.outputFileName, fs.readFileSync(absPath));
        }
      }

      const zipBuffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
      const zipName = `Ho_so_${sanitizeFileBase(caseItem.customerName || caseItem.id)}.zip`;
      setContentDisposition(res, zipName);
      res.setHeader('Content-Type', 'application/zip');
      res.send(zipBuffer);
    } catch (err) {
      res.status(500).json({ error: (err as Error).message });
    }
  });

  // ---------------------------------------------------------------------------
  // VITE DEV MIDDLEWARE OR PRODUCTION STATIC DIST
  // ---------------------------------------------------------------------------
  if (process.env.NODE_ENV !== 'production') {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distDir = path.resolve(process.cwd(), 'dist');
    app.use(express.static(distDir));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(distDir, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`HỒ SƠ AI server running at http://0.0.0.0:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error('Failed to start HỒ SƠ AI server:', err);
  process.exit(1);
});
