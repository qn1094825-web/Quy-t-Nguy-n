import fs from 'node:fs';
import path from 'node:path';
import PizZip from 'pizzip';
import type { CustomerCase, WordTemplate } from '../src/shared/types.ts';
import { isLegacyNumericBracketMapping } from '../src/shared/types.ts';
import { inspectDocx } from './docxEngine.ts';
import { createStarterDocxTemplates } from './sampleTemplates.ts';

const DATA_ROOT = path.resolve(process.cwd(), 'data');
const TEMPLATES_DIR = path.join(DATA_ROOT, 'templates');
const TEMPLATE_FILES_DIR = path.join(TEMPLATES_DIR, 'files');
const TEMPLATES_META_PATH = path.join(TEMPLATES_DIR, 'meta.json');

const CASES_DIR = path.join(DATA_ROOT, 'cases');
const CASES_META_PATH = path.join(CASES_DIR, 'cases.json');

function ensureDir(dirPath: string) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

function repairDocxContentTypesIfNeeded(filePath: string) {
  if (!filePath.toLowerCase().endsWith('.docx') || !fs.existsSync(filePath)) return;
  try {
    const buf = fs.readFileSync(filePath);
    const zip = new PizZip(buf);
    const ct = zip.file('[Content_Types].xml');
    if (!ct) return;
    const text = ct.asText();
    if (text.includes('schemas.openxmlformats.org/package/2006/contenttypes"')) {
      zip.file(
        '[Content_Types].xml',
        text.replace(
          'schemas.openxmlformats.org/package/2006/contenttypes"',
          'schemas.openxmlformats.org/package/2006/content-types"'
        )
      );
      const fixed = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }) as Buffer;
      fs.writeFileSync(filePath, fixed);
    }
  } catch {
    // Ignore unreadable files
  }
}

function repairStoredDocxFiles() {
  if (fs.existsSync(TEMPLATE_FILES_DIR)) {
    for (const name of fs.readdirSync(TEMPLATE_FILES_DIR)) {
      repairDocxContentTypesIfNeeded(path.join(TEMPLATE_FILES_DIR, name));
    }
  }
  if (fs.existsSync(CASES_DIR)) {
    for (const caseId of fs.readdirSync(CASES_DIR)) {
      const outDir = path.join(CASES_DIR, caseId, 'outputs');
      if (fs.existsSync(outDir)) {
        for (const outName of fs.readdirSync(outDir)) {
          repairDocxContentTypesIfNeeded(path.join(outDir, outName));
        }
      }
    }
  }
}

export function initStorage() {
  ensureDir(DATA_ROOT);
  ensureDir(TEMPLATES_DIR);
  ensureDir(TEMPLATE_FILES_DIR);
  ensureDir(CASES_DIR);
  repairStoredDocxFiles();

  if (!fs.existsSync(TEMPLATES_META_PATH)) {
    const seeds = createStarterDocxTemplates();
    const initialTemplates: WordTemplate[] = [];
    const now = new Date().toISOString();

    for (let i = 0; i < seeds.length; i++) {
      const seed = seeds[i];
      const id = `tpl_starter_${i + 1}`;
      const diskName = `${id}_${seed.fileName}`;
      const absPath = path.join(TEMPLATE_FILES_DIR, diskName);
      fs.writeFileSync(absPath, seed.buffer);

      const inspected = inspectDocx(seed.buffer);
      initialTemplates.push({
        id,
        name: seed.name,
        originalFileName: seed.fileName,
        category: seed.category,
        fileNameOnDisk: diskName,
        sizeBytes: seed.buffer.length,
        placeholdersDetected: inspected.placeholders,
        mappings: inspected.defaultMappings,
        createdAt: now,
        updatedAt: now,
      });
    }

    fs.writeFileSync(TEMPLATES_META_PATH, JSON.stringify(initialTemplates, null, 2), 'utf-8');
  } else {
    try {
      const existing = JSON.parse(fs.readFileSync(TEMPLATES_META_PATH, 'utf-8')) as WordTemplate[];
      const seeds = createStarterDocxTemplates();
      for (let i = 0; i < seeds.length; i++) {
        const id = `tpl_starter_${i + 1}`;
        const tpl = existing.find((t) => t.id === id);
        if (tpl) {
          const absPath = getTemplateFilePath(tpl.fileNameOnDisk);
          if (!fs.existsSync(absPath)) {
            fs.writeFileSync(absPath, seeds[i].buffer);
          }
        }
      }
    } catch {
      // Ignore malformed meta.json here
    }
  }

  if (!fs.existsSync(CASES_META_PATH)) {
    const now = new Date().toISOString();
    const firstCase: CustomerCase = {
      id: `case_${Date.now()}`,
      title: 'Hồ sơ khách hàng mới #1',
      customerName: 'Chưa xác định',
      status: 'draft',
      selectedTemplateIds: [],
      documents: [],
      extractedFields: {},
      isVerified: false,
      generatedOutputs: [],
      createdAt: now,
      updatedAt: now,
    };
    ensureCaseDirs(firstCase.id);
    fs.writeFileSync(CASES_META_PATH, JSON.stringify([firstCase], null, 2), 'utf-8');
  }
}

export function ensureCaseDirs(caseId: string) {
  const safeId = path.basename(caseId);
  ensureDir(path.join(CASES_DIR, safeId));
  ensureDir(path.join(CASES_DIR, safeId, 'docs'));
  ensureDir(path.join(CASES_DIR, safeId, 'outputs'));
}

export function getTemplateFilePath(fileNameOnDisk: string): string {
  return path.join(TEMPLATE_FILES_DIR, path.basename(fileNameOnDisk));
}

export function getCaseDocFilePath(caseId: string, fileNameOnDisk: string): string {
  return path.join(CASES_DIR, path.basename(caseId), 'docs', path.basename(fileNameOnDisk));
}

export function getCaseOutputFilePath(caseId: string, outputFileName: string): string {
  return path.join(CASES_DIR, path.basename(caseId), 'outputs', path.basename(outputFileName));
}

export function readTemplates(): WordTemplate[] {
  initStorage();
  const raw = fs.readFileSync(TEMPLATES_META_PATH, 'utf-8');
  const templates = JSON.parse(raw) as WordTemplate[];
  let changed = false;

  for (const tpl of templates) {
    const absPath = getTemplateFilePath(tpl.fileNameOnDisk);
    const existsOnDisk = fs.existsSync(absPath);
    if (tpl.mappings.length > 0 && tpl.mappings.every(isLegacyNumericBracketMapping)) {
      tpl.mappings = [];
      if (existsOnDisk) {
        try {
          const inspected = inspectDocx(fs.readFileSync(absPath));
          tpl.placeholdersDetected = inspected.placeholders;
          if (inspected.defaultMappings.length > 0) {
            tpl.mappings = inspected.defaultMappings;
          }
        } catch {
          tpl.placeholdersDetected = [];
        }
      } else {
        tpl.placeholdersDetected = tpl.placeholdersDetected.filter((p) => !/^\[\d+\]$/.test(p));
      }
      changed = true;
    }
  }

  if (changed) {
    writeTemplates(templates);
  }
  return templates.map((tpl) => ({
    ...tpl,
    fileMissing: !fs.existsSync(getTemplateFilePath(tpl.fileNameOnDisk)),
  }));
}

export function writeTemplates(templates: WordTemplate[]) {
  ensureDir(TEMPLATES_DIR);
  const clean = templates.map(({ fileMissing: _fm, ...rest }) => rest);
  fs.writeFileSync(TEMPLATES_META_PATH, JSON.stringify(clean, null, 2), 'utf-8');
}

export function readCases(): CustomerCase[] {
  initStorage();
  const raw = fs.readFileSync(CASES_META_PATH, 'utf-8');
  return JSON.parse(raw) as CustomerCase[];
}

export function writeCases(cases: CustomerCase[]) {
  ensureDir(CASES_DIR);
  fs.writeFileSync(CASES_META_PATH, JSON.stringify(cases, null, 2), 'utf-8');
}

export function deleteCaseDirectory(caseId: string) {
  const dir = path.join(CASES_DIR, path.basename(caseId));
  if (fs.existsSync(dir)) {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

export function restoreStarterTemplates(): WordTemplate[] {
  const existing = readTemplates();
  const existingIds = new Set(existing.map((t) => t.id));
  const seeds = createStarterDocxTemplates();
  const now = new Date().toISOString();

  for (let i = 0; i < seeds.length; i++) {
    const seed = seeds[i];
    const id = `tpl_starter_${i + 1}`;
    if (existingIds.has(id)) continue;

    const diskName = `${id}_${seed.fileName}`;
    const absPath = path.join(TEMPLATE_FILES_DIR, diskName);
    fs.writeFileSync(absPath, seed.buffer);

    const inspected = inspectDocx(seed.buffer);
    existing.push({
      id,
      name: seed.name,
      originalFileName: seed.fileName,
      category: seed.category,
      fileNameOnDisk: diskName,
      sizeBytes: seed.buffer.length,
      placeholdersDetected: inspected.placeholders,
      mappings: inspected.defaultMappings,
      createdAt: now,
      updatedAt: now,
    });
  }

  writeTemplates(existing);
  return existing;
}
