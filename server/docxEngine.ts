import PizZip from 'pizzip';
import {
  DOMParser,
  XMLSerializer,
  type Document as XmlDocument,
  type Element as XmlElement,
  type Node as XmlNode,
} from '@xmldom/xmldom';
import type {
  DocxParagraphInspect,
  ExtractedFieldValue,
  TemplateFieldMapping,
} from '../src/shared/types.ts';
import { normalizeKeyToVietnameseLabel } from '../src/shared/types.ts';

interface RunSpan {
  tNode: XmlElement | null;
  rNode: XmlElement | null;
  text: string;
  start: number;
  end: number;
}

function getTargetXmlParts(zip: PizZip): string[] {
  const parts: string[] = [];
  if (zip.file('word/document.xml')) {
    parts.push('word/document.xml');
  }
  const extraFiles = zip.file(/^word\/(header|footer)\d*\.xml$/);
  for (const f of extraFiles) {
    if (!parts.includes(f.name)) {
      parts.push(f.name);
    }
  }
  return parts;
}

function isClosestParagraph(node: XmlNode, targetP: XmlElement): boolean {
  let curr: XmlNode | null = node.parentNode;
  while (curr) {
    if (curr.nodeType === 1 && (curr as XmlElement).tagName === 'w:p') {
      return curr === targetP;
    }
    curr = curr.parentNode;
  }
  return false;
}

function collectParagraphRuns(pNode: XmlElement): { fullText: string; runs: RunSpan[] } {
  const runs: RunSpan[] = [];
  let cursor = 0;

  function walk(node: XmlNode, currentR: XmlElement | null) {
    if (node.nodeType !== 1) return;
    const el = node as XmlElement;
    const tag = el.tagName;

    // Do not descend into nested paragraphs (e.g. inside textboxes or nested tables)
    if (tag === 'w:p' && el !== pNode) return;

    const nextR = tag === 'w:r' ? el : currentR;

    if (tag === 'w:t') {
      const text = el.textContent ?? '';
      const start = cursor;
      const end = cursor + text.length;
      runs.push({ tNode: el, rNode: nextR, text, start, end });
      cursor = end;
      return;
    }

    if (tag === 'w:tab') {
      const start = cursor;
      const end = cursor + 1;
      runs.push({ tNode: null, rNode: nextR, text: '\t', start, end });
      cursor = end;
      return;
    }

    if (tag === 'w:br' || tag === 'w:cr') {
      const start = cursor;
      const end = cursor + 1;
      runs.push({ tNode: null, rNode: nextR, text: '\n', start, end });
      cursor = end;
      return;
    }

    const children = el.childNodes;
    for (let i = 0; i < children.length; i++) {
      const child = children.item(i);
      if (child) walk(child, nextR);
    }
  }

  walk(pNode, null);

  return {
    fullText: runs.map((r) => r.text).join(''),
    runs,
  };
}

function describeParagraphLocation(pNode: XmlElement, indexInPart: number, partName: string): string {
  const partPrefix =
    partName === 'word/document.xml'
      ? ''
      : partName.includes('header')
        ? 'Đầu trang · '
        : 'Chân trang · ';

  let curr: XmlNode | null = pNode.parentNode;
  let inTable = false;
  while (curr) {
    if (curr.nodeType === 1 && (curr as XmlElement).tagName === 'w:tc') {
      inTable = true;
      break;
    }
    curr = curr.parentNode;
  }

  return `${partPrefix}${inTable ? 'Ô bảng' : 'Đoạn'} #${indexInPart + 1}`;
}

export function extractPlaceholderKeyFromTag(rawTag: string): string {
  return rawTag
    .replace(/^\{\{|\}\}$/g, '')
    .replace(/^\{|\}$/g, '')
    .replace(/^\[|\]$/g, '')
    .trim();
}

const PLACEHOLDER_REGEX = /\{\{([^{}]+)\}\}|\{([a-zA-Z0-9_.\-À-ỹ]+)\}|\[([A-Z][A-Z0-9_]+)\]/g;
const BLANK_MARKER_REGEX = /(\.{3,}|…{2,}|_{3,})/g;

function isFirstParagraphInCell(pNode: XmlElement): boolean {
  let curr: XmlNode | null = pNode.parentNode;
  while (curr) {
    if (curr.nodeType === 1 && (curr as XmlElement).tagName === 'w:tc') {
      const ps = (curr as XmlElement).getElementsByTagName('w:p');
      return ps.length > 0 && ps.item(0) === pNode;
    }
    curr = curr.parentNode;
  }
  return false;
}

export function inspectDocx(buffer: Buffer): {
  placeholders: string[];
  defaultMappings: TemplateFieldMapping[];
  paragraphs: DocxParagraphInspect[];
} {
  const zip = new PizZip(buffer);
  const parts = getTargetXmlParts(zip);
  const parser = new DOMParser();

  const placeholderSet = new Set<string>();
  const paragraphs: DocxParagraphInspect[] = [];

  for (const partName of parts) {
    const file = zip.file(partName);
    if (!file) continue;
    const xmlStr = file.asText();
    const doc = parser.parseFromString(xmlStr, 'application/xml');
    const pNodes = doc.getElementsByTagName('w:p');

    for (let i = 0; i < pNodes.length; i++) {
      const pNode = pNodes.item(i);
      if (!pNode) continue;

      const { fullText } = collectParagraphRuns(pNode);
      const trimmed = fullText.trim();
      const emptyTableCell =
        isParagraphInTableCell(pNode) &&
        isFirstParagraphInCell(pNode) &&
        (trimmed === '' || trimmed === '.' || trimmed === '-');
      if (!trimmed && !emptyTableCell) continue;

      const detectedPlaceholders: string[] = [];
      let m: RegExpExecArray | null;
      PLACEHOLDER_REGEX.lastIndex = 0;
      while ((m = PLACEHOLDER_REGEX.exec(fullText)) !== null) {
        const rawTag = m[0];
        const key = (m[1] || m[2] || m[3] || '').trim();
        if (key) {
          placeholderSet.add(rawTag);
          if (!detectedPlaceholders.includes(rawTag)) {
            detectedPlaceholders.push(rawTag);
          }
        }
      }

      const detectedBlanks: string[] = [];
      BLANK_MARKER_REGEX.lastIndex = 0;
      while ((m = BLANK_MARKER_REGEX.exec(fullText)) !== null) {
        detectedBlanks.push(m[0]);
      }

      paragraphs.push({
        partName,
        paragraphIndex: i,
        locationLabel: describeParagraphLocation(pNode, i, partName),
        text: fullText,
        detectedPlaceholders,
        detectedBlanks,
      });
    }
  }

  const placeholders = Array.from(placeholderSet);
  const seenKeys = new Set<string>();
  const defaultMappings: TemplateFieldMapping[] = [];

  for (const rawTag of placeholders) {
    const fieldKey = extractPlaceholderKeyFromTag(rawTag);
    if (!fieldKey || seenKeys.has(fieldKey)) continue;
    seenKeys.add(fieldKey);

    const { label, group } = normalizeKeyToVietnameseLabel(fieldKey);
    defaultMappings.push({
      id: `map_${fieldKey}_${Math.random().toString(36).slice(2, 7)}`,
      fieldKey,
      label,
      group,
      mode: 'placeholder',
      placeholderTag: rawTag,
    });
  }

  return { placeholders, defaultMappings, paragraphs };
}

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function sanitizeXmlText(text: string): string {
  // Strip XML 1.0 forbidden control chars while keeping \t, \n, \r
  return text.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]/g, '');
}

function setNodeTextWithNewlines(
  doc: XmlDocument,
  tNode: XmlElement,
  rNode: XmlElement | null,
  newText: string
) {
  const lines = sanitizeXmlText(newText).split(/\r?\n/);
  tNode.textContent = lines[0] ?? '';
  tNode.setAttribute('xml:space', 'preserve');

  if (lines.length > 1 && rNode) {
    let anchor: XmlNode = tNode;
    for (let idx = 1; idx < lines.length; idx++) {
      const br = doc.createElementNS(W_NS, 'w:br');
      const nextT = doc.createElementNS(W_NS, 'w:t');
      nextT.setAttribute('xml:space', 'preserve');
      nextT.textContent = lines[idx] ?? '';
      if (anchor.nextSibling) {
        rNode.insertBefore(br, anchor.nextSibling);
        rNode.insertBefore(nextT, br.nextSibling);
      } else {
        rNode.appendChild(br);
        rNode.appendChild(nextT);
      }
      anchor = nextT;
    }
  }
}

interface PendingReplacement {
  start: number;
  end: number;
  replacement: string;
  fieldKey: string;
}

function normalizeWs(s: string): string {
  return s.replace(/\u00A0/g, ' ');
}

function isAllUpperCaseVi(s: string): boolean {
  const letters = s.replace(/[^a-zA-ZÀ-ỹ]/g, '');
  if (!letters) return false;
  return letters === letters.toLocaleUpperCase('vi-VN') && letters !== letters.toLocaleLowerCase('vi-VN');
}

function applyCaseLikeMatched(matchedInDoc: string, replacement: string): string {
  if (isAllUpperCaseVi(matchedInDoc)) {
    return replacement.toLocaleUpperCase('vi-VN');
  }
  return replacement;
}

function findNthOccurrence(haystack: string, needle: string, occurrence: number): number {
  if (!needle) return -1;
  const normHay = normalizeWs(haystack);
  const normNeedle = normalizeWs(needle);
  let fromIndex = 0;
  for (let count = 0; count <= occurrence; count++) {
    const idx = normHay.indexOf(normNeedle, fromIndex);
    if (idx === -1) return -1;
    if (count === occurrence) return idx;
    fromIndex = idx + normNeedle.length;
  }
  return -1;
}

function findAllTextMatches(
  haystack: string,
  needle: string,
  matchCase: boolean
): Array<{ start: number; end: number; matched: string }> {
  if (!needle) return [];
  const normHay = normalizeWs(haystack);
  const normNeedle = normalizeWs(needle);

  const searchHay = matchCase ? normHay : normHay.toLocaleLowerCase('vi-VN');
  const searchNeedle = matchCase ? normNeedle : normNeedle.toLocaleLowerCase('vi-VN');
  if (!searchNeedle) return [];

  const results: Array<{ start: number; end: number; matched: string }> = [];
  let fromIdx = 0;
  while (fromIdx <= searchHay.length - searchNeedle.length) {
    const idx = searchHay.indexOf(searchNeedle, fromIdx);
    if (idx === -1) break;
    results.push({
      start: idx,
      end: idx + normNeedle.length,
      matched: haystack.slice(idx, idx + normNeedle.length),
    });
    fromIdx = idx + normNeedle.length;
  }
  return results;
}

export interface FillDocxOptions {
  eraseUnfilledOldText?: boolean;
  autoInsertSpousePartyA?: boolean;
  autoInsertSpousePartyB?: boolean;
  customSpouseTextPartyA?: string;
  customSpouseTextPartyB?: string;
}

function isBlankMarkerOnly(s: string): boolean {
  return /^[.\u2026_\s-]+$/.test(s);
}

function buildSpouseInfoSentence(
  prefix: 'ben_ban' | 'ben_mua',
  valueByKey: Record<string, string>,
  customText?: string
): string {
  if (customText && customText.trim()) return customText.trim();
  const hoTen = (valueByKey[`${prefix}_vo_chong_ho_ten`] ?? '').trim();
  if (!hoTen) return '';

  const gioiTinh = (valueByKey[`${prefix}_vo_chong_gioi_tinh`] ?? '').trim().toLowerCase();
  const honorific =
    gioiTinh.includes('nữ') || gioiTinh === 'nu'
      ? 'Bà'
      : gioiTinh.includes('nam')
        ? 'Ông'
        : 'Ông/Bà';

  const parts: string[] = [`Và vợ/chồng là ${honorific}: ${hoTen}`];
  const ngaySinh = (valueByKey[`${prefix}_vo_chong_ngay_sinh`] ?? '').trim();
  if (ngaySinh) parts.push(`sinh ngày: ${ngaySinh}`);
  const cccd = (valueByKey[`${prefix}_vo_chong_cccd`] ?? '').trim();
  if (cccd) parts.push(`số CCCD: ${cccd}`);
  const ngayCap = (valueByKey[`${prefix}_vo_chong_ngay_cap`] ?? '').trim();
  if (ngayCap) parts.push(`cấp ngày: ${ngayCap}`);
  const noiCap = (valueByKey[`${prefix}_vo_chong_noi_cap`] ?? '').trim();
  if (noiCap) parts.push(`tại ${noiCap}`);
  const diaChi = (valueByKey[`${prefix}_vo_chong_dia_chi`] ?? '').trim();
  if (diaChi) parts.push(`địa chỉ thường trú: ${diaChi}`);

  return parts.join(', ') + '.';
}

function isParagraphInTableCell(pNode: XmlElement): boolean {
  let curr: XmlNode | null = pNode.parentNode;
  while (curr) {
    if (curr.nodeType === 1 && (curr as XmlElement).tagName === 'w:tc') {
      return true;
    }
    curr = curr.parentNode;
  }
  return false;
}

function cloneParagraphWithText(doc: XmlDocument, templateP: XmlElement, text: string): XmlElement {
  const newP = doc.createElementNS(W_NS, 'w:p');

  // Copy paragraph formatting (<w:pPr>) if present
  const pPrList = templateP.getElementsByTagName('w:pPr');
  if (pPrList.length > 0 && pPrList.item(0)?.parentNode === templateP) {
    newP.appendChild(pPrList.item(0)!.cloneNode(true));
  }

  const newR = doc.createElementNS(W_NS, 'w:r');
  // Copy run formatting (<w:rPr>) from first run of template paragraph, stripping bold so inserted prose matches body text
  const rList = templateP.getElementsByTagName('w:r');
  for (let i = 0; i < rList.length; i++) {
    const r = rList.item(i);
    if (!r || !isClosestParagraph(r, templateP)) continue;
    const rPrList = r.getElementsByTagName('w:rPr');
    if (rPrList.length > 0 && rPrList.item(0)?.parentNode === r) {
      const clonedRPr = rPrList.item(0)!.cloneNode(true) as XmlElement;
      const bNodes = clonedRPr.getElementsByTagName('w:b');
      while (bNodes.length > 0) {
        const b = bNodes.item(0);
        if (b && b.parentNode) b.parentNode.removeChild(b);
        else break;
      }
      newR.appendChild(clonedRPr);
    }
    break;
  }

  const lines = sanitizeXmlText(text).split(/\r?\n/);
  const firstT = doc.createElementNS(W_NS, 'w:t');
  firstT.setAttribute('xml:space', 'preserve');
  firstT.textContent = lines[0] ?? '';
  newR.appendChild(firstT);

  for (let idx = 1; idx < lines.length; idx++) {
    const br = doc.createElementNS(W_NS, 'w:br');
    const nextT = doc.createElementNS(W_NS, 'w:t');
    nextT.setAttribute('xml:space', 'preserve');
    nextT.textContent = lines[idx] ?? '';
    newR.appendChild(br);
    newR.appendChild(nextT);
  }

  newP.appendChild(newR);
  return newP;
}

function fillEmptyParagraphWithText(doc: XmlDocument, pNode: XmlElement, text: string) {
  const safeText = sanitizeXmlText(text);
  // Reuse first existing <w:t> if present (e.g. single-dot placeholder cell), clearing any extra dot <w:t>s
  const existingTs: XmlElement[] = [];
  const tNodes = pNode.getElementsByTagName('w:t');
  for (let i = 0; i < tNodes.length; i++) {
    const t = tNodes.item(i);
    if (t && isClosestParagraph(t, pNode)) {
      existingTs.push(t);
    }
  }
  if (existingTs.length > 0) {
    existingTs[0].setAttribute('xml:space', 'preserve');
    existingTs[0].textContent = safeText;
    for (let i = 1; i < existingTs.length; i++) {
      existingTs[i].textContent = '';
    }
    return;
  }

  const newR = doc.createElementNS(W_NS, 'w:r');
  // Inherit <w:rPr> from an existing run in this paragraph, or from <w:pPr>/<w:rPr>
  let inheritedRPr: XmlElement | null = null;
  const existingRs = pNode.getElementsByTagName('w:r');
  for (let i = 0; i < existingRs.length; i++) {
    const r = existingRs.item(i);
    if (!r || !isClosestParagraph(r, pNode)) continue;
    const rPrs = r.getElementsByTagName('w:rPr');
    if (rPrs.length > 0 && rPrs.item(0)?.parentNode === r) {
      inheritedRPr = rPrs.item(0)!.cloneNode(true) as XmlElement;
      break;
    }
  }
  if (!inheritedRPr) {
    const pPrList = pNode.getElementsByTagName('w:pPr');
    if (pPrList.length > 0 && pPrList.item(0)?.parentNode === pNode) {
      const pPr = pPrList.item(0)!;
      const pRPr = pPr.getElementsByTagName('w:rPr');
      if (pRPr.length > 0 && pRPr.item(0)?.parentNode === pPr) {
        inheritedRPr = pRPr.item(0)!.cloneNode(true) as XmlElement;
      }
    }
  }
  if (!inheritedRPr) {
    // Check sibling cells in the same table row (or previous rows in the table) so font/size match the form
    let tcNode: XmlNode | null = pNode.parentNode;
    while (tcNode && (tcNode.nodeType !== 1 || (tcNode as XmlElement).tagName !== 'w:tc')) {
      tcNode = tcNode.parentNode;
    }
    const trNode = tcNode?.parentNode;
    if (trNode && trNode.nodeType === 1 && (trNode as XmlElement).tagName === 'w:tr') {
      const candidateRows: XmlElement[] = [trNode as XmlElement];
      let prev = trNode.previousSibling;
      while (prev) {
        if (prev.nodeType === 1 && (prev as XmlElement).tagName === 'w:tr') {
          candidateRows.push(prev as XmlElement);
        }
        prev = prev.previousSibling;
      }
      for (const row of candidateRows) {
        if (inheritedRPr) break;
        const rowRs = row.getElementsByTagName('w:r');
        for (let i = 0; i < rowRs.length; i++) {
          const r = rowRs.item(i);
          if (!r) continue;
          const rPrs = r.getElementsByTagName('w:rPr');
          if (rPrs.length > 0 && rPrs.item(0)?.parentNode === r) {
            const cloned = rPrs.item(0)!.cloneNode(true) as XmlElement;
            // Strip bold/italic if borrowed from another column/header so inserted data stays regular body weight
            for (const tag of ['w:b', 'w:bCs', 'w:i', 'w:iCs']) {
              const nodes = cloned.getElementsByTagName(tag);
              while (nodes.length > 0) {
                const n = nodes.item(0);
                if (n && n.parentNode) n.parentNode.removeChild(n);
                else break;
              }
            }
            inheritedRPr = cloned;
            break;
          }
        }
      }
    }
  }
  if (inheritedRPr) {
    newR.appendChild(inheritedRPr);
  }

  const tNode = doc.createElementNS(W_NS, 'w:t');
  tNode.setAttribute('xml:space', 'preserve');
  tNode.textContent = safeText;
  newR.appendChild(tNode);
  pNode.appendChild(newR);
}

function stripDuplicateUnitSuffix(
  fieldKey: string,
  replacement: string,
  followingDocText: string
): string {
  if (!replacement) return replacement;
  const after = followingDocText.trimStart();
  if (!after) return replacement;

  if (fieldKey === 'dien_tich' && /^(m2|m²)/i.test(after)) {
    return replacement.replace(/\s*(m2|m²)\s*$/i, '').trim();
  }
  if (fieldKey.startsWith('ty_le_so_huu') && after.startsWith('%')) {
    return replacement.replace(/\s*%\s*$/, '').trim();
  }
  if (fieldKey === 'gia_chuyen_nhuong' && /^(đồng|vnd|vnđ)/i.test(after)) {
    return replacement.replace(/\s*(đồng|vnd|vnđ|đ)\s*$/i, '').trim();
  }
  return replacement;
}

export function fillDocx(
  templateBuffer: Buffer,
  mappings: TemplateFieldMapping[],
  extractedFields: Record<string, ExtractedFieldValue>,
  options: FillDocxOptions = {}
): {
  outputBuffer: Buffer;
  filledCount: number;
  totalMappings: number;
  missingKeys: string[];
} {
  const eraseUnfilledOldText = options.eraseUnfilledOldText ?? true;
  const autoInsertSpousePartyA = options.autoInsertSpousePartyA ?? false;
  const autoInsertSpousePartyB = options.autoInsertSpousePartyB ?? false;

  const zip = new PizZip(templateBuffer);
  const parts = getTargetXmlParts(zip);
  const parser = new DOMParser();
  const serializer = new XMLSerializer();

  const valueByKey: Record<string, string> = {};
  for (const [k, v] of Object.entries(extractedFields)) {
    if (v && typeof v.value === 'string') {
      valueByKey[k] = v.value;
    }
  }

  const placeholderToKey = new Map<string, string>();
  for (const m of mappings) {
    if (m.mode === 'placeholder' && m.placeholderTag) {
      placeholderToKey.set(m.placeholderTag, m.fieldKey);
    }
  }

  // Check if template already has explicit spouse fields
  const templateHasPartyASpouse = mappings.some((m) => m.fieldKey.startsWith('ben_ban_vo_chong_'));
  const templateHasPartyBSpouse = mappings.some((m) => m.fieldKey.startsWith('ben_mua_vo_chong_'));

  // Sort global text_replace and erase_text rules longest-first to avoid partial substring collisions
  const textReplaceMappings = mappings
    .filter(
      (m) =>
        (m.mode === 'text_replace' || m.mode === 'erase_text') &&
        (m.targetText ?? '').trim().length > 0
    )
    .sort((a, b) => (b.targetText?.length ?? 0) - (a.targetText?.length ?? 0));

  let filledCount = 0;
  const missingSet = new Set<string>();
  const matchedMappingIds = new Set<string>();

  // Track best anchor paragraph to insert spouse info line after (if template has no separate spouse slot)
  let partyAInfoAnchorP: XmlElement | null = null;
  let partyBInfoAnchorP: XmlElement | null = null;
  let partyADocRef: XmlDocument | null = null;
  let partyBDocRef: XmlDocument | null = null;

  for (const partName of parts) {
    const file = zip.file(partName);
    if (!file) continue;
    const xmlStr = file.asText();
    const doc = parser.parseFromString(xmlStr, 'application/xml');
    const pNodes = doc.getElementsByTagName('w:p');
    let partModified = false;

    for (let pIdx = 0; pIdx < pNodes.length; pIdx++) {
      const pNode = pNodes.item(pIdx);
      if (!pNode) continue;

      const { fullText, runs } = collectParagraphRuns(pNode);

      // 0. Empty table cell fills (mode === 'cell_fill')
      const cellFillMappings = mappings.filter(
        (m) =>
          m.mode === 'cell_fill' &&
          (m.partName || 'word/document.xml') === partName &&
          m.paragraphIndex === pIdx
      );
      const trimmedPara = fullText.trim();
      if (
        cellFillMappings.length > 0 &&
        (trimmedPara === '' || trimmedPara === '.' || trimmedPara === '-')
      ) {
        const m = cellFillMappings[0];
        const val = (valueByKey[m.fieldKey] ?? '').trim();
        if (val) {
          fillEmptyParagraphWithText(doc, pNode, val);
          matchedMappingIds.add(m.id);
          filledCount++;
          partModified = true;
        } else {
          missingSet.add(m.fieldKey);
        }
        continue;
      }

      if (!fullText || runs.length === 0) continue;

      const replacements: PendingReplacement[] = [];

      // 1. Fixed-position anchor replacements & insert_after configured for this paragraph
      const anchorMappings = mappings.filter(
        (m) =>
          (m.mode === 'anchor_replace' || m.mode === 'insert_after') &&
          (m.partName || 'word/document.xml') === partName &&
          m.paragraphIndex === pIdx &&
          m.targetText
      );

      for (const m of anchorMappings) {
        const idx = findNthOccurrence(fullText, m.targetText!, m.occurrenceInParagraph ?? 0);
        if (idx !== -1) {
          const val = (valueByKey[m.fieldKey] ?? '').trim();
          const end = idx + m.targetText!.length;
          const overlaps = replacements.some((r) => !(end <= r.start || idx >= r.end));
          if (overlaps) continue;

          if (m.mode === 'insert_after') {
            if (val) {
              const matchedInDoc = fullText.slice(idx, end);
              const afterStr = fullText.slice(end);
              const trailingSpacesMatch = afterStr.match(/^( {3,})/);
              const consumedSpaces = trailingSpacesMatch
                ? Math.min(trailingSpacesMatch[1].length, val.length + 1)
                : 0;
              const cleanVal = stripDuplicateUnitSuffix(
                m.fieldKey,
                val,
                fullText.slice(end + consumedSpaces)
              );
              const spacer = /\s$/.test(matchedInDoc) ? '' : ' ';
              replacements.push({
                start: idx,
                end: end + consumedSpaces,
                replacement: `${matchedInDoc}${spacer}${cleanVal}`,
                fieldKey: m.fieldKey,
              });
              matchedMappingIds.add(m.id);
            } else {
              missingSet.add(m.fieldKey);
            }
            continue;
          }

          if (val) {
            const matchedInDoc = fullText.slice(idx, end);
            const cleanVal = stripDuplicateUnitSuffix(m.fieldKey, val, fullText.slice(end));
            replacements.push({
              start: idx,
              end,
              replacement: applyCaseLikeMatched(matchedInDoc, cleanVal),
              fieldKey: m.fieldKey,
            });
            matchedMappingIds.add(m.id);
          } else if (eraseUnfilledOldText && !isBlankMarkerOnly(m.targetText!)) {
            // Erase old literal value from template so old client data never leaks
            replacements.push({
              start: idx,
              end,
              replacement: '',
              fieldKey: m.fieldKey,
            });
            missingSet.add(m.fieldKey);
          } else {
            missingSet.add(m.fieldKey);
          }
        }
      }

      // 2. Placeholder tag replacements ({{key}}, {key}, [KEY])
      PLACEHOLDER_REGEX.lastIndex = 0;
      let match: RegExpExecArray | null;
      while ((match = PLACEHOLDER_REGEX.exec(fullText)) !== null) {
        const rawTag = match[0];
        const mappedKey = placeholderToKey.get(rawTag) || extractPlaceholderKeyFromTag(rawTag);
        const val = (valueByKey[mappedKey] ?? '').trim();

        const start = match.index;
        const end = start + rawTag.length;

        const overlaps = replacements.some((r) => !(end <= r.start || start >= r.end));
        if (overlaps) continue;

        if (val) {
          const cleanVal = stripDuplicateUnitSuffix(mappedKey, val, fullText.slice(end));
          replacements.push({
            start,
            end,
            replacement: cleanVal,
            fieldKey: mappedKey,
          });
          const matchingMapping = mappings.find(
            (m) => m.mode === 'placeholder' && (m.placeholderTag === rawTag || m.fieldKey === mappedKey)
          );
          if (matchingMapping) matchedMappingIds.add(matchingMapping.id);
        } else {
          if (eraseUnfilledOldText) {
            replacements.push({
              start,
              end,
              replacement: '...............',
              fieldKey: mappedKey,
            });
          }
          missingSet.add(mappedKey);
        }
      }

      // 3. Global text_replace & erase_text mappings
      for (const m of textReplaceMappings) {
        const target = m.targetText!;
        const matches = findAllTextMatches(fullText, target, Boolean(m.matchCase));
        if (matches.length === 0) continue;

        const val = m.mode === 'erase_text' ? '' : (valueByKey[m.fieldKey] ?? '').trim();
        for (const found of matches) {
          const overlaps = replacements.some((r) => !(found.end <= r.start || found.start >= r.end));
          if (overlaps) continue;

          if (val) {
            const cleanVal = stripDuplicateUnitSuffix(m.fieldKey, val, fullText.slice(found.end));
            replacements.push({
              start: found.start,
              end: found.end,
              replacement: applyCaseLikeMatched(found.matched, cleanVal),
              fieldKey: m.fieldKey,
            });
            matchedMappingIds.add(m.id);
          } else if (m.mode === 'erase_text' || (eraseUnfilledOldText && !isBlankMarkerOnly(target))) {
            replacements.push({
              start: found.start,
              end: found.end,
              replacement: '',
              fieldKey: m.fieldKey,
            });
            if (m.mode !== 'erase_text') missingSet.add(m.fieldKey);
          } else {
            missingSet.add(m.fieldKey);
          }
        }
      }

      if (replacements.length === 0) continue;

      // Track paragraph anchors for optional automatic spouse insertion in main body (only when explicitly enabled)
      if (partName === 'word/document.xml' && !isParagraphInTableCell(pNode)) {
        const hasBodyInfoA = replacements.some((r) =>
          ['ben_ban_dia_chi', 'ben_ban_cccd', 'ben_ban_ngay_sinh'].includes(r.fieldKey)
        );
        const hasNameOnlyA = replacements.some((r) => r.fieldKey === 'ben_ban_ho_ten');
        if (hasBodyInfoA || (hasNameOnlyA && !partyAInfoAnchorP && fullText.trim().length > 25)) {
          partyAInfoAnchorP = pNode;
          partyADocRef = doc;
        }

        const hasBodyInfoB = replacements.some((r) =>
          ['ben_mua_dia_chi', 'ben_mua_cccd', 'ben_mua_ngay_sinh'].includes(r.fieldKey)
        );
        const hasNameOnlyB = replacements.some((r) => r.fieldKey === 'ben_mua_ho_ten');
        if (hasBodyInfoB || (hasNameOnlyB && !partyBInfoAnchorP && fullText.trim().length > 25)) {
          partyBInfoAnchorP = pNode;
          partyBDocRef = doc;
        }
      }

      // Also if it's a signature table paragraph containing ONLY ben_ban_ho_ten or ben_mua_ho_ten, append spouse name on new line
      if (partName === 'word/document.xml' && isParagraphInTableCell(pNode)) {
        const trimmedFull = fullText.trim();
        for (const rep of replacements) {
          if (
            rep.fieldKey === 'ben_ban_ho_ten' &&
            autoInsertSpousePartyA &&
            !templateHasPartyASpouse &&
            rep.end - rep.start >= trimmedFull.length - 2
          ) {
            const spouseName = (valueByKey.ben_ban_vo_chong_ho_ten ?? '').trim();
            if (spouseName && rep.replacement) {
              const formattedSpouse = isAllUpperCaseVi(rep.replacement)
                ? spouseName.toLocaleUpperCase('vi-VN')
                : spouseName;
              rep.replacement = `${rep.replacement}\n${formattedSpouse}`;
            }
          }
          if (
            rep.fieldKey === 'ben_mua_ho_ten' &&
            autoInsertSpousePartyB &&
            !templateHasPartyBSpouse &&
            rep.end - rep.start >= trimmedFull.length - 2
          ) {
            const spouseName = (valueByKey.ben_mua_vo_chong_ho_ten ?? '').trim();
            if (spouseName && rep.replacement) {
              const formattedSpouse = isAllUpperCaseVi(rep.replacement)
                ? spouseName.toLocaleUpperCase('vi-VN')
                : spouseName;
              rep.replacement = `${rep.replacement}\n${formattedSpouse}`;
            }
          }
        }
      }

      // Sort descending by start index so right-to-left edits don't shift earlier offsets
      replacements.sort((a, b) => b.start - a.start);
      const modifiedRuns = new Set<RunSpan>();

      for (const rep of replacements) {
        const firstRunIdx = runs.findIndex((r) => r.end > rep.start);
        let lastRunIdx = -1;
        for (let i = runs.length - 1; i >= 0; i--) {
          if (runs[i].start < rep.end) {
            lastRunIdx = i;
            break;
          }
        }
        if (firstRunIdx === -1 || lastRunIdx === -1 || lastRunIdx < firstRunIdx) continue;

        // Find actual editable text runs within the span (skipping tab/br separator entries if any)
        let firstEditableIdx = -1;
        let lastEditableIdx = -1;
        for (let i = firstRunIdx; i <= lastRunIdx; i++) {
          if (runs[i].tNode !== null) {
            if (firstEditableIdx === -1) firstEditableIdx = i;
            lastEditableIdx = i;
          }
        }
        if (firstEditableIdx === -1 || lastEditableIdx === -1) continue;

        const firstRun = runs[firstEditableIdx];
        const lastRun = runs[lastEditableIdx];

        const prefix = firstRun.text.slice(0, Math.max(0, rep.start - firstRun.start));
        const suffix = lastRun.text.slice(Math.max(0, rep.end - lastRun.start));

        if (firstEditableIdx === lastEditableIdx) {
          const updated = sanitizeXmlText(prefix + rep.replacement + suffix);
          firstRun.text = updated;
          if (firstRun.tNode) {
            firstRun.tNode.textContent = updated;
            firstRun.tNode.setAttribute('xml:space', 'preserve');
          }
          modifiedRuns.add(firstRun);
        } else {
          const updatedFirst = sanitizeXmlText(prefix + rep.replacement);
          firstRun.text = updatedFirst;
          if (firstRun.tNode) {
            firstRun.tNode.textContent = updatedFirst;
            firstRun.tNode.setAttribute('xml:space', 'preserve');
          }
          modifiedRuns.add(firstRun);

          for (let i = firstEditableIdx + 1; i < lastEditableIdx; i++) {
            const midRun = runs[i];
            if (midRun.tNode) {
              midRun.tNode.textContent = '';
              midRun.text = '';
            }
          }

          if (lastRun.tNode) {
            const safeSuffix = sanitizeXmlText(suffix);
            lastRun.tNode.textContent = safeSuffix;
            lastRun.tNode.setAttribute('xml:space', 'preserve');
            lastRun.text = safeSuffix;
          }
        }

        if (rep.replacement !== '') {
          filledCount++;
        }
        partModified = true;
      }

      // Expand any newline characters in modified runs into proper <w:br/> elements once per run
      for (const run of modifiedRuns) {
        if (run.tNode && run.text.includes('\n')) {
          setNodeTextWithNewlines(doc, run.tNode, run.rNode, run.text);
        }
      }
    }

    // Insert flexible spouse paragraph(s) right after Party A / Party B block in word/document.xml if needed
    if (partName === 'word/document.xml') {
      if (autoInsertSpousePartyA && !templateHasPartyASpouse && partyAInfoAnchorP && partyADocRef === doc) {
        const spouseLineA = buildSpouseInfoSentence(
          'ben_ban',
          valueByKey,
          options.customSpouseTextPartyA
        );
        if (spouseLineA && partyAInfoAnchorP.parentNode) {
          const newP = cloneParagraphWithText(doc, partyAInfoAnchorP, spouseLineA);
          if (partyAInfoAnchorP.nextSibling) {
            partyAInfoAnchorP.parentNode.insertBefore(newP, partyAInfoAnchorP.nextSibling);
          } else {
            partyAInfoAnchorP.parentNode.appendChild(newP);
          }
          filledCount++;
          partModified = true;
        }
      }

      if (autoInsertSpousePartyB && !templateHasPartyBSpouse && partyBInfoAnchorP && partyBDocRef === doc) {
        const spouseLineB = buildSpouseInfoSentence(
          'ben_mua',
          valueByKey,
          options.customSpouseTextPartyB
        );
        if (spouseLineB && partyBInfoAnchorP.parentNode) {
          const newP = cloneParagraphWithText(doc, partyBInfoAnchorP, spouseLineB);
          if (partyBInfoAnchorP.nextSibling) {
            partyBInfoAnchorP.parentNode.insertBefore(newP, partyBInfoAnchorP.nextSibling);
          } else {
            partyBInfoAnchorP.parentNode.appendChild(newP);
          }
          filledCount++;
          partModified = true;
        }
      }
    }

    if (partModified) {
      const updatedXml = serializer.serializeToString(doc);
      zip.file(partName, updatedXml);
    }
  }

  const outputBuffer = zip.generate({
    type: 'nodebuffer',
    compression: 'DEFLATE',
  }) as Buffer;

  return {
    outputBuffer,
    filledCount,
    totalMappings: Math.max(mappings.length, filledCount + missingSet.size),
    missingKeys: Array.from(missingSet),
  };
}
