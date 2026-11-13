import fs from 'node:fs';
import { GoogleGenAI, Type } from '@google/genai';
import type {
  DocxParagraphInspect,
  ExtractedFieldValue,
  FieldDefinition,
  FieldGroup,
  TemplateFieldMapping,
  UploadedDoc,
} from '../src/shared/types.ts';
import {
  STANDARD_FIELD_CATALOG,
  deriveGranularFields,
  normalizeKeyToVietnameseLabel,
} from '../src/shared/types.ts';

function getAiClient(): GoogleGenAI {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    throw new Error('Thiếu cấu hình GEMINI_API_KEY trên máy chủ.');
  }
  return new GoogleGenAI({
    apiKey,
    httpOptions: {
      headers: {
        'User-Agent': 'aistudio-build',
      },
    },
  });
}

export interface ExtractionResult {
  summary: string;
  suggestedCustomerName: string;
  fields: Record<string, ExtractedFieldValue>;
}

const DOC_HINT_LABELS: Record<string, string> = {
  cccd: 'Căn cước công dân (CCCD)',
  so_do: 'Sổ đỏ / Giấy chứng nhận QSDĐ',
  hop_dong_chuyen_nhuong: 'Hợp đồng chuyển nhượng',
  hop_dong_di_san: 'Hợp đồng / Văn bản phân chia di sản',
  pdf: 'Tài liệu PDF hồ sơ',
  other: 'Giấy tờ hồ sơ khác',
};

export async function extractCustomerDocuments(
  docs: Array<{ doc: UploadedDoc; absolutePath: string }>,
  templateTargetFields: FieldDefinition[]
): Promise<ExtractionResult> {
  if (docs.length === 0) {
    throw new Error('Chưa có giấy tờ nào được tải lên để phân tích.');
  }

  const ai = getAiClient();

  // Merge standard catalog + template-specific custom fields
  const fieldMap = new Map<string, FieldDefinition>();
  for (const f of STANDARD_FIELD_CATALOG) {
    fieldMap.set(f.key, f);
  }
  for (const tf of templateTargetFields) {
    if (!fieldMap.has(tf.key)) {
      fieldMap.set(tf.key, tf);
    }
  }
  const allFields = Array.from(fieldMap.values());

  const parts: Array<{ text: string } | { inlineData: { mimeType: string; data: string } }> = [];

  for (let i = 0; i < docs.length; i++) {
    const { doc, absolutePath } = docs[i];
    const fileBuf = fs.readFileSync(absolutePath);
    const base64 = fileBuf.toString('base64');
    const hintLabel = DOC_HINT_LABELS[doc.docTypeHint] || 'Giấy tờ hồ sơ';

    parts.push({
      text: `Tài liệu #${i + 1}: "${doc.originalName}" (Loại gợi ý: ${hintLabel}).`,
    });
    parts.push({
      inlineData: {
        mimeType: doc.mimeType || 'application/octet-stream',
        data: base64,
      },
    });
  }

  const fieldSpecText = allFields
    .map((f) => `- key: "${f.key}" | Tên trường: "${f.label}" | Nhóm: ${f.group}`)
    .join('\n');

  parts.push({
    text: `Hãy đọc toàn bộ ${docs.length} tài liệu trên của CÙNG MỘT HỒ SƠ KHÁCH HÀNG và trích xuất chính xác thông tin cho các trường sau đây:\n${fieldSpecText}\n\nQuy tắc bắt buộc:\n1. Chỉ trích xuất dữ liệu thực sự xuất hiện trong tài liệu. Không tự bịa hoặc suy đoán số CCCD, ngày tháng, số thửa, giá tiền nếu tài liệu không ghi.\n2. Nếu thông tin không có trong tài liệu, bỏ qua hoặc để chuỗi rỗng "".\n3. Phân biệt rõ Bên bán / Bên chuyển nhượng / Chủ sở hữu trên Sổ đỏ (party_a: ben_ban_*) với Bên mua / Bên nhận chuyển nhượng / Người thừa kế (party_b: ben_mua_*). Nếu trên hồ sơ có nhiều CCCD (ví dụ chồng và vợ hoặc 2 đồng sở hữu của Bên nhận / Bên bán), hãy trích xuất đầy đủ người thứ nhất vào ben_mua_* (hoặc ben_ban_*) và người vợ/chồng còn lại vào ben_mua_vo_chong_* (hoặc ben_ban_vo_chong_*): họ tên, ngày sinh, giới tính, số CCCD, ngày cấp, nơi cấp, địa chỉ thường trú.\n4. Nếu hồ sơ chỉ tải CCCD mới của Bên nhận/người thừa kế mà không có CCCD Bên bán, hãy xếp đúng vào vai trò thực tế (hoặc ben_mua_* nếu là người nhận/thừa kế).\n5. Chuẩn hóa ngày sinh, ngày cấp theo định dạng DD/MM/YYYY; riêng ngay_lap nếu hợp đồng ghi dạng chữ (ví dụ "ngày 13 tháng 11 năm 2026") thì giữ định dạng trang trọng, đồng thời tách thêm ngay_lap_ngay, ngay_lap_thang, ngay_lap_nam.\n6. Tách thêm các trường con địa chỉ nếu có: ben_ban_thon_xom, ben_ban_xa_phuong, ben_ban_tinh_thanh; ben_mua_thon_xom, ben_mua_xa_phuong, ben_mua_tinh_thanh; thua_dat_thon_xom, thua_dat_xa_phuong, thua_dat_tinh_thanh; và loai_dat_ky_hieu (vd: ONT, ODT, CLN), nguon_goc_chuyen_giao (vd: nhận chuyển nhượng / nhận tặng cho), ty_le_so_huu_ben_ban, ty_le_so_huu_ben_mua.\n7. Giữ đúng chính tả tiếng Việt có dấu, viết hoa tên riêng đúng chuẩn.\n8. Với mỗi trường trích xuất được, ghi rõ sourceNote ngắn gọn (vd: "CCCD chồng mặt trước", "CCCD vợ mặt sau", "Sổ đỏ trang 2").`,
  });

  const response = await ai.models.generateContent({
    model: 'gemini-4-hydrogen',
    contents: { parts },
    config: {
      systemInstruction:
        'Bạn là chuyên viên pháp lý và bất động sản Việt Nam chuyên đọc, đối chiếu thông tin từ CCCD, Giấy chứng nhận quyền sử dụng đất (Sổ đỏ/Sổ hồng), Hợp đồng chuyển nhượng, Văn bản phân chia di sản và hồ sơ PDF. Nhiệm vụ duy nhất của bạn là trích xuất dữ liệu có cấu trúc thật chính xác để người dùng kiểm tra và điền vào mẫu Word có sẵn.',
      responseMimeType: 'application/json',
      responseSchema: {
        type: Type.OBJECT,
        properties: {
          summary: {
            type: Type.STRING,
            description: 'Tóm tắt ngắn gọn bằng tiếng Việt về những giấy tờ đã đọc và các thông tin chính.',
          },
          suggestedCustomerName: {
            type: Type.STRING,
            description: 'Tên khách hàng đại diện cho hồ sơ (ví dụ tên bên mua hoặc bên bán chính).',
          },
          extractedItems: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                key: { type: Type.STRING },
                value: { type: Type.STRING },
                sourceNote: { type: Type.STRING },
              },
              required: ['key', 'value'],
            },
          },
        },
        required: ['summary', 'suggestedCustomerName', 'extractedItems'],
      },
    },
  });

  const rawText = (response.text ?? '{}').trim();
  const parsed = JSON.parse(rawText) as {
    summary?: string;
    suggestedCustomerName?: string;
    extractedItems?: Array<{ key: string; value: string; sourceNote?: string }>;
  };

  const fields: Record<string, ExtractedFieldValue> = {};

  // Populate all catalog + template fields (default empty if not extracted so user can edit any field)
  for (const def of allFields) {
    fields[def.key] = {
      key: def.key,
      label: def.label,
      group: def.group,
      value: '',
      sourceNote: '',
    };
  }

  for (const item of parsed.extractedItems ?? []) {
    const cleanKey = (item.key || '').trim();
    if (!cleanKey) continue;
    const val = (item.value ?? '').trim();
    if (!val) continue;

    const existingDef = fieldMap.get(cleanKey);
    const meta = existingDef
      ? { label: existingDef.label, group: existingDef.group as FieldGroup }
      : normalizeKeyToVietnameseLabel(cleanKey);

    fields[cleanKey] = {
      key: cleanKey,
      label: meta.label,
      group: meta.group,
      value: val,
      sourceNote: (item.sourceNote ?? '').trim(),
    };
  }

  const enrichedFields = deriveGranularFields(fields);

  return {
    summary: (parsed.summary || 'Đã hoàn tất đọc dữ liệu từ hồ sơ tải lên.').trim(),
    suggestedCustomerName: (parsed.suggestedCustomerName || '').trim(),
    fields: enrichedFields,
  };
}

export async function detectTemplateMappingsWithAI(
  paragraphs: Array<Pick<DocxParagraphInspect, 'partName' | 'paragraphIndex' | 'text'> & { locationLabel?: string }>
): Promise<TemplateFieldMapping[]> {
  if (paragraphs.length === 0) return [];

  const ai = getAiClient();
  const docPreview = paragraphs
    .map((p) => {
      const trimmed = p.text.trim();
      const body =
        trimmed === '' || trimmed === '.' || trimmed === '-'
          ? '[Ô TRỐNG]'
          : p.text;
      const loc = p.locationLabel ? ` (${p.locationLabel})` : '';
      return `[${p.partName}#${p.paragraphIndex}]${loc} ${body}`;
    })
    .join('\n')
    .slice(0, 45000);

  const catalogSpec = STANDARD_FIELD_CATALOG.map(
    (f) => `- "${f.key}": ${f.label} (${f.group})`
  ).join('\n');

  const response = await ai.models.generateContent({
    model: 'gemini-4-hydrogen',
    contents: `Dưới đây là toàn bộ nội dung các đoạn văn bản và ô bảng trích từ một file Word mẫu có sẵn (có thể là mẫu hợp đồng/tờ khai thuế cũ đã điền thông tin khách hàng trước đây, hoặc mẫu để trống dấu chấm/tab/ô bảng trống):\n\n${docPreview}\n\nDanh sách mã trường chuẩn:\n${catalogSpec}\n\nNhiệm vụ: Hãy rà soát TỪ TRÊN XUỐNG DƯỚI và liệt kê TẤT CẢ các vị trí dữ liệu biến thiên cần thay thế/điền mới (tên người bán/người chuyển nhượng cũ, vợ/chồng bên bán cũ, tên người mua/người nhận/người nộp thuế cũ, vợ/chồng bên mua cũ, ngày sinh, số CCCD, ngày cấp, nơi cấp, thôn/xóm, xã/phường, tỉnh/thành phố, số thửa, tờ bản đồ, diện tích, loại đất, seri Sổ đỏ, ngày cấp GCN, cơ quan cấp GCN, văn phòng công chứng, ngày lập/ngày công chứng, giá chuyển nhượng, các chỗ trống dấu chấm .... hoặc ô bảng trống [Ô TRỐNG]) sao cho KHÔNG làm thay đổi form hay cấu trúc câu chữ mặc định của mẫu.\n\nQuy tắc BẮT BUỘC:\n1. TUYỆT ĐỐI KHÔNG chọn mã chỉ tiêu của biểu mẫu như [01], [02], [04.1], [19.3], [20.1]..., không chọn tiêu đề biểu mẫu (Mẫu số 01/TK-SDDPNN, 03/BĐS-TNCN), tên cơ quan hướng dẫn, nhãn trường (như "Họ và tên:", "Số thửa đất:"), hay chú thích hướng dẫn.\n2. Các chế độ (mode) hợp lệ:\n   - "text_replace": Dùng cho chuỗi tên/địa chỉ/CCCD/cơ quan cũ ĐỦ DÀI (>= 6 ký tự, không trùng số pháp lý) để thay thế toàn văn bản (ví dụ: "Lê Văn Hảo", "Lê Văn Túy", "Đỗ Thị Cảnh", "038090038951", "AA 09851653", "Văn phòng công chứng Nguyễn Thị Giang", "chi nhánh VPĐK đất đai Thiệu Hóa"). KHÔNG gộp nhãn "Họ và tên:" vào "targetText".\n   - "anchor_replace": BẮT BUỘC dùng kèm đúng "partName", "paragraphIndex" và "occurrenceInParagraph" (tính từ 0) đối với:\n     + Các số ngắn, ngày tháng, tỷ lệ, diện tích (ví dụ: "1482", "42", "243,4", "308,15", "26", "8", "2026", "100%", "09/6/2026", "02/5/1990"). Lưu ý với "243,4 m2" hoặc "308,15m2" chỉ chọn phần số "243,4" hoặc "308,15" để giữ nguyên chữ "m2" của mẫu.\n     + Các chỗ trống dấu chấm/gạch chân (ví dụ: "…………………", "…………..", "........") gắn đúng trường cần điền.\n     + Tên thôn/xã/tỉnh nếu xuất hiện ở nhiều mục khác nhau (ví dụ vừa ở địa chỉ cư trú Bên mua vừa ở địa chỉ thửa đất) để tách đúng ben_mua_xa_phuong/ben_mua_tinh_thanh với thua_dat_xa_phuong/thua_dat_tinh_thanh.\n   - "insert_after": Dùng khi vị trí trống chỉ có nhãn kết thúc bằng dấu hai chấm + khoảng trắng/tab và KHÔNG có dấu chấm (ví dụ: "[18.3] Ngày lập: \\t" hoặc "[20.9] Giá trị đất thực tế chuyển giao (nếu có): \\tđồng"). Khi đó đặt "targetText" là đúng nhãn ngay trước chỗ trống kết thúc bằng ":" (ví dụ: "Ngày lập:" hoặc "Giá trị đất thực tế chuyển giao (nếu có):"), kèm "partName", "paragraphIndex", "occurrenceInParagraph": 0.\n   - "cell_fill": Dùng cho các ô bảng đang đánh dấu [Ô TRỐNG] thuộc dòng dữ liệu của Bên bán / Bên mua / Vợ chồng đồng sở hữu (ví dụ các ô [Ô TRỐNG] ở cột Ngày cấp, Nơi cấp, hoặc dòng số 2 dành cho vợ/chồng đồng nhận chuyển nhượng / đồng sử dụng đất). Đặt "partName", "paragraphIndex", và "targetText": "".\n3. Phân biệt vai trò theo mẫu hành chính Việt Nam:\n   - Với Tờ khai TNCN (Mẫu 03/BĐS-TNCN): Mục I (Người chuyển nhượng) dòng 1 là "ben_ban_*", dòng 2 là "ben_ban_vo_chong_*"; Mục II (Người nhận chuyển nhượng) dòng 1 là "ben_mua_*", dòng 2 là "ben_mua_vo_chong_*".\n   - Với Tờ khai thuế PNN (Mẫu 01/TK-SDDPNN): Mục [04]..[10] Người nộp thuế và tên ký cuối trang là người đứng tên kê khai mới ("ben_mua_*"), bảng [14] Người sử dụng đất là "ben_mua_*" (dòng 1) và "ben_mua_vo_chong_*" (dòng 2), mục [15.1]/[15.3] (người chuyển giao) là "ben_ban_ho_ten" / "ben_ban_cccd".\n4. CHỈ dùng các mode: "text_replace", "anchor_replace", "insert_after", "cell_fill". KHÔNG dùng "erase_text".`,
    config: {
      systemInstruction:
        'Bạn là chuyên gia phân tích biểu mẫu hành chính, thuế và hợp đồng đất đai Việt Nam. Bạn chỉ xác định chính xác từng vị trí giá trị biến thiên và ô trống cần điền vào file Word gốc.',
      responseMimeType: 'application/json',
      responseSchema: {
        type: Type.OBJECT,
        properties: {
          mappings: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                fieldKey: { type: Type.STRING },
                mode: { type: Type.STRING },
                targetText: { type: Type.STRING },
                partName: { type: Type.STRING },
                paragraphIndex: { type: Type.INTEGER },
                occurrenceInParagraph: { type: Type.INTEGER },
              },
              required: ['fieldKey', 'mode'],
            },
          },
        },
        required: ['mappings'],
      },
    },
  });

  const rawText = (response.text ?? '{}').trim();
  const parsed = JSON.parse(rawText) as {
    mappings?: Array<{
      fieldKey?: string;
      mode?: string;
      targetText?: string;
      partName?: string;
      paragraphIndex?: number;
      occurrenceInParagraph?: number;
    }>;
  };

  const paraMap = new Map<string, string>();
  for (const p of paragraphs) {
    paraMap.set(`${p.partName}#${p.paragraphIndex}`, p.text);
  }
  const allDocText = paragraphs.map((p) => p.text.replace(/\u00A0/g, ' ').toLocaleLowerCase('vi-VN')).join('\n');

  const results: TemplateFieldMapping[] = [];
  const dedup = new Set<string>();

  for (const item of parsed.mappings ?? []) {
    const fieldKey = (item.fieldKey ?? '').trim();
    const rawTarget = (item.targetText ?? '').trim();
    const rawMode = (item.mode ?? '').trim();

    if (!fieldKey || rawMode === 'erase_text' || fieldKey === 'old_info') continue;
    // Block form indicator codes like [01], [04.1], [20.9]
    if (/^\[\d+(\.\d+)?\]$/.test(rawTarget)) continue;

    const meta = normalizeKeyToVietnameseLabel(fieldKey);

    if (rawMode === 'cell_fill') {
      const partName = (item.partName || 'word/document.xml').trim();
      const paragraphIndex = typeof item.paragraphIndex === 'number' ? item.paragraphIndex : -1;
      const paraText = paraMap.get(`${partName}#${paragraphIndex}`);
      if (paraText === undefined) continue;
      const trimmedPara = paraText.trim();
      if (trimmedPara !== '' && trimmedPara !== '.' && trimmedPara !== '-') continue;

      const sig = `cell_fill|${partName}|${paragraphIndex}|${fieldKey}`;
      if (dedup.has(sig)) continue;
      dedup.add(sig);

      results.push({
        id: `map_ai_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 6)}`,
        fieldKey,
        label: meta.label,
        group: meta.group,
        mode: 'cell_fill',
        partName,
        paragraphIndex,
        targetText: '',
        sampleContext: '[Ô bảng trống]',
      });
      continue;
    }

    if (!rawTarget) continue;
    if (rawTarget.length < 2 && !/^\d+$/.test(rawTarget)) continue;

    if (rawMode === 'insert_after') {
      const partName = (item.partName || 'word/document.xml').trim();
      const paragraphIndex = typeof item.paragraphIndex === 'number' ? item.paragraphIndex : 0;
      const paraText = paraMap.get(`${partName}#${paragraphIndex}`);
      if (!paraText) continue;
      if (!paraText.replace(/\u00A0/g, ' ').includes(rawTarget.replace(/\u00A0/g, ' '))) continue;

      const occ =
        typeof item.occurrenceInParagraph === 'number' && item.occurrenceInParagraph >= 0
          ? item.occurrenceInParagraph
          : 0;
      const sig = `insert_after|${partName}|${paragraphIndex}|${occ}|${rawTarget}|${fieldKey}`;
      if (dedup.has(sig)) continue;
      dedup.add(sig);

      results.push({
        id: `map_ai_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 6)}`,
        fieldKey,
        label: meta.label,
        group: meta.group,
        mode: 'insert_after',
        partName,
        paragraphIndex,
        targetText: rawTarget,
        occurrenceInParagraph: occ,
        sampleContext: paraText.slice(0, 120),
      });
      continue;
    }

    const isAnchor =
      rawMode === 'anchor_replace' || rawTarget.length <= 5 || /^[.\u2026_]+$/.test(rawTarget);

    if (isAnchor) {
      const partName = (item.partName || 'word/document.xml').trim();
      const paragraphIndex = typeof item.paragraphIndex === 'number' ? item.paragraphIndex : 0;
      const paraText = paraMap.get(`${partName}#${paragraphIndex}`);
      if (!paraText) continue;

      const occ =
        typeof item.occurrenceInParagraph === 'number' && item.occurrenceInParagraph >= 0
          ? item.occurrenceInParagraph
          : 0;

      let resolvedTarget = rawTarget;
      const normPara = paraText.replace(/\u00A0/g, ' ');
      if (/^[.\u2026_\s-]+$/.test(rawTarget)) {
        const blankRuns = normPara.match(/(\.{3,}|…{2,}|_{3,})/g) ?? [];
        if (blankRuns[occ]) {
          resolvedTarget = blankRuns[occ];
        } else if (blankRuns.length > 0 && !normPara.includes(rawTarget.replace(/\u00A0/g, ' '))) {
          resolvedTarget = blankRuns[blankRuns.length - 1];
        }
      }
      if (!normPara.includes(resolvedTarget.replace(/\u00A0/g, ' '))) continue;

      const sig = `anchor|${partName}|${paragraphIndex}|${occ}|${resolvedTarget}|${fieldKey}`;
      if (dedup.has(sig)) continue;
      dedup.add(sig);

      results.push({
        id: `map_ai_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 6)}`,
        fieldKey,
        label: meta.label,
        group: meta.group,
        mode: 'anchor_replace',
        partName,
        paragraphIndex,
        targetText: resolvedTarget,
        occurrenceInParagraph: occ,
        sampleContext: paraText.slice(0, 120),
      });
    } else {
      const normTarget = rawTarget.replace(/\u00A0/g, ' ').toLocaleLowerCase('vi-VN');
      if (!allDocText.includes(normTarget)) continue;

      const mode = 'text_replace' as const;
      const sig = `${mode}|${normTarget}|${fieldKey}`;
      if (dedup.has(sig)) continue;
      dedup.add(sig);

      results.push({
        id: `map_ai_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 6)}`,
        fieldKey,
        label: meta.label,
        group: meta.group,
        mode,
        targetText: rawTarget,
        matchCase: false,
      });
    }
  }

  return results;
}

