export type FieldGroup = 'general' | 'party_a' | 'party_b' | 'property' | 'custom';

export interface FieldDefinition {
  key: string;
  label: string;
  group: FieldGroup;
  description?: string;
}

export const FIELD_GROUP_LABELS: Record<FieldGroup, string> = {
  general: 'Thông tin chung & Giao dịch',
  party_a: 'Bên bán / Chuyển nhượng / Để lại di sản',
  party_b: 'Bên mua / Nhận chuyển nhượng / Thừa kế',
  property: 'Thông tin Thửa đất / Sổ đỏ / Di sản',
  custom: 'Trường bổ sung theo mẫu Word',
};

export const STANDARD_FIELD_CATALOG: FieldDefinition[] = [
  // Thông tin chung & Giao dịch
  { key: 'so_hop_dong', label: 'Số hợp đồng / Văn bản', group: 'general' },
  { key: 'ngay_lap', label: 'Ngày lập văn bản', group: 'general' },
  { key: 'ngay_cong_chung', label: 'Ngày công chứng', group: 'general' },
  { key: 'van_phong_cong_chung', label: 'Văn phòng công chứng / UBND', group: 'general' },
  { key: 'dia_diem_ky', label: 'Địa điểm ký kết', group: 'general' },
  { key: 'gia_chuyen_nhuong', label: 'Giá chuyển nhượng (số)', group: 'general' },
  { key: 'gia_chuyen_nhuong_bang_chu', label: 'Giá chuyển nhượng (bằng chữ)', group: 'general' },
  { key: 'phuong_thuc_thanh_toan', label: 'Phương thức thanh toán', group: 'general' },

  // Bên A / Bên bán
  { key: 'ben_ban_ho_ten', label: 'Họ và tên (Bên A / Bên bán)', group: 'party_a' },
  { key: 'ben_ban_ngay_sinh', label: 'Ngày sinh (Bên A)', group: 'party_a' },
  { key: 'ben_ban_gioi_tinh', label: 'Giới tính (Bên A)', group: 'party_a' },
  { key: 'ben_ban_cccd', label: 'Số CCCD / CMND (Bên A)', group: 'party_a' },
  { key: 'ben_ban_ngay_cap', label: 'Ngày cấp CCCD (Bên A)', group: 'party_a' },
  { key: 'ben_ban_noi_cap', label: 'Nơi cấp CCCD (Bên A)', group: 'party_a' },
  { key: 'ben_ban_dien_thoai', label: 'Số điện thoại (Bên A)', group: 'party_a' },
  { key: 'ben_ban_dia_chi', label: 'Địa chỉ thường trú (Bên A)', group: 'party_a' },
  { key: 'ben_ban_thon_xom', label: 'Thôn / Xóm / Tổ (Bên A)', group: 'party_a' },
  { key: 'ben_ban_xa_phuong', label: 'Xã / Phường (Bên A)', group: 'party_a' },
  { key: 'ben_ban_tinh_thanh', label: 'Tỉnh / Thành phố (Bên A)', group: 'party_a' },
  { key: 'ty_le_so_huu_ben_ban', label: 'Tỷ lệ sở hữu % (Bên A)', group: 'party_a' },
  { key: 'ben_ban_vo_chong_ho_ten', label: 'Họ tên vợ/chồng (Bên A)', group: 'party_a' },
  { key: 'ben_ban_vo_chong_ngay_sinh', label: 'Ngày sinh vợ/chồng (Bên A)', group: 'party_a' },
  { key: 'ben_ban_vo_chong_gioi_tinh', label: 'Giới tính vợ/chồng (Bên A)', group: 'party_a' },
  { key: 'ben_ban_vo_chong_cccd', label: 'Số CCCD vợ/chồng (Bên A)', group: 'party_a' },
  { key: 'ben_ban_vo_chong_ngay_cap', label: 'Ngày cấp CCCD vợ/chồng (Bên A)', group: 'party_a' },
  { key: 'ben_ban_vo_chong_noi_cap', label: 'Nơi cấp CCCD vợ/chồng (Bên A)', group: 'party_a' },
  { key: 'ben_ban_vo_chong_dia_chi', label: 'Địa chỉ vợ/chồng (Bên A)', group: 'party_a' },

  // Bên B / Bên mua
  { key: 'ben_mua_ho_ten', label: 'Họ và tên (Bên B / Bên nhận)', group: 'party_b' },
  { key: 'ben_mua_ngay_sinh', label: 'Ngày sinh (Bên B)', group: 'party_b' },
  { key: 'ben_mua_gioi_tinh', label: 'Giới tính (Bên B)', group: 'party_b' },
  { key: 'ben_mua_cccd', label: 'Số CCCD / CMND (Bên B)', group: 'party_b' },
  { key: 'ben_mua_ngay_cap', label: 'Ngày cấp CCCD (Bên B)', group: 'party_b' },
  { key: 'ben_mua_noi_cap', label: 'Nơi cấp CCCD (Bên B)', group: 'party_b' },
  { key: 'ben_mua_dien_thoai', label: 'Số điện thoại (Bên B)', group: 'party_b' },
  { key: 'ben_mua_dia_chi', label: 'Địa chỉ thường trú (Bên B)', group: 'party_b' },
  { key: 'ben_mua_thon_xom', label: 'Thôn / Xóm / Tổ (Bên B)', group: 'party_b' },
  { key: 'ben_mua_xa_phuong', label: 'Xã / Phường (Bên B)', group: 'party_b' },
  { key: 'ben_mua_tinh_thanh', label: 'Tỉnh / Thành phố (Bên B)', group: 'party_b' },
  { key: 'ty_le_so_huu_ben_mua', label: 'Tỷ lệ sở hữu % (Bên B)', group: 'party_b' },
  { key: 'ben_mua_vo_chong_ho_ten', label: 'Họ tên vợ/chồng (Bên B)', group: 'party_b' },
  { key: 'ben_mua_vo_chong_ngay_sinh', label: 'Ngày sinh vợ/chồng (Bên B)', group: 'party_b' },
  { key: 'ben_mua_vo_chong_gioi_tinh', label: 'Giới tính vợ/chồng (Bên B)', group: 'party_b' },
  { key: 'ben_mua_vo_chong_cccd', label: 'Số CCCD vợ/chồng (Bên B)', group: 'party_b' },
  { key: 'ben_mua_vo_chong_ngay_cap', label: 'Ngày cấp CCCD vợ/chồng (Bên B)', group: 'party_b' },
  { key: 'ben_mua_vo_chong_noi_cap', label: 'Nơi cấp CCCD vợ/chồng (Bên B)', group: 'party_b' },
  { key: 'ben_mua_vo_chong_dia_chi', label: 'Địa chỉ vợ/chồng (Bên B)', group: 'party_b' },

  // Thông tin bổ sung ngày tháng & giao dịch
  { key: 'ngay_lap_ngay', label: 'Ngày lập (chỉ số ngày, vd: 05)', group: 'general' },
  { key: 'ngay_lap_thang', label: 'Tháng lập (chỉ số tháng, vd: 11)', group: 'general' },
  { key: 'ngay_lap_nam', label: 'Năm lập (4 số, vd: 2026)', group: 'general' },
  { key: 'nguon_goc_chuyen_giao', label: 'Hình thức chuyển giao (chuyển nhượng/tặng cho/thừa kế)', group: 'general' },

  // Thửa đất / Sổ đỏ / Di sản
  { key: 'thua_dat_so', label: 'Số thửa đất', group: 'property' },
  { key: 'to_ban_do_so', label: 'Số tờ bản đồ', group: 'property' },
  { key: 'dia_chi_thua_dat', label: 'Địa chỉ thửa đất (đầy đủ)', group: 'property' },
  { key: 'thua_dat_thon_xom', label: 'Thôn / Xóm / Tổ (Thửa đất)', group: 'property' },
  { key: 'thua_dat_xa_phuong', label: 'Xã / Phường / Đặc khu (Thửa đất)', group: 'property' },
  { key: 'thua_dat_tinh_thanh', label: 'Tỉnh / Thành phố (Thửa đất)', group: 'property' },
  { key: 'dien_tich', label: 'Diện tích (m²)', group: 'property' },
  { key: 'dien_tich_bang_chu', label: 'Diện tích (bằng chữ)', group: 'property' },
  { key: 'hinh_thuc_su_dung', label: 'Hình thức sử dụng', group: 'property' },
  { key: 'muc_dich_su_dung', label: 'Mục đích sử dụng đất', group: 'property' },
  { key: 'loai_dat_ky_hieu', label: 'Ký hiệu loại đất (vd: ONT, ODT, CLN)', group: 'property' },
  { key: 'thoi_han_su_dung', label: 'Thời hạn sử dụng', group: 'property' },
  { key: 'nguon_goc_su_dung', label: 'Nguồn gốc sử dụng', group: 'property' },
  { key: 'so_phat_hanh_gcn', label: 'Số phát hành Sổ đỏ (Seri)', group: 'property' },
  { key: 'so_vao_so_gcn', label: 'Số vào sổ cấp GCN', group: 'property' },
  { key: 'ngay_cap_gcn', label: 'Ngày cấp GCN', group: 'property' },
  { key: 'co_quan_cap_gcn', label: 'Cơ quan cấp GCN', group: 'property' },
  { key: 'tai_san_gan_lien', label: 'Tài sản gắn liền với đất', group: 'property' },
  { key: 'nguoi_de_lai_di_san', label: 'Người để lại di sản', group: 'property' },
  { key: 'ngay_mat_nguoi_de_lai', label: 'Ngày mất / Giấy chứng tử', group: 'property' },
  { key: 'noi_dung_di_san', label: 'Thông tin phân chia di sản', group: 'property' },
];

export interface TemplateFieldMapping {
  id: string;
  fieldKey: string;
  label: string;
  group: FieldGroup;
  mode: 'placeholder' | 'anchor_replace' | 'text_replace' | 'erase_text' | 'insert_after' | 'cell_fill';
  placeholderTag?: string;
  partName?: string;
  paragraphIndex?: number;
  targetText?: string;
  occurrenceInParagraph?: number;
  matchCase?: boolean;
  sampleContext?: string;
}

export interface DocxParagraphInspect {
  partName: string;
  paragraphIndex: number;
  locationLabel: string;
  text: string;
  detectedPlaceholders: string[];
  detectedBlanks: string[];
}

export interface WordTemplate {
  id: string;
  name: string;
  originalFileName: string;
  category: string;
  fileNameOnDisk: string;
  sizeBytes: number;
  placeholdersDetected: string[];
  mappings: TemplateFieldMapping[];
  fileMissing?: boolean;
  createdAt: string;
  updatedAt: string;
}

export type DocumentTypeHint =
  | 'cccd'
  | 'so_do'
  | 'hop_dong_chuyen_nhuong'
  | 'hop_dong_di_san'
  | 'pdf'
  | 'other';

export interface UploadedDoc {
  id: string;
  originalName: string;
  mimeType: string;
  sizeBytes: number;
  docTypeHint: DocumentTypeHint;
  fileNameOnDisk: string;
  uploadedAt: string;
}

export interface ExtractedFieldValue {
  key: string;
  label: string;
  group: FieldGroup;
  value: string;
  sourceNote?: string;
}

export interface GeneratedOutput {
  templateId: string;
  templateName: string;
  outputFileName: string;
  generatedAt: string;
  filledCount: number;
  totalMappings: number;
  missingKeys: string[];
}

export type CaseStatus = 'draft' | 'extracted' | 'verified' | 'exported';

export interface CaseFillOptions {
  eraseUnfilledOldText: boolean;
  autoInsertSpousePartyA: boolean;
  autoInsertSpousePartyB: boolean;
  customSpouseTextPartyA?: string;
  customSpouseTextPartyB?: string;
}

export interface CustomerCase {
  id: string;
  title: string;
  customerName: string;
  status: CaseStatus;
  selectedTemplateIds: string[];
  documents: UploadedDoc[];
  extractedFields: Record<string, ExtractedFieldValue>;
  fillOptions?: CaseFillOptions;
  aiSummary?: string;
  isVerified: boolean;
  generatedOutputs: GeneratedOutput[];
  createdAt: string;
  updatedAt: string;
}

export function normalizeKeyToVietnameseLabel(key: string): { label: string; group: FieldGroup } {
  const found = STANDARD_FIELD_CATALOG.find((f) => f.key === key);
  if (found) return { label: found.label, group: found.group };

  const clean = key
    .replace(/^[{\[]+|[}\]]+$/g, '')
    .replace(/_/g, ' ')
    .trim();
  const capitalized = clean.charAt(0).toUpperCase() + clean.slice(1);
  return { label: capitalized || key, group: 'custom' };
}

export function isLegacyNumericBracketMapping(m: TemplateFieldMapping): boolean {
  return m.mode === 'placeholder' && /^\[\d+\]$/.test((m.placeholderTag ?? '').trim());
}

function parseVietnameseAddressParts(raw: string): { thon: string; xa: string; tinh: string } {
  const cleaned = raw.trim().replace(/\)+$/g, '').trim();
  if (!cleaned) return { thon: '', xa: '', tinh: '' };

  // Prefer updated administrative name inside "(nay là ...)" if present
  const nayLaMatch = cleaned.match(/\(?\s*nay\s+là\s+([^)]+)\)?/i);
  const target = nayLaMatch ? nayLaMatch[1].trim() : cleaned;

  const segments = target
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);

  let thon = '';
  let xa = '';
  let tinh = '';

  for (const seg of segments) {
    if (/^(tỉnh|tp\.?|thành\s*phố)\s+/i.test(seg)) {
      tinh = seg.replace(/^(tỉnh|tp\.?|thành\s*phố)\s+/i, '').trim();
    } else if (/^(xã|phường|thị\s*trấn|đặc\s*khu)\s+/i.test(seg)) {
      xa = seg.replace(/^(xã|phường|thị\s*trấn|đặc\s*khu)\s+/i, '').trim();
    } else if (/^(thôn|xóm|tổ(\s+dân\s+phố)?|khu(\s+phố)?|ấp|bản|làng)\s+/i.test(seg)) {
      thon = seg;
    }
  }

  // Fallback if "(nay là Nga An, tỉnh Thanh Hoá)" omitted prefix "xã"
  if (!xa && nayLaMatch && segments.length >= 2) {
    const first = segments[0];
    if (!/^(huyện|quận|thị\s*xã|tỉnh|tp)/i.test(first)) {
      xa = first;
    }
  }

  return { thon, xa, tinh };
}

export function deriveGranularFields(
  fields: Record<string, ExtractedFieldValue>
): Record<string, ExtractedFieldValue> {
  const next: Record<string, ExtractedFieldValue> = { ...fields };

  const setIfEmpty = (key: string, value: string, sourceNote = 'Tự tách từ dữ liệu đã đọc') => {
    const val = value.trim();
    if (!val) return;
    if (next[key] && next[key].value.trim()) return;
    const meta = normalizeKeyToVietnameseLabel(key);
    next[key] = {
      key,
      label: next[key]?.label || meta.label,
      group: next[key]?.group || meta.group,
      value: val,
      sourceNote,
    };
  };

  // Address sub-parts
  if (next.ben_ban_dia_chi?.value) {
    const p = parseVietnameseAddressParts(next.ben_ban_dia_chi.value);
    setIfEmpty('ben_ban_thon_xom', p.thon);
    setIfEmpty('ben_ban_xa_phuong', p.xa);
    setIfEmpty('ben_ban_tinh_thanh', p.tinh);
  }
  if (next.ben_mua_dia_chi?.value) {
    const p = parseVietnameseAddressParts(next.ben_mua_dia_chi.value);
    setIfEmpty('ben_mua_thon_xom', p.thon);
    setIfEmpty('ben_mua_xa_phuong', p.xa);
    setIfEmpty('ben_mua_tinh_thanh', p.tinh);
  }
  if (next.dia_chi_thua_dat?.value) {
    const p = parseVietnameseAddressParts(next.dia_chi_thua_dat.value);
    setIfEmpty('thua_dat_thon_xom', p.thon);
    setIfEmpty('thua_dat_xa_phuong', p.xa);
    setIfEmpty('thua_dat_tinh_thanh', p.tinh);
  }

  // Date sub-parts from ngay_lap or ngay_cong_chung
  const dateSource = (next.ngay_lap?.value || next.ngay_cong_chung?.value || '').trim();
  if (dateSource) {
    const mLong = dateSource.match(/ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})/i);
    const mSlash = dateSource.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    const m = mLong || mSlash;
    if (m) {
      setIfEmpty('ngay_lap_ngay', m[1].padStart(2, '0'));
      setIfEmpty('ngay_lap_thang', m[2].padStart(2, '0'));
      setIfEmpty('ngay_lap_nam', m[3]);
    }
    if (!next.ngay_cong_chung?.value && next.ngay_lap?.value) {
      if (m) {
        setIfEmpty('ngay_cong_chung', `${m[1].padStart(2, '0')}/${m[2].padStart(2, '0')}/${m[3]}`);
      }
    }
  }

  // Land type abbreviation
  const landText = `${next.thoi_han_su_dung?.value ?? ''} ${next.muc_dich_su_dung?.value ?? ''}`;
  const abbrMatch = landText.match(/\b(ONT|ODT|CLN|LUC|LUK|RSX|SKC|BHK)\b/);
  if (abbrMatch) {
    setIfEmpty('loai_dat_ky_hieu', abbrMatch[1]);
  } else if (/đất ở tại nông thôn|đất ở nông thôn/i.test(landText)) {
    setIfEmpty('loai_dat_ky_hieu', 'ONT');
  } else if (/đất ở tại đô thị|đất ở đô thị/i.test(landText)) {
    setIfEmpty('loai_dat_ky_hieu', 'ODT');
  }

  // Default transfer origin if price exists
  if (next.gia_chuyen_nhuong?.value) {
    setIfEmpty('nguon_goc_chuyen_giao', 'nhận chuyển nhượng');
  }

  // Default ownership ratio
  if (next.ben_ban_ho_ten?.value) {
    setIfEmpty('ty_le_so_huu_ben_ban', '100%');
  }
  if (next.ben_mua_ho_ten?.value) {
    const hasSpouse = Boolean(next.ben_mua_vo_chong_ho_ten?.value?.trim());
    setIfEmpty('ty_le_so_huu_ben_mua', hasSpouse ? '50%' : '100%');
  }

  return next;
}
