import PizZip from 'pizzip';

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

interface ParagraphOpts {
  align?: 'left' | 'center' | 'right' | 'both';
  bold?: boolean;
  italic?: boolean;
  sizePt?: number;
  spaceBefore?: number;
  spaceAfter?: number;
}

function p(text: string, opts: ParagraphOpts = {}): string {
  const align = opts.align ?? 'left';
  const sz = (opts.sizePt ?? 13) * 2;
  const before = opts.spaceBefore ?? 40;
  const after = opts.spaceAfter ?? 80;
  const jc = align === 'left' ? 'left' : align;

  return `<w:p>
    <w:pPr>
      <w:spacing w:before="${before}" w:after="${after}" w:line="276" w:lineRule="auto"/>
      <w:jc w:val="${jc}"/>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        ${opts.bold ? '<w:b/>' : ''}
        ${opts.italic ? '<w:i/>' : ''}
        <w:sz w:val="${sz}"/>
        <w:szCs w:val="${sz}"/>
      </w:rPr>
      <w:t xml:space="preserve">${escapeXml(text)}</w:t>
    </w:r>
  </w:p>`;
}

function pRich(runs: Array<{ text: string; bold?: boolean; italic?: boolean }>, opts: ParagraphOpts = {}): string {
  const align = opts.align ?? 'left';
  const sz = (opts.sizePt ?? 13) * 2;
  const before = opts.spaceBefore ?? 40;
  const after = opts.spaceAfter ?? 80;

  const runsXml = runs
    .map(
      (r) => `<w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        ${r.bold ? '<w:b/>' : ''}
        ${r.italic ? '<w:i/>' : ''}
        <w:sz w:val="${sz}"/>
        <w:szCs w:val="${sz}"/>
      </w:rPr>
      <w:t xml:space="preserve">${escapeXml(r.text)}</w:t>
    </w:r>`
    )
    .join('');

  return `<w:p>
    <w:pPr>
      <w:spacing w:before="${before}" w:after="${after}" w:line="276" w:lineRule="auto"/>
      <w:jc w:val="${align}"/>
    </w:pPr>
    ${runsXml}
  </w:p>`;
}

function signatureTable(leftTitle: string, leftSub: string, rightTitle: string, rightSub: string): string {
  return `<w:tbl>
    <w:tblPr>
      <w:tblW w:w="9600" w:type="dxa"/>
      <w:tblBorders>
        <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/>
        <w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/>
      </w:tblBorders>
    </w:tblPr>
    <w:tblGrid><w:gridCol w:w="4800"/><w:gridCol w:w="4800"/></w:tblGrid>
    <w:tr>
      <w:tc>
        <w:tcPr><w:tcW w:w="4800" w:type="dxa"/></w:tcPr>
        ${p(leftTitle, { align: 'center', bold: true, sizePt: 13, spaceAfter: 20 })}
        ${p('(Ký, ghi rõ họ tên)', { align: 'center', italic: true, sizePt: 11, spaceAfter: 480 })}
        ${p(leftSub, { align: 'center', bold: true, sizePt: 13 })}
      </w:tc>
      <w:tc>
        <w:tcPr><w:tcW w:w="4800" w:type="dxa"/></w:tcPr>
        ${p(rightTitle, { align: 'center', bold: true, sizePt: 13, spaceAfter: 20 })}
        ${p('(Ký, ghi rõ họ tên)', { align: 'center', italic: true, sizePt: 11, spaceAfter: 480 })}
        ${p(rightSub, { align: 'center', bold: true, sizePt: 13 })}
      </w:tc>
    </w:tr>
  </w:tbl>`;
}

function infoTable(rows: Array<[string, string]>): string {
  const rowsXml = rows
    .map(
      ([label, val]) => `<w:tr>
      <w:tc>
        <w:tcPr><w:tcW w:w="3400" w:type="dxa"/></w:tcPr>
        ${p(label, { bold: true, sizePt: 12, spaceBefore: 40, spaceAfter: 40 })}
      </w:tc>
      <w:tc>
        <w:tcPr><w:tcW w:w="6200" w:type="dxa"/></w:tcPr>
        ${p(val, { sizePt: 12, spaceBefore: 40, spaceAfter: 40 })}
      </w:tc>
    </w:tr>`
    )
    .join('');

  return `<w:tbl>
    <w:tblPr>
      <w:tblW w:w="9600" w:type="dxa"/>
      <w:tblBorders>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="64748B"/>
        <w:left w:val="single" w:sz="4" w:space="0" w:color="64748B"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="64748B"/>
        <w:right w:val="single" w:sz="4" w:space="0" w:color="64748B"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
      </w:tblBorders>
    </w:tblPr>
    <w:tblGrid><w:gridCol w:w="3400"/><w:gridCol w:w="6200"/></w:tblGrid>
    ${rowsXml}
  </w:tbl>`;
}

function buildDocxBuffer(bodyXml: string): Buffer {
  const zip = new PizZip();

  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`
  );

  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );

  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
  );

  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
      <w:sz w:val="26"/>
    </w:rPr>
  </w:style>
</w:styles>`
  );

  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${bodyXml}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`
  );

  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }) as Buffer;
}

export interface SampleTemplateSeed {
  name: string;
  fileName: string;
  category: string;
  buffer: Buffer;
}

export function createStarterDocxTemplates(): SampleTemplateSeed[] {
  const hdChuyenNhuongBody = [
    p('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', { align: 'center', bold: true, sizePt: 13, spaceAfter: 20 }),
    p('Độc lập - Tự do - Hạnh phúc', { align: 'center', bold: true, sizePt: 13, spaceAfter: 160 }),
    p('HỢP ĐỒNG CHUYỂN NHƯỢNG QUYỀN SỬ DỤNG ĐẤT', { align: 'center', bold: true, sizePt: 15, spaceAfter: 40 }),
    p('Số: {{so_hop_dong}}', { align: 'center', italic: true, sizePt: 12, spaceAfter: 160 }),
    pRich(
      [
        { text: 'Hôm nay, {{ngay_lap}}, tại {{van_phong_cong_chung}}, chúng tôi gồm có:' },
      ],
      { align: 'both' }
    ),
    p('I. BÊN CHUYỂN NHƯỢNG (BÊN A):', { bold: true, spaceBefore: 120 }),
    pRich([
      { text: 'Ông/Bà: ', bold: true },
      { text: '{{ben_ban_ho_ten}}' },
      { text: '   —   Sinh ngày: ', bold: true },
      { text: '{{ben_ban_ngay_sinh}}' },
    ]),
    pRich([
      { text: 'Số CCCD: ', bold: true },
      { text: '{{ben_ban_cccd}}' },
      { text: '   —   Cấp ngày: ', bold: true },
      { text: '{{ben_ban_ngay_cap}}' },
      { text: '   —   Tại: ', bold: true },
      { text: '{{ben_ban_noi_cap}}' },
    ]),
    pRich([{ text: 'Địa chỉ thường trú: ', bold: true }, { text: '{{ben_ban_dia_chi}}' }]),
    p('II. BÊN NHẬN CHUYỂN NHƯỢNG (BÊN B):', { bold: true, spaceBefore: 120 }),
    pRich([
      { text: 'Ông/Bà: ', bold: true },
      { text: '{{ben_mua_ho_ten}}' },
      { text: '   —   Sinh ngày: ', bold: true },
      { text: '{{ben_mua_ngay_sinh}}' },
    ]),
    pRich([
      { text: 'Số CCCD: ', bold: true },
      { text: '{{ben_mua_cccd}}' },
      { text: '   —   Cấp ngày: ', bold: true },
      { text: '{{ben_mua_ngay_cap}}' },
      { text: '   —   Tại: ', bold: true },
      { text: '{{ben_mua_noi_cap}}' },
    ]),
    pRich([{ text: 'Địa chỉ thường trú: ', bold: true }, { text: '{{ben_mua_dia_chi}}' }]),
    p('III. THÔNG TIN THỬA ĐẤT CHUYỂN NHƯỢNG:', { bold: true, spaceBefore: 120, spaceAfter: 80 }),
    infoTable([
      ['Thửa đất số / Tờ bản đồ số', 'Thửa số {{thua_dat_so}}, Tờ bản đồ số {{to_ban_do_so}}'],
      ['Địa chỉ thửa đất', '{{dia_chi_thua_dat}}'],
      ['Diện tích', '{{dien_tich}} m² (Bằng chữ: {{dien_tich_bang_chu}})'],
      ['Hình thức & Mục đích sử dụng', '{{hinh_thuc_su_dung}} — {{muc_dich_su_dung}}'],
      ['Thời hạn sử dụng', '{{thoi_han_su_dung}}'],
      ['Giấy chứng nhận QSDĐ', 'Số phát hành: {{so_phat_hanh_gcn}}; Số vào sổ: {{so_vao_so_gcn}}; Cấp ngày: {{ngay_cap_gcn}} tại {{co_quan_cap_gcn}}'],
    ]),
    p('IV. GIÁ CHUYỂN NHƯỢNG VÀ PHƯƠNG THỨC THANH TOÁN:', { bold: true, spaceBefore: 140 }),
    pRich(
      [
        { text: 'Giá chuyển nhượng quyền sử dụng đất nêu trên là: ' },
        { text: '{{gia_chuyen_nhuong}} đồng', bold: true },
        { text: ' (Bằng chữ: {{gia_chuyen_nhuong_bang_chu}}).' },
      ],
      { align: 'both' }
    ),
    pRich([{ text: 'Phương thức thanh toán: {{phuong_thuc_thanh_toan}}.' }], { align: 'both' }),
    p('Hai bên đã đọc kỹ, hiểu rõ quyền và nghĩa vụ, đồng ý ký tên xác nhận dưới đây.', {
      align: 'both',
      spaceAfter: 160,
    }),
    signatureTable('BÊN CHUYỂN NHƯỢNG (BÊN A)', '{{ben_ban_ho_ten}}', 'BÊN NHẬN CHUYỂN NHƯỢNG (BÊN B)', '{{ben_mua_ho_ten}}'),
  ].join('');

  const phanChiaDiSanBody = [
    p('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', { align: 'center', bold: true, sizePt: 13, spaceAfter: 20 }),
    p('Độc lập - Tự do - Hạnh phúc', { align: 'center', bold: true, sizePt: 13, spaceAfter: 160 }),
    p('VĂN BẢN THỎA THUẬN PHÂN CHIA DI SẢN THỪA KẾ', { align: 'center', bold: true, sizePt: 15, spaceAfter: 40 }),
    p('Số công chứng: {{so_hop_dong}} — Ngày: {{ngay_cong_chung}}', { align: 'center', italic: true, sizePt: 12, spaceAfter: 140 }),
    pRich([
      { text: 'Tại {{van_phong_cong_chung}}, chúng tôi là những người thừa kế hợp pháp của Ông/Bà ' },
      { text: '{{nguoi_de_lai_di_san}}', bold: true },
      { text: ' (đã mất ngày {{ngay_mat_nguoi_de_lai}}), cùng nhau lập văn bản thỏa thuận phân chia di sản thừa kế với thông tin cụ thể sau:' },
    ], { align: 'both' }),
    p('1. Người đại diện / Người hưởng di sản:', { bold: true, spaceBefore: 100 }),
    pRich([
      { text: 'Ông/Bà: ', bold: true },
      { text: '{{ben_mua_ho_ten}}' },
      { text: ' — Sinh ngày: {{ben_mua_ngay_sinh}} — CCCD số: {{ben_mua_cccd}} (cấp ngày {{ben_mua_ngay_cap}} tại {{ben_mua_noi_cap}}).' },
    ]),
    pRich([{ text: 'Thường trú tại: ', bold: true }, { text: '{{ben_mua_dia_chi}}' }]),
    p('2. Di sản thừa kế là Quyền sử dụng đất:', { bold: true, spaceBefore: 100, spaceAfter: 80 }),
    infoTable([
      ['Thửa đất / Tờ bản đồ', 'Thửa số {{thua_dat_so}}, Tờ bản đồ số {{to_ban_do_so}}'],
      ['Địa chỉ thửa đất', '{{dia_chi_thua_dat}}'],
      ['Diện tích', '{{dien_tich}} m² ({{dien_tich_bang_chu}})'],
      ['Mục đích sử dụng', '{{muc_dich_su_dung}}'],
      ['Giấy chứng nhận QSDĐ', 'Seri: {{so_phat_hanh_gcn}}, Vào sổ: {{so_vao_so_gcn}}, Cấp ngày: {{ngay_cap_gcn}}'],
      ['Nội dung thỏa thuận phân chia', '{{noi_dung_di_san}}'],
    ]),
    p('Những người thừa kế cam kết chịu trách nhiệm trước pháp luật về tính chính xác của hồ sơ và tự nguyện ký tên.', {
      align: 'both',
      spaceBefore: 120,
      spaceAfter: 180,
    }),
    signatureTable('NGƯỜI LẬP VĂN BẢN / ĐẠI DIỆN', '{{ben_mua_ho_ten}}', 'CÔNG CHỨNG VIÊN', '{{van_phong_cong_chung}}'),
  ].join('');

  const donBienDongBody = [
    p('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', { align: 'center', bold: true, sizePt: 13, spaceAfter: 20 }),
    p('Độc lập - Tự do - Hạnh phúc', { align: 'center', bold: true, sizePt: 13, spaceAfter: 160 }),
    p('ĐƠN ĐĂNG KÝ BIẾN ĐỘNG ĐẤT ĐAI, TÀI SẢN GẮN LIỀN VỚI ĐẤT', { align: 'center', bold: true, sizePt: 14, spaceAfter: 40 }),
    p('Kính gửi: Văn phòng Đăng ký đất đai / Chi nhánh khu vực', { align: 'center', italic: true, sizePt: 12, spaceAfter: 140 }),
    p('1. Người sử dụng đất / Chủ sở hữu tài sản (Bên nhận biến động):', { bold: true }),
    pRich([
      { text: '- Họ và tên: ', bold: true },
      { text: '{{ben_mua_ho_ten}}' },
      { text: '  |  Ngày sinh: ', bold: true },
      { text: '{{ben_mua_ngay_sinh}}' },
    ]),
    pRich([
      { text: '- Số CCCD: ', bold: true },
      { text: '{{ben_mua_cccd}}' },
      { text: '  |  Ngày cấp: ', bold: true },
      { text: '{{ben_mua_ngay_cap}}' },
      { text: '  |  Nơi cấp: ', bold: true },
      { text: '{{ben_mua_noi_cap}}' },
    ]),
    pRich([{ text: '- Địa chỉ thường trú: ', bold: true }, { text: '{{ben_mua_dia_chi}}' }]),
    p('2. Thông tin thửa đất đăng ký biến động:', { bold: true, spaceBefore: 100 }),
    infoTable([
      ['Giấy chứng nhận đã cấp', 'Số phát hành: {{so_phat_hanh_gcn}} — Số vào sổ: {{so_vao_so_gcn}} — Cấp ngày: {{ngay_cap_gcn}}'],
      ['Thửa đất số / Tờ bản đồ số', 'Thửa {{thua_dat_so}}, Tờ bản đồ {{to_ban_do_so}}'],
      ['Địa chỉ thửa đất', '{{dia_chi_thua_dat}}'],
      ['Diện tích', '{{dien_tich}} m² — Mục đích: {{muc_dich_su_dung}}'],
      ['Cơ sở pháp lý biến động', 'Hợp đồng số {{so_hop_dong}} lập ngày {{ngay_cong_chung}} tại {{van_phong_cong_chung}}'],
      ['Bên chuyển quyền (Bên A)', '{{ben_ban_ho_ten}} (CCCD: {{ben_ban_cccd}})'],
    ]),
    p('Tôi xin cam đoan nội dung kê khai trên đơn là đúng sự thật và chịu trách nhiệm trước pháp luật.', {
      align: 'both',
      spaceBefore: 120,
      spaceAfter: 140,
    }),
    p('{{dia_diem_ky}}, {{ngay_lap}}', { align: 'right', italic: true, spaceAfter: 40 }),
    p('NGƯỜI LÀM ĐƠN', { align: 'right', bold: true, spaceAfter: 360 }),
    p('{{ben_mua_ho_ten}}', { align: 'right', bold: true }),
  ].join('');

  return [
    {
      name: 'Hợp đồng chuyển nhượng QSDĐ (Mẫu chuẩn)',
      fileName: 'Hop_dong_chuyen_nhuong_QSDD_mau.docx',
      category: 'Hợp đồng chuyển nhượng',
      buffer: buildDocxBuffer(hdChuyenNhuongBody),
    },
    {
      name: 'Văn bản thỏa thuận phân chia di sản thừa kế',
      fileName: 'Van_ban_phan_chia_di_san_thua_ke.docx',
      category: 'Di sản thừa kế',
      buffer: buildDocxBuffer(phanChiaDiSanBody),
    },
    {
      name: 'Đơn đăng ký biến động đất đai (Mẫu 09/ĐK)',
      fileName: 'Don_dang_ky_bien_dong_dat_dai.docx',
      category: 'Đơn đăng ký biến động',
      buffer: buildDocxBuffer(donBienDongBody),
    },
  ];
}
