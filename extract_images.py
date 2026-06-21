"""
Extract job images from xlsx, rename theo mã job (KL00002_avatar.jpg...).
Chạy: python3 extract_images.py "path/to/file.xlsx"
Output: thư mục job_images/ cạnh file xlsx
"""

import zipfile, re, os, sys
from pathlib import Path

def parse_drawing_map(drawing_xml, rels_xml):
    """drawing_row (0-indexed) -> image filename"""
    rid_to_file = {}
    for m in re.finditer(r'Id="(rId\d+)"[^>]*Target="[^"]*media/([^"]+)"', rels_xml):
        rid_to_file[m.group(1)] = m.group(2)

    row_to_file = {}
    for anchor in re.findall(r'<xdr:oneCellAnchor>[\s\S]*?</xdr:oneCellAnchor>', drawing_xml):
        rm = re.search(r'<xdr:row>(\d+)</xdr:row>', anchor)
        em = re.search(r'r:embed="(rId\d+)"', anchor)
        if rm and em and em.group(1) in rid_to_file:
            row_to_file[int(rm.group(1))] = rid_to_file[em.group(1)]
    return row_to_file

def parse_row_codes(xlsx_zip):
    """sheet row index (0-indexed in data array) -> job code"""
    import pandas as pd
    import io
    sheet_bytes = xlsx_zip.read('xl/worksheets/sheet1.xml')
    # Dùng pandas để đọc data (đơn giản hơn parse XML)
    xlsx_bytes = io.BytesIO()
    # Cần đọc lại file gốc qua pandas
    return None  # handled separately

def main(xlsx_path):
    xlsx_path = Path(xlsx_path)
    out_dir = xlsx_path.parent / 'job_images'
    out_dir.mkdir(exist_ok=True)

    import pandas as pd

    # 1. Đọc job codes từ sheet (pandas xử lý formula/shared strings tự động)
    df = pd.read_excel(xlsx_path, header=None)
    # Tìm header row có "Mã SP"
    header_row = None
    for i, row in df.iterrows():
        if 'Mã SP' in row.values:
            header_row = i
            break
    if header_row is None:
        print("Không tìm thấy header row có 'Mã SP'")
        sys.exit(1)

    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    # data array index 0 = xlsx row (header_row + 2), drawing row = header_row + 1
    # drawing_row cho data[0] = header_row + 1 (0-indexed)

    # 2. Parse image mapping từ xlsx zip
    with zipfile.ZipFile(xlsx_path) as z:
        names = z.namelist()

        # Tìm drawing file của sheet 1
        drawing_xml = rels_xml = None
        sheet_rels_path = 'xl/worksheets/_rels/sheet1.xml.rels'
        if sheet_rels_path in names:
            rels_content = z.read(sheet_rels_path).decode()
            dm = re.search(r'Target="\.\./drawings/(drawing\d+\.xml)"', rels_content)
            if dm:
                draw_key = 'xl/drawings/' + dm.group(1)
                rels_key = 'xl/drawings/_rels/' + dm.group(1) + '.rels'
                if draw_key in names:
                    drawing_xml = z.read(draw_key).decode()
                if rels_key in names:
                    rels_xml = z.read(rels_key).decode()

        if not drawing_xml or not rels_xml:
            print("Không tìm thấy drawing XML")
            sys.exit(1)

        row_to_img = parse_drawing_map(drawing_xml, rels_xml)

        # 3. Match và extract
        saved = 0
        skipped = 0
        for data_idx, row in df.iterrows():
            code = str(row.get('Mã SP', '') or '').strip()
            if not code:
                skipped += 1
                continue

            # drawing_row = header_row + 1 + data_idx
            drawing_row = header_row + 1 + data_idx
            img_name = row_to_img.get(drawing_row)

            if not img_name:
                print(f"  [{code}] Không có ảnh (drawing row {drawing_row})")
                skipped += 1
                continue

            img_path = 'xl/media/' + img_name
            if img_path not in names:
                print(f"  [{code}] File ảnh không tồn tại: {img_name}")
                skipped += 1
                continue

            ext = img_name.rsplit('.', 1)[-1].lower()
            out_name = f"{code}_avatar.{ext}"
            out_path = out_dir / out_name
            out_path.write_bytes(z.read(img_path))
            print(f"  [{code}] -> {out_name}")
            saved += 1

    print(f"\nXong! {saved} ảnh đã lưu vào: {out_dir}")
    print(f"Bỏ qua: {skipped} dòng")
    print(f"\nBước tiếp: upload thư mục '{out_dir.name}' lên Google Drive,")
    print("rồi chạy GAS function migrateImages('FOLDER_ID')")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python3 extract_images.py <path_to_xlsx>")
        sys.exit(1)
    main(sys.argv[1])
