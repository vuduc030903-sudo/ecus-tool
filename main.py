import sys
from pathlib import Path
from copy import copy
from datetime import datetime
import re
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell


# ===== CAU HINH =====
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
INPUT_FOLDER = BASE_DIR / "input" if (BASE_DIR / "input").exists() else BASE_DIR
TEMPLATE = BASE_DIR / "Template.xlsx"
OUTPUT = BASE_DIR / "output"

LABELS = [
    "Mode of Delivery:",
    "Forwarder/Consolidator:",
    "Terms of sale:",
    "Ship via:",
    "Container Stuffing Location:", BAO NGU :V
]

LABEL_COL = 1        # cot A (nhan)
VALUE_FIRST_COL = 3  # cot C (bat dau vung du lieu)
VALUE_LAST_COL = 9   # cot I (ket thuc vung du lieu)
SCAN_FROM_COL = 2    # quet tu cot B de khong bo sot vung merge bat dau tu B

# Cau hinh rieng cho tung nhan (nhan nao khong co thi dung mac dinh o tren)
#   label_col: cot chua nhan (None = tim o moi cot)
#   scan_from: cot bat dau quet gia tri
#   first    : cot ghi gia tri vao template
#   last     : cot cuoi vung gia tri
LABEL_CONFIG = {
    "Terms of sale:": {
        "label_col": None,   # K-Q: nhan co the nam o cot bat ky truoc cot K
        "scan_from": 11,     # K
        "first": 3,         # K
        "last": 17,          # Q
    },
}


# ===== CAU HINH BANG HANG HOA =====
IV_MARKER = "invoice no"           # sheet IV: o tieu de 'INVOICE NO.' (bang hang hoa lay o sheet nay)
PL_MARKER = "packing list no"      # sheet PL: o tieu de 'PACKING LIST NO.' (dong dong goi + KGS lay o sheet nay)
TABLE_HEADER_TEXT = "Aldila part"  # tieu de cot dau tien cua bang o file nguon
TABLE_STOP_TEXT = "total"          # dung khi gap dong co o = TOTAL
TABLE_FIRST_COL = 9                # cot I o file nguon (Aldila part, o tieu de merge I:J)
TABLE_LAST_COL = 17                # cot Q o file nguon ((2) Custom amount - Extension)
COPY_HEADER = False                # True = copy luon cac dong tieu de cua nguon
HEADER_ROWS_ABOVE = 1              # so dong tieu de nam tren dong 'Aldila part' (chi dung khi COPY_HEADER = True)

# Noi dan trong template
TARGET_SHEET = "Sheet2"            # ten sheet trong template
TARGET_START_ROW = 2               # dong dau tien dan du lieu (dong 1 la tieu de co san)
TARGET_START_COL = 1               # cot A
COPY_STYLE = False                 # False = giu nguyen dinh dang co san cua template

# Don cac dong cong thuc du (dong bi #N/A) sau khi dan du lieu
FORMULA_SHEETS = ["Sheet2", "Sheet3"]  # cac sheet co cong thuc can don dong du
EXTRA_ROWS_MODE = "hide"               # "hide" = an dong du (giu cong thuc) | "delete" = xoa han dong du

# O dich trong template cho dong dong goi + KGS: (ten sheet, o). De None = chua ghi.
# Vi du: "qty": ("Sheet1", "C30")
PACK_TARGETS = {
    "qty": ("Sheet1", "B16"),    # so luong, vi du 9
    "kind": ("Sheet1", "B17"),   # loai, vi du BOX / PLASTIC PALLET_big
    "kgs": ("Sheet1", "B18"),    # trong luong KGS, vi du 31.5 / 319.5
}

# Nhan dang dong dong goi: o chi chua dung 1 trong cac chu nay (khop dung ca o, khong nham voi 'BOXES')
PACK_EXACT_KEYWORDS = ("box", "bundle", "package")
# Luu y: script chi ghi CHU loai (BOX / BUNDLE / PACKAGE / PLASTIC PALLET...) vao o "kind" (B17).
# Ma don vi (BX, PP, PK) do cong thuc trong template (D17) tu doi tu chu do.


# O dich trong template cho so invoice + ngay ship (sheet IV). De None = chua ghi.
# Vi du: "invoice_no": ("Sheet1", "C2")
INFO_TARGETS = {
    "invoice_no": ("Sheet1", "B21"),   # vi du G-26-1040
    "ship_date": ("Sheet1", "B22"),    # vi du 16/10/2026
}
DATE_FORMAT = "dd/mm/yyyy"   # dinh dang ngay khi ghi vao template (None = giu dinh dang cua o template)

# O dich trong template cho 2 tong tien (dong TOTAL, sheet IV). De None = chua ghi.
TOTALS_TARGETS = {
    "processing_total": ("Sheet1", "B25"),   # tong tien gia cong (Processing Fee - Extension), vi du 154.80
    "custom_total": ("Sheet1", "B27"),       # tong tien hang (Custom amount - Extension), vi du 272.20
}

# O dich trong template cho Currency + Ship to Name (sheet IV). De None = chua ghi.
MISC_TARGETS = {
    "currency": ("Sheet1", "C12"),       # vi du USD
    "ship_to_name": ("Sheet1", "C13"),   # vi du ACUSHNET EUROPE LTD.
}

# Nho Excel tinh san cong thuc roi luu lai (can Excel + pywin32: pip install pywin32)
RECALC_WITH_EXCEL = True


def get_cfg(label):
    base = {
        "label_col": LABEL_COL,
        "scan_from": SCAN_FROM_COL,
        "first": VALUE_FIRST_COL,
        "last": VALUE_LAST_COL,
    }
    base.update(LABEL_CONFIG.get(label, {}))
    return base


def norm(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().rstrip(":").casefold()


def get_inv(filename):
    match = re.search(r"G-\d{2}-\d+", filename, re.I)
    return match.group(0).upper() if match else None


def merged_range(ws, cell):
    for r in ws.merged_cells.ranges:
        if cell.coordinate in r:
            return r
    return None


def real_cell(ws, row, col):
    """Neu o nam trong vung merge thi tra ve o goc tren-trai (noi chua gia tri)."""
    cell = ws.cell(row, col)
    rng = merged_range(ws, cell)
    if rng:
        return ws.cell(rng.min_row, rng.min_col)
    return cell


def find_all_labels(wb, label):
    """Tra ve danh sach (ws, cell) cua TAT CA o khop nhan, tren moi sheet."""
    cfg = get_cfg(label)
    target = norm(label)
    found = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cfg["label_col"] is not None and cell.column != cfg["label_col"]:
                    continue
                text = norm(cell.value)
                if text and (text == target or text.startswith(target)):
                    found.append((ws, cell))
    return found


def read_row_value(ws, label_cell, label):
    """Quet vung gia tri cua nhan tren cung dong, lay gia tri dau tien khac rong."""
    cfg = get_cfg(label)
    seen = set()
    for col in range(cfg["scan_from"], cfg["last"] + 1):
        cell = real_cell(ws, label_cell.row, col)
        if cell.coordinate in seen:
            continue
        seen.add(cell.coordinate)
        if cell.coordinate == label_cell.coordinate:
            continue
        if cell.value is not None and str(cell.value).strip() != "":
            return cell.value
    return None


def get_value(wb_values, wb_formulas, label):
    """Tra ve (value, found). Duyet moi sheet, lay nhan dau tien co du lieu."""
    cfg = get_cfg(label)
    matches = find_all_labels(wb_values, label)
    if not matches:
        return None, False

    for ws, cell in matches:
        value = read_row_value(ws, cell, label)
        if value is not None:
            print(f"    (lay tu {ws.title}!{cell.coordinate})")
            return value, True

    # Khong sheet nao co du lieu -> debug
    for ws, cell in matches:
        row_vals = [(c.coordinate, c.value) for c in ws[cell.row] if c.value is not None]
        print(f"    [DEBUG] {ws.title}!{cell.coordinate} -> cac o co du lieu trong dong: {row_vals}")

        # Kiem tra co phai cong thuc chua duoc Excel luu gia tri khong
        ws_f = wb_formulas[ws.title]
        formula_vals = [
            (c.coordinate, c.value)
            for c in ws_f[cell.row]
            if c.value is not None and c.column >= cfg["scan_from"]
        ]
        if formula_vals:
            print(f"    [CANH BAO] Dong nay co cong thuc/gia tri trong file goc: {formula_vals}")
            print("    -> Mo file nguon bang Excel, bam Ctrl+S roi chay lai.")
    return None, True


def put_value(wb, label, value):
    """Ghi vao cot dau cua vung gia tri, tren TAT CA sheet co nhan nay."""
    cfg = get_cfg(label)
    matches = find_all_labels(wb, label)
    if not matches:
        print(f"  Khong tim thay nhan trong template: {label}")
        return False

    for ws, cell in matches:
        target = real_cell(ws, cell.row, cfg["first"])
        target.value = value
        print(f"    -> ghi vao {ws.title}!{target.coordinate}")
    return True


def find_sheet(wb, marker, max_row=15):
    """Tim sheet co o tieu de khop marker (vi du 'INVOICE NO.' / 'PACKING LIST NO.') o ~15 dong dau."""
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=TABLE_LAST_COL):
            for cell in row:
                if norm(cell.value).rstrip(".").strip() == marker:
                    return ws
    return None


def find_text(wb, text, only_ws=None):
    """Tim o co noi dung bang text (khong phan biet hoa thuong). only_ws: gioi han trong 1 sheet."""
    target = norm(text)
    sheets = [only_ws] if only_ws is not None else wb.worksheets
    for ws in sheets:
        for row in ws.iter_rows():
            for cell in row:
                if norm(cell.value) == target:
                    return ws, cell
    return None, None


def is_total(value):
    """Nhan dang o TOTAL, ke ca dang '------TOTAL------' (bo dau gach, sao, khoang trang)."""
    return re.sub(r"[\s\-*_=]+", "", norm(value)) == "total"


def to_number(value):
    """Tra ve so neu o la so (ke ca chuoi '31.5', '1,234.5'), nguoc lai None."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return value
    try:
        return float(str(value).strip().replace(",", ""))
    except ValueError:
        return None


def find_total_row(ws, start_row):
    """Dong TOTAL dau tien (tu start_row tro xuong) trong cot A..Q."""
    for r in range(start_row, ws.max_row + 1):
        for c in range(1, TABLE_LAST_COL + 1):
            if is_total(ws.cell(r, c).value):
                return r
    return None


def has_part(ws, r):
    """Dong hang that su: o 'Aldila part' (cot I hoac J) co du lieu."""
    for c in (TABLE_FIRST_COL, TABLE_FIRST_COL + 1):
        v = real_cell(ws, r, c).value
        if v is not None and str(v).strip() != "":
            return True
    return False


def copy_table(source_wb, target_wb):
    """Copy bang hang hoa (cot I-Q) tu dong duoi tieu de 'Aldila part' den truoc dong TOTAL,
    dan vao TARGET_SHEET cua template, bat dau tu (TARGET_START_ROW, TARGET_START_COL).
    Chi lay dong co 'Aldila part' (bo dong trong / dong ghi chu nhu 'GOLF SHAFTS ****')."""
    iv_ws = find_sheet(source_wb, IV_MARKER)
    if iv_ws is None:
        print("  Khong thay sheet 'INVOICE NO.', tim bang o tat ca cac sheet")
    s_ws, s_head = find_text(source_wb, TABLE_HEADER_TEXT, iv_ws)
    if s_ws is None:
        print(f"  Khong tim thay tieu de bang nguon: {TABLE_HEADER_TEXT}")
        return False

    data_first = s_head.row + 1
    total_row = find_total_row(s_ws, data_first)
    if total_row is None:
        print("  Khong tim thay dong TOTAL trong file nguon")
        return False

    data_rows = [r for r in range(data_first, total_row) if has_part(s_ws, r)]
    if not data_rows:
        print("  Bang khong co dong du lieu nao")
        return False

    header_rows = []
    if COPY_HEADER:
        header_rows = list(range(s_head.row - HEADER_ROWS_ABOVE, data_first))

    src_rows = header_rows + data_rows
    row_map = {r: TARGET_START_ROW + i for i, r in enumerate(src_rows)}  # dong nguon -> dong dich
    col_offset = TARGET_START_COL - TABLE_FIRST_COL

    # Tim sheet dich trong template (khong phan biet hoa thuong)
    t_ws = None
    for name in target_wb.sheetnames:
        if name.strip().casefold() == TARGET_SHEET.strip().casefold():
            t_ws = target_wb[name]
            break
    if t_ws is None:
        print(f"  Template khong co sheet '{TARGET_SHEET}', tao moi")
        t_ws = target_wb.create_sheet(TARGET_SHEET)

    for r, tr in row_map.items():
        if COPY_STYLE:
            height = s_ws.row_dimensions[r].height
            if height:
                t_ws.row_dimensions[tr].height = height

        for c in range(TABLE_FIRST_COL, TABLE_LAST_COL + 1):
            src = s_ws.cell(r, c)
            if isinstance(src, MergedCell):
                continue
            dst = t_ws.cell(tr, c + col_offset)
            if isinstance(dst, MergedCell):
                continue

            dst.value = src.value
            if COPY_STYLE and src.has_style:
                dst.font = copy(src.font)
                dst.border = copy(src.border)
                dst.fill = copy(src.fill)
                dst.number_format = src.number_format
                dst.alignment = copy(src.alignment)

    # Copy cac vung merge trong bang (chi khi copy ca dinh dang)
    if COPY_STYLE:
        for rng in s_ws.merged_cells.ranges:
            if (
                rng.min_row in row_map and rng.max_row in row_map
                and rng.min_col >= TABLE_FIRST_COL and rng.max_col <= TABLE_LAST_COL
            ):
                try:
                    t_ws.merge_cells(
                        start_row=row_map[rng.min_row], start_column=rng.min_col + col_offset,
                        end_row=row_map[rng.max_row], end_column=rng.max_col + col_offset,
                    )
                except Exception:
                    pass

    n = len(src_rows)
    skipped = (total_row - data_first) - len(data_rows)
    print(f"  Da copy bang: {len(data_rows)} dong hang (nguon dong {data_first}-{total_row - 1}, "
          f"bo qua {skipped} dong trong/ghi chu) -> {t_ws.title}!A{TARGET_START_ROW}")
    return n


# ===== DONG DONG GOI + TRONG LUONG (CHI DOC, IN RA DE DOI CHIEU) =====
def is_pack_keyword(value):
    t = norm(value)
    return (
        t in PACK_EXACT_KEYWORDS         # khop dung ca o, de khong nham voi 'BOXES'
        or t.startswith("plastic pallet")
        or t.startswith("pallet consists")
        or t.startswith("pallet size")
    )


def find_total_weight(source_wb):
    """Tim o 'TOTAL WEIGHT' (cot A..Q) o sheet PL ('PACKING LIST NO.'); neu khong co sheet PL thi tim moi sheet."""
    pl_ws = find_sheet(source_wb, PL_MARKER)
    if pl_ws is None:
        print("  Khong thay sheet 'PACKING LIST NO.', tim TOTAL WEIGHT o tat ca cac sheet")
    sheets = [pl_ws] if pl_ws is not None else source_wb.worksheets
    for ws in sheets:
        for row in ws.iter_rows(max_col=TABLE_LAST_COL):
            for cell in row:
                if norm(cell.value).startswith("total weight"):
                    return ws, cell
    return None, None


def dump_pack_debug(source_wb, limit=40):
    """Khi khong tim thay: in cac o co chu total / weight / box / pallet de doi chieu."""
    print("  [DEBUG] Cac o co chu total / weight / box / pallet trong file nguon:")
    count = 0
    for ws in source_wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                t = norm(cell.value)
                if t and any(k in t for k in ("total", "weight", "box", "pallet")):
                    print(f"     {ws.title.strip()}!{cell.coordinate}: {str(cell.value)[:50]!r}")
                    count += 1
                    if count >= limit:
                        print("     ... (con nua)")
                        return


def extract_packing(source_wb):
    """Doc dong dong goi (so luong + loai) va KGS cua TOTAL WEIGHT. Tra ve dict, chua ghi vao template."""
    ws, tw_cell = find_total_weight(source_wb)
    if ws is None:
        return None

    tw_row, tw_col = tw_cell.row, tw_cell.column

    # Dong TOTAL gan nhat phia tren dong TOTAL WEIGHT
    total_row = None
    for r in range(tw_row - 1, 0, -1):
        if any(is_total(ws.cell(r, c).value) for c in range(1, TABLE_LAST_COL + 1)):
            total_row = r
            break
    if total_row is None:
        total_row = max(1, tw_row - 10)

    result = {"sheet": ws.title, "qty": None, "qty_cell": None, "kind": None,
              "kind_cell": None, "kgs": None, "kgs_cell": None}

    # --- Dong dong goi: trong khung tu dong TOTAL den truoc TOTAL WEIGHT, tren xuong, trai sang phai ---
    done = False
    for r in range(total_row, tw_row):
        for c in range(1, TABLE_LAST_COL + 1):
            cell = ws.cell(r, c)
            if not is_pack_keyword(cell.value):
                continue
            result["kind"] = str(cell.value).strip()
            result["kind_cell"] = cell.coordinate
            # So luong = o so gan nhat ben TRAI cung dong
            for cc in range(c - 1, 0, -1):
                left = real_cell(ws, r, cc)
                num = to_number(left.value)
                if num is not None:
                    result["qty"] = int(num) if float(num).is_integer() else num
                    result["qty_cell"] = left.coordinate
                    break
            done = True
            break
        if done:
            break

    # --- KGS: so cuoi cung ben phai tren dong TOTAL WEIGHT (LBS truoc, KGS sau) ---
    nums, seen = [], set()
    for c in range(tw_col + 1, TABLE_LAST_COL + 1):
        cell = real_cell(ws, tw_row, c)
        if cell.coordinate in seen:
            continue
        seen.add(cell.coordinate)
        num = to_number(cell.value)
        if num is not None:
            nums.append((num, cell.coordinate))
    if nums:
        result["kgs"], result["kgs_cell"] = nums[-1]
        if len(nums) >= 2 and nums[-1][0]:
            ratio = nums[-2][0] / nums[-1][0]
            if abs(ratio - 2.2046) > 0.05:
                print(f"    [CANH BAO] LBS/KGS = {ratio:.3f} (ky vong ~2.205) -> kiem tra lai: {nums}")
    else:
        # Co the la cong thuc chua duoc Excel luu gia tri
        raw = [c.coordinate for c in ws[tw_row] if c.value is not None and c.column > tw_col]
        print(f"    [DEBUG] Dong TOTAL WEIGHT ({ws.title.strip()}!{tw_cell.coordinate}) khong co so. "
              f"O co du lieu cung dong: {raw}")
    return result


def report_packing(source_wb):
    info = extract_packing(source_wb)
    if info is None:
        print("  [DONG GOI] Khong tim thay 'TOTAL WEIGHT' o sheet nao")
        dump_pack_debug(source_wb)
        return None
    sheet = info["sheet"].strip()
    if info["kind"] is None:
        print(f"  [DONG GOI] ({sheet}) Khong tim thay dong BOX / PALLET")
    else:
        print(f"  [DONG GOI] ({sheet}) so luong = {info['qty']} ({info['qty_cell']}), "
              f"loai = {info['kind']} ({info['kind_cell']})")
    if info["kgs"] is None:
        print(f"  [KGS] ({sheet}) Khong co so tren dong TOTAL WEIGHT")
    else:
        print(f"  [KGS] ({sheet}) {info['kgs']} ({info['kgs_cell']})")
    return info


def write_packing(target_wb, info):
    """Ghi so luong / loai / KGS vao cac o dich da khai bao trong PACK_TARGETS."""
    if not info:
        return
    if all(v is None for v in PACK_TARGETS.values()):
        print("  [DONG GOI] Chua khai bao o dich (PACK_TARGETS) nen chua ghi vao template")
        return

    for key, target in PACK_TARGETS.items():
        if target is None:
            continue
        value = info.get(key)
        if value is None:
            print(f"  [DONG GOI] Khong co gia tri '{key}' de ghi")
            continue

        sheet_name, addr = target
        ws = None
        for name in target_wb.sheetnames:
            if name.strip().casefold() == sheet_name.strip().casefold():
                ws = target_wb[name]
                break
        if ws is None:
            print(f"  [DONG GOI] Template khong co sheet '{sheet_name}'")
            continue

        cell = ws[addr]
        for rng in ws.merged_cells.ranges:       # o nam trong vung merge -> ghi vao o goc tren-trai
            if cell.coordinate in rng:
                cell = ws.cell(rng.min_row, rng.min_col)
                break
        cell.value = value
        print(f"    -> {key} = {value} ghi vao {ws.title.strip()}!{cell.coordinate}")


# ===== SO INVOICE + SHIP DATE (SHEET IV) =====
def find_cell_in(ws, text, max_col=TABLE_LAST_COL):
    """Tim o co chu khop text (bo qua hoa thuong, khoang trang thua, dau '.' va ':' o cuoi) trong 1 sheet."""
    for row in ws.iter_rows(max_col=max_col):
        for cell in row:
            if norm(cell.value).rstrip(".").strip() == text:
                return cell
    return None


def parse_date(value):
    """Giu nguyen neu la datetime; neu la chuoi thi thu doi sang ngay, khong duoc thi tra ve chuoi goc."""
    if isinstance(value, datetime) or value is None:
        return value
    text = str(value).strip()
    for fmt in ("%d/%m/%Y", "%d-%b-%y", "%d-%b-%Y", "%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return value


def extract_header_info(source_wb):
    """Doc so invoice (dong ngay duoi 'INVOICE NO.') va Ship Date (ben phai nhan 'Ship Date') o sheet IV."""
    iv_ws = find_sheet(source_wb, IV_MARKER)
    sheets = [iv_ws] if iv_ws is not None else source_wb.worksheets
    info = {"invoice_no": None, "invoice_cell": None, "ship_date": None, "ship_cell": None, "sheet": None}

    for ws in sheets:
        # --- So invoice: o nam duoi nhan (o co the merge rong, nen doc theo vung merge) ---
        if info["invoice_no"] is None:
            label = find_cell_in(ws, IV_MARKER)
            if label is not None:
                for dr in (1, 2):
                    r = label.row + dr
                    cell = real_cell(ws, r, label.column)
                    candidates = [cell] + [real_cell(ws, r, c) for c in range(1, 14)]
                    for cand in candidates:
                        if cand.value is not None and str(cand.value).strip():
                            info["invoice_no"] = str(cand.value).strip()
                            info["invoice_cell"] = cand.coordinate
                            break
                    if info["invoice_no"]:
                        break
                info["sheet"] = ws.title

        # --- Ship Date: o dau tien co du lieu ben phai nhan, cung dong ---
        if info["ship_date"] is None:
            label = find_cell_in(ws, "ship date")
            if label is not None:
                rng = merged_range(ws, label)
                start = (rng.max_col if rng else label.column) + 1
                seen = {label.coordinate}
                for c in range(start, TABLE_LAST_COL + 1):
                    cand = real_cell(ws, label.row, c)
                    if cand.coordinate in seen:
                        continue
                    seen.add(cand.coordinate)
                    if cand.value is not None and str(cand.value).strip():
                        info["ship_date"] = parse_date(cand.value)
                        info["ship_cell"] = cand.coordinate
                        break
                info["sheet"] = info["sheet"] or ws.title
    return info


def report_header(source_wb, inv=None):
    info = extract_header_info(source_wb)
    sheet = (info["sheet"] or "").strip()
    if info["invoice_no"] is None:
        print("  [INVOICE NO] Khong tim thay so invoice duoi nhan 'INVOICE NO.'")
    else:
        msg = f"  [INVOICE NO] ({sheet}) {info['invoice_no']} ({info['invoice_cell']})"
        if inv and info["invoice_no"].upper().replace(" ", "") != inv:
            msg += f"  [CANH BAO] khac so trong ten file ({inv})"
        print(msg)
    if info["ship_date"] is None:
        print("  [SHIP DATE] Khong tim thay ngay ben phai nhan 'Ship Date'")
    else:
        d = info["ship_date"]
        shown = d.strftime("%d/%m/%Y") if isinstance(d, datetime) else d
        print(f"  [SHIP DATE] ({sheet}) {shown} ({info['ship_cell']})")
    return info


def write_header(target_wb, info):
    """Ghi so invoice / ship date vao cac o dich da khai bao trong INFO_TARGETS."""
    if not info:
        return
    if all(v is None for v in INFO_TARGETS.values()):
        print("  [INVOICE/SHIP DATE] Chua khai bao o dich (INFO_TARGETS) nen chua ghi vao template")
        return

    for key, target in INFO_TARGETS.items():
        if target is None:
            continue
        value = info.get(key)
        if value is None:
            print(f"  [INVOICE/SHIP DATE] Khong co gia tri '{key}' de ghi")
            continue

        sheet_name, addr = target
        ws = None
        for name in target_wb.sheetnames:
            if name.strip().casefold() == sheet_name.strip().casefold():
                ws = target_wb[name]
                break
        if ws is None:
            print(f"  [INVOICE/SHIP DATE] Template khong co sheet '{sheet_name}'")
            continue

        cell = ws[addr]
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                cell = ws.cell(rng.min_row, rng.min_col)
                break
        cell.value = value
        if isinstance(value, datetime) and DATE_FORMAT:
            cell.number_format = DATE_FORMAT
        print(f"    -> {key} ghi vao {ws.title.strip()}!{cell.coordinate}")


# ===== TONG TIEN GIA CONG + TONG TIEN HANG (DONG TOTAL, SHEET IV) =====
def extract_totals(source_wb):
    """Doc 2 o tong tren dong TOTAL: Processing Fee - Extension va Custom amount - Extension."""
    iv_ws = find_sheet(source_wb, IV_MARKER)
    ws, head = find_text(source_wb, TABLE_HEADER_TEXT, iv_ws)
    if ws is None:
        return None
    total_row = find_total_row(ws, head.row + 1)
    if total_row is None:
        return None

    # Cot 'Extension' tim theo tieu de (2 cot: gia cong truoc, hang sau); khong thay thi dung O va Q
    ext_cols = set()
    for r in range(max(1, head.row - HEADER_ROWS_ABOVE), head.row + 1):
        for c in range(TABLE_FIRST_COL, TABLE_LAST_COL + 1):
            if norm(ws.cell(r, c).value) == "extension":
                ext_cols.add(c)
    ext_cols = sorted(ext_cols)
    if len(ext_cols) >= 2:
        proc_col, cust_col = ext_cols[0], ext_cols[-1]
    else:
        proc_col, cust_col = TABLE_FIRST_COL + 6, TABLE_LAST_COL   # O, Q

    info = {"sheet": ws.title}
    data_rows = [r for r in range(head.row + 1, total_row) if has_part(ws, r)]
    for key, col in (("processing_total", proc_col), ("custom_total", cust_col)):
        cell = real_cell(ws, total_row, col)
        info[key] = to_number(cell.value)
        info[key + "_cell"] = cell.coordinate
        # Doi chieu: cong cac dong hang voi so o dong TOTAL
        parts = [to_number(real_cell(ws, r, col).value) for r in data_rows]
        parts = [x for x in parts if x is not None]
        info[key + "_sum"] = round(sum(parts), 2) if parts else None
    return info


def report_totals(source_wb):
    info = extract_totals(source_wb)
    if info is None:
        print("  [TONG TIEN] Khong tim thay dong TOTAL cua bang hang hoa o sheet IV")
        return None
    sheet = info["sheet"].strip()
    for key, name in (("processing_total", "TONG GIA CONG"), ("custom_total", "TONG TIEN HANG")):
        val = info[key]
        if val is None:
            print(f"  [{name}] ({sheet}) o {info[key + '_cell']} khong co so "
                  "(neu la cong thuc: mo file nguon, Ctrl+S roi chay lai)")
            continue
        msg = f"  [{name}] ({sheet}) {val} ({info[key + '_cell']})"
        total_calc = info[key + "_sum"]
        if total_calc is not None and abs(total_calc - val) > 0.01:
            msg += f"  [CANH BAO] cong cac dong hang = {total_calc}, khac dong TOTAL"
        print(msg)
    return info


def write_totals(target_wb, info):
    """Ghi 2 tong tien vao cac o dich da khai bao trong TOTALS_TARGETS."""
    if not info:
        return
    if all(v is None for v in TOTALS_TARGETS.values()):
        print("  [TONG TIEN] Chua khai bao o dich (TOTALS_TARGETS) nen chua ghi vao template")
        return

    for key, target in TOTALS_TARGETS.items():
        if target is None:
            continue
        value = info.get(key)
        if value is None:
            print(f"  [TONG TIEN] Khong co gia tri '{key}' de ghi")
            continue

        sheet_name, addr = target
        ws = None
        for name in target_wb.sheetnames:
            if name.strip().casefold() == sheet_name.strip().casefold():
                ws = target_wb[name]
                break
        if ws is None:
            print(f"  [TONG TIEN] Template khong co sheet '{sheet_name}'")
            continue

        cell = ws[addr]
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                cell = ws.cell(rng.min_row, rng.min_col)
                break
        cell.value = value
        print(f"    -> {key} = {value} ghi vao {ws.title.strip()}!{cell.coordinate}")


# ===== CURRENCY + SHIP TO NAME (SHEET IV) =====
def extract_misc(source_wb):
    """Currency: gia tri ben phai nhan 'Currency'. Ship to name: dong dau cua o ngay duoi tieu de 'Ship to Name & Address'."""
    iv_ws = find_sheet(source_wb, IV_MARKER)
    sheets = [iv_ws] if iv_ws is not None else source_wb.worksheets
    info = {"currency": None, "currency_cell": None, "ship_to_name": None, "ship_to_cell": None, "sheet": None}

    for ws in sheets:
        # --- Currency ---
        if info["currency"] is None:
            for row in ws.iter_rows(max_col=TABLE_LAST_COL):
                for label in row:
                    if norm(label.value) != "currency":
                        continue
                    rng = merged_range(ws, label)
                    start = (rng.max_col if rng else label.column) + 1
                    seen = {label.coordinate}
                    for c in range(start, TABLE_LAST_COL + 1):
                        cand = real_cell(ws, label.row, c)
                        if cand.coordinate in seen:
                            continue
                        seen.add(cand.coordinate)
                        if cand.value is not None and str(cand.value).strip():
                            info["currency"] = str(cand.value).strip()
                            info["currency_cell"] = cand.coordinate
                            info["sheet"] = ws.title
                            break
                    if info["currency"] is not None:
                        break
                if info["currency"] is not None:
                    break

        # --- Ship to name: o ngay duoi tieu de 'Ship to Name & Address /Importer' ---
        if info["ship_to_name"] is None:
            for row in ws.iter_rows(max_col=TABLE_LAST_COL):
                for label in row:
                    if not norm(label.value).startswith("ship to name"):
                        continue
                    for dr in (1, 2):
                        cand = real_cell(ws, label.row + dr, label.column)
                        if cand.value is not None and str(cand.value).strip():
                            lines = [ln.strip() for ln in str(cand.value).splitlines() if ln.strip()]
                            info["ship_to_name"] = lines[0]
                            info["ship_to_cell"] = cand.coordinate
                            info["sheet"] = info["sheet"] or ws.title
                            break
                    if info["ship_to_name"] is not None:
                        break
                if info["ship_to_name"] is not None:
                    break
    return info


def report_misc(source_wb):
    info = extract_misc(source_wb)
    sheet = (info["sheet"] or "").strip()
    if info["currency"] is None:
        print("  [CURRENCY] Khong tim thay gia tri ben phai nhan 'Currency'")
    else:
        print(f"  [CURRENCY] ({sheet}) {info['currency']} ({info['currency_cell']})")
    if info["ship_to_name"] is None:
        print("  [SHIP TO NAME] Khong tim thay o duoi tieu de 'Ship to Name & Address'")
    else:
        print(f"  [SHIP TO NAME] ({sheet}) {info['ship_to_name']} ({info['ship_to_cell']})")
    return info


def write_misc(target_wb, info):
    """Ghi Currency / Ship to name vao cac o dich da khai bao trong MISC_TARGETS."""
    if not info:
        return
    if all(v is None for v in MISC_TARGETS.values()):
        print("  [CURRENCY/SHIP TO] Chua khai bao o dich (MISC_TARGETS) nen chua ghi vao template")
        return

    for key, target in MISC_TARGETS.items():
        if target is None:
            continue
        value = info.get(key)
        if value is None:
            print(f"  [CURRENCY/SHIP TO] Khong co gia tri '{key}' de ghi")
            continue

        sheet_name, addr = target
        ws = None
        for name in target_wb.sheetnames:
            if name.strip().casefold() == sheet_name.strip().casefold():
                ws = target_wb[name]
                break
        if ws is None:
            print(f"  [CURRENCY/SHIP TO] Template khong co sheet '{sheet_name}'")
            continue

        cell = ws[addr]
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                cell = ws.cell(rng.min_row, rng.min_col)
                break
        cell.value = value
        print(f"    -> {key} = {value} ghi vao {ws.title.strip()}!{cell.coordinate}")


def trim_extra_rows(target_wb, n_rows):
    """Sau khi dan n_rows dong du lieu, an hoac xoa cac dong cong thuc du (#N/A) ben duoi."""
    first_extra = TARGET_START_ROW + n_rows

    for name in FORMULA_SHEETS:
        ws = None
        for sheet_name in target_wb.sheetnames:
            if sheet_name.strip().casefold() == name.strip().casefold():
                ws = target_wb[sheet_name]
                break
        if ws is None:
            print(f"  Khong co sheet '{name}' de don dong du")
            continue

        last_row = ws.max_row
        if first_extra > last_row:
            print(f"  {ws.title}: khong co dong du de don "
                  f"(template chi co den dong {last_row}, du lieu den dong {first_extra - 1})")
            continue

        if EXTRA_ROWS_MODE == "delete":
            ws.delete_rows(first_extra, last_row - first_extra + 1)
            print(f"  {ws.title}: da xoa dong {first_extra}-{last_row}")
        else:
            for r in range(first_extra, last_row + 1):
                ws.row_dimensions[r].hidden = True
            print(f"  {ws.title}: da an dong {first_extra}-{last_row}")


def recalc_with_excel(paths):
    """Mo ngam tung file bang Excel, tinh lai cong thuc, luu lai -> file co san ket qua."""
    if not paths:
        return
    try:
        import win32com.client as win32
    except ImportError:
        print("\nChua cai pywin32 nen chua tinh san ket qua. Chay lenh: pip install pywin32")
        return

    print("\nDang nho Excel tinh cong thuc (chay ngam)...")
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        for path in paths:
            wb = None
            try:
                wb = excel.Workbooks.Open(str(Path(path).resolve()))
                excel.CalculateFull()
                wb.Save()
                print(f"  Da tinh xong: {Path(path).name}")
            except Exception as e:
                print(f"  LOI khi tinh {Path(path).name}: {e}")
            finally:
                if wb is not None:
                    wb.Close(SaveChanges=False)
    finally:
        if excel is not None:
            excel.Quit()


def main():
    if not TEMPLATE.exists():
        print(f"Khong tim thay file mau: {TEMPLATE}")
        return

    OUTPUT.mkdir(exist_ok=True)

    files = [
        f for f in INPUT_FOLDER.glob("IV + PL G-*.xlsx")
        if not f.name.startswith("~$")
        and f.resolve() != TEMPLATE.resolve()
    ]
    if not files and INPUT_FOLDER != BASE_DIR:
        files = [
            f for f in BASE_DIR.glob("IV + PL G-*.xlsx")
            if not f.name.startswith("~$")
            and f.resolve() != TEMPLATE.resolve()
        ]

    if not files:
        print(f"Khong tim thay file IV + PL G-*.xlsx trong {INPUT_FOLDER}")
        return

    created = []

    for source_path in files:
        inv = get_inv(source_path.name)

        if not inv:
            continue

        print(f"\nDang xu ly: {source_path.name}")

        try:
            source_vals = load_workbook(source_path, data_only=True)
            source_forms = load_workbook(source_path, data_only=False)
            target_wb = load_workbook(TEMPLATE)

            iv_s = find_sheet(source_vals, IV_MARKER)
            pl_s = find_sheet(source_vals, PL_MARKER)
            print(f"  Cac sheet nguon: {source_vals.sheetnames}")
            print(f"  Sheet IV: {iv_s.title.strip() if iv_s else 'KHONG THAY'} | "
                  f"Sheet PL: {pl_s.title.strip() if pl_s else 'KHONG THAY'}")

            for label in LABELS:
                value, found = get_value(source_vals, source_forms, label)

                if not found:
                    print(f"  Khong tim thay nhan nguon: {label}")
                    continue

                if value is None:
                    print(f"  O gia tri dang trong: {label}")
                    continue

                if put_value(target_wb, label, value):
                    print(f"  Da dien {label}: {value}")

            # Doc dong dong goi + KGS (chi in ra de doi chieu, chua ghi vao template)
            pack_info = report_packing(source_vals)
            head_info = report_header(source_vals, inv)
            totals_info = report_totals(source_vals)
            misc_info = report_misc(source_vals)

            # Copy bang hang hoa (Aldila part -> Extension, dung truoc TOTAL)
            n_rows = copy_table(source_vals, target_wb)
            if n_rows:
                trim_extra_rows(target_wb, n_rows)

            # Ghi dong dong goi + KGS vao template (neu da khai bao PACK_TARGETS)
            write_packing(target_wb, pack_info)
            write_header(target_wb, head_info)
            write_totals(target_wb, totals_info)
            write_misc(target_wb, misc_info)

            # De Excel tu tinh lai cong thuc khi mo file
            target_wb.calculation.fullCalcOnLoad = True

            output_path = OUTPUT / f"Template {inv}.xlsx"
            target_wb.save(output_path)
            created.append(output_path)

            source_vals.close()
            source_forms.close()
            target_wb.close()

            print(f"  Da tao file: {output_path}")

        except Exception as e:
            print(f"  LOI: {e}")

    if RECALC_WITH_EXCEL:
        recalc_with_excel(created)

    print("\nHoan tat!")


if __name__ == "__main__":
    main()
