import os
import io
import json
from datetime import datetime
from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.auth.transport.requests import Request

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

PARENT_FOLDER_ID = os.getenv("GDRIVE_PARENT_FOLDER_ID_2", "")

MONTH_SHEET_MAP = {
    1: "JAN", 2: "FEB", 3: "MAR", 4: "APR",
    5: "MAY", 6: "JUNE", 7: "JULY", 8: "AUG",
    9: "SEPT", 10: "OKT", 11: "NOV", 12: "DEC",
}


def _hex_to_rgb(hex_color: str) -> dict:
    h = hex_color.lstrip("#")
    return {
        "red": int(h[0:2], 16) / 255.0,
        "green": int(h[2:4], 16) / 255.0,
        "blue": int(h[4:6], 16) / 255.0,
    }


def _get_google_services():
    """Menggunakan OAuth token.json akun Google pribadi (punya kuota storage penuh)."""
    if os.getenv("GOOGLE_TOKEN_JSON"):
        info = json.loads(os.getenv("GOOGLE_TOKEN_JSON"))
        creds = Credentials.from_authorized_user_info(info, scopes=SCOPES)
    else:
        token_path = os.getenv("GOOGLE_TOKEN_FILE", "token_new.json")
        creds = Credentials.from_authorized_user_file(token_path, scopes=SCOPES)

    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        if not os.getenv("GOOGLE_TOKEN_JSON"):
            with open(os.getenv("GOOGLE_TOKEN_FILE", "token_new.json"), "w") as token_file:
                token_file.write(creds.to_json())

    drive_svc = build("drive", "v3", credentials=creds, cache_discovery=False)
    sheets_svc = build("sheets", "v4", credentials=creds, cache_discovery=False)
    return drive_svc, sheets_svc


def _get_or_create_subfolder(drive_svc, folder_name: str, parent_id: str) -> str:
    q = (
        f"name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' "
        f"and '{parent_id}' in parents and trashed = false"
    )
    res = drive_svc.files().list(q=q, fields="files(id, name)", pageSize=5).execute()
    files = res.get("files", [])
    if files:
        return files[0]["id"]

    meta = {
        "name": folder_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_id],
    }
    created = drive_svc.files().create(body=meta, fields="id").execute()
    return created["id"]


def upload_attachment_to_drive(file_storage, prefix: str, tanggal_str: str) -> str:
    if not file_storage or not file_storage.filename:
        return ""
    if not PARENT_FOLDER_ID:
        raise ValueError("GDRIVE_PARENT_FOLDER_ID_2 belum diisi di environment (.env).")

    drive_svc, _ = _get_google_services()
    attach_folder_id = _get_or_create_subfolder(drive_svc, "JFS_Attachments", PARENT_FOLDER_ID)
    ym_folder_id = _get_or_create_subfolder(drive_svc, tanggal_str[:7], attach_folder_id)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    clean_name = f"{prefix}_{stamp}_{file_storage.filename}"

    media = MediaIoBaseUpload(
        io.BytesIO(file_storage.read()),
        mimetype=file_storage.mimetype or "application/octet-stream",
        resumable=False,
    )
    meta = {"name": clean_name, "parents": [ym_folder_id]}
    uploaded = drive_svc.files().create(body=meta, media_body=media, fields="id, webViewLink").execute()

    drive_svc.permissions().create(
        fileId=uploaded["id"],
        body={"type": "anyone", "role": "reader"},
    ).execute()

    return uploaded.get("webViewLink", "")


def _get_or_create_jfs_spreadsheet(drive_svc, sheets_svc, year: int) -> dict:
    if not PARENT_FOLDER_ID:
        raise ValueError("GDRIVE_PARENT_FOLDER_ID_2 belum diisi di environment (.env).")

    title = f"JFS_LEDGER_{year}"
    q = (
        f"name = '{title}' and mimeType = 'application/vnd.google-apps.spreadsheet' "
        f"and '{PARENT_FOLDER_ID}' in parents and trashed = false"
    )
    res = drive_svc.files().list(q=q, fields="files(id, name, webViewLink)", pageSize=5).execute()
    files = res.get("files", [])
    if files:
        return {
            "spreadsheet_id": files[0]["id"],
            "title": files[0]["name"],
            "url": files[0].get("webViewLink", f"https://docs.google.com/spreadsheets/d/{files[0]['id']}/edit"),
        }

    file_meta = {
        "name": title,
        "mimeType": "application/vnd.google-apps.spreadsheet",
        "parents": [PARENT_FOLDER_ID],
    }
    created = drive_svc.files().create(body=file_meta, fields="id, name, webViewLink").execute()
    ss_id = created["id"]

    ss_info = sheets_svc.spreadsheets().get(spreadsheetId=ss_id).execute()
    default_sheet_id = ss_info["sheets"][0]["properties"]["sheetId"]
    init_month_name = MONTH_SHEET_MAP[datetime.now().month]

    sheets_svc.spreadsheets().batchUpdate(
        spreadsheetId=ss_id,
        body={
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {"sheetId": default_sheet_id, "title": init_month_name},
                        "fields": "title",
                    }
                }
            ]
        },
    ).execute()

    _initialize_modern_month_tab(sheets_svc, ss_id, default_sheet_id, init_month_name, year, saldo_awal=0.0)

    return {
        "spreadsheet_id": ss_id,
        "title": title,
        "url": created.get("webViewLink", f"https://docs.google.com/spreadsheets/d/{ss_id}/edit"),
    }


def _initialize_modern_month_tab(
    sheets_svc, ss_id: str, sheet_id: int, sheet_name: str, year: int, saldo_awal: float = 0.0
):
    """
    Membangun layout Modern Fintech Ledger di Google Sheets:
    - Baris 1-2: Brand Banner CV. Jaskurindo Mitra Abadi
    - Baris 4-5: 6 Kartu KPI Summary Dinamis (Saldo Awal, Debet, Kredit, Saldo Akhir, Omzet, Komisi 40%)
    - Baris 7: Header Tabel Bersih 15 Kolom (Freeze Pane)
    - Baris 8+: Data Transaksi
    """
    values = [
        # Row 1: Title Banner
        ["CV. JASKURINDO MITRA ABADI — OPERASIONAL J&T CARGO (JFS)", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        # Row 2: Subtitle
        [f"Buku Kas & Rekonsiliasi Harian JFS | Periode: {sheet_name} {year}", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        # Row 3: Spacer
        [""],
        # Row 4: KPI Card Labels
        ["SALDO AWAL JFS", "", "TOTAL DEBET (MASUK)", "", "TOTAL KREDIT (KELUAR)", "", "SALDO AKHIR JFS", "", "TOTAL OMZET (100%)", "", "KOMISI AGEN (40%)", "", "TOTAL PIUTANG (TEMPO)", "", ""],
        # Row 5: KPI Card Formulas (Col A = Saldo Awal input, sisanya rumus otomatis)
        [
            saldo_awal, "",
            "=SUM(E8:E1000)", "",
            "=SUM(F8:F1000)", "",
            "=A5+C5-E5", "",
            "=SUM(H8:H1000)", "",
            "=SUM(I8:I1000)", "",
            "=SUM(K8:K1000)", "", ""
        ],
        # Row 6: Spacer
        [""],
        # Row 7: Clean 15-Column Table Header
        [
            "NO",               # A (0)
            "TANGGAL",          # B (1)
            "KATEGORI",         # C (2)
            "KETERANGAN / RESI",# D (3)
            "DEBET (MASUK)",    # E (4)
            "KREDIT (KELUAR)",  # F (5)
            "SALDO JFS",        # G (6)
            "OMZET (100%)",     # H (7)
            "KOMISI (40%)",     # I (8)
            "PAID (LUNAS)",     # J (9)
            "BALANCE (TEMPO)",  # K (10)
            "STATUS",           # L (11)
            "METODE & CATATAN", # M (12)
            "BUKTI TRANSAKSI",  # N (13)
            "FOTO BARANG",      # O (14)
        ],
    ]

    sheets_svc.spreadsheets().values().update(
        spreadsheetId=ss_id,
        range=f"'{sheet_name}'!A1:O7",
        valueInputOption="USER_ENTERED",
        body={"values": values},
    ).execute()

    # Styling & Layout Requests via batchUpdate
    col_widths = [55, 110, 145, 290, 140, 140, 150, 140, 135, 135, 135, 110, 220, 130, 130]
    requests = [
        # 1. Freeze baris header 1..7 & tampilkan gridlines halus
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {"frozenRowCount": 7, "hideGridlines": False},
                },
                "fields": "gridProperties.frozenRowCount,gridProperties.hideGridlines",
            }
        },
        # 2. Font default seluruh sheet
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1000, "startColumnIndex": 0, "endColumnIndex": 15},
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {"fontFamily": "Inter", "fontSize": 10, "foregroundColor": _hex_to_rgb("#1E293B")},
                        "verticalAlignment": "MIDDLE",
                    }
                },
                "fields": "userEnteredFormat(textFormat,verticalAlignment)",
            }
        },
        # 3. Banner Baris 1-2 (Dark Slate / J&T Accent)
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 15},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": _hex_to_rgb("#0F172A"),
                        "textFormat": {"fontFamily": "Inter", "fontSize": 13, "bold": True, "foregroundColor": _hex_to_rgb("#FFFFFF")},
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat)",
            }
        },
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": 15},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": _hex_to_rgb("#1E293B"),
                        "textFormat": {"fontFamily": "Inter", "fontSize": 9, "foregroundColor": _hex_to_rgb("#CBD5E1")},
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat)",
            }
        },
        # 4. KPI Card Labels (Row 4) & Values (Row 5)
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 3, "endRowIndex": 4, "startColumnIndex": 0, "endColumnIndex": 14},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": _hex_to_rgb("#F1F5F9"),
                        "textFormat": {"fontSize": 8, "bold": True, "foregroundColor": _hex_to_rgb("#475569")},
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat)",
            }
        },
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 4, "endRowIndex": 5, "startColumnIndex": 0, "endColumnIndex": 14},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": _hex_to_rgb("#F8FAFC"),
                        "numberFormat": {"type": "CURRENCY", "pattern": "\"Rp \"#,##0"},
                        "textFormat": {"fontSize": 11, "bold": True, "foregroundColor": _hex_to_rgb("#0F172A")},
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,numberFormat,textFormat)",
            }
        },
        # 5. Table Header Row 7 (Merah J&T Cargo / Deep Navy)
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 6, "endRowIndex": 7, "startColumnIndex": 0, "endColumnIndex": 15},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": _hex_to_rgb("#B91C1C"),
                        "horizontalAlignment": "CENTER",
                        "textFormat": {"fontFamily": "Inter", "fontSize": 9, "bold": True, "foregroundColor": _hex_to_rgb("#FFFFFF")},
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)",
            }
        },
        # 6. Format Angka Rupiah untuk Kolom E..K (Index 4..10) mulai baris 8
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 7, "endRowIndex": 1000, "startColumnIndex": 4, "endColumnIndex": 11},
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {"type": "CURRENCY", "pattern": "\"Rp \"#,##0"},
                        "horizontalAlignment": "RIGHT",
                    }
                },
                "fields": "userEnteredFormat(numberFormat,horizontalAlignment)",
            }
        },
        # 7. Center Alignment untuk NO, TANGGAL, KATEGORI, STATUS, LINK
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 7, "endRowIndex": 1000, "startColumnIndex": 0, "endColumnIndex": 3},
                "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                "fields": "userEnteredFormat.horizontalAlignment",
            }
        },
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 7, "endRowIndex": 1000, "startColumnIndex": 11, "endColumnIndex": 12},
                "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER", "textFormat": {"bold": True}}},
                "fields": "userEnteredFormat(horizontalAlignment,textFormat)",
            }
        },
        {
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": 7, "endRowIndex": 1000, "startColumnIndex": 13, "endColumnIndex": 15},
                "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                "fields": "userEnteredFormat.horizontalAlignment",
            }
        },
        # 8. Tinggi baris header lebih lega
        {
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
                "properties": {"pixelSize": 36},
                "fields": "pixelSize",
            }
        },
        {
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": 4, "endIndex": 5},
                "properties": {"pixelSize": 32},
                "fields": "pixelSize",
            }
        },
        {
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": 6, "endIndex": 7},
                "properties": {"pixelSize": 34},
                "fields": "pixelSize",
            }
        },
        # 9. Conditional Formatting Badge Status LUNAS (Hijau) & TEMPO (Kuning) di Kolom L (index 11)
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{"sheetId": sheet_id, "startRowIndex": 7, "endRowIndex": 1000, "startColumnIndex": 11, "endColumnIndex": 12}],
                    "booleanRule": {
                        "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": "LUNAS"}]},
                        "format": {
                            "backgroundColor": _hex_to_rgb("#DCFCE7"),
                            "textFormat": {"foregroundColor": _hex_to_rgb("#166534"), "bold": True},
                        },
                    },
                },
                "index": 0,
            }
        },
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{"sheetId": sheet_id, "startRowIndex": 7, "endRowIndex": 1000, "startColumnIndex": 11, "endColumnIndex": 12}],
                    "booleanRule": {
                        "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": "TEMPO"}]},
                        "format": {
                            "backgroundColor": _hex_to_rgb("#FEF3C7"),
                            "textFormat": {"foregroundColor": _hex_to_rgb("#92400E"), "bold": True},
                        },
                    },
                },
                "index": 1,
            }
        },
    ]

    # Set lebar setiap kolom
    for idx, px in enumerate(col_widths):
        requests.append({
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": idx, "endIndex": idx + 1},
                "properties": {"pixelSize": px},
                "fields": "pixelSize",
            }
        })

    sheets_svc.spreadsheets().batchUpdate(spreadsheetId=ss_id, body={"requests": requests}).execute()


def _ensure_month_tab_exists(sheets_svc, ss_id: str, month: int, year: int) -> tuple[str, int]:
    sheet_name = MONTH_SHEET_MAP[month]
    ss_info = sheets_svc.spreadsheets().get(spreadsheetId=ss_id).execute()
    sheets_map = {s["properties"]["title"]: s["properties"]["sheetId"] for s in ss_info.get("sheets", [])}

    if sheet_name in sheets_map:
        sheet_id = sheets_map[sheet_name]
        chk = sheets_svc.spreadsheets().values().get(
            spreadsheetId=ss_id, range=f"'{sheet_name}'!A7"
        ).execute()
        if not chk.get("values"):
            _initialize_modern_month_tab(sheets_svc, ss_id, sheet_id, sheet_name, year, saldo_awal=0.0)
        return sheet_name, sheet_id

    # Ambil Saldo Akhir bulan sebelumnya dari sel G5 (KPI Saldo Akhir JFS)
    prev_saldo = 0.0
    if month > 1:
        prev_name = MONTH_SHEET_MAP[month - 1]
        if prev_name in sheets_map:
            prev_vals = sheets_svc.spreadsheets().values().get(
                spreadsheetId=ss_id,
                range=f"'{prev_name}'!G5",
                valueRenderOption="UNFORMATTED_VALUE",
            ).execute().get("values", [])
            if prev_vals and prev_vals[0] and isinstance(prev_vals[0][0], (int, float)):
                prev_saldo = float(prev_vals[0][0])

    res = sheets_svc.spreadsheets().batchUpdate(
        spreadsheetId=ss_id,
        body={"requests": [{"addSheet": {"properties": {"title": sheet_name}}}]},
    ).execute()
    new_sheet_id = res["replies"][0]["addSheet"]["properties"]["sheetId"]

    _initialize_modern_month_tab(sheets_svc, ss_id, new_sheet_id, sheet_name, year, saldo_awal=prev_saldo)
    return sheet_name, new_sheet_id


def append_jfs_transaction_to_gdrive(payload: dict) -> dict:
    drive_svc, sheets_svc = _get_google_services()

    tgl_str = payload["tanggal"]
    dt_tx = datetime.strptime(tgl_str, "%Y-%m-%d")
    year = dt_tx.year
    month = dt_tx.month

    ss_meta = _get_or_create_jfs_spreadsheet(drive_svc, sheets_svc, year)
    ss_id = ss_meta["spreadsheet_id"]
    sheet_name, sheet_id = _ensure_month_tab_exists(sheets_svc, ss_id, month, year)

    # Cari baris terakhir mulai baris 8
    col_d = sheets_svc.spreadsheets().values().get(
        spreadsheetId=ss_id, range=f"'{sheet_name}'!D1:D2000"
    ).execute().get("values", [])
    next_row = max(8, len(col_d) + 1)

    jenis = payload["jenis_transaksi"]
    pengirim = (payload.get("nama_pengirim") or "").strip()
    no_resi = (payload.get("no_resi") or "").strip()
    catatan = (payload.get("catatan") or "").strip()

    nominal_jfs = float(payload.get("nominal_jfs") or 0)
    asuransi = float(payload.get("asuransi") or 0)
    packing = float(payload.get("packing") or 0)
    nominal_paid = float(payload.get("nominal_paid") or 0)

    status_bayar = payload.get("status_bayar", "LUNAS")
    metode_bayar = payload.get("metode_bayar", "BCA MULYATI")
    kategori_pengeluaran = payload.get("kategori_pengeluaran", "")

    bukti_url = payload.get("bukti_url", "")
    foto_url = payload.get("foto_url", "")

    def make_row(r_idx: int, spec: dict) -> list:
        no_formula = 1 if r_idx == 8 else f"=A{r_idx - 1}+1"
        saldo_formula = f"=A5+E{r_idx}-F{r_idx}" if r_idx == 8 else f"=G{r_idx - 1}+E{r_idx}-F{r_idx}"

        omzet_val = f"=F{r_idx}*5/3" if spec.get("is_penjualan") else (f"=F{r_idx}" if spec.get("is_asuransi") else "")
        komisi_val = f"=H{r_idx}-F{r_idx}" if spec.get("is_penjualan") else ""

        paid_val = spec["paid"] if spec.get("paid") is not None else ""
        bal_val = f"=H{r_idx}" if spec.get("is_balance") else ""

        bukti_cell = f'=HYPERLINK("{spec["bukti_url"]}", "Lihat Bukti ↗")' if spec.get("bukti_url") else "-"
        foto_cell = f'=HYPERLINK("{spec["foto_url"]}", "Lihat Foto ↗")' if spec.get("foto_url") else "-"

        return [
            no_formula,                                                   # A: NO
            tgl_str,                                                      # B: TANGGAL
            spec["kategori"],                                             # C: KATEGORI
            spec["ket"],                                                  # D: KETERANGAN / RESI
            spec["debet"] if spec["debet"] is not None else "",           # E: DEBET
            spec["kredit"] if spec["kredit"] is not None else "",         # F: KREDIT
            saldo_formula,                                                # G: SALDO JFS
            omzet_val,                                                    # H: OMZET
            komisi_val,                                                   # I: KOMISI 40%
            paid_val,                                                     # J: PAID
            bal_val,                                                      # K: BALANCE
            spec.get("status", "-"),                                      # L: STATUS
            spec.get("metode_catatan", "-"),                              # M: METODE & CATATAN
            bukti_cell,                                                   # N: BUKTI
            foto_cell,                                                    # O: FOTO BARANG
        ]

    specs = []
    if jenis == "PENJUALAN":
        ket_main = f"Penjualan {pengirim}".strip()
        if no_resi:
            ket_main += f" — Resi {no_resi}"
        info_notes = f"{metode_bayar}"
        if packing > 0:
            info_notes += f" | Packing Rp {int(packing):,}"
        if catatan:
            info_notes += f" | {catatan}"

        specs.append({
            "kategori": "PENJUALAN",
            "ket": ket_main,
            "debet": None,
            "kredit": nominal_jfs,
            "is_penjualan": True,
            "paid": nominal_paid if status_bayar == "LUNAS" else None,
            "is_balance": status_bayar != "LUNAS",
            "status": "LUNAS" if status_bayar == "LUNAS" else "TEMPO",
            "metode_catatan": info_notes,
            "bukti_url": bukti_url,
            "foto_url": foto_url,
        })

        if asuransi > 0:
            specs.append({
                "kategori": "ASURANSI",
                "ket": f"Asuransi Pengiriman ({pengirim or no_resi})",
                "debet": None,
                "kredit": asuransi,
                "is_asuransi": True,
                "paid": asuransi if status_bayar == "LUNAS" else None,
                "status": "LUNAS" if status_bayar == "LUNAS" else "TEMPO",
                "metode_catatan": metode_bayar,
            })

    elif jenis in ("FEE_INCOMING", "COD_FEE", "CASHBACK", "TOP_UP", "REFUND_DFOD"):
        label_map = {
            "FEE_INCOMING": ("FEE INCOMING", "Fee INCOMING J&T Cargo"),
            "COD_FEE": ("COD FEE", "COD Fee Harian"),
            "CASHBACK": ("CASHBACK", "Cash back LEVEL A"),
            "TOP_UP": ("TOP UP", "TOP UP Saldo JFS"),
            "REFUND_DFOD": ("REFUND DFOD", f"Refund DFOD {pengirim} {no_resi}".strip()),
        }
        kat_lbl, ket_lbl = label_map[jenis]
        specs.append({
            "kategori": kat_lbl,
            "ket": ket_lbl,
            "debet": nominal_jfs,
            "kredit": None,
            "status": "MASUK JFS",
            "metode_catatan": catatan or "-",
            "bukti_url": bukti_url,
            "foto_url": foto_url,
        })
        biaya_adm = float(payload.get("biaya_admin") or 0)
        if jenis == "TOP_UP" and biaya_adm > 0:
            specs.append({
                "kategori": "BIAYA ADMIN",
                "ket": "Biaya Admin TOP UP JFS",
                "debet": None,
                "kredit": biaya_adm,
                "status": "POTONG JFS",
                "metode_catatan": "-",
            })

    elif jenis in ("DFOD_INCOMING", "PAYMENT_COD"):
        kat_lbl = "DFOD INCOMING" if jenis == "DFOD_INCOMING" else "PAYMENT COD"
        ket_lbl = f"{kat_lbl} — {pengirim}" if pengirim else kat_lbl
        specs.append({
            "kategori": kat_lbl,
            "ket": ket_lbl,
            "debet": None,
            "kredit": nominal_jfs,
            "status": "POTONG JFS",
            "metode_catatan": f"{metode_bayar} {catatan}".strip(),
            "bukti_url": bukti_url,
            "foto_url": foto_url,
        })

    elif jenis == "PENGELUARAN":
        specs.append({
            "kategori": f"EXP: {kategori_pengeluaran or 'OPERASIONAL'}",
            "ket": f"{kategori_pengeluaran or 'Operasional Outlet'} — {catatan}".strip(" —"),
            "debet": None,
            "kredit": nominal_jfs,
            "status": "KELUAR",
            "metode_catatan": metode_bayar,
            "bukti_url": bukti_url,
            "foto_url": foto_url,
        })

    else:
        arah = payload.get("arah_mutasi", "KREDIT")
        specs.append({
            "kategori": "LAIN-LAIN",
            "ket": catatan or "Transaksi Lain-lain",
            "debet": nominal_jfs if arah == "DEBET" else None,
            "kredit": nominal_jfs if arah == "KREDIT" else None,
            "status": arah,
            "metode_catatan": metode_bayar,
            "bukti_url": bukti_url,
            "foto_url": foto_url,
        })

    rows_values = [make_row(next_row + i, sp) for i, sp in enumerate(specs)]
    end_row = next_row + len(rows_values) - 1

    sheets_svc.spreadsheets().values().update(
        spreadsheetId=ss_id,
        range=f"'{sheet_name}'!A{next_row}:O{end_row}",
        valueInputOption="USER_ENTERED",
        body={"values": rows_values},
    ).execute()

    # Terapkan Zebra Striping & Border Halus pada baris yang baru ditambahkan
    style_reqs = []
    for i in range(len(rows_values)):
        r_zero = (next_row - 1) + i
        bg_hex = "#F8FAFC" if (r_zero % 2 == 1) else "#FFFFFF"
        style_reqs.append({
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": r_zero, "endRowIndex": r_zero + 1, "startColumnIndex": 0, "endColumnIndex": 15},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": _hex_to_rgb(bg_hex),
                        "wrapStrategy": "WRAP",
                        "borders": {
                            "bottom": {"style": "SOLID", "width": 1, "color": _hex_to_rgb("#E2E8F0")},
                        },
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,wrapStrategy,borders)",
            }
        })
    sheets_svc.spreadsheets().batchUpdate(spreadsheetId=ss_id, body={"requests": style_reqs}).execute()

    return get_gdrive_month_overview(month, year)


def get_gdrive_month_overview(month: int, year: int) -> dict:
    drive_svc, sheets_svc = _get_google_services()
    ss_meta = _get_or_create_jfs_spreadsheet(drive_svc, sheets_svc, year)
    ss_id = ss_meta["spreadsheet_id"]
    sheet_name, _ = _ensure_month_tab_exists(sheets_svc, ss_id, month, year)

    res = sheets_svc.spreadsheets().values().get(
        spreadsheetId=ss_id,
        range=f"'{sheet_name}'!A5:O1000",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()
    raw_rows = res.get("values", [])

    # Ambil URL asli dari formula HYPERLINK untuk ditampilkan di tabel web
    res_formulas = sheets_svc.spreadsheets().values().get(
        spreadsheetId=ss_id,
        range=f"'{sheet_name}'!N8:O1000",
        valueRenderOption="FORMULA",
    ).execute().get("values", [])

    saldo_awal = 0.0
    if raw_rows and len(raw_rows[0]) > 0 and isinstance(raw_rows[0][0], (int, float)):
        saldo_awal = float(raw_rows[0][0])

    tx_list = []
    tot_deb = 0.0
    tot_kred = 0.0
    tot_omzet = 0.0
    tot_komisi = 0.0
    last_saldo = saldo_awal

    # Data transaksi dimulai dari index 3 pada raw_rows (Row 8 di Sheets)
    for idx, r in enumerate(raw_rows[3:]):
        def cell(i, default=""):
            return r[i] if i < len(r) else default

        ket = str(cell(3, "")).strip()
        deb = cell(4, 0)
        kred = cell(5, 0)
        if not ket and not deb and not kred:
            continue

        deb_f = float(deb) if isinstance(deb, (int, float)) else 0.0
        kred_f = float(kred) if isinstance(kred, (int, float)) else 0.0
        sal_val = cell(6, None)
        last_saldo = float(sal_val) if isinstance(sal_val, (int, float)) else round(last_saldo + deb_f - kred_f, 2)

        omz_f = float(cell(7, 0)) if isinstance(cell(7, 0), (int, float)) else 0.0
        kom_f = float(cell(8, 0)) if isinstance(cell(8, 0), (int, float)) else 0.0
        paid_f = float(cell(9, 0)) if isinstance(cell(9, 0), (int, float)) else 0.0
        bal_f = float(cell(10, 0)) if isinstance(cell(10, 0), (int, float)) else 0.0

        tot_deb += deb_f
        tot_kred += kred_f
        tot_omzet += omz_f
        tot_komisi += kom_f

        # Ekstrak URL dari =HYPERLINK("url", "label")
        f_row = res_formulas[idx] if idx < len(res_formulas) else []
        bukti_url = ""
        foto_url = ""
        if len(f_row) > 0 and str(f_row[0]).startswith("=HYPERLINK"):
            parts = str(f_row[0]).split('"')
            if len(parts) >= 2:
                bukti_url = parts[1]
        if len(f_row) > 1 and str(f_row[1]).startswith("=HYPERLINK"):
            parts = str(f_row[1]).split('"')
            if len(parts) >= 2:
                foto_url = parts[1]

        tx_list.append({
            "row": 8 + idx,
            "tanggal": str(cell(1, "-")),
            "keterangan": f"[{cell(2, '')}] {ket}".strip(),
            "debet": deb_f,
            "kredit": kred_f,
            "saldo": last_saldo,
            "omzet": omz_f,
            "komisi": kom_f,
            "paid": paid_f,
            "balance": bal_f,
            "ket_bayar": f"{cell(11, '')} — {cell(12, '')}".strip(" —"),
            "bukti_url": bukti_url,
            "foto_url": foto_url,
        })

    tx_list.reverse()

    return {
        "spreadsheet_title": ss_meta["title"],
        "spreadsheet_url": ss_meta["url"],
        "sheet": sheet_name,
        "saldo_awal": round(saldo_awal, 2),
        "total_debet": round(tot_deb, 2),
        "total_kredit": round(tot_kred, 2),
        "saldo_akhir": round(last_saldo, 2),
        "total_omzet": round(tot_omzet, 2),
        "total_komisi": round(tot_komisi, 2),
        "transactions": tx_list,
    }