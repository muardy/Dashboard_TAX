import io
import re
from datetime import datetime, date
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BANK_SIGNATURES = [
    ("BCA", ["GRUP BCA", "REKENING TAHAPAN", "KLIKBCA", "MYBCA", "BANK CENTRAL ASIA"]),
    ("BRI", ["LAPORAN TRANSAKSI FINANSIAL", "BRIMO", "BRITAMA", "BANK RAKYAT INDONESIA", "QLOLA"]),
    ("MANDIRI", ["BANK MANDIRI", "LIVIN", "KOPRA", "MANDIRI ONLINE"]),
    ("BNI", ["BANK NEGARA INDONESIA", "BNI MOBILE", "BNI DIRECT", "WONDERS"]),
    ("BSI", ["BANK SYARIAH INDONESIA", "BYOND", "BSI NET"]),
    ("CIMB", ["CIMB NIAGA", "OCTO MOBILE", "OCTO CLICKS"]),
    ("PERMATA", ["BANK PERMATA", "PERMATAMOBILE", "PERMATANET"]),
    ("JAGO", ["BANK JAGO", "JAGO SYARIAH"]),
    ("SEABANK", ["SEABANK INDONESIA", "PT BANK SEABANK"]),
    ("OCBC", ["BANK OCBC", "OCBC NISP", "ONE MOBILE"]),
    ("DANAMON", ["BANK DANAMON", "D-BANK PRO"]),
    ("BTPN", ["BANK BTPN", "JENIUS", "SMBC INDONESIA"]),
    ("BTN", ["BANK TABUNGAN NEGARA", "BTN MOBILE"]),
    ("MAYBANK", ["MAYBANK INDONESIA", "M2U ID"]),
    ("MEGA", ["BANK MEGA", "M-SMILE"]),
    ("PANIN", ["BANK PANIN", "MOBILE PANIN"]),
    ("SINARMAS", ["BANK SINARMAS", "SIMOBI"]),
]

MONTH_NAMES = {
    "JAN": "01", "JANUARI": "01", "JANUARY": "01",
    "FEB": "02", "FEBRUARI": "02", "FEBRUARY": "02",
    "MAR": "03", "MARET": "03", "MARCH": "03",
    "APR": "04", "APRIL": "04",
    "MEI": "05", "MAY": "05",
    "JUN": "06", "JUNI": "06", "JUNE": "06",
    "JUL": "07", "JULI": "07", "JULY": "07",
    "AGU": "08", "AGUSTUS": "08", "AUG": "08", "AUGUST": "08",
    "SEP": "09", "SEPTEMBER": "09",
    "OKT": "10", "OKTOBER": "10", "OCT": "10", "OCTOBER": "10",
    "NOV": "11", "NOVEMBER": "11",
    "DES": "12", "DESEMBER": "12", "DEC": "12", "DECEMBER": "12",
}


def detect_bank_code(full_text_upper: str) -> str:
    for code, keywords in BANK_SIGNATURES:
        if any(kw in full_text_upper for kw in keywords):
            return code
    bpd_match = re.search(
        r"\b(BANK\s+(?:DKI|BJB|JATENG|JATIM|SUMUT|NAGARI|RIAU|SUMSEL|KALTIM|SULSELBAR|BALI|NTB|NTT|PAPUA))\b",
        full_text_upper,
    )
    if bpd_match:
        return bpd_match.group(1)
    return "UNIVERSAL"


def parse_id_amount(raw_str: str) -> float:
    """Normalisasi semua format nominal bank Indonesia: 1.500.000,00 | 1,500,000.00 | 700.000.00"""
    if not raw_str:
        return 0.0
    cleaned = re.sub(r"[^0-9.,]", "", str(raw_str).strip())
    if not cleaned:
        return 0.0
    if len(cleaned) >= 4 and cleaned[-3] in [".", ","]:
        dec_part = cleaned[-2:]
        int_part = re.sub(r"[.,]", "", cleaned[:-3])
        return float(f"{int_part or '0'}.{dec_part}")
    return float(re.sub(r"[.,]", "", cleaned))


def normalize_date_str(raw_date: str, default_year: int) -> str | None:
    """Konversi berbagai format tanggal bank Indonesia ke ISO YYYY-MM-DD."""
    s = raw_date.strip().upper()
    m = re.search(r"\b(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})\b", s)
    if m:
        dd, mm, yyyy = m.groups()
        return f"{yyyy}-{mm.zfill(2)}-{dd.zfill(2)}"

    m = re.search(r"\b(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})\b", s)
    if m:
        yyyy, mm, dd = m.groups()
        return f"{yyyy}-{mm.zfill(2)}-{dd.zfill(2)}"

    m = re.search(r"\b(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2})\b", s)
    if m:
        dd, mm, yy = m.groups()
        return f"20{yy}-{mm.zfill(2)}-{dd.zfill(2)}"

    m = re.search(r"\b(\d{1,2})[\s\-/]([A-Z]{3,9})[\s\-/](\d{2,4})\b", s)
    if m and m.group(2) in MONTH_NAMES:
        dd, mon_str, yr = m.groups()
        yyyy = yr if len(yr) == 4 else f"20{yr}"
        return f"{yyyy}-{MONTH_NAMES[mon_str]}-{dd.zfill(2)}"

    m = re.match(r"^(\d{2})/(\d{2})$", s)
    if m:
        dd, mm = m.groups()
        return f"{default_year}-{mm}-{dd}"

    return None


def to_excel_date(iso_str: str):
    """Konversi YYYY-MM-DD ke objek datetime.date agar format dd/mm di Excel aktif."""
    try:
        parts = [int(p) for p in str(iso_str).split("-")]
        if len(parts) == 3:
            return date(parts[0], parts[1], parts[2])
    except Exception:
        pass
    return iso_str


def format_remark_and_note(ket_lines: list[str], masuk: float, keluar: float) -> tuple[str, str]:
    """
    Menyederhanakan keterangan mutasi bank menjadi:
    1) REMARK 1 baris standar Buku Bank (misal: 'TRANSFER E-BANKING SURYA', 'TRANSFER E-BANKING KE SRI ZULISTIANA')
    2) KETERANGAN / catatan berita transfer (misal: 'Belanja oli')
    """
    raw = [ln.strip() for ln in ket_lines if ln and ln.strip()]
    if not raw:
        return "-", ""

    first = raw[0].upper()
    is_db = keluar > 0 and masuk == 0

    # 1. Pola khusus BCA
    if first.startswith("BIAYA ADM"):
        return "BIAYA ADMIN", ""
    if first == "BUNGA":
        return "BUNGA", ""
    if first.startswith("PAJAK BUNGA"):
        return "PAJAK BUNGA", ""
    if first.startswith("DB OTOMATIS"):
        return "DB OTOMATIS AUTO DEBIT", ""
    if first.startswith("SETORAN TUNAI"):
        depositor = " ".join(ln for ln in raw[1:] if not re.match(r"^\d+$", ln))
        return "SETORAN TUNAI", depositor

    if first.startswith("SWITCHING"):
        prefix = "SWITCHING KE" if is_db else "SWITCHING"
        names = [
            ln for ln in raw[1:]
            if not re.match(r"^(DR|CR)\s*\d+$", ln, re.I)
            and not ln.startswith("/")
        ]
        return f"{prefix} {' '.join(names).strip()}".strip(), ""

    if first.startswith("BI-FAST"):
        prefix = "BI-FAST KE" if is_db else "BI-FAST"
        filtered = [
            ln for ln in raw[1:]
            if not re.match(r"^BIF TRANSFER\s+(DR|KE)$", ln, re.I)
            and not re.match(r"^\d{3}$", ln)
            and not re.match(r"^TANGGAL\s*:", ln, re.I)
        ]
        joined_u = " ".join(filtered).upper()
        if "DOMPET ANAK BANGSA" in joined_u:
            return f"{prefix} GOPAY DOMPET ANAK BANGSA", ""
        if "AIRPAY INTERNATION" in joined_u:
            return f"{prefix} SHOPEEPAY", ""
        if len(filtered) >= 2:
            counterparty = filtered[-1]
            note = " ".join(filtered[:-1]).strip()
            if note and len(note) > 2:
                note = note[0].upper() + note[1:]
            return f"{prefix} {counterparty}".strip(), note
        elif len(filtered) == 1:
            return f"{prefix} {filtered[0]}".strip(), ""
        return prefix, ""

    if first.startswith("TRSF E-BANKING"):
        prefix = "TRANSFER E-BANKING KE" if is_db else "TRANSFER E-BANKING"
        filtered = []
        for ln in raw[1:]:
            if re.match(r"^\d{4}/[A-Z0-9]+/[A-Z0-9]+$", ln, re.I):
                continue
            if re.match(r"^[\d,]+\.\d{2}$", ln):
                continue
            if re.match(r"^TANGGAL\s*:", ln, re.I):
                continue
            if re.match(
                r"^(RW\s*\d+|RT\s*\d+|V\s+NO\s+\d+A|JL\s+KEBON|TAMAN\s+SARI|MAPHAR|JAKARTA\s+\d+|INDONESIA)$",
                ln,
                re.I,
            ):
                continue
            filtered.append(ln)

        joined_u = " ".join(filtered).upper()
        if "AIRPAY INTERNATION" in joined_u:
            return f"{prefix} SHOPEEPAY", ""
        if "DOMPET ANAK BANGSA" in joined_u:
            return f"{prefix} GOPAY DOMPET ANAK BANGSA", ""
        if len(filtered) >= 2:
            counterparty = filtered[-1]
            note = " ".join(filtered[:-1]).strip()
            if note and len(note) > 2:
                note = note[0].upper() + note[1:]
            return f"{prefix} {counterparty}".strip(), note
        elif len(filtered) == 1:
            return f"{prefix} {filtered[0]}".strip(), ""
        return prefix, ""

    # 2. Pola bank umum (BRI, Mandiri, BNI, dll.)
    single_line = re.sub(r"\s+", " ", " ".join(raw)).strip()
    return single_line, ""


def auto_categorize(desc: str, masuk: float, keluar: float) -> tuple[str, str]:
    """Klasifikasi otomatis akun pembukuan & ekstraksi nama lawan transaksi."""
    upper = desc.upper()
    lines = [ln.strip() for ln in desc.split("\n") if ln.strip()]
    counterparty = lines[-1] if lines else "-"

    if " - " in desc and ("BI-FAST" in upper or "TRANSFER" in upper):
        parts = [p.strip() for p in desc.split(" - ") if p.strip()]
        if len(parts) >= 2:
            counterparty = parts[-1]

    if any(k in upper for k in ["BIAYA ADM", "ADMIN FEE", "MONTHLY FEE", "BIF BIAYA", "BIAYA TRANSFER", "MATERAI"]):
        return "Biaya Administrasi Bank", "Bank"
    if any(k in upper for k in ["BUNGA", "JASA GIRO", "BAGI HASIL", "INTEREST"]):
        return "Pendapatan Bunga / Bagi Hasil", "Bank"
    if any(k in upper for k in ["PAJAK BUNGA", "PPH BUNGA", "TAX"]):
        return "Pajak Bunga Bank (PPh Final)", "Negara / Bank"
    if "SETORAN VIA CDM" in upper or "SETOR TUNAI" in upper or "SETORAN TUNAI" in upper or "CASH DEPOSIT" in upper:
        return "Setoran Tunai", counterparty
    if any(k in upper for k in ["TARIKAN ATM", "PENARIKAN TUNAI", "CASH WITHDRAWAL"]):
        return "Tarik Tunai ATM", "Kas Tunai"
    if any(k in upper for k in ["QR ", "QRC", "QRIS", "PRCH ", "EDC", "DEBIT CARD"]):
        return "Pembayaran QRIS / EDC", counterparty
    if any(k in upper for k in ["FTFVA", "BRIVA", "VIRTUAL ACCOUNT", "VA ", "PEMBAYARAN"]):
        return "Pembayaran Virtual Account / Tagihan", counterparty
    if any(k in upper for k in ["SALARY", "GAJIH", "GAJI", "PAYROLL", "UPAH"]):
        return "Beban Gaji / Payroll", counterparty
    if masuk > 0:
        return "Penerimaan Transfer / Usaha", counterparty
    return "Pengeluaran Transfer / Operasional", counterparty


def extract_generic_metadata(first_text: str, bank_label: str) -> dict:
    meta = {
        "bank": bank_label,
        "no_rekening": "-",
        "nama_pemilik": "-",
        "periode": "-",
        "mata_uang": "IDR",
        "saldo_awal": 0.0,
        "saldo_akhir_pdf": 0.0,
    }
    rek_m = re.search(
        r"(?:NO\.?\s*REKENING|ACCOUNT\s*NO\.?|NOMOR\s*REKENING|REKENING)\s*[:\-]?\s*([\d\s]{10,20})",
        first_text,
        re.IGNORECASE,
    )
    if rek_m:
        digits = re.sub(r"\s+", "", rek_m.group(1))
        if len(digits) >= 10:
            meta["no_rekening"] = digits

    per_m = re.search(
        r"(?:PERIODE|PERIOD)\s*[:\-]?\s*([A-Z0-9\s/\-\.]{6,35})",
        first_text,
        re.IGNORECASE,
    )
    if per_m:
        meta["periode"] = per_m.group(1).splitlines()[0].strip()

    name_m = re.search(
        r"(?:NAMA\s*PEMILIK|NAMA\s*NASABAH|ACCOUNT\s*NAME|KEPADA\s*YTH\.?(?:\s*/\s*TO)?)\s*[:\-]?\s*\n?\s*([A-Z][A-Z\s.,']{3,45})",
        first_text,
        re.IGNORECASE,
    )
    if name_m:
        candidate = name_m.group(1).splitlines()[0].strip()
        if not any(noise in candidate.upper() for noise in ["REKENING", "HALAMAN", "PERIODE", "TANGGAL"]):
            meta["nama_pemilik"] = candidate

    return meta


def parse_bca_pdf(pdf_bytes: bytes) -> dict:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        first_text = pdf.pages[0].extract_text() or ""
        meta = extract_generic_metadata(first_text, "Bank Central Asia (BCA)")
        year = datetime.now().year

        per_m = re.search(r"PERIODE\s*:\s*([A-Z]+)\s+(\d{4})", first_text)
        if per_m:
            meta["periode"] = f"{per_m.group(1)} {per_m.group(2)}"
            year = int(per_m.group(2))

        lines_p1 = first_text.splitlines()
        for idx, line in enumerate(lines_p1):
            compact_l = re.sub(r"\s+", "", line).upper()
            if ("KCP" in compact_l or "KCU" in compact_l) and idx + 1 < len(lines_p1):
                meta["nama_pemilik"] = lines_p1[idx + 1].strip()
                break

        transactions = []
        current_tx = None

        for page in pdf.pages:
            words = page.extract_words(keep_blank_chars=False, y_tolerance=3)
            rows_by_y = {}
            for w in words:
                y_key = round(w["top"] / 4) * 4
                rows_by_y.setdefault(y_key, []).append(w)

            in_table = False
            for y_key in sorted(rows_by_y.keys()):
                row_words = sorted(rows_by_y[y_key], key=lambda item: item["x0"])
                line_text = " ".join(w["text"] for w in row_words)
                compact = re.sub(r"\s+", "", line_text).upper()

                # Kunci area tabel: hanya baca baris di bawah header TANGGAL KETERANGAN SALDO
                if "TANGGAL" in compact and "KETERANGAN" in compact and "SALDO" in compact:
                    in_table = True
                    continue
                if not in_table:
                    continue

                if "BERSAMBUNGKEHALAMAN" in compact:
                    break
                if compact.startswith("SALDOAWAL:"):
                    meta["saldo_awal"] = parse_id_amount(line_text.split(":", 1)[1])
                    break
                if compact.startswith("SALDOAKHIR:"):
                    meta["saldo_akhir_pdf"] = parse_id_amount(line_text.split(":", 1)[1])
                    break
                if compact.startswith(("MUTASICR", "MUTASIDB")):
                    break

                col_tgl = " ".join(w["text"] for w in row_words if w["x0"] < 85)
                col_ket = " ".join(w["text"] for w in row_words if 85 <= w["x0"] < 310)
                col_mutasi = " ".join(w["text"] for w in row_words if 345 <= w["x0"] < 475)
                col_saldo = " ".join(w["text"] for w in row_words if w["x0"] >= 475)

                if re.match(r"^\d{2}/\d{2}$", col_tgl.strip()):
                    if "SALDO AWAL" in col_ket.upper():
                        if col_saldo:
                            meta["saldo_awal"] = parse_id_amount(col_saldo)
                        continue
                    if current_tx:
                        transactions.append(current_tx)

                    tgl_iso = normalize_date_str(col_tgl.strip(), year)
                    is_db = "DB" in col_mutasi.upper()
                    amt = parse_id_amount(col_mutasi)

                    current_tx = {
                        "tanggal": tgl_iso,
                        "keterangan_lines": [col_ket] if col_ket else [],
                        "masuk": 0.0 if is_db else amt,
                        "keluar": amt if is_db else 0.0,
                    }
                elif current_tx:
                    if col_ket:
                        current_tx["keterangan_lines"].append(col_ket)

        if current_tx:
            transactions.append(current_tx)

    return finalize_rows(meta, transactions)


def parse_bri_pdf(pdf_bytes: bytes) -> dict:
    watermark_tokens = re.compile(
        r"\b(BRI|RIBRI|BRIBRI|KBRI|BIBRI|BRIB|ARIBRI|KARL|BRL|RBR|ARI|IBR|BS|R|B)\b"
    )
    glued_noise = re.compile(r"^[A-Z]+(?=[\d,]+\.\d{2}$)")

    def clean_bri_text(txt: str) -> str:
        txt = watermark_tokens.sub(" ", txt)
        return re.sub(r"\s+", " ", txt).strip()

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        first_text = clean_bri_text(pdf.pages[0].extract_text() or "")
        meta = extract_generic_metadata(first_text, "Bank Rakyat Indonesia (BRI)")

        rek_m = re.search(r"\b(\d{15})\b", first_text)
        if rek_m:
            meta["no_rekening"] = rek_m.group(1)

        per_m = re.search(r"(\d{2}/\d{2}/\d{2})\s*-\s*(\d{2}/\d{2}/\d{2})", first_text)
        if per_m:
            meta["periode"] = f"{per_m.group(1)} s.d. {per_m.group(2)}"

        name_m = re.search(r"Kepada Yth\.?\s*/?\s*To\s*:?\s*([A-Z][A-Z\s.,']{3,45})", first_text)
        if name_m:
            meta["nama_pemilik"] = clean_bri_text(name_m.group(1))

        last_words = [
            w for w in pdf.pages[-1].extract_words(keep_blank_chars=False, y_tolerance=3)
            if w.get("upright", True)
        ]
        last_text = clean_bri_text(
            " ".join(w["text"] for w in sorted(last_words, key=lambda x: (round(x["top"] / 4), x["x0"])))
        )
        sum_m = re.search(
            r"([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+Terbilang",
            last_text,
        )
        if sum_m:
            meta["saldo_awal"] = parse_id_amount(sum_m.group(1))
            meta["saldo_akhir_pdf"] = parse_id_amount(sum_m.group(4))

        transactions = []

        for page in pdf.pages:
            words = [
                w for w in page.extract_words(keep_blank_chars=False, y_tolerance=3)
                if w.get("upright", True)
            ]
            if not words:
                continue

            y_header = 0.0
            y_footer = page.height
            col_x = {"uraian": 105.0, "teller": 315.0, "debet": 365.0, "kredit": 445.0, "saldo": 515.0}

            rows_by_y = {}
            for w in words:
                y_key = round(w["top"] / 4) * 4
                rows_by_y.setdefault(y_key, []).append(w)

            for y_key in sorted(rows_by_y.keys()):
                rw = sorted(rows_by_y[y_key], key=lambda x: x["x0"])
                line_u = " ".join(w["text"].upper() for w in rw)

                if ("URAIAN" in line_u or "DESCRIPTION" in line_u) and ("DEBET" in line_u or "DEBIT" in line_u):
                    y_header = max(w["bottom"] for w in rw)
                    for w in rw:
                        wt = w["text"].upper()
                        if wt in ("URAIAN", "TRANSACTION") and w["x0"] > 80:
                            col_x["uraian"] = min(col_x["uraian"], w["x0"] - 5)
                        elif wt in ("TELLER", "USER"):
                            col_x["teller"] = w["x0"] - 8
                        elif wt in ("DEBET", "DEBIT"):
                            col_x["debet"] = w["x0"] - 18
                        elif wt in ("KREDIT", "CREDIT"):
                            col_x["kredit"] = w["x0"] - 18
                        elif wt in ("SALDO", "BALANCE"):
                            col_x["saldo"] = w["x0"] - 18

                if y_header > 0 and any(k in line_u for k in ["UNIT KERJA", "BUSINESS UNIT", "SALDO AWAL", "OPENING BALANCE"]):
                    y_footer = min(y_footer, min(w["top"] for w in rw) - 2)

            body_words = [w for w in words if y_header < w["top"] < y_footer]
            date_anchors = sorted(
                [
                    w for w in body_words
                    if w["x0"] < col_x["uraian"] and re.match(r"^\d{2}/\d{2}/\d{2}$", w["text"].strip())
                ],
                key=lambda x: x["top"],
            )

            if not date_anchors:
                continue

            for idx, d_word in enumerate(date_anchors):
                y_start = y_header if idx == 0 else (date_anchors[idx - 1]["top"] + d_word["top"]) / 2.0
                y_end = y_footer if idx == len(date_anchors) - 1 else (d_word["top"] + date_anchors[idx + 1]["top"]) / 2.0

                band_words = [w for w in body_words if y_start <= w["top"] < y_end]
                band_words.sort(key=lambda x: (round(x["top"] / 4), x["x0"]))

                tgl_parts, desc_parts, deb_parts, kred_parts = [], [], [], []
                for w in band_words:
                    x0 = w["x0"]
                    txt = glued_noise.sub("", w["text"].strip())
                    if not txt:
                        continue
                    if x0 < col_x["uraian"]:
                        tgl_parts.append(txt)
                    elif x0 < col_x["teller"]:
                        desc_parts.append(txt)
                    elif col_x["debet"] <= x0 < col_x["kredit"]:
                        deb_parts.append(txt)
                    elif col_x["kredit"] <= x0 < col_x["saldo"]:
                        kred_parts.append(txt)

                tgl_iso = normalize_date_str(d_word["text"].strip(), datetime.now().year)
                if not tgl_iso:
                    continue

                jam_m = re.search(r"\b(\d{2}:\d{2}:\d{2})\b", " ".join(tgl_parts))
                jam_str = f" [{jam_m.group(1)}]" if jam_m else ""

                desc_clean = clean_bri_text(" ".join(desc_parts))
                debet = parse_id_amount(" ".join(deb_parts))
                kredit = parse_id_amount(" ".join(kred_parts))

                if debet == 0.0 and kredit == 0.0 and not desc_clean:
                    continue

                transactions.append({
                    "tanggal": tgl_iso,
                    "keterangan_lines": [f"{desc_clean}{jam_str}".strip()],
                    "masuk": kredit,
                    "keluar": debet,
                })

    return finalize_rows(meta, transactions)


def parse_universal_bank_pdf(pdf_bytes: bytes, bank_code: str) -> dict:
    """
    Universal Adaptive Parser untuk Mandiri, BNI, BSI, CIMB, Permata, Jago, SeaBank,
    OCBC, Danamon, BTN, Jenius, dan seluruh BPD di Indonesia.
    """
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        first_text = pdf.pages[0].extract_text() or ""
        bank_label = f"Bank {bank_code}" if bank_code != "UNIVERSAL" else "Rekening Koran Bank"
        meta = extract_generic_metadata(first_text, bank_label)

        year_m = re.search(r"\b(20\d{2})\b", first_text)
        default_year = int(year_m.group(1)) if year_m else datetime.now().year

        sa_m = re.search(
            r"(?:SALDO\s*AWAL|OPENING\s*BALANCE|BEGINNING\s*BALANCE)\s*[:\-]?\s*(?:IDR|RP\.?)?\s*([\d.,]+)",
            first_text,
            re.IGNORECASE,
        )
        if sa_m:
            meta["saldo_awal"] = parse_id_amount(sa_m.group(1))

        transactions = []
        col_bounds = None

        for page in pdf.pages:
            words = page.extract_words(keep_blank_chars=False, y_tolerance=3)
            rows_by_y = {}
            for w in words:
                y_key = round(w["top"] / 4) * 4
                rows_by_y.setdefault(y_key, []).append(w)

            for y_key in sorted(rows_by_y.keys()):
                row_words = sorted(rows_by_y[y_key], key=lambda item: item["x0"])
                row_upper = " ".join(w["text"].upper() for w in row_words)

                has_date_hdr = any(k in row_upper for k in ["TANGGAL", "TGL", "DATE"])
                has_amt_hdr = any(
                    k in row_upper
                    for k in ["DEBET", "DEBIT", "KREDIT", "CREDIT", "MUTASI", "NOMINAL", "SALDO", "BALANCE"]
                )
                if has_date_hdr and has_amt_hdr:
                    detected = {}
                    for w in row_words:
                        wt = w["text"].upper()
                        if wt in ["TANGGAL", "TGL", "DATE"] and "tgl" not in detected:
                            detected["tgl"] = w["x0"]
                        elif wt in ["KETERANGAN", "URAIAN", "DESKRIPSI", "TRANSAKSI", "DESCRIPTION", "REMARK"] and "ket" not in detected:
                            detected["ket"] = w["x0"]
                        elif wt in ["DEBET", "DEBIT", "KELUAR", "PENARIKAN"] and "deb" not in detected:
                            detected["deb"] = w["x0"]
                        elif wt in ["KREDIT", "CREDIT", "MASUK", "SETORAN"] and "kred" not in detected:
                            detected["kred"] = w["x0"]
                        elif wt in ["MUTASI", "NOMINAL", "JUMLAH", "AMOUNT"] and "mut" not in detected:
                            detected["mut"] = w["x0"]
                        elif wt in ["SALDO", "BALANCE"] and "sal" not in detected:
                            detected["sal"] = w["x0"]

                    if "tgl" in detected:
                        col_bounds = detected
                    continue

                sak_m = re.search(
                    r"(?:SALDO\s*AKHIR|CLOSING\s*BALANCE|ENDING\s*BALANCE)\s*[:\-]?\s*(?:IDR|RP\.?)?\s*([\d.,]+)",
                    row_upper,
                )
                if sak_m:
                    meta["saldo_akhir_pdf"] = parse_id_amount(sak_m.group(1))
                    continue

                if not col_bounds:
                    continue

                buckets = {"tgl": [], "ket": [], "deb": [], "kred": [], "mut": [], "sal": []}
                sorted_cols = sorted(col_bounds.items(), key=lambda kv: kv[1])

                for w in row_words:
                    chosen_col = sorted_cols[0][0]
                    for col_name, col_x in sorted_cols:
                        if w["x0"] >= col_x - 18:
                            chosen_col = col_name
                    buckets[chosen_col].append(w["text"])

                tgl_text = " ".join(buckets["tgl"]).strip()
                tgl_iso = normalize_date_str(tgl_text, default_year)

                if tgl_iso:
                    ket_text = " ".join(buckets["ket"]).strip()
                    if "SALDO AWAL" in ket_text.upper() or "OPENING BALANCE" in ket_text.upper():
                        sal_val = parse_id_amount(" ".join(buckets["sal"]))
                        if sal_val > 0:
                            meta["saldo_awal"] = sal_val
                        continue

                    deb_str = " ".join(buckets["deb"]).strip()
                    kred_str = " ".join(buckets["kred"]).strip()
                    mut_str = " ".join(buckets["mut"]).strip()
                    sal_str = " ".join(buckets["sal"]).strip()

                    masuk = 0.0
                    keluar = 0.0

                    if deb_str or kred_str:
                        keluar = parse_id_amount(deb_str)
                        masuk = parse_id_amount(kred_str)
                    elif mut_str:
                        amt = parse_id_amount(mut_str)
                        mut_u = mut_str.upper()
                        if any(tag in mut_u for tag in ["DB", "DEBIT", "DEBET", "-"]) or mut_u.endswith("D"):
                            keluar = amt
                        elif any(tag in mut_u for tag in ["CR", "CREDIT", "KREDIT", "+"]) or mut_u.endswith("K"):
                            masuk = amt
                        else:
                            masuk = amt

                    saldo_cp = parse_id_amount(sal_str) if sal_str else None
                    transactions.append({
                        "tanggal": tgl_iso,
                        "keterangan_lines": [ket_text] if ket_text else [],
                        "masuk": masuk,
                        "keluar": keluar,
                        "saldo_cp": saldo_cp,
                        "ambiguous_sign": bool(
                            mut_str and not any(t in mut_str.upper() for t in ["DB", "CR", "+", "-", "D", "K"])
                        ),
                    })
                elif transactions:
                    extra_ket = " ".join(buckets["ket"]).strip()
                    if extra_ket and not any(skip in extra_ket.upper() for skip in ["HALAMAN", "PAGE ", "BERSAMBUNG"]):
                        transactions[-1]["keterangan_lines"].append(extra_ket)

        prev_bal = meta["saldo_awal"]
        for tx in transactions:
            cp = tx.get("saldo_cp")
            if tx.get("ambiguous_sign") and cp is not None and prev_bal > 0:
                amt = tx["masuk"]
                if round(prev_bal - amt, 0) == round(cp, 0):
                    tx["keluar"] = amt
                    tx["masuk"] = 0.0
            if cp is not None:
                prev_bal = cp

    return finalize_rows(meta, transactions)


def finalize_rows(meta: dict, raw_transactions: list[dict]) -> dict:
    running_balance = meta["saldo_awal"]
    cleaned_rows = []

    for idx, tx in enumerate(raw_transactions, start=1):
        desc = "\n".join(tx["keterangan_lines"]).strip() or "-"
        remark, catatan = format_remark_and_note(tx["keterangan_lines"], tx["masuk"], tx["keluar"])
        running_balance = round(running_balance + tx["masuk"] - tx["keluar"], 2)
        kategori, lawan = auto_categorize(desc, tx["masuk"], tx["keluar"])
        cleaned_rows.append({
            "no": idx,
            "tanggal": tx["tanggal"],
            "remark": remark,
            "acc_no": "",
            "catatan": catatan,
            "keterangan": desc,
            "lawan_transaksi": lawan,
            "kategori": kategori,
            "masuk": round(tx["masuk"], 2),
            "keluar": round(tx["keluar"], 2),
            "saldo": running_balance,
        })

    if meta["saldo_akhir_pdf"] == 0.0 and cleaned_rows:
        meta["saldo_akhir_pdf"] = cleaned_rows[-1]["saldo"]

    total_masuk = round(sum(r["masuk"] for r in cleaned_rows), 2)
    total_keluar = round(sum(r["keluar"] for r in cleaned_rows), 2)
    saldo_akhir_hitung = cleaned_rows[-1]["saldo"] if cleaned_rows else meta["saldo_awal"]
    selisih = round(saldo_akhir_hitung - meta["saldo_akhir_pdf"], 2)

    meta["total_transaksi"] = len(cleaned_rows)
    meta["total_masuk"] = total_masuk
    meta["total_keluar"] = total_keluar
    meta["saldo_akhir_hitung"] = saldo_akhir_hitung
    meta["is_reconciled"] = abs(selisih) < 1.0
    meta["selisih"] = selisih

    return {"meta": meta, "transactions": cleaned_rows}


def parse_statement_pdf(pdf_bytes: bytes) -> dict:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        if not pdf.pages:
            raise ValueError("File PDF kosong atau tidak memiliki halaman.")
        sample_text = (pdf.pages[0].extract_text() or "").upper()

    if len(sample_text.strip()) < 30:
        raise ValueError(
            "PDF ini merupakan hasil scan gambar (tanpa layer teks). Gunakan e-Statement PDF asli dari internet/mobile banking."
        )

    bank_code = detect_bank_code(sample_text)
    if bank_code == "BCA":
        result = parse_bca_pdf(pdf_bytes)
    elif bank_code == "BRI":
        result = parse_bri_pdf(pdf_bytes)
    else:
        result = parse_universal_bank_pdf(pdf_bytes, bank_code)

    if not result["transactions"]:
        raise ValueError("Tidak ada baris transaksi yang terbaca. Pastikan file adalah e-Statement rekening koran resmi.")

    return result


def build_buku_bank_excel(meta: dict, rows: list[dict]) -> io.BytesIO:
    """
    Membangun file Excel Buku Bank sesuai template standar:
    Kolom B..H = TGL | REMARK | ACC NO. | DEBIT | KREDIT | SALDO | KETERANGAN
    - DEBIT  = Mutasi Uang Masuk (CR)
    - KREDIT = Mutasi Uang Keluar (DB)
    - SALDO  = Rumus berjalan =G{r-1}+E{r}-F{r}
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Buku Bank"

    thin = Side(style="thin", color="000000")
    med = Side(style="medium", color="000000")
    num_fmt = "#,##0.00"
    date_fmt = "dd/mm"

    owner_name = str(meta.get("nama_pemilik") or "BUKU BANK").strip().upper()
    periode_str = str(meta.get("periode") or "").strip().upper()

    # Baris 1: Nama Pemilik Rekening (Merge B1:H1)
    ws.merge_cells("B1:H1")
    ws["B1"] = owner_name
    ws["B1"].font = Font(name="Calibri", size=24, bold=True)
    ws["B1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 42
    for col_idx in range(2, 9):
        c = ws.cell(row=1, column=col_idx)
        c.border = Border(
            top=med,
            bottom=thin,
            left=med if col_idx == 2 else None,
            right=med if col_idx == 8 else None,
        )

    # Baris 2: Judul Periode (Merge B2:H2)
    ws.merge_cells("B2:H2")
    ws["B2"] = f"BUKU BANK PERIODE {periode_str}".strip()
    ws["B2"].font = Font(name="Calibri", size=16, bold=True)
    ws["B2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 24
    for col_idx in range(2, 9):
        c = ws.cell(row=2, column=col_idx)
        c.border = Border(
            top=thin,
            bottom=thin,
            left=med if col_idx == 2 else None,
            right=med if col_idx == 8 else None,
        )

    # Baris 3: Header Kolom B3:H3
    headers = ["TGL", "REMARK", "ACC NO.", "DEBIT", "KREDIT", "SALDO", "KETERANGAN"]
    ws.row_dimensions[3].height = 19
    for offset, h_text in enumerate(headers):
        col_idx = 2 + offset
        cell = ws.cell(row=3, column=col_idx, value=h_text)
        cell.font = Font(name="Calibri", size=11, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(
            top=thin,
            bottom=med,
            left=med if col_idx == 2 else thin,
            right=med if col_idx == 8 else thin,
        )

    # Baris 4: Saldo Awal (B4 = Tgl awal, C4:F4 = SALDO AWAL, G4 = Nominal Saldo Awal)
    first_date = to_excel_date(rows[0]["tanggal"]) if rows else ""
    ws.cell(row=4, column=2, value=first_date)
    ws["B4"].font = Font(name="Calibri", size=11)
    ws["B4"].alignment = Alignment(horizontal="center", vertical="center")
    ws["B4"].number_format = date_fmt

    ws.merge_cells("C4:F4")
    ws["C4"] = "SALDO AWAL"
    ws["C4"].font = Font(name="Calibri", size=11, bold=True)
    ws["C4"].alignment = Alignment(horizontal="center", vertical="center")

    ws["G4"] = float(meta.get("saldo_awal", 0.0))
    ws["G4"].font = Font(name="Calibri", size=11, bold=True)
    ws["G4"].number_format = num_fmt
    ws["G4"].alignment = Alignment(horizontal="right", vertical="center")

    for col_idx in range(2, 9):
        ws.cell(row=4, column=col_idx).border = Border(
            top=thin,
            bottom=thin,
            left=med if col_idx == 2 else thin,
            right=med if col_idx == 8 else thin,
        )

    # Baris 5 dst: Data Transaksi
    reg_font = Font(name="Calibri", size=11)
    for i, tx in enumerate(rows):
        r_idx = 5 + i
        tgl_val = to_excel_date(tx.get("tanggal", ""))
        remark_val = tx.get("remark") or tx.get("keterangan", "-")
        acc_val = tx.get("acc_no") or None
        masuk_val = float(tx.get("masuk", 0.0))
        keluar_val = float(tx.get("keluar", 0.0))
        catatan_val = tx.get("catatan") or None

        c_tgl = ws.cell(row=r_idx, column=2, value=tgl_val)
        c_tgl.alignment = Alignment(horizontal="center", vertical="center")
        c_tgl.number_format = date_fmt

        c_rem = ws.cell(row=r_idx, column=3, value=remark_val)
        c_rem.alignment = Alignment(horizontal="left", vertical="center")

        c_acc = ws.cell(row=r_idx, column=4, value=acc_val)
        c_acc.alignment = Alignment(horizontal="center", vertical="center")

        # DEBIT = Uang Masuk (CR), dikosongkan jika 0 agar bersih sesuai template
        c_deb = ws.cell(row=r_idx, column=5, value=masuk_val if masuk_val > 0 else None)
        c_deb.number_format = num_fmt
        c_deb.alignment = Alignment(horizontal="right", vertical="center")

        # KREDIT = Uang Keluar (DB), dikosongkan jika 0 agar bersih sesuai template
        c_kre = ws.cell(row=r_idx, column=6, value=keluar_val if keluar_val > 0 else None)
        c_kre.number_format = num_fmt
        c_kre.alignment = Alignment(horizontal="right", vertical="center")

        # SALDO = Rumus dinamis berjalan
        c_sal = ws.cell(row=r_idx, column=7, value=f"=G{r_idx - 1}+E{r_idx}-F{r_idx}")
        c_sal.number_format = num_fmt
        c_sal.alignment = Alignment(horizontal="right", vertical="center")

        c_ket = ws.cell(row=r_idx, column=8, value=catatan_val)
        c_ket.alignment = Alignment(horizontal="left", vertical="center")

        for col_idx in range(2, 9):
            cell = ws.cell(row=r_idx, column=col_idx)
            cell.font = reg_font
            cell.border = Border(
                top=thin,
                bottom=thin,
                left=med if col_idx == 2 else thin,
                right=med if col_idx == 8 else thin,
            )

    # Lebar kolom mengikuti template BUKU BANK 2026.xlsx
    col_widths = {
        "A": 4,
        "B": 11,
        "C": 52,
        "D": 12,
        "E": 17,
        "F": 17,
        "G": 18,
        "H": 38,
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    ws.freeze_panes = "B5"

    # Sheet 2: Rekap Kategori Akun (opsional untuk audit/ekualisasi)
    ws_cat = wb.create_sheet(title="Rekap Kategori")
    navy_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    white_bold = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Calibri", size=14, bold=True, color="1E3A8A")
    border_cat = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws_cat["A1"] = "REKAPITULASI MUTASI PER KATEGORI AKUN"
    ws_cat["A1"].font = title_font
    cat_headers = ["Kategori Akun", "Jumlah Transaksi", "Total Debit / Masuk (Rp)", "Total Kredit / Keluar (Rp)"]
    for col_idx, h_text in enumerate(cat_headers, start=1):
        c = ws_cat.cell(row=3, column=col_idx, value=h_text)
        c.fill = navy_fill
        c.font = white_bold
        c.border = border_cat

    cat_summary = {}
    for tx in rows:
        cat = tx.get("kategori", "Lain-lain")
        if cat not in cat_summary:
            cat_summary[cat] = [0, 0.0, 0.0]
        cat_summary[cat][0] += 1
        cat_summary[cat][1] += float(tx.get("masuk", 0.0))
        cat_summary[cat][2] += float(tx.get("keluar", 0.0))

    for idx, (cat, vals) in enumerate(sorted(cat_summary.items()), start=4):
        ws_cat.cell(row=idx, column=1, value=cat).border = border_cat
        c_cnt = ws_cat.cell(row=idx, column=2, value=vals[0])
        c_in = ws_cat.cell(row=idx, column=3, value=round(vals[1], 2))
        c_out = ws_cat.cell(row=idx, column=4, value=round(vals[2], 2))
        for c in (c_cnt, c_in, c_out):
            c.border = border_cat
        c_in.number_format = num_fmt
        c_out.number_format = num_fmt

    for col_idx, width in {1: 34, 2: 18, 3: 24, 4: 24}.items():
        ws_cat.column_dimensions[get_column_letter(col_idx)].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output