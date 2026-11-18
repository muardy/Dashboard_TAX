import os
import time
import math
import psycopg2

_TAX_CACHE = {}
_CACHE_LOADED_AT = 0
CACHE_TTL = 300

def clear_tax_cache():
    global _CACHE_LOADED_AT
    _TAX_CACHE.clear()
    _CACHE_LOADED_AT = 0

def _get_conn():
    return psycopg2.connect(os.getenv("DATABASE_URL"))

def _ensure_cache_loaded():
    global _CACHE_LOADED_AT
    if _TAX_CACHE and (time.time() - _CACHE_LOADED_AT < CACHE_TTL):
        return

    conn = None
    try:
        conn = _get_conn()
        cur = conn.cursor()

        cur.execute("SELECT param_key, param_value FROM ref_parameter_pajak")
        _TAX_CACHE['params'] = {r[0]: float(r[1]) for r in cur.fetchall()}

        cur.execute("""
            SELECT kode_ptkp, nilai_ptkp, kategori_ter, deskripsi
            FROM ref_ptkp
            WHERE is_active = TRUE
            ORDER BY nilai_ptkp ASC, kode_ptkp ASC
        """)
        _TAX_CACHE['ptkp'] = {
            r[0]: {'ptkp': int(r[1]), 'kategori': r[2], 'label': r[3]}
            for r in cur.fetchall()
        }

        cur.execute("""
            SELECT kapasitas_lapisan, tarif_persen, label
            FROM ref_tarif_pasal17
            WHERE is_active = TRUE
            ORDER BY urutan ASC
        """)
        _TAX_CACHE['pasal17'] = cur.fetchall()

        cur.execute("""
            SELECT kategori, batas_atas, tarif_persen
            FROM ref_ter_tarif
            WHERE is_active = TRUE
            ORDER BY batas_atas ASC NULLS LAST
        """)
        ter_grouped = {'A': [], 'B': [], 'C': []}
        for kat, batas, tarif in cur.fetchall():
            if kat in ter_grouped:
                ter_grouped[kat].append((batas, float(tarif)))
        _TAX_CACHE['ter'] = ter_grouped

        _CACHE_LOADED_AT = time.time()
        cur.close()
    finally:
        if conn:
            conn.close()

def get_ptkp_dict():
    _ensure_cache_loaded()
    return _TAX_CACHE.get('ptkp', {})

def get_ter_rate_from_cache(kategori, bruto_bulanan):
    _ensure_cache_loaded()
    for batas_atas, tarif in _TAX_CACHE.get('ter', {}).get(kategori, []):
        if batas_atas is None or bruto_bulanan <= batas_atas:
            return tarif
    return 34.00

def hitung_pph21_bulanan_ter(status_ptkp, gaji_pokok, tunjangan, bonus_thr, premi_pemberi_kerja):
    ptkp_map = get_ptkp_dict()
    info_ptkp = ptkp_map.get(status_ptkp, {'ptkp': 54_000_000, 'kategori': 'A'})
    kategori = info_ptkp['kategori']

    total_bruto = gaji_pokok + tunjangan + bonus_thr + premi_pemberi_kerja
    tarif_ter = get_ter_rate_from_cache(kategori, total_bruto)
    pph21_terutang = math.floor(total_bruto * (tarif_ter / 100.0))

    return {
        'mode': 'bulanan',
        'status_ptkp': status_ptkp,
        'kategori_ter': kategori,
        'gaji_pokok': gaji_pokok,
        'tunjangan': tunjangan,
        'bonus_thr': bonus_thr,
        'premi_pemberi_kerja': premi_pemberi_kerja,
        'total_bruto': total_bruto,
        'tarif_ter': tarif_ter,
        'pph21_terutang': pph21_terutang
    }

def hitung_pph21_tahunan_pasal17(status_ptkp, bruto_setahun, iuran_pensiun_setahun, zakat_setahun, jumlah_bulan_kerja, pph21_sudah_dipotong):
    _ensure_cache_loaded()

    params = _TAX_CACHE.get('params', {})
    persen_jabatan = params.get('BIAYA_JABATAN_PERSEN', 5.0) / 100.0
    max_jabatan_bln = int(params.get('BIAYA_JABATAN_MAX_BULAN', 500_000))

    ptkp_info = _TAX_CACHE.get('ptkp', {}).get(status_ptkp, {'ptkp': 54_000_000})
    nilai_ptkp = ptkp_info['ptkp']

    lapisan_db = _TAX_CACHE.get('pasal17', [])

    batas_biaya_jabatan = max_jabatan_bln * min(max(jumlah_bulan_kerja, 1), 12)
    biaya_jabatan = min(math.floor(bruto_setahun * persen_jabatan), batas_biaya_jabatan)

    total_pengurang = biaya_jabatan + iuran_pensiun_setahun + zakat_setahun
    neto_setahun = max(bruto_setahun - total_pengurang, 0)

    pkp_raw = max(neto_setahun - nilai_ptkp, 0)
    pkp = math.floor(pkp_raw / 1000) * 1000

    sisa_pkp = pkp
    pph21_setahun = 0
    rincian_lapisan = []

    for kapasitas_raw, tarif_raw, label in lapisan_db:
        if sisa_pkp <= 0:
            break
        kapasitas = int(kapasitas_raw) if kapasitas_raw is not None else float('inf')
        tarif = float(tarif_raw) / 100.0

        kena_lapisan = min(sisa_pkp, kapasitas)
        pajak_lapisan = math.floor(kena_lapisan * tarif)
        pph21_setahun += pajak_lapisan
        rincian_lapisan.append({
            'label': f"{label} - {tarif_raw:g}%",
            'dasar_pengenaan': kena_lapisan,
            'pajak': pajak_lapisan
        })
        sisa_pkp -= kena_lapisan

    selisih_desember = pph21_setahun - pph21_sudah_dipotong

    return {
        'mode': 'tahunan',
        'status_ptkp': status_ptkp,
        'jumlah_bulan_kerja': jumlah_bulan_kerja,
        'bruto_setahun': bruto_setahun,
        'biaya_jabatan': biaya_jabatan,
        'iuran_pensiun_setahun': iuran_pensiun_setahun,
        'zakat_setahun': zakat_setahun,
        'neto_setahun': neto_setahun,
        'nilai_ptkp': nilai_ptkp,
        'pkp': pkp,
        'rincian_lapisan': rincian_lapisan,
        'pph21_setahun': pph21_setahun,
        'pph21_sudah_dipotong': pph21_sudah_dipotong,
        'selisih_desember': selisih_desember
    }