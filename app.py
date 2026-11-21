import os
from datetime import datetime
import psycopg2
from flask import Flask, redirect, url_for, session, render_template, flash, request, send_file, jsonify
from authlib.integrations.flask_client import OAuth
from dotenv import load_dotenv
from werkzeug.middleware.proxy_fix import ProxyFix

from drive_service import upload_file_to_drive
from pph21_service import get_ptkp_dict, hitung_pph21_bulanan_ter, hitung_pph21_tahunan_pasal17
from services.bank_converter import parse_statement_pdf, build_buku_bank_excel
from services.jfs_ledger import (
    append_jfs_transaction_to_gdrive,
    get_gdrive_month_overview,
    upload_attachment_to_drive,
)
from services.project_explorer import project_bp

load_dotenv()

app = Flask(__name__)
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)
app.secret_key = os.getenv("SECRET_KEY", "rahasia")

# Daftarkan Blueprint Folder Project Explorer
app.register_blueprint(project_bp)


@app.template_filter('rupiah')
def format_rupiah(value):
    try:
        return f"{int(value):,}".replace(",", ".")
    except (ValueError, TypeError):
        return "0"


oauth = OAuth(app)
google = oauth.register(
    name='google',
    client_id=os.getenv("GOOGLE_CLIENT_ID"),
    client_secret=os.getenv("GOOGLE_CLIENT_SECRET"),
    server_metadata_url='https://accounts.google.com/.well-known/openid-configuration',
    client_kwargs={'scope': 'openid email profile'}
)


def get_db_connection():
    return psycopg2.connect(os.getenv("DATABASE_URL"))


def get_user_from_db(email):
    conn = None
    try:
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute(
            "SELECT user_id, nama_lengkap, user_role, is_active FROM master_user WHERE email = %s",
            (email,)
        )
        row = cur.fetchone()
        cur.close()
        return row
    finally:
        if conn:
            conn.close()


@app.route('/')
def index():
    if 'user' in session:
        return redirect(url_for('dashboard'))
    return render_template('login.html')


@app.route('/login')
def login():
    redirect_uri = url_for('authorize', _external=True)
    return google.authorize_redirect(redirect_uri)


@app.route('/authorize')
def authorize():
    token = google.authorize_access_token()
    user_info = token.get('userinfo')
    email = user_info.get('email')

    db_user = get_user_from_db(email)

    if not db_user:
        flash(f"Akses Ditolak: Email {email} belum terdaftar di sistem internal.", "error")
        return redirect(url_for('index'))

    user_id, nama_lengkap, user_role, is_active = db_user

    if not is_active:
        flash("Akun lo dinonaktifkan oleh Admin.", "error")
        return redirect(url_for('index'))

    session['user'] = {
        'id': user_id,
        'email': email,
        'nama': nama_lengkap,
        'role': user_role
    }
    return redirect(url_for('dashboard'))


# ==========================================
# DASHBOARD UTAMA = WINDOWS FILE EXPLORER
# ==========================================
@app.route('/dashboard')
def dashboard():
    if 'user' not in session:
        return redirect(url_for('index'))
    return render_template('dashboard.html', user=session['user'])


# ==========================================
# ARSIP DOKUMEN DRIVE (DASHBOARD LAMA)
# ==========================================
@app.route('/arsip-drive')
def arsip_drive():
    if 'user' not in session:
        return redirect(url_for('index'))

    conn = None
    documents = []
    try:
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("""
            SELECT doc_id, nama_klien, npwp, jenis_dokumen, masa_pajak, 
                   nama_file, gdrive_link, uploaded_by, tanggal_upload
            FROM dokumen_pajak
            ORDER BY tanggal_upload DESC
            LIMIT 50
        """)
        documents = cur.fetchall()
        cur.close()
    except Exception as e:
        flash(f"Gagal memuat riwayat dokumen: {str(e)}", "error")
    finally:
        if conn:
            conn.close()

    return render_template('arsip_drive.html', user=session['user'], documents=documents)


@app.route('/upload', methods=['POST'])
def upload_dokumen():
    if 'user' not in session:
        return redirect(url_for('index'))

    nama_klien = request.form.get('nama_klien', '').strip()
    npwp = request.form.get('npwp', '').strip()
    jenis_dokumen = request.form.get('jenis_dokumen', '').strip()
    masa_pajak = request.form.get('masa_pajak', '').strip()
    file = request.files.get('dokumen')

    if not nama_klien or not jenis_dokumen or not masa_pajak or not file or file.filename == '':
        flash("Semua kolom wajib (kecuali NPWP) dan file dokumen harus diisi!", "error")
        return redirect(url_for('arsip_drive'))

    parent_folder_id = os.getenv("GDRIVE_PARENT_FOLDER_ID")
    if not parent_folder_id:
        flash("Konfigurasi GDRIVE_PARENT_FOLDER_ID di .env belum diisi!", "error")
        return redirect(url_for('arsip_drive'))

    conn = None
    try:
        drive_result = upload_file_to_drive(
            file_storage=file,
            filename=file.filename,
            client_name=nama_klien,
            parent_folder_id=parent_folder_id
        )

        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("""
            INSERT INTO dokumen_pajak 
            (nama_klien, npwp, jenis_dokumen, masa_pajak, nama_file, gdrive_file_id, gdrive_link, uploaded_by)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
        """, (
            nama_klien,
            npwp if npwp else '-',
            jenis_dokumen,
            masa_pajak,
            file.filename,
            drive_result['file_id'],
            drive_result['web_link'],
            session['user']['nama']
        ))
        conn.commit()
        cur.close()

        flash(f"Sukses! Dokumen '{file.filename}' berhasil masuk ke folder '{nama_klien}' di Google Drive.", "success")

    except Exception as e:
        if conn:
            conn.rollback()
        flash(f"Terjadi kesalahan saat upload: {str(e)}", "error")
    finally:
        if conn:
            conn.close()

    return redirect(url_for('arsip_drive'))


@app.route('/kalkulator-pph21', methods=['GET', 'POST'])
def kalkulator_pph21():
    if 'user' not in session:
        return redirect(url_for('index'))

    hasil = None
    active_tab = 'bulanan'
    ptkp_list = get_ptkp_dict()

    if request.method == 'POST':
        mode = request.form.get('mode', 'bulanan')
        active_tab = mode
        status_ptkp = request.form.get('status_ptkp', 'TK/0')

        def parse_angka(field_name):
            val = request.form.get(field_name, '0').replace('.', '').replace(',', '').strip()
            return int(val) if val.isdigit() else 0

        if mode == 'bulanan':
            gaji_pokok = parse_angka('gaji_pokok')
            tunjangan = parse_angka('tunjangan')
            bonus_thr = parse_angka('bonus_thr')
            premi_pemberi_kerja = parse_angka('premi_pemberi_kerja')
            hasil = hitung_pph21_bulanan_ter(status_ptkp, gaji_pokok, tunjangan, bonus_thr, premi_pemberi_kerja)

        elif mode == 'tahunan':
            bruto_setahun = parse_angka('bruto_setahun')
            iuran_pensiun = parse_angka('iuran_pensiun')
            zakat = parse_angka('zakat')
            jumlah_bulan = int(request.form.get('jumlah_bulan', 12) or 12)
            pph21_dipotong = parse_angka('pph21_dipotong')
            hasil = hitung_pph21_tahunan_pasal17(
                status_ptkp, bruto_setahun, iuran_pensiun, zakat, jumlah_bulan, pph21_dipotong
            )

    return render_template(
        'kalkulator_pph21.html',
        user=session['user'],
        ptkp_list=ptkp_list,
        hasil=hasil,
        active_tab=active_tab
    )


@app.route('/admin/tarif-pajak', methods=['GET', 'POST'])
def admin_tarif_pajak():
    if 'user' not in session:
        return redirect(url_for('index'))
    if session['user']['role'] != 'Admin':
        flash("Halaman ini hanya dapat diakses oleh Admin.", "error")
        return redirect(url_for('dashboard'))

    kategori_ter = request.args.get('kategori', 'A')
    conn = None
    try:
        conn = get_db_connection()
        cur = conn.cursor()

        if request.method == 'POST':
            action = request.form.get('action')

            if action == 'update_param':
                param_key = request.form.get('param_key')
                param_value = float(request.form.get('param_value'))
                cur.execute("UPDATE ref_parameter_pajak SET param_value = %s WHERE param_key = %s", (param_value, param_key))
                conn.commit()
                flash(f"Parameter {param_key} berhasil diperbarui!", "success")

            elif action == 'update_ptkp':
                kode_ptkp = request.form.get('kode_ptkp')
                nilai_ptkp = int(request.form.get('nilai_ptkp').replace('.', ''))
                kat_ter = request.form.get('kategori_ter')
                cur.execute(
                    "UPDATE ref_ptkp SET nilai_ptkp = %s, kategori_ter = %s WHERE kode_ptkp = %s",
                    (nilai_ptkp, kat_ter, kode_ptkp)
                )
                conn.commit()
                flash(f"Data PTKP {kode_ptkp} berhasil diperbarui!", "success")

            elif action == 'update_pasal17':
                lapisan_id = int(request.form.get('lapisan_id'))
                kap_raw = request.form.get('kapasitas_lapisan', '').replace('.', '').strip()
                kapasitas = int(kap_raw) if kap_raw else None
                tarif = float(request.form.get('tarif_persen'))
                cur.execute(
                    "UPDATE ref_tarif_pasal17 SET kapasitas_lapisan = %s, tarif_persen = %s WHERE lapisan_id = %s",
                    (kapasitas, tarif, lapisan_id)
                )
                conn.commit()
                flash("Lapisan tarif Pasal 17 berhasil diperbarui!", "success")

            elif action == 'update_ter':
                ter_id = int(request.form.get('ter_id'))
                batas_raw = request.form.get('batas_atas', '').replace('.', '').strip()
                batas_atas = int(batas_raw) if batas_raw else None
                tarif = float(request.form.get('tarif_persen'))
                kategori_ter = request.form.get('kategori_aktif', 'A')
                cur.execute(
                    "UPDATE ref_ter_tarif SET batas_atas = %s, tarif_persen = %s WHERE ter_id = %s",
                    (batas_atas, tarif, ter_id)
                )
                conn.commit()
                flash("Tarif TER berhasil diperbarui!", "success")
                return redirect(url_for('admin_tarif_pajak', kategori=kategori_ter))

        cur.execute("SELECT param_key, param_value, deskripsi FROM ref_parameter_pajak ORDER BY param_key")
        params = cur.fetchall()

        cur.execute("SELECT kode_ptkp, nilai_ptkp, kategori_ter, deskripsi FROM ref_ptkp ORDER BY nilai_ptkp, kode_ptkp")
        ptkp_rows = cur.fetchall()

        cur.execute("SELECT lapisan_id, urutan, kapasitas_lapisan, tarif_persen, label FROM ref_tarif_pasal17 ORDER BY urutan")
        pasal17_rows = cur.fetchall()

        cur.execute("""
            SELECT ter_id, kategori, batas_atas, tarif_persen
            FROM ref_ter_tarif
            WHERE kategori = %s
            ORDER BY batas_atas ASC NULLS LAST
        """, (kategori_ter,))
        ter_rows = cur.fetchall()

        cur.close()
        return render_template(
            'admin_tarif.html',
            user=session['user'],
            params=params,
            ptkp_rows=ptkp_rows,
            pasal17_rows=pasal17_rows,
            ter_rows=ter_rows,
            kategori_ter=kategori_ter
        )
    finally:
        if conn:
            conn.close()


@app.route("/buku-bank")
def buku_bank_page():
    return render_template("buku_bank.html")


@app.route("/api/buku-bank/preview", methods=["POST"])
def api_buku_bank_preview():
    if "file" not in request.files:
        return jsonify({"error": "File PDF rekening koran belum dipilih."}), 400

    file = request.files["file"]
    if not file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Format file harus berupa PDF e-Statement."}), 400

    try:
        result = parse_statement_pdf(file.read())
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 400


@app.route("/api/buku-bank/export", methods=["POST"])
def api_buku_bank_export():
    payload = request.get_json(silent=True)
    if not payload or "meta" not in payload or "transactions" not in payload:
        return jsonify({"error": "Data transaksi tidak valid untuk diekspor."}), 400

    excel_stream = build_buku_bank_excel(payload["meta"], payload["transactions"])
    safe_rek = str(payload["meta"].get("no_rekening", "Rekening")).replace(" ", "_")
    return send_file(
        excel_stream,
        as_attachment=True,
        download_name=f"Buku_Bank_{safe_rek}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/jfs-transaksi")
def jfs_transaksi_page():
    return render_template("jfs_transaksi.html")


@app.route("/api/jfs/list", methods=["GET"])
def api_jfs_list():
    try:
        month = int(request.args.get("month", datetime.now().month))
        year = int(request.args.get("year", datetime.now().year))
        return jsonify(get_gdrive_month_overview(month, year))
    except Exception as e:
        return jsonify({"error": str(e)}), 400


@app.route("/api/jfs/save", methods=["POST"])
def api_jfs_save():
    try:
        tanggal = request.form.get("tanggal") or datetime.now().strftime("%Y-%m-%d")

        bukti_url = upload_attachment_to_drive(request.files.get("bukti_transaksi"), "BUKTI", tanggal)
        foto_url = upload_attachment_to_drive(request.files.get("foto_barang"), "FOTO_BARANG", tanggal)

        payload = {
            "tanggal": tanggal,
            "jenis_transaksi": request.form.get("jenis_transaksi", "PENJUALAN"),
            "nama_pengirim": request.form.get("nama_pengirim", ""),
            "no_resi": request.form.get("no_resi", ""),
            "nominal_jfs": request.form.get("nominal_jfs", 0),
            "asuransi": request.form.get("asuransi", 0),
            "packing": request.form.get("packing", 0),
            "jumlah_resi": request.form.get("jumlah_resi", 0),
            "status_bayar": request.form.get("status_bayar", "LUNAS"),
            "metode_bayar": request.form.get("metode_bayar", "BCA MULYATI"),
            "nominal_paid": request.form.get("nominal_paid", 0),
            "tanggal_paid": request.form.get("tanggal_paid", tanggal),
            "biaya_admin": request.form.get("biaya_admin", 0),
            "kategori_pengeluaran": request.form.get("kategori_pengeluaran", ""),
            "arah_mutasi": request.form.get("arah_mutasi", "KREDIT"),
            "catatan": request.form.get("catatan", ""),
            "bukti_url": bukti_url,
            "foto_url": foto_url,
        }
        overview = append_jfs_transaction_to_gdrive(payload)
        return jsonify(overview)
    except Exception as e:
        return jsonify({"error": str(e)}), 400


@app.route('/logout')
def logout():
    session.pop('user', None)
    return redirect(url_for('index'))


if __name__ == '__main__':
    app.run(debug=True)