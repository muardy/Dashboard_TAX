import io
import os
import psycopg2
from psycopg2.extras import RealDictCursor
from flask import Blueprint, request, jsonify, session
from googleapiclient.http import MediaIoBaseUpload
from drive_service import get_drive_service

project_bp = Blueprint("project_bp", __name__)


def _get_db():
    return psycopg2.connect(os.getenv("DATABASE_URL"), cursor_factory=RealDictCursor)


def _create_gdrive_subfolder(folder_name):
    """Membuat subfolder baru di dalam GDRIVE_PARENT_FOLDER_ID."""
    parent_id = os.getenv("GDRIVE_PARENT_FOLDER_ID")
    try:
        drive_svc = get_drive_service()
        metadata = {
            "name": f"[PROJECT] {folder_name}",
            "mimeType": "application/vnd.google-apps.folder",
        }
        if parent_id:
            metadata["parents"] = [parent_id]
        folder = drive_svc.files().create(body=metadata, fields="id").execute()
        return folder.get("id")
    except Exception as e:
        print(f"[WARN] Gagal buat subfolder Drive, fallback ke parent utama: {e}")
        return parent_id


def _recalculate_progress(cur, project_id):
    """Hitung ulang persentase & status otomatis berdasarkan checklist."""
    cur.execute(
        "SELECT COUNT(*) AS total, SUM(CASE WHEN is_done THEN 1 ELSE 0 END) AS done_cnt "
        "FROM project_tasks WHERE project_id = %s",
        (project_id,),
    )
    row = cur.fetchone()
    total = row["total"] or 0
    done_cnt = row["done_cnt"] or 0

    if total > 0:
        pct = round((done_cnt / total) * 100)
        if done_cnt == total:
            status = "DONE"
        elif done_cnt > 0:
            status = "ON_PROGRESS"
        else:
            status = "IDLE"

        cur.execute(
            "UPDATE projects SET progress_pct = %s, status = %s, updated_at = NOW() WHERE project_id = %s",
            (pct, status, project_id),
        )


# ==========================================
# 1. API AMBIL DAFTAR USER AKTIF (MASTER_USER)
# ==========================================
@project_bp.route("/api/users", methods=["GET"])
def api_list_users():
    conn = _get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT user_id, nama_lengkap, email, user_role
                FROM master_user
                WHERE is_active = TRUE
                ORDER BY nama_lengkap ASC
            """)
            users = cur.fetchall()
        return jsonify({"status": "ok", "users": users})
    finally:
        conn.close()


# ==========================================
# 2. API LIST SEMUA PROJECT CARD + ASSIGNEES
# ==========================================
@project_bp.route("/api/projects", methods=["GET"])
def api_list_projects():
    conn = _get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT p.*,
                       (SELECT COUNT(*) FROM project_files f WHERE f.project_id = p.project_id) AS file_count,
                       (SELECT COUNT(*) FROM project_tasks t WHERE t.project_id = p.project_id) AS task_total,
                       (SELECT COUNT(*) FROM project_tasks t WHERE t.project_id = p.project_id AND t.is_done) AS task_done,
                       COALESCE(
                           (SELECT json_agg(
                               json_build_object('user_id', u.user_id, 'nama_lengkap', u.nama_lengkap)
                               ORDER BY u.nama_lengkap ASC
                           )
                           FROM project_assignees pa
                           JOIN master_user u ON u.user_id = pa.user_id
                           WHERE pa.project_id = p.project_id),
                           '[]'::json
                       ) AS assignees
                FROM projects p
                ORDER BY
                    CASE p.status WHEN 'ON_PROGRESS' THEN 1 WHEN 'IDLE' THEN 2 ELSE 3 END,
                    p.updated_at DESC
            """)
            rows = cur.fetchall()
            for r in rows:
                if r.get("updated_at"):
                    r["updated_at"] = r["updated_at"].strftime("%d/%m/%Y %H:%M")
        return jsonify({"status": "ok", "projects": rows})
    finally:
        conn.close()


# ==========================================
# 3. API BUAT PROJECT CARD BARU
# ==========================================
@project_bp.route("/api/projects", methods=["POST"])
def api_create_project():
    data = request.get_json() or {}
    nama = (data.get("nama_project") or "").strip()
    kategori = (data.get("kategori") or "Pajak Klien").strip()
    catatan = (data.get("catatan_terakhir") or "").strip()
    initial_tasks = data.get("tasks") or []
    assignee_ids = data.get("assignee_ids") or []

    if not nama:
        return jsonify({"status": "error", "message": "Nama pekerjaan / project wajib diisi!"}), 400

    gdrive_folder_id = _create_gdrive_subfolder(nama)
    creator = (session.get("user") or {}).get("nama", "Admin")

    conn = _get_db()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO projects (nama_project, kategori, status, progress_pct, catatan_terakhir, gdrive_folder_id, created_by)
                VALUES (%s, %s, 'IDLE', 0, %s, %s, %s)
                RETURNING project_id
                """,
                (nama, kategori, catatan, gdrive_folder_id, creator),
            )
            project_id = cur.fetchone()["project_id"]

            for task_title in initial_tasks:
                t = task_title.strip()
                if t:
                    cur.execute(
                        "INSERT INTO project_tasks (project_id, judul_task) VALUES (%s, %s)",
                        (project_id, t),
                    )

            for uid in assignee_ids:
                cur.execute(
                    "INSERT INTO project_assignees (project_id, user_id) VALUES (%s, %s) ON CONFLICT DO NOTHING",
                    (project_id, int(uid)),
                )

            if assignee_ids:
                _push_project_notification(
                    cur,
                    project_id,
                    assignee_ids,
                    f"menugaskan Anda pada pekerjaan baru \"{nama}\".",
                )

        conn.commit()
        return jsonify({"status": "ok", "project_id": project_id})
    except Exception as e:
        conn.rollback()
        return jsonify({"status": "error", "message": str(e)}), 500
    finally:
        conn.close()


# ==========================================
# 4. API DETAIL PROJECT (SAAT CARD DIKLIK)
# ==========================================
@project_bp.route("/api/projects/<int:project_id>", methods=["GET"])
def api_project_detail(project_id):
    conn = _get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM projects WHERE project_id = %s", (project_id,))
            project = cur.fetchone()
            if not project:
                return jsonify({"status": "error", "message": "Pekerjaan tidak ditemukan"}), 404

            cur.execute(
                "SELECT user_id FROM project_assignees WHERE project_id = %s",
                (project_id,),
            )
            assigned_ids = [r["user_id"] for r in cur.fetchall()]

            cur.execute(
                "SELECT * FROM project_tasks WHERE project_id = %s ORDER BY task_id ASC",
                (project_id,),
            )
            tasks = cur.fetchall()

            cur.execute(
                "SELECT * FROM project_files WHERE project_id = %s ORDER BY uploaded_at DESC",
                (project_id,),
            )
            files = cur.fetchall()
            for f in files:
                if f.get("uploaded_at"):
                    f["uploaded_at"] = f["uploaded_at"].strftime("%d/%m/%Y %H:%M")

        return jsonify({
            "status": "ok",
            "project": project,
            "assigned_ids": assigned_ids,
            "tasks": tasks,
            "files": files,
        })
    finally:
        conn.close()


# ==========================================
# 5. API TOGGLE ASSIGN / UNASSIGN USER
# ==========================================
@project_bp.route("/api/projects/<int:project_id>/assignees/toggle", methods=["POST"])
def api_toggle_assignee(project_id):
    data = request.get_json() or {}
    user_id = data.get("user_id")
    if not user_id:
        return jsonify({"status": "error", "message": "user_id wajib diisi"}), 400

    conn = _get_db()
    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT 1 FROM project_assignees WHERE project_id = %s AND user_id = %s",
                (project_id, int(user_id)),
            )
            exists = cur.fetchone()
            if exists:
                cur.execute(
                    "DELETE FROM project_assignees WHERE project_id = %s AND user_id = %s",
                    (project_id, int(user_id)),
                )
            else:
                cur.execute(
                    "INSERT INTO project_assignees (project_id, user_id) VALUES (%s, %s)",
                    (project_id, int(user_id)),
                )
            cur.execute("UPDATE projects SET updated_at = NOW() WHERE project_id = %s", (project_id,))
        conn.commit()
        return jsonify({"status": "ok"})
    finally:
        conn.close()


# ==========================================
# 6. API UPDATE STATUS & RINGKASAN / CATATAN
# ==========================================
@project_bp.route("/api/projects/<int:project_id>/status", methods=["POST"])
def api_update_project_status(project_id):
    data = request.get_json() or {}
    new_status = data.get("status")
    catatan = data.get("catatan_terakhir")

    conn = _get_db()
    try:
        with conn.cursor() as cur:
            if new_status in ("IDLE", "ON_PROGRESS", "DONE"):
                pct = 100 if new_status == "DONE" else (0 if new_status == "IDLE" else None)
                if pct is not None:
                    cur.execute(
                        "UPDATE projects SET status = %s, progress_pct = %s, updated_at = NOW() WHERE project_id = %s",
                        (new_status, pct, project_id),
                    )
                else:
                    cur.execute(
                        "UPDATE projects SET status = %s, progress_pct = GREATEST(progress_pct, 10), updated_at = NOW() WHERE project_id = %s",
                        (new_status, project_id),
                    )

            if catatan is not None:
                cur.execute(
                    "UPDATE projects SET catatan_terakhir = %s, updated_at = NOW() WHERE project_id = %s",
                    (catatan.strip(), project_id),
                )
        conn.commit()
        return jsonify({"status": "ok"})
    finally:
        conn.close()


# ==========================================
# 7. API TAMBAH & CENTANG CHECKLIST
# ==========================================
@project_bp.route("/api/projects/<int:project_id>/tasks", methods=["POST"])
def api_add_task(project_id):
    data = request.get_json() or {}
    judul = (data.get("judul_task") or "").strip()
    if not judul:
        return jsonify({"status": "error", "message": "Nama tahapan tidak boleh kosong"}), 400

    conn = _get_db()
    try:
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO project_tasks (project_id, judul_task) VALUES (%s, %s)",
                (project_id, judul),
            )
            _recalculate_progress(cur, project_id)
        conn.commit()
        return jsonify({"status": "ok"})
    finally:
        conn.close()


@project_bp.route("/api/projects/tasks/<int:task_id>/toggle", methods=["POST"])
def api_toggle_task(task_id):
    conn = _get_db()
    try:
        with conn.cursor() as cur:
            cur.execute(
                "UPDATE project_tasks SET is_done = NOT is_done WHERE task_id = %s RETURNING project_id",
                (task_id,),
            )
            row = cur.fetchone()
            if row:
                _recalculate_progress(cur, row["project_id"])
        conn.commit()
        return jsonify({"status": "ok"})
    finally:
        conn.close()


# ==========================================
# 7B. API SIMPAN SEMUA PERUBAHAN DETAIL CARD SEKALIGUS
# ==========================================
@project_bp.route("/api/projects/<int:project_id>/save", methods=["POST"])
def api_save_project_changes(project_id):
    data = request.get_json() or {}
    new_status = data.get("status", "IDLE")
    status_manually_set = bool(data.get("status_manually_set", False))
    catatan = (data.get("catatan_terakhir") or "").strip()
    assignee_ids = [int(x) for x in (data.get("assignee_ids") or [])]
    existing_tasks = data.get("existing_tasks") or []
    deleted_task_ids = data.get("deleted_task_ids") or []
    new_tasks = data.get("new_tasks") or []

    conn = _get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT nama_project, status, catatan_terakhir FROM projects WHERE project_id = %s", (project_id,))
            old_proj = cur.fetchone() or {}
            nama_project = old_proj.get("nama_project", "Pekerjaan")

            cur.execute("SELECT user_id FROM project_assignees WHERE project_id = %s", (project_id,))
            old_assignees = set(r["user_id"] for r in cur.fetchall())
            new_assignees_set = set(assignee_ids)

            cur.execute("SELECT task_id, is_done FROM project_tasks WHERE project_id = %s", (project_id,))
            old_task_states = {r["task_id"]: bool(r["is_done"]) for r in cur.fetchall()}

            # 1. Sinkronkan daftar PIC (project_assignees)
            cur.execute("DELETE FROM project_assignees WHERE project_id = %s", (project_id,))
            for uid in new_assignees_set:
                cur.execute(
                    "INSERT INTO project_assignees (project_id, user_id) VALUES (%s, %s)",
                    (project_id, uid),
                )

            # 2. Hapus task yang dibatalkan/dihapus
            for tid in deleted_task_ids:
                cur.execute(
                    "DELETE FROM project_tasks WHERE task_id = %s AND project_id = %s",
                    (int(tid), project_id),
                )

            # 3. Update status centang task lama
            tasks_toggled = False
            for t in existing_tasks:
                tid = int(t["task_id"])
                new_done = bool(t.get("is_done"))
                if tid in old_task_states and old_task_states[tid] != new_done:
                    tasks_toggled = True
                cur.execute(
                    "UPDATE project_tasks SET is_done = %s WHERE task_id = %s AND project_id = %s",
                    (new_done, tid, project_id),
                )

            # 4. Masukkan task baru
            added_task_count = 0
            for nt in new_tasks:
                judul = (nt.get("judul_task") or "").strip()
                if judul:
                    added_task_count += 1
                    cur.execute(
                        "INSERT INTO project_tasks (project_id, judul_task, is_done) VALUES (%s, %s, %s)",
                        (project_id, judul, bool(nt.get("is_done", False))),
                    )

            # 5. Hitung ulang persentase & status
            cur.execute(
                "SELECT COUNT(*) AS total, SUM(CASE WHEN is_done THEN 1 ELSE 0 END) AS done_cnt "
                "FROM project_tasks WHERE project_id = %s",
                (project_id,),
            )
            row = cur.fetchone()
            total = row["total"] or 0
            done_cnt = row["done_cnt"] or 0

            if status_manually_set and new_status in ("IDLE", "ON_PROGRESS", "DONE"):
                final_status = new_status
                if final_status == "DONE":
                    final_pct = 100
                elif final_status == "IDLE":
                    final_pct = 0
                else:
                    final_pct = round((done_cnt / total) * 100) if total > 0 else 10
                    if final_pct == 0:
                        final_pct = 10
            elif total > 0:
                final_pct = round((done_cnt / total) * 100)
                if done_cnt == total:
                    final_status = "DONE"
                elif done_cnt > 0:
                    final_status = "ON_PROGRESS"
                else:
                    final_status = "IDLE"
            else:
                final_status = new_status if new_status in ("IDLE", "ON_PROGRESS", "DONE") else "IDLE"
                final_pct = 100 if final_status == "DONE" else (10 if final_status == "ON_PROGRESS" else 0)

            cur.execute(
                """
                UPDATE projects
                SET status = %s, progress_pct = %s, catatan_terakhir = %s, updated_at = NOW()
                WHERE project_id = %s
                """,
                (final_status, final_pct, catatan, project_id),
            )

            # 6. Kirim notifikasi ke PIC baru & PIC lama yang di-assign
            newly_added_pics = new_assignees_set - old_assignees
            if newly_added_pics:
                _push_project_notification(
                    cur,
                    project_id,
                    newly_added_pics,
                    f"menambahkan Anda sebagai PIC pada pekerjaan \"{nama_project}\".",
                )

            changes_desc = []
            if old_proj.get("status") != final_status:
                changes_desc.append(f"status → {final_status.replace('_', ' ')}")
            if added_task_count > 0:
                changes_desc.append(f"+{added_task_count} tahapan baru")
            if deleted_task_ids:
                changes_desc.append(f"{len(deleted_task_ids)} tahapan dihapus")
            if tasks_toggled:
                changes_desc.append(f"progres tahapan ({done_cnt}/{total})")
            if (old_proj.get("catatan_terakhir") or "").strip() != catatan:
                changes_desc.append("ringkasan diperbarui")

            retained_pics = new_assignees_set & old_assignees
            if retained_pics and changes_desc:
                summary_str = ", ".join(changes_desc)
                _push_project_notification(
                    cur,
                    project_id,
                    retained_pics,
                    f"memperbarui \"{nama_project}\" ({summary_str}).",
                )

        conn.commit()
        return jsonify({"status": "ok"})
    except Exception as e:
        conn.rollback()
        return jsonify({"status": "error", "message": str(e)}), 500
    finally:
        conn.close()


# ==========================================
# 8. API UPLOAD FILE KE GOOGLE DRIVE
# ==========================================
@project_bp.route("/api/projects/<int:project_id>/upload", methods=["POST"])
def api_upload_project_file(project_id):
    if "file" not in request.files:
        return jsonify({"status": "error", "message": "Tidak ada file yang dipilih"}), 400

    uploaded_file = request.files["file"]
    if not uploaded_file.filename:
        return jsonify({"status": "error", "message": "Nama file kosong"}), 400

    filename = uploaded_file.filename
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else "file"
    uploader = (session.get("user") or {}).get("nama", "Admin")

    conn = _get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT nama_project, gdrive_folder_id FROM projects WHERE project_id = %s", (project_id,))
            proj = cur.fetchone()
            if not proj:
                return jsonify({"status": "error", "message": "Project tidak ditemukan"}), 404

            target_folder = proj["gdrive_folder_id"] or os.getenv("GDRIVE_PARENT_FOLDER_ID")

            drive_svc = get_drive_service()
            file_metadata = {"name": filename}
            if target_folder:
                file_metadata["parents"] = [target_folder]

            media = MediaIoBaseUpload(
                io.BytesIO(uploaded_file.read()),
                mimetype=uploaded_file.mimetype or "application/octet-stream",
                resumable=False,
            )
            created = (
                drive_svc.files()
                .create(body=file_metadata, media_body=media, fields="id, webViewLink")
                .execute()
            )

            gdrive_id = created.get("id")
            gdrive_link = created.get("webViewLink") or f"https://drive.google.com/file/d/{gdrive_id}/view"

            cur.execute(
                """
                INSERT INTO project_files (project_id, nama_file, ekstensi, gdrive_file_id, gdrive_link, uploaded_by)
                VALUES (%s, %s, %s, %s, %s, %s)
                """,
                (project_id, filename, ext, gdrive_id, gdrive_link, uploader),
            )
            cur.execute("UPDATE projects SET updated_at = NOW() WHERE project_id = %s", (project_id,))

            cur.execute("SELECT user_id FROM project_assignees WHERE project_id = %s", (project_id,))
            pic_ids = [r["user_id"] for r in cur.fetchall()]
            if pic_ids:
                _push_project_notification(
                    cur,
                    project_id,
                    pic_ids,
                    f"mengunggah dokumen \"{filename}\" ke \"{proj['nama_project']}\".",
                )

        conn.commit()
        return jsonify({"status": "ok"})
    except Exception as e:
        conn.rollback()
        return jsonify({"status": "error", "message": str(e)}), 500
    finally:
        conn.close()


# ==========================================
# 9. API HAPUS PEKERJAAN / PROJECT
# ==========================================
@project_bp.route("/api/projects/<int:project_id>", methods=["DELETE"])
def api_delete_project(project_id):
    conn = _get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT gdrive_folder_id FROM projects WHERE project_id = %s", (project_id,))
            proj = cur.fetchone()
            if not proj:
                return jsonify({"status": "error", "message": "Pekerjaan tidak ditemukan"}), 404

            gdrive_folder_id = proj.get("gdrive_folder_id")
            parent_id = os.getenv("GDRIVE_PARENT_FOLDER_ID")

            # Pindahkan subfolder ke Trash Google Drive (aman, bukan hapus permanen)
            if gdrive_folder_id and gdrive_folder_id != parent_id:
                try:
                    drive_svc = get_drive_service()
                    drive_svc.files().update(
                        fileId=gdrive_folder_id,
                        body={"trashed": True}
                    ).execute()
                except Exception as e:
                    print(f"[WARN] Gagal memindahkan folder Drive ke trash: {e}")

            # Hapus project dari PostgreSQL (tasks, files, assignees otomatis terhapus)
            cur.execute("DELETE FROM projects WHERE project_id = %s", (project_id,))
        conn.commit()
        return jsonify({"status": "ok"})
    except Exception as e:
        conn.rollback()
        return jsonify({"status": "error", "message": str(e)}), 500
    finally:
        conn.close()


# ==========================================
# 10. SISTEM NOTIFIKASI USER (ASSIGNED PROJECTS)
# ==========================================
_NOTIF_TABLE_INITIALIZED = False


def _ensure_notif_table(cur):
    global _NOTIF_TABLE_INITIALIZED
    if _NOTIF_TABLE_INITIALIZED:
        return
    cur.execute("""
        CREATE TABLE IF NOT EXISTS user_notifications (
            notif_id SERIAL PRIMARY KEY,
            user_id INTEGER NOT NULL REFERENCES master_user(user_id) ON DELETE CASCADE,
            project_id INTEGER REFERENCES projects(project_id) ON DELETE CASCADE,
            actor_name VARCHAR(150),
            message TEXT NOT NULL,
            is_read BOOLEAN NOT NULL DEFAULT FALSE,
            created_at TIMESTAMP DEFAULT NOW()
        )
    """)
    _NOTIF_TABLE_INITIALIZED = True


def _get_current_user_info(cur):
    user_sess = session.get("user") or {}
    email = (user_sess.get("email") or "").strip()
    nama = (user_sess.get("nama") or user_sess.get("nama_lengkap") or email or "Rekan Tim").strip()
    uid = user_sess.get("user_id")
    if not uid and email:
        cur.execute("SELECT user_id, nama_lengkap FROM master_user WHERE LOWER(email) = LOWER(%s) LIMIT 1", (email,))
        row = cur.fetchone()
        if row:
            uid = row["user_id"]
            nama = row.get("nama_lengkap") or nama
    return uid, nama


def _push_project_notification(cur, project_id, target_user_ids, message, exclude_actor=True):
    _ensure_notif_table(cur)
    actor_id, actor_name = _get_current_user_info(cur)
    for uid in set(int(x) for x in (target_user_ids or [])):
        if exclude_actor and actor_id and int(uid) == int(actor_id):
            continue
        cur.execute(
            """
            INSERT INTO user_notifications (user_id, project_id, actor_name, message)
            VALUES (%s, %s, %s, %s)
            """,
            (uid, project_id, actor_name, f"{actor_name} {message}"),
        )


@project_bp.route("/api/notifications", methods=["GET"])
def api_get_notifications():
    conn = _get_db()
    try:
        with conn.cursor() as cur:
            _ensure_notif_table(cur)
            uid, _ = _get_current_user_info(cur)
            if not uid:
                conn.commit()
                return jsonify({"status": "ok", "notifications": [], "unread_count": 0})

            cur.execute(
                """
                SELECT n.notif_id, n.project_id, n.actor_name, n.message, n.is_read,
                       TO_CHAR(n.created_at, 'DD Mon YYYY, HH24:MI') AS created_at,
                       p.nama_project
                FROM user_notifications n
                LEFT JOIN projects p ON p.project_id = n.project_id
                WHERE n.user_id = %s
                ORDER BY n.created_at DESC
                LIMIT 25
                """,
                (uid,),
            )
            rows = cur.fetchall()
            unread_count = sum(1 for r in rows if not r["is_read"])
        conn.commit()
        return jsonify({"status": "ok", "notifications": rows, "unread_count": unread_count})
    except Exception as e:
        conn.rollback()
        return jsonify({"status": "error", "message": str(e)}), 500
    finally:
        conn.close()


@project_bp.route("/api/notifications/read", methods=["POST"])
def api_mark_notifications_read():
    data = request.get_json() or {}
    notif_id = data.get("notif_id")
    conn = _get_db()
    try:
        with conn.cursor() as cur:
            _ensure_notif_table(cur)
            uid, _ = _get_current_user_info(cur)
            if uid:
                if notif_id:
                    cur.execute(
                        "UPDATE user_notifications SET is_read = TRUE WHERE notif_id = %s AND user_id = %s",
                        (int(notif_id), uid),
                    )
                else:
                    cur.execute(
                        "UPDATE user_notifications SET is_read = TRUE WHERE user_id = %s AND is_read = FALSE",
                        (uid,),
                    )
        conn.commit()
        return jsonify({"status": "ok"})
    except Exception as e:
        conn.rollback()
        return jsonify({"status": "error", "message": str(e)}), 500
    finally:
        conn.close()

# ==========================================
# 11. API REAL-TIME BOARD & NOTIF SYNC
# ==========================================
@project_bp.route("/api/board-sync", methods=["GET"])
def api_board_sync():
    conn = _get_db()
    try:
        with conn.cursor() as cur:
            _ensure_notif_table(cur)
            uid, _ = _get_current_user_info(cur)

            cur.execute("""
                SELECT p.*,
                       (SELECT COUNT(*) FROM project_files f WHERE f.project_id = p.project_id) AS file_count,
                       (SELECT COUNT(*) FROM project_tasks t WHERE t.project_id = p.project_id) AS task_total,
                       (SELECT COUNT(*) FROM project_tasks t WHERE t.project_id = p.project_id AND t.is_done) AS task_done,
                       COALESCE(
                           (SELECT json_agg(
                               json_build_object('user_id', u.user_id, 'nama_lengkap', u.nama_lengkap)
                               ORDER BY u.nama_lengkap ASC
                           )
                           FROM project_assignees pa
                           JOIN master_user u ON u.user_id = pa.user_id
                           WHERE pa.project_id = p.project_id),
                           '[]'::json
                       ) AS assignees
                FROM projects p
                ORDER BY
                    CASE p.status WHEN 'ON_PROGRESS' THEN 1 WHEN 'IDLE' THEN 2 ELSE 3 END,
                    p.updated_at DESC
            """)
            projects = cur.fetchall()
            for r in projects:
                if r.get("updated_at"):
                    r["updated_at"] = r["updated_at"].strftime("%d/%m/%Y %H:%M")

            notifications = []
            unread_count = 0
            if uid:
                cur.execute(
                    """
                    SELECT n.notif_id, n.project_id, n.actor_name, n.message, n.is_read,
                           TO_CHAR(n.created_at, 'DD Mon YYYY, HH24:MI') AS created_at,
                           p.nama_project
                    FROM user_notifications n
                    LEFT JOIN projects p ON p.project_id = n.project_id
                    WHERE n.user_id = %s
                    ORDER BY n.created_at DESC
                    LIMIT 25
                    """,
                    (uid,),
                )
                notifications = cur.fetchall()
                unread_count = sum(1 for r in notifications if not r["is_read"])

        conn.commit()
        return jsonify({
            "status": "ok",
            "projects": projects,
            "notifications": notifications,
            "unread_count": unread_count,
        })
    except Exception as e:
        conn.rollback()
        return jsonify({"status": "error", "message": str(e)}), 500
    finally:
        conn.close()