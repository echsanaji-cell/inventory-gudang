import requests
import os
from flask import send_file
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
import io
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, session, flash
from werkzeug.security import check_password_hash
from functools import wraps
from db import get_db_connection

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key")
from flask_wtf import CSRFProtect
from datetime import timedelta

csrf = CSRFProtect(app)

# ------------------ DECORATOR: wajib login ------------------
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

# ------------------ LOGIN ------------------
MAX_FAILED_ATTEMPTS = 5
LOCKOUT_MINUTES = 15

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')

        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute(
            "SELECT * FROM users WHERE username = %s AND is_active = TRUE",
            (username,)
        )
        user = cur.fetchone()

        if not user:
            flash('Username atau password salah.', 'danger')
            cur.close()
            conn.close()
            return render_template('login.html')

        # cek apakah akun sedang terkunci
        if user['locked_until'] and user['locked_until'] > datetime.now():
            sisa_menit = int((user['locked_until'] - datetime.now()).total_seconds() / 60) + 1
            flash(f'Akun terkunci karena terlalu banyak percobaan gagal. Coba lagi dalam {sisa_menit} menit.', 'danger')
            cur.close()
            conn.close()
            return render_template('login.html')

        if check_password_hash(user['password_hash'], password):
            # login berhasil: reset failed_attempts
            cur.execute(
                "UPDATE users SET failed_attempts = 0, locked_until = NULL WHERE id = %s",
                (user['id'],)
            )
            conn.commit()

            session['user_id'] = user['id']
            session['username'] = user['username']
            session['role'] = user['role']
            session['nama_lengkap'] = user['nama_lengkap']

            cur.close()
            conn.close()
            return redirect(url_for('dashboard'))
        else:
            # login gagal: tambah counter
            new_attempts = user['failed_attempts'] + 1

            if new_attempts >= MAX_FAILED_ATTEMPTS:
                locked_until = datetime.now() + timedelta(minutes=LOCKOUT_MINUTES)
                cur.execute(
                    "UPDATE users SET failed_attempts = %s, locked_until = %s WHERE id = %s",
                    (new_attempts, locked_until, user['id'])
                )
                conn.commit()
                flash(f'Terlalu banyak percobaan gagal. Akun dikunci selama {LOCKOUT_MINUTES} menit.', 'danger')
            else:
                cur.execute(
                    "UPDATE users SET failed_attempts = %s WHERE id = %s",
                    (new_attempts, user['id'])
                )
                conn.commit()
                sisa = MAX_FAILED_ATTEMPTS - new_attempts
                flash(f'Username atau password salah. Percobaan tersisa: {sisa}.', 'danger')

        cur.close()
        conn.close()

    return render_template('login.html')

# ------------------ LOGOUT ------------------
@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# ------------------ DASHBOARD ------------------
@app.route('/')
@login_required
def dashboard():
    conn = get_db_connection()
    cur = conn.cursor()

    # total buku & total stok
    cur.execute("SELECT COUNT(*) as total_judul, COALESCE(SUM(stok), 0) as total_stok FROM buku")
    ringkasan = cur.fetchone()

    # buku dengan stok menipis (stok <= stok_minimum, dan stok_minimum > 0 biar buku tanpa batas minimum nggak ikut muncul)
    cur.execute(
        """SELECT * FROM buku 
           WHERE stok_minimum > 0 AND stok <= stok_minimum 
           ORDER BY stok ASC LIMIT 10"""
    )
    stok_menipis = cur.fetchall()

    # transaksi hari ini
    cur.execute(
        """SELECT tipe, COUNT(*) as jumlah_transaksi, COALESCE(SUM(jumlah), 0) as total_item
           FROM transaksi
           WHERE tanggal = CURRENT_DATE
           GROUP BY tipe"""
    )
    transaksi_hari_ini_raw = cur.fetchall()
    transaksi_hari_ini = {'masuk': {'jumlah_transaksi': 0, 'total_item': 0},
                           'keluar': {'jumlah_transaksi': 0, 'total_item': 0}}
    for row in transaksi_hari_ini_raw:
        transaksi_hari_ini[row['tipe']] = row

    # 5 transaksi terakhir (aktivitas terbaru)
    cur.execute(
        """SELECT t.*, b.judul
           FROM transaksi t
           JOIN buku b ON t.buku_id = b.id
           ORDER BY t.created_at DESC
           LIMIT 5"""
    )
    aktivitas_terbaru = cur.fetchall()

    cur.close()
    conn.close()

    return render_template(
        'dashboard.html',
        ringkasan=ringkasan,
        stok_menipis=stok_menipis,
        transaksi_hari_ini=transaksi_hari_ini,
        aktivitas_terbaru=aktivitas_terbaru
    )

# ------------------ LIST BUKU ------------------
@app.route('/buku')
@login_required
def buku_list():
    search = request.args.get('search', '').strip()

    conn = get_db_connection()
    cur = conn.cursor()

    if search:
        cur.execute(
            """SELECT * FROM buku 
               WHERE judul ILIKE %s OR isbn ILIKE %s OR penulis ILIKE %s
               ORDER BY judul ASC""",
            (f'%{search}%', f'%{search}%', f'%{search}%')
        )
    else:
        cur.execute("SELECT * FROM buku ORDER BY judul ASC")

    daftar_buku = cur.fetchall()
    cur.close()
    conn.close()

    return render_template('buku/list.html', daftar_buku=daftar_buku, search=search)

# ------------------ EXPORT BUKU - EXCEL ------------------
@app.route('/buku/export/excel')
@login_required
def buku_export_excel():
    search = request.args.get('search', '').strip()

    conn = get_db_connection()
    cur = conn.cursor()
    if search:
        cur.execute(
            """SELECT * FROM buku 
               WHERE judul ILIKE %s OR isbn ILIKE %s OR penulis ILIKE %s
               ORDER BY judul ASC""",
            (f'%{search}%', f'%{search}%', f'%{search}%')
        )
    else:
        cur.execute("SELECT * FROM buku ORDER BY judul ASC")
    daftar_buku = cur.fetchall()
    cur.close()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Data Buku"

    headers = ['ISBN', 'Judul', 'Penulis', 'Penerbit', 'Kategori', 'Stok', 'Stok Minimum']
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='0D6EFD', end_color='0D6EFD', fill_type='solid')

    for buku in daftar_buku:
        ws.append([
            buku['isbn'], buku['judul'], buku['penulis'] or '-',
            buku['penerbit'] or '-', buku['kategori'] or '-',
            buku['stok'], buku['stok_minimum']
        ])

    # lebar kolom otomatis biar rapi
    for col in ws.columns:
        max_length = max(len(str(cell.value)) for cell in col if cell.value)
        ws.column_dimensions[col[0].column_letter].width = max_length + 3

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"data-buku-{datetime.now().strftime('%Y%m%d')}.xlsx"
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )
# ------------------ EXPORT BUKU - PDF ------------------
@app.route('/buku/export/pdf')
@login_required
def buku_export_pdf():
    search = request.args.get('search', '').strip()

    conn = get_db_connection()
    cur = conn.cursor()
    if search:
        cur.execute(
            """SELECT * FROM buku 
               WHERE judul ILIKE %s OR isbn ILIKE %s OR penulis ILIKE %s
               ORDER BY judul ASC""",
            (f'%{search}%', f'%{search}%', f'%{search}%')
        )
    else:
        cur.execute("SELECT * FROM buku ORDER BY judul ASC")
    daftar_buku = cur.fetchall()
    cur.close()
    conn.close()

    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=landscape(A4), topMargin=1*cm, bottomMargin=1*cm)
    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph("Laporan Data Buku", styles['Title']))
    elements.append(Paragraph(f"Dicetak: {datetime.now().strftime('%d-%m-%Y %H:%M')}", styles['Normal']))
    elements.append(Spacer(1, 0.5*cm))

    data = [['ISBN', 'Judul', 'Penulis', 'Penerbit', 'Kategori', 'Stok', 'Min']]
    for buku in daftar_buku:
        data.append([
            buku['isbn'], buku['judul'], buku['penulis'] or '-',
            buku['penerbit'] or '-', buku['kategori'] or '-',
            str(buku['stok']), str(buku['stok_minimum'])
        ])

    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0D6EFD')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F2F2F2')]),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    elements.append(table)

    doc.build(elements)
    output.seek(0)

    filename = f"data-buku-{datetime.now().strftime('%Y%m%d')}.pdf"
    return send_file(output, mimetype='application/pdf', as_attachment=True, download_name=filename)
# ------------------ EXPORT TRANSAKSI - EXCEL ------------------
@app.route('/transaksi/export/excel')
@login_required
def transaksi_export_excel():
    tipe_filter = request.args.get('tipe', '').strip()

    conn = get_db_connection()
    cur = conn.cursor()
    query = """
        SELECT t.*, b.judul, b.isbn, u.nama_lengkap, u.username
        FROM transaksi t
        JOIN buku b ON t.buku_id = b.id
        LEFT JOIN users u ON t.user_id = u.id
    """
    params = []
    if tipe_filter in ('masuk', 'keluar'):
        query += " WHERE t.tipe = %s"
        params.append(tipe_filter)
    query += " ORDER BY t.tanggal DESC, t.created_at DESC"
    cur.execute(query, tuple(params))
    daftar_transaksi = cur.fetchall()
    cur.close()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Riwayat Transaksi"

    headers = ['Tanggal', 'Tipe', 'ISBN', 'Judul', 'Jumlah', 'Pihak Terkait', 'Keterangan', 'Dicatat Oleh']
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='0D6EFD', end_color='0D6EFD', fill_type='solid')

    for t in daftar_transaksi:
        ws.append([
            str(t['tanggal']), t['tipe'].capitalize(), t['isbn'], t['judul'],
            t['jumlah'], t['pihak_terkait'] or '-', t['keterangan'] or '-',
            t['nama_lengkap'] or t['username'] or '-'
        ])

    for col in ws.columns:
        max_length = max(len(str(cell.value)) for cell in col if cell.value)
        ws.column_dimensions[col[0].column_letter].width = max_length + 3

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"riwayat-transaksi-{datetime.now().strftime('%Y%m%d')}.xlsx"
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


# ------------------ EXPORT TRANSAKSI - PDF ------------------
@app.route('/transaksi/export/pdf')
@login_required
def transaksi_export_pdf():
    tipe_filter = request.args.get('tipe', '').strip()

    conn = get_db_connection()
    cur = conn.cursor()
    query = """
        SELECT t.*, b.judul, b.isbn, u.nama_lengkap, u.username
        FROM transaksi t
        JOIN buku b ON t.buku_id = b.id
        LEFT JOIN users u ON t.user_id = u.id
    """
    params = []
    if tipe_filter in ('masuk', 'keluar'):
        query += " WHERE t.tipe = %s"
        params.append(tipe_filter)
    query += " ORDER BY t.tanggal DESC, t.created_at DESC"
    cur.execute(query, tuple(params))
    daftar_transaksi = cur.fetchall()
    cur.close()
    conn.close()

    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=landscape(A4), topMargin=1*cm, bottomMargin=1*cm)
    styles = getSampleStyleSheet()
    elements = []

    judul_laporan = "Laporan Riwayat Transaksi"
    if tipe_filter == 'masuk':
        judul_laporan += " - Barang Masuk"
    elif tipe_filter == 'keluar':
        judul_laporan += " - Barang Keluar"

    elements.append(Paragraph(judul_laporan, styles['Title']))
    elements.append(Paragraph(f"Dicetak: {datetime.now().strftime('%d-%m-%Y %H:%M')}", styles['Normal']))
    elements.append(Spacer(1, 0.5*cm))

    data = [['Tanggal', 'Tipe', 'ISBN', 'Judul', 'Jml', 'Pihak Terkait', 'Keterangan', 'Oleh']]
    for t in daftar_transaksi:
        data.append([
            str(t['tanggal']), t['tipe'].capitalize(), t['isbn'], t['judul'],
            str(t['jumlah']), t['pihak_terkait'] or '-', t['keterangan'] or '-',
            t['nama_lengkap'] or t['username'] or '-'
        ])

    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0D6EFD')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F2F2F2')]),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    elements.append(table)

    doc.build(elements)
    output.seek(0)

    filename = f"riwayat-transaksi-{datetime.now().strftime('%Y%m%d')}.pdf"
    return send_file(output, mimetype='application/pdf', as_attachment=True, download_name=filename)
# ------------------ TAMBAH BUKU ------------------
@app.route('/buku/tambah', methods=['GET', 'POST'])
@login_required
def buku_tambah():
    if request.method == 'POST':
        isbn = request.form.get('isbn', '').strip()
        judul = request.form.get('judul', '').strip()
        penulis = request.form.get('penulis', '').strip()
        penerbit = request.form.get('penerbit', '').strip()
        kategori = request.form.get('kategori', '').strip()
        stok = request.form.get('stok', '0').strip()
        stok_minimum = request.form.get('stok_minimum', '0').strip()

        if not isbn or not judul:
            flash('ISBN dan Judul wajib diisi.', 'danger')
            return render_template('buku/form.html', buku=request.form)

        conn = get_db_connection()
        cur = conn.cursor()

        # cek ISBN sudah ada atau belum
        cur.execute("SELECT id FROM buku WHERE isbn = %s", (isbn,))
        existing = cur.fetchone()
        if existing:
            flash(f'ISBN {isbn} sudah terdaftar di database.', 'danger')
            cur.close()
            conn.close()
            return render_template('buku/form.html', buku=request.form)

        cur.execute(
            """INSERT INTO buku (isbn, judul, penulis, penerbit, kategori, stok, stok_minimum)
               VALUES (%s, %s, %s, %s, %s, %s, %s)""",
            (isbn, judul, penulis, penerbit, kategori, stok, stok_minimum)
        )
        conn.commit()
        cur.close()
        conn.close()

        flash(f'Buku "{judul}" berhasil ditambahkan.', 'success')
        return redirect(url_for('buku_list'))

    return render_template('buku/form.html', buku=None)


# ------------------ EDIT BUKU ------------------
@app.route('/buku/edit/<int:buku_id>', methods=['GET', 'POST'])
@login_required
def buku_edit(buku_id):
    conn = get_db_connection()
    cur = conn.cursor()

    if request.method == 'POST':
        isbn = request.form.get('isbn', '').strip()
        judul = request.form.get('judul', '').strip()
        penulis = request.form.get('penulis', '').strip()
        penerbit = request.form.get('penerbit', '').strip()
        kategori = request.form.get('kategori', '').strip()
        stok = request.form.get('stok', '0').strip()
        stok_minimum = request.form.get('stok_minimum', '0').strip()

        if not isbn or not judul:
            flash('ISBN dan Judul wajib diisi.', 'danger')
            cur.close()
            conn.close()
            return render_template('buku/form.html', buku=request.form, buku_id=buku_id)

        cur.execute(
            """UPDATE buku 
               SET isbn=%s, judul=%s, penulis=%s, penerbit=%s, kategori=%s, 
                   stok=%s, stok_minimum=%s, updated_at=NOW()
               WHERE id=%s""",
            (isbn, judul, penulis, penerbit, kategori, stok, stok_minimum, buku_id)
        )
        conn.commit()
        cur.close()
        conn.close()

        flash(f'Buku "{judul}" berhasil diupdate.', 'success')
        return redirect(url_for('buku_list'))

    cur.execute("SELECT * FROM buku WHERE id = %s", (buku_id,))
    buku = cur.fetchone()
    cur.close()
    conn.close()

    if not buku:
        flash('Buku tidak ditemukan.', 'danger')
        return redirect(url_for('buku_list'))

    return render_template('buku/form.html', buku=buku, buku_id=buku_id)


# ------------------ HAPUS BUKU ------------------
@app.route('/buku/hapus/<int:buku_id>', methods=['POST'])
@login_required
def buku_hapus(buku_id):
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("SELECT judul FROM buku WHERE id = %s", (buku_id,))
    buku = cur.fetchone()

    if not buku:
        flash('Buku tidak ditemukan.', 'danger')
    else:
        try:
            cur.execute("DELETE FROM buku WHERE id = %s", (buku_id,))
            conn.commit()
            flash(f'Buku "{buku["judul"]}" berhasil dihapus.', 'success')
        except Exception as e:
            conn.rollback()
            flash('Gagal menghapus buku — mungkin masih ada transaksi terkait.', 'danger')

    cur.close()
    conn.close()
    return redirect(url_for('buku_list'))

# ------------------ BARANG MASUK ------------------
@app.route('/transaksi/masuk', methods=['GET', 'POST'])
@login_required
def transaksi_masuk():
    conn = get_db_connection()
    cur = conn.cursor()

    if request.method == 'POST':
        buku_id = request.form.get('buku_id', '').strip()
        jumlah = request.form.get('jumlah', '').strip()
        keterangan = request.form.get('keterangan', '').strip()
        pihak_terkait = request.form.get('pihak_terkait', '').strip()
        tanggal = request.form.get('tanggal', '').strip()

        if not buku_id or not jumlah or int(jumlah) <= 0:
            flash('Buku dan jumlah (harus lebih dari 0) wajib diisi.', 'danger')
            cur.execute("SELECT * FROM buku ORDER BY judul ASC")
            daftar_buku = cur.fetchall()
            cur.close()
            conn.close()
            return render_template('transaksi/masuk.html', daftar_buku=daftar_buku)

        try:
            # cek buku ada
            cur.execute("SELECT * FROM buku WHERE id = %s", (buku_id,))
            buku = cur.fetchone()
            if not buku:
                flash('Buku tidak ditemukan.', 'danger')
                cur.close()
                conn.close()
                return redirect(url_for('transaksi_masuk'))

            # insert transaksi
            cur.execute(
                """INSERT INTO transaksi (buku_id, tipe, jumlah, keterangan, pihak_terkait, user_id, tanggal)
                   VALUES (%s, 'masuk', %s, %s, %s, %s, %s)""",
                (buku_id, jumlah, keterangan, pihak_terkait, session['user_id'],
                 tanggal or None)
            )

            # update stok buku
            cur.execute(
                "UPDATE buku SET stok = stok + %s, updated_at = NOW() WHERE id = %s",
                (jumlah, buku_id)
            )

            conn.commit()
            flash(f'Barang masuk: {buku["judul"]} +{jumlah} berhasil dicatat.', 'success')
        except Exception as e:
            conn.rollback()
            flash('Gagal mencatat transaksi. Coba lagi.', 'danger')
        finally:
            cur.close()
            conn.close()

        return redirect(url_for('transaksi_masuk'))

    cur.execute("SELECT * FROM buku ORDER BY judul ASC")
    daftar_buku = cur.fetchall()
    cur.close()
    conn.close()
    return render_template('transaksi/masuk.html', daftar_buku=daftar_buku)


# ------------------ BARANG KELUAR ------------------
@app.route('/transaksi/keluar', methods=['GET', 'POST'])
@login_required
def transaksi_keluar():
    conn = get_db_connection()
    cur = conn.cursor()

    if request.method == 'POST':
        buku_id = request.form.get('buku_id', '').strip()
        jumlah = request.form.get('jumlah', '').strip()
        keterangan = request.form.get('keterangan', '').strip()
        pihak_terkait = request.form.get('pihak_terkait', '').strip()
        tanggal = request.form.get('tanggal', '').strip()

        if not buku_id or not jumlah or int(jumlah) <= 0:
            flash('Buku dan jumlah (harus lebih dari 0) wajib diisi.', 'danger')
            cur.execute("SELECT * FROM buku ORDER BY judul ASC")
            daftar_buku = cur.fetchall()
            cur.close()
            conn.close()
            return render_template('transaksi/keluar.html', daftar_buku=daftar_buku)

        jumlah = int(jumlah)

        try:
            cur.execute("SELECT * FROM buku WHERE id = %s", (buku_id,))
            buku = cur.fetchone()

            if not buku:
                flash('Buku tidak ditemukan.', 'danger')
                cur.close()
                conn.close()
                return redirect(url_for('transaksi_keluar'))

            # validasi stok cukup
            if buku['stok'] < jumlah:
                flash(f'Stok tidak cukup. Stok tersedia: {buku["stok"]}, diminta: {jumlah}.', 'danger')
                cur.close()
                conn.close()
                return redirect(url_for('transaksi_keluar'))

            cur.execute(
                """INSERT INTO transaksi (buku_id, tipe, jumlah, keterangan, pihak_terkait, user_id, tanggal)
                   VALUES (%s, 'keluar', %s, %s, %s, %s, %s)""",
                (buku_id, jumlah, keterangan, pihak_terkait, session['user_id'],
                 tanggal or None)
            )

            cur.execute(
                "UPDATE buku SET stok = stok - %s, updated_at = NOW() WHERE id = %s",
                (jumlah, buku_id)
            )

            conn.commit()
            flash(f'Barang keluar: {buku["judul"]} -{jumlah} berhasil dicatat.', 'success')
        except Exception as e:
            conn.rollback()
            flash('Gagal mencatat transaksi. Coba lagi.', 'danger')
        finally:
            cur.close()
            conn.close()

        return redirect(url_for('transaksi_keluar'))

    cur.execute("SELECT * FROM buku ORDER BY judul ASC")
    daftar_buku = cur.fetchall()
    cur.close()
    conn.close()
    return render_template('transaksi/keluar.html', daftar_buku=daftar_buku)


# ------------------ RIWAYAT TRANSAKSI ------------------
@app.route('/transaksi/riwayat')
@login_required
def transaksi_riwayat():
    tipe_filter = request.args.get('tipe', '').strip()

    conn = get_db_connection()
    cur = conn.cursor()

    query = """
        SELECT t.*, b.judul, b.isbn, u.nama_lengkap, u.username
        FROM transaksi t
        JOIN buku b ON t.buku_id = b.id
        LEFT JOIN users u ON t.user_id = u.id
    """
    params = []

    if tipe_filter in ('masuk', 'keluar'):
        query += " WHERE t.tipe = %s"
        params.append(tipe_filter)

    query += " ORDER BY t.tanggal DESC, t.created_at DESC"

    cur.execute(query, tuple(params))
    daftar_transaksi = cur.fetchall()
    cur.close()
    conn.close()

    return render_template('transaksi/riwayat.html', daftar_transaksi=daftar_transaksi, tipe_filter=tipe_filter)

# ------------------ API: LOOKUP ISBN (Google Books) ------------------
@app.route('/api/isbn/<isbn>')
@login_required
def api_lookup_isbn(isbn):
    isbn = isbn.strip().replace('-', '').replace(' ', '')

    try:
        resp = requests.get(
            f'https://www.googleapis.com/books/v1/volumes?q=isbn:{isbn}',
            timeout=5
        )
        data = resp.json()

        if data.get('totalItems', 0) > 0:
            info = data['items'][0]['volumeInfo']
            return {
                'found': True,
                'judul': info.get('title', ''),
                'penulis': ', '.join(info.get('authors', [])),
                'penerbit': info.get('publisher', ''),
                'sampul_url': info.get('imageLinks', {}).get('thumbnail', '')
            }
        else:
            return {'found': False}
    except Exception as e:
        return {'found': False, 'error': str(e)}
if __name__ == '__main__':
    app.run(debug=True)