import requests
import os
import barcode
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from barcode.writer import ImageWriter
from flask import send_file
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from openpyxl import load_workbook
import io
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, session, flash
from werkzeug.security import check_password_hash, generate_password_hash
from functools import wraps
from db import get_db_connection
from reportlab.lib.styles import ParagraphStyle
from google.oauth2 import service_account
from googleapiclient.discovery import build
import json as json_lib

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key")
app.config['MAX_CONTENT_LENGTH'] = 5 * 1024 * 1024  # 5MB
app.config['SESSION_COOKIE_SECURE'] = os.environ.get('RENDER') is not None
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
from flask_wtf import CSRFProtect
from datetime import timedelta

csrf = CSRFProtect(app)

def kirim_email_stok_kritis():
    mail_username = os.environ.get('MAIL_USERNAME')
    mail_password = os.environ.get('MAIL_PASSWORD')
    mail_to = os.environ.get('MAIL_TO')

    if not all([mail_username, mail_password, mail_to]):
        return False, "Konfigurasi email belum lengkap di environment variables."

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute(
        """SELECT * FROM buku 
           WHERE stok_minimum > 0 AND stok <= stok_minimum 
           ORDER BY stok ASC"""
    )
    stok_menipis = cur.fetchall()
    cur.close()
    conn.close()

    if not stok_menipis:
        return True, "Tidak ada buku dengan stok menipis, email tidak dikirim."

    baris_tabel = ""
    for buku in stok_menipis:
        baris_tabel += f"""
        <tr>
            <td style="padding:6px 10px; border-bottom:1px solid #ddd;">{buku['judul']}</td>
            <td style="padding:6px 10px; border-bottom:1px solid #ddd;">{buku['isbn']}</td>
            <td style="padding:6px 10px; border-bottom:1px solid #ddd; color:#A63A2C; font-weight:bold;">{buku['stok']}</td>
            <td style="padding:6px 10px; border-bottom:1px solid #ddd;">{buku['stok_minimum']}</td>
        </tr>"""

    html_body = f"""
    <html><body style="font-family: Arial, sans-serif;">
        <h2>⚠️ Peringatan Stok Menipis — Inventory Gudang</h2>
        <p>Ada {len(stok_menipis)} buku dengan stok di bawah atau sama dengan batas minimum:</p>
        <table style="border-collapse: collapse; width:100%;">
            <tr style="background:#202A33; color:white;">
                <th style="padding:6px 10px; text-align:left;">Judul</th>
                <th style="padding:6px 10px; text-align:left;">ISBN</th>
                <th style="padding:6px 10px; text-align:left;">Stok</th>
                <th style="padding:6px 10px; text-align:left;">Minimum</th>
            </tr>
            {baris_tabel}
        </table>
        <p style="margin-top:20px; color:#888; font-size:12px;">
            Email otomatis dari sistem Inventory Gudang — dikirim {datetime.now().strftime('%d-%m-%Y %H:%M')}
        </p>
    </body></html>
    """

    msg = MIMEMultipart('alternative')
    msg['Subject'] = f'⚠️ {len(stok_menipis)} Buku Stok Menipis - Inventory Gudang'
    msg['From'] = mail_username
    msg['To'] = mail_to
    msg.attach(MIMEText(html_body, 'html'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(mail_username, mail_password)
        server.sendmail(mail_username, mail_to.split(','), msg.as_string())
        server.quit()
        return True, f"Email berhasil dikirim ke {mail_to} ({len(stok_menipis)} buku)."
    except Exception as e:
        return False, f"Gagal kirim email: {str(e)}"

def ambil_int(form, nama_field, default=0):
    """Ambil nilai integer dari form dengan aman, fallback ke default kalau kosong/tidak valid"""
    nilai = form.get(nama_field, '').strip()
    if not nilai:
        return default
    try:
        return int(float(nilai))
    except (ValueError, TypeError):
        return default
    
def sync_ke_google_sheets():
    creds_json = os.environ.get('GOOGLE_SHEETS_CREDENTIALS')
    sheet_id = os.environ.get('GOOGLE_SHEETS_ID')

    if not creds_json or not sheet_id:
        return False, "Konfigurasi Google Sheets belum lengkap di environment variables."

    try:
        creds_dict = json_lib.loads(creds_json)
        creds = service_account.Credentials.from_service_account_info(
            creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets']
        )
        service = build('sheets', 'v4', credentials=creds)
    except Exception as e:
        return False, f"Gagal autentikasi Google Sheets: {str(e)}"

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM buku ORDER BY judul ASC")
    daftar_buku = cur.fetchall()
    cur.close()
    conn.close()

    header = ['No', 'ISBN', 'Judul', 'Penulis', 'Penerbit', 'Diterima', 'Rencana', 'Stok Minimum', 'Tanggal Masuk', 'Catatan']

    def buat_rows(buku_list):
        rows = [header]
        for i, buku in enumerate(buku_list, start=1):
            rows.append([
                i, buku['isbn'], buku['judul'], buku['penulis'] or '-',
                buku['penerbit'] or '-', buku['stok'], buku['jumlah_rencana'],
                buku['stok_minimum'], str(buku['tanggal_masuk']) if buku['tanggal_masuk'] else '-',
                buku['catatan'] or '-'
            ])
        return rows

    # filter buku dengan status "Sebagian"
    daftar_sebagian = [
        b for b in daftar_buku
        if b['jumlah_rencana'] > 0 and 0 < b['stok'] < b['jumlah_rencana']
    ]

    try:
        sheet = service.spreadsheets()

        # cek tab yang sudah ada
        meta = sheet.get(spreadsheetId=sheet_id).execute()
        tab_ada = [s['properties']['title'] for s in meta['sheets']]

        # buat tab "Sebagian" kalau belum ada
        if 'Sebagian' not in tab_ada:
            sheet.batchUpdate(
                spreadsheetId=sheet_id,
                body={'requests': [{'addSheet': {'properties': {'title': 'Sebagian'}}}]}
            ).execute()

        # kosongkan kedua tab sekaligus
        sheet.values().batchClear(
            spreadsheetId=sheet_id,
            body={'ranges': ['Sheet1', "'Sebagian'"]}
        ).execute()

        # tulis data ke kedua tab sekaligus
        sheet.values().batchUpdate(
            spreadsheetId=sheet_id,
            body={
                'valueInputOption': 'RAW',
                'data': [
                    {'range': 'Sheet1!A1', 'values': buat_rows(daftar_buku)},
                    {'range': "'Sebagian'!A1", 'values': buat_rows(daftar_sebagian)}
                ]
            }
        ).execute()

        return True, f"Berhasil sync {len(daftar_buku)} buku ke Sheet1, dan {len(daftar_sebagian)} buku status Sebagian ke tab terpisah."
    except Exception as e:
        return False, f"Gagal menulis ke Google Sheets: {str(e)}"
# ------------------ DECORATOR: wajib login ------------------
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        if session.get('role') != 'admin':
            flash('Halaman ini khusus untuk admin.', 'danger')
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated_function
def viewer_blocked(f):
    """Blokir role viewer dari akses fitur selain Dashboard & Data Buku"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if session.get('role') == 'viewer':
            flash('Akun ini hanya memiliki akses lihat Data Buku.', 'danger')
            return redirect(url_for('buku_list'))
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

    # total buku, total stok, total rencana
    cur.execute(
        """SELECT COUNT(*) as total_judul, 
                  COALESCE(SUM(stok), 0) as total_stok,
                  COALESCE(SUM(jumlah_rencana), 0) as total_rencana
           FROM buku"""
    )
    ringkasan = cur.fetchone()

    # hitung jumlah buku per status penerimaan
    cur.execute(
        """SELECT 
             COUNT(*) FILTER (WHERE jumlah_rencana > 0 AND stok = 0) as belum_ada,
             COUNT(*) FILTER (WHERE jumlah_rencana > 0 AND stok > 0 AND stok < jumlah_rencana) as sebagian,
             COUNT(*) FILTER (WHERE jumlah_rencana > 0 AND stok >= jumlah_rencana) as lengkap
           FROM buku"""
    )
    status_penerimaan = cur.fetchone()

    # buku dengan stok menipis
    cur.execute(
        """SELECT * FROM buku 
           WHERE stok_minimum > 0 AND stok <= stok_minimum 
           ORDER BY stok ASC LIMIT 10"""
    )
    stok_menipis = cur.fetchall()

    # transaksi hari ini (total eksemplar masuk & keluar)
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

    # jumlah judul buku unik yang masuk hari ini
    cur.execute(
        """SELECT COUNT(DISTINCT buku_id) as total_judul_masuk
           FROM transaksi
           WHERE tipe = 'masuk' AND tanggal = CURRENT_DATE"""
    )
    judul_masuk_hari_ini = cur.fetchone()['total_judul_masuk']

    # 5 transaksi terakhir
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
        aktivitas_terbaru=aktivitas_terbaru,
        status_penerimaan=status_penerimaan,
        judul_masuk_hari_ini=judul_masuk_hari_ini
    )

# ------------------ LIST BUKU ------------------
@app.route('/buku')
@login_required
def buku_list():
    search = request.args.get('search', '').strip()
    status_filter = request.args.get('status', '').strip()
    penerbit_filter = request.args.get('penerbit', '').strip()
    tgl_dari = request.args.get('tgl_dari', '').strip()
    tgl_sampai = request.args.get('tgl_sampai', '').strip()
    catatan_filter = request.args.get('catatan', '').strip()
    catatan_filter = request.args.get('catatan', '').strip()

    conn = get_db_connection()
    cur = conn.cursor()

    query = "SELECT * FROM buku WHERE 1=1"
    params = []

    if search:
        query += " AND (judul ILIKE %s OR isbn ILIKE %s OR penulis ILIKE %s)"
        params.extend([f'%{search}%', f'%{search}%', f'%{search}%'])

    if status_filter == 'belum':
        query += " AND jumlah_rencana > 0 AND stok = 0"
    elif status_filter == 'sebagian':
        query += " AND jumlah_rencana > 0 AND stok > 0 AND stok < jumlah_rencana"
    elif status_filter == 'lengkap':
        query += " AND jumlah_rencana > 0 AND stok >= jumlah_rencana"

    if penerbit_filter:
        query += " AND penerbit = %s"
        params.append(penerbit_filter)

    if tgl_dari:
        query += " AND tanggal_masuk >= %s"
        params.append(tgl_dari)

    if tgl_sampai:
        query += " AND tanggal_masuk <= %s"
        params.append(tgl_sampai)

    if catatan_filter == 'kosong':
        query += " AND (catatan IS NULL OR catatan = '')"
    elif catatan_filter == 'isi':
        query += " AND catatan IS NOT NULL AND catatan != ''"

    query += " ORDER BY judul ASC"

    cur.execute(query, tuple(params))
    daftar_buku = cur.fetchall()

    # daftar penerbit unik untuk dropdown filter
    cur.execute("SELECT DISTINCT penerbit FROM buku WHERE penerbit IS NOT NULL AND penerbit != '' ORDER BY penerbit ASC")
    daftar_penerbit = [row['penerbit'] for row in cur.fetchall()]

    cur.close()
    conn.close()

    return render_template(
        'buku/list.html',
        daftar_buku=daftar_buku,
        search=search,
        status_filter=status_filter,
        penerbit_filter=penerbit_filter,
        daftar_penerbit=daftar_penerbit,
        tgl_dari=tgl_dari,
        tgl_sampai=tgl_sampai,
        catatan_filter=catatan_filter
    )

# ------------------ EXPORT BUKU - EXCEL ------------------
@app.route('/buku/export/excel')
@login_required
def buku_export_excel():
    search = request.args.get('search', '').strip()
    status_filter = request.args.get('status', '').strip()
    penerbit_filter = request.args.get('penerbit', '').strip()
    tgl_dari = request.args.get('tgl_dari', '').strip()
    tgl_sampai = request.args.get('tgl_sampai', '').strip()
    catatan_filter = request.args.get('catatan', '').strip()

    conn = get_db_connection()
    cur = conn.cursor()

    query = "SELECT * FROM buku WHERE 1=1"
    params = []

    if search:
        query += " AND (judul ILIKE %s OR isbn ILIKE %s OR penulis ILIKE %s)"
        params.extend([f'%{search}%', f'%{search}%', f'%{search}%'])

    if status_filter == 'belum':
        query += " AND jumlah_rencana > 0 AND stok = 0"
    elif status_filter == 'sebagian':
        query += " AND jumlah_rencana > 0 AND stok > 0 AND stok < jumlah_rencana"
    elif status_filter == 'lengkap':
        query += " AND jumlah_rencana > 0 AND stok >= jumlah_rencana"

    if penerbit_filter:
        query += " AND penerbit = %s"
        params.append(penerbit_filter)

    if tgl_dari:
        query += " AND tanggal_masuk >= %s"
        params.append(tgl_dari)

    if tgl_sampai:
        query += " AND tanggal_masuk <= %s"
        params.append(tgl_sampai)

    if catatan_filter == 'kosong':
        query += " AND (catatan IS NULL OR catatan = '')"
    elif catatan_filter == 'isi':
        query += " AND catatan IS NOT NULL AND catatan != ''"

    query += " ORDER BY judul ASC"

    cur.execute(query, tuple(params))
    daftar_buku = cur.fetchall()
    cur.close()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Data Buku"

    headers = ['No', 'ISBN', 'Judul', 'Penulis', 'Penerbit', 'Diterima', 'Rencana', 'Stok Minimum', 'Tgl Masuk']
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='0D6EFD', end_color='0D6EFD', fill_type='solid')

    from openpyxl.styles import Alignment
    wrap_style = Alignment(wrap_text=True, vertical='top')

    for i, buku in enumerate(daftar_buku, start=1):
        ws.append([
            i, buku['isbn'], buku['judul'], buku['penulis'] or '-',
            buku['penerbit'] or '-',
            buku['stok'], buku['jumlah_rencana'], buku['stok_minimum'],
            str(buku['tanggal_masuk']) if buku['tanggal_masuk'] else '-'
        ])
        row_num = ws.max_row
        ws.cell(row=row_num, column=3).alignment = wrap_style  # Judul
        ws.cell(row=row_num, column=4).alignment = wrap_style  # Penulis
        ws.cell(row=row_num, column=5).alignment = wrap_style  # Penerbit

    lebar_kolom = {'A': 6, 'B': 16, 'C': 45, 'D': 30, 'E': 25, 'F': 10, 'G': 10, 'H': 13, 'I': 13}
    for kolom, lebar in lebar_kolom.items():
        ws.column_dimensions[kolom].width = lebar

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    label_status = {'belum': 'belum-diterima', 'sebagian': 'sebagian-diterima', 'lengkap': 'lengkap-diterima'}.get(status_filter, 'semua')
    filename = f"data-buku-{label_status}-{datetime.now().strftime('%Y%m%d')}.xlsx"
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
    status_filter = request.args.get('status', '').strip()
    penerbit_filter = request.args.get('penerbit', '').strip()
    tgl_dari = request.args.get('tgl_dari', '').strip()
    tgl_sampai = request.args.get('tgl_sampai', '').strip()

    conn = get_db_connection()
    cur = conn.cursor()

    query = "SELECT * FROM buku WHERE 1=1"
    params = []

    if search:
        query += " AND (judul ILIKE %s OR isbn ILIKE %s OR penulis ILIKE %s)"
        params.extend([f'%{search}%', f'%{search}%', f'%{search}%'])

    if status_filter == 'belum':
        query += " AND jumlah_rencana > 0 AND stok = 0"
    elif status_filter == 'sebagian':
        query += " AND jumlah_rencana > 0 AND stok > 0 AND stok < jumlah_rencana"
    elif status_filter == 'lengkap':
        query += " AND jumlah_rencana > 0 AND stok >= jumlah_rencana"

    if penerbit_filter:
        query += " AND penerbit = %s"
        params.append(penerbit_filter)

    if tgl_dari:
        query += " AND tanggal_masuk >= %s"
        params.append(tgl_dari)

    if tgl_sampai:
        query += " AND tanggal_masuk <= %s"
        params.append(tgl_sampai)

    if catatan_filter == 'kosong':
        query += " AND (catatan IS NULL OR catatan = '')"
    elif catatan_filter == 'isi':
        query += " AND catatan IS NOT NULL AND catatan != ''"

    query += " ORDER BY judul ASC"

    cur.execute(query, tuple(params))
    daftar_buku = cur.fetchall()
    cur.close()
    conn.close()

    output = io.BytesIO()
    doc = SimpleDocTemplate(
        output, pagesize=landscape(A4),
        topMargin=1*cm, bottomMargin=1*cm, leftMargin=1*cm, rightMargin=1*cm
    )
    styles = getSampleStyleSheet()
    elements = []

    judul_map = {'belum': ' - Belum Diterima', 'sebagian': ' - Diterima Sebagian', 'lengkap': ' - Sudah Lengkap'}
    judul_laporan = "Laporan Data Buku" + judul_map.get(status_filter, '')
    elements.append(Paragraph(judul_laporan, styles['Title']))
    elements.append(Paragraph(f"Dicetak: {datetime.now().strftime('%d-%m-%Y %H:%M')}", styles['Normal']))
    elements.append(Spacer(1, 0.5*cm))

    style_sel = ParagraphStyle('sel', fontSize=7.5, leading=9, fontName='Helvetica')
    style_header = ParagraphStyle('header', fontSize=8, leading=10, fontName='Helvetica-Bold', textColor=colors.white)

    data = [[
        Paragraph('No', style_header), Paragraph('ISBN', style_header), Paragraph('Judul', style_header),
        Paragraph('Penulis', style_header), Paragraph('Penerbit', style_header),
        Paragraph('Diterima', style_header), Paragraph('Rencana', style_header),
        Paragraph('Min', style_header), Paragraph('Tgl Masuk', style_header)
    ]]
    for i, buku in enumerate(daftar_buku, start=1):
        data.append([
            Paragraph(str(i), style_sel),
            Paragraph(buku['isbn'], style_sel),
            Paragraph(buku['judul'], style_sel),
            Paragraph(buku['penulis'] or '-', style_sel),
            Paragraph(buku['penerbit'] or '-', style_sel),
            Paragraph(str(buku['stok']), style_sel),
            Paragraph(str(buku['jumlah_rencana']), style_sel),
            Paragraph(str(buku['stok_minimum']), style_sel),
            Paragraph(str(buku['tanggal_masuk']) if buku['tanggal_masuk'] else '-', style_sel),
        ])

    col_widths = [1.0*cm, 2.4*cm, 6.5*cm, 5.0*cm, 4.0*cm, 1.8*cm, 1.8*cm, 1.3*cm, 2.2*cm]

    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0D6EFD')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F2F2F2')]),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    elements.append(table)

    doc.build(elements)
    output.seek(0)

    label_status = {'belum': 'belum-diterima', 'sebagian': 'sebagian-diterima', 'lengkap': 'lengkap-diterima'}.get(status_filter, 'semua')
    filename = f"data-buku-{label_status}-{datetime.now().strftime('%Y%m%d')}.pdf"
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
@admin_required 
def buku_tambah():
    if request.method == 'POST':
        isbn = request.form.get('isbn', '').strip()
        judul = request.form.get('judul', '').strip()
        penulis = request.form.get('penulis', '').strip()
        penerbit = request.form.get('penerbit', '').strip()
        
        stok = ambil_int(request.form, 'stok', 0)
        stok_minimum = ambil_int(request.form, 'stok_minimum', 0)
        jumlah_rencana = ambil_int(request.form, 'jumlah_rencana', 0)
        tanggal_masuk = request.form.get('tanggal_masuk', '').strip() or None
        catatan = request.form.get('catatan', '').strip()

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
            """INSERT INTO buku (isbn, judul, penulis, penerbit, stok, stok_minimum, jumlah_rencana, tanggal_masuk, catatan)
               VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)""",
            (isbn, judul, penulis, penerbit, stok, stok_minimum, jumlah_rencana, tanggal_masuk, catatan)
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
@admin_required
def buku_edit(buku_id):
    conn = get_db_connection()
    cur = conn.cursor()

    if request.method == 'POST':
        isbn = request.form.get('isbn', '').strip()
        judul = request.form.get('judul', '').strip()
        penulis = request.form.get('penulis', '').strip()
        penerbit = request.form.get('penerbit', '').strip()
        stok = ambil_int(request.form, 'stok', 0)
        stok_minimum = ambil_int(request.form, 'stok_minimum', 0)
        jumlah_rencana = ambil_int(request.form, 'jumlah_rencana', 0)
        tanggal_masuk = request.form.get('tanggal_masuk', '').strip() or None
        catatan = request.form.get('catatan', '').strip()

        if not isbn or not judul:
            flash('ISBN dan Judul wajib diisi.', 'danger')
            cur.close()
            conn.close()
            return render_template('buku/form.html', buku=request.form, buku_id=buku_id)

        cur.execute(
            """UPDATE buku 
               SET isbn=%s, judul=%s, penulis=%s, penerbit=%s, 
                   stok=%s, stok_minimum=%s, jumlah_rencana=%s, tanggal_masuk=%s, catatan=%s, updated_at=NOW()
               WHERE id=%s""",
            (isbn, judul, penulis, penerbit, stok, stok_minimum, jumlah_rencana, tanggal_masuk, catatan, buku_id)
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

# ------------------ DETAIL BUKU + RIWAYAT MUTASI ------------------
@app.route('/buku/<int:buku_id>/detail')
@login_required
def buku_detail(buku_id):
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("SELECT * FROM buku WHERE id = %s", (buku_id,))
    buku = cur.fetchone()

    if not buku:
        cur.close()
        conn.close()
        flash('Buku tidak ditemukan.', 'danger')
        return redirect(url_for('buku_list'))

    cur.execute(
        """SELECT t.*, u.nama_lengkap, u.username
           FROM transaksi t
           LEFT JOIN users u ON t.user_id = u.id
           WHERE t.buku_id = %s
           ORDER BY t.tanggal DESC, t.created_at DESC""",
        (buku_id,)
    )
    riwayat = cur.fetchall()

    # ringkasan total masuk & keluar sepanjang waktu untuk buku ini
    cur.execute(
        """SELECT tipe, COALESCE(SUM(jumlah), 0) as total
           FROM transaksi WHERE buku_id = %s GROUP BY tipe""",
        (buku_id,)
    )
    ringkasan_raw = cur.fetchall()
    ringkasan = {'masuk': 0, 'keluar': 0}
    for row in ringkasan_raw:
        ringkasan[row['tipe']] = row['total']

    cur.close()
    conn.close()

    return render_template('buku/detail.html', buku=buku, riwayat=riwayat, ringkasan=ringkasan)

# ------------------ HAPUS BUKU ------------------
@app.route('/buku/hapus/<int:buku_id>', methods=['POST'])
@login_required
@admin_required
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
@viewer_blocked
def transaksi_masuk():
    conn = get_db_connection()
    cur = conn.cursor()

    if request.method == 'POST':
        buku_id = request.form.get('buku_id', '').strip()
        jumlah = ambil_int(request.form, 'jumlah', 0)
        catatan_buku = request.form.get('catatan_buku', '').strip()
        pihak_terkait = request.form.get('pihak_terkait', '').strip()
        tanggal = request.form.get('tanggal', '').strip()

        if not buku_id or jumlah <= 0:
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
                """INSERT INTO transaksi (buku_id, tipe, jumlah, pihak_terkait, user_id, tanggal)
                   VALUES (%s, 'masuk', %s, %s, %s, %s)""",
                (buku_id, jumlah, pihak_terkait, session['user_id'], tanggal or None)
            )

            # update stok buku + tanggal masuk terakhir + catatan
            tgl_transaksi = tanggal or datetime.now().date()
            cur.execute(
                """UPDATE buku 
                   SET stok = stok + %s, updated_at = NOW(),
                       tanggal_masuk = GREATEST(COALESCE(tanggal_masuk, %s), %s),
                       catatan = %s
                   WHERE id = %s""",
                (jumlah, tgl_transaksi, tgl_transaksi, catatan_buku, buku_id)
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
@viewer_blocked
def transaksi_keluar():
    conn = get_db_connection()
    cur = conn.cursor()

    if request.method == 'POST':
        buku_id = request.form.get('buku_id', '').strip()
        jumlah = ambil_int(request.form, 'jumlah', 0)
        keterangan = request.form.get('keterangan', '').strip()
        pihak_terkait = request.form.get('pihak_terkait', '').strip()
        tanggal = request.form.get('tanggal', '').strip()
        if not buku_id or jumlah <= 0:
            flash('Buku dan jumlah (harus lebih dari 0) wajib diisi.', 'danger')
            cur.execute("SELECT * FROM buku ORDER BY judul ASC")
            daftar_buku = cur.fetchall()
            cur.close()
            conn.close()
            return render_template('transaksi/keluar.html', daftar_buku=daftar_buku)

        try:
            cur.execute("SELECT * FROM buku WHERE id = %s", (buku_id,))
            buku = cur.fetchone()

            if not buku:
                flash('Buku tidak ditemukan.', 'danger')
                cur.close()
                conn.close()
                return redirect(url_for('transaksi_keluar'))

            if buku['stok'] < jumlah:
                flash(f'Stok tidak cukup. Stok tersedia: {buku["stok"]}, diminta: {jumlah}.', 'danger')
                cur.close()
                conn.close()
                return redirect(url_for('transaksi_keluar'))

            cur.execute(
                """INSERT INTO transaksi (buku_id, tipe, jumlah, keterangan, pihak_terkait, user_id, tanggal)
                   VALUES (%s, 'keluar', %s, %s, %s, %s, %s)""",
                (buku_id, jumlah, keterangan, pihak_terkait, session['user_id'], tanggal or None)
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
@viewer_blocked
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
    # ------------------ API: GENERATE KODE INTERNAL BARU ------------------
@app.route('/api/kode-internal-baru')
@login_required
def api_kode_internal_baru():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT isbn FROM buku WHERE isbn LIKE 'GDG-%' ORDER BY isbn DESC LIMIT 1")
    last = cur.fetchone()
    cur.close()
    conn.close()

    last_num = 0
    if last:
        try:
            last_num = int(last['isbn'].split('-')[1])
        except (IndexError, ValueError):
            last_num = 0

    return {'kode': f"GDG-{last_num + 1:04d}"}


# ------------------ GENERATE GAMBAR BARCODE ------------------
@app.route('/buku/<int:buku_id>/barcode.png')
@login_required
def buku_barcode_image(buku_id):
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT isbn FROM buku WHERE id = %s", (buku_id,))
    buku = cur.fetchone()
    cur.close()
    conn.close()

    if not buku:
        return '', 404

    code128 = barcode.get('code128', buku['isbn'], writer=ImageWriter())
    output = io.BytesIO()
    code128.write(output, options={'write_text': False, 'module_height': 8.0, 'quiet_zone': 2})
    output.seek(0)
    return send_file(output, mimetype='image/png')


# ------------------ HALAMAN PILIH BUKU UNTUK CETAK LABEL ------------------
@app.route('/buku/cetak-label', methods=['GET', 'POST'])
@login_required
def buku_cetak_label():
    conn = get_db_connection()
    cur = conn.cursor()

    if request.method == 'POST':
        ids = request.form.getlist('buku_ids')
        ids = [int(i) for i in ids if i.isdigit()]

        if not ids:
            flash('Pilih minimal satu buku untuk dicetak labelnya.', 'danger')
            cur.execute("SELECT * FROM buku ORDER BY judul ASC")
            daftar_buku = cur.fetchall()
            cur.close()
            conn.close()
            return render_template('buku/pilih_label.html', daftar_buku=daftar_buku)

        cur.execute("SELECT * FROM buku WHERE id = ANY(%s) ORDER BY judul ASC", (ids,))
        terpilih = cur.fetchall()
        cur.close()
        conn.close()
        return render_template('buku/label.html', daftar_buku=terpilih)

    cur.execute("SELECT * FROM buku ORDER BY judul ASC")
    daftar_buku = cur.fetchall()
    cur.close()
    conn.close()
    return render_template('buku/pilih_label.html', daftar_buku=daftar_buku)


# ------------------ CETAK LABEL UNTUK 1 BUKU (dari halaman edit) ------------------
@app.route('/buku/<int:buku_id>/label')
@login_required
def buku_label(buku_id):
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM buku WHERE id = %s", (buku_id,))
    buku = cur.fetchone()
    cur.close()
    conn.close()

    if not buku:
        flash('Buku tidak ditemukan.', 'danger')
        return redirect(url_for('buku_list'))

    return render_template('buku/label.html', daftar_buku=[buku])

# ------------------ DAFTAR PEMINJAMAN AKTIF ------------------
@app.route('/peminjaman')
@login_required
@viewer_blocked
def peminjaman_list():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute(
        """SELECT t.*, b.judul, b.isbn
           FROM transaksi t
           JOIN buku b ON t.buku_id = b.id
           WHERE t.tipe = 'keluar' AND t.jenis_keluar = 'pinjam' 
                 AND t.tanggal_kembali_aktual IS NULL
           ORDER BY t.tanggal_kembali_rencana ASC"""
    )
    daftar_pinjam = cur.fetchall()
    cur.close()
    conn.close()
    return render_template('peminjaman.html', daftar_pinjam=daftar_pinjam, today=datetime.now().date())


# ------------------ TANDAI SUDAH KEMBALI ------------------
@app.route('/peminjaman/<int:transaksi_id>/kembali', methods=['POST'])
@login_required
@viewer_blocked
def peminjaman_kembali(transaksi_id):
    conn = get_db_connection()
    cur = conn.cursor()

    try:
        cur.execute(
            """SELECT * FROM transaksi 
               WHERE id = %s AND tipe = 'keluar' AND jenis_keluar = 'pinjam' 
                     AND tanggal_kembali_aktual IS NULL""",
            (transaksi_id,)
        )
        transaksi = cur.fetchone()

        if not transaksi:
            flash('Data peminjaman tidak ditemukan atau sudah ditandai kembali.', 'danger')
        else:
            cur.execute(
                "UPDATE transaksi SET tanggal_kembali_aktual = CURRENT_DATE WHERE id = %s",
                (transaksi_id,)
            )
            cur.execute(
                "UPDATE buku SET stok = stok + %s, updated_at = NOW() WHERE id = %s",
                (transaksi['jumlah'], transaksi['buku_id'])
            )
            conn.commit()
            flash('Berhasil ditandai sudah kembali, stok bertambah kembali.', 'success')
    except Exception as e:
        conn.rollback()
        flash('Gagal memproses. Coba lagi.', 'danger')
    finally:
        cur.close()
        conn.close()

    return redirect(url_for('peminjaman_list'))
# ------------------ ADMIN: DAFTAR USER ------------------
@app.route('/admin/users')
@login_required
@admin_required
def user_list():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM users ORDER BY username ASC")
    daftar_user = cur.fetchall()
    cur.close()
    conn.close()
    return render_template('admin/users.html', daftar_user=daftar_user)


# ------------------ ADMIN: TAMBAH USER ------------------
@app.route('/admin/users/tambah', methods=['GET', 'POST'])
@login_required
@admin_required
def user_tambah():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        nama_lengkap = request.form.get('nama_lengkap', '').strip()
        role = request.form.get('role', 'staff').strip()

        if not username or not password:
            flash('Username dan password wajib diisi.', 'danger')
            return render_template('admin/user_form.html')

        if len(password) < 8:
            flash('Password minimal 8 karakter.', 'danger')
            return render_template('admin/user_form.html')

        conn = get_db_connection()
        cur = conn.cursor()

        cur.execute("SELECT id FROM users WHERE username = %s", (username,))
        if cur.fetchone():
            flash('Username sudah dipakai.', 'danger')
            cur.close()
            conn.close()
            return render_template('admin/user_form.html')

        password_hash = generate_password_hash(password)
        cur.execute(
            """INSERT INTO users (username, password_hash, nama_lengkap, role)
               VALUES (%s, %s, %s, %s)""",
            (username, password_hash, nama_lengkap, role)
        )
        conn.commit()
        cur.close()
        conn.close()

        flash(f'User "{username}" berhasil ditambahkan.', 'success')
        return redirect(url_for('user_list'))

    return render_template('admin/user_form.html')


# ------------------ ADMIN: AKTIFKAN/NONAKTIFKAN USER ------------------
@app.route('/admin/users/<int:user_id>/toggle', methods=['POST'])
@login_required
@admin_required
def user_toggle(user_id):
    if user_id == session['user_id']:
        flash('Tidak bisa menonaktifkan akun sendiri.', 'danger')
        return redirect(url_for('user_list'))

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT is_active FROM users WHERE id = %s", (user_id,))
    user = cur.fetchone()

    if user:
        cur.execute("UPDATE users SET is_active = %s WHERE id = %s", (not user['is_active'], user_id))
        conn.commit()
        flash('Status user berhasil diubah.', 'success')

    cur.close()
    conn.close()
    return redirect(url_for('user_list'))


# ------------------ ADMIN: RESET PASSWORD USER ------------------
@app.route('/admin/users/<int:user_id>/reset-password', methods=['POST'])
@login_required
@admin_required
def user_reset_password(user_id):
    password_baru = request.form.get('password_baru', '').strip()

    if len(password_baru) < 8:
        flash('Password minimal 8 karakter.', 'danger')
        return redirect(url_for('user_list'))

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute(
        "UPDATE users SET password_hash = %s, failed_attempts = 0, locked_until = NULL WHERE id = %s",
        (generate_password_hash(password_baru), user_id)
    )
    conn.commit()
    cur.close()
    conn.close()

    flash('Password berhasil direset.', 'success')
    return redirect(url_for('user_list'))
# ------------------ ADMIN: KIRIM NOTIFIKASI MANUAL ------------------
@app.route('/admin/kirim-notifikasi-stok', methods=['POST'])
@login_required
@admin_required
def kirim_notifikasi_stok_manual():
    sukses, pesan = kirim_email_stok_kritis()
    flash(pesan, 'success' if sukses else 'danger')
    return redirect(url_for('dashboard'))


# ------------------ CRON: DIPANGGIL SCHEDULER EKSTERNAL ------------------
@app.route('/cron/cek-stok-kritis')
def cron_cek_stok_kritis():
    secret = request.args.get('secret', '')
    if secret != os.environ.get('CRON_SECRET'):
        return {'error': 'unauthorized'}, 401

    sukses, pesan = kirim_email_stok_kritis()
    return {'success': sukses, 'message': pesan}

# ------------------ DOWNLOAD TEMPLATE IMPORT ------------------
@app.route('/buku/import/template')
@login_required
@admin_required
def buku_import_template():
    wb = Workbook()
    ws = wb.active
    ws.title = "Template Import"

    headers = ['isbn', 'judul', 'penulis', 'penerbit', 'jumlah_rencana', 'stok', 'stok_minimum']
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='0D6EFD', end_color='0D6EFD', fill_type='solid')

    ws.append(['9786020633178', 'Contoh Judul Buku', 'Nama Penulis', 'Nama Penerbit', 40, 0, 3])

    for col in ws.columns:
        max_length = max(len(str(cell.value)) for cell in col if cell.value)
        ws.column_dimensions[col[0].column_letter].width = max_length + 3

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='template-import-buku.xlsx'
    )

# ------------------ IMPORT BUKU DARI EXCEL ------------------
@app.route('/buku/import', methods=['GET', 'POST'])
@login_required
@admin_required
def buku_import():
    if request.method == 'POST':
        file = request.files.get('file_excel')

        if not file or file.filename == '':
            flash('Pilih file Excel dulu.', 'danger')
            return render_template('buku/import.html')

        if not file.filename.endswith(('.xlsx', '.xls')):
            flash('File harus berformat .xlsx atau .xls', 'danger')
            return render_template('buku/import.html')

        try:
            wb = load_workbook(file, data_only=True)
            ws = wb.active
        except Exception as e:
            flash('Gagal membaca file Excel. Pastikan formatnya benar.', 'danger')
            return render_template('buku/import.html')

        # alias nama kolom yang dikenali (huruf kecil, sudah di-strip)
        alias_kolom = {
            'isbn': ['isbn'],
            'judul': ['judul', 'judul buku'],
            'penulis': ['penulis'],
            'penerbit': ['penerbit'],
            'jumlah_rencana': ['eksemplar', 'jumlah rencana', 'jumlah_rencana', 'rencana'],
            'stok': ['stok', 'stok awal'],
            'stok_minimum': ['stok_minimum', 'stok minimum'],
        }

        # cari baris header: scan 15 baris pertama, cari baris yang punya sel "isbn"
        header_row_num = None
        kolom_index = {}

        for row_num in range(1, min(16, ws.max_row + 1)):
            row_values = [str(cell.value).strip().lower() if cell.value else '' for cell in ws[row_num]]
            if 'isbn' in row_values:
                header_row_num = row_num
                for field, kemungkinan_nama in alias_kolom.items():
                    for nama in kemungkinan_nama:
                        if nama in row_values:
                            kolom_index[field] = row_values.index(nama)
                            break
                break

        if header_row_num is None or 'isbn' not in kolom_index or 'judul' not in kolom_index:
            flash('Kolom ISBN dan Judul tidak ditemukan di file. Pastikan ada baris header dengan kolom "ISBN" dan "Judul"/"Judul Buku".', 'danger')
            return render_template('buku/import.html')

        conn = get_db_connection()
        cur = conn.cursor()

        berhasil = 0
        dilewati = []
        baris_ke = header_row_num

        for row in ws.iter_rows(min_row=header_row_num + 1, values_only=True):
            baris_ke += 1

            def ambil(nama_field, default=''):
                idx = kolom_index.get(nama_field)
                if idx is None or idx >= len(row) or row[idx] is None:
                    return default
                return row[idx]

            isbn = str(ambil('isbn')).strip()
            judul = str(ambil('judul')).strip()

            if not isbn or not judul or isbn.lower() == 'none':
                continue  # baris kosong, dilewati diam-diam (bukan error)

            cur.execute("SELECT id FROM buku WHERE isbn = %s", (isbn,))
            if cur.fetchone():
                dilewati.append(f"Baris {baris_ke}: ISBN {isbn} sudah ada")
                continue

            try:
                jumlah_rencana = int(float(ambil('jumlah_rencana', 0) or 0))
            except (ValueError, TypeError):
                jumlah_rencana = 0
            try:
                stok = int(float(ambil('stok', 0) or 0))
            except (ValueError, TypeError):
                stok = 0
            try:
                stok_minimum = int(float(ambil('stok_minimum', 0) or 0))
            except (ValueError, TypeError):
                stok_minimum = 0

            cur.execute(
                """INSERT INTO buku (isbn, judul, penulis, penerbit, stok, stok_minimum, jumlah_rencana)
                   VALUES (%s, %s, %s, %s, %s, %s, %s)""",
                (isbn, judul, str(ambil('penulis')), str(ambil('penerbit')),
                 stok, stok_minimum, jumlah_rencana)
            )
            berhasil += 1

        conn.commit()
        cur.close()
        conn.close()

        pesan = f'{berhasil} buku berhasil diimpor.'
        if dilewati:
            pesan += f' {len(dilewati)} baris dilewati (ISBN dobel).'
        flash(pesan, 'success' if berhasil > 0 else 'danger')

        return render_template('buku/import.html', dilewati=dilewati, berhasil=berhasil)

    return render_template('buku/import.html')

# ------------------ DOWNLOAD TEMPLATE IMPORT MASSAL BARANG MASUK ------------------
@app.route('/transaksi/masuk/import/template')
@login_required
@viewer_blocked
def import_masuk_template():
    wb = Workbook()
    ws = wb.active
    ws.title = "Template Barang Masuk"

    headers = ['isbn', 'jumlah_masuk', 'tanggal', 'keterangan']
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='0D6EFD', end_color='0D6EFD', fill_type='solid')

    ws.append(['9786347345264', 40, '2026-08-12', 'Kiriman tahap 1'])

    for col in ws.columns:
        max_length = max(len(str(cell.value)) for cell in col if cell.value)
        ws.column_dimensions[col[0].column_letter].width = max_length + 3

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='template-barang-masuk-massal.xlsx'
    )
# ------------------ IMPORT MASSAL BARANG MASUK ------------------
@app.route('/transaksi/masuk/import', methods=['GET', 'POST'])
@login_required
@viewer_blocked
@admin_required
def import_masuk_massal():
    if request.method == 'POST':
        file = request.files.get('file_excel')

        if not file or file.filename == '':
            flash('Pilih file Excel dulu.', 'danger')
            return render_template('transaksi/import_masuk.html')

        if not file.filename.endswith(('.xlsx', '.xls')):
            flash('File harus berformat .xlsx atau .xls', 'danger')
            return render_template('transaksi/import_masuk.html')

        try:
            wb = load_workbook(file, data_only=True)
            ws = wb.active
        except Exception:
            flash('Gagal membaca file Excel. Pastikan formatnya benar.', 'danger')
            return render_template('transaksi/import_masuk.html')

        header_row = [str(cell.value).strip().lower() if cell.value else '' for cell in ws[1]]

        if 'isbn' not in header_row or 'jumlah_masuk' not in header_row:
            flash('Kolom "isbn" dan "jumlah_masuk" wajib ada di baris pertama. Gunakan template yang disediakan.', 'danger')
            return render_template('transaksi/import_masuk.html')

        idx_isbn = header_row.index('isbn')
        idx_jumlah = header_row.index('jumlah_masuk')
        idx_keterangan = header_row.index('keterangan') if 'keterangan' in header_row else None
        idx_tanggal = header_row.index('tanggal') if 'tanggal' in header_row else None

        conn = get_db_connection()
        cur = conn.cursor()

        berhasil = []
        gagal = []
        baris_ke = 1

        for row in ws.iter_rows(min_row=2, values_only=True):
            baris_ke += 1

            isbn = str(row[idx_isbn]).strip() if idx_isbn < len(row) and row[idx_isbn] else ''
            if not isbn:
                continue

            try:
                jumlah = int(float(row[idx_jumlah])) if idx_jumlah < len(row) and row[idx_jumlah] else 0
            except (ValueError, TypeError):
                jumlah = 0

            keterangan = ''
            if idx_keterangan is not None and idx_keterangan < len(row) and row[idx_keterangan]:
                keterangan = str(row[idx_keterangan])

            tgl_masuk = datetime.now().date()
            if idx_tanggal is not None and idx_tanggal < len(row) and row[idx_tanggal]:
                nilai_tanggal = row[idx_tanggal]
                if hasattr(nilai_tanggal, 'date'):
                    tgl_masuk = nilai_tanggal.date()
                elif isinstance(nilai_tanggal, str):
                    try:
                        tgl_masuk = datetime.strptime(nilai_tanggal.strip(), '%Y-%m-%d').date()
                    except ValueError:
                        pass

            if jumlah <= 0:
                gagal.append(f"Baris {baris_ke}: jumlah tidak valid untuk ISBN {isbn}")
                continue

            cur.execute("SELECT * FROM buku WHERE isbn = %s", (isbn,))
            buku = cur.fetchone()

            if not buku:
                gagal.append(f"Baris {baris_ke}: ISBN {isbn} tidak ditemukan di master data")
                continue

            try:
                cur.execute(
                    """INSERT INTO transaksi (buku_id, tipe, jumlah, keterangan, user_id, tanggal)
                       VALUES (%s, 'masuk', %s, %s, %s, %s)""",
                    (buku['id'], jumlah, keterangan or 'Import massal barang masuk', session['user_id'], tgl_masuk)
                )
                cur.execute(
                    """UPDATE buku 
                       SET stok = stok + %s, updated_at = NOW(),
                           tanggal_masuk = GREATEST(COALESCE(tanggal_masuk, %s), %s)
                       WHERE id = %s""",
                    (jumlah, tgl_masuk, tgl_masuk, buku['id'])
                )
                berhasil.append(f"{buku['judul']} (+{jumlah})")
            except Exception:
                gagal.append(f"Baris {baris_ke}: gagal memproses ISBN {isbn}")

        conn.commit()
        cur.close()
        conn.close()

        pesan = f'{len(berhasil)} buku berhasil diupdate stoknya.'
        if gagal:
            pesan += f' {len(gagal)} baris gagal/dilewati.'
        flash(pesan, 'success' if berhasil else 'danger')

        return render_template('transaksi/import_masuk.html', berhasil=berhasil, gagal=gagal)

    return render_template('transaksi/import_masuk.html')

# ------------------ PROGRESS PENERIMAAN PER PENERBIT ------------------
@app.route('/buku/progress-penerbit')
@login_required
@viewer_blocked
def progress_penerbit():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute(
        """SELECT 
             COALESCE(NULLIF(penerbit, ''), '(Tanpa Penerbit)') as penerbit,
             COUNT(*) as total_judul,
             COUNT(*) FILTER (WHERE jumlah_rencana > 0 AND stok >= jumlah_rencana) as judul_lengkap,
             COALESCE(SUM(stok), 0) as total_diterima,
             COALESCE(SUM(jumlah_rencana), 0) as total_rencana
           FROM buku
           GROUP BY COALESCE(NULLIF(penerbit, ''), '(Tanpa Penerbit)')
           ORDER BY penerbit ASC"""
    )
    data_penerbit = cur.fetchall()
    cur.close()
    conn.close()

    # hitung persentase & urutkan dari yang paling tertinggal
    hasil = []
    for p in data_penerbit:
        persen_eks = (p['total_diterima'] / p['total_rencana'] * 100) if p['total_rencana'] > 0 else 0
        persen_judul = (p['judul_lengkap'] / p['total_judul'] * 100) if p['total_judul'] > 0 else 0
        hasil.append({
            'penerbit': p['penerbit'],
            'total_judul': p['total_judul'],
            'judul_lengkap': p['judul_lengkap'],
            'total_diterima': p['total_diterima'],
            'total_rencana': p['total_rencana'],
            'persen_eks': round(persen_eks, 1),
            'persen_judul': round(persen_judul, 1)
        })

    # klasifikasi berdasarkan judul: lengkap, sebagian, kosong
    for item in hasil:
        if item['total_judul'] > 0 and item['judul_lengkap'] >= item['total_judul']:
            item['status_label'] = 'Lengkap'
        elif item['judul_lengkap'] > 0:
            item['status_label'] = 'Sebagian'
        else:
            item['status_label'] = 'Kosong'

    grup_lengkap = sorted([x for x in hasil if x['status_label'] == 'Lengkap'], key=lambda x: -x['persen_judul'])
    grup_sebagian = sorted([x for x in hasil if x['status_label'] == 'Sebagian'], key=lambda x: -x['persen_judul'])
    grup_kosong = sorted([x for x in hasil if x['status_label'] == 'Kosong'], key=lambda x: x['penerbit'])

    return render_template(
        'buku/progress_penerbit.html',
        grup_lengkap=grup_lengkap,
        grup_sebagian=grup_sebagian,
        grup_kosong=grup_kosong
    )

# ------------------ ADMIN: KELOLA PENERBIT ------------------
@app.route('/admin/penerbit')
@login_required
@admin_required
def kelola_penerbit():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute(
        """SELECT COALESCE(NULLIF(penerbit, ''), '(Tanpa Penerbit)') as penerbit, COUNT(*) as jumlah_buku
           FROM buku
           GROUP BY COALESCE(NULLIF(penerbit, ''), '(Tanpa Penerbit)')
           ORDER BY penerbit ASC"""
    )
    daftar_penerbit = cur.fetchall()
    cur.close()
    conn.close()
    return render_template('admin/penerbit.html', daftar_penerbit=daftar_penerbit)


# ------------------ ADMIN: GABUNG/GANTI NAMA PENERBIT ------------------
@app.route('/admin/penerbit/gabung', methods=['POST'])
@login_required
@admin_required
def gabung_penerbit():
    penerbit_lama = request.form.get('penerbit_lama', '').strip()
    penerbit_baru = request.form.get('penerbit_baru', '').strip()

    if not penerbit_lama or not penerbit_baru:
        flash('Pilih penerbit lama dan isi nama penerbit baru.', 'danger')
        return redirect(url_for('kelola_penerbit'))

    if penerbit_lama == penerbit_baru:
        flash('Nama baru sama dengan nama lama, tidak ada yang diubah.', 'danger')
        return redirect(url_for('kelola_penerbit'))

    conn = get_db_connection()
    cur = conn.cursor()

    try:
        cur.execute(
            "UPDATE buku SET penerbit = %s WHERE penerbit = %s",
            (penerbit_baru, penerbit_lama)
        )
        jumlah_terupdate = cur.rowcount
        conn.commit()
        print(f"[GABUNG PENERBIT] oleh {session.get('username')} pada {datetime.now()}: '{penerbit_lama}' -> '{penerbit_baru}' ({jumlah_terupdate} buku)")
        flash(f'{jumlah_terupdate} buku dari "{penerbit_lama}" berhasil diubah jadi "{penerbit_baru}".', 'success')
    except Exception as e:
        conn.rollback()
        flash('Gagal memproses. Coba lagi.', 'danger')
    finally:
        cur.close()
        conn.close()

    return redirect(url_for('kelola_penerbit'))

@app.route('/buku/sync-sheets', methods=['POST'])
@login_required
@viewer_blocked
@admin_required
def buku_sync_sheets():
    sukses, pesan = sync_ke_google_sheets()
    flash(pesan, 'success' if sukses else 'danger')
    return redirect(url_for('buku_list'))   

@app.errorhandler(413)
def file_terlalu_besar(e):
    flash('File terlalu besar. Maksimal ukuran file adalah 5MB.', 'danger')
    return redirect(request.referrer or url_for('dashboard'))
if __name__ == '__main__':
    app.run(debug=True)