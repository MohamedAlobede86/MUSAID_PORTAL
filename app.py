from flask import Flask, render_template, request, redirect, url_for, session, flash, send_from_directory
import sqlite3
import os
import time
from werkzeug.utils import secure_filename
import pypdf
import pytesseract
from docx import Document
from pptx import Presentation
from PIL import Image
import fitz 
import psycopg2
import psycopg2.extras

app = Flask(__name__)
app.secret_key = 'musaid_secret_key' 

# -----------------------------
# الإعدادات والدوال المساعدة
# -----------------------------
DATABASE = 'musaid_ist.db'
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

import psycopg2.extras # تأكد من وجود هذا السطر في أعلى الملف

def get_db_connection():
    database_url = os.environ.get('DATABASE_URL')
    if database_url:
        # تصحيح رابط PostgreSQL
        if database_url.startswith("postgresql://"):
            database_url = database_url.replace("postgresql://", "postgres://", 1)
        
        # إنشاء الاتصال مع خاصية التثبيت التلقائي (Autocommit) لتقليل الضغط
        conn = psycopg2.connect(database_url, sslmode='require')
        conn.autocommit = True
        
        # فئة محاكاة لـ SQLite لضمان عمل كودك القديم (266 أو 576 سطر)
        class DBWrapper:
            def __init__(self, connection):
                self.connection = connection
            def execute(self, sql, params=None):
                # تحويل العلامات تلقائياً من ? إلى %s
                sql = sql.replace('?', '%s')
                cursor = self.connection.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
                cursor.execute(sql, params)
                return cursor
            def commit(self): pass # Autocommit مفعل
            def close(self): self.connection.close()
            
        return DBWrapper(conn)
    else:
        # الاتصال المحلي في طبرق
        conn = sqlite3.connect('musaid_ist.db')
        conn.row_factory = sqlite3.Row
        return conn

def clean_text(text):
    import re
    text = re.sub(r'[^\u0600-\u06FFa-zA-Z0-9\s.,!?؟]', '', text)
    text = re.sub(r'(.)\1{2,}', r'\1', text)
    return text.strip()

def extract_text_from_file(file_path):
    text = ""
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext == ".pdf":
            try:
                reader = pypdf.PdfReader(file_path)
                for page in reader.pages:
                    text += page.extract_text() or ""
            except:
                doc = fitz.open(file_path)
                for page in doc:
                    pix = page.get_pixmap()
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    text += pytesseract.image_to_string(img, lang="ara+eng")
        elif ext == ".docx":
            doc = Document(file_path)
            for para in doc.paragraphs:
                if para.text.strip(): text += para.text + "\n"
        elif ext == ".pptx":
            prs = Presentation(file_path)
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text.strip():
                        text += shape.text + "\n"
    except Exception as e:
        text = f"⚠️ خطأ في القراءة: {e}"
    return clean_text(text)

def summarize_handout(file_name):
    full_path = os.path.join(app.config['UPLOAD_FOLDER'], file_name)
    if not os.path.exists(full_path): return "الملف غير موجود."
    content = extract_text_from_file(full_path)
    if not content or len(content.strip()) < 20: return "لا يوجد محتوى كافي للتلخيص."
    import re
    sentences = re.split(r'[.!?؟]', content)
    return " ".join(sentences[:3]).strip()

# -----------------------------
# المسارات (Routes)
# -----------------------------

@app.route('/download/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=False)

@app.route('/')
def index():
    conn = get_db_connection()
    db_url = os.environ.get('DATABASE_URL')
    if db_url:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute('SELECT * FROM departments')
        depts = cur.fetchall()
        cur.close()
    else:
        depts = conn.execute('SELECT * FROM departments').fetchall()
    conn.close()
    return render_template('index.html', depts=depts)

@app.route('/search')
def search():
    dept_id = request.args.get('dept')
    semester = request.args.get('semester')
    conn = get_db_connection()
    db_url = os.environ.get('DATABASE_URL')
    
    query = '''
        SELECT h.*, s.subject_name, t.full_name as teacher_name
        FROM handouts h
        JOIN subjects s ON h.subject_id = s.id
        JOIN teachers t ON h.teacher_id = t.id
        WHERE h.dept_id = %s AND h.semester = %s
    ''' if db_url else '''
        SELECT h.*, s.subject_name, t.full_name as teacher_name
        FROM handouts h
        JOIN subjects s ON h.subject_id = s.id
        JOIN teachers t ON h.teacher_id = t.id
        WHERE h.dept_id = ? AND h.semester = ?
    '''

    if db_url:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute(query, (dept_id, semester))
        results = cur.fetchall()
        cur.execute('SELECT dept_name FROM departments WHERE id = %s', (dept_id,))
        dept_row = cur.fetchone()
        cur.close()
    else:
        results = conn.execute(query, (dept_id, semester)).fetchall()
        dept_row = conn.execute('SELECT dept_name FROM departments WHERE id = ?', (dept_id,)).fetchone()

    processed_results = []
    for row in results:
        item = dict(row)
        item['flash_summary'] = summarize_handout(item.get('file_path', ''))
        processed_results.append(item)
    
    conn.close()
    return render_template('results.html', results=processed_results, dept_name=dept_row['dept_name'] if dept_row else "غير معروف", semester=semester)

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form['email'].strip()
        password = request.form['password'].strip()
        conn = get_db_connection()
        db_url = os.environ.get('DATABASE_URL')
        
        sql = 'SELECT * FROM teachers WHERE email = %s AND password = %s' if db_url else 'SELECT * FROM teachers WHERE email = ? AND password = ?'
        
        if db_url:
            cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            cur.execute(sql, (email, password))
            user = cur.fetchone()
            cur.close()
        else:
            user = conn.execute(sql, (email, password)).fetchone()
        
        conn.close()
        if user:
            session['user_id'] = user['id']
            session['user_name'] = user['full_name']
            session['role'] = 'admin' if email == 'admin@musaid.edu.ly' else 'teacher'
            return redirect(url_for('admin_dashboard' if session['role'] == 'admin' else 'teacher_dashboard'))
        flash('خطأ في البيانات')
    return render_template('login.html')

@app.route('/teacher')
def teacher_dashboard():
    if 'user_id' not in session or session['role'] != 'teacher': return redirect(url_for('login'))
    conn = get_db_connection()
    db_url = os.environ.get('DATABASE_URL')
    
    sql_h = '''SELECT h.*, s.subject_name FROM handouts h JOIN subjects s ON h.subject_id = s.id 
               WHERE h.teacher_id = %s ORDER BY h.id DESC''' if db_url else \
            '''SELECT h.*, s.subject_name FROM handouts h JOIN subjects s ON h.subject_id = s.id 
               WHERE h.teacher_id = ? ORDER BY h.id DESC'''

    if db_url:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute('SELECT id, dept_name FROM departments')
        depts = cur.fetchall()
        cur.execute(sql_h, (session['user_id'],))
        my_handouts = cur.fetchall()
        cur.execute("SELECT COUNT(*) FROM handouts")
        total_count = cur.fetchone()['count']
        cur.close()
    else:
        depts = conn.execute('SELECT id, dept_name FROM departments').fetchall()
        my_handouts = conn.execute(sql_h, (session['user_id'],)).fetchall()
        total_count = conn.execute("SELECT COUNT(*) FROM handouts").fetchone()[0]

    my_handouts_list = [dict(row) for row in my_handouts]
    my_count = len(my_handouts_list)
    participation = round((my_count / total_count) * 100, 1) if total_count > 0 else 0
    conn.close()
    return render_template('teacher_dashboard.html', name=session['user_name'], depts=depts, my_handouts=my_handouts_list, my_count=my_count, participation=participation)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'user_id' not in session or session['role'] != 'teacher': return redirect(url_for('login'))
    dept_id = request.form.get('dept_id')
    semester = request.form.get('semester')
    subject_id = request.form.get('subject_id')
    title = request.form.get('title')
    notes = request.form.get('notes')
    files = request.files.getlist('files[]')
    
    conn = get_db_connection()
    db_url = os.environ.get('DATABASE_URL')
    
    try:
        for file in files:
            if file and file.filename != '':
                ext = os.path.splitext(file.filename)[1]
                unique_filename = f"{int(time.time())}_{secure_filename(file.filename)}{ext}"
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], unique_filename))
                
                sql = "INSERT INTO handouts (teacher_id, subject_id, dept_id, semester, title, notes, file_path) VALUES (%s, %s, %s, %s, %s, %s, %s)" if db_url else \
                      "INSERT INTO handouts (teacher_id, subject_id, dept_id, semester, title, notes, file_path) VALUES (?, ?, ?, ?, ?, ?, ?)"
                
                if db_url:
                    cur = conn.cursor()
                    cur.execute(sql, (session['user_id'], subject_id, dept_id, semester, title, notes, unique_filename))
                    cur.close()
                else:
                    conn.execute(sql, (session['user_id'], subject_id, dept_id, semester, title, notes, unique_filename))
        conn.commit()
        flash('✅ تم الرفع!')
    except Exception as e:
        flash(f'❌ خطأ: {e}')
    finally:
        conn.close()
    return redirect(url_for('teacher_dashboard'))

@app.route('/admin')
def admin_dashboard():
    if 'user_id' not in session or session['role'] != 'admin': return redirect(url_for('login'))
    conn = get_db_connection()
    db_url = os.environ.get('DATABASE_URL')
    sql = "SELECT * FROM teachers WHERE email != 'admin@musaid.edu.ly' ORDER BY id DESC"
    if db_url:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute(sql)
        teachers = cur.fetchall()
        cur.close()
    else:
        teachers = conn.execute(sql).fetchall()
    conn.close()
    return render_template('admin/dashboard.html', teachers=teachers)

@app.route('/logout')
def logout():
    session.clear() 
    return redirect(url_for('login'))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=False)