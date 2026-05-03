import sqlite3
import psycopg2
import os

# بيانات الاتصال (تأكد من وضع رابط PostgreSQL الخاص بك هنا أو سيقرأه من البيئة)
SQLITE_DB = 'musaid_ist.db'
# الرابط الذي أرسلته أنت الآن
POSTGRES_URL = "postgresql://musaid_db_user:bhP7p8VMKEy4dUlZOB7j6x1YKzVA1GKX@dpg-d7rn6rd7vvec738q0p8g-a.virginia-postgres.render.com/musaid_db"
def migrate():
    # الاتصال بـ SQLite
    if not os.path.exists(SQLITE_DB):
        print(f"خطأ: ملف {SQLITE_DB} غير موجود في المجلد الحالي.")
        return

    sqlite_conn = sqlite3.connect(SQLITE_DB)
    sqlite_cur = sqlite_conn.cursor()

    # الاتصال بـ PostgreSQL
    try:
        pg_conn = psycopg2.connect(POSTGRES_URL)
        pg_cur = pg_conn.cursor()
        print("تم الاتصال بـ PostgreSQL بنجاح.")
    except Exception as e:
        print(f"فشل الاتصال بـ PostgreSQL: {e}")
        return

    # نقل الجداول (مثال لجدول الأقسام)
    tables = ['departments', 'teachers', 'subjects', 'handouts']
    
    for table in tables:
        print(f"جاري نقل بيانات الجدول: {table}...")
        sqlite_cur.execute(f"SELECT * FROM {table}")
        rows = sqlite_cur.fetchall()
        
        if rows:
            # مسح البيانات القديمة في بوسطجرس لتجنب التكرار
            pg_cur.execute(f"TRUNCATE TABLE {table} RESTART IDENTITY CASCADE")
            
            # تجهيز أمر الإدخال (نستخدم %s لبوسطجرس)
            placeholders = ', '.join(['%s'] * len(rows[0]))
            insert_query = f"INSERT INTO {table} VALUES ({placeholders})"
            
            pg_cur.executemany(insert_query, rows)
            print(f"تم نقل {len(rows)} سجل بنجاح.")

    pg_conn.commit()
    sqlite_conn.close()
    pg_conn.close()
    print("✅ اكتملت عملية الترحيل بنجاح!")

if __name__ == "__main__":
    migrate()