import sqlite3
import psycopg2
from psycopg2 import extras

# 1. الرابط الذي نسخته من Render (External Database URL)
POSTGRES_URL = "postgresql://musaid_db_user:bhP7p8VMKEy4dUlZOB7j6x1YKzVA1GKX@dpg-d7rn6rd7vvec738q0p8g-a.virginia-postgres.render.com/musaid_db"

# 2. اسم ملف قاعدة البيانات المحلي في جهازك
SQLITE_DB = "musaid_ist.db"

def migrate():
    sqlite_conn = None
    pg_conn = None
    try:
        print("🔄 بدء عملية الاتصال بقواعد البيانات...")
        # الاتصال بـ SQLite
        sqlite_conn = sqlite3.connect(SQLITE_DB)
        sqlite_cursor = sqlite_conn.cursor()

        # الاتصال بـ PostgreSQL
        pg_conn = psycopg2.connect(POSTGRES_URL)
        pg_cursor = pg_conn.cursor()

        # جلب أسماء الجداول من SQLite (باستثناء جداول النظام)
        sqlite_cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%';")
        tables = sqlite_cursor.fetchall()

        if not tables:
            print("⚠️ لم يتم العثور على جداول في ملف SQLite المحلي.")
            return

        for table_row in tables:
            table_name = table_row[0]
            print(f"📦 جاري نقل الجدول: {table_name}...")

            # 1. جلب البيانات من SQLite
            sqlite_cursor.execute(f"SELECT * FROM {table_name}")
            rows = sqlite_cursor.fetchall()
            
            # جلب أسماء الأعمدة
            col_names = [description[0] for description in sqlite_cursor.description]
            
            # 2. إنشاء الجدول في PostgreSQL إذا لم يكن موجوداً (بناءً على أول سجل)
            # ملاحظة: سنقوم بتفريغ الجدول في Postgres أولاً لتجنب تكرار البيانات
            pg_cursor.execute(f"DROP TABLE IF EXISTS {table_name} CASCADE;")
            
            # بناء استعلام الإنشاء (تحويل بسيط للأنواع)
            # ملاحظة: هذا سكريبت نقل سريع، يفترض أن هيكل الجداول بسيط
            cols_with_types = []
            for col in col_names:
                cols_with_types.append(f'"{col}" TEXT') # نستخدم TEXT للتبسيط لضمان نجاح النقل
            
            create_query = f"CREATE TABLE {table_name} ({', '.join(cols_with_types)});"
            pg_cursor.execute(create_query)

            # 3. إدخال البيانات في PostgreSQL
            if rows:
                placeholders = ", ".join(["%s"] * len(col_names))
                columns_joined = ", ".join([f'"{c}"' for c in col_names])
                insert_query = f"INSERT INTO {table_name} ({columns_joined}) VALUES ({placeholders})"
                pg_cursor.executemany(insert_query, rows)
                print(f"✅ تم نقل {len(rows)} سجل في جدول {table_name}")

        pg_conn.commit()
        print("\n✨ مبروك دكتور محمد! تم نقل كل البيانات بنجاح إلى Render.")

    except Exception as e:
        print(f"\n❌ حدث خطأ أثناء النقل: {e}")
        if pg_conn:
            pg_conn.rollback()
    finally:
        if sqlite_conn: sqlite_conn.close()
        if pg_conn: pg_conn.close()

if __name__ == "__main__":
    migrate()