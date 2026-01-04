import os
from sqlalchemy import create_engine, inspect, text
from dotenv import load_dotenv

load_dotenv()

db_url = os.environ.get("DATABASE_URL")
if not db_url:
    print("No DATABASE_URL found in .env")
    exit(1)

print(f"Connecting to: {db_url.split('@')[1] if '@' in db_url else 'Local SQLite'}")

try:
    engine = create_engine(db_url)
    inspector = inspect(engine)
    
    tables = inspector.get_table_names()
    print(f"\nTablas encontradas: {tables}")
    
    with engine.connect() as conn:
        for table in tables:
            print(f"\n--- Tabla: {table} ---")
            # Count rows
            try:
                result = conn.execute(text(f"SELECT COUNT(*) FROM {table}"))
                count = result.scalar()
                print(f"Total registros: {count}")
                
                # Show last 3 rows if any
                if count > 0:
                    print("Últimos 3 registros:")
                    # Assuming there is an 'id' or 'fecha' column, but simple select limit is safer for generic
                    rows = conn.execute(text(f"SELECT * FROM {table} LIMIT 3")).fetchall()
                    for row in rows:
                        print(row)
            except Exception as e:
                print(f"Error querying table {table}: {e}")

except Exception as e:
    print(f"Error connecting to database: {e}")
