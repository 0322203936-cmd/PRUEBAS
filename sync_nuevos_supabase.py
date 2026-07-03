import os
import pandas as pd
from supabase import create_client, Client
import datetime

url = os.environ.get("SUPABASE_URL")
key = os.environ.get("SUPABASE_KEY")

try:
    supabase: Client = create_client(url, key)
except Exception as e:
    print(f"Error conectando a supabase: {e}")
    exit(1)

# Obtener la fecha más reciente registrada en Supabase
try:
    res = supabase.table('facturas_folios').select('diario').order('diario', desc=True).limit(1).execute()
    max_date_str = None
    if res.data and len(res.data) > 0:
        max_date_str = res.data[0].get('diario')
    print(f"La fecha más reciente encontrada en Supabase es: {max_date_str}")
except Exception as e:
    print(f"Error al buscar la fecha más reciente: {e}")
    max_date_str = None

print("Leyendo archivo de Excel 'Analisis CFBC MAYO-JUNIO.xlsx'...")
try:
    df = pd.read_excel('Analisis CFBC MAYO-JUNIO.xlsx')
except Exception as e:
    print(f"Error leyendo excel: {e}")
    exit(1)

df = df.fillna("")

records_to_insert = []
for index, row in df.iterrows():
    diario_val = row.get('Diario', '')
    if isinstance(diario_val, datetime.datetime):
        diario_val = diario_val.strftime('%Y-%m-%d')
    elif str(diario_val).strip() == "":
        diario_val = None
    else:
        try:
            diario_val = pd.to_datetime(diario_val).strftime('%Y-%m-%d')
        except:
            diario_val = None

    # Filtrar: Si tenemos una fecha máxima, solo tomamos las fechas estrictamente mayores
    if max_date_str and diario_val:
        if diario_val <= max_date_str:
            continue # Saltamos las fechas que ya están subidas o son anteriores

    def clean_num(val):
        try:
            v = float(val)
            return v if pd.notna(v) else 0.0
        except:
            return 0.0

    record = {
        "diario": diario_val,
        "ruta": str(row.get('Ruta', '')),
        "viajes_por_ruta": str(row.get('Viajes Por Ruta', '')),
        "prueba": str(row.get('Prueba', '')),
        "sem": str(row.get('SEM', '')),
        "tienda": str(row.get('Nombre Tienda/Club', '')),
        "salida": str(row.get('Salida', '')),
        "folio": str(row.get('Folio contpaq', '')),
        "producto": str(row.get('Producto', '')),
        "unidades": clean_num(row.get('Unidades', 0)),
        "precio_unidad": clean_num(row.get('Precio Unidad', 0)),
        "venta_total": clean_num(row.get('Venta Total', 0))
    }
    records_to_insert.append(record)

if len(records_to_insert) == 0:
    print("No se encontraron registros nuevos posteriores a la fecha máxima. No hay nada que subir.")
    exit(0)

print(f"Preparados {len(records_to_insert)} registros NUEVOS. Subiendo a Supabase...")

BATCH_SIZE = 100
total_inserted = 0

for i in range(0, len(records_to_insert), BATCH_SIZE):
    batch = records_to_insert[i:i+BATCH_SIZE]
    try:
        response = supabase.table('facturas_folios').insert(batch).execute()
        total_inserted += len(batch)
        print(f"Progreso: {total_inserted}/{len(records_to_insert)}")
    except Exception as e:
        print(f"Error en el lote {i}-{i+BATCH_SIZE}: {e}")

print("¡Proceso de subida de datos nuevos finalizado!")
