import os
from supabase import create_client, Client

url = os.environ.get("SUPABASE_URL")
key = os.environ.get("SUPABASE_KEY")

try:
    supabase: Client = create_client(url, key)
    print("Conectando a Supabase para borrar los datos...")
    
    # El filtro .neq() con un valor que nunca existirá asegura que se borren todas las filas
    response = supabase.table('facturas_folios').delete().neq('folio', 'BORRAR_TODO_XXXXX').execute()
    
    print(f"¡Éxito! Se borraron {len(response.data)} registros antiguos.")
except Exception as e:
    print(f"Error al borrar los datos: {e}")
