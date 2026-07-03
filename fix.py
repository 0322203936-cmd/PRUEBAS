import re

with open("app.py", "r", encoding="utf-8") as f:
    content = f.read()

# We need to find the start of api_analizar_factura
match_start = re.search(r"@app\.route\(\"/api/analizar_factura\", methods=\[\"POST\"\]\)\ndef api_analizar_factura\(\):", content)
start_idx = match_start.start()

# We need to find the end of the function (before @app.route('/api/analizar_recibo')
match_end = re.search(r"@app\.route\(\'/api/analizar_recibo\', methods=\[\'POST\'\]\)", content)
end_idx = match_end.start()

new_func = """@app.route("/api/analizar_factura", methods=["POST"])
def api_analizar_factura():
    if 'imagen' not in request.files:
        return jsonify({"ok": False, "error": "No image provided"}), 400
        
    file = request.files['imagen']
    if file.filename == '':
        return jsonify({"ok": False, "error": "No selected file"}), 400
        
    try:
        raw_bytes = file.read()
        file_bytes = np.frombuffer(raw_bytes, np.uint8)
        
        prompt = '''Analiza esta imagen que contiene una factura o ticket.
Extrae la siguiente información y responde ÚNICAMENTE con un JSON válido con esta estructura:
{
  "folio": "número de folio (solo números/letras, ej: 810)",
  "fecha": "fecha del documento (ej: 30/Jun/2026)",
  "total": "el monto total a pagar (solo números y decimales, ej: 3166.00)",
  "productos": [
    {
      "cantidad": "cantidad del producto",
      "descripcion": "descripción o nombre del producto"
    }
  ],
  "texto_crudo": "una transcripción del texto más importante de la factura, para poder compararlo"
}
IMPORTANTE: La imagen puede tener texto en el reverso transparente. Concéntrate en la factura principal (Emisor, Receptor, Productos, Total). NO incluyas markdown, SOLO EL JSON.'''

        model = genai.GenerativeModel("gemini-1.5-flash")
        imagen_part = {"mime_type": file.mimetype or "image/jpeg", "data": raw_bytes}
        response = model.generate_content([prompt, imagen_part])
        response_text = response.text.strip()
        
        if response_text.startswith("```json"): response_text = response_text[7:]
        if response_text.startswith("```"): response_text = response_text[3:]
        if response_text.endswith("```"): response_text = response_text[:-3]
            
        import json
        datos = json.loads(response_text.strip())
        
        folio_encontrado = datos.get("folio", "No detectado")
        fecha_detectada = datos.get("fecha", "No detectada")
        total_detectado = datos.get("total", "No detectado")
        productos_gemini = datos.get("productos", [])
        full_text = datos.get("texto_crudo", "")
        
        print(f"Extracción Gemini exitosa. Folio: {folio_encontrado}", flush=True)
        
        folio_manual = request.form.get('folio_manual', '').strip()
        if folio_manual:
            folio_encontrado = folio_manual
            print(f"Folio sobrescrito manualmente: {folio_encontrado}", flush=True)
            
        db_productos = []
        db_status = "NOT_FOUND"
        
        if folio_encontrado != "No detectado":
            db_res = supabase_client.table("facturas_folios").select("*").eq("folio", folio_encontrado).execute()
            db_productos = db_res.data
            if db_productos:
                db_status = "FOUND"
                
        serie = "Desconocida"
        
        if not folio_encontrado or folio_encontrado == "No detectado":
            return jsonify({
                "ok": True,
                "factura": {"serie": serie, "folio": folio_encontrado, "fecha": fecha_detectada, "total": total_detectado, "url_factura": ""},
                "db_status": "NOT_FOUND",
                "comparacion": [],
                "ocr_raw_text": full_text
            })
            
        comparacion = []
        if db_status == "FOUND":
            import unicodedata
            def remove_accents(input_str):
                return unicodedata.normalize('NFKD', input_str).encode('ASCII', 'ignore').decode('utf-8')
                
            ocr_clean = remove_accents(full_text.lower())
            
            for db_p in db_productos:
                prod_name = str(db_p.get("producto", "")).strip()
                cant_db = float(db_p.get("unidades", 0))
                precio_db = float(db_p.get("precio_unidad", 0))
                
                prod_clean = remove_accents(prod_name.lower())
                words = prod_clean.split()
                matched_words = sum(1 for w in words if w in ocr_clean)
                
                is_match = False
                if len(words) > 0 and (matched_words / len(words)) >= 0.5:
                    is_match = True
                    
                comparacion.append({
                    "producto_db": prod_name,
                    "cantidad_db": cant_db,
                    "precio_db": precio_db,
                    "estado": "OK" if is_match else "DIFF"
                })
                    
        import base64, time
        url_factura_temp = ""
        try:
            b64_str = base64.b64encode(file_bytes).decode('utf-8')
            ruta_supa = f"Facturas/factura_{int(time.time())}.jpg"
            url_factura_temp = subir_foto_supabase(b64_str, ruta_supa)
        except Exception as ex:
            print(f"Error subiendo foto factura: {ex}")
            
        return jsonify({
              "ok": True,
              "factura": {
                  "serie": serie,
                  "folio": folio_encontrado,
                  "fecha": fecha_detectada,
                  "total": total_detectado,
                  "url_factura": url_factura_temp
              },
              "db_status": db_status,
              "comparacion": comparacion,
              "productos_gemini": productos_gemini,
              "ocr_raw_text": full_text
          })
        
    except Exception as e:
        print(f"Error en analisis Gemini: {e}")
        return jsonify({"ok": False, "error": str(e)}), 500

"""

new_content = content[:start_idx] + new_func + content[end_idx:]

with open("app.py", "w", encoding="utf-8") as f:
    f.write(new_content)
    
print("Archivo actualizado")
