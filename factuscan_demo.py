# ---------------------------------------------------------------------------
# FactuScan — versión de DEMOSTRACIÓN pública.
#
# Generado automáticamente a partir del programa original con una herramienta
# interna.
# La lógica de extracción es EXACTAMENTE la misma (el generador lo comprueba).
# Sólo cambian dos cosas:
#   - Los nombres de columna de salida: los originales son los del importador
#     de un sistema contable y aquí se renombraron por genéricos
#     (factura_numero, nif, linea_precio...).
#   - El bloque de configuración: empresa y proveedores inventados.
# No editar a mano: los cambios se hacen en el programa original y se regenera.
# ---------------------------------------------------------------------------
from email.mime import text
import fitz
import zipfile
import pandas as pd
import re
import os
import shutil

# =====================================================================
# CONFIGURACIÓN DEL NEGOCIO  (DEMO: empresa y proveedores ficticios)
# Todos los valores fijos están aquí: cambiarlos aquí, no dentro de las
# funciones.
# =====================================================================

# Series del importador contable. Las define contabilidad:
# REVISARLAS AL EMPEZAR CADA AÑO.
SERIE_VENTAS = "DV26"
SERIE_COMPRAS_AMAZON = "DA26"
SERIE_COMPRAS_META = "DM26"

# NIF de la empresa que emite las ventas: se ignora al buscar el NIF del cliente
NIF_PROPIO = "B00000000"

# Palabras que aparecen cerca de un NIF y no lo son
PALABRAS_NO_NIF = ["PERSONA", "SUJETO", "JURIDICA", "COMERCIAL", "SOCIEDAD", "BUSINESS", "NUBEMARKET"]

# Proveedor de publicidad tipo "Amazon": cómo reconocerlo y datos fijos
AMAZON_NOMBRES = ["anuncios nube online"]
AMAZON_NIF = "B00000011"
AMAZON_DESCRIPCION = "Tarifas publicitarias"

# Proveedor de publicidad tipo "Meta": cómo reconocer el número de factura y datos fijos
META_ETIQUETA_FACTURA = "Factura número "
META_PREFIJO_FACTURA = "SOCADS-"
META_NOMBRE = "Social Ads Demo Ltd."
META_DIRECCION = "1 Demo Street Dublin 2 Ireland"
META_CP = "D02 X0X0"
META_NIF = "IE0000000X"
META_DESCRIPCION = "Publicidad"


def show_menu():
    print("\n" + "="*50)
    print("FACTUSCAN - PROCESADOR DE FACTURAS")
    print("="*50)
    print("Selecciona el tipo de factura a procesar:")
    print("1. Factura de Venta")
    print("2. Factura de Compra (Amazon)")
    print("3. Factura de Compra (Meta)")
    print("4. Salir")
    print("="*50)
    
    while True:
        choice = input("Ingresa tu opción (1, 2, 3 o 4): ").strip()
        if choice in ['1', '2', '3', '4']:
            return choice
        print("❌ Opción no válida. Por favor, selecciona 1, 2, 3 o 4.")

def select_file():
    """Función mejorada para seleccionar archivos con mejor manejo de errores"""
    try:
        # tkinter sólo hace falta para esta ventana: se importa aquí para que la
        # extracción se pueda usar en equipos sin interfaz gráfica
        import tkinter as tk
        from tkinter import filedialog

        print("\n📁 Abriendo ventana de selección de archivos...")
        root = tk.Tk()
        root.withdraw()  # Ocultar la ventana principal de tkinter
        root.lift()      # Traer al frente
        root.attributes('-topmost', True)  # Mantener en primer plano
        
        file_path = filedialog.askopenfilename(
            title="Selecciona el archivo PDF o ZIP",
            filetypes=[
                ("Archivos PDF", "*.pdf"), 
                ("Archivos ZIP", "*.zip"),
                ("Todos los archivos", "*.*")
            ]
        )
        
        root.destroy()  # Cerrar completamente la ventana de tkinter
        
        if file_path:
            print(f"✅ Archivo seleccionado: {os.path.basename(file_path)}")
            return file_path
        else:
            print("❌ No se seleccionó ningún archivo.")
            return None
            
    except Exception as e:
        print(f"❌ Error al abrir la ventana de selección: {e}")
        print("💡 Intenta ejecutar el script desde la terminal o verifica que tkinter esté instalado correctamente.")
        return None

def transform_date(date_str):
    meses = {
        # Inglés
        "january": "01", "february": "02", "march": "03", "april": "04", "may": "05", "june": "06",
        "july": "07", "august": "08", "september": "09", "october": "10", "november": "11", "december": "12",
        # Español
        "enero": "01", "febrero": "02", "marzo": "03", "abril": "04", "mayo": "05", "junio": "06",
        "julio": "07", "agosto": "08", "septiembre": "09", "octubre": "10", "noviembre": "11", "diciembre": "12",
        "ene": "01", "feb": "02", "mar": "03", "abr": "04", "may": "05", "jun": "06",
        "jul": "07", "ago": "08", "sep": "09", "oct": "10", "nov": "11", "dic": "12",
        # Italiano
        "gennaio": "01", "febbraio": "02", "marzo": "03", "aprile": "04", "maggio": "05", "giugno": "06",
        "luglio": "07", "agosto": "08", "settembre": "09", "ottobre": "10", "novembre": "11", "dicembre": "12",
        # Francés
        "janvier": "01", "février": "02", "mars": "03", "avril": "04", "mai": "05", "juin": "06",
        "juillet": "07", "août": "08", "septembre": "09", "octobre": "10", "novembre": "11", "décembre": "12",
        # Alemán
        "januar": "01", "februar": "02", "märz": "03", "april": "04", "mai": "05", "juni": "06",
        "juli": "07", "august": "08", "september": "09", "oktober": "10", "november": "11", "dezember": "12"
    }

    date_str = date_str.lower().strip()

    # Formato numérico con puntos, guiones o barras (Ej: 15.12.2025 | 15-12-25 | 15/12/2025)
    match = re.search(r"(\d{1,2})[./\-](\d{1,2})[./\-](\d{2,4})", date_str)
    if match:
        day, month, year = match.groups()
        if len(year) == 2:
            year = "20" + year
        return f"{day}/{month}/{year}"

    # Formato texto (Ej: 10 de diciembre 2025 | 10 dicembre 2025)
    # Buscamos: dia + espacio + (opcional 'de') + mes + año
    match = re.search(r"(\d{1,2})\.?\s*(?:de)?\s*([a-záéíóúäöüß]+)\.?\s*(\d{4})", date_str)
    if match:
        day, month_text, year = match.groups()
        month = meses.get(month_text, "00")
        if month != "00":
            return f"{day}/{month}/{year}"

    return "FECHA NO RECONOCIDA"

def extract_ventas(pdf_path):
    all_products = []
    try:
        doc = fitz.open(pdf_path)
        # 1. Analizamos la PAGINA 1 para Cabeceras y Dirección
        page1 = doc[0]
        words_list = page1.get_text("words")
        text_page1 = page1.get_text("text")
        lines_p1 = [l.strip() for l in text_page1.split("\n") if l.strip()]

        # --- 1. MOTOR DE COORDENADAS (BÚSQUEDA EXACTA) ---
        ancla_y = 0
        cabeceras_busqueda = [
            "Geschäftsadresse", "Dirección comercial", "Adresse commerciale",
            "Indirizzo sede legale", "Billing Address", "Rechnungsadresse"
        ]
        for cabecera in cabeceras_busqueda:
            resultados = page1.search_for(cabecera)
            if resultados:
                ancla_y = resultados[0].y1
                break

        if ancla_y == 0:
            kw_ancla_suave = ["contacto", "contact-us", "besuche", "visitez", "contactar"]
            for w in words_list:
                if any(k in w[4].lower() for k in kw_ancla_suave):
                    ancla_y = w[1]
                    break

        lineas_coord_dict = {}
        if ancla_y > 0:
            for w in words_list:
                x0, y0, x1, y1, word_text, b_no, l_no, w_no = w
                if y0 > (ancla_y + 2) and x1 < 210:
                    stoppers_dir = [
                        "fecha del pedido", "bestelldatum", "date de la commande",
                        "pedido nº", "order #", "data ordine", "informazioni sull'ordine",
                        "dettagli fattura", "rechnungsdetails", "bestellinformationen",
                        "informations de la commande", "lieferadresse", "dirección de envío",
                        "verkauft von", "sold by"
                    ]
                    if any(k in word_text.lower() for k in stoppers_dir):
                        break
                    y_idx = round(y0)
                    if y_idx not in lineas_coord_dict: lineas_coord_dict[y_idx] = []
                    lineas_coord_dict[y_idx].append(word_text)

        lineas_cliente = [" ".join(lineas_coord_dict[y]) for y in sorted(lineas_coord_dict.keys())]

        trash_labels = [
            "dirección comercial", "geschäftsadresse", "adresse commerciale", "billing address",
            "indirizzo sede legale", "indirizzo di spedizione", "dirección de envío", "lieferadresse",
            "versendungsland", "spanien", "rechnung", "seite", "page", "página", "verkauft von"
        ]
        # --- LIMPIEZA DE BASURA ---
        lineas_cliente_limpias = []
        for l in lineas_cliente:
            l_lower = l.lower()
            if any(t in l_lower for t in trash_labels): continue
            if len(l.strip()) == 2 and l.strip().isupper(): continue
            lineas_cliente_limpias.append(l)
        lineas_cliente = lineas_cliente_limpias

        # --- 2. LÓGICA DE NOMBRE Y DIRECCIÓN ---
        nombre_cliente = "No encontrado"
        direccion_cliente = ""
        if lineas_cliente:
            primera = lineas_cliente[0]
            termina_punto = primera.endswith(".")
            es_amazon_name = "Amazon Business" in primera
            tiene_sucursal = "sucursal" in primera.lower() or "branch" in primera.lower()
            segunda_es_pais = False
            if len(lineas_cliente) > 1:
                paises_comunes = ["deutschland", "germany", "allemagne", "france", "francia", "italia", "italy", "españa", "spain", "belgique", "belgium"]
                if lineas_cliente[1].lower().strip() in paises_comunes:
                    segunda_es_pais = True

            unir_segunda = (("," in primera or tiene_sucursal) and not (termina_punto and not es_amazon_name)) or segunda_es_pais
            if unir_segunda and len(lineas_cliente) > 1:
                nombre_cliente = f"{primera} {lineas_cliente[1]}"
                restante = lineas_cliente[2:]
            else:
                nombre_cliente = primera
                restante = lineas_cliente[1:]

            nombre_cliente = re.sub(r'\s+[CcPp]/.*$|\s+CALLE.*$|\s+ESPAÑA\s+[A-Z].*$', '', nombre_cliente, flags=re.IGNORECASE).strip()
            dir_final_parts = []
            for i, r_line in enumerate(restante):
                if len(r_line) < 6 and i > 0 and not re.search(r'\d{5}', r_line):
                    dir_final_parts[-1] = f"{dir_final_parts[-1]} {r_line}"
                else:
                    dir_final_parts.append(r_line)
            direccion_cliente = dir_final_parts[0] if dir_final_parts else ""

        # --- 3. EXTRACCIÓN DE FECHA (zona superior derecha de la página 1) ---
        fecha_factura = "No encontrada"
        rect_fecha = fitz.Rect(340, 70, 600, 200)
        texto_zona_fecha = page1.get_text("text", clip=rect_fecha)
        lines_zona_fecha = [l.strip() for l in texto_zona_fecha.split("\n") if l.strip()]

        encontrada_vip = False
        for line in lines_zona_fecha:
            match = re.search(r'(\d{1,2}[\s./-](?:de\s|del\s)?[a-zA-Z0-9\u00C0-\u00FF]+[\s./-]\d{2,4})', line, re.IGNORECASE)
            if match:
                res = transform_date(match.group(1))
                if res != "FECHA NO RECONOCIDA":
                    fecha_factura = res
                    encontrada_vip = True
                    break

        if not encontrada_vip:
            kw_fecha = ["fecha de la factura", "data di fatturazione", "invoice date", "rechnungsdatum"]
            for i, line in enumerate(lines_p1[:50]):
                if any(k in line.lower() for k in kw_fecha):
                    texto_busqueda = " ".join(lines_p1[i:i+2])
                    match = re.search(r'(\d{1,2}[\s./-](?:de\s|del\s)?[a-zA-Z0-9\u00C0-\u00FF]+[\s./-]\d{2,4})', texto_busqueda, re.IGNORECASE)
                    if match:
                        res = transform_date(match.group(1))
                        if res != "FECHA NO RECONOCIDA":
                            fecha_factura = res
                            break

        # --- 4. FACTURA Y PEDIDO ---
        match_f = re.search(r'ES\d[A-Z0-9]{5,}', text_page1)
        num_factura = re.sub(r'[^A-Z0-9]', '', match_f.group().upper()) if match_f else "No encontrado"
        match_p = re.search(r'(\d{3}-\d{7}-\d{7})', text_page1)
        num_pedido = match_p.group(1) if match_p else "No encontrado"

        # --- 5. NIF Y CÓDIGO POSTAL ---
        col_full = " ".join(lineas_cliente)
        nif_final = "No encontrado"
        dtocli_final = ""
        es_extranjero = False

        cp_m = re.search(r"\b\d{5}\b", col_full)
        temp_cp = cp_m.group() if cp_m else ""
        # Regex NIF Universal
        # Etiqueta sin distinguir mayúsculas; el NIF en MAYÚSCULAS y con al menos un dígito
        labels = r'(?i:NIF|CIF|TVA|VAT|USt\.?[- ]?IDNr\.?|P\.?\s*IVA)'
        sep = r'[\s\S]*?'
        # [A-Z]\d{7}[A-Z0-9]: CIF de sociedades (B12345678) y de entidades con letra de control (R0000000A)
        patterns = r'\b([A-Z]{2}(?=[A-Z0-9]*\d)[A-Z0-9]{8,15}|[XYZ]\d{7,8}[A-Z]|\d{8}[A-Z]|[A-Z]\d{7}[A-Z0-9]|\d{9,11})\b'
        regex_nif = rf'{labels}{sep}{patterns}'
        # INTENTO 1: Buscar en el bloque del cliente
        nif_m = re.search(regex_nif, col_full)
        if not nif_m:
            nif_m = re.search(r'\b([A-Z]{2}\d{9,12}|[XYZ]\d{7,8}[A-Z]|\d{8}[A-Z]|[A-Z]\d{7}[A-Z0-9]|ESW\d{7,9}[A-Z])\b', col_full)
        # INTENTO 3 (BÚSQUEDA GLOBAL): Buscar en TODA la página
        if not nif_m:
            all_nifs = re.findall(regex_nif, text_page1)
            all_nifs_2 = re.findall(r'\b([A-Z]{2}\d{9,12}|[XYZ]\d{7,8}[A-Z]|\d{8}[A-Z]|[A-Z]\d{7}[A-Z0-9]|ESW\d{7,9}[A-Z]|\d{9,11})\b', text_page1)
            all_nifs = all_nifs + all_nifs_2
            tu_nif = [f"ES{NIF_PROPIO}", NIF_PROPIO, NIF_PROPIO[1:]]
            bad_words = PALABRAS_NO_NIF
            nifs_validos = []
            for n in all_nifs:
                n_upper = n.upper()
                if any(x in n_upper for x in tu_nif): continue
                if any(bw in n_upper for bw in bad_words): continue
                if n_upper.isdigit() and len(n_upper) != 11: continue
                nifs_validos.append(n_upper)
            if nifs_validos:
                class FakeMatch:
                    def __init__(self, text): self.text = text
                    def group(self, n): return self.text
                nif_m = FakeMatch(nifs_validos[0])

        if nif_m:
            val_nif = nif_m.group(1).upper()
            if NIF_PROPIO not in val_nif and f"ES{NIF_PROPIO}" not in val_nif:
                prefijos_int = ["PT", "BE", "FR", "DE", "IT", "AT", "NL", "PL", "IE"]
                es_numerico_largo = val_nif.isdigit() and len(val_nif) > 8
                if any(val_nif.startswith(p) for p in prefijos_int) or es_numerico_largo:
                    es_extranjero = True
                    nif_final = val_nif
                elif val_nif.startswith("ES") and len(val_nif) > 10:
                    nif_final = val_nif[2:]
                else:
                    nif_final = val_nif

        if es_extranjero:
            dtocli_final = ""
        else:
            dtocli_final = temp_cp

        has_real_vat = "21%" in text_page1 or "21 %" in text_page1
        tipo_iva, tipo_op = ("ORD21", "VNAC") if (not es_extranjero or has_real_vat) else ("EXE", "VCEE")

        # --- 6. EXTRACCIÓN DE PRODUCTOS ---
        full_doc_lines = []
        for p in doc:
            full_doc_lines.extend([l.strip() for l in p.get_text("text").split("\n") if l.strip()])

        asin_indices = [i for i, line in enumerate(full_doc_lines) if "ASIN:" in line]
        header_stoppers = [
            "Descripción", "Descrizione", "Description", "Beschreibung",
            "Cant.", "Qté", "Menge", "Quant.",
            "P. Unitario", "Prix unitaire", "Einzelpreis",
            "IVA %", "TVA", "MwSt", "USt.",
            "Precio total", "Montant", "Gesamtpreis", "Prezzo Totale",
            "incluido", "excluido", "inclusa", "esclusa"
        ]
        # Orden de columnas según el importador
        base_item = {
            "factura_numero": num_factura,
            "pedido_numero": num_pedido,
            "factura_fecha": fecha_factura,
            "nombre": nombre_cliente,
            "direccion": direccion_cliente,
            "codigo_postal": dtocli_final,
            "nif": nif_final,
            "linea_descripcion": "",
            "linea_unidades": "",
            "linea_precio": 0.0,
            "cliente_codigo": "",
            "serie": SERIE_VENTAS,
            "factura_numero_interno": "",
            "linea_tipo_iva": tipo_iva,
            "linea_tipo_operacion": tipo_op,
            "pago_documento": "TR",
            "pago_forma": "C",
            "texto": "",
            "analitica_nivel1": "",
            "analitica_nivel2": "",
            "rectificada_serie": "",
            "rectificada_numero": ""
        }

        for idx in asin_indices:
            desc_lines = []
            for j in range(idx - 1, -1, -1):
                line_to_check = full_doc_lines[j].strip()
                if any(s.lower() in line_to_check.lower() for s in header_stoppers): break
                if re.search(r'\(IVA (excluido|incluido|esclusa|inclusa)\)', line_to_check, re.I): break
                # Fin del producto anterior: su línea ASIN o sus importes (regla de legacy/pruebaFACTUSCAN.py)
                if "ASIN:" in line_to_check or re.fullmatch(r'[\d.,€%\s\-]+', line_to_check): break
                desc_lines.insert(0, line_to_check)

            full_desc = " ".join(desc_lines).strip()
            full_desc = re.sub(r'^(TTC\s*Total\s*TTC|Total\s*TTC)\s*', '', full_desc, flags=re.IGNORECASE).strip()
            for s in header_stoppers:
                if s in full_desc: full_desc = full_desc.split(s)[-1].strip()

            bloque_datos = " ".join(full_doc_lines[idx+1:idx+6])
            cant_m = re.search(r"\b(\d+)\b", bloque_datos)
            prec_m = re.search(r"(\d+[.,]\d+)", bloque_datos)

            if cant_m and prec_m:
                item = base_item.copy()
                item["linea_descripcion"] = full_desc
                item["linea_unidades"] = cant_m.group(1)

                raw_price = prec_m.group(1)
                if "," in raw_price and "." in raw_price:
                    raw_price = raw_price.replace(".", "").replace(",", ".")
                elif "," in raw_price:
                    raw_price = raw_price.replace(",", ".")
                item["linea_precio"] = float(raw_price)
                all_products.append(item)

        # --- 7. GASTOS DE ENVÍO ---
        # Se busca en los bloques de la página 1 una línea de envío con importe
        shipping_keywords = ["envío", "shipping", "frais de port", "livraison", "spedizione", "porto"]
        blocks = page1.get_text("blocks")

        for b in blocks:
            block_text = b[4].replace("\n", " ")
            text_lower = block_text.lower()
            primera_linea = b[4].strip().split("\n")[0].strip().lower()

            # ¿El bloque es la FILA de envío? Tiene que empezar por la palabra clave:
            # así no se confunde con etiquetas como "Date de la livraison 20.02.2026"
            if any(primera_linea.startswith(k) for k in shipping_keywords):

                # Evitar los bloques de dirección de envío
                if "dirección" in text_lower or "address" in text_lower or "indirizzo" in text_lower:
                    continue

                # Importes de la fila: con dos decimales y seguidos de € (una fecha nunca lo está)
                precios = re.findall(r'(\d{1,3}(?:\.\d{3})+,\d{2}|\d+[.,]\d{2})\s*€', block_text)

                if precios:
                    # El primero es el precio unitario SIN IVA, como en los productos
                    # (los siguientes son el precio con IVA y el total)
                    raw_ship_price = precios[0]
                    try:
                        if ',' in raw_ship_price and '.' in raw_ship_price:
                            price_val = float(raw_ship_price.replace('.', '').replace(',', '.'))
                        elif ',' in raw_ship_price:
                            price_val = float(raw_ship_price.replace(',', '.'))
                        else:
                            price_val = float(raw_ship_price)

                        if price_val > 0:
                            item_envio = base_item.copy()
                            item_envio["linea_descripcion"] = "Envío"
                            item_envio["linea_unidades"] = "1"
                            item_envio["linea_precio"] = price_val
                            all_products.append(item_envio)
                            break
                    except:
                        pass

        doc.close()

    except Exception as e:
        print(f"❌ Error crítico en extract_ventas: {e}")

    return all_products


def extract_compras_amazon(pdf_path):
    """
    Extrae datos de facturas de Amazon (Soporte Español e Inglés).
    Corrige error de conversión de puntos sueltos y respeta orden Exemple.xls.
    """
    try:
        doc = fitz.open(pdf_path)
        text = "\n".join([page.get_text("text") for page in doc])
        lines = [line.strip() for line in text.split("\n") if line.strip()]
        doc.close()
    except Exception as e:
        print(f"❌ Error leyendo PDF {pdf_path}: {e}")
        return None

    # Palabras clave (español e inglés)
    kw_factura = ["número de factura", "invoice number", "invoice #"]
    kw_fecha = ["fecha de la factura", "invoice date"]
    kw_amazon_name = AMAZON_NOMBRES
    kw_stop_address = ["impuesto", "factura", "vat", "invoice", "nif", "cif"]

    # Orden de columnas según Exemple.xls
    data = {
        "factura_numero": "No encontrado",
        "pedido_numero": "",
        "factura_fecha": "No encontrada",
        "nombre": "No encontrado",
        "direccion": "No encontrada",
        "codigo_postal": "No encontrado",
        "nif": AMAZON_NIF,


        "linea_descripcion": AMAZON_DESCRIPCION,
        "linea_unidades": "1",
        "linea_precio": "No encontrado",

        "cliente_codigo": "",
        "serie": SERIE_COMPRAS_AMAZON,
        "factura_numero_interno": "",
        "linea_tipo_iva": "ORD21",
        "linea_tipo_operacion": "CNAC",
        "pago_documento": "TR",
        "pago_forma": "C",
        "texto": "",
        "analitica_nivel1": "",
        "analitica_nivel2": "",
        "rectificada_serie": "",
        "rectificada_numero": ""
    }

    seller_address_lines = []
    is_capturing_address = False

    # --- Búsqueda y Extracción de Datos ---
    for i, line in enumerate(lines):
        line_lower = line.lower()

        # D: Nombre del Vendedor (activa la captura de dirección)
        if any(k in line_lower for k in kw_amazon_name) and not is_capturing_address:
            data["nombre"] = line
            is_capturing_address = True
            continue

        # A: N° de Factura (línea de abajo)
        elif any(k in line_lower for k in kw_factura) and i + 1 < len(lines):
            data["factura_numero"] = lines[i + 1]
            continue

        # C: Fecha de la factura (línea de abajo)
        elif any(k in line_lower for k in kw_fecha) and i + 1 < len(lines):
            raw_date = lines[i + 1]
            # Solo si parece una fecha
            if len(raw_date) > 5:
                data["factura_fecha"] = transform_date(raw_date)
            continue

        # J: Precio (busca "Subtotal" y toma el valor de la línea siguiente)
        elif "subtotal" in line_lower and data["linea_precio"] == "No encontrado" and i + 1 < len(lines):
            linea_del_precio = lines[i + 1]

            # Regex: número con separador de miles opcional (1.024,50 | 1,024.50) o sin él (516,49)
            price_match = re.search(r'(\d{1,3}(?:[.,]\d{3})+(?:[.,]\d+)?|\d+[.,]?\d*)', linea_del_precio)

            if price_match:
                price_str = price_match.group(1)

                try:
                    # Normalización de separadores (1.234,56 | 1,234.56 | 12,34)
                    if ',' in price_str and '.' in price_str:
                        # Formato europeo (punto de miles, coma decimal)
                        if price_str.find('.') < price_str.find(','):
                            price_str = price_str.replace('.', '').replace(',', '.')
                        # Formato inglés (coma de miles, punto decimal)
                        else:
                            price_str = price_str.replace(',', '')
                    # Solo coma: es decimal
                    elif ',' in price_str:
                        # 12,34 -> 12.34
                        price_str = price_str.replace(',', '.')

                    val = float(price_str)
                    if val > 0:
                        data["linea_precio"] = val
                except ValueError:
                    # Valor no numérico: se ignora
                    continue
            continue

        # --- Captura de la DIRECCIÓN del vendedor ---
        elif is_capturing_address:
            if any(k in line_lower for k in kw_stop_address):
                is_capturing_address = False
            else:
                seller_address_lines.append(line)

    # --- Procesamiento de Datos Compuestos ---
    if seller_address_lines:
        if len(seller_address_lines) > 0:
            data["direccion"] = seller_address_lines[0]

        full_address_block = " ".join(seller_address_lines)
        cp_match = re.search(r'\b(\d{5})\b', full_address_block)
        if cp_match:
            data["codigo_postal"] = cp_match.group(1)

    if data["factura_numero"] != "No encontrado":
        return [data]
    else:
        return None

def extract_compras_meta(pdf_path):
    """Extracción de facturas de publicidad de Meta (datos fijos en la configuración)"""
    doc = fitz.open(pdf_path)

    # Plantilla de datos para Meta (columnas fijas conocidas)
    data_template = {
        "factura_numero": "No encontrado",      # A: Número factura (META_PREFIJO_FACTURA + número)
        "pedido_numero": "No encontrado",       # B: Número referencia  
        "factura_fecha": "No encontrada",     # C: Fecha
        "nombre": META_NOMBRE,           # D: Fijo
        "direccion": META_DIRECCION,        # E: Fijo
        "codigo_postal": META_CP,               # F: Fijo
        "nif": META_NIF,              # G: Fijo
        "linea_descripcion": META_DESCRIPCION, # H: Fijo
        "linea_unidades": "1",                   # I: Fijo
        "linea_precio": "No encontrado",    # J: Precio total
        "proveedor_codigo": "",
        "serie": SERIE_COMPRAS_META,  # L: Fijo Meta
        "factura_numero_interno": "",
        "linea_tipo_iva": "ORD21",            # N: Fijo
        "linea_tipo_operacion": "CCES",        # O: Cambio a ES
        "pago_documento": "TR",                  # P: Fijo
        "pago_forma": "C",                   # Q: Fijo
        "texto": "",
        "analitica_nivel1": "",
        "analitica_nivel2": "",
        "rectificada_serie": "",
        "rectificada_numero": ""
    }

    # Extraer texto de todas las páginas
    text = "\n".join([page.get_text("text") for page in doc])
    lines = text.split("\n")

    for i, line in enumerate(lines):
        line = line.strip()
        
        # A: Número de factura (línea que empieza con META_ETIQUETA_FACTURA + META_PREFIJO_FACTURA)
        if line.startswith(META_ETIQUETA_FACTURA + META_PREFIJO_FACTURA):
            data_template["factura_numero"] = line.replace(META_ETIQUETA_FACTURA, "")
        
        # B: Número de referencia (línea que dice "Número de referencia:")
        elif "Número de referencia:" in line:
            # El valor está en la misma línea, después de los dos puntos.
            # Dividimos la línea por ":" y tomamos el último elemento.
            valor = line.split(":")[-1].strip()
            data_template["pedido_numero"] = valor
        
        # C: Fecha (buscar fecha después de "Fecha de la factura/pago")
        elif "Fecha de la factura/pago" in line:
            if i + 1 < len(lines):
                raw_date = lines[i + 1].strip()
                data_template["factura_fecha"] = transform_date(raw_date)
        
        # J: Precio total (línea con EUR y paréntesis)
        elif "€ (EUR)" in line:
            price_match = re.search(r"([\d.,]+)\s*€\s*\(EUR\)", line)
            if price_match:
                price_str = price_match.group(1).replace(',', '.')
                data_template["linea_precio"] = float(price_str)

    doc.close()
    return [data_template] if data_template["factura_numero"] != "No encontrado" else None

def es_rectificativa(pdf_path):
    """
    Detecta facturas rectificativas / notas de crédito de Amazon.
    Devuelve {"numero", "original"} si lo es, o None si es una factura normal.
    Por ahora no se importan: se listan aparte para cargarlas a mano.
    """
    doc = fitz.open(pdf_path)
    text = doc[0].get_text("text")
    doc.close()

    # El número propio termina en ...ALC (las normales en ...ALI) y el texto
    # dice "nota de crédito" aunque el título sea "Factura rectificativa"
    numeros_alc = re.findall(r'ES\d[A-Z0-9]{5,}ALC', text)
    if not numeros_alc and "nota de crédito" not in text.lower():
        return None

    # La factura original va tras "nota de crédito para factura", en la misma línea o en la siguiente
    original = ""
    match_o = re.search(r'nota de crédito para factura\s+(\S+)', text, re.IGNORECASE)
    if match_o:
        original = match_o.group(1)

    return {"numero": numeros_alc[0] if numeros_alc else "", "original": original}

# Archivo y hoja de salida de cada opción del menú
ARCHIVOS_SALIDA = {
    '1': ("Facturas Ventas.xlsx", "VENTAS"),
    '2': ("Facturas Compra Amazon.xlsx", "COMPRAS"),
    '3': ("Facturas Compra Meta.xlsx", "COMPRAS"),
}

def facturas_ya_guardadas(choice):
    """Números de factura (factura_numero) que ya están en el Excel de salida de esta opción."""
    output_file, sheet_name = ARCHIVOS_SALIDA[choice]
    if not os.path.exists(output_file):
        return set()
    try:
        df_existing = pd.read_excel(output_file, sheet_name=sheet_name, dtype={'factura_numero': str})
    except Exception as e:
        print(f"⚠️  No se pudo leer '{output_file}' para buscar duplicados: {e}")
        return set()
    if 'factura_numero' not in df_existing.columns:
        return set()
    return set(df_existing['factura_numero'].dropna().astype(str).str.strip())

def process_files(choice, pdf_paths):
    """Función para procesar los archivos PDF según la opción elegida"""
    all_data = []
    included_files = []
    excluded_files = []
    rectificativas = []
    duplicadas = []

    # Una factura es duplicada si su número ya está en el Excel o ya salió en esta tanda.
    # Se compara la factura entera (no pedido + producto): un pedido puede tener varias
    # facturas, y dos pedidos distintos con los mismos productos tienen números distintos.
    ya_guardadas = facturas_ya_guardadas(choice)
    vistas_en_tanda = set()

    for pdf_path in pdf_paths:
        print(f"📄 Procesando: {os.path.basename(pdf_path)}")
        data = None

        try:
            if choice == '1':
                rect = es_rectificativa(pdf_path)
                if rect:
                    rectificativas.append((os.path.basename(pdf_path), rect["numero"], rect["original"]))
                    print(f"⚠️  {os.path.basename(pdf_path)} es una factura rectificativa: no se importa")
                    continue
                data = extract_ventas(pdf_path)
            elif choice == '2':
                data = extract_compras_amazon(pdf_path)
            elif choice == '3':
                data = extract_compras_meta(pdf_path)

            if data:
                numero = str(data[0].get("factura_numero", "")).strip()
                if numero and numero != "No encontrado":
                    if numero in ya_guardadas or numero in vistas_en_tanda:
                        motivo = "ya está en el Excel" if numero in ya_guardadas else "repetida en esta tanda"
                        duplicadas.append((os.path.basename(pdf_path), numero, motivo))
                        print(f"⚠️  {os.path.basename(pdf_path)}: la factura {numero} {motivo}, no se vuelve a agregar")
                        continue
                    vistas_en_tanda.add(numero)
                all_data.extend(data)
                included_files.append(pdf_path)
                print(f"✅ {os.path.basename(pdf_path)} procesado correctamente")
            else:
                excluded_files.append(pdf_path)
                print(f"❌ {os.path.basename(pdf_path)} no pudo ser procesado")

        except Exception as e:
            print(f"❌ Error procesando {os.path.basename(pdf_path)}: {e}")
            excluded_files.append(pdf_path)

    return all_data, included_files, excluded_files, rectificativas, duplicadas

def save_to_excel(all_data, choice):
    """
    Guarda los datos en un Excel. Lógica reestructurada para manejar correctamente
    la creación y adición de datos para cada tipo de factura.
    """
    if not all_data:
        print("ℹ️ No hay nuevos datos para guardar.")
        return False

    # 1. Creamos el DataFrame con el nombre 'df_new'
    df_new = pd.DataFrame(all_data)

    # --- CORRECCIÓN AQUÍ ---
    # Usamos 'df_new' en lugar de 'df' para limpiar los datos
    df_new = df_new.fillna("")
    df_new = df_new.replace("nan", "")
    
    try:
        # --- Lógica para FACTURAS DE VENTA (Opción 1) ---
        if choice == '1':
            output_file, sheet_name = ARCHIVOS_SALIDA['1']

            if os.path.exists(output_file):
                # Leemos el existente asegurando que codigo_postal sea texto
                df_existing = pd.read_excel(output_file, sheet_name=sheet_name, dtype={'codigo_postal': str})
                # Concatenamos con los nuevos datos LIMPIOS (df_new)
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            else:
                df_combined = df_new

            # Aseguramos formato texto en CP
            if 'codigo_postal' in df_combined.columns:
                df_combined['codigo_postal'] = df_combined['codigo_postal'].astype(str)
                # Una última limpieza por si al leer el excel viejo vinieron 'nan' literales
                df_combined['codigo_postal'] = df_combined['codigo_postal'].replace("nan", "")
            
            df_combined.to_excel(output_file, sheet_name=sheet_name, index=False)

        # --- Lógica para FACTURAS DE COMPRA (Opciones 2 y 3) ---
        elif choice in ['2', '3']:
            output_file = ARCHIVOS_SALIDA[choice][0]

            if os.path.exists(output_file):
                df_existing = pd.read_excel(output_file, sheet_name="COMPRAS")
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            else:
                df_combined = df_new
            
            if 'codigo_postal' in df_combined.columns:
                df_combined['codigo_postal'] = df_combined['codigo_postal'].astype(str)
                df_combined['codigo_postal'] = df_combined['codigo_postal'].replace("nan", "")

            df_ventas_vacia = pd.DataFrame()
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df_combined.to_excel(writer, sheet_name='COMPRAS', index=False)
                df_ventas_vacia.to_excel(writer, sheet_name='VENTAS')
        
        else:
            return False

        print(f"✅ ¡Proceso completado! Archivo '{output_file}' actualizado.")
        return True
        
    except Exception as e:
        print(f"❌ Error al escribir en el archivo Excel: {e}")
        return False

def main():
    """Función principal para el procesamiento de facturas."""
    print("🚀 Iniciando FactuScan...")
    
    while True:
        choice = show_menu()
        
        # Opción para salir
        if choice == '4':
            print("👋 ¡Hasta luego!")
            break
        
        # Seleccionar archivo
        file_path = select_file()
        if not file_path:
            print("❌ No se seleccionó archivo. Volviendo al menú principal.")
            continue

        # Verificar que el archivo existe
        if not os.path.exists(file_path):
            print(f"❌ El archivo {file_path} no existe.")
            continue

        all_data = []
        included_files = []
        excluded_files = []
        temp_dir = ""

        try:
            # Manejar archivos ZIP o PDF
            if file_path.endswith('.zip'):
                print("📦 Procesando archivo ZIP...")
                temp_dir = "temp_extracted"
                os.makedirs(temp_dir, exist_ok=True)
                
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    pdf_files = [f for f in zip_ref.namelist() if f.endswith('.pdf')]
                    
                    if not pdf_files:
                        print("❌ No se encontraron archivos PDF en el ZIP.")
                        continue
                    
                    print(f"📁 Encontrados {len(pdf_files)} archivos PDF en el ZIP")
                    for pdf_file in pdf_files:
                        zip_ref.extract(pdf_file, temp_dir)
                    
                    pdf_paths_to_process = [os.path.join(temp_dir, f) for f in pdf_files]
            else:
                pdf_paths_to_process = [file_path]

            # Procesar archivos
            all_data, included_files, excluded_files, rectificativas, duplicadas = process_files(choice, pdf_paths_to_process)

            # Guardar en Excel
            if all_data:
                save_to_excel(all_data, choice)

            # Resumen de resultados
            print(f"\n📊 RESUMEN:")
            print(f"✅ Archivos procesados: {len(included_files)}")
            print(f"❌ Archivos excluidos: {len(excluded_files)}")
            
            if excluded_files:
                print(f"🗂️  Archivos excluidos: {[os.path.basename(f) for f in excluded_files]}")

            if rectificativas:
                print(f"\n⚠️  Rectificativas NO importadas (cargar a mano): {len(rectificativas)}")
                for archivo, numero, original in rectificativas:
                    print(f"   - {archivo}: rectificativa {numero or '?'} de la factura {original or '?'}")

            if duplicadas:
                print(f"\n🔁 Facturas duplicadas NO agregadas: {len(duplicadas)}")
                for archivo, numero, motivo in duplicadas:
                    print(f"   - {archivo}: factura {numero} ({motivo})")

        except Exception as e:
            print(f"❌ Error general: {e}")
            
        finally:
            # Limpiar archivos temporales
            if temp_dir and os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
                print("🧹 Archivos temporales limpiados")
        
        # Preguntar si quiere continuar
        print("\n" + "="*50)
        while True:
            continue_choice = input("¿Quieres procesar otro archivo? (s/n): ").lower().strip()
            if continue_choice in ['s', 'si', 'sí', 'y', 'yes']:
                break
            elif continue_choice in ['n', 'no']:
                print("👋 ¡Hasta luego!")
                return
            else:
                print("❌ Por favor, responde 's' para sí o 'n' para no.")

if __name__ == "__main__":
    main()