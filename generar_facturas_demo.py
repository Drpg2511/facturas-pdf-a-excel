"""
Genera las facturas PDF sintéticas de la demo y su ground_truth.xlsx.

    python demo-publica/generar_facturas_demo.py

Todo es inventado: la empresa (Lumora Hogar S.L.), sus clientes, el mercado
online (Nubemarket) y los proveedores de publicidad (Anuncios Nube Online,
Social Ads Demo). Cada página lleva la marca "DOCUMENTO FICTICIO - DEMO".
Las facturas imitan la ESTRUCTURA de las reales (etiquetas, posiciones, orden
del texto) para que la lógica de extracción funcione igual, sin copiar datos.

Salida:
  facturas/ventas/*.pdf          12 ventas + 3 rectificativas
  facturas/compra-amazon/*.pdf   6 compras tipo Amazon
  facturas/compra-meta/*.pdf     4 compras tipo Meta
  ground_truth.xlsx              valores exactos usados y valores esperados
"""
import os
import shutil
from datetime import date

import fitz  # PyMuPDF
import pandas as pd

AQUI = os.path.dirname(os.path.abspath(__file__))
SALIDA = os.path.join(AQUI, "facturas")
FUENTE = fitz.Font("helv")
FUENTE_NEGRITA = fitz.Font("hebo")
A4 = (595.28, 841.88)

EMPRESA = {"nombre": "LUMORA HOGAR S.L.", "nif": "B00000000", "calle": "Calle del Ejemplo 12",
           "ciudad": "Villademo, Madrid, 28999"}
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]
MESES_FR = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août",
            "septembre", "octobre", "novembre", "décembre"]
MESES_META = ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sept.", "oct.", "nov.", "dic."]


# ----------------------------------------------------------------------------- utilidades
class Pagina:
    """Escribe texto en una página. y = borde superior de la línea, como en las facturas reales."""

    def __init__(self, doc):
        self.page = doc.new_page(width=A4[0], height=A4[1])
        self.tw = fitz.TextWriter(self.page.rect)

    def t(self, x, y, texto, size=7.5, negrita=False):
        self.tw.append((x, y + size * 0.85), texto, font=FUENTE_NEGRITA if negrita else FUENTE, fontsize=size)

    def cerrar(self):
        self.tw.write_text(self.page)
        # Marca de agua: en gris claro y fuera de las zonas que lee el programa
        marca = fitz.TextWriter(self.page.rect, color=(0.75, 0.75, 0.75))
        marca.append((150, 820), "DOCUMENTO FICTICIO - DEMO - NO VÁLIDO COMO FACTURA", font=FUENTE_NEGRITA, fontsize=11)
        marca.write_text(self.page)


def partir(texto, ancho, size=7.5):
    """Parte un texto en líneas que quepan en 'ancho' puntos, como hacen las facturas reales."""
    lineas, actual = [], ""
    for palabra in texto.split():
        prueba = f"{actual} {palabra}".strip()
        if actual and FUENTE.text_length(prueba, fontsize=size) > ancho:
            lineas.append(actual); actual = palabra
        else:
            actual = prueba
    return lineas + [actual]


def euros(valor):
    """1234.5 -> '1.234,50 €' (formato de las facturas)."""
    s = f"{abs(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("-" if valor < 0 else "") + s + " €"


def fecha_texto(d, idioma="es"):
    meses = MESES_FR if idioma == "fr" else MESES
    return f"{d.day:02d} {meses[d.month - 1]} {d.year}"


def fecha_puntos(d):
    return f"{d.day:02d}.{d.month:02d}.{d.year}"


def esperado_fecha(d):
    return f"{d.day:02d}/{d.month:02d}/{d.year}"


# ----------------------------------------------------------------------------- datos
# Cada venta: cliente, productos (líneas de descripción, ASIN, cantidad, precio unitario SIN IVA)
VENTAS = [
    dict(archivo="DV-001.pdf", numero="ES9000A1MC1ALI", pedido="901-0000001-0000101", fecha=date(2026, 2, 10),
         formato_fecha="texto", cliente=["Lucía Ferrán Ortega", "Calle de los Almendros 14", "Valencia, Valencia, 46021"],
         nif_etiqueta="NIF persona jurídica", nif_impreso="99000001R", nif_esperado="99000001R", pais="ES", iva=21,
         productos=[(["Lámpara de Pie Regulable con Tres Brazos - Luz Cálida", "para Salón y Dormitorio - Blanco"], "B0DEMO0001", 1, 41.32)]),
    dict(archivo="DV-002.pdf", numero="ES9000A2MC1ALI", pedido="902-0000002-0000202", fecha=date(2026, 3, 5),
         formato_fecha="puntos", cliente=["Talleres Brisamar S.L.", "Avenida del Puerto 221", "Castellón, Castellón, 12100"],
         nif_etiqueta="NIF sujeto de IVA", nif_impreso="ESB99000002", nif_esperado="B99000002", pais="ES", iva=21,
         productos=[(["Estantería Metálica 5 Baldas 180 x 90 cm - Carga 150 kg", "por Balda - Gris"], "B0DEMO0002", 2, 57.85)]),
    dict(archivo="DV-003.pdf", numero="ES9000A3MC1ALI", pedido="903-0000003-0000303", fecha=date(2026, 2, 12),
         formato_fecha="texto", cliente=["Marc Soler Vidal", "Carrer de la Riera 8, 2-1", "Barcelona, Barcelona, 08003"],
         nif_etiqueta="NIF persona jurídica", nif_impreso="99000003W", nif_esperado="99000003W", pais="ES", iva=21,
         productos=[(["Organizador de Cajones Bambú 6 Compartimentos - Cocina"], "B0DEMO0003", 1, 16.49)]),
    dict(archivo="DV-004.pdf", numero="ES9000A4MC1ALI", pedido="904-0000004-0000404", fecha=date(2026, 2, 16),
         formato_fecha="texto", cliente=["Hostal Las Dunas S.L.", "Paseo Marítimo 3", "Almería, Almería, 04007"],
         nif_etiqueta="NIF sujeto de IVA", nif_impreso="ESB99000004", nif_esperado="B99000004", pais="ES", iva=21,
         productos=[(["Juego de Toallas de Algodón 4 Piezas - Baño y Lavabo", "Suaves y Absorbentes - Arena"], "B0DEMO0041", 3, 20.66),
                    (["Perchero de Pie de Madera con 8 Ganchos - Entrada y", "Dormitorio - Natural"], "B0DEMO0042", 1, 33.05),
                    (["Cesto para Ropa Plegable 60 L con Tapa - Lino Gris"], "B0DEMO0043", 2, 12.39)]),
    dict(archivo="DV-005.pdf", numero="ES9000A5MC1ALI", pedido="905-0000005-0000505", fecha=date(2026, 2, 18),
         formato_fecha="puntos", cliente=["Casa Atlântica Lda", "Rua das Flores 27", "Porto, Porto, 4050-262"],
         nif_etiqueta="NIF sujeto de IVA", nif_impreso="PT509999990", nif_esperado="PT509999990", pais="PT", iva=0,
         productos=[(["Lámpara de Pie Regulable con Tres Brazos - Luz Cálida", "para Salón y Dormitorio - Negro"], "B0DEMO0005", 4, 41.32)]),
    dict(archivo="DV-006.pdf", numero="ES9000A6MC1ALI", pedido="906-0000006-0000606", fecha=date(2026, 2, 20),
         formato_fecha="puntos", idioma="fr", cliente=["Maison Lavande SARL", "12 Rue des Tilleuls", "Lyon, Rhône, 69003"],
         nif_etiqueta="TVA", nif_impreso="FR00999999990", nif_esperado="FR00999999990", pais="FR", iva=0,
         productos=[(["Étagère Murale Flottante en Chêne 80 cm - Lot de 2"], "B0DEMO0006", 1, 28.10)]),
    dict(archivo="DV-007.pdf", numero="ES9000A7MC1ALI", pedido="907-0000007-0000707", fecha=date(2026, 2, 23),
         formato_fecha="texto", cliente=["Andrés Peña Molina", "Calle Mayor 45, 3B", "Zaragoza, Zaragoza, 50001"],
         nif_etiqueta="NIF persona jurídica", nif_impreso="99000007C", nif_esperado="99000007C", pais="ES", iva=21,
         productos=[(["Silla de Escritorio Ergonómica con Reposabrazos", "Ajustables - Malla Transpirable - Negro"], "B0DEMO0007", 1, 74.38)],
         envio=5.99),
    dict(archivo="DV-008.pdf", numero="ES9000A8MC1ALI", pedido="908-0000008-0000808", fecha=date(2026, 2, 25),
         formato_fecha="texto", cliente=["Yuliya Kovalenko", "Calle Olmo 2, 1A", "Málaga, Málaga, 29004"],
         nif_etiqueta="NIF persona jurídica", nif_impreso="X9900008L", nif_esperado="X9900008L", pais="ES", iva=21,
         productos=[(["Organizador de Cajones Bambú 6 Compartimentos - Cocina"], "B0DEMO0003", 2, 16.49)]),
    dict(archivo="DV-009.pdf", numero="ES9000A9MC1ALI", pedido="909-0000009-0000909", fecha=date(2026, 3, 2),
         formato_fecha="texto", cliente=["Parroquia de San Demo", "Plaza de la Iglesia 1", "León, León, 24003"],
         nif_etiqueta="NIF sujeto de IVA", nif_impreso="R9900009E", nif_esperado="R9900009E", pais="ES", iva=21,
         productos=[(["Estantería Metálica 5 Baldas 180 x 90 cm - Carga 150 kg", "por Balda - Gris"], "B0DEMO0002", 1, 57.85)]),
    dict(archivo="DV-010.pdf", numero="ES9000B0MC1ALI", pedido="910-0000010-0001010", fecha=date(2026, 3, 4),
         formato_fecha="texto", cliente=["Nubemarket Business EU S.L., Sucursal en España", "Calle de la Nube 7", "Madrid, Madrid, 28045"],
         nif_etiqueta="NIF sujeto de IVA", nif_impreso="ESW9900010J", nif_esperado="W9900010J", pais="ES", iva=21,
         productos=[(["Perchero de Pie de Madera con 8 Ganchos - Entrada y", "Dormitorio - Natural"], "B0DEMO0042", 5, 33.05)]),
    dict(archivo="DV-011.pdf", numero="ES9000B1MC1ALI", pedido="911-0000011-0001111", fecha=date(2026, 3, 9),
         formato_fecha="puntos", cliente=["Clínica Dental Alba S.L.P.", "Calle Recoletos 30", "Burgos, Burgos, 09003"],
         nif_etiqueta="NIF sujeto de IVA", nif_impreso="ESB99000011", nif_esperado="B99000011", pais="ES", iva=21,
         productos=[(["Lámpara de Pie Regulable con Tres Brazos - Luz Cálida", "para Salón y Dormitorio - Blanco"], "B0DEMO0001", 2, 41.32)]),
    dict(archivo="DV-012.pdf", numero="ES9000B2MC1ALI", pedido="912-0000012-0001212", fecha=date(2026, 3, 11),
         formato_fecha="texto", cliente=["Irene Castaño Ruiz", "Avenida de la Constitución 9", "Sevilla, Sevilla, 41001"],
         nif_etiqueta="NIF persona jurídica", nif_impreso="99000012M", nif_esperado="99000012M", pais="ES", iva=21,
         productos=[(["Cesto para Ropa Plegable 60 L con Tapa - Lino Gris"], "B0DEMO0043", 1, 12.39)]),
]

# Rectificativas: tres variantes de etiquetas, como las reales
RECTIFICATIVAS = [
    dict(archivo="RDV-001.pdf", variante="nota_credito", numero="ES90001AMC1ALC", original="ES9000A3MC1ALI",
         fecha=date(2026, 3, 10), base=VENTAS[2]),
    dict(archivo="RDV-002.pdf", variante="rectificativa", numero="ES90001BMC1ALC", original="ES9000A7MC1ALI",
         fecha=date(2026, 3, 12), base=VENTAS[6]),
    dict(archivo="RDV-003.pdf", variante="partida", numero="ES90001CMC1ALC", original="ES9000B2MC1ALI",
         fecha=date(2026, 3, 16), base=VENTAS[11]),
]

COMPRAS_AMAZON = [
    dict(archivo="ANO-001.pdf", numero="500101DEMO26", fecha=date(2026, 1, 1), subtotal=486.20),
    dict(archivo="ANO-002.pdf", numero="500102DEMO26", fecha=date(2026, 1, 8), subtotal=512.75),
    dict(archivo="ANO-003.pdf", numero="500103DEMO26", fecha=date(2026, 1, 15), subtotal=1024.50),
    dict(archivo="ANO-004.pdf", numero="500104DEMO26", fecha=date(2026, 1, 22), subtotal=73.10),
    dict(archivo="ANO-005.pdf", numero="500105DEMO26", fecha=date(2026, 2, 1), subtotal=530.00),
    dict(archivo="ANO-006.pdf", numero="500106DEMO26", fecha=date(2026, 2, 8), subtotal=498.64),
]
PROVEEDOR_AMAZON = {"nombre": "Anuncios Nube Online, S.L.U.", "calle": "Calle de la Prueba 5",
                    "ciudad": "Madrid ES 28999", "nif": "ESB00000011", "cp": "28999"}

COMPRAS_META = [
    dict(archivo="SAD-001.pdf", numero="SOCADS-100-000000001", referencia="DEMOREF01", fecha=date(2026, 1, 31), hora="15:00",
         campanas=[("Primavera hogar", 212.40), ("Remarketing", 133.27)]),
    dict(archivo="SAD-002.pdf", numero="SOCADS-100-000000002", referencia="DEMOREF02", fecha=date(2026, 2, 28), hora="09:30",
         campanas=[("Lámparas", 640.00), ("Estanterías", 512.35), ("Marca", 806.65)]),
    dict(archivo="SAD-003.pdf", numero="SOCADS-100-000000003", referencia="DEMOREF03", fecha=date(2026, 3, 31), hora="23:10",
         campanas=[("Baño y textil", 98.90)]),
    dict(archivo="SAD-004.pdf", numero="SOCADS-100-000000004", referencia="DEMOREF04", fecha=date(2026, 8, 15), hora="12:00",
         campanas=[("Vuelta al cole", 301.15), ("Ofertas agosto", 145.80)]),
]
PROVEEDOR_META = {"nombre": "Social Ads Demo Ltd.", "ciudad": "Dublin 2", "nif": "IE0000000X"}


# ----------------------------------------------------------------------------- ventas
ETIQUETAS = {
    "es": dict(titulo="Factura", pagado="Pagado", ref_pago="Nº de referencia de pago", vendido="Vendido por",
               iva="IVA", fecha="Fecha de la factura/Fecha de la entrega", numero="Número de la factura",
               total_pend="Total pendiente", contacto="Si tienes preguntas sobre tus pedidos, visita",
               dir_com="Dirección comercial", dir_env="Dirección de envío", info="Información del pedido",
               f_pedido="Fecha del pedido", n_pedido="Número del pedido", pedido_por="Pedido por",
               detalles="Detalles de la factura", envio="Envío", pagina="Página",
               cab=[(38.1, 0, "Descripción"), (307.5, 0, "Cant."), (357.6, 0, "P. Unitario"), (346.0, 1, "(IVA excluido)"),
                    (407.0, 0, "IVA %"), (457.5, 0, "P. Unitario"), (447.9, 1, "(IVA incluido)"),
                    (517.6, 0, "Precio total"), (512.6, 1, "(IVA incluido)")]),
    "fr": dict(titulo="Facture", pagado="Payé", ref_pago="Nº de référence de paiement", vendido="Vendu par",
               iva="TVA", fecha="Date de la facture/Date de la livraison", numero="Numéro de la facture",
               total_pend="Total à payer", contacto="Pour toute question sur vos commandes, visitez",
               dir_com="Adresse commerciale", dir_env="Adresse de livraison", info="Informations de la commande",
               f_pedido="Date de la commande", n_pedido="Numéro de la commande", pedido_por="Commandé par",
               detalles="Détails de la facture", envio="Livraison", pagina="Page",
               cab=[(38.1, 0, "Description"), (309.0, 0, "Qté"), (347.2, 0, "Prix Unitaire"), (378.1, 1, "HT"),
                    (411.5, 0, "Taux TVA"), (471.1, 0, "Prix Unitaire"), (497.4, 1, "TTC"),
                    (538.4, 0, "Total"), (542.6, 1, "TTC")]),
}


def importes(iva, unitario, cantidad):
    con_iva = round(unitario * (1 + iva / 100), 2)
    return con_iva, round(con_iva * cantidad, 2)


def dibujar_venta(v, rect=None):
    """Dibuja una venta (o una rectificativa si rect no es None) en el orden del texto de las reales."""
    e = ETIQUETAS[v.get("idioma", "es")]
    doc = fitz.open()
    p = Pagina(doc)
    signo = -1 if rect else 1
    iva_txt = f"{v['iva']}%" if v["iva"] else "0 %  (1)"
    lineas = [(d, a, c, u) + importes(v["iva"], u, c) for d, a, c, u in v["productos"]]
    envio_bruto = v.get("envio", 0.0)
    total = round(sum(l[5] for l in lineas) + envio_bruto, 2) * signo
    base = round(sum(l[3] * l[2] for l in lineas) + envio_bruto / (1 + v["iva"] / 100), 2) * signo

    # 1) título, pie y resumen de importes (en las reales van primero en el texto)
    titulo = e["titulo"] if not rect else ("Nota de crédito" if rect["variante"] == "nota_credito" else "Factura rectificativa")
    p.t(516.4 if not rect else 441.1, 29.1, titulo, 14.2, True)
    p.t(511.3, 777.8, f"{e['pagina']}  1  de  1", 9.0)
    y_res = 560 + 30 * len(lineas)
    p.t(298.4, y_res - 26, "Total", 11.2, True); p.t(522.6, y_res - 26, euros(total), 11.2, True)
    p.t(350.0, y_res, "IVA %"); p.t(447.6, y_res, "Precio total"); p.t(545.1, y_res, "IVA")
    p.t(440.5, y_res + 13.5, "(IVA excluido)")
    p.t(355.8, y_res + 34.5, iva_txt); p.t(462.1, y_res + 34.5, euros(base)); p.t(532.1, y_res + 34.5, euros(round(total - base, 2)))
    p.t(301.4, y_res + 56.3, "Total"); p.t(462.1, y_res + 56.3, euros(base)); p.t(532.1, y_res + 56.3, euros(round(total - base, 2)))

    # 2) datos del pedido
    p.t(38.1, 413.3, e["detalles"], 11.2, True)
    p.t(35.1, 367.9, e["f_pedido"]); p.t(153.6, 367.9, fecha_texto(date.fromordinal(v["fecha"].toordinal() - 1), v.get("idioma", "es")))
    p.t(35.1, 380.7, e["n_pedido"]); p.t(153.6, 380.7, v["pedido"])
    p.t(35.1, 393.4, e["pedido_por"]); p.t(153.6, 393.4, v["cliente"][0][:28])

    # 3) bloque superior derecho: vendedor, fecha y número
    y = 73.2
    if rect:
        p.t(342.0, 69.6, "Reembolso", 11.2, True)
        if rect["variante"] == "partida":
            p.t(342.0, 88.5, "Esta es una nota de crédito para factura")
            p.t(342.0, 102.1, rect["original"]); y = 115.7
        else:
            p.t(342.0, 88.5, f"Esta es una nota de crédito para factura  {rect['original']}"); y = 102.1
        p.t(342.0, y, f"{e['vendido']}  {EMPRESA['nombre']}"); p.t(342.0, y + 13.6, f"{e['iva']} ES{EMPRESA['nif']}")
        f_txt = fecha_puntos(rect["fecha"]) if rect["variante"] == "nota_credito" else fecha_texto(rect["fecha"])
        if rect["variante"] == "nota_credito":
            filas = [("Fecha de la nota de crédito", f_txt), ("Número de la nota de crédito", rect["numero"]),
                     ("Número de la factura original", rect["original"])]
        elif rect["variante"] == "rectificativa":
            filas = [("Fecha de la factura rectificativa", f_txt), ("Número de la factura rectificativa", rect["numero"]),
                     ("Número de la factura original", rect["original"])]
        else:
            filas = [(("Fecha de la factura", "rectificativa"), f_txt), (("Número de la factura", "rectificativa"), rect["numero"]),
                     (("Número de la factura", "original"), rect["original"])]
        yy = y + 41.5
        for etiqueta, valor in filas + [(e["total_pend"], euros(total))]:
            if isinstance(etiqueta, tuple):          # etiqueta partida en dos líneas, valor en medio
                p.t(344.3, yy, etiqueta[0]); p.t(344.3, yy + 10.5, etiqueta[1]); p.t(442.8, yy + 5.1, valor); yy += 24
            else:
                p.t(344.3, yy, etiqueta); p.t(470.3, yy, valor); yy += 13.5
    else:
        p.t(342.0, 73.2, e["pagado"], 11.2, True)
        if v.get("idioma", "es") == "es":          # las francesas reales no llevan referencia de pago
            p.t(342.0, 91.0, f"{e['ref_pago']} DemoPago{v['numero'][-8:]}")
        p.t(342.0, 104.5, f"{e['vendido']} {EMPRESA['nombre']}")
        p.t(342.0, 118.0, f"{e['iva']} ES{EMPRESA['nif']}")
        f_txt = fecha_texto(v["fecha"], v.get("idioma", "es")) if v["formato_fecha"] == "texto" else fecha_puntos(v["fecha"])
        p.t(344.3, 145.7, e["fecha"]); p.t(484.0, 145.7, f_txt)
        p.t(344.3, 159.2, e["numero"]); p.t(484.0, 159.2, v["numero"])
        p.t(344.3, 172.7, e["total_pend"]); p.t(484.0, 172.7, euros(total))

    # 4) dirección postal (arriba a la izquierda), contacto y bloques de direcciones
    for i, linea in enumerate(v["cliente"] + [v["pais"]]):
        p.t(62.7, 150.9 + 13.5 * i, linea.upper(), 9.0)
    p.t(34.3, 220.5, f"{e['contacto']} https://www.nubemarket.example/contacto")
    p.t(35.1, 242.1, e["dir_com"], 9.8, True)
    # Columnas de direcciones: ~170 pt de ancho; los textos largos se parten como en las reales
    bloque = [t for linea in v["cliente"] + [v["pais"], f"{v['nif_etiqueta']} {v['nif_impreso']}"] for t in partir(linea, 168)]
    for i, linea in enumerate(bloque):
        p.t(35.1, 257.3 + 13.5 * i, linea)
    p.t(210.6, 242.1, e["dir_env"], 9.8, True)
    for i, linea in enumerate([t for linea in v["cliente"] + [v["pais"]] for t in partir(linea, 168)]):
        p.t(210.6, 257.3 + 13.5 * i, linea)
    p.t(386.1, 242.1, e["vendido"], 9.8, True)
    for i, linea in enumerate([EMPRESA["nombre"], EMPRESA["calle"], EMPRESA["ciudad"], "ES", f"{e['iva']} ES{EMPRESA['nif']}"]):
        p.t(386.1, 257.3 + 13.5 * i, linea)
    p.t(34.3, 350.1, e["info"], 9.8, True)

    # 5) tabla de productos
    tablas = [1] if not rect or rect["variante"] != "rectificativa" else [1, -1]   # la "rectificativa" repite la tabla original
    y = 437.3
    for signo_tabla in tablas:
        for x, fila, texto in e["cab"]:
            p.t(x, y + 13.5 * fila, texto)
        y += 31.5
        for desc, asin, cant, unit, con_iva, tot in lineas:
            s = signo if signo_tabla < 0 or not rect or rect["variante"] != "rectificativa" else 1
            y_fila = y
            for d in desc:
                p.t(38.1, y, d); y += 13.5
            p.t(38.1, y, f"ASIN: {asin}"); y += 13.5
            if rect and rect["variante"] == "rectificativa" and signo_tabla < 0:
                p.t(38.1, y, "Razón de la rectificación: Devolución por el cliente"); y += 13.5
            p.t(321.2, y_fila, str(cant)); p.t(367.6, y_fila, euros(unit * s)); p.t(410.7, y_fila, iva_txt)
            p.t(467.4, y_fila, euros(con_iva * s)); p.t(532.1, y_fila, euros(tot * s))
            y += 7.5
        p.t(38.1, y, e["envio"])
        neto_envio = round(envio_bruto / (1 + v["iva"] / 100), 2)
        p.t(371.8, y, euros(neto_envio * (1 if signo_tabla > 0 else -1)))
        p.t(471.6, y, euros(envio_bruto)); p.t(536.3, y, euros(envio_bruto))
        y += 40
    p.cerrar()
    return doc


def esperado_venta(v):
    """Valores que el programa DEBERÍA producir (precio unitario sin IVA, como en los productos)."""
    extranjero = v["pais"] != "ES"
    cab = {
        "factura_numero": v["numero"], "pedido_numero": v["pedido"], "factura_fecha": esperado_fecha(v["fecha"]),
        "nombre": v["cliente"][0], "direccion": v["cliente"][1],
        "codigo_postal": "" if extranjero else v["cliente"][2][-5:], "nif": v["nif_esperado"],
        "linea_tipo_iva": "EXE" if extranjero else "ORD21", "linea_tipo_operacion": "VCEE" if extranjero else "VNAC",
    }
    lineas = [{"linea_descripcion": " ".join(d), "linea_unidades": str(c), "linea_precio": u} for d, _, c, u in v["productos"]]
    if v.get("envio"):
        lineas.append({"linea_descripcion": "Envío", "linea_unidades": "1",
                       "linea_precio": round(v["envio"] / (1 + v["iva"] / 100), 2)})
    return cab, lineas


# ----------------------------------------------------------------------------- compras
def dibujar_amazon(c):
    doc = fitz.open(); p = Pagina(doc); pr = PROVEEDOR_AMAZON
    total_tarifas = round(c["subtotal"] * 0.03, 2)
    cargos = round(c["subtotal"] - total_tarifas, 2)
    iva = round(c["subtotal"] * 0.21, 2)
    eur = lambda v: euros(v).replace(" €", "   EUR")
    p.t(51.2, 123.7, "Desde", 9, True); p.t(187.0, 123.7, "Para", 9, True); p.t(322.8, 123.7, "Información de facturación", 9, True)
    p.t(51.2, 144.0, pr["nombre"]); p.t(51.2, 153.8, pr["calle"]); p.t(51.2, 163.7, pr["ciudad"])
    p.t(51.2, 179.6, "Número de"); p.t(51.2, 189.5, "impuesto"); p.t(100.6, 177.7, pr["nif"])
    p.t(187.0, 144.0, EMPRESA["nombre"]); p.t(187.0, 153.8, EMPRESA["calle"]); p.t(187.0, 163.7, "Villademo Madrid ES 28999")
    p.t(187.0, 189.5, "Número de"); p.t(187.0, 199.3, "impuesto"); p.t(236.4, 187.5, f"ES{EMPRESA['nif']}")
    f = c["fecha"]
    for i, (etq, val) in enumerate([("Número de factura", c["numero"]), ("Fecha de la factura", f"{f.day:02d}-{f.month:02d}-{f.year}"),
                                    ("Periodo de facturación", f"{f.day:02d}-{f.month:02d}-{f.year} - {f.day:02d}-{f.month:02d}-{f.year}"),
                                    ("Moneda de la factura", "EUR")]):
        p.t(322.8, 146.0 + 9.8 * i, etq); p.t(411.7, 144.0 + 9.8 * i, val)
    p.t(51.2, 237.1, "Resumen de la factura", 11, True); p.t(434.7, 237.4, "Importe de factura adeudado")
    p.t(485.8, 251.8, eur(c["subtotal"] + iva))
    p.t(51.2, 283.5, "Cargos Totales De La Campaña"); p.t(504.9, 281.5, eur(cargos))
    p.t(51.2, 301.0, "Total de tarifas publicitarias regulatorias"); p.t(509.0, 299.0, eur(total_tarifas))
    p.t(51.2, 337.5, "VAT (21%) - SPAIN"); p.t(504.9, 335.5, eur(iva))
    p.t(51.2, 384.5, "Importe Total a Pagar", 9, True); p.t(504.9, 382.5, eur(c["subtotal"] + iva))
    p.t(51.2, 418.2, "Instrucciones de pago", 11, True); p.t(51.2, 442.4, "Tu método de pago")
    p.t(51.2, 548.5, "Resumen de cargos del"); p.t(51.2, 562.6, "portafolio")
    p.t(51.2, 580.3, "Los subtotales del portafolio incluyen los ajustes y cargos que se")
    p.t(51.2, 590.2, "reflejan en el subtotal del portafolio.")
    p.t(51.2, 319.2, "Subtotal"); p.t(504.9, 317.3, eur(c["subtotal"]))
    p.t(51.2, 356.4, "Subtotal de"); p.t(51.2, 366.2, "Impuesto"); p.t(504.9, 359.4, eur(iva))
    p.t(474.0, 50.9, "Número de factura"); p.t(477.7, 62.1, f"  {c['numero']}")
    p.t(459.0, 81.8, "Factura fiscal", 11, True)
    p.cerrar()
    return doc


def dibujar_meta(c):
    doc = fitz.open(); p = Pagina(doc); pr = PROVEEDOR_META
    total = round(sum(v for _, v in c["campanas"]), 2)
    f = c["fecha"]
    fecha_meta = f"{f.day} {MESES_META[f.month - 1]} {f.year}, {c['hora']}"
    importe = f"{total:.2f}".replace(".", ",")          # Meta escribe 1959,00 (sin punto de miles)
    p.t(50.0, 44.4, "Factura para LUMORA", 14, True)
    p.t(50.0, 59.8, "Identificador de cuenta: 0000000000000001")
    p.t(50.0, 107.3, "Fecha de la factura/pago", 9, True); p.t(50.0, 117.2, fecha_meta)
    p.t(50.0, 137.8, "Método de pago", 9, True); p.t(50.0, 147.7, "Tarjeta · 0000")
    p.t(50.0, 157.6, f"Número de referencia: {c['referencia']}")
    p.t(50.0, 178.1, "Identificador de la transacción", 9, True); p.t(50.0, 188.1, "0000000000000000-0000000000000000")
    p.t(50.0, 208.6, "Tipo de producto", 9, True); p.t(50.0, 218.5, "Anuncios en redes")
    p.t(507.4, 130.8, "Pagado", 9, True); p.t(384.0, 150.4, f"{importe} € (EUR)", 16, True)
    p.t(50.0, 282.0, "Campañas", 12, True)
    y = 306.8
    for nombre, valor in c["campanas"]:
        p.t(50.0, y, nombre, 9, True); p.t(50.0, y + 13.4, f"De 1 {MESES_META[f.month - 1]} {f.year}, 0:00 a {fecha_meta}")
        p.t(518.1, y + 6.4, euros(valor)); y += 59
    p.t(50.0, 725.9, pr["nombre"]); p.t(50.0, 745.8, pr["ciudad"]); p.t(50.0, 775.5, f"VAT Reg. No. {pr['nif']}")
    p.t(465.7, 755.7, f"Nº de IVA: ES{EMPRESA['nif']}")
    p.t(419.3, 765.6, f"Factura número {c['numero']}")
    p.t(497.1, 775.5, "Factura original")
    p.cerrar()
    return doc, total


# ----------------------------------------------------------------------------- principal
def main():
    if os.path.exists(SALIDA):
        shutil.rmtree(SALIDA)
    for sub in ("ventas", "compra-amazon", "compra-meta"):
        os.makedirs(os.path.join(SALIDA, sub))

    facturas, lineas, rects = [], [], []
    for v in VENTAS:
        dibujar_venta(v).save(os.path.join(SALIDA, "ventas", v["archivo"]))
        cab, lin = esperado_venta(v)
        facturas.append({"tipo": "venta", "archivo": v["archivo"], "serie": "DV26", "formato_fecha": v["formato_fecha"],
                         "idioma": v.get("idioma", "es"), "nif_impreso": v["nif_impreso"], **cab})
        for i, l in enumerate(lin, 1):
            lineas.append({"tipo": "venta", "archivo": v["archivo"], "factura_numero": v["numero"], "linea": i, **l})
    for r in RECTIFICATIVAS:
        dibujar_venta(r["base"], rect=r).save(os.path.join(SALIDA, "ventas", r["archivo"]))
        rects.append({"archivo": r["archivo"], "variante": r["variante"], "numero": r["numero"], "original": r["original"]})
    for c in COMPRAS_AMAZON:
        dibujar_amazon(c).save(os.path.join(SALIDA, "compra-amazon", c["archivo"]))
        facturas.append({"tipo": "amazon", "archivo": c["archivo"], "serie": "DA26", "factura_numero": c["numero"],
                         "factura_fecha": esperado_fecha(c["fecha"]), "nombre": PROVEEDOR_AMAZON["nombre"],
                         "direccion": PROVEEDOR_AMAZON["calle"], "codigo_postal": PROVEEDOR_AMAZON["cp"],
                         "nif": "B00000011", "linea_tipo_iva": "ORD21", "linea_tipo_operacion": "CNAC"})
        lineas.append({"tipo": "amazon", "archivo": c["archivo"], "factura_numero": c["numero"], "linea": 1,
                       "linea_descripcion": "Tarifas publicitarias", "linea_unidades": "1", "linea_precio": c["subtotal"]})
    for c in COMPRAS_META:
        doc, total = dibujar_meta(c)
        doc.save(os.path.join(SALIDA, "compra-meta", c["archivo"]))
        facturas.append({"tipo": "meta", "archivo": c["archivo"], "serie": "DM26", "factura_numero": c["numero"],
                         "pedido_numero": c["referencia"], "factura_fecha": esperado_fecha(c["fecha"]),
                         "nombre": PROVEEDOR_META["nombre"], "direccion": "1 Demo Street Dublin 2 Ireland",
                         "codigo_postal": "D02 X0X0", "nif": PROVEEDOR_META["nif"],
                         "linea_tipo_iva": "ORD21", "linea_tipo_operacion": "CCES"})
        lineas.append({"tipo": "meta", "archivo": c["archivo"], "factura_numero": c["numero"], "linea": 1,
                       "linea_descripcion": "Publicidad", "linea_unidades": "1", "linea_precio": total})

    with pd.ExcelWriter(os.path.join(AQUI, "ground_truth.xlsx")) as w:
        pd.DataFrame(facturas).to_excel(w, sheet_name="facturas", index=False)
        pd.DataFrame(lineas).to_excel(w, sheet_name="lineas", index=False)
        pd.DataFrame(rects).to_excel(w, sheet_name="rectificativas", index=False)
    print(f"✅ {len(VENTAS)} ventas, {len(RECTIFICATIVAS)} rectificativas, {len(COMPRAS_AMAZON)} compras tipo Amazon, "
          f"{len(COMPRAS_META)} compras tipo Meta, y ground_truth.xlsx ({len(facturas)} facturas, {len(lineas)} líneas)")


if __name__ == "__main__":
    main()
