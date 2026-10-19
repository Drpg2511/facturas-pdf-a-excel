"""
Ejecuta factuscan_demo.py sobre las facturas sintéticas y compara la salida con
ground_truth.xlsx.

    python demo-publica/verificar.py

1. Procesa facturas/ventas, facturas/compra-amazon y facturas/compra-meta con las
   mismas funciones que usa el menú del programa, y guarda los Excel en salida/.
2. Compara, por tipo de factura y por campo, lo obtenido con lo esperado.
3. Comprueba que las rectificativas quedaron fuera del Excel y aparecen en el
   listado aparte con su número y el de la factura original.

El informe se muestra en pantalla y se guarda en salida/informe_verificacion.txt.
"""
import contextlib
import glob
import importlib.util
import io
import os
import shutil
from collections import defaultdict

import pandas as pd

AQUI = os.path.dirname(os.path.abspath(__file__))
FACTURAS = os.path.join(AQUI, "facturas")
SALIDA = os.path.join(AQUI, "salida")
TIPOS = [("1", "venta", "ventas"), ("2", "amazon", "compra-amazon"), ("3", "meta", "compra-meta")]
CAMPOS_CABECERA = ["factura_numero", "pedido_numero", "factura_fecha", "nombre", "direccion", "codigo_postal",
                   "nif", "serie", "linea_tipo_iva", "linea_tipo_operacion"]
CAMPOS_LINEA = ["linea_descripcion", "linea_unidades", "linea_precio"]


def cargar_programa():
    spec = importlib.util.spec_from_file_location("factuscan_demo", os.path.join(AQUI, "factuscan_demo.py"))
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return modulo


def ejecutar(fs):
    """Procesa los tres tipos como lo haría el menú y guarda los Excel en salida/."""
    if os.path.exists(SALIDA):
        shutil.rmtree(SALIDA)
    os.makedirs(SALIDA)
    rectificativas, consola = [], io.StringIO()
    anterior = os.getcwd()
    os.chdir(SALIDA)
    try:
        with contextlib.redirect_stdout(consola):
            for choice, _, carpeta in TIPOS:
                pdfs = sorted(glob.glob(os.path.join(FACTURAS, carpeta, "*.pdf")))
                datos, _, excluidas, rects, _ = fs.process_files(choice, pdfs)
                rectificativas += rects
                if datos:
                    fs.save_to_excel(datos, choice)
    finally:
        os.chdir(anterior)
    with open(os.path.join(SALIDA, "consola.txt"), "w", encoding="utf-8") as f:
        f.write(consola.getvalue())
    salida = {}
    for choice, tipo, _ in TIPOS:
        archivo, hoja = fs.ARCHIVOS_SALIDA[choice]
        ruta = os.path.join(SALIDA, archivo)
        salida[tipo] = pd.read_excel(ruta, sheet_name=hoja, dtype=str).fillna("") if os.path.exists(ruta) else pd.DataFrame()
    return salida, rectificativas


def iguales(esperado, obtenido, campo):
    if campo == "linea_precio":
        try:
            return abs(float(esperado) - float(obtenido)) < 0.005
        except ValueError:
            return False
    return str(esperado).strip() == str(obtenido).strip()


def main():
    fs = cargar_programa()
    salida, rectificativas = ejecutar(fs)
    gt_fact = pd.read_excel(os.path.join(AQUI, "ground_truth.xlsx"), sheet_name="facturas", dtype=str).fillna("")
    gt_lin = pd.read_excel(os.path.join(AQUI, "ground_truth.xlsx"), sheet_name="lineas", dtype=str).fillna("")
    gt_rect = pd.read_excel(os.path.join(AQUI, "ground_truth.xlsx"), sheet_name="rectificativas", dtype=str).fillna("")

    conteo = defaultdict(lambda: [0, 0])          # (tipo, campo) -> [coinciden, fallan]
    fallos = []                                    # (tipo, archivo, campo, esperado, obtenido)

    def anotar(tipo, archivo, campo, esperado, obtenido):
        ok = iguales(esperado, obtenido, campo)
        conteo[(tipo, campo)][0 if ok else 1] += 1
        if not ok:
            fallos.append((tipo, archivo, campo, esperado, obtenido))

    for _, f in gt_fact.iterrows():
        tipo, out = f["tipo"], salida[f["tipo"]]
        filas = out[out["factura_numero"] == f["factura_numero"]] if len(out) else out
        cab = filas.iloc[0] if len(filas) else None
        for campo in CAMPOS_CABECERA:
            if campo in f and f[campo] != "" or campo in ("codigo_postal",) and tipo == "venta":
                anotar(tipo, f["archivo"], campo, f[campo], cab[campo] if cab is not None else "(factura ausente)")
        # líneas, en orden
        esperadas = gt_lin[gt_lin["archivo"] == f["archivo"]].reset_index(drop=True)
        for i in range(max(len(esperadas), len(filas))):
            for campo in CAMPOS_LINEA:
                e = esperadas.iloc[i][campo] if i < len(esperadas) else "(no debería existir)"
                o = filas.iloc[i][campo] if i < len(filas) else "(falta la línea)"
                anotar(tipo, f"{f['archivo']} línea {i + 1}", campo, e, o)

    # rectificativas: fuera del Excel y en el listado aparte
    listadas = {archivo: (numero, original) for archivo, numero, original in rectificativas}
    numeros_venta = set(salida["venta"]["factura_numero"]) if len(salida["venta"]) else set()
    for _, r in gt_rect.iterrows():
        anotar("rectificativa", r["archivo"], "fuera_del_excel", "sí", "no" if r["numero"] in numeros_venta else "sí")
        num, orig = listadas.get(r["archivo"], ("(no listada)", "(no listada)"))
        anotar("rectificativa", r["archivo"], "numero_listado", r["numero"], num)
        anotar("rectificativa", r["archivo"], "original_listada", r["original"], orig)

    # informe
    lineas = ["VERIFICACIÓN DE FACTUSCAN (DEMO) CONTRA ground_truth.xlsx", "=" * 60]
    for tipo in ("venta", "amazon", "meta", "rectificativa"):
        campos = [(c, v) for (t, c), v in conteo.items() if t == tipo]
        ok = sum(v[0] for _, v in campos); ko = sum(v[1] for _, v in campos)
        lineas.append(f"\n[{tipo.upper()}]  {ok} coinciden, {ko} fallan")
        for campo, (a, b) in campos:
            lineas.append(f"   {campo:<22} {a:>3} ok   {b:>3} fallan" + ("   <--" if b else ""))
    lineas.append("\nFALLOS (esperado -> obtenido)")
    if not fallos:
        lineas.append("   ninguno")
    for tipo, archivo, campo, e, o in fallos:
        lineas.append(f"   [{tipo}] {archivo} · {campo}: {e!r} -> {o!r}")
    total_ok = sum(v[0] for v in conteo.values()); total_ko = sum(v[1] for v in conteo.values())
    lineas.append(f"\nTOTAL: {total_ok} comprobaciones correctas, {total_ko} fallos")
    informe = "\n".join(lineas)
    print(informe)
    with open(os.path.join(SALIDA, "informe_verificacion.txt"), "w", encoding="utf-8") as f:
        f.write(informe + "\n")


if __name__ == "__main__":
    main()
