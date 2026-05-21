import streamlit as st
import pdfplumber
import re
import io
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.styles.numbers import FORMAT_NUMBER_COMMA_SEPARATED1

st.set_page_config(page_title="Extractor de Piezas GNP", page_icon="🔧", layout="centered")
st.title("🔧 Extractor de Piezas Sustituidas")
st.markdown("**Valuaciones GNP / Audatex** — Sube tu PDF y descarga el Excel al instante.")
st.divider()

def extraer_texto(contenido_bytes):
    texto_completo = []
    with pdfplumber.open(io.BytesIO(contenido_bytes)) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if texto:
                texto_completo.append(texto)
    return "\n".join(texto_completo)

def parsear_piezas(texto):
    lineas = texto.splitlines()

    inicio = None
    for i, linea in enumerate(lineas):
        if "PIEZAS SUSTITUIDAS" in linea.upper():
            inicio = i + 1
            break
    if inicio is None:
        return []

    fin = len(lineas)
    for i in range(inicio, len(lineas)):
        if re.search(r"ahorro|sub\s*total|total\s*piezas", lineas[i], re.IGNORECASE):
            fin = i
            break

    bloque = lineas[inicio:fin]
    piezas = []

    for linea in bloque:
        linea = linea.strip()
        if not linea:
            continue

        m_precio = re.match(r'^([\$][\d,]+\.\d{2})[\*A-Za-z]?\s+', linea)
        if not m_precio:
            continue

        precio_str = m_precio.group(1)
        # Convertir a número: quitar $ y comas
        try:
            precio_num = float(precio_str.replace('$', '').replace(',', ''))
        except:
            precio_num = 0.0

        resto  = linea[m_precio.end():]
        tokens = resto.split()
        if len(tokens) < 4:
            continue

        descripcion = ' '.join(tokens[1:-2]).strip()
        if not descripcion:
            continue
        if re.match(r'^(precio|referencia|descripci)', descripcion, re.IGNORECASE):
            continue

        piezas.append({"precio": precio_num, "descripcion": descripcion})

    return piezas

def generar_excel(piezas, nombre):
    wb = Workbook()
    ws = wb.active
    ws.title = "Piezas Sustituidas"

    azul  = "1F4E79"
    claro = "D9E1F2"
    fill_hdr  = PatternFill("solid", start_color=azul,     end_color=azul)
    fill_alt  = PatternFill("solid", start_color=claro,    end_color=claro)
    fill_blco = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")
    fill_tot  = PatternFill("solid", start_color=azul,     end_color=azul)
    borde = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin")
    )

    # Título
    ws.merge_cells("A1:B1")
    ws["A1"] = f"PIEZAS SUSTITUIDAS  |  {nombre}"
    ws["A1"].font      = Font(name="Arial", bold=True, size=13, color=azul)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    # Encabezados
    for col, txt in [("A","Precio"), ("B","Descripción")]:
        c = ws[f"{col}2"]
        c.value     = txt
        c.fill      = fill_hdr
        c.font      = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = borde
    ws.row_dimensions[2].height = 20

    # Datos
    for i, p in enumerate(piezas, start=3):
        fill = fill_alt if i % 2 == 0 else fill_blco

        # Precio como número con formato de moneda mexicana
        cp = ws.cell(row=i, column=1, value=p["precio"])
        cp.number_format = '"$"#,##0.00'
        cp.fill      = fill
        cp.font      = Font(name="Arial", size=10)
        cp.border    = borde
        cp.alignment = Alignment(horizontal="right", vertical="center")

        cd = ws.cell(row=i, column=2, value=p["descripcion"])
        cd.fill      = fill
        cd.font      = Font(name="Arial", size=10)
        cd.border    = borde
        cd.alignment = Alignment(horizontal="left", vertical="center")

    # Fila de TOTAL con fórmula SUMA
    fila_datos_inicio = 3
    fila_datos_fin    = len(piezas) + 2
    ft = len(piezas) + 3

    # Celda de etiqueta
    ws.cell(row=ft, column=2, value="TOTAL").font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    ws.cell(row=ft, column=2).fill      = fill_tot
    ws.cell(row=ft, column=2).alignment = Alignment(horizontal="right", vertical="center")
    ws.cell(row=ft, column=2).border    = borde

    # Celda de suma
    ct = ws.cell(row=ft, column=1,
                 value=f"=SUM(A{fila_datos_inicio}:A{fila_datos_fin})")
    ct.number_format = '"$"#,##0.00'
    ct.font      = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    ct.fill      = fill_tot
    ct.alignment = Alignment(horizontal="right", vertical="center")
    ct.border    = borde
    ws.row_dimensions[ft].height = 20

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 40

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Interfaz ─────────────────────────────────────────────────
pdf_file = st.file_uploader(
    "📄 Sube tu valuación en PDF",
    type=["pdf"],
    help="Archivos PDF de valuaciones GNP / Audatex"
)

if pdf_file is not None:
    with st.spinner("Procesando PDF..."):
        contenido = pdf_file.read()
        texto     = extraer_texto(contenido)
        piezas    = parsear_piezas(texto)

    if not piezas:
        st.error("❌ No se encontró la sección 'PIEZAS SUSTITUIDAS'. Verifica que sea una valuación GNP/Audatex.")
    else:
        total = sum(p["precio"] for p in piezas)
        st.success(f"✅ {len(piezas)} piezas encontradas  |  Total: ${total:,.2f}")
        st.subheader("Piezas sustituidas")
        st.dataframe(
            {"Precio": [f"${p['precio']:,.2f}" for p in piezas],
             "Descripción": [p["descripcion"] for p in piezas]},
            use_container_width=True,
            hide_index=True
        )
        nombre_base  = Path(pdf_file.name).stem
        excel_buf    = generar_excel(piezas, nombre_base)
        nombre_excel = nombre_base + "_piezas.xlsx"
        st.download_button(
            label="📥 Descargar Excel",
            data=excel_buf,
            file_name=nombre_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

st.divider()
st.caption("Vanguardia Body & Paint — Extractor de valuaciones GNP")

