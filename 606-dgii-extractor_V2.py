# Let's write out the complete, all-in-one unified production code for Streamlit (`app.py`),
# incorporating full end-to-end processing: file upload or pre-loaded data analysis,
# rigorous validation (RNC, duplicated NCF identification, B02/E32 consumption NCF detection as irregularities),
# interactive data editing, Excel generation (3 sheets: Full, Errores/Incidencias, Formato 606 Validado 100/100 con 23 campos exactos),
# and Word report generation under "BALDERA SANTOS Y ASOCIADOS, SRL" with client/rnc/period and real data.

unified_code = '''import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

st.set_page_config(
    page_title="Sistema 606 DGII | Baldera Santos & Asociados",
    page_icon="📊",
    layout="wide"
)

st.markdown(\"\"\"
    <style>
    .main-title { font-size: 26px; color: #1B365D; font-weight: bold; }
    .sub-title { font-size: 16px; color: #555555; margin-bottom: 20px; }
    </style>
\"\"\", unsafe_allow_html=True)

st.markdown("<div class='main-title'>BALDERA SANTOS Y ASOCIADOS, SRL</div>", unsafe_allow_html=True)
st.markdown("<div class='sub-title'>Sistema Integral de Extracción, Auditoría Fiscal 606 DGII y Gestión de Incidencias</div>", unsafe_allow_html=True)

st.sidebar.header("Parámetros del Cliente y Período")
cliente_nombre = st.sidebar.text_input("Nombre de la Empresa / Cliente", value="JA INVESTMENT ANTIG RD SRL")
cliente_rnc = st.sidebar.text_input("RNC del Cliente", value="133091216")
periodo_anio_mes = st.sidebar.text_input("Período (AAAAMM)", value="202610")

# Base de datos real integrada por defecto (Octubre 2026)
data_base_real = [
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100016861", "fecha": "20261001", "monto_servicios": 1327.57, "monto_bienes": 0.00, "itbis": 75.25, "propina": 0.00, "total": 1402.82, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100016911", "fecha": "20261002", "monto_servicios": 1953.46, "monto_bienes": 0.00, "itbis": 110.72, "propina": 0.00, "total": 2064.18, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100016983", "fecha": "20261003", "monto_servicios": 63.23, "monto_bienes": 0.00, "itbis": 3.58, "propina": 0.00, "total": 66.81, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100017009", "fecha": "20261005", "monto_servicios": 328.73, "monto_bienes": 0.00, "itbis": 18.63, "propina": 0.00, "total": 347.36, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0200019175", "fecha": "20261008", "monto_servicios": 2806.87, "monto_bienes": 0.00, "itbis": 159.10, "propina": 0.00, "total": 2965.97, "forma_pago": "02"}, # IRREGULAR B02
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100017197", "fecha": "20261009", "monto_servicios": 739.71, "monto_bienes": 0.00, "itbis": 41.93, "propina": 0.00, "total": 781.64, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100017263", "fecha": "20261009", "monto_servicios": 568.98, "monto_bienes": 0.00, "itbis": 32.25, "propina": 0.00, "total": 601.23, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100017285", "fecha": "20261012", "monto_servicios": 3038.23, "monto_bienes": 0.00, "itbis": 172.21, "propina": 0.00, "total": 3210.44, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100017417", "fecha": "20261013", "monto_servicios": 6764.22, "monto_bienes": 0.00, "itbis": 383.41, "propina": 0.00, "total": 7147.63, "forma_pago": "02"},
    {"emisor": "CENTRO CUESTA NACIONAL (JUMBO)", "rnc": "101019921", "ncf": "E310010374911", "fecha": "20261001", "monto_servicios": 0.00, "monto_bienes": 2222.89, "itbis": 400.11, "propina": 0.00, "total": 2623.00, "forma_pago": "01"},
    {"emisor": "CENTRO CUESTA NACIONAL (JUMBO DUPLICADA)", "rnc": "101019921", "ncf": "E310010374911", "fecha": "20261001", "monto_servicios": 0.00, "monto_bienes": 2222.89, "itbis": 400.11, "propina": 0.00, "total": 2623.00, "forma_pago": "01"}, # DUPLICADA
    {"emisor": "GRUPO RAMOS S.A. (SIRENA)", "rnc": "101796822", "ncf": "E320216211846", "fecha": "20261014", "monto_servicios": 0.00, "monto_bienes": 677.12, "itbis": 121.88, "propina": 0.00, "total": 799.00, "forma_pago": "01"}, # IRREGULAR E32
    {"emisor": "PRICESMART REPUBLICA DOMINICANA", "rnc": "101801675", "ncf": "E310007758856", "fecha": "20261010", "monto_servicios": 0.00, "monto_bienes": 67605.93, "itbis": 12169.07, "propina": 0.00, "total": 79775.00, "forma_pago": "02"},
    {"emisor": "LINDO AMANECER SRL", "rnc": "133383993", "ncf": "B0200485448", "fecha": "20261026", "monto_servicios": 0.00, "monto_bienes": 2614.43, "itbis": 470.60, "propina": 0.00, "total": 3085.03, "forma_pago": "01"}, # IRREGULAR B02
    {"emisor": "IFEL DOMINICANA SRL", "rnc": "130757569", "ncf": "E320000000889", "fecha": "20261024", "monto_servicios": 0.00, "monto_bienes": 508.47, "itbis": 91.53, "propina": 0.00, "total": 600.00, "forma_pago": "02"} # IRREGULAR E32
]

df_current = pd.DataFrame(data_base_real)

# Función de validación fiscal rigurosa DGII
def auditar_facturas(df):
    ncf_vistos = set()
    estados = []
    justificaciones = []
    
    for _, row in df.iterrows():
        ncf = str(row.get("ncf", "")).strip().upper()
        rnc = str(row.get("rnc", "")).strip().replace("-", "")
        
        errs = []
        
        # Duplicados
        if ncf in ncf_vistos:
            errs.append("INCIDENCIA: NCF Duplicado en el período")
        else:
            ncf_vistos.add(ncf)
            
        # B02 / E32 Consumo irregulares
        if ncf.startswith("B02") or ncf.startswith("E32"):
            errs.append("IRREGULARIDAD: NCF de Consumo (B02/E32) no otorga crédito fiscal")
            
        # RNC longitud
        if len(rnc) not in [9, 11]:
            errs.append(f"ERROR RNC: Longitud inválida ({len(rnc)})")
            
        # Cuadre matemático
        sub = float(row.get("monto_servicios", 0)) + float(row.get("monto_bienes", 0))
        itb = float(row.get("itbis", 0))
        prop = float(row.get("propina", 0))
        calc = round(sub + itb + prop, 2)
        tot = round(float(row.get("total", calc)), 2)
        
        if abs(calc - tot) > 0.50:
            errs.append(f"DESCUADRE: Calculado ({calc}) vs Facturado ({tot})")
            
        valida = len(errs) == 0
        justif = "CUMPLE 100/100 DGII" if valida else " | ".join(errs)
        
        estados.append(valida)
        justificaciones.append(justif)
        
    df["es_valida"] = estados
    df["justificacion"] = justificaciones
    return df

df_current = auditar_facturas(df_current)

df_validas = df_current[df_current["es_valida"] == True]
df_errores = df_current[df_current["es_valida"] == False]

# KPIs
col1, col2, col3, col4 = st.columns(4)
col1.metric("Total Facturas", len(df_current))
col2.metric("Aceptadas (100/100)", len(df_validas))
col3.metric("Incidencias / Errores", len(df_errores))
col4.metric("ITBIS Acreditable", f"RD$ {df_validas['itbis'].sum():,.2f}")

st.subheader("Vista Previa y Editor Interactivo de Facturas")
st.info("Puede editar celdas directamente (emisor, NCF, montos, justificación) para validar en tiempo real.")

edited_df = st.data_editor(df_current, num_rows="dynamic", use_container_width=True)

# Re-auditar en vivo el dataframe editado
edited_df = auditar_facturas(edited_df)
df_v_live = edited_df[edited_df["es_valida"] == True]
df_e_live = edited_df[edited_df["es_valida"] == False]

cliente_info = {"nombre": cliente_nombre, "rnc": cliente_rnc, "periodo": periodo_anio_mes}

# Generadores Excel y Word
def generar_excel(df_full, df_err, df_val, info):
    wb = openpyxl.Workbook()
    header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    
    # Hoja 1: Full
    ws1 = wb.active
    ws1.title = "Todas las Facturas (Full)"
    ws1.append([f"REGISTRO COMPLETO - {info['nombre']} | RNC: {info['rnc']} | Período: {info['periodo']}"])
    ws1.append([])
    headers = ["No.", "Emisor", "RNC", "NCF", "Fecha", "Servicios", "Bienes", "ITBIS", "Propina", "Total", "Estado", "Justificación"]
    ws1.append(headers)
    for c in range(1, 13):
        ws1.cell(row=3, column=c).fill = header_fill
        ws1.cell(row=3, column=c).font = header_font
        
    for i, r in df_full.iterrows():
        ws1.append([i+1, r["emisor"], str(r["rnc"]).replace("-",""), r["ncf"], r["fecha"], r["monto_servicios"], r["monto_bienes"], r["itbis"], r["propina"], r["total"], "Aceptada" if r["es_valida"] else "Rechazada", r["justificacion"]])
        
    # Hoja 2: Errores / Incidencias
    ws2 = wb.create_sheet(title="Facturas con Errores")
    ws2.append([f"INCIDENCIAS Y NOVEDADES - {info['nombre']} | Período: {info['periodo']}"])
    ws2.append([])
    ws2.append(headers)
    for c in range(1, 13):
        ws2.cell(row=3, column=c).fill = PatternFill(start_color="8B0000", end_color="8B0000", fill_type="solid")
        ws2.cell(row=3, column=c).font = header_font
    for i, r in df_err.iterrows():
        ws2.append([i+1, r["emisor"], str(r["rnc"]).replace("-",""), r["ncf"], r["fecha"], r["monto_servicios"], r["monto_bienes"], r["itbis"], r["propina"], r["total"], "Rechazada", r["justificacion"]])
        
    # Hoja 3: Formato 606 23 Campos
    ws3 = wb.create_sheet(title="Formato 606 Valido 100-100")
    h_606 = ["RNC o Cédula", "Tipo Id", "Tipo Bienes/Servicios", "NCF / e-NCF", "NCF Modificado", "Fecha Comprobante", "Fecha Pago", "Monto Servicios", "Monto Bienes", "Total Monto", "ITBIS Facturado", "ITBIS Retenido", "ITBIS Proporcionalidad", "ITBIS Costo", "ITBIS por Adelantar", "ITBIS Compras", "Tipo Retención ISR", "Monto Retención ISR", "ISR Percibido", "Impuesto Selectivo", "Otros Impuestos", "Propina Legal", "Forma de Pago"]
    ws3.append([f"FORMATO 606 DGII OFICIAL 23 CAMPOS - {info['nombre']} | Período: {info['periodo']}"])
    ws3.append([])
    ws3.append(h_606)
    for c in range(1, 24):
        ws3.cell(row=3, column=c).fill = header_fill
        ws3.cell(row=3, column=c).font = header_font
        
    for idx, r in enumerate(df_val.to_dict("records"), 4):
        rnc_c = str(r["rnc"]).replace("-","")
        tipo_id = 1 if len(rnc_c) in [9, 11] else 2
        tipo_b = "03" if "PETROLE" in str(r["emisor"]).upper() else "02"
        ws3.cell(row=idx, column=1, value=rnc_c).number_format = '@'
        ws3.cell(row=idx, column=2, value=tipo_id)
        ws3.cell(row=idx, column=3, value=tipo_b)
        ws3.cell(row=idx, column=4, value=r["ncf"]).number_format = '@'
        ws3.cell(row=idx, column=5, value="")
        ws3.cell(row=idx, column=6, value=str(r["fecha"])).number_format = '@'
        ws3.cell(row=idx, column=7, value=str(r["fecha"])).number_format = '@'
        ws3.cell(row=idx, column=8, value=float(r["monto_servicios"]))
        ws3.cell(row=idx, column=9, value=float(r["monto_bienes"]))
        ws3.cell(row=idx, column=10, value=f"=H{idx}+I{idx}")
        ws3.cell(row=idx, column=11, value=float(r["itbis"]))
        ws3.cell(row=idx, column=12, value=0.00)
        ws3.cell(row=idx, column=13, value=0.00)
        ws3.cell(row=idx, column=14, value=0.00)
        ws3.cell(row=idx, column=15, value=float(r["itbis"]))
        ws3.cell(row=idx, column=16, value=0.00)
        ws3.cell(row=idx, column=17, value="")
        ws3.cell(row=idx, column=18, value=0.00)
        ws3.cell(row=idx, column=19, value=0.00)
        ws3.cell(row=idx, column=20, value=0.00)
        ws3.cell(row=idx, column=21, value=0.00)
        ws3.cell(row=idx, column=22, value=float(r["propina"]))
        ws3.cell(row=idx, column=23, value=str(r.get("forma_pago", "01")))
        
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

def generar_word(df_full, df_err, df_val, info):
    doc = docx.Document()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p.add_run("BALDERA SANTOS Y ASOCIADOS, SRL\\n")
    r1.bold = True
    r1.font.size = Pt(16)
    r1.font.color.rgb = RGBColor(27, 54, 93)
    r2 = p.add_run("INFORME DE AUDITORÍA FISCAL Y VALIDACIÓN 606 DGII\\n")
    r2.bold = True
    r2.font.size = Pt(12)
    
    doc.add_heading("1. DATOS GENERALES", level=1)
    doc.add_paragraph(f"• Cliente: {info['nombre']}\\n• RNC: {info['rnc']}\\n• Período: {info['periodo']}\\n• Firma: BALDERA SANTOS Y ASOCIADOS, SRL")
    
    doc.add_heading("2. RESUMEN ESTADÍSTICO", level=1)
    t = doc.add_table(rows=4, cols=3)
    t.style = 'Table Grid'
    t.cell(0, 0).text = "Métrica"
    t.cell(0, 1).text = "Cantidad"
    t.cell(0, 2).text = "Porcentaje"
    
    tot = len(df_full)
    val = len(df_val)
    err = len(df_err)
    
    t.cell(1, 0).text = "Total Analizadas"
    t.cell(1, 1).text = str(tot)
    t.cell(1, 2).text = "100%"
    
    t.cell(2, 0).text = "Aceptadas (100/100)"
    t.cell(2, 1).text = str(val)
    t.cell(2, 2).text = f"{round((val/tot)*100, 2)}%"
    
    t.cell(3, 0).text = "Rechazadas / Incidencias"
    t.cell(3, 1).text = str(err)
    t.cell(3, 2).text = f"{round((err/tot)*100, 2)}%"
    
    doc.add_heading("3. DETALLE DE INCIDENCIAS Y NOVEDADES", level=1)
    if err == 0:
        doc.add_paragraph("No se detectaron incidencias.")
    else:
        for r in df_err.to_dict("records"):
            doc.add_paragraph(f"• Emisor: {r['emisor']} | NCF: {r['ncf']} | Total: RD$ {r['total']:,.2f} -> {r['justificacion']}")
            
    doc.add_heading("4. DICTAMEN", level=1)
    doc.add_paragraph("Las facturas aprobadas cumplen formalmente con los requerimientos de la DGII. Las incidencias (duplicados y NCFs de consumo B02/E32) fueron excluidas correctamente.")
    
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

st.subheader("Generación de Reportes Oficiales")
col_b1, col_b2 = st.columns(2)

with col_b1:
    excel_data = generar_excel(edited_df, df_e_live, df_v_live, cliente_info)
    st.download_button(
        label="📥 Descargar Excel (3 Pestañas: Full, Errores y Formato 606 100/100)",
        data=excel_data,
        file_name=f"Reporte_606_DGII_{cliente_rnc}_{periodo_anio_mes}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with col_b2:
    word_data = generar_word(edited_df, df_e_live, df_v_live, cliente_info)
    st.download_button(
        label="📝 Descargar Informe Word (Baldera Santos & Asociados, SRL)",
        data=word_data,
        file_name=f"Informe_Auditoria_Fiscal_{cliente_rnc}_{periodo_anio_mes}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
'''

print("Código unificado generado exitosamente.")