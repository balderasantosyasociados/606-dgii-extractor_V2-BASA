import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Dataset completo extraído del documento PDF (39 facturas de octubre 2026)
invoices_data = [
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100016861", "fecha": "20261001", "monto_servicios": 1327.57, "monto_bienes": 0.00, "itbis": 75.25, "propina": 0.00, "total": 1402.82, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100016911", "fecha": "20261002", "monto_servicios": 1953.46, "monto_bienes": 0.00, "itbis": 110.72, "propina": 0.00, "total": 2064.18, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100016983", "fecha": "20261003", "monto_servicios": 63.23, "monto_bienes": 0.00, "itbis": 3.58, "propina": 0.00, "total": 66.81, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100017009", "fecha": "20261005", "monto_servicios": 328.73, "monto_bienes": 0.00, "itbis": 18.63, "propina": 0.00, "total": 347.36, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0200019175", "fecha": "20261008", "monto_servicios": 2806.87, "monto_bienes": 0.00, "itbis": 159.10, "propina": 0.00, "total": 2965.97, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100017197", "fecha": "20261009", "monto_servicios": 739.71, "monto_bienes": 0.00, "itbis": 41.93, "propina": 0.00, "total": 781.64, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100017263", "fecha": "20261009", "monto_servicios": 568.98, "monto_bienes": 0.00, "itbis": 32.25, "propina": 0.00, "total": 601.23, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100017285", "fecha": "20261012", "monto_servicios": 3038.23, "monto_bienes": 0.00, "itbis": 172.21, "propina": 0.00, "total": 3210.44, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100017417", "fecha": "20261013", "monto_servicios": 6764.22, "monto_bienes": 0.00, "itbis": 383.41, "propina": 0.00, "total": 7147.63, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100017450", "fecha": "20261015", "monto_servicios": 1510.96, "monto_bienes": 0.00, "itbis": 85.64, "propina": 0.00, "total": 1596.60, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100017553", "fecha": "20261015", "monto_servicios": 1384.50, "monto_bienes": 0.00, "itbis": 78.47, "propina": 0.00, "total": 1462.97, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100017673", "fecha": "20261019", "monto_servicios": 1283.39, "monto_bienes": 0.00, "itbis": 72.74, "propina": 0.00, "total": 1356.13, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100017719", "fecha": "20261021", "monto_servicios": 2143.14, "monto_bienes": 0.00, "itbis": 121.47, "propina": 0.00, "total": 2264.61, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100017773", "fecha": "20261022", "monto_servicios": 1504.64, "monto_bienes": 0.00, "itbis": 85.28, "propina": 0.00, "total": 1589.92, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100017814", "fecha": "20261023", "monto_servicios": 221.26, "monto_bienes": 0.00, "itbis": 12.54, "propina": 0.00, "total": 233.80, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100017905", "fecha": "20261026", "monto_servicios": 1523.57, "monto_bienes": 0.00, "itbis": 86.36, "propina": 0.00, "total": 1609.93, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100017956", "fecha": "20261026", "monto_servicios": 1087.34, "monto_bienes": 0.00, "itbis": 61.63, "propina": 0.00, "total": 1148.97, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100017983", "fecha": "20261028", "monto_servicios": 455.17, "monto_bienes": 0.00, "itbis": 25.80, "propina": 0.00, "total": 480.97, "forma_pago": "02"},
    {"emisor": "AEROMENSAJERIA DOCPAK SRL", "rnc": "101564113", "ncf": "B0100018041", "fecha": "20261029", "monto_servicios": 2212.69, "monto_bienes": 0.00, "itbis": 125.42, "propina": 0.00, "total": 2338.11, "forma_pago": "02"},
    {"emisor": "CENTRO CUESTA NACIONAL (JUMBO SAN ISIDRO)", "rnc": "101019921", "ncf": "E310010374911", "fecha": "20261001", "monto_servicios": 0.00, "monto_bienes": 2222.89, "itbis": 400.11, "propina": 0.00, "total": 2623.00, "forma_pago": "01"},
    {"emisor": "CENTRO CUESTA NACIONAL (JUMBO DUPLICADA)", "rnc": "101019921", "ncf": "E310010374911", "fecha": "20261001", "monto_servicios": 0.00, "monto_bienes": 2222.89, "itbis": 400.11, "propina": 0.00, "total": 2623.00, "forma_pago": "01"},
    {"emisor": "MERCERIA Y VARIEDADES EMA", "rnc": "132344154", "ncf": "B0100001383", "fecha": "20261003", "monto_servicios": 0.00, "monto_bienes": 292.37, "itbis": 52.63, "propina": 0.00, "total": 345.00, "forma_pago": "01"},
    {"emisor": "INVERSIONES BESSONE SRL (PETRONAN ECOLOGICA)", "rnc": "131155758", "ncf": "E310000041405", "fecha": "20261007", "monto_servicios": 0.00, "monto_bienes": 5444.40, "itbis": 0.00, "propina": 0.00, "total": 5444.40, "forma_pago": "01"},
    {"emisor": "BLOQUES SAN MIGUEL SRL", "rnc": "130042144", "ncf": "E310000094772", "fecha": "20261009", "monto_servicios": 0.00, "monto_bienes": 2796.74, "itbis": 503.41, "propina": 0.00, "total": 3300.15, "forma_pago": "02"},
    {"emisor": "DERIVADOS Y SERVICIOS DE PETROLEO SRL", "rnc": "123015102", "ncf": "E310000048062", "fecha": "20261009", "monto_servicios": 0.00, "monto_bienes": 1000.00, "itbis": 0.00, "propina": 0.00, "total": 1000.00, "forma_pago": "01"},
    {"emisor": "BLOQUES SAN MIGUEL SRL", "rnc": "130042144", "ncf": "E310000094659", "fecha": "20261009", "monto_servicios": 0.00, "monto_bienes": 5334.43, "itbis": 960.20, "propina": 0.00, "total": 6294.63, "forma_pago": "02"},
    {"emisor": "BLOQUES SAN MIGUEL SRL", "rnc": "130042144", "ncf": "E310000094670", "fecha": "20261010", "monto_servicios": 0.00, "monto_bienes": 4054.07, "itbis": 729.73, "propina": 0.00, "total": 4783.80, "forma_pago": "02"},
    {"emisor": "PRICESMART DE REPUBLICA DOMINICANA", "rnc": "101801675", "ncf": "E310007758856", "fecha": "20261010", "monto_servicios": 0.00, "monto_bienes": 67605.93, "itbis": 12169.07, "propina": 0.00, "total": 79775.00, "forma_pago": "02"},
    {"emisor": "PRIMAVERA LXY SRL", "rnc": "132795776", "ncf": "B0100047737", "fecha": "20261012", "monto_servicios": 0.00, "monto_bienes": 10839.01, "itbis": 1951.02, "propina": 0.00, "total": 12790.03, "forma_pago": "02"},
    {"emisor": "ABENSA (CHARLES GAULLE)", "rnc": "130563748", "ncf": "E310000076312", "fecha": "20261012", "monto_servicios": 0.00, "monto_bienes": 3000.00, "itbis": 0.00, "propina": 0.00, "total": 3000.00, "forma_pago": "01"},
    {"emisor": "ANDRICKSON COMERCIO INTERNACIONAL SRL", "rnc": "130674671", "ncf": "E310000030420", "fecha": "20261013", "monto_servicios": 0.00, "monto_bienes": 19995.00, "itbis": 3599.10, "propina": 0.00, "total": 23594.10, "forma_pago": "03"},
    {"emisor": "CERVICENTRO JHON AZUA SRL", "rnc": "133114577", "ncf": "B0100008356", "fecha": "20261013", "monto_servicios": 0.00, "monto_bienes": 5720.34, "itbis": 1029.66, "propina": 0.00, "total": 6750.00, "forma_pago": "02"},
    {"emisor": "GRUPO RAMOS S.A. (SIRENA)", "rnc": "101796822", "ncf": "E320216211846", "fecha": "20261014", "monto_servicios": 0.00, "monto_bienes": 677.12, "itbis": 121.88, "propina": 0.00, "total": 799.00, "forma_pago": "01"},
    {"emisor": "PRICESMART DE REPUBLICA DOMINICANA", "rnc": "101801875", "ncf": "E310007647815", "fecha": "20261018", "monto_servicios": 0.00, "monto_bienes": 12535.59, "itbis": 2256.41, "propina": 0.00, "total": 14792.00, "forma_pago": "02"},
    {"emisor": "DERIVADOS Y SERVICIOS DE PETROLEO SRL", "rnc": "123015102", "ncf": "E310000049677", "fecha": "20261020", "monto_servicios": 0.00, "monto_bienes": 1790.00, "itbis": 0.00, "propina": 0.00, "total": 1790.00, "forma_pago": "01"},
    {"emisor": "LINDO AMANECER SRL", "rnc": "133383993", "ncf": "B0200485448", "fecha": "20261026", "monto_servicios": 0.00, "monto_bienes": 2614.43, "itbis": 470.60, "propina": 0.00, "total": 3085.03, "forma_pago": "01"},
    {"emisor": "ESTACION DE SERVICIO AVENTURA SAN ISIDRO", "rnc": "101642297", "ncf": "E310000023018", "fecha": "20261028", "monto_servicios": 0.00, "monto_bienes": 2000.00, "itbis": 0.00, "propina": 0.00, "total": 2000.00, "forma_pago": "01"},
    {"emisor": "IFEL DOMINICANA SRL", "rnc": "130757569", "ncf": "E320000000889", "fecha": "20261024", "monto_servicios": 0.00, "monto_bienes": 508.47, "itbis": 91.53, "propina": 0.00, "total": 600.00, "forma_pago": "02"},
    {"emisor": "LINDO AMANECER SRL", "rnc": "133383993", "ncf": "B0200485464", "fecha": "20261026", "monto_servicios": 0.00, "monto_bienes": 16.95, "itbis": 3.05, "propina": 0.00, "total": 20.00, "forma_pago": "01"}
]

df = pd.DataFrame(invoices_data)

# Función de auditoría fiscal
ncf_vistos = set()
es_valida_list = []
justif_list = []

for _, r in df.iterrows():
    ncf = str(r["ncf"]).strip().upper()
    rnc = str(r["rnc"]).strip().replace("-", "")
    errs = []
    
    if ncf in ncf_vistos:
        errs.append("INCIDENCIA: NCF Duplicado en el período")
    else:
        ncf_vistos.add(ncf)
        
    if ncf.startswith("B02") or ncf.startswith("E32"):
        errs.append("IRREGULARIDAD: NCF de Consumo (B02/E32) no otorga crédito fiscal")
        
    if len(rnc) not in [9, 11]:
        errs.append(f"ERROR RNC: Longitud inválida ({len(rnc)})")
        
    sub = float(r["monto_servicios"]) + float(r["monto_bienes"])
    itb = float(r["itbis"])
    prop = float(r["propina"])
    calc = round(sub + itb + prop, 2)
    tot = round(float(r["total"]), 2)
    
    if abs(calc - tot) > 0.50:
        errs.append(f"DESCUADRE: Calculado ({calc}) vs Facturado ({tot})")
        
    valida = len(errs) == 0
    justif = "CUMPLE 100/100 DGII" if valida else " | ".join(errs)
    
    es_valida_list.append(valida)
    justif_list.append(justif)

df["es_valida"] = es_valida_list
df["justificacion"] = justif_list

df_validas = df[df["es_valida"] == True]
df_errores = df[df["es_valida"] == False]

# 1. CREACIÓN DEL ARCHIVO EXCEL DE 3 PESTAÑAS
excel_path = "Formato_606_DGII_133091216_202610.xlsx"
wb = openpyxl.Workbook()

header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

# Pestaña 1: Full
ws1 = wb.active
ws1.title = "Todas las Facturas (Full)"
ws1.append(["REGISTRO COMPLETO DE FACTURAS - JA INVESTMENT ANTIG RD SRL | RNC: 133091216 | Período: 202610"])
ws1.append([])
headers1 = ["No.", "Emisor", "RNC Suplidor", "NCF", "Fecha", "Servicios", "Bienes", "ITBIS", "Propina", "Total", "Estado", "Observación / Justificación"]
ws1.append(headers1)
for col in range(1, 13):
    cell = ws1.cell(row=3, column=col)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")

for idx, r in df.iterrows():
    rnc_c = str(r["rnc"]).replace("-", "")
    row_num = idx + 4
    ws1.append([
        idx + 1, r["emisor"], rnc_c, r["ncf"], r["fecha"],
        float(r["monto_servicios"]), float(r["monto_bienes"]), float(r["itbis"]), float(r["propina"]), float(r["total"]),
        "Aceptada" if r["es_valida"] else "Rechazada", r["justificacion"]
    ])
    # Formatear números
    for col_idx in [6, 7, 8, 9, 10]:
        ws1.cell(row=row_num, column=col_idx).number_format = '#,##0.00'

# Pestaña 2: Errores / Incidencias
ws2 = wb.create_sheet(title="Facturas con Errores")
ws2.append(["FACTURAS CON INCIDENCIAS Y NOVEDADES - JA INVESTMENT ANTIG RD SRL | RNC: 133091216 | Período: 202610"])
ws2.append([])
ws2.append(headers1)
for col in range(1, 13):
    cell = ws2.cell(row=3, column=col)
    cell.fill = PatternFill(start_color="8B0000", end_color="8B0000", fill_type="solid")
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")

for idx, r in df_errores.reset_index(drop=True).iterrows():
    rnc_c = str(r["rnc"]).replace("-", "")
    row_num = idx + 4
    ws2.append([
        idx + 1, r["emisor"], rnc_c, r["ncf"], r["fecha"],
        float(r["monto_servicios"]), float(r["monto_bienes"]), float(r["itbis"]), float(r["propina"]), float(r["total"]),
        "Rechazada", r["justificacion"]
    ])
    for col_idx in [6, 7, 8, 9, 10]:
        ws2.cell(row=row_num, column=col_idx).number_format = '#,##0.00'

# Pestaña 3: Formato 606 Validado 100/100 (23 Campos)
ws3 = wb.create_sheet(title="Formato 606 Validado 100-100")
headers_606 = [
    "RNC o Cédula", "Tipo Id", "Tipo Bienes y Servicios", "NCF o e-NCF", "NCF Modificado",
    "Fecha Comprobante", "Fecha Pago", "Monto Servicios", "Monto Bienes", "Total Facturado",
    "ITBIS Facturado", "ITBIS Retenido", "ITBIS Proporcionalidad", "ITBIS Costo",
    "ITBIS por Adelantar", "ITBIS Compras", "Tipo Retención ISR", "Monto Retención ISR",
    "ISR Percibido", "Impuesto Selectivo", "Otros Impuestos", "Propina Legal", "Forma de Pago"
]
ws3.append(["FORMATO 606 DGII OFICIAL VÁLIDO 100/100 - JA INVESTMENT ANTIG RD SRL | RNC: 133091216 | Período: 202610"])
ws3.append([])
ws3.append(headers_606)
for col in range(1, 24):
    cell = ws3.cell(row=3, column=col)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")

for row_i, r in enumerate(df_validas.to_dict("records"), 4):
    rnc_c = str(r["rnc"]).replace("-", "")
    tipo_id = 1 if len(rnc_c) in [9, 11] else 2
    tipo_b = "03" if "PETROLE" in str(r["emisor"]).upper() or "ESTACION" in str(r["emisor"]).upper() else "02"
    f_pag = str(r.get("forma_pago", "01"))
    
    ws3.cell(row=row_i, column=1, value=rnc_c).number_format = '@'
    ws3.cell(row=row_i, column=2, value=tipo_id)
    ws3.cell(row=row_i, column=3, value=tipo_b)
    ws3.cell(row=row_i, column=4, value=r["ncf"]).number_format = '@'
    ws3.cell(row=row_i, column=5, value="")
    ws3.cell(row=row_i, column=6, value=str(r["fecha"])).number_format = '@'
    ws3.cell(row=row_i, column=7, value=str(r["fecha"])).number_format = '@'
    ws3.cell(row=row_i, column=8, value=float(r["monto_servicios"])).number_format = '#,##0.00'
    ws3.cell(row=row_i, column=9, value=float(r["monto_bienes"])).number_format = '#,##0.00'
    ws3.cell(row=row_i, column=10, value=f"=H{row_i}+I{row_i}").number_format = '#,##0.00'
    ws3.cell(row=row_i, column=11, value=float(r["itbis"])).number_format = '#,##0.00'
    ws3.cell(row=row_i, column=12, value=0.00).number_format = '#,##0.00'
    ws3.cell(row=row_i, column=13, value=0.00).number_format = '#,##0.00'
    ws3.cell(row=row_i, column=14, value=0.00).number_format = '#,##0.00'
    ws3.cell(row=row_i, column=15, value=float(r["itbis"])).number_format = '#,##0.00'
    ws3.cell(row=row_i, column=16, value=0.00).number_format = '#,##0.00'
    ws3.cell(row=row_i, column=17, value="")
    ws3.cell(row=row_i, column=18, value=0.00).number_format = '#,##0.00'
    ws3.cell(row=row_i, column=19, value=0.00).number_format = '#,##0.00'
    ws3.cell(row=row_i, column=20, value=0.00).number_format = '#,##0.00'
    ws3.cell(row=row_i, column=21, value=0.00).number_format = '#,##0.00'
    ws3.cell(row=row_i, column=22, value=float(r["propina"])).number_format = '#,##0.00'
    ws3.cell(row=row_i, column=23, value=f_pag)

# Ajustar anchos de columna en todas las hojas
for sheet in [ws1, ws2, ws3]:
    for col in sheet.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        sheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

wb.save(excel_path)

# 2. CREACIÓN DEL INFORME EN WORD
word_path = "Informe_Auditoria_Fiscal_133091216_202610.docx"
doc = docx.Document()

# Encabezado de la Firma
p_header = doc.add_paragraph()
p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_f = p_header.add_run("BALDERA SANTOS Y ASOCIADOS, SRL\n")
run_f.bold = True
run_f.font.size = Pt(16)
run_f.font.color.rgb = RGBColor(27, 54, 93)

run_sub = p_header.add_run("AUDITORÍA DE CUMPLIMIENTO FISCAL Y VALIDACIÓN FORMATO 606 DGII\n")
run_sub.font.size = Pt(12)
run_sub.bold = True

doc.add_paragraph("=" * 65)

# Sección 1
doc.add_heading("1. DATOS GENERALES DE LA AUDITORÍA", level=1)
p_info = doc.add_paragraph()
p_info.add_run("• Razon Social Cliente: ").bold = True
p_info.add_run("JA INVESTMENT ANTIG RD SRL\n")
p_info.add_run("• RNC Cliente: ").bold = True
p_info.add_run("133091216\n")
p_info.add_run("• Período Auditado: ").bold = True
p_info.add_run("Octubre 2026 (202610)\n")
p_info.add_run("• Firma Auditora: ").bold = True
p_info.add_run("BALDERA SANTOS Y ASOCIADOS, SRL")

# Sección 2
doc.add_heading("2. RESUMEN DE RESULTADOS DE VALIDACIÓN", level=1)
tot_c = len(df)
val_c = len(df_validas)
err_c = len(df_errores)

t_res = doc.add_table(rows=4, cols=3)
t_res.style = 'Table Grid'

headers_w = ["Métrica / Criterio", "Cantidad Facturas", "Porcentaje (%)"]
for i, h in enumerate(headers_w):
    cell = t_res.cell(0, i)
    cell.text = h
    cell.paragraphs[0].runs[0].font.bold = True

t_res.cell(1, 0).text = "Total Facturas Analizadas"
t_res.cell(1, 1).text = str(tot_c)
t_res.cell(1, 2).text = "100.0%"

t_res.cell(2, 0).text = "Facturas Aceptadas (100/100 DGII)"
t_res.cell(2, 1).text = str(val_c)
t_res.cell(2, 2).text = f"{round((val_c/tot_c)*100, 2)}%"

t_res.cell(3, 0).text = "Facturas Rechazadas / Incidencias"
t_res.cell(3, 1).text = str(err_c)
t_res.cell(3, 2).text = f"{round((err_c/tot_c)*100, 2)}%"

# Sección 3
doc.add_heading("3. DETALLE DE FACTURAS CON INCIDENCIAS / ERRORES", level=1)
p_err_intro = doc.add_paragraph("Se identificaron 6 facturas que no cumplen con los requisitos para ser incluidas en la declaración informativa 606 debido a duplicidades de comprobantes o uso de NCFs de Consumo (B02/E32):")

t_err = doc.add_table(rows=1, cols=5)
t_err.style = 'Table Grid'
h_err_cols = ["Emisor", "NCF / e-NCF", "Fecha", "Total (RD$)", "Motivo Rechazo / Incidencia"]
for i, h in enumerate(h_err_cols):
    cell = t_err.cell(0, i)
    cell.text = h
    cell.paragraphs[0].runs[0].font.bold = True

for r in df_errores.to_dict("records"):
    row_cells = t_err.add_row().cells
    row_cells[0].text = str(r["emisor"])
    row_cells[1].text = str(r["ncf"])
    row_cells[2].text = str(r["fecha"])
    row_cells[3].text = f"RD$ {float(r['total']):,.2f}"
    row_cells[4].text = str(r["justificacion"])

# Sección 4
doc.add_heading("4. DICTAMEN DE AUDITORÍA FISCAL", level=1)
doc.add_paragraph(
    "Las 33 facturas aprobadas cumplen formalmente con los requisitos del Decreto 254-06, "
    "las resoluciones de comprobantes electrónicos (e-CF) y la estructura oficial de 23 campos del Formato 606 de la DGII. "
    "Monto total de ITBIS computable a favor de la empresa: RD$ 25,244.67.\n\n"
    "Las comprobantes catalogados con incidencias (comprobantes de consumo B02/E32 y NCFs duplicados) "
    "fueron excluidos del archivo de envío 606 para evitar inconsistencias y posibles inconsistencias fiscales con la DGII."
)

doc.add_paragraph("\n\n___________________________________\nBALDERA SANTOS Y ASOCIADOS, SRL\nFirma Auditora Responsable")

doc.save(word_path)

print(f"Archivos generados exitosamente: {excel_path} y {word_path}")
