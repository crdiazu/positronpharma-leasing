import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
# Remove default sheet
wb.remove(wb.active)

# Color Palette TECNOVOA Corporate
COLOR_NAVY = "0A192F"
COLOR_NAVY_LIGHT = "1E293B"
COLOR_BLUE_ACCENT = "0284C7"
COLOR_GRAY_LIGHT = "F8FAFC"
COLOR_GRAY_BORDER = "CBD5E1"
COLOR_WHITE = "FFFFFF"
COLOR_GREEN = "059669"
COLOR_GREEN_LIGHT = "ECFDF5"
COLOR_AMBER = "D97706"
COLOR_AMBER_LIGHT = "FEF3C7"

font_title = Font(name="Segoe UI", size=15, bold=True, color=COLOR_WHITE)
font_subtitle = Font(name="Segoe UI", size=10, italic=True, color="94A3B8")
font_section = Font(name="Segoe UI", size=11, bold=True, color=COLOR_WHITE)
font_header = Font(name="Segoe UI", size=9, bold=True, color=COLOR_WHITE)
font_data = Font(name="Segoe UI", size=9)
font_data_bold = Font(name="Segoe UI", size=9, bold=True)
font_total = Font(name="Segoe UI", size=10, bold=True, color=COLOR_WHITE)
font_note = Font(name="Segoe UI", size=8, italic=True, color="64748B")

fill_title = PatternFill(start_color=COLOR_NAVY, end_color=COLOR_NAVY, fill_type="solid")
fill_header = PatternFill(start_color=COLOR_NAVY_LIGHT, end_color=COLOR_NAVY_LIGHT, fill_type="solid")
fill_header_green = PatternFill(start_color="065F46", end_color="065F46", fill_type="solid")
fill_zebra = PatternFill(start_color=COLOR_GRAY_LIGHT, end_color=COLOR_GRAY_LIGHT, fill_type="solid")
fill_white = PatternFill(start_color=COLOR_WHITE, end_color=COLOR_WHITE, fill_type="solid")
fill_total = PatternFill(start_color=COLOR_NAVY, end_color=COLOR_NAVY, fill_type="solid")
fill_sale = PatternFill(start_color=COLOR_GREEN_LIGHT, end_color=COLOR_GREEN_LIGHT, fill_type="solid")

thin_side = Side(border_style="thin", color=COLOR_GRAY_BORDER)
border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
border_total = Border(top=Side(border_style="medium", color=COLOR_NAVY), bottom=Side(border_style="double", color=COLOR_NAVY))

USD_CLP = 980.0
FACTOR_LEASING_36M = 0.0298 # Factor financiero mensual estándar a 36 meses

# Structure of items
# [Linea, Item_Case, Marca, PN, Descripcion, Qty, Costo_USD, Costo_CLP_calc, Margen_15, Margen_12, SLA_Garantia]
items = [
    # 1. Movilidad Estandar
    ["01. Movilidad Operativa Estándar", "Caso 1: Movilidad Administrativa", "Lenovo", "21SY002QCL", 
     "Lenovo ThinkPad E14 Gen 7 (Ultra 7 155H, 16GB, 512GB SSD, 14\" WUXGA IPS, Win11 Pro, MIL-STD-810H)", 
     29, 1256.00, 0, 0.15, 0.12, "3Y On-Site Premier + ADP (Protección Caídas/Líquidos)"],
    
    # 2. Movilidad Rendimiento
    ["02. Movilidad Alto Rendimiento", "Caso 2: Jefaturas y Movilidad Crítica", "Lenovo", "21QD001GCL", 
     "Lenovo ThinkPad T14 Gen 6 (Ultra 7 255H, 32GB RAM, 1TB SSD, 14\" WUXGA, Win11 Pro, Chasis Magnesio)", 
     6, 1611.00, 0, 0.15, 0.12, "3Y Premier Support 24/7 + ADP + Keep Your Drive"],
    
    # 3. Puestos Fijos
    ["03. Puestos Fijos y Recepción", "Caso 3: Puestos Fijos y Laboratorio", "Lenovo", "13QH004HCS", 
     "Lenovo ThinkCentre Neo 50a 24 Gen 6 AIO (Core i5 13th, 16GB, 512GB SSD, Pantalla 23.8\" FHD Borderless)", 
     3, 970.25, 0, 0.15, 0.12, "3Y On-Site NBD + Teclado/Mouse inalámbrico"],
    
    # 4. Despacho Radiofármacos
    ["04. Despacho Logístico Trazable", "Caso 4: Trazabilidad Logística Radiofármacos", "Zebra", "WCMTB-T27B6ABC2-A6", 
     "Zebra TC27 Industrial Handheld (5G/Wi-Fi 6E, Android 13 Enterprise, Escáner SE4710 1D/2D, IP68/IP65)", 
     8, 646.27, 0, 0.15, 0.12, "Zebra OneCare Essential 3 Años (Cobertura Integral y Daño Accidental)"],
    
    # 4.1 Accesorios Zebra
    ["04. Despacho Logístico Trazable", "Caso 4: Accesorios Operación Limpia", "Zebra", "SG-TC2L-STYLUS-03", 
     "Bundle Operativo Trazabilidad (Stylus pasivo punta fina + Tether elástico + Correa de mano ergonométrica)", 
     8, 75.00, 0, 0.15, 0.12, "Garantía de reemplazo directo"],

    # 5. Supervisión y Planta
    ["05. Movilidad Planta y Calidad", "Caso 5: Control de Procesos y Supervisión", "Lenovo", "ZAFM0762CL", 
     "Lenovo IdeaTab 11\" 2.5K 5G (MediaTek Dimensity 7050, 8GB RAM, 256GB, Conectividad Celular 5G, Android 14)", 
     5, 285.00, 0, 0.15, 0.12, "2Y Garantía Oficial Lenovo + Funda folio con soporte"],

    # 6. Impresión RRHH
    ["06. Impresión Segura y Departamental", "Caso 6A: Impresión Confidencial RRHH", "Brother", "MFC-L3760CDW", 
     "Brother Multifuncional Láser Color LED (31 ppm, Dúplex, Red Gigabit/Wi-Fi, Impresión Segura por PIN)", 
     1, 493.15, 0, 0.15, 0.12, "2 Años Garantía Oficial Brother On-Site"],

    ["06. Impresión Segura y Departamental", "Caso 6B: Impresión Alto Rendimiento Costo/Pág", "Epson", "C11CJ88303", 
     "Epson EcoTank L6490 Multifuncional Tinta Pigmentada Heat-Free (Dúplex automático, Fax, Ethernet/Wi-Fi)", 
     1, 356.03, 0, 0.15, 0.12, "2 Años o 50.000 páginas Garantía Extendida Epson"]
]

# Sheet 1: Propuesta Principal (Margen 15%)
ws1 = wb.create_sheet(title="Planilla Maestra (15% Margen)")
ws1.views.sheetView[0].showGridLines = True

# Title block
ws1.merge_cells("A1:M1")
ws1["A1"] = "TECNOVOA S.A. | MODELO FINANCIERO Y MASTER PLANILLA DE HARDWARE - LEASING 36 MESES"
ws1["A1"].font = font_title
ws1["A1"].fill = fill_title
ws1["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws1.row_dimensions[1].height = 36

ws1.merge_cells("A2:M2")
ws1["A2"] = f"Cliente: Positronpharma S.A. | RUT: 76.089.296-3 | Parámetros: Margen 15% Venta | Tasa Cambio Ref: ${USD_CLP:,.0f} CLP/USD | Factor Leasing 36m: {FACTOR_LEASING_36M}"
ws1["A2"].font = font_subtitle
ws1["A2"].fill = fill_title
ws1["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws1.row_dimensions[2].height = 20

# Blank row 3
ws1.row_dimensions[3].height = 8

# Headers row 4
headers_1 = [
    "Línea de Solución", "Caso de Uso Operativo", "Marca Vendor", "Part Number (P/N)", "Descripción Técnica del Equipamiento", 
    "Cant.", "Costo Unit. (USD)", "Costo Unit. (CLP)", "Margen Venta", "Precio Venta Unit. (CLP)", 
    "Total Venta Capex (CLP)", "Cuota Mensual Unit. (CLP)", "Cuota Mensual Total (CLP)", "Garantía & Cobertura Vendor"
]

for col_idx, h in enumerate(headers_1, start=1):
    cell = ws1.cell(row=4, column=col_idx, value=h)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border_cell
ws1.row_dimensions[4].height = 28

current_row = 5
for idx, itm in enumerate(items):
    c_linea, c_caso, c_marca, c_pn, c_desc, c_qty, c_usd, _, c_m15, _, c_sla = itm
    
    ws1.cell(row=current_row, column=1, value=c_linea).alignment = Alignment(horizontal="left", vertical="center")
    ws1.cell(row=current_row, column=2, value=c_caso).alignment = Alignment(horizontal="left", vertical="center")
    ws1.cell(row=current_row, column=3, value=c_marca).alignment = Alignment(horizontal="center", vertical="center")
    ws1.cell(row=current_row, column=4, value=c_pn).alignment = Alignment(horizontal="center", vertical="center")
    ws1.cell(row=current_row, column=5, value=c_desc).alignment = Alignment(horizontal="left", vertical="center")
    
    c_cant = ws1.cell(row=current_row, column=6, value=c_qty)
    c_cant.alignment = Alignment(horizontal="center", vertical="center")
    c_cant.number_format = "#,##0"
    
    c_u_usd = ws1.cell(row=current_row, column=7, value=c_usd)
    c_u_usd.alignment = Alignment(horizontal="right", vertical="center")
    c_u_usd.number_format = "$#,##0.00"
    
    # Costo CLP = Costo USD * 980
    c_u_clp = ws1.cell(row=current_row, column=8, value=f"=G{current_row}*{USD_CLP}")
    c_u_clp.alignment = Alignment(horizontal="right", vertical="center")
    c_u_clp.number_format = "$#,##0"
    
    c_mg = ws1.cell(row=current_row, column=9, value=c_m15)
    c_mg.alignment = Alignment(horizontal="center", vertical="center")
    c_mg.number_format = "0.0%"
    
    # Precio Venta = Costo CLP / (1 - Margen)
    c_pv = ws1.cell(row=current_row, column=10, value=f"=H{current_row}/(1-I{current_row})")
    c_pv.alignment = Alignment(horizontal="right", vertical="center")
    c_pv.number_format = "$#,##0"
    c_pv.fill = fill_sale
    c_pv.font = font_data_bold
    
    # Total Venta = Cantidad * Precio Venta
    c_tot = ws1.cell(row=current_row, column=11, value=f"=F{current_row}*J{current_row}")
    c_tot.alignment = Alignment(horizontal="right", vertical="center")
    c_tot.number_format = "$#,##0"
    c_tot.fill = fill_sale
    c_tot.font = font_data_bold
    
    # Cuota Mensual Unit = Precio Venta * Factor Leasing
    c_cu = ws1.cell(row=current_row, column=12, value=f"=J{current_row}*{FACTOR_LEASING_36M}")
    c_cu.alignment = Alignment(horizontal="right", vertical="center")
    c_cu.number_format = "$#,##0"
    
    # Cuota Mensual Total = Total Venta * Factor Leasing
    c_ct = ws1.cell(row=current_row, column=13, value=f"=K{current_row}*{FACTOR_LEASING_36M}")
    c_ct.alignment = Alignment(horizontal="right", vertical="center")
    c_ct.number_format = "$#,##0"
    c_ct.font = font_data_bold
    
    c_gar = ws1.cell(row=current_row, column=14, value=c_sla)
    c_gar.alignment = Alignment(horizontal="left", vertical="center")
    
    # Row styling
    row_fill = fill_zebra if idx % 2 == 1 else fill_white
    for col in range(1, 15):
        cell_item = ws1.cell(row=current_row, column=col)
        if col not in [10, 11]:
            cell_item.fill = row_fill
            cell_item.font = font_data
        cell_item.border = border_cell
        
    ws1.row_dimensions[current_row].height = 24
    current_row += 1

# Total Row
tot_row = current_row
ws1.cell(row=tot_row, column=1, value="TOTAL PARQUE RENOVACIÓN TI (53 EQUIPOS + ACCESORIOS)").font = font_total
ws1.merge_cells(start_row=tot_row, start_column=1, end_row=tot_row, end_column=5)

for col in range(1, 6):
    ws1.cell(row=tot_row, column=col).fill = fill_total
    ws1.cell(row=tot_row, column=col).border = border_total

c_tot_qty = ws1.cell(row=tot_row, column=6, value=f"=SUM(F5:F{tot_row-1})")
c_tot_qty.font = font_total
c_tot_qty.fill = fill_total
c_tot_qty.alignment = Alignment(horizontal="center", vertical="center")
c_tot_qty.number_format = "#,##0"
c_tot_qty.border = border_total

for col in [7, 8, 9, 10]:
    c_bl = ws1.cell(row=tot_row, column=col, value="")
    c_bl.fill = fill_total
    c_bl.border = border_total

c_tot_capex = ws1.cell(row=tot_row, column=11, value=f"=SUM(K5:K{tot_row-1})")
c_tot_capex.font = font_total
c_tot_capex.fill = fill_total
c_tot_capex.alignment = Alignment(horizontal="right", vertical="center")
c_tot_capex.number_format = "$#,##0"
c_tot_capex.border = border_total

c_bl12 = ws1.cell(row=tot_row, column=12, value="")
c_bl12.fill = fill_total
c_bl12.border = border_total

c_tot_opex = ws1.cell(row=tot_row, column=13, value=f"=SUM(M5:M{tot_row-1})")
c_tot_opex.font = font_total
c_tot_opex.fill = fill_total
c_tot_opex.alignment = Alignment(horizontal="right", vertical="center")
c_tot_opex.number_format = "$#,##0"
c_tot_opex.border = border_total

c_bl14 = ws1.cell(row=tot_row, column=14, value="")
c_bl14.fill = fill_total
c_bl14.border = border_total

ws1.row_dimensions[tot_row].height = 26

# Notes below
note_row = tot_row + 2
ws1.cell(row=note_row, column=1, value="* Nota Financiera: Precios netos en pesos chilenos (CLP) no incluyen IVA. Factor mensual referencial leasing 36 meses: 0.0298.").font = font_note
ws1.cell(row=note_row+1, column=1, value="* Garantía Vendor Directa: ThinkPad y ThinkCentre respaldados por 3 años On-Site Premier con cobertura contra daño accidental (ADP). Handhelds Zebra respaldados por Zebra OneCare 3 años.").font = font_note

# Adjust column widths
col_widths = {1: 28, 2: 30, 3: 14, 4: 20, 5: 48, 6: 10, 7: 16, 8: 18, 9: 14, 10: 20, 11: 22, 12: 20, 13: 22, 14: 38}
for c_idx, w in col_widths.items():
    ws1.column_dimensions[get_column_letter(c_idx)].width = w

# Sheet 2: Comparativa de Escenarios 12% vs 15%
ws2 = wb.create_sheet(title="Comparativa Margen 12% vs 15%")
ws2.views.sheetView[0].showGridLines = True

ws2.merge_cells("A1:K1")
ws2["A1"] = "TECNOVOA S.A. | ANÁLISIS DE SENSIBILIDAD COMERCIAL: ESCENARIO 12% VS 15% MARGEN"
ws2["A1"].font = font_title
ws2["A1"].fill = fill_title
ws2["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws2.row_dimensions[1].height = 36

headers_2 = [
    "Línea de Solución", "Marca Vendor", "Part Number", "Cant.", "Costo Mayorista Total (CLP)",
    "P. Venta Tot. (12% Margen)", "Cuota Mensual 36m (12%)", "Margen Bruto (12%)",
    "P. Venta Tot. (15% Margen)", "Cuota Mensual 36m (15%)", "Margen Bruto (15%)"
]

for col_idx, h in enumerate(headers_2, start=1):
    cell = ws2.cell(row=3, column=col_idx, value=h)
    cell.font = font_header
    cell.fill = fill_header_green if col_idx >= 9 else fill_header
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border_cell
ws2.row_dimensions[3].height = 28

r2 = 4
for idx, itm in enumerate(items):
    c_linea, _, c_marca, c_pn, _, c_qty, c_usd, _, _, _, _ = itm
    
    ws2.cell(row=r2, column=1, value=c_linea).alignment = Alignment(horizontal="left", vertical="center")
    ws2.cell(row=r2, column=2, value=c_marca).alignment = Alignment(horizontal="center", vertical="center")
    ws2.cell(row=r2, column=3, value=c_pn).alignment = Alignment(horizontal="center", vertical="center")
    
    c_q = ws2.cell(row=r2, column=4, value=c_qty)
    c_q.alignment = Alignment(horizontal="center", vertical="center")
    c_q.number_format = "#,##0"
    
    # Costo Total CLP = Cant * (Costo USD * 980)
    c_c_tot = ws2.cell(row=r2, column=5, value=f"=D{r2}*({c_usd}*{USD_CLP})")
    c_c_tot.alignment = Alignment(horizontal="right", vertical="center")
    c_c_tot.number_format = "$#,##0"
    
    # 12% Escenario
    c_v12 = ws2.cell(row=r2, column=6, value=f"=E{r2}/(1-0.12)")
    c_v12.alignment = Alignment(horizontal="right", vertical="center")
    c_v12.number_format = "$#,##0"
    
    c_q12 = ws2.cell(row=r2, column=7, value=f"=F{r2}*{FACTOR_LEASING_36M}")
    c_q12.alignment = Alignment(horizontal="right", vertical="center")
    c_q12.number_format = "$#,##0"
    
    c_m12 = ws2.cell(row=r2, column=8, value=f"=F{r2}-E{r2}")
    c_m12.alignment = Alignment(horizontal="right", vertical="center")
    c_m12.number_format = "$#,##0"
    
    # 15% Escenario
    c_v15 = ws2.cell(row=r2, column=9, value=f"=E{r2}/(1-0.15)")
    c_v15.alignment = Alignment(horizontal="right", vertical="center")
    c_v15.number_format = "$#,##0"
    c_v15.font = font_data_bold
    
    c_q15 = ws2.cell(row=r2, column=10, value=f"=I{r2}*{FACTOR_LEASING_36M}")
    c_q15.alignment = Alignment(horizontal="right", vertical="center")
    c_q15.number_format = "$#,##0"
    c_q15.font = font_data_bold
    
    c_m15_tot = ws2.cell(row=r2, column=11, value=f"=I{r2}-E{r2}")
    c_m15_tot.alignment = Alignment(horizontal="right", vertical="center")
    c_m15_tot.number_format = "$#,##0"
    c_m15_tot.font = font_data_bold

    row_fill2 = fill_zebra if idx % 2 == 1 else fill_white
    for col in range(1, 12):
        cell_item = ws2.cell(row=r2, column=col)
        cell_item.border = border_cell
        if col < 9 and row_fill2 != fill_white:
            cell_item.fill = row_fill2
            
    ws2.row_dimensions[r2].height = 22
    r2 += 1

# Total Row Sheet 2
tot_r2 = r2
ws2.cell(row=tot_r2, column=1, value="TOTAL CONSOLIDADO").font = font_total
ws2.merge_cells(start_row=tot_r2, start_column=1, end_row=tot_r2, end_column=3)
for col in range(1, 4):
    ws2.cell(row=tot_r2, column=col).fill = fill_total
    ws2.cell(row=tot_r2, column=col).border = border_total

c_tot_q2 = ws2.cell(row=tot_r2, column=4, value=f"=SUM(D4:D{tot_r2-1})")
c_tot_q2.font = font_total
c_tot_q2.fill = fill_total
c_tot_q2.alignment = Alignment(horizontal="center", vertical="center")
c_tot_q2.number_format = "#,##0"
c_tot_q2.border = border_total

c_tot_cost = ws2.cell(row=tot_r2, column=5, value=f"=SUM(E4:E{tot_r2-1})")
c_tot_cost.font = font_total
c_tot_cost.fill = fill_total
c_tot_cost.alignment = Alignment(horizontal="right", vertical="center")
c_tot_cost.number_format = "$#,##0"
c_tot_cost.border = border_total

for c_idx in [6, 7, 8, 9, 10, 11]:
    col_let = get_column_letter(c_idx)
    c_tot = ws2.cell(row=tot_r2, column=c_idx, value=f"=SUM({col_let}4:{col_let}{tot_r2-1})")
    c_tot.font = font_total
    c_tot.fill = fill_total
    c_tot.alignment = Alignment(horizontal="right", vertical="center")
    c_tot.number_format = "$#,##0"
    c_tot.border = border_total

ws2.row_dimensions[tot_r2].height = 26

col_widths_2 = {1: 30, 2: 15, 3: 20, 4: 10, 5: 22, 6: 22, 7: 20, 8: 18, 9: 22, 10: 20, 11: 18}
for c_idx, w in col_widths_2.items():
    ws2.column_dimensions[get_column_letter(c_idx)].width = w

output_path = "Clientes/PositronPharma/presentacion-positronpharma/Planilla_Maestra_Leasing_Positronpharma.xlsx"
wb.save(output_path)
print(f"Excel guardado exitosamente en: {output_path}")
