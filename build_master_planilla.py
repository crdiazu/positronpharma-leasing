import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
wb.remove(wb.active)

# Paleta Impeccable TECNOVOA (Editorial, sobria, sin saturación)
C_NAVY_DARK = "0F172A"      # Slate 900
C_NAVY_MED = "1E293B"       # Slate 800
C_SKY_DARK = "0369A1"       # Sky 700
C_SKY_LIGHT = "E0F2FE"      # Sky 100
C_EMERALD_DARK = "065F46"   # Emerald 800
C_EMERALD_LIGHT = "ECFDF5"  # Emerald 50
C_AMBER_LIGHT = "FEF3C7"    # Amber 100 (Entradas del usuario)
C_GRAY_LIGHT = "F8FAFC"     # Slate 50
C_BORDER = "CBD5E1"         # Slate 300
C_WHITE = "FFFFFF"

f_title = Font(name="Segoe UI", size=13, bold=True, color=C_WHITE)
f_subtitle = Font(name="Segoe UI", size=9, italic=True, color="94A3B8")
f_head = Font(name="Segoe UI", size=9, bold=True, color=C_WHITE)
f_data = Font(name="Segoe UI", size=9)
f_data_bold = Font(name="Segoe UI", size=9, bold=True)
f_param_lbl = Font(name="Segoe UI", size=9, bold=True, color="1E293B")
f_param_val = Font(name="Segoe UI", size=10, bold=True, color="0369A1")
f_tot = Font(name="Segoe UI", size=10, bold=True, color=C_WHITE)
f_note = Font(name="Segoe UI", size=8, italic=True, color="64748B")

fill_title = PatternFill(start_color=C_NAVY_DARK, fill_type="solid")
fill_head = PatternFill(start_color=C_NAVY_MED, fill_type="solid")
fill_param_box = PatternFill(start_color=C_AMBER_LIGHT, fill_type="solid")
fill_sale = PatternFill(start_color=C_EMERALD_LIGHT, fill_type="solid")
fill_zebra = PatternFill(start_color=C_GRAY_LIGHT, fill_type="solid")
fill_white = PatternFill(start_color=C_WHITE, fill_type="solid")
fill_tot = PatternFill(start_color=C_NAVY_DARK, fill_type="solid")

thin = Side(border_style="thin", color=C_BORDER)
border_cell = Border(left=thin, right=thin, top=thin, bottom=thin)
border_tot = Border(top=Side(border_style="medium", color=C_NAVY_DARK), bottom=Side(border_style="double", color=C_NAVY_DARK))

# ==========================================
# HOJA 1: CALCULADORA DINÁMICA DE LEASING
# ==========================================
ws = wb.create_sheet(title="Calculadora Dinámica")
ws.views.sheetView[0].showGridLines = True

# Banner
ws.merge_cells("A1:O1")
ws["A1"] = "TECNOVOA S.A. | MODELO FINANCIERO DINÁMICO & SIMULADOR DE LEASING 36 MESES"
ws["A1"].font = f_title
ws["A1"].fill = fill_title
ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws.row_dimensions[1].height = 32

ws.merge_cells("A2:O2")
ws["A2"] = "Cliente: Positronpharma S.A. | Rut: 76.089.296-3 | Renovación Parque TI (53 Dispositivos + Accesorios)"
ws["A2"].font = f_subtitle
ws["A2"].fill = fill_title
ws["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws.row_dimensions[2].height = 18

# Bloque de Parámetros Globales (Filas 4 y 5)
ws["B4"] = "Tasa de Cambio (CLP/USD):"
ws["B4"].font = f_param_lbl
ws["C4"] = 980
ws["C4"].font = f_param_val
ws["C4"].fill = fill_param_box
ws["C4"].border = border_cell
ws["C4"].number_format = "$#,##0"
ws["C4"].alignment = Alignment(horizontal="right")

ws["E4"] = "Margen Objetivo Venta (%):"
ws["E4"].font = f_param_lbl
ws["F4"] = 0.15
ws["F4"].font = f_param_val
ws["F4"].fill = fill_param_box
ws["F4"].border = border_cell
ws["F4"].number_format = "0.0%"
ws["F4"].alignment = Alignment(horizontal="right")

ws["H4"] = "Factor Mensual Leasing 36m:"
ws["H4"].font = f_param_lbl
ws["I4"] = 0.0298
ws["I4"].font = f_param_val
ws["I4"].fill = fill_param_box
ws["I4"].border = border_cell
ws["I4"].number_format = "0.0000"
ws["I4"].alignment = Alignment(horizontal="right")

ws["K4"] = "Tasa IVA (Chile):"
ws["K4"].font = f_param_lbl
ws["L4"] = 0.19
ws["L4"].font = f_param_val
ws["L4"].fill = fill_param_box
ws["L4"].border = border_cell
ws["L4"].number_format = "0.0%"
ws["L4"].alignment = Alignment(horizontal="right")

ws["B5"] = "* Celdas en amarillo son editables: modifique el Margen Objetivo (ej. 12% a 15%) o el Dólar para recalcular automáticamente toda la propuesta."
ws["B5"].font = f_note
ws.merge_cells("B5:L5")
ws.row_dimensions[4].height = 24
ws.row_dimensions[5].height = 18

# Encabezados de Tabla (Fila 7)
headers = [
    "Ítem / Caso de Uso", "Vendor", "Part Number (P/N)", "Descripción Técnica Detallada",
    "Cant.", "Costo Unit. (USD)", "Costo Unit. (CLP)", "Costo Total (CLP)",
    "Margen", "Precio Venta Unit. Neto (CLP)", "Total Venta Capex Neto (CLP)",
    "Cuota Mensual Unit. Neta (CLP)", "Cuota Mensual Total Neta (CLP)",
    "Cuota Total c/IVA (CLP)", "Garantía & Cobertura Oficial Vendor"
]

for col_idx, h in enumerate(headers, start=1):
    c = ws.cell(row=7, column=col_idx, value=h)
    c.font = f_head
    c.fill = fill_head
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = border_cell
ws.row_dimensions[7].height = 28

items_data = [
    ("Caso 1: Movilidad Administrativa", "Lenovo", "21SY002QCL", "Lenovo ThinkPad E14 Gen 7 (Core Ultra 7 155H, 16GB RAM, 512GB SSD Opal, 14 pulg IPS, Win11 Pro, MIL-STD)", 29, 1256.00, "3Y On-Site Premier + ADP (Líquidos/Caídas)"),
    ("Caso 2: Jefaturas y Movilidad Crítica", "Lenovo", "21QD001GCL", "Lenovo ThinkPad T14 Gen 6 (Core Ultra 7 255H, 32GB RAM, 1TB SSD, 14 pulg WUXGA, Chasis Magnesio)", 6, 1611.00, "3Y Premier Support 24/7 + ADP + Keep Your Drive"),
    ("Caso 3: Puestos Fijos y Laboratorio", "Lenovo", "13QH004HCS", "Lenovo ThinkCentre Neo 50a 24 G6 AIO (Core i5 13th, 16GB, 512GB SSD, 23.8 pulg Borderless + Kit Wireless)", 3, 970.25, "3Y On-Site NBD + Kit Inalámbrico"),
    ("Caso 4: Despacho Trazable Radiofármacos", "Zebra", "WCMTB-T27B6ABC2-A6", "Zebra TC27 Industrial Handheld 5G/Wi-Fi 6E (Qualcomm Octa-Core, SE4710 1D/2D, IP68/IP65, Gorilla Glass)", 8, 646.27, "Zebra OneCare Essential 3 Años Integral"),
    ("Caso 4: Accesorios Operativos Despacho", "Zebra", "SG-TC2L-STYLUS-03", "Kit Operativo Limpio (Stylus cautivo punta fina con tether retráctil + Correa de mano ergonómica)", 8, 75.00, "Garantía oficial Zebra"),
    ("Caso 5: Supervisión y Movilidad Planta", "Lenovo", "ZAFM0762CL", "Lenovo IdeaTab 11 pulg 2.5K 5G (Dimensity 7050, 8GB RAM, 256GB, Sim 5G + Folio Teclado + Lenovo Pen)", 5, 285.00, "2Y Oficial Lenovo + Teclado y Lápiz"),
    ("Caso 6A: Impresión Confidencial RRHH", "Brother", "MFC-L3760CDW", "Brother Multifuncional Láser Color LED (27 ppm, Dúplex, Red/Wi-Fi, Secure Print con liberación por PIN)", 1, 493.15, "2 Años Garantía Oficial On-Site"),
    ("Caso 6B: Impresión Volumen Heat-Free", "Epson", "C11CJ88303", "Epson EcoTank L6490 Tinta Pigmentada DURABrite (PrecisionCore Sin Calor, Dúplex, ADF 35h, 7500 págs negro)", 1, 356.03, "2 Años o 50.000 págs Garantía Oficial")
]

r = 8
for idx, itm in enumerate(items_data):
    caso, vendor, pn, desc, qty, usd, sla = itm
    
    ws.cell(row=r, column=1, value=caso).alignment = Alignment(horizontal="left", vertical="center")
    ws.cell(row=r, column=2, value=vendor).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=r, column=3, value=pn).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=r, column=4, value=desc).alignment = Alignment(horizontal="left", vertical="center")
    
    c_q = ws.cell(row=r, column=5, value=qty)
    c_q.alignment = Alignment(horizontal="center", vertical="center")
    c_q.number_format = "#,##0"
    
    c_usd = ws.cell(row=r, column=6, value=usd)
    c_usd.alignment = Alignment(horizontal="right", vertical="center")
    c_usd.number_format = "$#,##0.00"
    
    # Costo CLP Unit = Costo USD * $C$4 (TC)
    c_cu = ws.cell(row=r, column=7, value=f"=F{r}*$C$4")
    c_cu.alignment = Alignment(horizontal="right", vertical="center")
    c_cu.number_format = "$#,##0"
    
    # Costo Total CLP = Cantidad * Costo Unit CLP
    c_ct = ws.cell(row=r, column=8, value=f"=E{r}*G{r}")
    c_ct.alignment = Alignment(horizontal="right", vertical="center")
    c_ct.number_format = "$#,##0"
    
    # Margen = $F$4
    c_mg = ws.cell(row=r, column=9, value="=$F$4")
    c_mg.alignment = Alignment(horizontal="center", vertical="center")
    c_mg.number_format = "0.0%"
    
    # Precio Venta Unit Neto = Costo Unit CLP / (1 - Margen)
    c_pvu = ws.cell(row=r, column=10, value=f"=G{r}/(1-I{r})")
    c_pvu.alignment = Alignment(horizontal="right", vertical="center")
    c_pvu.number_format = "$#,##0"
    c_pvu.fill = fill_sale
    c_pvu.font = f_data_bold
    
    # Total Venta Capex = Cantidad * Precio Venta Unit
    c_totv = ws.cell(row=r, column=11, value=f"=E{r}*J{r}")
    c_totv.alignment = Alignment(horizontal="right", vertical="center")
    c_totv.number_format = "$#,##0"
    c_totv.fill = fill_sale
    c_totv.font = f_data_bold
    
    # Cuota Mensual Unit Neta = Precio Venta Unit * $I$4 (Factor Leasing)
    c_qmu = ws.cell(row=r, column=12, value=f"=J{r}*$I$4")
    c_qmu.alignment = Alignment(horizontal="right", vertical="center")
    c_qmu.number_format = "$#,##0"
    
    # Cuota Mensual Total Neta = Total Venta Capex * $I$4
    c_qmt = ws.cell(row=r, column=13, value=f"=K{r}*$I$4")
    c_qmt.alignment = Alignment(horizontal="right", vertical="center")
    c_qmt.number_format = "$#,##0"
    c_qmt.font = f_data_bold
    
    # Cuota Total c/IVA = Cuota Neta * (1 + $L$4)
    c_qm_iva = ws.cell(row=r, column=14, value=f"=M{r}*(1+$L$4)")
    c_qm_iva.alignment = Alignment(horizontal="right", vertical="center")
    c_qm_iva.number_format = "$#,##0"
    
    ws.cell(row=r, column=15, value=sla).alignment = Alignment(horizontal="left", vertical="center")
    
    rf = fill_zebra if idx % 2 == 1 else fill_white
    for col in range(1, 16):
        cell_item = ws.cell(row=r, column=col)
        cell_item.border = border_cell
        if col not in [10, 11] and rf != fill_white:
            cell_item.fill = rf
            
    ws.row_dimensions[r].height = 22
    r += 1

tot_r = r
ws.merge_cells(start_row=tot_r, start_column=1, end_row=tot_r, end_column=4)
ws.cell(row=tot_r, column=1, value="TOTAL CONSOLIDADO PARQUE TI (53 EQUIPOS + ACCESORIOS)").font = f_tot

for col in range(1, 5):
    ws.cell(row=tot_r, column=col).fill = fill_tot
    ws.cell(row=tot_r, column=col).border = border_tot

# Total Cantidad
c_tq = ws.cell(row=tot_r, column=5, value=f"=SUM(E8:E{tot_r-1})")
c_tq.font = f_tot
c_tq.fill = fill_tot
c_tq.alignment = Alignment(horizontal="center", vertical="center")
c_tq.number_format = "#,##0"
c_tq.border = border_tot

for col in [6, 7]:
    c_b = ws.cell(row=tot_r, column=col, value="")
    c_b.fill = fill_tot
    c_b.border = border_tot

# Total Costo CLP
c_tc = ws.cell(row=tot_r, column=8, value=f"=SUM(H8:H{tot_r-1})")
c_tc.font = f_tot
c_tc.fill = fill_tot
c_tc.alignment = Alignment(horizontal="right", vertical="center")
c_tc.number_format = "$#,##0"
c_tc.border = border_tot

# Margen Ponderado Real = (Total Venta - Total Costo) / Total Venta
c_tmg = ws.cell(row=tot_r, column=9, value=f"=(K{tot_r}-H{tot_r})/K{tot_r}")
c_tmg.font = f_tot
c_tmg.fill = fill_tot
c_tmg.alignment = Alignment(horizontal="center", vertical="center")
c_tmg.number_format = "0.0%"
c_tmg.border = border_tot

c_b10 = ws.cell(row=tot_r, column=10, value="")
c_b10.fill = fill_tot
c_b10.border = border_tot

# Total Venta Capex
c_tv = ws.cell(row=tot_r, column=11, value=f"=SUM(K8:K{tot_r-1})")
c_tv.font = f_tot
c_tv.fill = fill_tot
c_tv.alignment = Alignment(horizontal="right", vertical="center")
c_tv.number_format = "$#,##0"
c_tv.border = border_tot

c_b12 = ws.cell(row=tot_r, column=12, value="")
c_b12.fill = fill_tot
c_b12.border = border_tot

# Total Cuota Neta
c_tqn = ws.cell(row=tot_r, column=13, value=f"=SUM(M8:M{tot_r-1})")
c_tqn.font = f_tot
c_tqn.fill = fill_tot
c_tqn.alignment = Alignment(horizontal="right", vertical="center")
c_tqn.number_format = "$#,##0"
c_tqn.border = border_tot

# Total Cuota c/IVA
c_tqi = ws.cell(row=tot_r, column=14, value=f"=SUM(N8:N{tot_r-1})")
c_tqi.font = f_tot
c_tqi.fill = fill_tot
c_tqi.alignment = Alignment(horizontal="right", vertical="center")
c_tqi.number_format = "$#,##0"
c_tqi.border = border_tot

c_b15 = ws.cell(row=tot_r, column=15, value="")
c_b15.fill = fill_tot
c_b15.border = border_tot

ws.row_dimensions[tot_r].height = 26

# Cuadro Resumen Ejecutivo (Kpis)
kpi_r = tot_r + 3
ws.merge_cells(start_row=kpi_r, start_column=2, end_row=kpi_r, end_column=5)
ws.cell(row=kpi_r, column=2, value="RESUMEN EJECUTIVO DEL CONTRATO LEASING (36 MESES)").font = f_tot
ws.cell(row=kpi_r, column=2).fill = fill_head

kpi_data = [
    ("Costo Total de Compra Hardware (Neto):", f"=H{tot_r}", "$#,##0"),
    ("Valor Venta Capex Financiado (Neto):", f"=K{tot_r}", "$#,##0"),
    ("Margen Comercial Bruto TECNOVOA (Neto):", f"=K{tot_r}-H{tot_r}", "$#,##0"),
    ("Porcentaje de Margen Efectivo:", f"=I{tot_r}", "0.00%"),
    ("Cuota Mensual de Leasing Operativo (Neto):", f"=M{tot_r}", "$#,##0"),
    ("IVA Mensual (19% Crédito Fiscal):", f"=M{tot_r}*$L$4", "$#,##0"),
    ("Cuota Mensual Facturada Total (c/IVA):", f"=N{tot_r}", "$#,##0"),
    ("Ahorro Impositivo Estimado (Tasa Corp. 27%):", f"=(M{tot_r}*36)*0.27", "$#,##0")
]

for idx_k, (k_lbl, k_val, k_fmt) in enumerate(kpi_data):
    curr_k = kpi_r + 1 + idx_k
    ws.cell(row=curr_k, column=2, value=k_lbl).font = f_param_lbl
    ws.merge_cells(start_row=curr_k, start_column=2, end_row=curr_k, end_column=4)
    c_val = ws.cell(row=curr_k, column=5, value=k_val)
    c_val.font = f_data_bold
    c_val.number_format = k_fmt
    c_val.alignment = Alignment(horizontal="right")
    if "Margen Comercial" in k_lbl or "Cuota Mensual de Leasing" in k_lbl:
        c_val.fill = fill_sale
        c_val.font = Font(name="Segoe UI", size=10, bold=True, color="047857")

col_w = {1: 30, 2: 12, 3: 20, 4: 48, 5: 8, 6: 15, 7: 16, 8: 18, 9: 10, 10: 18, 11: 20, 12: 18, 13: 20, 14: 20, 15: 36}
for c_i, w in col_w.items():
    ws.column_dimensions[get_column_letter(c_i)].width = w

# ==========================================
# HOJA 2: COMPARADOR DE SENSIBILIDAD 12% - 15%
# ==========================================
ws2 = wb.create_sheet(title="Sensibilidad 12% vs 15%")
ws2.views.sheetView[0].showGridLines = True

ws2.merge_cells("A1:L1")
ws2["A1"] = "TECNOVOA S.A. | MATRIZ DE SENSIBILIDAD: COMPARATIVA DE MARGEN 12% VS 15%"
ws2["A1"].font = f_title
ws2["A1"].fill = fill_title
ws2["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws2.row_dimensions[1].height = 30

h2 = [
    "Ítem / Caso de Uso", "Vendor", "Part Number", "Cant.", "Costo Total Compra (CLP)",
    "Venta Capex (12%)", "Cuota 36m (12%)", "Margen $ (12%)",
    "Venta Capex (15%)", "Cuota 36m (15%)", "Margen $ (15%)", "Delta Utilidad (15% vs 12%)"
]

for col_idx, h in enumerate(h2, start=1):
    c = ws2.cell(row=3, column=col_idx, value=h)
    c.font = f_head
    c.fill = PatternFill(start_color=C_EMERALD_DARK if col_idx in [9,10,11,12] else C_NAVY_MED, fill_type="solid")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = border_cell
ws2.row_dimensions[3].height = 28

r2 = 4
for idx, itm in enumerate(items_data):
    caso, vendor, pn, _, qty, usd, _ = itm
    ws2.cell(row=r2, column=1, value=caso).alignment = Alignment(horizontal="left", vertical="center")
    ws2.cell(row=r2, column=2, value=vendor).alignment = Alignment(horizontal="center", vertical="center")
    ws2.cell(row=r2, column=3, value=pn).alignment = Alignment(horizontal="center", vertical="center")
    
    cq = ws2.cell(row=r2, column=4, value=qty)
    cq.alignment = Alignment(horizontal="center", vertical="center")
    cq.number_format = "#,##0"
    
    # Costo vinculado a Hoja 1
    c_cost = ws2.cell(row=r2, column=5, value=f"='Calculadora Dinámica'!H{idx+8}")
    c_cost.alignment = Alignment(horizontal="right", vertical="center")
    c_cost.number_format = "$#,##0"
    
    # 12% Margen
    c_v12 = ws2.cell(row=r2, column=6, value=f"=E{r2}/(1-0.12)")
    c_v12.alignment = Alignment(horizontal="right", vertical="center")
    c_v12.number_format = "$#,##0"
    
    c_q12 = ws2.cell(row=r2, column=7, value=f"=F{r2}*'Calculadora Dinámica'!$I$4")
    c_q12.alignment = Alignment(horizontal="right", vertical="center")
    c_q12.number_format = "$#,##0"
    
    c_m12 = ws2.cell(row=r2, column=8, value=f"=F{r2}-E{r2}")
    c_m12.alignment = Alignment(horizontal="right", vertical="center")
    c_m12.number_format = "$#,##0"
    
    # 15% Margen
    c_v15 = ws2.cell(row=r2, column=9, value=f"=E{r2}/(1-0.15)")
    c_v15.alignment = Alignment(horizontal="right", vertical="center")
    c_v15.number_format = "$#,##0"
    c_v15.font = f_data_bold
    
    c_q15 = ws2.cell(row=r2, column=10, value=f"=I{r2}*'Calculadora Dinámica'!$I$4")
    c_q15.alignment = Alignment(horizontal="right", vertical="center")
    c_q15.number_format = "$#,##0"
    c_q15.font = f_data_bold
    
    c_m15 = ws2.cell(row=r2, column=11, value=f"=I{r2}-E{r2}")
    c_m15.alignment = Alignment(horizontal="right", vertical="center")
    c_m15.number_format = "$#,##0"
    c_m15.font = f_data_bold
    
    # Delta
    c_del = ws2.cell(row=r2, column=12, value=f"=K{r2}-H{r2}")
    c_del.alignment = Alignment(horizontal="right", vertical="center")
    c_del.number_format = "$#,##0"
    c_del.font = Font(name="Segoe UI", size=9, bold=True, color="047857")
    
    rf2 = fill_zebra if idx % 2 == 1 else fill_white
    for col in range(1, 13):
        c_i = ws2.cell(row=r2, column=col)
        c_i.border = border_cell
        if col < 9 and rf2 != fill_white:
            c_i.fill = rf2
    ws2.row_dimensions[r2].height = 22
    r2 += 1

tot_r2 = r2
ws2.merge_cells(start_row=tot_r2, start_column=1, end_row=tot_r2, end_column=3)
ws2.cell(row=tot_r2, column=1, value="TOTAL CONSOLIDADO").font = f_tot

for col in range(1, 4):
    ws2.cell(row=tot_r2, column=col).fill = fill_tot
    ws2.cell(row=tot_r2, column=col).border = border_tot

c_tq2 = ws2.cell(row=tot_r2, column=4, value=f"=SUM(D4:D{tot_r2-1})")
c_tq2.font = f_tot
c_tq2.fill = fill_tot
c_tq2.alignment = Alignment(horizontal="center")
c_tq2.number_format = "#,##0"
c_tq2.border = border_tot

for c_i in range(5, 13):
    col_let = get_column_letter(c_i)
    c_t = ws2.cell(row=tot_r2, column=c_i, value=f"=SUM({col_let}4:{col_let}{tot_r2-1})")
    c_t.font = f_tot
    c_t.fill = fill_tot
    c_t.alignment = Alignment(horizontal="right")
    c_t.number_format = "$#,##0"
    c_t.border = border_tot

ws2.row_dimensions[tot_r2].height = 26

col_w2 = {1: 30, 2: 12, 3: 20, 4: 8, 5: 20, 6: 18, 7: 18, 8: 16, 9: 18, 10: 18, 11: 16, 12: 18}
for c_i, w in col_w2.items():
    ws2.column_dimensions[get_column_letter(c_i)].width = w

out_file = "Clientes/PositronPharma/presentacion-positronpharma/Planilla_Maestra_Leasing_Positronpharma.xlsx"
wb.save(out_file)
print("Planilla guardada con éxito en:", out_file)
