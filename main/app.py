import streamlit as st
import pandas as pd
from datetime import datetime
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Lista de oficinas basada en tu reporte
OFICINAS = [
    "RIVIERA ARMENIA",
    "RIVIERA BUENAVISTA BARRANQUILLA",
    "RIVIERA UNICENTRO 1-040",
    "RIVIERA PALATINO",
    "RIVIERA SANTA ANA",
    "INEDITO PERFUME DE AUTOR",
    "RIVIERA EL RETIRO",
    "RIVIERA SANTA FE",
    "RIVIERA GRAN ESTACION",
    "RIVIERA UNICENTRO 1-252",
    "RIVIERA SALITRE",
    "LA HORA TALLER",
    "RIVIERA ANDINO",
    "RIVIERA UNICENTRO CALI",
    "RIVIERA CHIPICHAPE",
    "RIVIERA CENTENARIO",
    "RIVIERA CARTAGENA",
    "RIVIERA CARIBE PLAZA",
    "RIVIERA PLAZA BOCAGRANDE",
    "RIVIERA VENTURA PLAZA",
    "RIVIERA UNICENTRO CUCUTA",
    "RIVIERA IBAGUE",
    "RIVIERA ACQUA IBAGUE",
    "RIVIERA MANIZALES",
    "RIVIERA TESORO MEDELLIN",
    "RIVIERA UNICENTRO MEDELLIN",
    "RIVIERA SANTA FE MEDELLIN",
    "RIVIERA MONTERIA ALAMEDA",
    "RIVIERA SANTA LUCIA",
    "RIVIERA UNICENTRO PEREIRA",
    "RIVIERA POPAYAN",
    "RIVIERA SANTA MARTA",
    "RIVIERA VALLEDUPAR",
    "RIVIERA VILLAVICENCIO",
    "RIVIERA YOPAL"
]

EXCEL_FILE = "ESTADO_TIENDAS 2026.xlsx"

st.set_page_config(page_title="Control de Equipos Tecnológicos", layout="centered")

# --- PANEL DE ADMINISTRADOR EN LA BARRA LATERAL (Protección de descarga) ---
with st.sidebar:
    st.header("🔒 Panel de Administrador")
    password_ingresada = st.text_input("Contraseña de Descarga:", type="password")
    PASSWORD_SECRETA = "Rivier@25"

st.title("🖥️ Formulario de Control de Equipos Tecnológicos")
st.write("Selecciona la oficina e ingresa la información de las cajas correspondientes.")

# Paso 1: Selección de Oficina
oficina_seleccionada = st.selectbox("Seleccione la Oficina / Sucursal:", OFICINAS)

# Paso 2: Cantidad de cajas (máximo 5)
cantidad_cajas = st.number_input("¿Cantidad cajas en el punto? (Máximo 5):", min_value=1, max_value=5, value=1, step=1)

st.divider()

# Diccionario para almacenar los datos de las cajas de forma temporal
registros_cajas = []

# Paso 3: Preguntas dinámicas por cada caja
for i in range(1, int(cantidad_cajas) + 1):
    st.subheader(f"📦 Configuración de la Caja #{i}")
    
    col1, col2 = st.columns(2)
    with col1:
        pc_activo = st.text_input(f"Computador - Activo Fijo (Caja {i})", key=f"pc_act_{i}")
        pant_activo = st.text_input(f"Pantalla - Activo Fijo (Caja {i})", key=f"pant_act_{i}")
        mou_activo = st.text_input(f"Mouse - Marca (Caja {i})", key=f"mou_act_{i}")
        tec_activo = st.text_input(f"Teclado - Marca (Caja {i})", key=f"tec_act_{i}")
        lec_activo = st.text_input(f"Lector de Código - Activo Fijo (Caja {i})", key=f"lec_act_{i}")
        caj_activo = st.text_input(f"Cajón Monedero - Activo Fijo (Caja {i})", key=f"caj_act_{i}")
        tel_activo = st.text_input(f"Teléfono - Activo Fijo (Caja {i})", key=f"tel_act_{i}")
        bio_activo = st.text_input(f"Biométrico - Activo Fijo (Caja {i})", key=f"bio_act_{i}")
        imp_activo = st.text_input(f"Impresora POS - Activo Fijo (Caja {i})", key=f"imp_act_{i}")

    with col2:
        pc_estado = st.selectbox(f"Computador - Estado (Caja {i})", ["BUENO", "REGULAR", "MALO", "N/A"], key=f"pc_est_{i}")
        pant_estado = st.selectbox(f"Pantalla - Estado (Caja {i})", ["BUENO", "REGULAR", "MALO", "N/A"], key=f"pant_est_{i}")
        mou_estado = st.selectbox(f"Mouse - Estado (Caja {i})", ["BUENO", "REGULAR", "MALO", "N/A"], key=f"mou_est_{i}")
        tec_estado = st.selectbox(f"Teclado - Estado (Caja {i})", ["BUENO", "REGULAR", "MALO", "N/A"], key=f"tec_est_{i}")
        lec_estado = st.selectbox(f"Lector de Código - Estado (Caja {i})", ["BUENO", "REGULAR", "MALO", "N/A"], key=f"lec_est_{i}")
        caj_estado = st.selectbox(f"Cajón Monedero - Estado (Caja {i})", ["BUENO", "REGULAR", "MALO", "N/A"], key=f"caj_est_{i}")
        tel_estado = st.selectbox(f"Teléfono - Estado (Caja {i})", ["BUENO", "REGULAR", "MALO", "N/A"], key=f"tel_est_{i}")
        bio_estado = st.selectbox(f"Biométrico - Estado (Caja {i})", ["BUENO", "REGULAR", "MALO", "N/A"], key=f"bio_est_{i}")
        imp_estado = st.selectbox(f"Impresora POS - Estado (Caja {i})", ["BUENO", "REGULAR", "MALO", "N/A"], key=f"imp_est_{i}")

    # Guardar datos estructurados incluyendo Pantalla y todas las llaves unicas
    registros_cajas.append({
        "Fecha_Registro": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Oficina": oficina_seleccionada,
        "Numero_Caja": f"Caja {i}",
        "PC_Activo": pc_activo,
        "PC_Estado": pc_estado,
        "Pantalla_Activo": pant_activo,
        "Pantalla_Estado": pant_estado,
        "Mouse_Activo": mou_activo,
        "Mouse_Estado": mou_estado,
        "Teclado_Activo": tec_activo,
        "Teclado_Estado": tec_estado,
        "Lector_Activo": lec_activo,
        "Lector_Estado": lec_estado,
        "Cajon_Activo": caj_activo,
        "Cajon_Estado": caj_estado,
        "Telefono_Activo": tel_activo,
        "Telefono_Estado": tel_estado,
        "Biometrico_Activo": bio_activo,
        "Biometrico_Estado": bio_estado,
        "Impresora_Activa": imp_activo,
        "Impresora_Estado": imp_estado
    })
    st.divider()

# Botón para enviar y guardar en el Excel con formato institucional
if st.button("💾 Guardar Información del Formulario", type="primary"):
    df_nuevo = pd.DataFrame(registros_cajas)
    
    if os.path.exists(EXCEL_FILE):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["Control de Equipos"] if "Control de Equipos" in wb.sheetnames else wb.active
        
        for idx, row in df_nuevo.iterrows():
            ws.append([
                "", row["Oficina"], "",
                row["PC_Activo"], row["PC_Estado"],
                row["Telefono_Activo"], row["Telefono_Estado"],
                row["Biometrico_Activo"], row["Biometrico_Estado"],
                row["Lector_Activo"], row["Lector_Estado"],
                row["Cajon_Activo"], row["Cajon_Estado"],
                row["Impresora_Activa"], row["Impresora_Estado"]
            ])
        wb.save(EXCEL_FILE)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Control de Equipos"
        
        green_fill = PatternFill(start_color="046A38", end_color="046A38", fill_type="solid")
        light_green = PatternFill(start_color="C2E2D2", end_color="C2E2D2", fill_type="solid")
        sub_green = PatternFill(start_color="BFD7CB", end_color="BFD7CB", fill_type="solid")
        white_font = Font(name="Avenir", size=11, bold=True, color="FFFFFF")
        dark_green_font = Font(name="Avenir", size=10, bold=True, color="024021")
        gray_font = Font(name="Avenir", size=10, bold=True, color="444444")

        ws.merge_cells("A1:E3")
        ws["A1"] = "SISTEMA DE GESTIÓN ADMINISTRATIVA\nFORMATO DE CONTROL Y SEGUIMIENTO\nESTADO DE EQUIPOS TECNOLÓGICOS"
        ws["A1"].fill = green_fill
        ws["A1"].font = white_font
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        ws["F1"] = "FECHA DE EMISON"; ws["G1"] = datetime.now().strftime("%d/%m/%Y")
        ws["F2"] = "VERSIÓN:"; ws["G2"] = "1.0.0"
        ws["F3"] = "FECHA REVISIÓN:"; ws["G3"] = datetime.now().strftime("%d/%m/%Y")
        ws["H1"] = "ELABORADO POR:"; ws["I1"] = "Cristian Arciniegas"
        ws["H2"] = "REVISADO POR:"; ws["H3"] = "APROBADO POR:"

        ws["A5"] = "C. COSTOS"; ws["B5"] = "TIENDA / SUCURSAL"; ws["C5"] = "CIUDAD"
        for col in ["A5", "B5", "C5"]:
            ws[col].fill = green_fill
            ws[col].font = white_font

        pos_headers = [("D", "Computador"), ("F", "Teléfono"), ("H", "Biométrico"), ("J", "Lector de Código"), ("L", "Cajón Monedero"), ("N", "Impresora POS")]
        for col_let, title in pos_headers:
            col_idx = openpyxl.utils.column_index_from_string(col_let)
            ws.merge_cells(start_row=6, start_column=col_idx, end_row=6, end_column=col_idx+1)
            cell = ws.cell(row=6, column=col_idx, value=title)
            cell.fill = light_green
            cell.font = dark_green_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        sub_headers = ["Activo", "Estado"]
        for col_idx in range(4, 16):
            cell = ws.cell(row=7, column=col_idx, value=sub_headers[(col_idx - 4) % 2])
            cell.fill = sub_green
            cell.font = gray_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for idx, row in df_nuevo.iterrows():
            ws.append([
                "", row["Oficina"], "",
                row["PC_Activo"], row["PC_Estado"],
                row["Telefono_Activo"], row["Telefono_Estado"],
                row["Biometrico_Activo"], row["Biometrico_Estado"],
                row["Lector_Activo"], row["Lector_Estado"],
                row["Cajon_Activo"], row["Cajon_Estado"],
                row["Impresora_Activa"], row["Impresora_Estado"]
            ])

        wb.save(EXCEL_FILE)

    st.success("¡Datos guardados exitosamente y sincronizados con el Excel!")

st.divider()

# --- ZONA RESTRINGIDA DE DESCARGA ---
st.subheader("📥 Zona de Descarga de Reportes")
if password_ingresada == PASSWORD_SECRETA:
    st.success("Acceso de administrador concedido.")
    if os.path.exists(EXCEL_FILE):
        with open(EXCEL_FILE, "rb") as f:
            st.download_button(
                label="📥 Descargar archivo Excel completo",
                data=f,
                file_name="ESTADO_TIENDAS 2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("Aún no hay registros guardados en el archivo Excel.")
else:
    st.info("🔒 Introduce la contraseña correcta en la barra lateral para habilitar la descarga del archivo Excel.")
