import streamlit as st
import pandas as pd
from datetime import datetime
import os

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

EXCEL_FILE = "control_equipos_tecnologicos.xlsx"

st.set_page_config(page_title="Control de Equipos Tecnológicos", layout="centered")

st.title("🖥️ Formulario de Control de Equipos Tecnológicos")
st.write("Selecciona la oficina e ingresa la información de las cajas correspondientes.")

# Paso 1: Selección de Oficina
oficina_seleccionada = st.selectbox("Seleccione la Oficina / Sucursal:", OFICINAS)

# Paso 2: Cantidad de cajas (máximo 5)
cantidad_cajas = st.number_input("¿Cuántas cajas va a registrar? (Máximo 5):", min_value=1, max_value=5, value=1, step=1)

st.divider()

# Diccionario para almacenar los datos de las cajas de forma temporal
registros_cajas = []

# Paso 3: Preguntas dinámicas por cada caja (Página / Sección por caja)
for i in range(1, int(cantidad_cajas) + 1):
    st.subheader(f"📦 Configuración de la Caja #{i}")
    
    col1, col2 = st.columns(2)
    with col1:
        pc_activo = st.text_input(f"Computador - Activo Fijo (Caja {i})", key=f"pc_act_{i}")
        pc_estado = st.selectbox(f"Computador - Estado (Caja {i})", ["BUENO", "REGULAR", "MALO", "N/A"], key=f"pc_est_{i}")
        
        tel_activo = st.text_input(f"Teléfono - Activo Fijo (Caja {i})", key=f"tel_act_{i}")
        tel_estado = st.selectbox(f"Teléfono - Estado (Caja {i})", ["BUENO", "REGULAR", "MALO", "N/A"], key=f"tel_est_{i}")
        
        bio_activo = st.text_input(f"Biométrico - Activo Fijo (Caja {i})", key=f"bio_act_{i}")
        bio_estado = st.selectbox(f"Biométrico - Estado (Caja {i})", ["BUENO", "REGULAR", "MALO", "N/A"], key=f"bio_est_{i}")

    with col2:
        lec_activo = st.text_input(f"Lector de Código - Activo Fijo (Caja {i})", key=f"lec_act_{i}")
        lec_estado = st.selectbox(f"Lector de Código - Estado (Caja {i})", ["BUENO", "REGULAR", "MALO", "N/A"], key=f"lec_est_{i}")
        
        caj_activo = st.text_input(f"Cajón Monedero - Activo Fijo (Caja {i})", key=f"caj_act_{i}")
        caj_estado = st.selectbox(f"Cajón Monedero - Estado (Caja {i})", ["BUENO", "REGULAR", "MALO", "N/A"], key=f"caj_est_{i}")
        
        imp_activo = st.text_input(f"Impresora POS - Activo Fijo (Caja {i})", key=f"imp_act_{i}")
        imp_estado = st.selectbox(f"Impresora POS - Estado (Caja {i})", ["BUENO", "REGULAR", "MALO", "N/A"], key=f"imp_est_{i}")

    # Guardar datos estructurados de esta caja
    registros_cajas.append({
        "Fecha_Registro": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Oficina": oficina_seleccionada,
        "Numero_Caja": f"Caja {i}",
        "PC_Activo": pc_activo,
        "PC_Estado": pc_estado,
        "Telefono_Activo": tel_activo,
        "Telefono_Estado": tel_estado,
        "Biometrico_Activo": bio_activo,
        "Biometrico_Estado": bio_estado,
        "Lector_Activo": lec_activo,
        "Lector_Estado": lec_estado,
        "Cajon_Activo": caj_activo,
        "Cajon_Estado": caj_estado,
        "Impresora_Activa": imp_activo,
        "Impresora_Estado": imp_estado
    })
    st.divider()

# Botón para enviar y alimentar el Excel
if st.button("💾 Guardar Información en Excel", type="primary"):
    df_nuevo = pd.DataFrame(registros_cajas)
    
    # Si ya existe el archivo, se le añade la nueva información (acumulativo)
    if os.path.exists(EXCEL_FILE):
        df_existente = pd.read_excel(EXCEL_FILE)
        df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
    else:
        df_final = df_nuevo
        
    df_final.to_excel(EXCEL_FILE, index=False)
    st.success("¡Datos guardados exitosamente en el archivo de Excel!")
    
    # Opción para descargar el Excel directamente
    with open(EXCEL_FILE, "rb") as f:
        st.download_button(
            label="📥 Descargar archivo Excel actualizado",
            data=f,
            file_name="control_equipos_tecnologicos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
