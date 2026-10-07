from excel_formato import crear_excel_formato, guardar_registros

if st.button("💾 Guardar Información del Formulario", type="primary"):
    if not os.path.exists(EXCEL_FILE):
        crear_excel_formato(EXCEL_FILE)          # crea el Excel idéntico al original

    no_encontradas = guardar_registros(EXCEL_FILE, registros_cajas)

    if no_encontradas:
        st.warning(f"Estas tiendas no coinciden con la lista del Excel: {no_encontradas}")
    else:
        st.success("¡Datos guardados exitosamente y sincronizados con el Excel!")
