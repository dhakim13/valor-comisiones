
    with colC:
        dist = st.text_input("Distribuidor (opcional para filtrar altas)", value="")

    if base_file and st.button("Generar reporte"):
        try:
            xls = pd.ExcelFile(base_file, engine="openpyxl")
            buf = calc_report(xls, int(year), int(month), dist_filtro=dist.strip())
            fname = f"COMISION VALOR DISTRIBUIDOR {dist.upper() or 'GENERAL'} {spanish_month_name(pd.Timestamp(int(year),int(month),1))} {int(year)}.xlsx"
            st.success("✅ Reporte generado.")
            st.download_button("⬇️ Descargar Excel", data=buf, file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error("Ocurrió un error al generar el reporte.")
            st.exception(e)
