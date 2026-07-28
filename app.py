import base64
from datetime import datetime
import io
import os
import qrcode
import streamlit as st
from fpdf import FPDF
from openpyxl import Workbook, load_workbook

# Configuración inicial de la página
st.set_page_config(
    page_title="Documento de Identificación (DI) - Residuos",
    page_icon="🚛",
    layout="wide",
)

EXCEL_PATH = "registro_documentos.xlsx"

# ==========================================
# BARRA LATERAL (Acceso constante al Excel)
# ==========================================
with st.sidebar:
    st.header("📊 Gestión de Registros")
    st.write("Descarga la base de datos completa de documentos generados:")

    if os.path.exists(EXCEL_PATH):
        with open(EXCEL_PATH, "rb") as f_excel:
            st.download_button(
                label="📥 Descargar Registro Excel Completo",
                data=f_excel,
                file_name="registro_documentos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="btn_excel_sidebar"
            )
    else:
        st.info("ℹ️ Aún no hay registros guardados. Genera el primer documento para crear el Excel.")

# ==========================================
# FUNCIONES AUXILIARES
# ==========================================
def generar_numero_di(nima_operador: str) -> str:
    """Genera automáticamente el Nº DI con el criterio:
    [NIMA 10 dígitos] + [Año 4 dígitos] + [Contador 7 dígitos (MMDDHHM)]
    """
    ahora = datetime.now()
    nima_limpio = nima_operador.strip() if nima_operador else "0"
    nima_10 = nima_limpio[:10] if len(nima_limpio) >= 10 else nima_limpio.zfill(10)
    anio = ahora.strftime("%Y")
    contador_7 = ahora.strftime("%m%d%H") + ahora.strftime("%M")[0]
    return f"{nima_10}{anio}{contador_7}"


# ==========================================
# 1. MODO VISOR (Lectura al escanear el QR)
# ==========================================
if "doc" in st.query_params:
    doc_id = st.query_params["doc"]
    st.title(f"🔎 Verificación de Documento: {doc_id}")

    pdf_filename = f"DI_{doc_id.replace('/', '_')}.pdf"

    if os.path.exists(pdf_filename):
        st.success("✅ Documento original encontrado y verificado.")

        with open(pdf_filename, "rb") as f:
            pdf_bytes = f.read()

        st.download_button(
            label="📥 Descargar Documento PDF Oficial",
            data=pdf_bytes,
            file_name=pdf_filename,
            mime="application/pdf",
            key="btn_visor_pdf"
        )

        st.markdown("---")
        st.subheader("📄 Vista previa del documento:")

        base64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")
        pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="600" type="application/pdf"></iframe>'
        st.markdown(pdf_display, unsafe_allow_html=True)

    else:
        st.error("❌ El documento no se encuentra en el servidor. Verifica que haya sido generado correctamente.")

    st.markdown("---")
    if st.button("⬅️ Volver a la aplicación principal"):
        st.query_params.clear()
        st.rerun()

    st.stop()


# ==========================================
# 2. MODO FORMULARIO (Creación del documento)
# ==========================================
st.title("🚛 Documento de Identificación de Residuos (DI) y Carta de Porte")
st.write("Rellena las secciones para generar el PDF oficial y volcar el registro completo en Excel.")

# Enlace base configurado con tu subdominio real
url_base = st.text_input(
    "🌐 URL Base de la Aplicación:",
    value="https://gestor-residuos-di-zv7k5cappd8wle3kzxlxskd.streamlit.app"
)

with st.form("di_form_completo"):
    st.header("1. OPERADOR Y DATOS GENERALES")
    c_op1, c_op2, c_op3 = st.columns(3)
    with c_op1:
        op_nima = st.text_input("NIMA Operador:", value="123456789")
        op_nif = st.text_input("NIF Operador:", value="B12345678")
        op_nombre = st.text_input("Razón Social / Nombre:", value="Empresa Operadora S.L.")
    with c_op2:
        op_inscripcion = st.text_input("Nº Inscripción:", value="INS-001")
        op_tipo = st.text_input("Tipo Operador:", value="Gestor")
        op_direccion = st.text_input("Dirección:", value="Calle Industria 12")
    with c_op3:
        op_cp = st.text_input("C.P.:", value="29000")
        op_muni = st.text_input("Municipio:", value="Málaga")
        op_prov = st.text_input("Provincia:", value="Málaga")
        op_telefono = st.text_input("Teléfono Operador:", value="952000000")
        op_email = st.text_input("Correo Electrónico Operador:", value="info@operador.com")

    # Autogeneración dinámica del Nº DI
    di_sugerido = generar_numero_di(op_nima)

    st.markdown("---")
    col_d1, col_d2, col_d3 = st.columns(3)
    with col_d1:
        di_num = st.text_input("🆔 Nº Documento (Autogenerado):", value=di_sugerido)
    with col_d2:
        fecha_inicio = st.text_input("Fecha inicio traslado:", value=datetime.now().strftime("%d/%m/%Y"))
    with col_d3:
        hora_inicio = st.text_input("Hora:", value=datetime.now().strftime("%H:%M"))

    st.markdown("---")
    st.header("2. ORIGEN DEL TRASLADO")
    c1, c2, c3 = st.columns(3)
    with c1:
        ori_nif = st.text_input("NIF Origen:", value="A98765432")
        ori_nombre = st.text_input("Razón Social Origen:", value="Fábrica Origen S.A.")
        ori_nima = st.text_input("NIMA Origen:", value="987654321")
    with c2:
        ori_inscripcion = st.text_input("Nº Inscripción Origen:", value="ORI-002")
        ori_tipo = st.text_input("Tipo Origen:", value="Productor")
        ori_direccion = st.text_input("Dirección Origen:", value="Polígono Industrial Norte 5")
    with c3:
        ori_cp = st.text_input("C.P. Origen:", value="29001")
        ori_muni = st.text_input("Municipio Origen:", value="Málaga")
        ori_prov = st.text_input("Provincia Origen:", value="Málaga")
        ori_telefono = st.text_input("Teléfono Origen:", value="952111222")
        ori_email = st.text_input("Email Origen:", value="origen@empresa.com")

    st.markdown("---")
    st.header("3. DESTINO DEL TRASLADO")
    c1, c2, c3 = st.columns(3)
    with c1:
        des_nif = st.text_input("NIF Destino:", value="B55544332")
        des_nombre = st.text_input("Razón Social Destino:", value="Planta Reciclaje Destino S.L.")
        des_nima = st.text_input("NIMA Destino:", value="555443322")
    with c2:
        des_inscripcion = st.text_input("Nº Inscripción Destino:", value="DES-003")
        des_tipo = st.text_input("Tipo Destino:", value="Planta Valorización")
        des_direccion = st.text_input("Dirección Destino:", value="Carretera Nacional Km 5")
    with c3:
        des_cp = st.text_input("C.P. Destino:", value="29002")
        des_muni = st.text_input("Municipio Destino:", value="Antequera")
        des_prov = st.text_input("Provincia Destino:", value="Málaga")
        des_telefono = st.text_input("Teléfono Destino:", value="952333444")
        des_email = st.text_input("Email Destino:", value="destino@reciclaje.com")

    st.markdown("---")
    st.header("4. INFORMACIÓN SOBRE EL RESIDUO QUE SE TRASLADA")
    c1, c2 = st.columns(2)
    with c1:
        ler = st.text_input("Código LER:", value="17 09 04")
        desc_residuo = st.text_area("Descripción del residuo:", value="Residuos mezclados de construcción y demolición")
        cantidad_kg = st.text_input("Cantidad (kg):", value="2500")
    with c2:
        operacion_tratam = st.text_input("Operación Tratamiento Destino:", value="R13")
        operacion_desagregada = st.text_input("Operación Destino Desagregada:", value="R1301")
        desc_operacion = st.text_input("Descripción Op. Tratamiento:", value="Acumulación de residuos previa a valorización")

    st.markdown("---")
    st.header("5. INFORMACIÓN RELATIVA AL TRANSPORTISTA")
    c1, c2, c3 = st.columns(3)
    with c1:
        trans_nif = st.text_input("N.I.F. Transportista:", value="B11223344")
        trans_nombre = st.text_input("Razón Social Transportista:", value="Transportes Rápidos S.L.")
        trans_nima = st.text_input("NIMA Transportista:", value="112233445")
    with c2:
        trans_inscripcion = st.text_input("Nº Inscripción Transportista:", value="TRA-004")
        trans_tipo = st.text_input("Tipo Transportista:", value="Transportista Profesional")
        trans_direccion = st.text_input("Dirección Transportista:", value="Av. Logística 8")
    with c3:
        trans_conductor = st.text_input("Conductor:", value="Juan Pérez")
        trans_matricula = st.text_input("Matrícula y Vehículo:", value="1234-XYZ / Camión")
        trans_telefono = st.text_input("Teléfono Transportista:", value="600112233")
        trans_email = st.text_input("Email Transportista:", value="trans@rapidos.com")

    st.markdown("---")
    st.header("6. INFORMACIÓN SOBRE LA ACEPTACIÓN DEL RESIDUO")
    c1, c2, c3 = st.columns(3)
    with c1:
        fecha_entrega = st.text_input("Fecha Entrega:", value=datetime.now().strftime("%d/%m/%Y"))
        kg_recibidos = st.text_input("Kg. Netos Recibidos:", value="2500")
    with c2:
        fecha_aceptacion = st.text_input("Fecha Aceptación/Rechazo:", value=datetime.now().strftime("%d/%m/%Y"))
        aceptacion_estado = st.selectbox("Aceptación:", ["Sí", "No"])
    with c3:
        motivo_rechazo = st.text_input("Motivo de rechazo (si aplica):", value="")

    btn_generar = st.form_submit_button("🚀 Generar PDF Oficial y Registrar")


# ==========================================
# PROCESAMIENTO TRAS PULSAR EL BOTÓN
# ==========================================
if btn_generar:
    if not di_num:
        st.error("Por favor, introduce el Número de Documento (DI).")
    else:
        base_limpia = url_base.strip().rstrip("/")
        enlace_qr = f"{base_limpia}?doc={di_num}"

        # 1. Generar código QR
        qr = qrcode.QRCode(box_size=10, border=2)
        qr.add_data(enlace_qr)
        qr.make(fit=True)
        img_qr = qr.make_image(fill_color="black", back_color="white")
        qr_path = "temp_qr.png"
        img_qr.save(qr_path)

        # 2. Generar PDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=10)

        def s(txt):
            return str(txt).encode("latin-1", "replace").decode("latin-1")

        # Dibujar QR (esquina superior derecha)
        pdf.image(qr_path, x=150, y=10, w=50, h=50)

        # Cabecera PDF
        pdf.set_font("Arial", "B", 10)
        pdf.cell(135, 6, s("DOCUMENTO DE IDENTIFICACIÓN DE RESIDUOS Y CARTA DE PORTE"), border=0, ln=True)
        pdf.ln(2)
        pdf.set_font("Arial", "", 8)
        pdf.cell(135, 6, s(f"Documento de Identificación nº: {di_num}"), border=1, ln=True)
        pdf.cell(67, 6, s(f"Fecha inicio traslado: {fecha_inicio}"), border=1)
        pdf.cell(68, 6, s(f"Hora: {hora_inicio}"), border=1, ln=True)

        pdf.set_y(63)

        # Operador
        pdf.set_font("Arial", "B", 8)
        pdf.set_fill_color(220, 220, 220)
        pdf.cell(190, 5, s("OPERADOR DEL TRASLADO"), border=1, ln=True, fill=True)
        pdf.set_font("Arial", "", 7.5)
        pdf.cell(50, 5, s(f"NIF: {op_nif}"), border=1)
        pdf.cell(140, 5, s(f"Razón social/Nombre: {op_nombre}"), border=1, ln=True)
        pdf.cell(50, 5, s(f"NIMA: {op_nima}"), border=1)
        pdf.cell(80, 5, s(f"Nº inscripción: {op_inscripcion}"), border=1)
        pdf.cell(60, 5, s(f"Tipo: {op_tipo}"), border=1, ln=True)
        pdf.cell(130, 5, s(f"Dirección: {op_direccion}"), border=1)
        pdf.cell(60, 5, s(f"C.P.: {op_cp}"), border=1, ln=True)
        pdf.cell(95, 5, s(f"Municipio: {op_muni}"), border=1)
        pdf.cell(95, 5, s(f"Provincia: {op_prov}"), border=1, ln=True)
        pdf.cell(95, 5, s(f"Teléfono: {op_telefono}"), border=1)
        pdf.cell(95, 5, s(f"Correo electrónico: {op_email}"), border=1, ln=True)
        pdf.ln(3)

        # Origen
        pdf.set_font("Arial", "B", 8)
        pdf.cell(190, 5, s("ORIGEN DEL TRASLADO"), border=1, ln=True, fill=True)
        pdf.set_font("Arial", "I", 7.5)
        pdf.cell(190, 4, s("Información de la instalación origen del traslado:"), border="LR", ln=True)
        pdf.set_font("Arial", "", 7.5)
        pdf.cell(50, 5, s(f"NIF: {ori_nif}"), border=1)
        pdf.cell(140, 5, s(f"Razón social/Nombre: {ori_nombre}"), border=1, ln=True)
        pdf.cell(50, 5, s(f"NIMA: {ori_nima}"), border=1)
        pdf.cell(80, 5, s(f"Nº inscripción: {ori_inscripcion}"), border=1)
        pdf.cell(60, 5, s(f"Tipo: {ori_tipo}"), border=1, ln=True)
        pdf.cell(130, 5, s(f"Dirección: {ori_direccion}"), border=1)
        pdf.cell(60, 5, s(f"C.P.: {ori_cp}"), border=1, ln=True)
        pdf.cell(95, 5, s(f"Municipio: {ori_muni}"), border=1)
        pdf.cell(95, 5, s(f"Provincia: {ori_prov}"), border=1, ln=True)
        pdf.cell(95, 5, s(f"Teléfono: {ori_telefono}"), border=1)
        pdf.cell(95, 5, s(f"Correo electrónico: {ori_email}"), border=1, ln=True)
        pdf.ln(3)

        # Destino
        pdf.set_font("Arial", "B", 8)
        pdf.cell(190, 5, s("DESTINO DEL TRASLADO"), border=1, ln=True, fill=True)
        pdf.set_font("Arial", "I", 7.5)
        pdf.cell(190, 4, s("Información de la instalación de destino:"), border="LR", ln=True)
        pdf.set_font("Arial", "", 7.5)
        pdf.cell(50, 5, s(f"NIF: {des_nif}"), border=1)
        pdf.cell(140, 5, s(f"Razón social/Nombre: {des_nombre}"), border=1, ln=True)
        pdf.cell(50, 5, s(f"NIMA: {des_nima}"), border=1)
        pdf.cell(80, 5, s(f"Nº inscripción: {des_inscripcion}"), border=1)
        pdf.cell(60, 5, s(f"Tipo: {des_tipo}"), border=1, ln=True)
        pdf.cell(130, 5, s(f"Dirección: {des_direccion}"), border=1)
        pdf.cell(60, 5, s(f"C.P.: {des_cp}"), border=1, ln=True)
        pdf.cell(95, 5, s(f"Municipio: {des_muni}"), border=1)
        pdf.cell(95, 5, s(f"Provincia: {des_prov}"), border=1, ln=True)
        pdf.cell(95, 5, s(f"Teléfono: {des_telefono}"), border=1)
        pdf.cell(95, 5, s(f"Correo electrónico: {des_email}"), border=1, ln=True)
        pdf.ln(3)

        # Residuo
        pdf.set_font("Arial", "B", 8)
        pdf.cell(190, 5, s("INFORMACIÓN SOBRE EL RESIDUO QUE SE TRASLADA"), border=1, ln=True, fill=True)
        pdf.set_font("Arial", "", 7.5)
        pdf.cell(50, 5, s(f"Código LER: {ler}"), border=1)
        pdf.cell(140, 5, s(f"Descripción: {desc_residuo}"), border=1, ln=True)
        pdf.cell(95, 5, s(f"Op. Tratamiento Destino: {operacion_tratam}"), border=1)
        pdf.cell(95, 5, s(f"Op. Tratamiento Desagregada: {operacion_desagregada}"), border=1, ln=True)
        pdf.cell(130, 5, s(f"Descripción Op. Tratamiento: {desc_operacion}"), border=1)
        pdf.cell(60, 5, s(f"Cantidad (kg): {cantidad_kg}"), border=1, ln=True)
        pdf.ln(3)

        # Transportista
        pdf.set_font("Arial", "B", 8)
        pdf.cell(190, 5, s("INFORMACIÓN RELATIVA AL TRANSPORTISTA"), border=1, ln=True, fill=True)
        pdf.set_font("Arial", "", 7.5)
        pdf.cell(50, 5, s(f"N.I.F.: {trans_nif}"), border=1)
        pdf.cell(140, 5, s(f"Razón social/Nombre: {trans_nombre}"), border=1, ln=True)
        pdf.cell(50, 5, s(f"NIMA: {trans_nima}"), border=1)
        pdf.cell(80, 5, s(f"Nº inscripción: {trans_inscripcion}"), border=1)
        pdf.cell(60, 5, s(f"Tipo: {trans_tipo}"), border=1, ln=True)
        pdf.cell(130, 5, s(f"Dirección: {trans_direccion}"), border=1)
        pdf.cell(60, 5, s(f"Conductor: {trans_conductor}"), border=1, ln=True)
        pdf.cell(95, 5, s(f"Matrícula y vehículo: {trans_matricula}"), border=1)
        pdf.cell(95, 5, s(f"Teléfono / Email: {trans_telefono} - {trans_email}"), border=1, ln=True)
        pdf.ln(3)

        # Aceptación
        pdf.set_font("Arial", "B", 8)
        pdf.cell(190, 5, s("INFORMACIÓN SOBRE LA ACEPTACIÓN DEL RESIDUO"), border=1, ln=True, fill=True)
        pdf.set_font("Arial", "", 7.5)
        pdf.cell(95, 5, s(f"Fecha entrega: {fecha_entrega}"), border=1)
        pdf.cell(95, 5, s(f"Kg. netos recibidos: {kg_recibidos}"), border=1, ln=True)
        pdf.cell(95, 5, s(f"Fecha aceptación/rechazo: {fecha_aceptacion}"), border=1)
        pdf.cell(95, 5, s(f"Aceptación: [{aceptacion_estado}]"), border=1, ln=True)
        if motivo_rechazo:
            pdf.cell(190, 5, s(f"Motivo de rechazo: {motivo_rechazo}"), border=1, ln=True)

        pdf_out_filename = f"DI_{di_num.replace('/', '_')}.pdf"
        
        # Guardado en disco
        pdf.output(pdf_out_filename)
        
        # Obtención de bytes directos para descarga segura
        pdf_bytes = pdf.output(dest='S').encode('latin-1')

        # 3. Guardar en Excel con TODOS los datos
        columnas_excel = [
            "Nº DI", "Fecha Inicio Traslado", "Hora Inicio", "NIF Operador", "Razón Social Operador",
            "NIMA Operador", "Nº Inscripción Operador", "Tipo Operador", "Dirección Operador", "CP Operador",
            "Municipio Operador", "Provincia Operador", "Teléfono Operador", "Email Operador",
            "NIF Origen", "Razón Social Origen", "NIMA Origen", "Nº Inscripción Origen", "Tipo Origen",
            "Dirección Origen", "CP Origen", "Municipio Origen", "Provincia Origen", "Teléfono Origen",
            "Email Origen", "NIF Destino", "Razón Social Destino", "NIMA Destino", "Nº Inscripción Destino",
            "Tipo Destino", "Dirección Destino", "CP Destino", "Municipio Destino", "Provincia Destino",
            "Teléfono Destino", "Email Destino", "Código LER", "Descripción Residuo", "Cantidad (kg)",
            "Op. Tratamiento Destino", "Op. Tratamiento Desagregada", "Descripción Op. Tratamiento",
            "NIF Transportista", "Razón Social Transportista", "NIMA Transportista", "Nº Inscripción Transportista",
            "Tipo Transportista", "Dirección Transportista", "Conductor", "Matrícula / Vehículo",
            "Teléfono Transportista", "Email Transportista", "Fecha Entrega", "Kg Netos Recibidos",
            "Fecha Aceptación/Rechazo", "Estado Aceptación", "Motivo Rechazo", "Enlace QR Verificación"
        ]

        fila_datos = [
            di_num, fecha_inicio, hora_inicio, op_nif, op_nombre,
            op_nima, op_inscripcion, op_tipo, op_direccion, op_cp,
            op_muni, op_prov, op_telefono, op_email,
            ori_nif, ori_nombre, ori_nima, ori_inscripcion, ori_tipo,
            ori_direccion, ori_cp, ori_muni, ori_prov, ori_telefono,
            ori_email, des_nif, des_nombre, des_nima, des_inscripcion,
            des_tipo, des_direccion, des_cp, des_muni, des_prov,
            des_telefono, des_email, ler, desc_residuo, cantidad_kg,
            operacion_tratam, operacion_desagregada, desc_operacion,
            trans_nif, trans_nombre, trans_nima, trans_inscripcion,
            trans_tipo, trans_direccion, trans_conductor, trans_matricula,
            trans_telefono, trans_email, fecha_entrega, kg_recibidos,
            fecha_aceptacion, aceptacion_estado, motivo_rechazo, enlace_qr
        ]

        if not os.path.exists(EXCEL_PATH):
            wb = Workbook()
            ws = wb.active
            ws.title = "Registros DI"
            ws.append(columnas_excel)
        else:
            wb = load_workbook(EXCEL_PATH)
            ws = wb.active

        ws.append(fila_datos)
        wb.save(EXCEL_PATH)

        st.success(f"✅ Documento generado y registrado correctamente: **{di_num}**")

        col_a, col_b = st.columns([1, 3])
        with col_a:
            st.image(qr_path, caption="Código QR", width=150)
        with col_b:
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                st.download_button(
                    label="📄 Descargar PDF Oficial",
                    data=pdf_bytes,
                    file_name=pdf_out_filename,
                    mime="application/pdf",
                    key="btn_pdf_main"
                )
            with col_btn2:
                with open(EXCEL_PATH, "rb") as f_excel:
                    st.download_button(
                        label="📊 Descargar Registro Excel",
                        data=f_excel,
                        file_name="registro_documentos.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_excel_main"
                    )
            st.code(f"URL del QR: {enlace_qr}", language="text")
