import base64
from datetime import datetime
import io
import os
import qrcode
import streamlit as st
import pytz
from fpdf import FPDF
from openpyxl import Workbook, load_workbook

# Configuración inicial de la página (optimizada para móviles)
st.set_page_config(
    page_title="Documento de Identificación (DI) - Residuos",
    page_icon="🚛",
    layout="wide",
    initial_sidebar_state="collapsed"
)

EXCEL_PATH = "registro_documentos.xlsx"
SPAIN_TZ = pytz.timezone("Europe/Madrid")

# URL Base oculta en el servidor para generar los QR de verificación
URL_BASE_APP = os.getenv("URL_BASE_APP", "https://gestor-residuos-di-zv7k5cappd8wle3kzxlxskd.streamlit.app")

def obtener_ahora_espana():
    return datetime.now(SPAIN_TZ)

def obtener_siguiente_correlativo() -> int:
    if not os.path.exists(EXCEL_PATH):
        return 1
    try:
        wb = load_workbook(EXCEL_PATH)
        ws = wb.active
        return max(1, ws.max_row)
    except Exception:
        return 1

def generar_numero_di(nima_operador: str, correlativo: int) -> str:
    ahora = obtener_ahora_espana()
    nima_limpio = nima_operador.strip() if nima_operador else "0"
    nima_10 = nima_limpio[:10] if len(nima_limpio) >= 10 else nima_limpio.zfill(10)
    anio = me = ahora.strftime("%Y")
    correlativo_str = str(correlativo).zfill(3)
    return f"{nima_10}{anio}{correlativo_str}"

# --- OPCIONES DE DESPLEGABLES CON DESCRIPCIÓN COMPLETA ---
OPCIONES_OPERADOR = [
    "A02",
    "P03 (Productor > 1000 Tn RNP)",
    "P04 (Productor < 1000 Tn RNP)",
    "G04 (Gestor RNP)",
    "G05 (Gestor Intermedio RNP)"
]

OPCIONES_ORIGEN = [
    "P03 (Productor > 1000 Tn RNP)",
    "P04 (Productor < 1000 Tn RNP)",
    "G04 (Gestor RNP)",
    "G05 (Gestor Intermedio RNP)"
]

OPCIONES_DESTINO = [
    "G04 (Gestor RNP)",
    "G05 (Gestor Intermedio RNP)"
]

OPCIONES_TRANSPORTISTA = [
    "T02 (Transportista RNP)",
    "T01",
    "T03"
]

# --- BARRA LATERAL ---
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
        st.info("ℹ️ Aún no hay registros guardados.")

# --- MODO VISOR (ACCESO DESDE CÓDIGO QR) ---
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
        base64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")
        st.markdown(f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="600" type="application/pdf"></iframe>', unsafe_allow_html=True)
    else:
        st.error("❌ El documento no se encuentra en el servidor.")

    if st.button("⬅️ Volver a la aplicación principal"):
        st.query_params.clear()
        st.rerun()
    st.stop()

# --- MODO FORMULARIO PRINCIPAL ---
ahora_espana = obtener_ahora_espana()

st.title("🚛 Documento de Identificación de Residuos (DI) y Carta de Porte")
st.write("Rellena las secciones para generar el PDF oficial y volcar el registro en Excel.")

with st.form("di_form_completo"):

    # 1. IDENTIFICACIÓN DEL DOCUMENTO Y FECHA
    siguiente_correlativo = obtener_siguiente_correlativo()
    di_sugerido = generar_numero_di("123456789", siguiente_correlativo)

    st.header("1. IDENTIFICACIÓN DEL DOCUMENTO Y FECHA")
    col_d1, col_d2, col_d3 = st.columns(3)
    with col_d1:
        di_num = st.text_input("🆔 Nº Documento (Autogenerado):", value=di_sugerido)
    with col_d2:
        fecha_inicio = st.text_input("Fecha inicio traslado:", value=ahora_espana.strftime("%d/%m/%Y"))
    with col_d3:
        hora_inicio = st.text_input("Hora (España):", value=ahora_espana.strftime("%H:%M"))

    st.markdown("---")

    # 2. OPERADOR DEL TRASLADO
    st.header("2. OPERADOR DEL TRASLADO")
    c_op1, c_op2, c_op3 = st.columns(3)
    with c_op1:
        op_nif = st.text_input("NIF Operador:", value="")
        op_nombre = st.text_input("Razón Social / Nombre:", value="")
        op_tipo = st.selectbox("Tipo Operador:", OPCIONES_OPERADOR)
    with c_op2:
        req_op = " *" if "P04" not in op_tipo else ""
        op_nima = st.text_input(f"NIMA Operador{req_op}:", value="")
        op_inscripcion = st.text_input(f"Nº Inscripción{req_op}:", value="")
        op_direccion = st.text_input("Dirección:", value="")
    with c_op3:
        op_cp = st.text_input("C.P.:", value="")
        op_muni = st.text_input("Municipio:", value="")
        op_prov = st.text_input("Provincia:", value="")
        op_telefono = st.text_input("Teléfono Operador:", value="")
        op_email = st.text_input("Correo Electrónico Operador:", value="")

    st.markdown("---")

    # 3. ORIGEN DEL TRASLADO
    st.header("3. ORIGEN DEL TRASLADO")
    c1, c2, c3 = st.columns(3)
    with c1:
        ori_nif = st.text_input("NIF Origen:", value="")
        ori_nombre = st.text_input("Razón Social Origen:", value="")
        ori_tipo = st.selectbox("Tipo Origen:", OPCIONES_ORIGEN)
    with c2:
        req_ori = " *" if "P04" not in ori_tipo else ""
        ori_nima = st.text_input(f"NIMA Origen{req_ori}:", value="")
        ori_inscripcion = st.text_input(f"Nº Inscripción Origen{req_ori}:", value="")
        ori_direccion = st.text_input("Dirección Origen:", value="")
    with c3:
        ori_cp = st.text_input("C.P. Origen:", value="")
        ori_muni = st.text_input("Municipio Origen:", value="")
        ori_prov = st.text_input("Provincia Origen:", value="")
        ori_telefono = st.text_input("Teléfono Origen:", value="")
        ori_email = st.text_input("Email Origen:", value="")

    st.markdown("---")

    # 4. DESTINO DEL TRASLADO
    st.header("4. DESTINO DEL TRASLADO")
    c1, c2, c3 = st.columns(3)
    with c1:
        des_nif = st.text_input("NIF Destino:", value="")
        des_nombre = st.text_input("Razón Social Destino:", value="")
        des_tipo = st.selectbox("Tipo Destino:", OPCIONES_DESTINO)
    with c2:
        req_des = " *" if "P04" not in des_tipo else ""
        des_nima = st.text_input(f"NIMA Destino{req_des}:", value="")
        des_inscripcion = st.text_input(f"Nº Inscripción Destino{req_des}:", value="")
        des_direccion = st.text_input("Dirección Destino:", value="")
    with c3:
        des_cp = st.text_input("C.P. Destino:", value="")
        des_muni = st.text_input("Municipio Destino:", value="")
        des_prov = st.text_input("Provincia Destino:", value="")
        des_telefono = st.text_input("Teléfono Destino:", value="")
        des_email = st.text_input("Email Destino:", value="")

    st.markdown("---")

    # 5. INFORMACIÓN SOBRE EL RESIDUO
    st.header("5. INFORMACIÓN SOBRE EL RESIDUO QUE SE TRASLADA")
    c1, c2 = st.columns(2)
    with c1:
        ler = st.text_input("Código LER:", value="")
        desc_residuo = st.text_area("Descripción del residuo:", value="")
        cantidad_kg = st.text_input("Cantidad (kg):", value="")
    with c2:
        operacion_tratam = st.text_input("Operación Tratamiento Destino:", value="")
        operacion_desagregada = st.text_input("Operación Destino Desagregada:", value="")
        desc_operacion = st.text_input("Descripción Op. Tratamiento:", value="")

    st.markdown("---")

    # 6. INFORMACIÓN RELATIVA AL TRANSPORTISTA
    st.header("6. INFORMACIÓN RELATIVA AL TRANSPORTISTA")
    c1, c2, c3 = st.columns(3)
    with c1:
        trans_nif = st.text_input("N.I.F. Transportista:", value="")
        trans_nombre = st.text_input("Razón Social / Nombre Transportista:", value="")
        trans_tipo = st.selectbox("Tipo Transportista:", OPCIONES_TRANSPORTISTA, index=0)
    with c2:
        req_trans = " *" if "P04" not in trans_tipo else ""
        trans_nima = st.text_input(f"NIMA Transportista{req_trans}:", value="")
        trans_inscripcion = st.text_input(f"Nº Inscripción / Autorización{req_trans}:", value="")
        trans_direccion = st.text_input("Dirección Transportista:", value="")
    with c3:
        trans_conductor = st.text_input("Conductor:", value="")
        trans_matricula = st.text_input("Matrícula y Vehículo:", value="")
        trans_telefono = st.text_input("Teléfono Transportista:", value="")
        trans_email = st.text_input("Email Transportista:", value="")

    st.markdown("---")

    # 7. ACEPTACIÓN DEL RESIDUO
    st.header("7. INFORMACIÓN SOBRE LA ACEPTACIÓN DEL RESIDUO")
    c1, c2, c3 = st.columns(3)
    with c1:
        fecha_entrega = st.text_input("Fecha Entrega:", value="")
        kg_recibidos = st.text_input("Kg. Netos Recibidos:", value="")
    with c2:
        fecha_aceptacion = st.text_input("Fecha Aceptación/Rechazo:", value="")
        aceptacion_estado = st.selectbox("Aceptación:", ["", "Sí", "No"], index=0)
    with c3:
        motivo_rechazo = st.text_input("Motivo de rechazo (si aplica):", value="")

    btn_generar = st.form_submit_button("🚀 Generar PDF Oficial y Registrar")

# --- PROCESAMIENTO Y VALIDACIÓN ---
if btn_generar:
    errores = []

    if not di_num:
        errores.append("El Número de Documento (DI) es obligatorio.")

    # Validaciones condicionales: Obligatorio si NO contiene P04
    if "P04" not in op_tipo:
        if not op_nima.strip():
            errores.append("El NIMA del Operador es obligatorio para el tipo seleccionado.")
        if not op_inscripcion.strip():
            errores.append("El Nº de Inscripción del Operador es obligatorio para el tipo seleccionado.")

    if "P04" not in ori_tipo:
        if not ori_nima.strip():
            errores.append("El NIMA de Origen es obligatorio para el tipo seleccionado.")
        if not ori_inscripcion.strip():
            errores.append("El Nº de Inscripción de Origen es obligatorio para el tipo seleccionado.")

    if "P04" not in des_tipo:
        if not des_nima.strip():
            errores.append("El NIMA de Destino es obligatorio para el tipo seleccionado.")
        if not des_inscripcion.strip():
            errores.append("El Nº de Inscripción de Destino es obligatorio para el tipo seleccionado.")

    if "P04" not in trans_tipo:
        if not trans_nima.strip():
            errores.append("El NIMA del Transportista es obligatorio para el tipo seleccionado.")
        if not trans_inscripcion.strip():
            errores.append("El Nº de Inscripción del Transportista es obligatorio para el tipo seleccionado.")

    if errores:
        for err in errores:
            st.error(f"⚠️ {err}")
    else:
        # Generación del PDF y Excel
        base_limpia = URL_BASE_APP.strip().rstrip("/")
        enlace_qr = f"{base_limpia}?doc={di_num}"

        # 1. Código QR
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

        pdf.image(qr_path, x=150, y=10, w=50, h=50)

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
        pdf.output(pdf_out_filename)
        pdf_bytes = pdf.output(dest='S').encode('latin-1')

        # Excel
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
            st.image(qr_path, caption="Código QR Generado", width=150)
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
