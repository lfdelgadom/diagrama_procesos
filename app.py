from __future__ import annotations

import io
import json
import re
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Modelador de procesos | SIPOC y caracterización",
    page_icon="◆",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------------------------------------------------------
# Estilo visual inspirado en los modeladores BPMN, sin copiar marcas o activos.
# -----------------------------------------------------------------------------
st.markdown(
    """
    <style>
      :root { --azul:#0b5cab; --azul2:#17365d; --cielo:#eaf3fb; --gris:#f5f7fa; }
      .stApp { background: #f4f6f8; }
      [data-testid="stSidebar"] { background: linear-gradient(180deg,#102a43 0%,#17365d 100%); }
      [data-testid="stSidebar"] * { color: #f7fbff; }
      .hero {background:linear-gradient(115deg,#123a63,#0b5cab);padding:1.15rem 1.4rem;
             border-radius:14px;color:white;box-shadow:0 6px 18px rgba(22,54,93,.18);margin-bottom:.8rem}
      .hero h1 {font-size:1.7rem;margin:0 0 .15rem 0;color:white;}
      .hero p {margin:0;color:#dcecff;}
      .proc-card {background:white;border:1px solid #dbe3ec;border-left:5px solid #0b5cab;
                  border-radius:10px;padding:.85rem 1rem;margin:.3rem 0 .8rem 0;}
      .small-note {font-size:.82rem;color:#5b6776}
      div[data-testid="stMetric"] {background:#fff;border:1px solid #dbe3ec;border-radius:10px;padding:.55rem .75rem;}
      .stTabs [data-baseweb="tab-list"] {gap:.4rem;}
      .stTabs [data-baseweb="tab"] {background:#fff;border-radius:8px 8px 0 0;border:1px solid #dbe3ec;padding:.45rem .8rem;}
      .stTabs [aria-selected="true"] {background:#eaf3fb;color:#0b5cab;}
    </style>
    """,
    unsafe_allow_html=True,
)

EXPECTED_ACT_COLS = [
    "id", "nombre", "tipo_elemento", "responsable", "carril", "entrada", "salida",
    "sistema", "documento", "es_decision", "pregunta_decision", "siguiente",
    "siguiente_si", "siguiente_no", "tiempo_dias", "observaciones",
]


def clean_col(value: object) -> str:
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9áéíóúüñ]+", "_", text)
    return text.strip("_")


def text(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def safe_id(value: object) -> str:
    raw = re.sub(r"[^A-Za-z0-9_]", "_", text(value))
    return raw or "nodo"


def boolish(value: object) -> bool:
    return text(value).lower() in {"sí", "si", "yes", "true", "1", "x"}


def excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name[:31], index=False)
    return output.getvalue()


# Datos de demostración incorporados. Así el despliegue requiere únicamente
# app.py y requirements.txt; los Excel cargados por el usuario los reemplazan.
DEMO_META = {
    "codigo_proceso": "AC-HOM-01",
    "nombre_proceso": "Gestión de solicitudes de homologación de asignaturas",
    "institucion": "Universidad Horizonte del Pacífico",
    "macroproceso": "Gestión Académica",
    "tipo_proceso": "Misional académico",
    "version": "1.0",
    "fecha_elaboracion": "2026-09-16",
}

DEMO_CARACTERIZACION = [
    ("Código", "AC-HOM-01"),
    ("Nombre del proceso", "Gestión de solicitudes de homologación de asignaturas"),
    ("Macroproceso", "Gestión Académica"), ("Tipo", "Misional académico"),
    ("Versión", "1.0"), ("Institución", "Universidad Horizonte del Pacífico"),
    ("Responsable", "Director del programa académico"),
    ("Dependencia líder", "Facultad correspondiente"),
    ("Objetivo", "Gestionar las solicitudes de homologación mediante verificación documental, evaluación de equivalencia, decisión institucional, registro y notificación, garantizando trazabilidad y aplicación uniforme de criterios académicos."),
    ("Alcance - inicio", "Inicia con el diligenciamiento y envío de la solicitud por parte del estudiante."),
    ("Alcance - fin", "Finaliza con el registro de la decisión, la notificación y el archivo del expediente."),
    ("Incluye", "Radicación, revisión documental, validación de pago, evaluación académica, decisión, registro, notificación y archivo."),
    ("No incluye", "Admisión, matrícula financiera, convalidación de títulos ni modificación del plan de estudios."),
    ("Requisitos ficticios", "Intensidad horaria mínima del 80%; similitud temática mínima del 75%; nota mínima de 3,5/5,0; documentos oficiales; solicitud dentro del periodo."),
    ("Recursos", "Portal estudiantil, sistema académico, repositorio documental, correo institucional y personal académico-administrativo."),
    ("Registros", "Solicitud, radicado, lista de chequeo, matriz de equivalencia, concepto, acta o decisión, notificación y expediente electrónico."),
]

DEMO_SIPOC = [
    ("SP-01", "Estudiante solicitante", "Formulario, identificación y solicitud firmada", "Recibir y radicar solicitud", "Solicitud radicada", "Auxiliar académico", 1),
    ("SP-02", "Institución educativa de origen", "Certificado oficial de calificaciones y syllabus", "Verificar requisitos y documentos", "Solicitud validada o devuelta", "Estudiante solicitante / Dirección de programa", 2),
    ("SP-03", "Tesorería", "Confirmación del pago del estudio", "Validar requisito financiero", "Pago confirmado o novedad financiera", "Dirección de programa", 3),
    ("SP-04", "Dirección del programa", "Plan de estudios vigente y asignación de evaluador", "Evaluar equivalencia académica", "Matriz de equivalencia y concepto académico", "Comité Curricular / Secretaría Académica", 4),
    ("SP-05", "Comité Curricular o Director de programa", "Concepto académico y expediente completo", "Aprobar o rechazar homologación", "Acta o decisión académica", "Registro Académico", 5),
    ("SP-06", "Secretaría Académica", "Decisión formalizada", "Registrar la decisión", "Historial académico actualizado", "Estudiante solicitante / Facultad", 6),
    ("SP-07", "Registro Académico", "Confirmación del registro", "Notificar, archivar y cerrar", "Comunicación oficial y expediente cerrado", "Estudiante solicitante / Aseguramiento de la Calidad", 7),
]

DEMO_ACTIVITIES = [
    ("A00","Inicio","Inicio","Estudiante","Estudiante","","Proceso iniciado","Portal estudiantil","","No","","A01","","",0,"Evento de inicio"),
    ("A01","Diligenciar solicitud","Actividad","Estudiante","Estudiante","Credenciales y formulario","Solicitud en borrador","Portal estudiantil","Formulario de homologación","No","","A02","","",1,""),
    ("A02","Adjuntar documentos y enviar","Actividad","Estudiante","Estudiante","Solicitud, certificado y syllabus","Solicitud enviada","Portal estudiantil","Anexos académicos","No","","A03","","",1,""),
    ("A03","Generar radicado","Actividad","Sistema","Sistema","Solicitud enviada","Solicitud radicada","Aplicativo de homologaciones","Número de radicado","No","","D01","","",0,"Automática"),
    ("D01","Verificar documentación","Decisión","Auxiliar académico","Facultad","Solicitud radicada","Resultado de revisión","Aplicativo de homologaciones","Lista de chequeo","Sí","¿La documentación está completa?","","A04","A05",2,""),
    ("A05","Devolver para corrección","Actividad","Auxiliar académico","Facultad","Observaciones de revisión","Solicitud devuelta","Aplicativo de homologaciones","Notificación de devolución","No","","D02","","",1,""),
    ("D02","Verificar corrección en plazo","Decisión","Sistema","Sistema","Solicitud devuelta","Resultado del plazo","Aplicativo de homologaciones","","Sí","¿El estudiante corrigió dentro del plazo?","","D01","A15",5,"Plazo ficticio"),
    ("A04","Validar documentos","Actividad","Auxiliar académico","Facultad","Documentación completa","Solicitud validada","Aplicativo de homologaciones","Lista de chequeo","No","","D03","","",1,""),
    ("D03","Verificar pago","Decisión","Sistema","Sistema","Solicitud validada","Estado financiero","Sistema financiero","","Sí","¿El pago está confirmado?","","A06","A07",0,""),
    ("A07","Notificar pago pendiente","Actividad","Sistema","Sistema","Pago no confirmado","Solicitud pendiente de pago","Correo institucional","Notificación de pago","No","","D03","","",0,"Bucle hasta pago o gestión manual"),
    ("A06","Asignar evaluador","Actividad","Director de programa","Facultad","Solicitud validada y pago confirmado","Evaluador asignado","Aplicativo de homologaciones","Asignación de evaluación","No","","A08","","",1,""),
    ("A08","Evaluar equivalencia","Actividad","Docente evaluador","Evaluación académica","Syllabus, notas y plan de estudios","Matriz de equivalencia","Aplicativo de homologaciones","Matriz de equivalencia","No","","D04","","",5,""),
    ("D04","Emitir concepto","Decisión","Docente evaluador","Evaluación académica","Matriz de equivalencia","Concepto académico","Aplicativo de homologaciones","Concepto académico","Sí","¿El concepto es favorable?","","A09","A10",1,""),
    ("A09","Aprobar solicitud","Actividad","Director o Comité Curricular","Decisión académica","Concepto favorable","Decisión aprobatoria","Gestor documental","Acta o decisión","No","","A11","","",2,""),
    ("A10","Rechazar solicitud","Actividad","Director o Comité Curricular","Decisión académica","Concepto desfavorable","Decisión de rechazo","Gestor documental","Acta o decisión","No","","A11","","",2,""),
    ("A11","Formalizar decisión","Actividad","Secretaría Académica","Secretaría Académica","Decisión académica","Acto académico","Gestor documental","Acto académico","No","","D05","","",1,""),
    ("D05","Determinar resultado","Decisión","Registro Académico","Registro Académico","Acto académico","Ruta de registro","Sistema académico","","Sí","¿La homologación fue aprobada?","","A12","A13",0,""),
    ("A12","Registrar asignaturas homologadas","Actividad","Registro Académico","Registro Académico","Decisión aprobatoria","Historial actualizado","Sistema académico","Registro de homologación","No","","A14","","",1,""),
    ("A13","Registrar rechazo","Actividad","Registro Académico","Registro Académico","Decisión de rechazo","Resultado registrado","Sistema académico","Registro de rechazo","No","","A14","","",1,""),
    ("A14","Notificar y archivar","Actividad","Sistema / Auxiliar académico","Cierre","Resultado registrado","Expediente cerrado","Correo y repositorio","Notificación y expediente","No","","A16","","",1,""),
    ("A15","Cerrar por vencimiento","Actividad","Sistema","Cierre","Plazo agotado","Caso cerrado por desistimiento","Aplicativo de homologaciones","Registro de cierre","No","","A16","","",0,""),
    ("A16","Fin","Fin","Sistema","Cierre","Expediente cerrado","Proceso finalizado","","","No","","","","",0,"Evento de fin"),
]

DEMO_RISKS = [
    ("R01","Solicitud incompleta","Ausencia de validaciones","Retrasos y reprocesos","Campos obligatorios y lista de chequeo","Sistema / Auxiliar académico"),
    ("R02","Evaluación inconsistente","Criterios no estandarizados","Decisiones diferentes","Matriz de equivalencia","Director de programa"),
    ("R03","Pérdida de documentos","Almacenamiento fragmentado","Falta de evidencia","Expediente electrónico centralizado","Secretaría Académica"),
    ("R04","Incumplimiento de tiempos","Falta de seguimiento","Insatisfacción del estudiante","Alertas y semáforos","Sistema"),
    ("R05","Registro incorrecto","Error humano","Historial inconsistente","Validación previa al registro","Registro Académico"),
]
DEMO_INDICATORS = [
    ("I01","Tiempo promedio de respuesta","Total de días de atención / solicitudes cerradas","Días","Mensual","<= 15"),
    ("I02","Solicitudes atendidas a tiempo","Solicitudes dentro del plazo / solicitudes cerradas * 100","Porcentaje","Mensual",">= 90%"),
    ("I03","Solicitudes devueltas","Solicitudes devueltas / solicitudes recibidas * 100","Porcentaje","Mensual","<= 20%"),
    ("I04","Índice de reproceso","Solicitudes revisadas más de una vez / solicitudes recibidas * 100","Porcentaje","Mensual","<= 15%"),
]


def demo_data() -> Dict[str, object]:
    return {
        "meta": dict(DEMO_META),
        "caracterizacion": pd.DataFrame(DEMO_CARACTERIZACION, columns=["Campo", "Descripción"]),
        "sipoc": pd.DataFrame(DEMO_SIPOC, columns=["ID", "Proveedor (S)", "Entrada (I)", "Etapa del proceso (P)", "Salida (O)", "Cliente (C)", "Orden_etapa"]),
        "actividades": pd.DataFrame(DEMO_ACTIVITIES, columns=EXPECTED_ACT_COLS),
        "riesgos": pd.DataFrame(DEMO_RISKS, columns=["id_riesgo","riesgo","causa","consecuencia","control","responsable"]),
        "indicadores": pd.DataFrame(DEMO_INDICATORS, columns=["id","indicador","formula_descriptiva","unidad","frecuencia","meta"]),
    }


def read_excel(source) -> Dict[str, pd.DataFrame]:
    book = pd.ExcelFile(source, engine="openpyxl")
    return {name: pd.read_excel(book, sheet_name=name) for name in book.sheet_names}


def normalize(frame: pd.DataFrame) -> pd.DataFrame:
    result = frame.copy()
    result.columns = [clean_col(c) for c in result.columns]
    return result.dropna(how="all").reset_index(drop=True)


def load_data(sipoc_file, ficha_file) -> Tuple[Dict[str, object], List[str]]:
    data = demo_data()
    notes: List[str] = []
    try:
        if sipoc_file is not None:
            sheets = read_excel(sipoc_file)
            by_key = {clean_col(k): v for k, v in sheets.items()}
            if "sipoc" in by_key:
                data["sipoc"] = by_key["sipoc"].dropna(how="all")
            if "metadatos" in by_key:
                md = normalize(by_key["metadatos"])
                if {"campo", "valor"}.issubset(md.columns):
                    data["meta"].update(dict(zip(md["campo"].astype(str), md["valor"])))
        if ficha_file is not None:
            sheets = read_excel(ficha_file)
            by_key = {clean_col(k): v for k, v in sheets.items()}
            if "caracterizacion" in by_key:
                data["caracterizacion"] = by_key["caracterizacion"].dropna(how="all")
            if "actividades" in by_key:
                acts = normalize(by_key["actividades"])
                for col in EXPECTED_ACT_COLS:
                    if col not in acts.columns:
                        acts[col] = ""
                data["actividades"] = acts[EXPECTED_ACT_COLS]
            if "riesgos_controles" in by_key:
                data["riesgos"] = normalize(by_key["riesgos_controles"])
            if "indicadores" in by_key:
                data["indicadores"] = normalize(by_key["indicadores"])
    except Exception as exc:
        notes.append(f"No fue posible leer uno de los archivos: {exc}. Se muestran los datos de demostración incorporados.")
    return data, notes


def escape_dot(value: object) -> str:
    return text(value).replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ")


def wrap_label(value: object, width: int = 26) -> str:
    words = text(value).split()
    lines, current = [], []
    for word in words:
        if len(" ".join(current + [word])) > width and current:
            lines.append(" ".join(current)); current = [word]
        else:
            current.append(word)
    if current: lines.append(" ".join(current))
    return "\\n".join(lines)


def build_dot(df: pd.DataFrame, orientation: str, show_ids: bool, show_time: bool,
              show_system: bool, lane_filter: Iterable[str]) -> str:
    acts = df.copy().fillna("")
    acts["id"] = acts["id"].map(text)
    acts = acts[acts["id"] != ""]
    allowed = set(lane_filter)
    if allowed:
        acts = acts[acts["carril"].map(text).isin(allowed)]
    visible_ids = set(acts["id"])
    rankdir = "LR" if orientation == "Horizontal" else "TB"
    lines = [
        "digraph proceso {", f"rankdir={rankdir};", "compound=true;", "newrank=true;", "splines=ortho;",
        'graph [bgcolor="transparent", pad="0.35", nodesep="0.55", ranksep="0.75", fontname="Arial"];',
        'node [fontname="Arial", fontsize=10, color="#35516f", penwidth=1.25, style="filled", fillcolor="#ffffff", margin="0.14,0.09"];',
        'edge [fontname="Arial", fontsize=9, color="#526b84", arrowsize=0.72, penwidth=1.15];'
    ]
    lanes = [text(x) or "Sin carril" for x in acts["carril"].drop_duplicates().tolist()]
    for idx, lane in enumerate(lanes):
        lane_rows = acts[acts["carril"].map(lambda x: text(x) or "Sin carril") == lane]
        lines += [f"subgraph cluster_{idx} {{", f'label="{escape_dot(lane)}";',
                  'color="#b8c7d9"; fillcolor="#f3f7fb"; style="rounded,filled"; penwidth=1.1;',
                  'fontname="Arial Bold"; fontsize=11; fontcolor="#17365d";']
        for _, row in lane_rows.iterrows():
            nid = safe_id(row["id"])
            tipo = text(row["tipo_elemento"]).lower()
            label_parts = []
            if show_ids: label_parts.append(f"[{text(row['id'])}]")
            label_parts.append(wrap_label(row["pregunta_decision"] if "decisi" in tipo and text(row["pregunta_decision"]) else row["nombre"]))
            if show_time and text(row["tiempo_dias"]): label_parts.append(f"⏱ {text(row['tiempo_dias'])} día(s)")
            if show_system and text(row["sistema"]): label_parts.append(wrap_label(f"Sistema: {text(row['sistema'])}", 30))
            label = "\\n".join(label_parts)
            if tipo == "inicio":
                attrs = 'shape=circle, width=0.58, fixedsize=true, fillcolor="#d9f2e6", color="#29845a", penwidth=2.0'
            elif tipo == "fin":
                attrs = 'shape=doublecircle, width=0.62, fixedsize=true, fillcolor="#fde3e3", color="#b83a3a", penwidth=2.0'
            elif "decisi" in tipo:
                attrs = 'shape=diamond, fillcolor="#fff2cc", color="#bf9000", width=1.45, height=0.92'
            else:
                attrs = 'shape=box, style="rounded,filled", fillcolor="#ffffff", color="#35516f"'
            tooltip = escape_dot(f"Responsable: {text(row['responsable'])} | Entrada: {text(row['entrada'])} | Salida: {text(row['salida'])}")
            lines.append(f'{nid} [label="{escape_dot(label)}", {attrs}, tooltip="{tooltip}"];')
        lines.append("}")
    seen_edges = set()
    for _, row in acts.iterrows():
        src = row["id"]
        candidates = []
        if boolish(row["es_decision"]):
            candidates = [(row["siguiente_si"], "Sí", "#238636"), (row["siguiente_no"], "No", "#c43d3d")]
            if text(row["siguiente"]): candidates.append((row["siguiente"], "", "#526b84"))
        else:
            candidates = [(row["siguiente"], "", "#526b84")]
        for target, lbl, color in candidates:
            target = text(target)
            key = (src, target, lbl)
            if target and target in visible_ids and key not in seen_edges:
                seen_edges.add(key)
                lines.append(f'{safe_id(src)} -> {safe_id(target)} [label="{lbl}", color="{color}", fontcolor="{color}"];')
    lines.append("}")
    return "\n".join(lines)


def quality_checks(df: pd.DataFrame) -> pd.DataFrame:
    acts = df.fillna("").copy()
    ids = [text(v) for v in acts["id"] if text(v)]
    idset = set(ids)
    issues = []
    duplicates = sorted({x for x in ids if ids.count(x) > 1})
    if duplicates: issues.append(("Error", "IDs duplicados", ", ".join(duplicates)))
    for _, row in acts.iterrows():
        rid = text(row["id"])
        if not rid: continue
        if boolish(row["es_decision"]):
            if not text(row["siguiente_si"]) or not text(row["siguiente_no"]):
                issues.append(("Advertencia", rid, "La decisión no tiene completas las rutas Sí/No."))
        targets = [text(row[c]) for c in ["siguiente", "siguiente_si", "siguiente_no"] if text(row[c])]
        for target in targets:
            if target not in idset:
                issues.append(("Error", rid, f"La conexión apunta a un ID inexistente: {target}"))
    if not any(text(x).lower() == "inicio" for x in acts["tipo_elemento"]):
        issues.append(("Error", "Modelo", "No se encontró un evento de inicio."))
    if not any(text(x).lower() == "fin" for x in acts["tipo_elemento"]):
        issues.append(("Error", "Modelo", "No se encontró un evento de fin."))
    if not issues: issues.append(("Correcto", "Modelo", "No se detectaron inconsistencias estructurales básicas."))
    return pd.DataFrame(issues, columns=["Nivel", "Elemento", "Detalle"])


# --- Entrada y estado ---------------------------------------------------------
with st.sidebar:
    st.markdown("## Modelador académico")
    st.caption("Importe los libros o trabaje con el caso de demostración incorporado.")
    sipoc_upload = st.file_uploader("Archivo SIPOC", type=["xlsx"], key="sipoc")
    ficha_upload = st.file_uploader("Ficha de caracterización", type=["xlsx"], key="ficha")
    st.divider()
    orientation = st.radio("Orientación", ["Horizontal", "Vertical"], horizontal=True)
    show_ids = st.checkbox("Mostrar ID", value=True)
    show_time = st.checkbox("Mostrar duración", value=True)
    show_system = st.checkbox("Mostrar sistema", value=False)

source_sipoc = sipoc_upload
source_ficha = ficha_upload
# En ejecución local, también toma los libros si están disponibles junto al script.
if source_sipoc is None and Path("01_SIPOC_Homologacion.xlsx").exists(): source_sipoc = "01_SIPOC_Homologacion.xlsx"
if source_ficha is None and Path("02_Ficha_Caracterizacion_Homologacion.xlsx").exists(): source_ficha = "02_Ficha_Caracterizacion_Homologacion.xlsx"

data, notes = load_data(source_sipoc, source_ficha)
for note in notes: st.warning(note)

st.markdown("""
<div class="hero"><h1>Modelador de procesos</h1>
<p>Diagrama por carriles, SIPOC, caracterización, riesgos e indicadores en una sola vista.</p></div>
""", unsafe_allow_html=True)

meta = data["meta"]
st.markdown(
    f"""<div class="proc-card"><b>{text(meta.get('codigo_proceso'))} · {text(meta.get('nombre_proceso'))}</b><br>
    <span class="small-note">{text(meta.get('institucion'))} · {text(meta.get('macroproceso'))} · {text(meta.get('tipo_proceso'))} · Versión {text(meta.get('version'))}</span></div>""",
    unsafe_allow_html=True,
)

acts_initial = data["actividades"].copy()
if "activities_editor" not in st.session_state:
    st.session_state.activities_editor = acts_initial

c1, c2, c3, c4 = st.columns(4)
c1.metric("Elementos", len(st.session_state.activities_editor))
c2.metric("Decisiones", int(st.session_state.activities_editor["es_decision"].map(boolish).sum()))
c3.metric("Carriles", st.session_state.activities_editor["carril"].map(text).replace("", pd.NA).nunique())
try:
    total_days = pd.to_numeric(st.session_state.activities_editor["tiempo_dias"], errors="coerce").fillna(0).sum()
    c4.metric("Duración referencial", f"{total_days:g} días")
except Exception:
    c4.metric("Duración referencial", "N/D")

available_lanes = [text(x) for x in st.session_state.activities_editor["carril"].drop_duplicates() if text(x)]
with st.sidebar:
    lane_filter = st.multiselect("Carriles visibles", available_lanes, default=available_lanes)
    st.caption("El filtro oculta también las conexiones hacia elementos no visibles.")

tab_diagram, tab_edit, tab_sipoc, tab_char, tab_risk, tab_kpi, tab_quality = st.tabs([
    "Diagrama", "Editar flujo", "SIPOC", "Caracterización", "Riesgos y controles", "Indicadores", "Calidad"
])

with tab_diagram:
    dot = build_dot(st.session_state.activities_editor, orientation, show_ids, show_time, show_system, lane_filter)
    st.graphviz_chart(dot, use_container_width=True)
    d1, d2 = st.columns([1, 1])
    d1.download_button("Descargar modelo DOT", dot.encode("utf-8"), "modelo_proceso.dot", "text/vnd.graphviz", use_container_width=True)
    model_json = st.session_state.activities_editor.fillna("").to_dict(orient="records")
    d2.download_button("Descargar datos JSON", json.dumps(model_json, ensure_ascii=False, indent=2).encode("utf-8"), "modelo_proceso.json", "application/json", use_container_width=True)
    st.caption("Convención: verde = inicio, rectángulo = actividad, rombo = decisión y rojo = fin. Pase el cursor sobre un elemento para consultar responsable, entrada y salida.")

with tab_edit:
    st.info("Puede editar, agregar o eliminar filas. Para actualizar el diagrama, pulse Aplicar cambios.")
    edited = st.data_editor(
        st.session_state.activities_editor,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "es_decision": st.column_config.SelectboxColumn("es_decision", options=["Sí", "No"]),
            "tiempo_dias": st.column_config.NumberColumn("tiempo_dias", min_value=0.0, step=1.0),
        },
        key="activity_grid",
    )
    b1, b2, b3 = st.columns(3)
    if b1.button("Aplicar cambios", type="primary", use_container_width=True):
        st.session_state.activities_editor = edited.copy()
        st.rerun()
    if b2.button("Restablecer", use_container_width=True):
        st.session_state.activities_editor = acts_initial.copy()
        st.rerun()
    activity_xlsx = excel_bytes({"Actividades": edited})
    b3.download_button("Descargar actividades", activity_xlsx, "actividades_editadas.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

with tab_sipoc:
    sipoc_frame = data["sipoc"]
    st.dataframe(sipoc_frame, use_container_width=True, hide_index=True)
    st.download_button("Descargar SIPOC", excel_bytes({"SIPOC": sipoc_frame}), "sipoc_exportado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with tab_char:
    char_frame = data["caracterizacion"]
    st.dataframe(char_frame, use_container_width=True, hide_index=True)

with tab_risk:
    risk_frame = data["riesgos"]
    st.dataframe(risk_frame, use_container_width=True, hide_index=True)

with tab_kpi:
    kpi_frame = data["indicadores"]
    st.dataframe(kpi_frame, use_container_width=True, hide_index=True)

with tab_quality:
    checks = quality_checks(st.session_state.activities_editor)
    st.dataframe(checks, use_container_width=True, hide_index=True)
    st.caption("La revisión comprueba IDs, referencias de conexión, rutas de decisión y presencia de eventos de inicio y fin.")

with st.expander("Estructura esperada de los archivos"):
    st.markdown("""
    **SIPOC:** hojas `SIPOC` y `Metadatos`.  
    **Ficha:** hojas `Caracterizacion`, `Actividades`, `Riesgos_Controles` e `Indicadores`.  
    En `Actividades`, las conexiones se forman con `siguiente`, `siguiente_si` y `siguiente_no` usando los valores de la columna `id`.
    """)
