import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta
import random
import os
import json
 
# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="FLOW — Ford Logistics",
    page_icon="🚛",
    layout="wide",
    initial_sidebar_state="collapsed"
)
 
EXCEL_PATH = "flow_registro.xlsx"
RECEPCIONES = ["GENERAL", "INTERPLANTA", "JIT", "RANGER"]
DARSENAS = {"GENERAL": 5, "INTERPLANTA": 2, "JIT": 2, "RANGER": 4}
 
# ─────────────────────────────────────────────
# ESTILOS
# ─────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .main { background-color: #ffffff; }
    .block-container { padding-top: 1rem; padding-bottom: 1rem; }
 
    /* Header */
    .flow-header {
        background: #ffffff;
        border-bottom: 2px solid #003478;
        padding: 12px 24px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 16px;
    }
    .flow-title {
        font-size: 26px;
        font-weight: 700;
        color: #003478;
        letter-spacing: -0.5px;
    }
    .flow-subtitle {
        font-size: 11px;
        color: #6b7280;
        letter-spacing: 0.1em;
        text-transform: uppercase;
    }
 
    /* Cards */
    .kpi-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 16px 20px;
        text-align: center;
    }
    .kpi-value { font-size: 28px; font-weight: 700; color: #003478; }
    .kpi-label { font-size: 11px; color: #6b7280; text-transform: uppercase; letter-spacing: 0.08em; margin-top: 4px; }
 
    /* Camión cards */
    .truck-critico {
        background: #fff5f5;
        border: 2px solid #ef4444;
        border-left: 6px solid #ef4444;
        border-radius: 10px;
        padding: 14px 16px;
        margin-bottom: 10px;
    }
    .truck-normal {
        background: #f0fdf4;
        border: 1px solid #bbf7d0;
        border-left: 6px solid #22c55e;
        border-radius: 10px;
        padding: 14px 16px;
        margin-bottom: 10px;
    }
    .truck-atrasado {
        background: #fffbeb;
        border: 1px solid #fde68a;
        border-left: 6px solid #f59e0b;
        border-radius: 10px;
        padding: 14px 16px;
        margin-bottom: 10px;
    }
    .truck-descargado {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-left: 6px solid #94a3b8;
        border-radius: 10px;
        padding: 14px 16px;
        margin-bottom: 10px;
        opacity: 0.6;
    }
 
    /* Priority badge */
    .badge-red { background:#fee2e2; color:#b91c1c; padding:3px 10px; border-radius:20px; font-size:11px; font-weight:600; }
    .badge-green { background:#dcfce7; color:#15803d; padding:3px 10px; border-radius:20px; font-size:11px; font-weight:600; }
    .badge-yellow { background:#fef9c3; color:#a16207; padding:3px 10px; border-radius:20px; font-size:11px; font-weight:600; }
    .badge-blue { background:#dbeafe; color:#1d4ed8; padding:3px 10px; border-radius:20px; font-size:11px; font-weight:600; }
    .badge-gray { background:#f1f5f9; color:#64748b; padding:3px 10px; border-radius:20px; font-size:11px; font-weight:600; }
 
    /* Botones custom */
    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
        font-size: 13px;
        padding: 8px 16px;
        transition: all 0.2s;
    }
 
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] { gap: 6px; border-bottom: 2px solid #e2e8f0; }
    .stTabs [data-baseweb="tab"] { 
        background: #f8fafc; border-radius: 8px 8px 0 0; 
        color: #64748b; font-weight: 500; font-size: 13px;
        padding: 8px 18px; border: 1px solid #e2e8f0;
    }
    .stTabs [aria-selected="true"] {
        background: #003478 !important; color: white !important;
        border-color: #003478 !important;
    }
 
    /* Section titles */
    .section-title {
        font-size: 13px; font-weight: 600; color: #374151;
        text-transform: uppercase; letter-spacing: 0.08em;
        border-bottom: 1px solid #e2e8f0; padding-bottom: 6px; margin-bottom: 12px;
    }
 
    /* Alert box */
    .alert-box {
        background: #fff5f5; border: 1px solid #fecaca;
        border-radius: 8px; padding: 10px 14px;
        color: #b91c1c; font-size: 12px; margin-bottom: 8px;
    }
    .info-box {
        background: #eff6ff; border: 1px solid #bfdbfe;
        border-radius: 8px; padding: 10px 14px;
        color: #1d4ed8; font-size: 12px; margin-bottom: 8px;
    }
 
    div[data-testid="stMetricValue"] { color: #003478 !important; }
</style>
""", unsafe_allow_html=True)
 
# ─────────────────────────────────────────────
# DATOS INICIALES
# ─────────────────────────────────────────────
def init_data():
    if "camiones" not in st.session_state:
        camiones = [
            # GENERAL
            {"id": "CAM-001", "patente": "AB 123 CD", "codigo": "GEN-001", "recepcion": "GENERAL",
             "proveedor": "Autoparts SA", "tipo": "LOTE", "material": "Paragolpes trasero",
             "ventana": "08:30", "llegada": "08:42", "estado": "Crítico",
             "criticidad_smart": 9, "punto_pedido": 50, "stock_actual": 8,
             "cantidad_lote": 120, "remito": "REM-2026-001", "darsena": 3,
             "descargado": False, "hora_descarga": None, "km": 45},
            {"id": "CAM-002", "patente": "EF 456 GH", "codigo": "GEN-002", "recepcion": "GENERAL",
             "proveedor": "Metalmec SRL", "tipo": "LOTE", "material": "Ejes delanteros",
             "ventana": "09:00", "llegada": "09:05", "estado": "Normal",
             "criticidad_smart": 5, "punto_pedido": 80, "stock_actual": 65,
             "cantidad_lote": 200, "remito": "REM-2026-002", "darsena": 1,
             "descargado": False, "hora_descarga": None, "km": 120},
            {"id": "CAM-003", "patente": "IJ 789 KL", "codigo": "GEN-003", "recepcion": "GENERAL",
             "proveedor": "Plastiford SA", "tipo": "LOTE", "material": "Molduras interiores",
             "ventana": "09:30", "llegada": "10:05", "estado": "Atrasado",
             "criticidad_smart": 6, "punto_pedido": 60, "stock_actual": 40,
             "cantidad_lote": 150, "remito": "REM-2026-003", "darsena": 2,
             "descargado": False, "hora_descarga": None, "km": 80},
            # INTERPLANTA
            {"id": "CAM-004", "patente": "MN 321 OP", "codigo": "INT-001", "recepcion": "INTERPLANTA",
             "proveedor": "Depósito Norte", "tipo": "LOTE", "material": "Varios (interplanta)",
             "ventana": "09:15", "llegada": "09:20", "estado": "Normal",
             "criticidad_smart": 6, "punto_pedido": 40, "stock_actual": 25,
             "cantidad_lote": 90, "remito": "REM-2026-004", "darsena": 1,
             "descargado": False, "hora_descarga": None, "km": 30},
            {"id": "CAM-005", "patente": "QR 654 ST", "codigo": "INT-002", "recepcion": "INTERPLANTA",
             "proveedor": "Depósito Sur", "tipo": "LOTE", "material": "Componentes motor",
             "ventana": "10:00", "llegada": "09:55", "estado": "Normal",
             "criticidad_smart": 7, "punto_pedido": 55, "stock_actual": 30,
             "cantidad_lote": 110, "remito": "REM-2026-005", "darsena": 2,
             "descargado": False, "hora_descarga": None, "km": 55},
            # JIT
            {"id": "CAM-006", "patente": "UV 987 WX", "codigo": "JIT-001", "recepcion": "JIT",
             "proveedor": "Secuencia Express", "tipo": "SECUENCIADO", "material": "Asientos delanteros",
             "ventana": "09:20", "llegada": "09:22", "estado": "Crítico",
             "criticidad_smart": 10, "punto_pedido": 70, "stock_actual": 12,
             "cantidad_lote": 70, "remito": "REM-2026-006", "darsena": 1,
             "modulos_disp": 0.17, "tiempo_reposicion_hs": 1,
             "descargado": False, "hora_descarga": None, "km": 95},
            {"id": "CAM-007", "patente": "YZ 111 AB", "codigo": "JIT-002", "recepcion": "JIT",
             "proveedor": "AutoSec SRL", "tipo": "SECUENCIADO", "material": "Tableros completos",
             "ventana": "10:30", "llegada": None, "estado": "Normal",
             "criticidad_smart": 8, "punto_pedido": 70, "stock_actual": 55,
             "cantidad_lote": 70, "remito": None, "darsena": None,
             "modulos_disp": 0.79, "tiempo_reposicion_hs": 1,
             "descargado": False, "hora_descarga": None, "km": 70},
            # RANGER
            {"id": "CAM-008", "patente": "CD 222 EF", "codigo": "RAN-001", "recepcion": "RANGER",
             "proveedor": "Ranger Parts", "tipo": "LOTE", "material": "Chasis Ranger",
             "ventana": "09:00", "llegada": "08:58", "estado": "Normal",
             "criticidad_smart": 7, "punto_pedido": 30, "stock_actual": 22,
             "cantidad_lote": 80, "remito": "REM-2026-008", "darsena": 2,
             "descargado": False, "hora_descarga": None, "km": 110},
            {"id": "CAM-009", "patente": "GH 333 IJ", "codigo": "RAN-002", "recepcion": "RANGER",
             "proveedor": "Motores del Sur", "tipo": "LOTE", "material": "Motores Ranger V6",
             "ventana": "10:15", "llegada": None, "estado": "Normal",
             "criticidad_smart": 8, "punto_pedido": 20, "stock_actual": 18,
             "cantidad_lote": 60, "remito": None, "darsena": None,
             "descargado": False, "hora_descarga": None, "km": 90},
        ]
        st.session_state.camiones = camiones
 
    if "registro" not in st.session_state:
        st.session_state.registro = []
 
    if "descarga_extra" not in st.session_state:
        st.session_state.descarga_extra = None
 
    if "confirmaciones_qr" not in st.session_state:
        st.session_state.confirmaciones_qr = set()
 
    if "ventanas_plan" not in st.session_state:
        st.session_state.ventanas_plan = [
            {"hora": "08:00", "recepcion": "GENERAL", "proveedor": "Autoparts SA", "tipo": "LOTE", "estado_ventana": "Cumplida"},
            {"hora": "08:30", "recepcion": "GENERAL", "proveedor": "Metalmec SRL", "tipo": "LOTE", "estado_ventana": "Cumplida"},
            {"hora": "09:00", "recepcion": "JIT", "proveedor": "Secuencia Express", "tipo": "SECUENCIADO", "estado_ventana": "En curso"},
            {"hora": "09:15", "recepcion": "INTERPLANTA", "proveedor": "Depósito Norte", "tipo": "LOTE", "estado_ventana": "En curso"},
            {"hora": "09:30", "recepcion": "GENERAL", "proveedor": "Plastiford SA", "tipo": "LOTE", "estado_ventana": "Demorada"},
            {"hora": "10:00", "recepcion": "RANGER", "proveedor": "Ranger Parts", "tipo": "LOTE", "estado_ventana": "Pendiente"},
            {"hora": "10:30", "recepcion": "JIT", "proveedor": "AutoSec SRL", "tipo": "SECUENCIADO", "estado_ventana": "Pendiente"},
            {"hora": "11:00", "recepcion": "INTERPLANTA", "proveedor": "Depósito Sur", "tipo": "LOTE", "estado_ventana": "Pendiente"},
            {"hora": "11:30", "recepcion": "RANGER", "proveedor": "Motores del Sur", "tipo": "LOTE", "estado_ventana": "Pendiente"},
            {"hora": "12:00", "recepcion": "GENERAL", "proveedor": "Autoparts SA", "tipo": "LOTE", "estado_ventana": "Pendiente"},
        ]
 
# ─────────────────────────────────────────────
# CÁLCULO IPD
# ─────────────────────────────────────────────
def calcular_ipd(c):
    score = 0
 
    # 1. Tipo: secuenciado tiene prioridad base
    if c["tipo"] == "SECUENCIADO":
        score += 40
 
    # 2. Criticidad SMART (1-10) → hasta 30 puntos
    score += c.get("criticidad_smart", 5) * 3
 
    # 3. Stock vs punto de pedido
    stock = c.get("stock_actual", 50)
    pp = c.get("punto_pedido", 50)
    ratio = stock / pp if pp > 0 else 1
    if ratio <= 0.2:
        score += 25
    elif ratio <= 0.5:
        score += 15
    elif ratio <= 0.8:
        score += 8
    else:
        score += 0
 
    # 4. Módulos secuenciados disponibles
    if c["tipo"] == "SECUENCIADO":
        modulos = c.get("modulos_disp", 1.0)
        if modulos <= 0.2:
            score += 20
        elif modulos <= 0.6:
            score += 10
        else:
            score += 5
 
    # 5. Desfasaje de ventana (llegó pero no descargó)
    if c.get("llegada") and not c.get("descargado"):
        try:
            llegada = datetime.strptime(c["llegada"], "%H:%M")
            ventana = datetime.strptime(c["ventana"], "%H:%M")
            diff = (datetime.now().replace(hour=datetime.now().hour, minute=datetime.now().minute) - llegada).total_seconds() / 60
            if diff > 30:
                score += 10
            elif diff > 15:
                score += 5
        except:
            pass
 
    # 6. Camión llegado suma sobre no llegado
    if c.get("llegada"):
        score += 5
 
    return min(score, 100)
 
def get_estado_color(estado):
    if estado == "Crítico": return "truck-critico"
    if estado == "Atrasado": return "truck-atrasado"
    if estado == "Descargado": return "truck-descargado"
    return "truck-normal"
 
def get_badge(estado):
    if estado == "Crítico": return "badge-red", "🔴 CRÍTICO"
    if estado == "Atrasado": return "badge-yellow", "🟡 ATRASADO"
    if estado == "Descargado": return "badge-gray", "✅ DESCARGADO"
    return "badge-green", "🟢 NORMAL"
 
def get_ipd_color(ipd):
    if ipd >= 75: return "#ef4444"
    if ipd >= 50: return "#f59e0b"
    if ipd >= 30: return "#3b82f6"
    return "#22c55e"
 
# ─────────────────────────────────────────────
# GUARDAR EXCEL
# ─────────────────────────────────────────────
def guardar_excel():
    try:
        df = pd.DataFrame(st.session_state.registro)
        df.to_excel(EXCEL_PATH, index=False)
    except Exception as e:
        pass
 
def registrar_accion(cam_id, accion, detalle=""):
    st.session_state.registro.append({
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "camion_id": cam_id,
        "accion": accion,
        "detalle": detalle,
        "usuario": "Operario"
    })
    guardar_excel()
 
# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
def render_header():
    col1, col2 = st.columns([3, 1])
    with col1:
        st.markdown("""
        <div style="padding: 8px 0 16px 0; border-bottom: 2px solid #003478; margin-bottom: 20px;">
            <div style="font-size:28px; font-weight:800; color:#003478; letter-spacing:-0.5px;">
                FLOW
            </div>
            <div style="font-size:11px; color:#6b7280; letter-spacing:0.12em; text-transform:uppercase; margin-top:2px;">
                Ford Logistics Operations Window
            </div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        now = datetime.now()
        st.markdown(f"""
        <div style="text-align:right; padding: 8px 0 16px 0; border-bottom: 2px solid #003478; margin-bottom: 20px;">
            <div style="font-size:22px; font-weight:700; color:#003478;">{now.strftime('%H:%M:%S')}</div>
            <div style="font-size:11px; color:#6b7280;">{now.strftime('%d / %m / %Y')}</div>
            <div style="font-size:11px; color:#003478; font-weight:600; margin-top:4px;">🏭 FORD PACHECO</div>
        </div>
        """, unsafe_allow_html=True)
 
# ─────────────────────────────────────────────
# RENDER CAMIÓN CARD
# ─────────────────────────────────────────────
def render_truck_card(c, show_buttons=True, preview_mode=False):
    ipd = calcular_ipd(c)
    estado = "Descargado" if c["descargado"] else c["estado"]
    css_class = get_estado_color(estado)
    badge_class, badge_text = get_badge(estado)
    ipd_color = get_ipd_color(ipd)
 
    llegada_str = c["llegada"] if c["llegada"] else "—"
    darsena_str = f"Dársena {c['darsena']}" if c.get("darsena") else "No asignada"
    remito_str = c["remito"] if c.get("remito") else "Pendiente"
 
    tipo_badge = "badge-blue" if c["tipo"] == "SECUENCIADO" else "badge-gray"
 
    st.markdown(f"""
    <div class="{css_class}">
        <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:8px;">
            <div>
                <span style="font-size:15px; font-weight:700; color:#111827;">{c['id']}</span>
                <span style="margin-left:8px; font-size:12px; color:#6b7280;">{c['patente']}</span>
                <span class="{badge_class}" style="margin-left:8px;">{badge_text}</span>
                <span class="{tipo_badge}" style="margin-left:6px;">{c['tipo']}</span>
            </div>
            <div style="text-align:right;">
                <div style="font-size:22px; font-weight:800; color:{ipd_color};">{ipd}</div>
                <div style="font-size:9px; color:#9ca3af; text-transform:uppercase;">IPD</div>
            </div>
        </div>
        <div style="display:grid; grid-template-columns: 1fr 1fr 1fr 1fr; gap:8px; font-size:12px;">
            <div><span style="color:#9ca3af;">Proveedor</span><br><b>{c['proveedor']}</b></div>
            <div><span style="color:#9ca3af;">Material</span><br><b>{c['material']}</b></div>
            <div><span style="color:#9ca3af;">Ventana</span><br><b>{c['ventana']}</b></div>
            <div><span style="color:#9ca3af;">Llegada</span><br><b>{llegada_str}</b></div>
            <div><span style="color:#9ca3af;">Remito</span><br><b>{remito_str}</b></div>
            <div><span style="color:#9ca3af;">Dársena</span><br><b>{darsena_str}</b></div>
            <div><span style="color:#9ca3af;">Stock actual</span><br><b>{c.get('stock_actual','—')} u.</b></div>
            <div><span style="color:#9ca3af;">Punto pedido</span><br><b>{c.get('punto_pedido','—')} u.</b></div>
        </div>
        {"<div style='margin-top:8px; padding:6px 10px; background:#fee2e2; border-radius:6px; font-size:11px; color:#b91c1c;'>⚠️ MATERIAL CRÍTICO — Riesgo de parada de línea</div>" if estado == "Crítico" else ""}
        {"<div style='margin-top:8px; padding:6px 10px; background:#fef9c3; border-radius:6px; font-size:11px; color:#a16207;'>⏱ CAMIÓN DEMORADO — Llegó fuera de ventana</div>" if estado == "Atrasado" else ""}
        {"<div style='margin-top:8px; padding:6px 10px; background:#dcfce7; border-radius:6px; font-size:11px; color:#15803d;'>✅ Descargado a las " + str(c.get('hora_descarga','—')) + "</div>" if c['descargado'] else ""}
    </div>
    """, unsafe_allow_html=True)
 
    if show_buttons and not c["descargado"] and not preview_mode:
        col_a, col_b, col_c = st.columns([1, 1, 2])
        with col_a:
            if st.button(f"✅ Descargado", key=f"desc_{c['id']}", type="primary"):
                hora_ahora = datetime.now().strftime("%H:%M")
                for cam in st.session_state.camiones:
                    if cam["id"] == c["id"]:
                        cam["descargado"] = True
                        cam["hora_descarga"] = hora_ahora
                        cam["estado"] = "Descargado"
                registrar_accion(c["id"], "DESCARGA_CONFIRMADA", f"Dársena {c.get('darsena','?')} a las {hora_ahora}")
                st.rerun()
        with col_b:
            if st.button(f"🔄 Desc. Extraordinaria", key=f"extra_{c['id']}"):
                st.session_state.descarga_extra = c["id"]
                st.rerun()
        with col_c:
            if not c.get("llegada"):
                remito_input = st.text_input("Confirmar llegada (Remito/QR)", key=f"qr_{c['id']}", placeholder="Ingresá el nro. de remito")
                if remito_input and remito_input == c.get("remito", ""):
                    for cam in st.session_state.camiones:
                        if cam["id"] == c["id"]:
                            cam["llegada"] = datetime.now().strftime("%H:%M")
                    registrar_accion(c["id"], "CONFIRMACION_LLEGADA", f"Remito {remito_input}")
                    st.success("✅ Llegada confirmada")
                    st.rerun()
 
# ─────────────────────────────────────────────
# KPIs POR RECEPCIÓN
# ─────────────────────────────────────────────
def render_kpis(recepcion, camiones_rec):
    total_darsenas = DARSENAS[recepcion]
    ocupadas = len([c for c in camiones_rec if c.get("darsena") and not c["descargado"]])
    pct_ocupacion = round(ocupadas / total_darsenas * 100) if total_darsenas > 0 else 0
 
    llegados = [c for c in camiones_rec if c.get("llegada")]
    en_ventana = 0
    for c in llegados:
        try:
            ll = datetime.strptime(c["llegada"], "%H:%M")
            vt = datetime.strptime(c["ventana"], "%H:%M")
            if abs((ll - vt).total_seconds()) <= 900:
                en_ventana += 1
        except:
            pass
    pct_ventana = round(en_ventana / len(llegados) * 100) if llegados else 0
 
    pendientes = len([c for c in camiones_rec if not c.get("llegada")])
 
    tiempos_espera = []
    for c in camiones_rec:
        if c.get("llegada") and not c["descargado"]:
            try:
                ll = datetime.strptime(c["llegada"], "%H:%M")
                ahora = datetime.now().replace(second=0, microsecond=0)
                ll_hoy = ahora.replace(hour=ll.hour, minute=ll.minute)
                espera = (ahora - ll_hoy).total_seconds() / 60
                if espera >= 0:
                    tiempos_espera.append(espera)
            except:
                pass
    espera_prom = round(sum(tiempos_espera) / len(tiempos_espera)) if tiempos_espera else 0
 
    co2_ahorrado = round(len([c for c in camiones_rec if c["descargado"]]) * 0.027 * 15, 1)
 
    cols = st.columns(5)
    metrics = [
        ("⏱ Espera prom.", f"{espera_prom} min", None),
        ("🔲 Ocupación dársenas", f"{pct_ocupacion}%", f"{ocupadas}/{total_darsenas}"),
        ("✅ Cumplimiento ventanas", f"{pct_ventana}%", f"{en_ventana}/{len(llegados)} llegados"),
        ("🚛 Camiones pendientes", str(pendientes), "por llegar"),
        ("🌱 CO₂ ahorrado", f"{co2_ahorrado} kg", "estimado hoy"),
    ]
    for i, (label, val, delta) in enumerate(metrics):
        with cols[i]:
            st.metric(label=label, value=val, delta=delta)
 
# ─────────────────────────────────────────────
# RENDER RECEPCIÓN
# ─────────────────────────────────────────────
def render_recepcion(recepcion):
    camiones_rec = [c for c in st.session_state.camiones if c["recepcion"] == recepcion]
    camiones_rec_sorted = sorted(camiones_rec, key=lambda c: -calcular_ipd(c))
 
    st.markdown(f"""
    <div style="background:#003478; color:white; padding:10px 16px; border-radius:10px; margin-bottom:16px;">
        <span style="font-size:16px; font-weight:700;">📍 Recepción {recepcion}</span>
        <span style="margin-left:12px; font-size:12px; opacity:0.8;">{DARSENAS[recepcion]} dársenas disponibles</span>
    </div>
    """, unsafe_allow_html=True)
 
    render_kpis(recepcion, camiones_rec)
    st.markdown("---")
 
    # Descarga extraordinaria activa
    if st.session_state.descarga_extra:
        cam_extra = next((c for c in st.session_state.camiones if c["id"] == st.session_state.descarga_extra), None)
        if cam_extra and cam_extra["recepcion"] == recepcion:
            st.markdown("""
            <div style="background:#fff7ed; border:2px solid #f97316; border-radius:10px; padding:14px 16px; margin-bottom:16px;">
                <b style="color:#c2410c;">⚠️ MODO DESCARGA EXTRAORDINARIA</b><br>
                <span style="font-size:12px; color:#6b7280;">Vista previa del impacto. Confirmá o cancelá la operación.</span>
            </div>
            """, unsafe_allow_html=True)
            st.markdown("**Vista previa — Camión seleccionado para descarga extraordinaria:**")
            render_truck_card(cam_extra, show_buttons=False, preview_mode=True)
 
            # Simulación impacto
            st.markdown("**Impacto simulado en indicadores:**")
            col1, col2, col3 = st.columns(3)
            col1.metric("IPD promedio restante", f"{round(sum(calcular_ipd(c) for c in camiones_rec if c['id'] != cam_extra['id'] and not c['descargado']) / max(1, len([c for c in camiones_rec if c['id'] != cam_extra['id'] and not c['descargado']])), 1)}", "-2.3 vs plan")
            col2.metric("Riesgo parada línea", "BAJO" if cam_extra["estado"] != "Crítico" else "ALTO", None)
            col3.metric("Dársena liberada", f"#{cam_extra.get('darsena','?')}", "+1 disponible")
 
            colA, colB = st.columns(2)
            with colA:
                if st.button("✅ Confirmar descarga extraordinaria", key="confirm_extra", type="primary"):
                    hora_ahora = datetime.now().strftime("%H:%M")
                    for cam in st.session_state.camiones:
                        if cam["id"] == cam_extra["id"]:
                            cam["descargado"] = True
                            cam["hora_descarga"] = hora_ahora
                    registrar_accion(cam_extra["id"], "DESCARGA_EXTRAORDINARIA", f"Confirmada a las {hora_ahora}")
                    st.session_state.descarga_extra = None
                    st.rerun()
            with colB:
                if st.button("❌ Cancelar — Seguir flujo normal", key="cancel_extra"):
                    st.session_state.descarga_extra = None
                    st.rerun()
            st.markdown("---")
 
    # Lista de camiones
    activos = [c for c in camiones_rec_sorted if not c["descargado"]]
    descargados = [c for c in camiones_rec_sorted if c["descargado"]]
 
    if not activos:
        st.markdown("""
        <div class="info-box">✅ No hay camiones pendientes en esta recepción.</div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="section-title">Cola de descarga — {len(activos)} camiones activos</div>', unsafe_allow_html=True)
        for i, c in enumerate(activos):
            col_num, col_card = st.columns([0.08, 0.92])
            with col_num:
                ipd = calcular_ipd(c)
                color = get_ipd_color(ipd)
                st.markdown(f"""
                <div style="text-align:center; padding-top:12px;">
                    <div style="font-size:22px; font-weight:800; color:{color};">#{i+1}</div>
                </div>
                """, unsafe_allow_html=True)
            with col_card:
                render_truck_card(c)
 
    if descargados:
        with st.expander(f"✅ Historial de descargados ({len(descargados)})"):
            for c in descargados:
                render_truck_card(c, show_buttons=False)
 
# ─────────────────────────────────────────────
# TAB: TABLERO GENERAL
# ─────────────────────────────────────────────
def render_tablero_general():
    st.markdown('<div class="section-title">Vista consolidada — Todas las recepciones</div>', unsafe_allow_html=True)
 
    cols = st.columns(4)
    for i, rec in enumerate(RECEPCIONES):
        with cols[i]:
            camiones_rec = [c for c in st.session_state.camiones if c["recepcion"] == rec]
            activos = sorted([c for c in camiones_rec if not c["descargado"]], key=lambda c: -calcular_ipd(c))
 
            criticos = [c for c in activos if c["estado"] == "Crítico"]
            header_color = "#ef4444" if criticos else "#003478"
 
            st.markdown(f"""
            <div style="background:{header_color}; color:white; padding:8px 12px; border-radius:8px 8px 0 0; text-align:center; font-weight:700; font-size:14px;">
                {rec}
                <br><span style="font-size:11px; opacity:0.8;">{len(activos)} activos · {DARSENAS[rec]} dársenas</span>
            </div>
            """, unsafe_allow_html=True)
 
            if not activos:
                st.markdown("""<div style="background:#f8fafc; border:1px solid #e2e8f0; border-top:none; border-radius:0 0 8px 8px; padding:20px; text-align:center; color:#9ca3af; font-size:12px;">Sin camiones pendientes</div>""", unsafe_allow_html=True)
            else:
                for j, c in enumerate(activos):
                    ipd = calcular_ipd(c)
                    color = get_ipd_color(ipd)
                    estado = c["estado"]
                    bg = "#fff5f5" if estado == "Crítico" else "#fffbeb" if estado == "Atrasado" else "#f0fdf4"
                    border = "#fecaca" if estado == "Crítico" else "#fde68a" if estado == "Atrasado" else "#bbf7d0"
                    icon = "🔴" if estado == "Crítico" else "🟡" if estado == "Atrasado" else "🟢"
 
                    st.markdown(f"""
                    <div style="background:{bg}; border:1px solid {border}; border-top:none; padding:10px 12px; {'border-radius:0 0 8px 8px' if j == len(activos)-1 else ''}">
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <div>
                                <span style="font-weight:700; font-size:13px; color:#111827;">#{j+1} {icon} {c['id']}</span>
                                <div style="font-size:11px; color:#6b7280;">{c['material'][:25]}{'...' if len(c['material'])>25 else ''}</div>
                                <div style="font-size:10px; color:#9ca3af;">{c['proveedor']} · {c['tipo']}</div>
                            </div>
                            <div style="text-align:right;">
                                <div style="font-size:18px; font-weight:800; color:{color};">{ipd}</div>
                                <div style="font-size:9px; color:#9ca3af;">IPD</div>
                            </div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
 
# ─────────────────────────────────────────────
# TAB: PLANIFICACIÓN DE VENTANAS
# ─────────────────────────────────────────────
def render_planificacion():
    st.markdown('<div class="section-title">Módulo de Planificación — Ventanas inteligentes</div>', unsafe_allow_html=True)
 
    col1, col2 = st.columns([2, 1])
 
    with col1:
        st.markdown("#### 📅 Plan de ventanas del día")
 
        # Gráfico Gantt de ventanas
        fig = go.Figure()
        colores = {"GENERAL": "#3b82f6", "INTERPLANTA": "#8b5cf6", "JIT": "#ef4444", "RANGER": "#22c55e"}
        y_pos = {"GENERAL": 3, "INTERPLANTA": 2, "JIT": 1, "RANGER": 0}
        estado_colors = {"Cumplida": "#22c55e", "En curso": "#3b82f6", "Demorada": "#ef4444", "Pendiente": "#9ca3af"}
 
        for v in st.session_state.ventanas_plan:
            try:
                hora = datetime.strptime(v["hora"], "%H:%M")
                fin = hora + timedelta(minutes=30)
                rec = v["recepcion"]
                ec = estado_colors.get(v["estado_ventana"], "#9ca3af")
                fig.add_trace(go.Bar(
                    x=[30], y=[rec], orientation='h',
                    base=[hora.hour * 60 + hora.minute],
                    marker_color=ec, marker_line_color="white", marker_line_width=1,
                    name=v["proveedor"],
                    hovertemplate=f"<b>{v['proveedor']}</b><br>{v['tipo']}<br>{v['hora']}<br>Estado: {v['estado_ventana']}<extra></extra>",
                    showlegend=False
                ))
            except:
                pass
 
        fig.update_layout(
            barmode='stack',
            xaxis=dict(title="Hora (minutos desde 00:00)", range=[440, 780],
                       tickvals=[480, 510, 540, 570, 600, 630, 660, 690, 720],
                       ticktext=["08:00","08:30","09:00","09:30","10:00","10:30","11:00","11:30","12:00"]),
            yaxis=dict(title="Recepción"),
            height=260, margin=dict(l=10, r=10, t=20, b=40),
            plot_bgcolor="#f8fafc", paper_bgcolor="#ffffff",
            font=dict(size=11)
        )
        st.plotly_chart(fig, use_container_width=True)
 
        # Tabla de ventanas
        st.markdown("**Detalle del plan:**")
        vc = {"Cumplida": "🟢", "En curso": "🔵", "Demorada": "🔴", "Pendiente": "⚪"}
        for v in st.session_state.ventanas_plan:
            ic = vc.get(v["estado_ventana"], "⚪")
            bg = "#fff5f5" if v["estado_ventana"] == "Demorada" else "#f0fdf4" if v["estado_ventana"] == "Cumplida" else "#f8fafc"
            st.markdown(f"""
            <div style="background:{bg}; border:1px solid #e2e8f0; border-radius:6px; padding:8px 12px; margin-bottom:4px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-weight:600; color:#111827; min-width:50px;">{v['hora']}</span>
                <span style="color:#374151; flex:1; margin:0 12px;">{v['proveedor']}</span>
                <span style="color:#6b7280; min-width:100px;">{v['recepcion']}</span>
                <span style="color:#6b7280; min-width:90px; font-size:12px;">{v['tipo']}</span>
                <span>{ic} {v['estado_ventana']}</span>
            </div>
            """, unsafe_allow_html=True)
 
    with col2:
        st.markdown("#### 🔴 Monitoreo Secuenciados")
        sec_camiones = [c for c in st.session_state.camiones if c["tipo"] == "SECUENCIADO"]
        for c in sec_camiones:
            modulos = c.get("modulos_disp", 1.0)
            stock = c.get("stock_actual", 0)
            pp = c.get("punto_pedido", 70)
            pct = min(stock / pp * 100, 100)
            color = "#ef4444" if pct < 25 else "#f59e0b" if pct < 65 else "#22c55e"
            alerta = "🔴 CRÍTICO — Llamar proveedor" if pct < 25 else "🟡 Monitorear" if pct < 65 else "🟢 OK"
 
            st.markdown(f"""
            <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:12px; margin-bottom:10px;">
                <div style="font-weight:700; font-size:13px; color:#111827;">{c['material']}</div>
                <div style="font-size:11px; color:#6b7280; margin-bottom:8px;">{c['proveedor']} · {c['id']}</div>
                <div style="background:#e2e8f0; border-radius:4px; height:10px; margin-bottom:6px;">
                    <div style="background:{color}; width:{pct:.0f}%; height:100%; border-radius:4px;"></div>
                </div>
                <div style="display:flex; justify-content:space-between; font-size:11px;">
                    <span style="color:{color}; font-weight:600;">{alerta}</span>
                    <span style="color:#6b7280;">{stock}/{pp} u.</span>
                </div>
                <div style="font-size:10px; color:#9ca3af; margin-top:4px;">
                    Próx. ventana: {c.get('ventana','—')} · 
                    {'Camión en planta' if c.get('llegada') else 'Sin confirmar llegada'}
                </div>
            </div>
            """, unsafe_allow_html=True)
 
        st.markdown("#### ➕ Agregar ventana")
        with st.form("nueva_ventana"):
            nv_hora = st.text_input("Hora (HH:MM)", "12:00")
            nv_rec = st.selectbox("Recepción", RECEPCIONES)
            nv_prov = st.text_input("Proveedor")
            nv_tipo = st.selectbox("Tipo", ["LOTE", "SECUENCIADO"])
            if st.form_submit_button("Agregar al plan"):
                st.session_state.ventanas_plan.append({
                    "hora": nv_hora, "recepcion": nv_rec,
                    "proveedor": nv_prov, "tipo": nv_tipo, "estado_ventana": "Pendiente"
                })
                st.success("✅ Ventana agregada")
                st.rerun()
 
# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    init_data()
    render_header()
 
    tabs = st.tabs([
        "📊 Tablero General",
        "📦 GENERAL",
        "🔄 INTERPLANTA",
        "⚡ JIT",
        "🛻 RANGER",
        "🗓️ Planificación"
    ])
 
    with tabs[0]:
        render_tablero_general()
 
    with tabs[1]:
        render_recepcion("GENERAL")
 
    with tabs[2]:
        render_recepcion("INTERPLANTA")
 
    with tabs[3]:
        render_recepcion("JIT")
 
    with tabs[4]:
        render_recepcion("RANGER")
 
    with tabs[5]:
        render_planificacion()
 
if __name__ == "__main__":
    main()