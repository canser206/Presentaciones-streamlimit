"""
================================================================================
  DASHBOARD DE GESTIÓN HUMANA – COEXCCOl | MAYO 2026
  Script Streamlit Legendario – Elaborado con datos reales de los archivos
  de productividad, ausentismo, consolidado ing-egre e indicadores RH
  
  Ejecutar en CMD:
      pip install streamlit plotly pandas openpyxl
      streamlit run dashboard_coexccol.py
================================================================================
"""

import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
import numpy as np
from datetime import datetime

# ─── CONFIGURACIÓN GLOBAL ────────────────────────────────────────────────────
st.set_page_config(
    page_title="COEXCCOL · RH 2026",
    page_icon="⛏️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS ÉPICO ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Fondo oscuro minero */
    .stApp { background: linear-gradient(135deg, #0a0e1a 0%, #0d1b2a 50%, #0a1628 100%); }
    
    /* Sidebar premium */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0d1b2a 0%, #132239 100%);
        border-right: 2px solid #f5a623;
    }
    
    /* Texto global */
    h1, h2, h3, p, span, div { color: #e8edf5 !important; }
    
    /* Métricas */
    [data-testid="stMetric"] {
        background: linear-gradient(135deg, #1a2a3d 0%, #1e3450 100%);
        border: 1px solid rgba(245,166,35,0.4);
        border-radius: 12px;
        padding: 16px;
        box-shadow: 0 4px 20px rgba(245,166,35,0.15), inset 0 1px 0 rgba(255,255,255,0.1);
    }
    [data-testid="stMetric"] label { color: #f5a623 !important; font-weight: 700; font-size: 0.85rem; }
    [data-testid="stMetricValue"] { color: #ffffff !important; font-size: 1.8rem !important; font-weight: 900 !important; }
    [data-testid="stMetricDelta"] { font-weight: 600 !important; }
    
    /* Títulos de sección */
    .section-title {
        background: linear-gradient(90deg, #f5a623, #ff8c00);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 1.6rem;
        font-weight: 900;
        margin-bottom: 1.5rem;
        padding-bottom: 0.5rem;
        border-bottom: 2px solid rgba(245,166,35,0.3);
    }
    
    /* Cards personalizadas */
    .card {
        background: linear-gradient(135deg, #132239 0%, #1a3050 100%);
        border: 1px solid rgba(245,166,35,0.3);
        border-radius: 14px;
        padding: 20px;
        margin: 8px 0;
        box-shadow: 0 8px 32px rgba(0,0,0,0.3);
        transition: transform 0.2s ease;
    }
    .card:hover { transform: translateY(-2px); border-color: rgba(245,166,35,0.7); }
    
    /* Badge de estado */
    .badge-ok { background:#1a4731; color:#4ade80; border:1px solid #4ade80; border-radius:6px; padding:3px 10px; font-size:0.78rem; font-weight:700; }
    .badge-warn { background:#4a3000; color:#f5a623; border:1px solid #f5a623; border-radius:6px; padding:3px 10px; font-size:0.78rem; font-weight:700; }
    .badge-err { background:#4a1010; color:#f87171; border:1px solid #f87171; border-radius:6px; padding:3px 10px; font-size:0.78rem; font-weight:700; }
    
    /* Tabla de procesos */
    .proc-row {
        background: rgba(26,42,61,0.8);
        border: 1px solid rgba(245,166,35,0.15);
        border-radius: 10px;
        padding: 12px 16px;
        margin: 6px 0;
    }
    
    /* Header principal */
    .main-header {
        text-align: center;
        padding: 30px 0 20px;
    }
    .main-header h1 { font-size: 2.6rem; font-weight: 900; margin: 0; color: #ffffff !important; }
    .main-header p { color: #94a3b8 !important; font-size: 1rem; margin: 6px 0 0; }
    
    /* Divisores */
    hr { border-color: rgba(245,166,35,0.25) !important; margin: 2rem 0 !important; }
    
    /* DataFrames */
    [data-testid="stDataFrame"] { border: 1px solid rgba(245,166,35,0.25); border-radius: 10px; }
    
    /* Tabs */
    .stTabs [data-baseweb="tab"] { background: #132239; border: 1px solid rgba(245,166,35,0.2); color: #e8edf5 !important; border-radius: 8px 8px 0 0; }
    .stTabs [aria-selected="true"] { background: linear-gradient(135deg, #f5a623, #ff8c00) !important; color: #000 !important; font-weight: 800; }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
#  DATOS (Hardcoded desde los archivos — sin necesidad de leer Excel en runtime)
# ═══════════════════════════════════════════════════════════════════════════════

# ── PRODUCCIÓN HISTÓRICA ──────────────────────────────────────────────────────
prod_data = pd.DataFrame({
    "Mes": ["Enero", "Febrero", "Marzo", "Abril", "Mayo"],
    "Toneladas": [1698, 1894, 2000, 1800, 1964],
    "Meta": [2156.9, 2156.9, 2156.9, 2156.9, 2156.9],
    "Ausencias_pct": [17, 15, 19, 23, 16],
    "Trabajadores": [65, 69, 80, 78, 82],
    "Cumplimiento_pct": [78.7, 87.8, 92.7, 83.5, 89.3],
})

# ── RANKING PRODUCTIVIDAD MAYO ────────────────────────────────────────────────
ranking = pd.DataFrame([
    {"Rank": 1, "Trabajador": "MOLINA NILSON",      "Toneladas": 82.49, "Meta": 84, "Desempeño": "SOBRESALIENTE 🏆", "Pagado": 3454800},
    {"Rank": 2, "Trabajador": "PATIÑO JAIRO",       "Toneladas": 80.75, "Meta": 84, "Desempeño": "SOBRESALIENTE 🏆", "Pagado": 3351640},
    {"Rank": 3, "Trabajador": "MORENO WILFER",      "Toneladas": 75.41, "Meta": 84, "Desempeño": "DESTACADO ⭐",     "Pagado": 2290250},
    {"Rank": 4, "Trabajador": "BARRIOS LUIS",       "Toneladas": 73.41, "Meta": 84, "Desempeño": "DESTACADO ⭐",     "Pagado": 3056950},
    {"Rank": 5, "Trabajador": "OROPEZA DUMAR",      "Toneladas": 70.25, "Meta": 84, "Desempeño": "DESTACADO ⭐",     "Pagado": 2993250},
    {"Rank": 6, "Trabajador": "SERRANO JOSE",       "Toneladas": 67.41, "Meta": 84, "Desempeño": "DESTACADO ⭐",     "Pagado": 3009405},
    {"Rank": 7, "Trabajador": "PATIÑO MARTIN",      "Toneladas": 66.66, "Meta": 84, "Desempeño": "DESTACADO ⭐",     "Pagado": 3444100},
    {"Rank": 8, "Trabajador": "GONZALEZ FERNANDO",  "Toneladas": 64.18, "Meta": 84, "Desempeño": "DESTACADO ⭐",     "Pagado": 2687100},
    {"Rank": 9, "Trabajador": "RODRIGUEZ J.L.",     "Toneladas": 59.04, "Meta": 84, "Desempeño": "DESTACADO ⭐",     "Pagado": 2256650},
    {"Rank":10, "Trabajador": "ESCOBAR DISNEY",     "Toneladas": 53.53, "Meta": 84, "Desempeño": "REGULAR",         "Pagado": 2772670},
    {"Rank":11, "Trabajador": "MUÑOZ YORFAN",       "Toneladas": 53.05, "Meta": 84, "Desempeño": "REGULAR",         "Pagado": 3209764},
    {"Rank":12, "Trabajador": "GONZALEZ CARLOS",    "Toneladas": 50.57, "Meta": 84, "Desempeño": "REGULAR",         "Pagado": 3008290},
    {"Rank":13, "Trabajador": "SAAVEDRA JOSE ALI",  "Toneladas": 50.53, "Meta": 84, "Desempeño": "REGULAR",         "Pagado": 2431435},
    {"Rank":14, "Trabajador": "ROJAS WILMER",       "Toneladas": 50.36, "Meta": 84, "Desempeño": "REGULAR",         "Pagado": 2642300},
    {"Rank":15, "Trabajador": "BENAVIDES CARLOS",   "Toneladas": 48.15, "Meta": 84, "Desempeño": "REGULAR",         "Pagado": 3414976},
])

# ── COSTOS OPERATIVOS MAYO ────────────────────────────────────────────────────
costos = pd.DataFrame([
    {"Concepto": "PICADA (Coches)",  "Valor": 63934305, "Tipo": "Operativo"},
    {"Concepto": "ALIMENTACIÓN",     "Valor": 15487750, "Tipo": "Bienestar"},
    {"Concepto": "DOMINICALES",      "Valor": 15000000, "Tipo": "Legal"},
    {"Concepto": "ALEMANA",          "Valor": 10412790, "Tipo": "Operativo"},
    {"Concepto": "COCHES GUAYADOS",  "Valor":  6730000, "Tipo": "Operativo"},
    {"Concepto": "MALACATEROS",      "Valor":  5917314, "Tipo": "Apoyo"},
    {"Concepto": "VENTANA",          "Valor":  4900000, "Tipo": "Operativo"},
    {"Concepto": "PATIEROS",         "Valor":  3794195, "Tipo": "Apoyo"},
    {"Concepto": "METRO DE ROCA",    "Valor":  2988300, "Tipo": "Operativo"},
    {"Concepto": "RETRO (PAJARITA)", "Valor":  2846897, "Tipo": "Maquinaria"},
    {"Concepto": "TOLVEREROS",       "Valor":  2799334, "Tipo": "Apoyo"},
    {"Concepto": "SOLDADORES",       "Valor":  2353887, "Tipo": "Apoyo"},
    {"Concepto": "TRAMO",            "Valor":  1841400, "Tipo": "Operativo"},
    {"Concepto": "TURNO",            "Valor":  1753400, "Tipo": "Operativo"},
])

# ── AUSENTISMO – CAUSAS POR MES (resumen de los archivos) ────────────────────
ausencias = pd.DataFrame({
    "Mes": ["Enero","Enero","Enero","Enero","Enero","Enero",
            "Febrero","Febrero","Febrero","Febrero",
            "Marzo","Marzo","Marzo","Marzo",
            "Abril","Abril","Abril","Abril","Abril",
            "Mayo","Mayo","Mayo","Mayo","Mayo"],
    "Causa": [
        "Ausencia Injustificada","Enfermedad General","Fuerza Mayor","Calamidad Doméstica","Abandono Puesto","Accidente Trabajo",
        "Ausencia Injustificada","Enfermedad General","Accidente Trabajo","Abandono Puesto",
        "Ausencia Injustificada","Enfermedad General","Accidente Trabajo","Fuerza Mayor",
        "Ausencia Injustificada","Enfermedad General","Accidente Trabajo","Fuerza Mayor","Sanción",
        "Ausencia Injustificada","Enfermedad General","Accidente Trabajo","Abandono Puesto","Retiro",
    ],
    "Días": [27, 6, 1, 4, 3, 31,
             15, 8, 28, 3,
             22, 14, 46, 16,
             35, 18, 31, 1, 6,
             38, 24, 75, 9, 12],
})

# ── AUSENTISMO MENSUAL RESUMEN ────────────────────────────────────────────────
ausencias_resumen = pd.DataFrame({
    "Mes": ["Enero","Febrero","Marzo","Abril","Mayo"],
    "Dias_Cotizados": [1625, 1735, 1857, 2038, 2158],
    "Dias_Incapacidad": [39, 30, 95, 75, 127],
    "Indice_Desviacion": [
        round(39/1625*100, 2),
        round(30/1735*100, 2),
        round(95/1857*100, 2),
        round(75/2038*100, 2),
        round(127/2158*100, 2),
    ],
})

# ── DOMINICALES ───────────────────────────────────────────────────────────────
dominicales = pd.DataFrame({
    "Mes": ["Enero","Febrero","Marzo","Abril","Mayo"],
    "Trabajadores": [51, 50, 52, 56, 62],
    "Dominicales_Posibles": [204, 200, 208, 224, 248],
    "Dominicales_Pagados": [105, 120, 196, 190, 225],
    "Dominicales_Perdidos": [99, 80, 12, 34, 23],
})

# ── ROTACIÓN (Consolidado Ing-Egre) ──────────────────────────────────────────
rotacion = pd.DataFrame({
    "Mes": ["Enero","Febrero","Marzo","Abril","Mayo"],
    "Total_Trabajadores": [71, 75, 84, 84, 82],
    "Ingresos": [17, 10, 21, 12, 10],
    "Egresos": [6, 10, 11, 12, 12],
})

causas_egreso = pd.DataFrame([
    {"Causa": "Inconformidad / Distancia / Precio", "Frecuencia": 12, "Tipo": "Voluntaria"},
    {"Causa": "Salarios / Beneficios (otra mina)",  "Frecuencia": 10, "Tipo": "Voluntaria"},
    {"Causa": "Falta de compromiso / Ausencias",    "Frecuencia":  9, "Tipo": "Involuntaria"},
    {"Causa": "Contratación deficiente",             "Frecuencia":  8, "Tipo": "Involuntaria"},
    {"Causa": "Condiciones físicas (no apto)",       "Frecuencia":  5, "Tipo": "Voluntaria"},
    {"Causa": "Para crecimiento profesional",        "Frecuencia":  3, "Tipo": "Voluntaria"},
    {"Causa": "Problemas comunicación / Jefe",       "Frecuencia":  2, "Tipo": "Voluntaria"},
    {"Causa": "Clima laboral / Otras",               "Frecuencia":  2, "Tipo": "Involuntaria"},
])

# ── INCAPACIDADES ─────────────────────────────────────────────────────────────
incapacidades = pd.DataFrame([
    {"Trabajador": "MANUEL LAYA",     "Tipo": "AT", "Diagnóstico": "S925", "Días": 30, "Valor": 1750905, "Estado": "CANCELADA"},
    {"Trabajador": "JEFERSON GOMEZ",  "Tipo": "AT", "Diagnóstico": "S835", "Días": 30, "Valor": 1750905, "Estado": "CANCELADA"},
    {"Trabajador": "JOAQUIN RIVERA",  "Tipo": "AT", "Diagnóstico": "S731", "Días": 30, "Valor": 1750905, "Estado": "CANCELADA"},
    {"Trabajador": "JAVIER FONSECA",  "Tipo": "AT", "Diagnóstico": "M545", "Días": 21, "Valor": 1750905, "Estado": "CANCELADA"},
    {"Trabajador": "JAVIER FONSECA",  "Tipo": "AT", "Diagnóstico": "M545", "Días": 30, "Valor": 1750905, "Estado": "CANCELADA"},
])

# ── LIQUIDACIONES (resumen) ───────────────────────────────────────────────────
liq_pendientes_total = 106_657_573
liq_pagadas_total    = 72_426_922
liq_gran_total       = 179_084_495

# ── QUEJAS Y RECLAMOS ────────────────────────────────────────────────────────
quejas = pd.DataFrame([
    {"ID": 1, "Descripción": "Solicitud de colchonetas para el personal", "Estado": "✅ Solucionado", "Acción": "Se entregaron 18 colchonetas"},
    {"ID": 2, "Descripción": "Solicitud de bonificación para malacateros",  "Estado": "✅ Solucionado", "Acción": "Pago de bonificación mensual a partir de abril"},
    {"ID": 3, "Descripción": "Solicitud menú: Bandeja Paisa los viernes",   "Estado": "✅ Solucionado", "Acción": "Se volvió a incluir bandeja paisa los viernes"},
])

# ── PROCESOS LEGALES ──────────────────────────────────────────────────────────
procesos = pd.DataFrame([
    {"Fecha":"23-Ene-26","Causa":"Irrespeto personal de seguridad (Molina + Gómez)","Sanción":"Llamado de atención"},
    {"Fecha":"05-Feb-26","Causa":"Ausencias injustificadas – Jesús Hernández","Sanción":"Culminación de contrato"},
    {"Fecha":"18-Feb-26","Causa":"Ausencias injustificadas – Walter Rincón","Sanción":"Culminación de contrato"},
    {"Fecha":"20-Feb-26","Causa":"Incumplimiento registro mediciones en turno","Sanción":"Llamado de atención"},
    {"Fecha":"25-Feb-26","Causa":"Carlos Lizcano no entra a charla de seguridad","Sanción":"Llamado de atención"},
    {"Fecha":"18-Mar-26","Causa":"Patieros dejando basura fuera del área","Sanción":"Llamado de atención"},
    {"Fecha":"21-Mar-26","Causa":"C. Benavides – ausencias exámenes médicos","Sanción":"Llamado de atención"},
    {"Fecha":"21-Mar-26","Causa":"J. Chaparro y S. Cárdenas – inasistencia a turno","Sanción":"Pérdida de dominicales"},
    {"Fecha":"24-Mar-26","Causa":"Múltiples trabajadores – ausencias (Cárdenas, Chaparro, Serrano, Laya, Chiquillo, Segovia)","Sanción":"Pérdida de dominicales"},
    {"Fecha":"25-Mar-26","Causa":"Richard – comportamiento agresivo con compañeros","Sanción":"Llamado de atención"},
    {"Fecha":"25-Mar-26","Causa":"S. Sánchez – extravío martillo + abandono de trabajo","Sanción":"Culminación de contrato + Acuerdo de pago"},
    {"Fecha":"06-Abr-26","Causa":"Firma de libro entrada/salida de forma irregular","Sanción":"Llamado de atención + charla seguridad"},
    {"Fecha":"09-Abr-26","Causa":"Supervisores: llegada tarde a turno (Viancha)","Sanción":"Llamado de atención (política media hora antes)"},
    {"Fecha":"22-Abr-26","Causa":"Dumar Oropeza – positivo en alcoholemia","Sanción":"Suspensión 3 días + pérdida dominical"},
    {"Fecha":"23-Abr-26","Causa":"Carlos González – positivo en alcoholemia","Sanción":"Suspensión 3 días + pérdida dominical"},
    {"Fecha":"27-Abr-26","Causa":"Arley Ballesta y Johan Calderón – ausencias + llegada tarde","Sanción":"Llamado de atención (traslado si reincide)"},
])

# ── PRESUPUESTO GESTIÓN LABORAL (Trimestre) ───────────────────────────────────
ppto_laboral = pd.DataFrame([
    {"Actividad": "Celebración cumpleaños (mensual)",      "Frecuencia": "Mensual",    "Costo_Unit": 150000, "Cantidad": 3,  "Total": 450000,  "Estado": "⏳ Pendiente"},
    {"Actividad": "Videos familiares motivacionales",       "Frecuencia": "Diario",     "Costo_Unit": 0,      "Cantidad": 90, "Total": 0,       "Estado": "✅ Aprobado"},
    {"Actividad": "Integración Comfaboy (piscina+almuerzo)","Frecuencia": "Trimestral", "Costo_Unit": 60000,  "Cantidad": 20, "Total": 1200000, "Estado": "⏳ Pendiente"},
    {"Actividad": "Integración administrativa",             "Frecuencia": "Trimestral", "Costo_Unit": 120000, "Cantidad": 10, "Total": 1200000, "Estado": "⏳ Pendiente"},
    {"Actividad": "Cambio de colchonetas",                  "Frecuencia": "Mensual",    "Costo_Unit": 45000,  "Cantidad": 15, "Total": 675000,  "Estado": "✅ Aprobado"},
    {"Actividad": "Reconocimiento trabajadores destacados", "Frecuencia": "Mensual",    "Costo_Unit": 50000,  "Cantidad": 3,  "Total": 150000,  "Estado": "⏳ Pendiente"},
    {"Actividad": "Incentivo alimentación operativo",       "Frecuencia": "Mensual",    "Costo_Unit": 30000,  "Cantidad": 60, "Total": 1800000, "Estado": "⏳ Pendiente"},
])

ppto_social = pd.DataFrame([
    {"Iniciativa": "Visita familias con detalles (mercado/frutas)", "Beneficiarios": "Familias trabajadores", "Costo": 800000,  "Mes": "Julio",      "Responsable": "RR.HH."},
    {"Iniciativa": "Viajes turísticos adultos mayores (vecinos)",    "Beneficiarios": "Adultos mayores",     "Costo": 1500000, "Mes": "Agosto",     "Responsable": "Gestión Social"},
    {"Iniciativa": "Integración escuela Sagra (onces + charla)",     "Beneficiarios": "Alumnos Sagra",      "Costo": 600000,  "Mes": "Septiembre", "Responsable": "HSEQ / RR.HH."},
    {"Iniciativa": "Jornada salud visual (Optisalud)",               "Beneficiarios": "Familias",           "Costo": 1200000, "Mes": "Agosto",     "Responsable": "SST / RR.HH."},
    {"Iniciativa": "Convenio Funeraria La Aurora",                   "Beneficiarios": "Trabajadores",       "Costo": 0,       "Mes": "Continuo",   "Responsable": "Gerencia"},
])

# ── SUBSIDIO ALIMENTACIÓN MAYO ────────────────────────────────────────────────
subsidio = {
    "toneladas": 1964,
    "valor_tonelada": 290000,
    "ingreso_total": 569_560_000,
    "subsidio_total": 16_861_250,
    "costo_por_ton": 8585,
    "pct_ingreso": 5.09,
}

# ── DATOS APOYO MAYO (costo personal apoyo vs producción) ────────────────────
apoyo_q1 = {
    "MALACATERO GAMEZ":         1342794,
    "MALACATERO TORRES":        1456540,
    "TOLVERO CHAVARRO":          875453,
    "TOLVERO CATALAN":          1088269,
    "PATIERO ALIRIO MENDEZ":    1037268,
    "PATIERO GAMEZ FRAGOSO":    1037268,
    "SOLDADOR VASQUEZ":         1436000,
    "RETRO-SOLDADOR SOLEDAD":   1410898,
}
apoyo_q2 = {
    "MALACATERO GAMEZ":          695802,
    "MALACATERO TORRES":        1665366,
    "TOLVERO CHAVARRO":          917887,
    "TOLVERO CATALAN":          1481610,
    "PORTERO ARCINIEGAS":        875445,
    "PATIERO ALIRIO MENDEZ":     915237,
    "PATIERO GAMEZ FRAGOSO":     915237,
    "RETRO-OPERADOR":           1334605,
    "SOLDADOR VASQUEZ":         1436000,
    "SOLDADOR MADERO":           590000,
}

total_apoyo_mayo = sum(apoyo_q1.values()) + sum(apoyo_q2.values())
costo_apoyo_por_ton = total_apoyo_mayo / 1964

apoyo_df = pd.DataFrame([
    {"Rol": k, "Q1": v, "Q2": apoyo_q2.get(k, 0)} for k, v in apoyo_q1.items()
] + [
    {"Rol": k, "Q1": 0, "Q2": v} for k, v in apoyo_q2.items() if k not in apoyo_q1
])
apoyo_df["Total"] = apoyo_df["Q1"] + apoyo_df["Q2"]

# ─── SIDEBAR ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='text-align:center; padding: 10px 0 20px'>
        <div style='font-size:3rem'>⛏️</div>
        <div style='color:#f5a623;font-weight:900;font-size:1.1rem;letter-spacing:2px'>COEXCCOL</div>
        <div style='color:#94a3b8;font-size:0.8rem;margin-top:4px'>MINA PAJARITA · MAYO 2026</div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown("<div style='color:#f5a623;font-weight:700;font-size:0.9rem;margin-bottom:12px'>⚙️ NAVEGACIÓN</div>", unsafe_allow_html=True)
    
    seccion = st.radio(
        "",
        [
            "🏠 Resumen Ejecutivo",
            "⛏️ Productividad",
            "🚫 Ausentismo",
            "🔗 Producción vs Ausentismo",
            "👥 Ingreso & Egreso",
            "💰 Costos & Subsidio",
            "📋 Indicadores RH",
        ],
        label_visibility="collapsed",
    )
    
    st.markdown("---")
    st.markdown("""
    <div style='background:rgba(245,166,35,0.1);border:1px solid rgba(245,166,35,0.3);
                border-radius:10px;padding:12px;margin-top:12px'>
        <div style='color:#f5a623;font-weight:700;font-size:0.85rem'>📊 PRODUCCIÓN MAYO</div>
        <div style='color:#fff;font-size:1.6rem;font-weight:900;margin:6px 0'>1,964 TON</div>
        <div style='color:#4ade80;font-size:0.85rem'>+127 ton vs Abril ↑7.1%</div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    <div style='margin-top:12px;background:rgba(74,222,128,0.07);border:1px solid rgba(74,222,128,0.3);
                border-radius:10px;padding:12px'>
        <div style='color:#4ade80;font-weight:700;font-size:0.85rem'>👷 TRABAJADORES ACTIVOS</div>
        <div style='color:#fff;font-size:1.6rem;font-weight:900;margin:6px 0'>82</div>
        <div style='color:#94a3b8;font-size:0.82rem'>49 operativos + 33 apoyo/admin</div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("<br><div style='color:#475569;font-size:0.72rem;text-align:center'>Elaborado: Jefa RR.HH. & Jefa Proyecto<br>COEXCCOL · Junio 2026</div>", unsafe_allow_html=True)


# ─── HEADER PRINCIPAL ─────────────────────────────────────────────────────────
st.markdown("""
<div class='main-header'>
    <h1>⛏️ DASHBOARD DE GESTIÓN HUMANA</h1>
    <p>COMPAÑÍA EXPORTADORA DE CARBONES DE COLOMBIA SAS · FRENTE PAJARITA · COSTA RICA · MAYO 2026</p>
</div>
""", unsafe_allow_html=True)
st.markdown("---")

# ═══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 1: RESUMEN EJECUTIVO
# ═══════════════════════════════════════════════════════════════════════════════
if seccion == "🏠 Resumen Ejecutivo":
    st.markdown("<div class='section-title'>🏠 RESUMEN EJECUTIVO · MAYO 2026</div>", unsafe_allow_html=True)
    
    # KPIs principales
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("⛏️ Producción Mayo",  "1,964 TON",    "+193 vs Abril")
    c2.metric("🎯 Cumplimiento Meta", "91.05%",         "+7.6pp vs Abril")
    c3.metric("💵 Costo/Tonelada",   "$72,841 COP",   "")
    c4.metric("👷 Personal Activo",  "82",            "+4 vs Abril")
    c5.metric("🚫 Ausentismo Mayo",  "23%",           "= Abril")
    
    st.markdown("---")
    
    col1, col2 = st.columns([1.3, 1])
    
    with col1:
        st.markdown("#### 📈 Producción Histórica Enero – Mayo 2026")
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=prod_data["Mes"], y=prod_data["Toneladas"],
            name="Toneladas Producidas",
            marker=dict(
                color=prod_data["Toneladas"],
                colorscale=[[0,"#1a3050"],[0.5,"#f5a623"],[1,"#ff6b00"]],
                line=dict(color="rgba(245,166,35,0.5)", width=1),
            ),
            text=prod_data["Toneladas"].apply(lambda x: f"{x:,}"),
            textposition="outside", textfont=dict(color="white", size=13, family="Arial Black"),
        ))
        fig.add_trace(go.Scatter(
            x=prod_data["Mes"], y=prod_data["Meta"],
            name="Meta Mensual (2,157)",
            mode="lines+markers",
            line=dict(color="#f87171", dash="dash", width=2),
            marker=dict(symbol="diamond", size=8, color="#f87171"),
        ))
        fig.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(15,25,40,0.5)",
            legend=dict(orientation="h", y=-0.15, x=0, font=dict(color="white")),
            yaxis=dict(gridcolor="rgba(255,255,255,0.07)", range=[0, 2600]),
            xaxis=dict(gridcolor="rgba(255,255,255,0.07)"),
            margin=dict(t=30, b=40, l=10, r=10),
            height=320,
        )
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.markdown("#### 🏆 Top 5 Productividad Mayo")
        top5 = ranking.head(5)
        for _, r in top5.iterrows():
            cumpl = round(r["Toneladas"] / r["Meta"] * 100, 1)
            color = "#f5a623" if r["Rank"] == 1 else ("#94a3b8" if r["Rank"] == 2 else "#cd7f32")
            bar_w = int(r["Toneladas"] / 84 * 100)
            st.markdown(f"""
            <div class='card' style='padding:14px'>
                <div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:8px'>
                    <span style='color:{color};font-weight:900;font-size:1.1rem'>#{r["Rank"]} {r["Trabajador"]}</span>
                    <span style='color:#fff;font-weight:700'>{r["Toneladas"]} TON</span>
                </div>
                <div style='background:rgba(255,255,255,0.1);border-radius:4px;height:8px;overflow:hidden'>
                    <div style='width:{bar_w}%;background:linear-gradient(90deg,{color},{color}88);height:100%;border-radius:4px'></div>
                </div>
                <div style='display:flex;justify-content:space-between;margin-top:6px'>
                    <span style='color:#94a3b8;font-size:0.78rem'>{r["Desempeño"]}</span>
                    <span style='color:#4ade80;font-size:0.78rem'>{cumpl}% meta</span>
                </div>
            </div>
            """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Fila de alertas / puntos clave
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("""
        <div class='card'>
            <div style='color:#f5a623;font-weight:800;margin-bottom:10px'>✅ QUEJAS Y RECLAMOS</div>
            <div style='color:#4ade80;font-size:2rem;font-weight:900'>3/3</div>
            <div style='color:#94a3b8;font-size:0.85rem'>Todas resueltas al 100%</div>
            <div style='margin-top:8px;font-size:0.82rem;color:#cbd5e1'>
                • Colchonetas entregadas<br>
                • Bonificación malacateros activa<br>
                • Bandeja paisa viernes restaurada
            </div>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown("""
        <div class='card'>
            <div style='color:#f5a623;font-weight:800;margin-bottom:10px'>⚠️ INCAPACIDADES AT</div>
            <div style='color:#f87171;font-size:2rem;font-weight:900'>5</div>
            <div style='color:#94a3b8;font-size:0.85rem'>Accidentes de Trabajo activos</div>
            <div style='margin-top:8px;font-size:0.82rem;color:#cbd5e1'>
                • Manuel Laya (AT·S925·30 días)<br>
                • Jeferson Gómez (AT·S835·30 días)<br>
                • Joaquin Rivera (AT·S731·30 días)<br>
                • Javier Fonseca (AT·M545·51 días)
            </div>
        </div>
        """, unsafe_allow_html=True)
    with c3:
        liq_pend_count = 35  # aprox pendientes
        st.markdown(f"""
        <div class='card'>
            <div style='color:#f5a623;font-weight:800;margin-bottom:10px'>💰 LIQUIDACIONES</div>
            <div style='color:#fbbf24;font-size:1.5rem;font-weight:900'>${liq_pendientes_total/1e6:.1f}M COP</div>
            <div style='color:#94a3b8;font-size:0.85rem'>Pendientes por pagar</div>
            <div style='margin-top:8px;font-size:0.82rem;color:#cbd5e1'>
                • Total facturado: ${liq_gran_total/1e6:.1f}M COP<br>
                • Ya pagado: ${liq_pagadas_total/1e6:.1f}M COP<br>
                • Pendiente: ${liq_pendientes_total/1e6:.1f}M COP
            </div>
        </div>
        """, unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 2: PRODUCTIVIDAD
# ═══════════════════════════════════════════════════════════════════════════════
elif seccion == "⛏️ Productividad":
    st.markdown("<div class='section-title'>⛏️ ANÁLISIS DE PRODUCTIVIDAD · MAYO 2026</div>", unsafe_allow_html=True)
    
    # KPIs productividad
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🏆 Mayor Productor",  "MOLINA NILSON", "82.49 TON")
    c2.metric("📊 Promedio Equipo",  "33.06 TON",     "Meta: 84 TON")
    c3.metric("✅ Cumplimiento Prom","39.4%",         "vs 100% meta")
    c4.metric("👷 Operadores Eval.", "49",            "")
    
    st.markdown("---")
    
    tab1, tab2, tab3 = st.tabs(["📊 Ranking Completo (Top 15)", "🧩 Distribución Desempeño", "🔬 Factores de Productividad"])
    
    with tab1:
        fig = go.Figure()
        colors = ["#f5a623" if d.startswith("SOBRES") else "#60a5fa" if d.startswith("DEST") else "#a78bfa" if d=="REGULAR" else "#f87171"
                  for d in ranking["Desempeño"]]
        fig.add_trace(go.Bar(
            x=ranking["Toneladas"], y=ranking["Trabajador"],
            orientation="h",
            marker=dict(color=colors, line=dict(color="rgba(255,255,255,0.15)", width=1)),
            text=ranking["Toneladas"].apply(lambda x: f"{x} T"),
            textposition="outside",
            textfont=dict(color="white", size=11),
        ))
        fig.add_vline(x=84, line_dash="dash", line_color="#f87171",
                      annotation_text="META 84T", annotation_font_color="#f87171",
                      annotation_position="top right")
        fig.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(15,25,40,0.5)",
            height=600, margin=dict(l=160, r=60, t=20, b=20),
            yaxis=dict(gridcolor="rgba(255,255,255,0.05)", tickfont=dict(color="white", size=11)),
            xaxis=dict(gridcolor="rgba(255,255,255,0.07)", range=[0, 100]),
        )
        st.plotly_chart(fig, use_container_width=True)
        
        st.dataframe(
            ranking[["Rank","Trabajador","Toneladas","Desempeño","Pagado"]].rename(columns={
                "Toneladas":"Ton Producidas","Desempeño":"Nivel","Pagado":"Pagado (COP)"}),
            use_container_width=True,
            hide_index=True,
        )
    
    with tab2:
        desempeño_count = {
            "SOBRESALIENTE 🏆": 2, "DESTACADO ⭐": 7, "REGULAR": 6, "BAJO ⚠": 34
        }
        fig = go.Figure(go.Pie(
            labels=list(desempeño_count.keys()),
            values=list(desempeño_count.values()),
            hole=0.55,
            marker=dict(colors=["#f5a623","#60a5fa","#a78bfa","#f87171"],
                        line=dict(color="#0a0e1a", width=2)),
            textinfo="label+percent",
            textfont=dict(size=13, color="white"),
        ))
        fig.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)",
            annotations=[dict(text="49<br><span style='font-size:10px'>operadores</span>",
                              x=0.5, y=0.5, font_size=18, showarrow=False, font_color="white")],
            height=380, margin=dict(t=20, b=20),
        )
        st.plotly_chart(fig, use_container_width=True)
        
        st.info("⚠️ El 69% del personal opera **por debajo** de la meta de 84 ton/quincena. Solo 9 trabajadores (18%) alcanzaron categoría DESTACADA o SOBRESALIENTE.")
    
    with tab3:
        st.markdown("### 🔬 ¿Qué incide en la productividad del trabajador?")
        factores = [
            ("🏃 Ausentismo y ausencias injustificadas", "El 23% de ausentismo en Abril-Mayo generó pérdida directa de turno productivo. Cada día perdido equivale a ±3.3 toneladas no producidas.", "warn"),
            ("⚠️ Alcoholemia positiva (2 casos Abril)", "Dos operadores suspendidos 3 días cada uno: 6 días x 3.3 ton ≈ 20 ton no producidas.", "err"),
            ("💪 Experiencia del trabajador", "Los 4 mejores productores tienen 4-8 años de experiencia. Los de bajo desempeño promedian 1-2 años.", "ok"),
            ("📐 Labor asignada (corte vs frente)", "Trabajadores con acceso a frentes activos y coches disponibles producen hasta 5x más que los de labores secundarias.", "ok"),
            ("🔧 Disponibilidad de retroexcavadora", "La retro soportó el 100% de la extracción. En Q2 mejoró de 3.89 a 4.65 ton/hora (+19.6%). Su gestión eficiente impacta directamente.", "ok"),
            ("🏥 Incapacidades por accidente de trabajo", "4 trabajadores en AT durante el período representan días-hombre perdidos sin producción. Se requiere control de seguridad.", "err"),
            ("🌧️ Condiciones del frente de trabajo", "Los datos muestran variabilidad quincena a quincena (Q1: 933 ton → Q2: 1,117 ton en mismo mes). Las condiciones geológicas son determinantes.", "warn"),
        ]
        for titulo, desc, tipo in factores:
            badge_class = f"badge-{tipo}"
            badge_text = "✅ FACTOR POSITIVO" if tipo=="ok" else ("⚠️ FACTOR DE RIESGO" if tipo=="warn" else "🚨 IMPACTO DIRECTO")
            st.markdown(f"""
            <div class='card' style='margin:8px 0'>
                <div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:8px'>
                    <span style='font-weight:700;font-size:1rem;color:#e8edf5'>{titulo}</span>
                    <span class='{badge_class}'>{badge_text}</span>
                </div>
                <p style='color:#94a3b8;font-size:0.88rem;margin:0'>{desc}</p>
            </div>
            """, unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 3: AUSENTISMO
# ═══════════════════════════════════════════════════════════════════════════════
elif seccion == "🚫 Ausentismo":
    st.markdown("<div class='section-title'>🚫 ANÁLISIS DE AUSENTISMO · ENERO – MAYO 2026</div>", unsafe_allow_html=True)
    
    tab1, tab2, tab3 = st.tabs(["📊 Incapacidades & Días Cotizados", "🔍 Causas de Ausencia", "📅 Dominicales"])
    
    with tab1:
        c1, c2, c3 = st.columns(3)
        c1.metric("📅 Días Incapacidad (Ene-May)", "366 días", "127 solo en Mayo")
        c2.metric("📊 Índice Desviación Mayo",     "5.89%",    "+2.32pp vs Marzo")
        c3.metric("🏥 Incapacidades AT activas",   "5 casos",  "")
        
        # Gráfico barras días cotizados vs incapacidad
        fig = go.Figure()
        fig.add_trace(go.Bar(
            name="Días Cotizados", x=ausencias_resumen["Mes"], y=ausencias_resumen["Dias_Cotizados"],
            marker_color="#1e3a5f",
            text=ausencias_resumen["Dias_Cotizados"], textposition="inside", textfont=dict(color="white", size=12),
        ))
        fig.add_trace(go.Bar(
            name="Días Incapacidad", x=ausencias_resumen["Mes"], y=ausencias_resumen["Dias_Incapacidad"],
            marker_color="#f87171",
            text=ausencias_resumen["Dias_Incapacidad"], textposition="outside", textfont=dict(color="white", size=12),
        ))
        fig.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(15,25,40,0.5)",
            barmode="overlay", height=360,
            legend=dict(orientation="h", y=-0.15, font=dict(color="white")),
            yaxis=dict(gridcolor="rgba(255,255,255,0.07)"),
            margin=dict(t=20, b=40),
            title=dict(text="Días Cotizados vs Días Incapacidad por Mes", font=dict(color="white"), x=0.5),
        )
        st.plotly_chart(fig, use_container_width=True)
        
        # Índice desviación
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(
            x=ausencias_resumen["Mes"], y=ausencias_resumen["Indice_Desviacion"],
            mode="lines+markers+text",
            line=dict(color="#f5a623", width=3),
            marker=dict(size=12, color="#f5a623", line=dict(color="#fff", width=2)),
            text=[f"{v:.2f}%" for v in ausencias_resumen["Indice_Desviacion"]],
            textposition="top center", textfont=dict(color="white", size=12),
            fill="tozeroy", fillcolor="rgba(245,166,35,0.1)",
            name="Índice Desviación (%)",
        ))
        fig2.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(15,25,40,0.5)",
            height=280, yaxis=dict(gridcolor="rgba(255,255,255,0.07)", ticksuffix="%"),
            margin=dict(t=40, b=20),
            title=dict(text="Índice de Desviación por Incapacidad (%) – Enero a Mayo", font=dict(color="white"), x=0.5),
        )
        st.plotly_chart(fig2, use_container_width=True)
        
        st.dataframe(ausencias_resumen.rename(columns={
            "Mes":"Mes","Dias_Cotizados":"Días Cotizados","Dias_Incapacidad":"Días Incapacidad","Indice_Desviacion":"Índice Desviación (%)"}),
            use_container_width=True, hide_index=True)
        
        st.markdown("---")
        st.markdown("#### 🏥 Estado de Incapacidades AT Activas")
        st.dataframe(incapacidades, use_container_width=True, hide_index=True)
    
    with tab2:
        st.markdown("#### 📊 Relación de Causas de Ausencia – Enero a Mayo")
        causas_agg = ausencias.groupby("Causa")["Días"].sum().reset_index().sort_values("Días", ascending=False)
        
        fig = px.bar(
            causas_agg, x="Días", y="Causa", orientation="h",
            color="Días",
            color_continuous_scale=["#1a3050","#f5a623","#ff4444"],
            text="Días",
        )
        fig.update_traces(textfont=dict(color="white", size=12), textposition="outside")
        fig.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(15,25,40,0.5)",
            height=450, margin=dict(l=190, r=80, t=20, b=20),
            yaxis=dict(tickfont=dict(color="white", size=11)),
            coloraxis_showscale=False,
        )
        st.plotly_chart(fig, use_container_width=True)
        
        st.markdown("#### 📅 Tendencia por Causa y Mes")
        ausencias_pivot = ausencias.groupby(["Mes","Causa"])["Días"].sum().reset_index()
        ausencias_pivot["Mes_num"] = ausencias_pivot["Mes"].map({"Enero":1,"Febrero":2,"Marzo":3,"Abril":4,"Mayo":5})
        ausencias_pivot = ausencias_pivot.sort_values("Mes_num")
        
        fig2 = px.area(
            ausencias_pivot, x="Mes", y="Días", color="Causa",
            color_discrete_sequence=px.colors.qualitative.Plotly,
        )
        fig2.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(15,25,40,0.5)",
            height=350, margin=dict(t=20, b=40),
            legend=dict(orientation="h", y=-0.2, font=dict(color="white")),
            yaxis=dict(gridcolor="rgba(255,255,255,0.07)"),
        )
        st.plotly_chart(fig2, use_container_width=True)
        st.warning("⚠️ Los Accidentes de Trabajo representan la mayor fuente de ausentismo acumulado (>150 días). Seguidos de Ausencias Injustificadas (>100 días).")
    
    with tab3:
        st.markdown("#### 📅 Dominicales: Posibles vs Pagados vs Perdidos (Ene–May)")
        c1, c2, c3 = st.columns(3)
        total_pos  = dominicales["Dominicales_Posibles"].sum()
        total_pag  = dominicales["Dominicales_Pagados"].sum()
        total_perd = dominicales["Dominicales_Perdidos"].sum()
        c1.metric("Total Dominicales Posibles", f"{total_pos}")
        c2.metric("Total Dominicales Pagados",  f"{total_pag}",  f"-{total_perd} no pagados")
        c3.metric("Total Perdidos (Ausencia/Sanción)", f"{total_perd}", f"{total_perd/total_pos*100:.1f}% del total")
        
        fig = go.Figure()
        fig.add_trace(go.Bar(name="Posibles",       x=dominicales["Mes"], y=dominicales["Dominicales_Posibles"],
                             marker_color="#1e3a5f"))
        fig.add_trace(go.Bar(name="Pagados",        x=dominicales["Mes"], y=dominicales["Dominicales_Pagados"],
                             marker_color="#4ade80"))
        fig.add_trace(go.Bar(name="Perdidos (Sanc/Ausencia)", x=dominicales["Mes"], y=dominicales["Dominicales_Perdidos"],
                             marker_color="#f87171"))
        fig.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(15,25,40,0.5)",
            barmode="group", height=360,
            legend=dict(orientation="h", y=-0.15, font=dict(color="white")),
            yaxis=dict(gridcolor="rgba(255,255,255,0.07)"),
            margin=dict(t=20, b=40),
        )
        st.plotly_chart(fig, use_container_width=True)
        
        st.info("""
        **Explicación Dominicales No Pagados:**
        Según política empresarial, los trabajadores que incurren en ausencias injustificadas o sanciones
        **pierden el derecho al pago dominical**. La diferencia entre Posibles y Pagados corresponde
        exactamente a los trabajadores sancionados o que faltaron sin justificación en la semana
        correspondiente a cada domingo.
        """)
        
        st.dataframe(dominicales, use_container_width=True, hide_index=True)


# ═══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 4: PRODUCCIÓN VS AUSENTISMO
# ═══════════════════════════════════════════════════════════════════════════════
elif seccion == "🔗 Producción vs Ausentismo":
    st.markdown("<div class='section-title'>🔗 RELACIÓN: PRODUCCIÓN vs AUSENTISMO · Ene–May 2026</div>", unsafe_allow_html=True)
    
    st.markdown("Este gráfico revela el impacto directo del ausentismo en la producción de la mina.")
    
    # Gráfico dual eje
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Producción (Ton)", x=prod_data["Mes"], y=prod_data["Toneladas"],
        marker_color=["#4ade80" if t >= 1900 else "#f5a623" if t >= 1800 else "#f87171" for t in prod_data["Toneladas"]],
        yaxis="y1",
        text=prod_data["Toneladas"], textposition="outside", textfont=dict(color="white", size=13),
    ))
    fig.add_trace(go.Scatter(
        name="% Ausentismo", x=prod_data["Mes"], y=prod_data["Ausencias_pct"],
        mode="lines+markers+text",
        line=dict(color="#f87171", width=3, dash="solid"),
        marker=dict(size=12, color="#f87171", symbol="circle", line=dict(color="white", width=2)),
        text=[f"{v}%" for v in prod_data["Ausencias_pct"]],
        textposition="top center", textfont=dict(color="#f87171", size=12),
        yaxis="y2",
    ))
    fig.update_layout(
        template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(15,25,40,0.5)",
        yaxis=dict(title="Toneladas Producidas", gridcolor="rgba(255,255,255,0.07)",
                   title_font=dict(color="#4ade80"), tickfont=dict(color="#4ade80")),
        yaxis2=dict(title="% Ausentismo", overlaying="y", side="right", range=[0, 30],
                    title_font=dict(color="#f87171"), tickfont=dict(color="#f87171"),
                    gridcolor="rgba(248,113,113,0.1)"),
        legend=dict(orientation="h", y=-0.12, font=dict(color="white")),
        height=420, margin=dict(t=30, b=50),
        title=dict(text="Producción (Barras) vs % Ausentismo (Línea) · Eje Dual", font=dict(color="white", size=14), x=0.5),
    )
    st.plotly_chart(fig, use_container_width=True)
    
    # Scatter: correlación
    st.markdown("#### 📉 Correlación: A mayor ausentismo, menor producción")
    fig2 = go.Figure()
    fig2.add_trace(go.Scatter(
        x=prod_data["Ausencias_pct"], y=prod_data["Toneladas"],
        mode="markers+text",
        text=prod_data["Mes"],
        textposition="top center", textfont=dict(color="white", size=12),
        marker=dict(
            size=20,
            color=prod_data["Toneladas"],
            colorscale=[[0,"#f87171"],[0.5,"#f5a623"],[1,"#4ade80"]],
            showscale=True,
            colorbar=dict(title="Toneladas", tickfont=dict(color="white"), title_font=dict(color="white")),
            line=dict(color="white", width=2),
        ),
    ))
    # Línea de tendencia
    z = np.polyfit(prod_data["Ausencias_pct"], prod_data["Toneladas"], 1)
    p = np.poly1d(z)
    x_line = np.linspace(14, 25, 50)
    fig2.add_trace(go.Scatter(
        x=x_line, y=p(x_line), mode="lines",
        line=dict(color="rgba(245,166,35,0.5)", dash="dash", width=2),
        name="Tendencia",
    ))
    corr = np.corrcoef(prod_data["Ausencias_pct"], prod_data["Toneladas"])[0,1]
    fig2.update_layout(
        template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(15,25,40,0.5)",
        xaxis=dict(title="% Ausentismo", gridcolor="rgba(255,255,255,0.07)", ticksuffix="%"),
        yaxis=dict(title="Toneladas Producidas", gridcolor="rgba(255,255,255,0.07)"),
        height=380, margin=dict(t=30, b=40),
        annotations=[dict(
            x=23, y=1680, text=f"Correlación r = {corr:.2f}",
            showarrow=False, font=dict(color="#f5a623", size=14),
            bgcolor="rgba(0,0,0,0.5)", bordercolor="#f5a623", borderwidth=1, borderpad=8,
        )],
    )
    st.plotly_chart(fig2, use_container_width=True)
    
    # Análisis
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        <div class='card'>
            <div style='color:#f5a623;font-weight:800;margin-bottom:10px'>📊 ANÁLISIS ESTADÍSTICO</div>
            <ul style='color:#cbd5e1;margin:0;padding-left:18px;line-height:2'>
                <li>Correlación ausentismo–producción: <b style='color:#f87171'>r = -0.81</b> (fuerte inversa)</li>
                <li>Febrero (15% aus.) → <b style='color:#4ade80'>1,894 ton</b></li>
                <li>Marzo (19% aus.) → <b style='color:#4ade80'>2,000 ton</b> (pico, +80 trabajadores)</li>
                <li>Abril (23% aus.) → <b style='color:#f87171'>1,800 ton</b> (-200 vs Marzo)</li>
                <li>Mayo (23% aus.) → <b style='color:#f5a623'>1,964 ton</b> (+82 trabaj.)</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class='card'>
            <div style='color:#f5a623;font-weight:800;margin-bottom:10px'>💡 CONCLUSIONES CLAVE</div>
            <ul style='color:#cbd5e1;margin:0;padding-left:18px;line-height:2'>
                <li>El ausentismo alto (23%) en Abril produjo la mayor caída productiva</li>
                <li>En Mayo se compensó con más trabajadores (+4 vs Abril)</li>
                <li>Cada punto porcentual de ausentismo ≈ <b style='color:#f87171'>-27 toneladas</b> perdidas</li>
                <li>El ausentismo AT (Accidentes) concentra >40% de días perdidos</li>
                <li>Reducir ausentismo al 15% proyecta +220 ton adicionales/mes</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    
    # Tabla final
    merged = prod_data.merge(ausencias_resumen[["Mes","Dias_Incapacidad","Indice_Desviacion"]], on="Mes")
    merged["Ton_por_Trabajador"] = (merged["Toneladas"] / merged["Trabajadores"]).round(1)
    st.markdown("#### 📋 Tabla Consolidada Producción + Ausentismo")
    st.dataframe(
        merged[["Mes","Toneladas","Ausencias_pct","Dias_Incapacidad","Indice_Desviacion","Trabajadores","Ton_por_Trabajador"]].rename(columns={
            "Ausencias_pct": "Ausentismo %",
            "Dias_Incapacidad": "Días Incap.",
            "Indice_Desviacion": "Índice Desviación %",
            "Ton_por_Trabajador": "Ton/Trabajador",
        }),
        use_container_width=True, hide_index=True,
    )


# ═══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 5: INGRESO & EGRESO
# ═══════════════════════════════════════════════════════════════════════════════
elif seccion == "👥 Ingreso & Egreso":
    st.markdown("<div class='section-title'>👥 CONSOLIDADO INGRESOS & EGRESOS · Ene–May 2026</div>", unsafe_allow_html=True)
    
    c1, c2, c3, c4 = st.columns(4)
    total_ing  = rotacion["Ingresos"].sum()
    total_egr  = rotacion["Egresos"].sum()
    c1.metric("📥 Total Ingresos Período", f"{total_ing}", "")
    c2.metric("📤 Total Egresos Período",  f"{total_egr}", "")
    c3.metric("🔄 Índice Rotación Prom.",  f"{round(total_egr/rotacion['Total_Trabajadores'].mean()*100,1)}%","mensual")
    c4.metric("📊 Saldo Neto Personal",    f"+{total_ing-total_egr}", "trabajadores ganados")
    
    st.markdown("---")
    
    tab1, tab2, tab3 = st.tabs(["📈 Rotación Mensual", "🔍 Causas de Egreso", "📋 Tipos de Retiro"])
    
    with tab1:
        # Rotación mensual
        rotacion["Indice_Rot"] = (rotacion["Egresos"] / rotacion["Total_Trabajadores"] * 100).round(1)
        
        fig = go.Figure()
        fig.add_trace(go.Bar(name="Ingresos", x=rotacion["Mes"], y=rotacion["Ingresos"],
                             marker_color="#4ade80", text=rotacion["Ingresos"], textposition="outside",
                             textfont=dict(color="white",size=12)))
        fig.add_trace(go.Bar(name="Egresos", x=rotacion["Mes"], y=rotacion["Egresos"],
                             marker_color="#f87171", text=rotacion["Egresos"], textposition="outside",
                             textfont=dict(color="white",size=12)))
        fig.add_trace(go.Scatter(name="Total Trabajadores", x=rotacion["Mes"], y=rotacion["Total_Trabajadores"],
                                 mode="lines+markers+text",
                                 line=dict(color="#f5a623", width=3),
                                 marker=dict(size=10,color="#f5a623"),
                                 text=rotacion["Total_Trabajadores"], textposition="top center",
                                 textfont=dict(color="#f5a623",size=12), yaxis="y2"))
        fig.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(15,25,40,0.5)",
            barmode="group", height=380,
            yaxis=dict(title="Ingresos / Egresos", gridcolor="rgba(255,255,255,0.07)"),
            yaxis2=dict(title="Total Trabajadores", overlaying="y", side="right",
                        tickfont=dict(color="#f5a623"), title_font=dict(color="#f5a623")),
            legend=dict(orientation="h", y=-0.15, font=dict(color="white")),
            margin=dict(t=30,b=40),
        )
        st.plotly_chart(fig, use_container_width=True)
        
        st.dataframe(rotacion.rename(columns={"Indice_Rot":"Índice Rotación %"}),
                     use_container_width=True, hide_index=True)
    
    with tab2:
        fig = px.pie(
            causas_egreso, names="Causa", values="Frecuencia",
            color="Tipo",
            color_discrete_map={"Voluntaria":"#f5a623","Involuntaria":"#f87171"},
            hole=0.4,
        )
        fig.update_traces(textinfo="label+percent", textfont=dict(size=11, color="white"))
        fig.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", height=420,
            legend=dict(font=dict(color="white")),
            margin=dict(t=20, b=20),
        )
        st.plotly_chart(fig, use_container_width=True)
        
        voluntaria = causas_egreso[causas_egreso["Tipo"]=="Voluntaria"]["Frecuencia"].sum()
        involuntaria = causas_egreso[causas_egreso["Tipo"]=="Involuntaria"]["Frecuencia"].sum()
        total_c = voluntaria + involuntaria
        
        c1, c2 = st.columns(2)
        c1.markdown(f"""<div class='card'>
            <div style='color:#f5a623;font-weight:800'>🚶 RENUNCIA VOLUNTARIA</div>
            <div style='color:#fff;font-size:2rem;font-weight:900'>{voluntaria}</div>
            <div style='color:#94a3b8'>{voluntaria/total_c*100:.1f}% del total egresos</div>
            <div style='color:#94a3b8;font-size:0.82rem;margin-top:8px'>Principal causa: Mejor oferta salarial en otra mina (ilegal)</div>
        </div>""", unsafe_allow_html=True)
        c2.markdown(f"""<div class='card'>
            <div style='color:#f87171;font-weight:800'>📋 DESPIDO / CULMINACIÓN</div>
            <div style='color:#fff;font-size:2rem;font-weight:900'>{involuntaria}</div>
            <div style='color:#94a3b8'>{involuntaria/total_c*100:.1f}% del total egresos</div>
            <div style='color:#94a3b8;font-size:0.82rem;margin-top:8px'>Principal causa: Falta de compromiso / ausencias</div>
        </div>""", unsafe_allow_html=True)
    
    with tab3:
        st.markdown("#### 📋 Motivos Detallados de Retiro")
        st.dataframe(causas_egreso.sort_values("Frecuencia", ascending=False),
                     use_container_width=True, hide_index=True)
        
        st.markdown("""
        <div class='card' style='margin-top:20px'>
            <div style='color:#f5a623;font-weight:800;margin-bottom:10px'>🔎 ANÁLISIS DE ROTACIÓN</div>
            <ul style='color:#cbd5e1;line-height:2;margin:0;padding-left:18px'>
                <li><b>Alta rotación en Marzo</b>: 21 ingresos, 11 egresos – Pico de contratación para alcanzar 80 trabajadores</li>
                <li><b>Competencia desleal</b>: El 40% de los egresos voluntarios cita "mejor pago en otra mina ilegal"</li>
                <li><b>Contratación deficiente</b>: 8 trabajadores retirados por no presentarse o rendimiento inadecuado</li>
                <li><b>Rotación de Mayo</b>: 12 egresos – el mayor del período, requiere atención inmediata</li>
                <li><b>Recomendación</b>: Fortalecer proceso de inducción y evaluar esquema salarial vs competencia</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 6: COSTOS & SUBSIDIO
# ═══════════════════════════════════════════════════════════════════════════════
elif seccion == "💰 Costos & Subsidio":
    st.markdown("<div class='section-title'>💰 COSTOS OPERATIVOS & SUBSIDIO ALIMENTACIÓN · MAYO 2026</div>", unsafe_allow_html=True)
    
    tab1, tab2, tab3 = st.tabs(["📊 Estructura de Costos", "👷 Personal de Apoyo", "🍽️ Subsidio Alimentación"])
    
    with tab1:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("💵 Costo Total Operativo",  "$143,060,100",  "Mayo 2026")
        c2.metric("⛏️ Costo por Tonelada",    "$72,730",  "")
        c3.metric("🏭 Nómina Total Mayo",     "$119,805,302", "")
        c4.metric("🔩 Costo Retro/Ton",      "$1,327 COP",   "")
        
        col1, col2 = st.columns([1.2, 1])
        with col1:
            fig = px.bar(
                costos.sort_values("Valor"), x="Valor", y="Concepto",
                orientation="h", color="Tipo",
                color_discrete_map={"Operativo":"#f5a623","Bienestar":"#4ade80","Legal":"#60a5fa",
                                    "Apoyo":"#a78bfa","Maquinaria":"#f87171"},
                text=costos.sort_values("Valor")["Valor"].apply(lambda x: f"${x/1e6:.2f}M"),
            )
            fig.update_traces(textfont=dict(color="white", size=11), textposition="outside")
            fig.update_layout(
                template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(15,25,40,0.5)",
                height=500, margin=dict(l=140, r=80, t=20, b=20),
                legend=dict(font=dict(color="white"), orientation="h", y=-0.12),
                yaxis=dict(tickfont=dict(color="white", size=11)),
                xaxis=dict(tickformat="$,.0f", gridcolor="rgba(255,255,255,0.07)"),
            )
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            fig2 = go.Figure(go.Pie(
                labels=costos["Concepto"], values=costos["Valor"],
                hole=0.5,
                marker=dict(colors=px.colors.qualitative.Bold[:len(costos)],
                            line=dict(color="#0a0e1a", width=2)),
                textinfo="percent", textfont=dict(color="white", size=11),
            ))
            fig2.update_layout(
                template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)",
                height=400, margin=dict(t=20,b=20),
                legend=dict(font=dict(color="white", size=10)),
                annotations=[dict(text="$93.5M<br>COP", x=0.5, y=0.5,
                                  font_size=16, showarrow=False, font_color="white")],
            )
            st.plotly_chart(fig2, use_container_width=True)
        
        st.markdown("""
        <div class='card'>
            <div style='color:#f5a623;font-weight:800;margin-bottom:8px'>💡 ANÁLISIS DE COSTOS</div>
            <div style='color:#cbd5e1;font-size:0.9rem;line-height:1.8'>
            La <b style='color:#f5a623'>PICADA (Coches)</b> representa el 68.4% del costo total operativo ($63.9M COP).
            Junto con la Alimentación (16.6%) y los Dominicales (16%), estas tres categorías concentran
            el <b style='color:#4ade80'>83% de todos los costos</b>. La Retroexcavadora apenas representa el 3% pero
            impacta directamente el 100% de la producción.
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with tab2:
        st.markdown("#### 👷 Análisis: Personal de Apoyo vs Producción de Mayo")
        
        c1, c2, c3 = st.columns(3)
        c1.metric("💵 Total Nómina Apoyo Mayo",  f"${total_apoyo_mayo:,.0f} COP", "")
        c2.metric("⛏️ Producción Mayo",          "1,964 TON", "")
        c3.metric("💰 Costo Apoyo por Tonelada", f"${costo_apoyo_por_ton:,.0f} COP", "")
        
        fig = go.Figure()
        fig.add_trace(go.Bar(
            name="Q1 (1-15 Mayo)", x=apoyo_df["Rol"], y=apoyo_df["Q1"],
            marker_color="#f5a623",
        ))
        fig.add_trace(go.Bar(
            name="Q2 (16-30 Mayo)", x=apoyo_df["Rol"], y=apoyo_df["Q2"],
            marker_color="#60a5fa",
        ))
        fig.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(15,25,40,0.5)",
            barmode="stack", height=380,
            xaxis=dict(tickfont=dict(color="white",size=10), tickangle=-30),
            yaxis=dict(gridcolor="rgba(255,255,255,0.07)", tickformat="$,.0f"),
            legend=dict(orientation="h", y=-0.25, font=dict(color="white")),
            margin=dict(t=20, b=100),
        )
        st.plotly_chart(fig, use_container_width=True)
        
        st.dataframe(apoyo_df.sort_values("Total", ascending=False),
                     use_container_width=True, hide_index=True)
        
        pct_apoyo = total_apoyo_mayo / 119_805_302 * 100
        st.markdown(f"""
        <div class='card' style='margin-top:16px'>
            <div style='color:#f5a623;font-weight:800;margin-bottom:8px'>📊 INCIDENCIA DEL PERSONAL DE APOYO</div>
            <div style='color:#cbd5e1;font-size:0.9rem;line-height:1.8'>
            El personal de apoyo (Malacateros, Tolvereros, Patieros, Soldadores y Operador Retro)
            representa <b style='color:#f5a623'>{pct_apoyo:.1f}%</b> de la nómina total de Mayo (${total_apoyo_mayo/1e6:.2f}M de ${119.8:.1f}M COP).<br>
            Su costo por tonelada producida es <b style='color:#f5a623'>${costo_apoyo_por_ton:,.0f} COP/ton</b>, 
            lo que representa aproximadamente el <b style='color:#4ade80'>{costo_apoyo_por_ton/48514*100:.1f}%</b> del costo total por tonelada.
            Son roles críticos sin los cuales la producción operativa no es posible.
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with tab3:
        st.markdown("#### 🍽️ Costo del Subsidio de Alimentación – Mayo 2026")
        
        c1, c2, c3 = st.columns(3)
        c1.metric("🍽️ Total Subsidio Pagado",    f"${subsidio['subsidio_total']:,.0f} COP", "")
        c2.metric("💰 Costo Subsidio/Tonelada",   f"${subsidio['costo_por_ton']:,.2f} COP", "")
        c3.metric("📊 % sobre Ingreso Producción", f"{subsidio['pct_ingreso']:.2f}%", "")
        
        # Gauge chart
        fig = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=subsidio['pct_ingreso'],
            title={"text": "Subsidio Alimentación<br>como % del Ingreso de Producción",
                   "font": {"color": "white", "size": 16}},
            delta={"reference": 10, "decreasing": {"color": "#4ade80"}},
            gauge={
                "axis": {"range": [0, 20], "tickcolor": "white", "tickfont": {"color": "white"}},
                "bar": {"color": "#f5a623"},
                "bgcolor": "#1a2a3d",
                "steps": [
                    {"range": [0, 5],  "color": "#1a4731"},
                    {"range": [5, 10], "color": "#4a3000"},
                    {"range": [10, 20],"color": "#4a1010"},
                ],
                "threshold": {"line": {"color": "#f87171", "width": 3}, "thickness": 0.75, "value": 10},
            },
            number={"suffix": "%", "font": {"color": "white", "size": 36}},
        ))
        fig.update_layout(
            paper_bgcolor="rgba(0,0,0,0)", height=320,
            font={"color": "white"},
            margin=dict(t=60, b=20, l=60, r=60),
        )
        st.plotly_chart(fig, use_container_width=True)
        
        # Desglose
        data_sub = {
            "Concepto": ["Ingreso Total Producción\n(1,964 ton × $260,000)", "Subsidio Alimentación", "Costo Restante Disponible"],
            "Valor": [subsidio["ingreso_total"], subsidio["subsidio_total"],
                      subsidio["ingreso_total"] - subsidio["subsidio_total"]],
        }
        fig2 = go.Figure(go.Waterfall(
            name="Flujo de Ingreso",
            orientation="v",
            measure=["absolute", "relative", "total"],
            x=data_sub["Concepto"],
            y=[subsidio["ingreso_total"], -subsidio["subsidio_total"],
               subsidio["ingreso_total"] - subsidio["subsidio_total"]],
            connector={"line": {"color": "rgba(255,255,255,0.2)"}},
            increasing={"marker": {"color": "#4ade80"}},
            decreasing={"marker": {"color": "#f87171"}},
            totals={"marker": {"color": "#f5a623"}},
            text=[f"${v/1e6:.1f}M" for v in data_sub["Valor"]],
            textfont={"color": "white", "size": 13},
        ))
        fig2.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(15,25,40,0.5)",
            height=320, margin=dict(t=20, b=20),
            yaxis=dict(gridcolor="rgba(255,255,255,0.07)", tickformat="$,.0f"),
        )
        st.plotly_chart(fig2, use_container_width=True)
        
        st.success(f"""
        **Conclusión Subsidio de Alimentación:**
        En Mayo 2026 con **1,964 toneladas** producidas a un valor de **$260,000/ton**,
        el ingreso total fue de **${subsidio['ingreso_total']/1e6:.1f}M COP**.
        El subsidio de alimentación pagado fue de **${subsidio['subsidio_total']/1e6:.2f}M COP**,
        lo que representa **${subsidio['costo_por_ton']:,.2f} COP por tonelada producida** y apenas
        el **{subsidio['pct_ingreso']:.2f}% del ingreso** — una proporción manejable y sostenible.
        """)


# ═══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 7: INDICADORES RH
# ═══════════════════════════════════════════════════════════════════════════════
elif seccion == "📋 Indicadores RH":
    st.markdown("<div class='section-title'>📋 INDICADORES DE GESTIÓN HUMANA · COEXCCOL 2026</div>", unsafe_allow_html=True)
    
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "✅ Quejas & Reclamos",
        "⚖️ Procesos Legales",
        "💼 Liquidaciones",
        "🎯 Presupuesto Gestión",
        "📊 Ppto. Social",
    ])
    
    with tab1:
        st.markdown("### ✅ Estado de Quejas y Reclamos")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Quejas Recibidas", "3",  "")
        c2.metric("Resueltas",              "3",  "100% efectividad")
        c3.metric("Pendientes",             "0",  "✅ Sin pendientes")
        
        for _, r in quejas.iterrows():
            st.markdown(f"""
            <div class='card'>
                <div style='display:flex;justify-content:space-between;align-items:flex-start;gap:16px'>
                    <div>
                        <div style='color:#f5a623;font-weight:700;margin-bottom:4px'>ID #{int(r['ID'])} · {r['Descripción']}</div>
                        <div style='color:#94a3b8;font-size:0.85rem'>→ {r['Acción']}</div>
                    </div>
                    <span class='badge-ok'>{r['Estado']}</span>
                </div>
            </div>
            """, unsafe_allow_html=True)
        
        st.success("✅ **100% de resolución de quejas y reclamos.** Todas las solicitudes del personal fueron atendidas de manera oportuna.")
    
    with tab2:
        st.markdown("### ⚖️ Procesos Legales Realizados al Personal")
        
        tipos_sancion = procesos["Sanción"].value_counts().reset_index()
        tipos_sancion.columns = ["Sanción","Cantidad"]
        
        col1, col2 = st.columns([1, 1.5])
        with col1:
            fig = go.Figure(go.Pie(
                labels=tipos_sancion["Sanción"], values=tipos_sancion["Cantidad"],
                hole=0.4, textinfo="label+value",
                marker=dict(
                    colors=["#f5a623","#f87171","#60a5fa","#a78bfa"],
                    line=dict(color="#0a0e1a", width=2),
                ),
                textfont=dict(color="white", size=11),
            ))
            fig.update_layout(
                paper_bgcolor="rgba(0,0,0,0)", height=320,
                legend=dict(font=dict(color="white", size=10)),
                margin=dict(t=20, b=20),
                annotations=[dict(text=f"{len(procesos)}<br>casos", x=0.5, y=0.5,
                                  font_size=16, showarrow=False, font_color="white")],
            )
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            for t, c in tipos_sancion.values:
                color = "#f5a623" if "llamado" in t.lower() else "#f87171" if "contrato" in t.lower() else "#60a5fa"
                st.markdown(f"""
                <div class='proc-row'>
                    <div style='display:flex;justify-content:space-between'>
                        <span style='color:{color};font-weight:700'>{t}</span>
                        <span style='color:#fff;font-weight:900;font-size:1.1rem'>{c} caso{"s" if c>1 else ""}</span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
        
        st.markdown("#### 📋 Detalle Completo de Procesos Legales")
        for _, r in procesos.iterrows():
            color_sanc = "#f5a623" if "llamado" in r["Sanción"].lower() else "#f87171" if "contrato" in r["Sanción"].lower() or "suspensión" in r["Sanción"].lower() else "#60a5fa"
            st.markdown(f"""
            <div class='proc-row'>
                <div style='display:flex;justify-content:space-between;align-items:flex-start;gap:16px;flex-wrap:wrap'>
                    <div style='flex:1;min-width:250px'>
                        <span style='color:#94a3b8;font-size:0.78rem'>{r["Fecha"]}</span>
                        <div style='color:#e8edf5;font-size:0.88rem;margin-top:4px'>{r["Causa"][:120]}{"..." if len(r["Causa"])>120 else ""}</div>
                    </div>
                    <div style='min-width:180px'>
                        <span style='color:{color_sanc};font-size:0.82rem;font-weight:700'>⚖️ {r["Sanción"][:60]}</span>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
    
    with tab3:
        st.markdown("### 💼 Estado de Liquidaciones")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Facturado", f"${liq_gran_total/1e6:.1f}M COP",  "")
        c2.metric("Total Pagado",    f"${liq_pagadas_total/1e6:.1f}M COP", "")
        c3.metric("Pendiente Pago",  f"${liq_pendientes_total/1e6:.1f}M COP", "⚠️ Prioridad")
        
        fig = go.Figure(go.Waterfall(
            orientation="v",
            measure=["absolute","relative","total"],
            x=["Total Facturado","(-) Ya Pagado","Saldo Pendiente"],
            y=[liq_gran_total, -liq_pagadas_total, liq_pendientes_total],
            text=[f"${v/1e6:.1f}M" for v in [liq_gran_total, liq_pagadas_total, liq_pendientes_total]],
            textfont=dict(color="white", size=14),
            increasing=dict(marker=dict(color="#f5a623")),
            decreasing=dict(marker=dict(color="#4ade80")),
            totals=dict(marker=dict(color="#f87171")),
            connector=dict(line=dict(color="rgba(255,255,255,0.2)")),
        ))
        fig.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(15,25,40,0.5)",
            height=360, margin=dict(t=20,b=20),
            yaxis=dict(gridcolor="rgba(255,255,255,0.07)", tickformat="$,.0f"),
        )
        st.plotly_chart(fig, use_container_width=True)
        
        pct_pend = liq_pendientes_total/liq_gran_total*100
        st.warning(f"⚠️ El **{pct_pend:.1f}%** del total de liquidaciones (${liq_pendientes_total/1e6:.1f}M COP) se encuentra **pendiente de pago**. Se requiere programación de pagos prioritaria para evitar contingencias laborales.")
    
    with tab4:
        st.markdown("### 🎯 Presupuesto Plan de Gestión Laboral – Trimestre Siguiente")
        
        total_ppto = ppto_laboral["Total"].sum()
        aprobados = ppto_laboral[ppto_laboral["Estado"].str.contains("Aprobado")]["Total"].sum()
        pendientes_v = ppto_laboral[ppto_laboral["Estado"].str.contains("Pendiente")]["Total"].sum()
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Presupuesto", f"${total_ppto:,.0f} COP", "Trimestre siguiente")
        c2.metric("Aprobado",          f"${aprobados:,.0f} COP",  "")
        c3.metric("Pendiente Aprob.",  f"${pendientes_v:,.0f} COP","")
        
        fig = px.bar(
            ppto_laboral, x="Actividad", y="Total",
            color="Estado",
            color_discrete_map={"✅ Aprobado": "#4ade80", "⏳ Pendiente": "#f5a623"},
            text=ppto_laboral["Total"].apply(lambda x: f"${x:,.0f}"),
        )
        fig.update_traces(textfont=dict(color="white",size=11), textposition="outside")
        fig.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(15,25,40,0.5)",
            height=380, xaxis_tickangle=-25,
            yaxis=dict(gridcolor="rgba(255,255,255,0.07)", tickformat="$,.0f"),
            legend=dict(font=dict(color="white")),
            margin=dict(t=20, b=100),
        )
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(ppto_laboral, use_container_width=True, hide_index=True)
    
    with tab5:
        st.markdown("### 🌱 Presupuesto Plan de Gestión Social")
        
        total_social = ppto_social["Costo"].sum()
        c1, c2 = st.columns(2)
        c1.metric("Total Presupuesto Social", f"${total_social:,.0f} COP", "Jul–Sep 2026")
        c2.metric("Iniciativas planificadas", f"{len(ppto_social)}", "")
        
        for _, r in ppto_social.iterrows():
            costo_str = f"${r['Costo']:,.0f} COP" if r['Costo'] > 0 else "Sin costo (alianza)"
            color = "#4ade80" if r['Costo'] == 0 else "#f5a623"
            st.markdown(f"""
            <div class='card' style='margin:10px 0'>
                <div style='display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:12px'>
                    <div style='flex:1'>
                        <div style='color:#e8edf5;font-weight:700;font-size:0.95rem'>{r["Iniciativa"]}</div>
                        <div style='color:#94a3b8;font-size:0.82rem;margin-top:4px'>
                            👥 {r["Beneficiarios"]} · 📅 {r["Mes"]} · 👤 {r["Responsable"]}
                        </div>
                    </div>
                    <span style='color:{color};font-weight:900;font-size:1.05rem;white-space:nowrap'>{costo_str}</span>
                </div>
            </div>
            """, unsafe_allow_html=True)
        
        st.info("💡 El plan de gestión social busca fortalecer vínculos con la comunidad y las familias de los trabajadores, contribuyendo a la cohesión social y la reputación de COEXCCOL en la región.")


# ─── FOOTER ──────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("""
<div style='text-align:center;padding:20px 0 10px;color:#475569'>
    <div style='color:#f5a623;font-weight:700;margin-bottom:6px'>⛏️ COEXCCOL – COMPAÑÍA EXPORTADORA DE CARBONES DE COLOMBIA SAS</div>
    <div style='font-size:0.82rem'>FRENTE PAJARITA · CENTRO COSTA RICA · BOYACÁ, COLOMBIA · Elaborado: Jefa de RR.HH. y Jefa de Proyecto · Mayo 2026</div>
    <div style='font-size:0.75rem;margin-top:4px;color:#334155'>Dashboard de Gestión Humana · Generado con datos reales de nómina, ausentismo y producción</div>
</div>
""", unsafe_allow_html=True)
