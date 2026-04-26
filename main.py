import streamlit as st

from remarks_generator import render_remarks_generator
from report_generator import render_report_generator
from rfd import render_rfd_mapper
from response_rate import response_rate


st.set_page_config(page_title="BPI Automation", layout="wide")


def apply_global_theme() -> None:
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Mono:wght@400;500&display=swap');

    :root {
        --bg-main: #ffffff;
        --bg-soft: #f0f0f0;
        --bg-panel: #ffffff;
        --bg-input: #ffffff;
        --bg-card: #f9f9f9;
        --text-main: #333333;
        --text-soft: #666666;
        --text-dim: #999999;
        --accent: #ff3b30;
        --accent-dark: #b30000;
        --border: #ddd;
        --success: #0f8b4c;
    }

    html, body, [data-testid="stAppViewContainer"] {
        background-color: var(--bg-main);
        color: var(--text-main);
        font-family: 'Syne', sans-serif;
    }

    * {
        font-family: 'Syne', sans-serif;
    }

    [data-testid="stHeader"] {
        background: transparent;
    }

    [data-testid="stToolbar"] {
        right: 0.75rem;
    }

    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 2rem !important;
        max-width: 1200px;
    }

    .app-title {
        font-size: 3rem;
        font-weight: 800;
        color: #000000;
        letter-spacing: -0.03em;
        margin-bottom: 1rem;
        text-transform: uppercase;
    }

    .app-subtitle {
        color: var(--text-soft);
        font-size: 1rem;
        margin-bottom: 1.25rem;
    }

    [data-testid="stSidebar"] {
        background: var(--bg-panel);
        border-right: 1px solid var(--border);
    }

    [data-testid="stSidebar"] > div:first-child {
        padding-top: 1.2rem;
    }

    [data-testid="stSidebar"] * {
        color: var(--text-main);
    }

    .sidebar-brand {
        font-size: 1.85rem;
        font-weight: 800;
        line-height: 1;
        margin-bottom: 1.2rem;
        color: #000000;
        letter-spacing: -0.03em;
        background: none;
        -webkit-text-fill-color: initial;
    }

    .sidebar-section {
        background: var(--bg-card);
        border: 1px solid var(--border);
        border-radius: 14px;
        padding: 0.9rem 0.85rem 0.95rem 0.85rem;
        margin-bottom: 0.9rem;
        box-shadow: 0 6px 24px rgba(0, 0, 0, 0.1);
    }

    .sidebar-label {
        color: var(--text-soft);
        font-size: 0.72rem;
        text-transform: uppercase;
        letter-spacing: 0.14em;
        font-weight: 700;
        margin-bottom: 0.45rem;
    }

    .stButton > button,
    .stDownloadButton > button {
        width: 100%;
        border: none !important;
        border-radius: 10px !important;
        padding: 0.7rem 1rem !important;
        font-weight: 700 !important;
        letter-spacing: 0.04em;
        background: linear-gradient(135deg, var(--accent), var(--accent-dark));
        color: white !important;
        transition: all 0.2s ease;
        box-shadow: 0 8px 22px rgba(179, 0, 0, 0.18);
    }

    .stButton > button:hover,
    .stDownloadButton > button:hover {
        transform: translateY(-1px);
        filter: brightness(1.06);
    }

    .stAlert {
        border-radius: 12px !important;
    }

    h1, h2, h3, h4, p, label, .stMarkdown, .stText, div {
        color: var(--text-main);
    }

    [data-testid="stTextInput"] input,
    [data-testid="stTextArea"] textarea,
    [data-testid="stNumberInput"] input,
    [data-baseweb="select"] > div {
        background: var(--bg-input) !important;
        color: var(--text-main) !important;
        border: 1px solid var(--border) !important;
        border-radius: 10px !important;
    }

    [data-testid="stTextInput"] input::placeholder,
    [data-testid="stTextArea"] textarea::placeholder {
        color: var(--text-dim) !important;
        opacity: 1 !important;
        font-family: 'DM Mono', monospace !important;
    }

    [data-testid="stTextInput"] input:focus,
    [data-testid="stTextArea"] textarea:focus {
        border-color: var(--accent) !important;
        box-shadow: 0 0 0 2px rgba(255, 59, 48, 0.15) !important;
    }

    .output-card {
        background: var(--bg-card);
        border: 1px solid var(--border);
        border-radius: 14px;
        padding: 1.15rem 1.2rem;
        margin-bottom: 0.9rem;
    }

    .output-label {
        font-weight: 700;
        margin-bottom: 0.45rem;
    }

    .output-text {
        font-family: 'DM Mono', monospace !important;
        font-size: 0.9rem;
        color: var(--text-main);
        line-height: 1.6;
    }

    .output-empty {
        color: var(--text-dim);
        font-style: italic;
    }
    </style>
    """, unsafe_allow_html=True)


apply_global_theme()

st.sidebar.markdown('<div class="sidebar-brand">BPI Tools</div>', unsafe_allow_html=True)

st.sidebar.markdown('<div class="sidebar-section">', unsafe_allow_html=True)
st.sidebar.markdown('<div class="sidebar-label">Module</div>', unsafe_allow_html=True)

main_mode = st.sidebar.selectbox(
    "Choose Module",
    [
        "📋 Remarks Generator",
        "📊 Report Generator",
    ],
    label_visibility="collapsed"
)

st.sidebar.markdown('</div>', unsafe_allow_html=True)

if main_mode == "📊 Report Generator":
    st.sidebar.markdown('<div class="sidebar-section">', unsafe_allow_html=True)
    st.sidebar.markdown('<div class="sidebar-label">Functions</div>', unsafe_allow_html=True)

    report_mode = st.sidebar.selectbox(
        "Choose Function",
        [
            "📂 DRR CSV Processor",
            "✅ POSITIVE Status",
            "❌ NEGATIVE Status",
            "🏍️ FIELD RESULT",
            "🧾 RFD Mapper",
            "RESPONSE RATE"
        ],
        label_visibility="collapsed"
    )

    st.sidebar.markdown('</div>', unsafe_allow_html=True)
else:
    report_mode = None

if main_mode == "📋 Remarks Generator":
    render_remarks_generator()
elif main_mode == "📊 Report Generator" and report_mode is not None:
    if report_mode == "🧾 RFD Mapper":
        render_rfd_mapper()
    if report_mode == "RESPONSE RATE":
        response_rate()
    else:
        render_report_generator(report_mode)
