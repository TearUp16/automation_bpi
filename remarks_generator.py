import streamlit as st
import streamlit.components.v1 as components


def render_page_header(title: str, subtitle: str) -> None:
    st.markdown(f'<div class="app-title">{title}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="app-subtitle">{subtitle}</div>', unsafe_allow_html=True)


def escape_js(text: str) -> str:
    return text.replace("\\", "\\\\").replace("`", "\\`").replace("</", "<\\/")


def render_remarks_generator() -> None:
    render_page_header("Remarks Generator", "Collections · Volare & F1 Format")

    if "clear_trigger" not in st.session_state:
        st.session_state.clear_trigger = 0

    ct = st.session_state.clear_trigger
    key = lambda k: f"{k}_{ct}"

    RFD_OPTIONS = {
        "": "— Select RFD —",
        "INSU": "INSU – Financial Difficulty (Lack/Awaiting Funds)",
        "UNEM": "UNEM – Unemployed",
        "BUSL": "BUSL – Business Closure/Slowdown",
        "OVLK": "OVLK – Forgot/Overlooked",
        "SICK": "SICK – Family Member/Client Hospitalized",
        "PRIO": "PRIO – Prioritize Other Expenses",
        "OOTC": "OOTC – OOTC/OOT",
        "NONB": "NONB – Payment Channel Issue",
        "TYPH": "TYPH – Natural Disaster or Calamity",
        "NLRS": "NLRS – NRS/LRS",
        "DISP": "DISP – Dispute",
        "PAPO": "PAPO – Pending Nego/Reversal",
        "PAPB": "PAPB – Pending Nego/Reversal",
        "BS": "BS – Business Slowdown",
        "CC": "CC – Contested Charges/Dispute",
        "DC": "DC – Deceased Cardholder",
        "FE": "FE – Family Emergency",
        "FTP": "FTP – Forgot to Pay/Overlooked Payment",
        "IL": "IL – ILL/Sickness in the family",
        "OL": "OL – Old age/Retired",
        "OT": "OT – Out of the country/Migrated",
        "OV": "OV – Over extended/Lack or Short of funds",
        "SK": "SK – Skip in both RA and BA/No lead",
        "UK": "UK – Unknown reason/Third party contact only",
        "UN": "UN – Loss of job/Unemployment",
        "AR": "AR – Awaiting Remittance",
        "AC": "AC – Awaiting Collection",
        "CV": "CV – Victim of Calamity (typhoon, fire, earthquake, pandemic or war)",
    }

    SRC_OPTIONS = {
        "": "— Select SRC —",
        "EML": "EML – Email",
        "FLD": "FLD – Field",
        "SMS": "SMS – SMS",
        "CAL": "CAL – Call",
    }

    col1, col2 = st.columns([1, 2])
    with col1:
        confidence = st.selectbox(
            "Confidence Level",
            ["", "1_", "0_"],
            key=key("conf"),
            format_func=lambda x: "— Select —" if x == "" else x
        )
    with col2:
        number_email = st.text_input(
            "Number / Email",
            placeholder="e.g. 09176308527 or email@example.com",
            key=key("num")
        )

    col3, col4 = st.columns(2)
    with col3:
        rfd_key = st.selectbox(
            "RFD – Reason for Delinquency",
            list(RFD_OPTIONS.keys()),
            format_func=lambda x: RFD_OPTIONS[x],
            key=key("rfd")
        )
    with col4:
        src_key = st.selectbox(
            "SRC – Source of Contact",
            list(SRC_OPTIONS.keys()),
            format_func=lambda x: SRC_OPTIONS[x],
            key=key("src")
        )

    soi = st.text_input(
        "SOI – Source of Income",
        placeholder="e.g. business, job, etc...",
        key=key("soi")
    )

    remarks = st.text_area(
        "Remarks",
        placeholder="Enter your remarks here…",
        key=key("remarks"),
        height=90
    )

    v_parts = []
    if confidence or number_email:
        v_parts.append(f"{confidence}{number_email}".strip())
    if rfd_key:
        v_parts.append(f"RFD:{rfd_key}")
    if src_key:
        v_parts.append(f"SRC:{src_key}")
    if soi:
        v_parts.append(f"SOI:{soi}")
    if remarks:
        v_parts.append(f"REMARKS:{remarks}")

    v_text = " | ".join([part for part in v_parts if part])
    f1_text = " - ".join([part for part in [number_email, remarks] if part])

    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("#### Generated Remarks")

    st.markdown(f"""
    <div class="output-card">
      <div class="output-label">🔴 For Volare</div>
      <div class="output-text {'output-empty' if not v_text else ''}">
        {v_text if v_text else 'Fill in the fields above to generate…'}
      </div>
    </div>
    <div class="output-card">
      <div class="output-label">⚫ For F1</div>
      <div class="output-text {'output-empty' if not f1_text else ''}">
        {f1_text if f1_text else 'Fill in Number/Email and Remarks…'}
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("""
    <style>
    div[data-testid="stButton"] > button {
        width: 100% !important;
        min-height: 42px !important;
        border-radius: 10px !important;
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
    }

    div[data-testid="stButton"] > button * {
        color: #ffffff !important;
        fill: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
    }

    div[data-testid="stButton"] > button:hover {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
    }
    </style>
    """, unsafe_allow_html=True)

    v_safe = escape_js(v_text)
    f1_safe = escape_js(f1_text)

    b1, b2, b3 = st.columns([1, 1, 1], gap="small")

    with b1:
        if st.button("🗑️ CLEAR", key="clear_btn", use_container_width=True):
            st.session_state.clear_trigger += 1
            st.rerun()

    with b2:
        components.html(f"""
        <style>
          @import url('https://fonts.googleapis.com/css2?family=Syne:wght@700&display=swap');
          html, body {{
            margin: 0;
            padding: 0;
            background: transparent;
          }}
          .btn-wrap {{
            width: 100%;
            padding-top: 1px;
          }}
          button {{
            width: 100%;
            min-height: 42px;
            padding: 0.65rem 1rem;
            background: linear-gradient(135deg,#ff3b30,#b30000);
            color: #ffffff;
            border: none;
            border-radius: 10px;
            font-family: 'Syne', sans-serif;
            font-weight: 700;
            font-size: 0.78rem;
            letter-spacing: 0.05em;
            cursor: pointer;
            transition: opacity 0.2s ease;
            box-sizing: border-box;
          }}
          button:hover {{
            opacity: 0.9;
          }}
          button.ok {{
            background: linear-gradient(135deg,#0f8b4c,#0a5e34);
          }}
        </style>
        <div class="btn-wrap">
          <button id="btn" onclick="copyText()">📋 VOLARE REMARKS</button>
        </div>
        <script>
          function copyText() {{
            const txt = `{v_safe}`;
            if (!txt.trim()) {{
              alert('Nothing to copy — fill in the fields first.');
              return;
            }}
            navigator.clipboard.writeText(txt).then(() => {{
              const b = document.getElementById('btn');
              b.textContent = '✅ COPIED!';
              b.classList.add('ok');
              setTimeout(() => {{
                b.textContent = '📋 VOLARE REMARKS';
                b.classList.remove('ok');
              }}, 2000);
            }}).catch(() => {{
              const ta = document.createElement('textarea');
              ta.value = txt;
              ta.style.position = 'fixed';
              ta.style.opacity = '0';
              document.body.appendChild(ta);
              ta.select();
              document.execCommand('copy');
              document.body.removeChild(ta);
              const b = document.getElementById('btn');
              b.textContent = '✅ COPIED!';
              b.classList.add('ok');
              setTimeout(() => {{
                b.textContent = '📋 VOLARE REMARKS';
                b.classList.remove('ok');
              }}, 2000);
            }});
          }}
        </script>
        """, height=50)

    with b3:
        components.html(f"""
        <style>
          @import url('https://fonts.googleapis.com/css2?family=Syne:wght@700&display=swap');
          html, body {{
            margin: 0;
            padding: 0;
            background: transparent;
          }}
          .btn-wrap {{
            width: 100%;
            padding-top: 1px;
          }}
          button {{
            width: 100%;
            min-height: 42px;
            padding: 0.65rem 1rem;
            background: linear-gradient(135deg,#ff3b30,#b30000);
            color: #ffffff;
            border: none;
            border-radius: 10px;
            font-family: 'Syne', sans-serif;
            font-weight: 700;
            font-size: 0.78rem;
            letter-spacing: 0.05em;
            cursor: pointer;
            transition: opacity 0.2s ease;
            box-sizing: border-box;
          }}
          button:hover {{
            opacity: 0.9;
          }}
          button.ok {{
            background: linear-gradient(135deg,#0f8b4c,#0a5e34);
          }}
        </style>
        <div class="btn-wrap">
          <button id="btn" onclick="copyText()">📋 F1 REMARKS</button>
        </div>
        <script>
          function copyText() {{
            const txt = `{f1_safe}`;
            if (!txt.trim()) {{
              alert('Nothing to copy — fill in Number/Email and Remarks first.');
              return;
            }}
            navigator.clipboard.writeText(txt).then(() => {{
              const b = document.getElementById('btn');
              b.textContent = '✅ COPIED!';
              b.classList.add('ok');
              setTimeout(() => {{
                b.textContent = '📋 F1 REMARKS';
                b.classList.remove('ok');
              }}, 2000);
            }}).catch(() => {{
              const ta = document.createElement('textarea');
              ta.value = txt;
              ta.style.position = 'fixed';
              ta.style.opacity = '0';
              document.body.appendChild(ta);
              ta.select();
              document.execCommand('copy');
              document.body.removeChild(ta);
              const b = document.getElementById('btn');
              b.textContent = '✅ COPIED!';
              b.classList.add('ok');
              setTimeout(() => {{
                b.textContent = '📋 F1 REMARKS';
                b.classList.remove('ok');
              }}, 2000);
            }});
          }}
        </script>
        """, height=50)

    with st.expander("📖 RFD Quick Reference"):
        ref_data = {k: v.split("–")[-1].strip() for k, v in RFD_OPTIONS.items() if k}
        cols = st.columns(2)
        items = list(ref_data.items())
        half = len(items) // 2

        for i, (code, desc) in enumerate(items):
            with cols[0 if i < half else 1]:
                st.markdown(f"**`{code}`** — {desc}")