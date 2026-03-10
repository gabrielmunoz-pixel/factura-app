import streamlit as st
import anthropic
import base64
import json
import io
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="FacturaAI",
    page_icon="🧾",
    layout="wide",
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Mono:wght@300;400;500&display=swap');

html, body, [class*="css"] { font-family: 'Syne', sans-serif; }

.main { background: #f5f2eb; }
.block-container { padding-top: 2rem; padding-bottom: 3rem; }

h1 { font-size: 2rem !important; font-weight: 800 !important; letter-spacing: -1px !important; }

.stButton > button {
    background: #0a0a0f !important;
    color: #f5f2eb !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    font-size: 15px !important;
    padding: 12px 28px !important;
    width: 100%;
    transition: background 0.2s !important;
}
.stButton > button:hover { background: #e8521a !important; }

.info-card {
    background: #fff;
    border: 1.5px solid #ddd9d0;
    border-radius: 10px;
    padding: 14px 16px;
    margin-bottom: 8px;
}
.info-label {
    font-family: 'DM Mono', monospace;
    font-size: 10px;
    text-transform: uppercase;
    letter-spacing: 0.6px;
    color: #8a8680;
    margin-bottom: 3px;
}
.info-value {
    font-size: 14px;
    font-weight: 700;
    color: #0a0a0f;
    line-height: 1.3;
}
.section-label {
    font-family: 'DM Mono', monospace;
    font-size: 10px;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    color: #8a8680;
    border-bottom: 1px solid #ddd9d0;
    padding-bottom: 6px;
    margin-bottom: 12px;
    margin-top: 20px;
}
.total-box {
    background: #f5f2eb;
    border-radius: 10px;
    padding: 18px;
}
.total-row {
    display: flex;
    justify-content: space-between;
    font-size: 13px;
    margin-bottom: 6px;
}
.total-grand {
    display: flex;
    justify-content: space-between;
    font-size: 20px;
    font-weight: 800;
    color: #e8521a;
    border-top: 2px solid #0a0a0f;
    padding-top: 10px;
    margin-top: 6px;
    font-family: 'DM Mono', monospace;
}
.note-box {
    background: #fff8f5;
    border: 1.5px solid #f5c4b2;
    border-radius: 8px;
    padding: 12px 14px;
    font-size: 13px;
    color: #c44a1a;
    font-weight: 600;
    margin-top: 12px;
}
.stFileUploader > div { border-radius: 10px !important; border: 2px dashed #ddd9d0 !important; background: #f5f2eb !important; }
.stRadio > div { gap: 8px; }
div[data-testid="stMetricValue"] { font-family: 'DM Mono', monospace !important; font-size: 18px !important; }
</style>
""", unsafe_allow_html=True)

# ── Prompts ───────────────────────────────────────────────────────────────────
PROMPT_STRUCTURED = """Eres un experto en lectura de facturas chilenas y latinoamericanas.
Analiza esta factura con máxima precisión y extrae TODOS los datos.
Responde SOLO con JSON puro, sin texto adicional, sin backticks, sin explicaciones.

Estructura exacta requerida:
{
  "emisor": { "nombre": "", "rut": "", "direccion": "", "giro": "" },
  "receptor": { "nombre": "", "rut": "", "direccion": "", "comuna": "", "ciudad": "", "giro": "" },
  "factura": { "numero": "", "fecha": "", "tipo": "", "condicion_venta": "", "vendedor": "", "nro_ov": "" },
  "items": [
    { "codigo": "", "descripcion": "", "unidad": "", "cantidad": 0, "precio_unitario": 0, "descuento": 0, "total": 0 }
  ],
  "totales": { "neto": 0, "exento": 0, "iva": 0, "otros_impuestos": [], "total": 0 },
  "notas": "",
  "despacho": ""
}

Lee con cuidado cada línea. Los montos son números sin puntos ni símbolo $.
Si no encuentras un campo usa "" para texto y 0 para números. No inventes datos."""

PROMPT_ITEMS = """Extrae SOLO los productos/items de esta factura con máxima precisión.
Responde SOLO con JSON puro, sin texto adicional:
{"items":[{"codigo":"","descripcion":"","unidad":"","cantidad":0,"precio_unitario":0,"total":0}],"total":0}"""

PROMPT_SUMMARY = """Resume esta factura en los campos principales.
Responde SOLO con JSON puro, sin texto adicional:
{"emisor":"","receptor":"","numero_factura":"","fecha":"","total":0,"cantidad_items":0,"descripcion_breve":""}"""


# ── Helpers ───────────────────────────────────────────────────────────────────
def parse_json_safe(text: str) -> dict:
    clean = re.sub(r"```json|```", "", text).strip()
    try:
        return json.loads(clean)
    except json.JSONDecodeError:
        match = re.search(r"\{[\s\S]*\}", clean)
        if match:
            return json.loads(match.group())
        raise ValueError("No se encontró JSON válido en la respuesta")


def fmt_money(n) -> str:
    if not n and n != 0:
        return "—"
    return f"${int(n):,}".replace(",", ".")


def analyze_invoice(api_key: str, file_bytes: bytes, media_type: str, mode: str) -> dict:
    client = anthropic.Anthropic(api_key=api_key)
    b64 = base64.standard_b64encode(file_bytes).decode("utf-8")

    prompt = {"structured": PROMPT_STRUCTURED, "items": PROMPT_ITEMS, "summary": PROMPT_SUMMARY}[mode]

    if media_type == "application/pdf":
        content = [
            {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": b64}},
            {"type": "text", "text": prompt},
        ]
    else:
        content = [
            {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64}},
            {"type": "text", "text": prompt},
        ]

    response = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=4096,
        messages=[{"role": "user", "content": content}],
    )
    raw = response.content[0].text
    return parse_json_safe(raw), raw


def build_excel(data: dict) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Factura"

    hdr_fill = PatternFill("solid", start_color="1F4E79")
    row_fill = PatternFill("solid", start_color="DEEAF1")
    tot_fill = PatternFill("solid", start_color="FFF2CC")
    hdr_font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    bold = Font(bold=True, name="Calibri", size=10)
    normal = Font(name="Calibri", size=10)
    thin = Side(style="thin")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Title
    ws.merge_cells("A1:G1")
    ws["A1"] = f"FACTURA {data.get('factura', {}).get('numero', '')} — {data.get('emisor', {}).get('nombre', '')}"
    ws["A1"].font = Font(bold=True, name="Calibri", size=13, color="1F4E79")
    ws["A1"].alignment = Alignment(horizontal="center")

    # Meta info
    meta = [
        ("Fecha", data.get("factura", {}).get("fecha", "")),
        ("Receptor", data.get("receptor", {}).get("nombre", "")),
        ("RUT Receptor", data.get("receptor", {}).get("rut", "")),
        ("Vendedor", data.get("factura", {}).get("vendedor", "")),
        ("Cond. Venta", data.get("factura", {}).get("condicion_venta", "")),
    ]
    for i, (k, v) in enumerate(meta, start=3):
        ws.cell(row=i, column=1, value=k).font = bold
        ws.cell(row=i, column=2, value=v).font = normal

    # Header
    headers = ["Código", "Descripción", "U/M", "Cantidad", "Precio Unit.", "Dto.", "Total"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=9, column=col, value=h)
        c.fill = hdr_fill; c.font = hdr_font
        c.alignment = Alignment(horizontal="center")
        c.border = bdr
    ws.row_dimensions[9].height = 22

    # Items
    items = data.get("items", [])
    for i, item in enumerate(items, start=10):
        fill = row_fill if i % 2 == 0 else PatternFill("solid", start_color="FFFFFF")
        vals = [item.get("codigo",""), item.get("descripcion",""), item.get("unidad",""),
                item.get("cantidad",0), item.get("precio_unitario",0),
                item.get("descuento",0), item.get("total",0)]
        for col, val in enumerate(vals, 1):
            c = ws.cell(row=i, column=col, value=val)
            c.font = normal; c.fill = fill; c.border = bdr
            if col in (4,5,6,7): c.alignment = Alignment(horizontal="right")

    # Totals
    t = data.get("totales", {})
    trow = len(items) + 11
    total_data = [
        ("Monto Neto", t.get("neto", 0)),
        ("Monto Exento", t.get("exento", 0)),
        ("IVA 19%", t.get("iva", 0)),
    ]
    for otros in (t.get("otros_impuestos") or []):
        total_data.append((otros.get("nombre","Impuesto"), otros.get("monto",0)))
    total_data.append(("TOTAL", t.get("total", 0)))

    for j, (label, val) in enumerate(total_data, start=trow):
        is_grand = label == "TOTAL"
        c_lbl = ws.cell(row=j, column=5, value=label)
        c_val = ws.cell(row=j, column=7, value=val)
        c_lbl.font = Font(bold=True, name="Calibri", size=11 if is_grand else 10, color="C44A1A" if is_grand else "000000")
        c_val.font = c_lbl.font
        c_lbl.alignment = Alignment(horizontal="right")
        c_val.alignment = Alignment(horizontal="right")
        if is_grand:
            c_val.fill = tot_fill
        c_val.border = bdr

    if data.get("notas"):
        ws.cell(row=trow + len(total_data) + 1, column=1, value=f"NOTA: {data['notas']}").font = Font(italic=True, color="C44A1A", name="Calibri")

    # Column widths
    for col, w in zip(range(1, 8), [14, 42, 7, 10, 14, 8, 14]):
        ws.column_dimensions[get_column_letter(col)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def build_csv(data: dict) -> str:
    lines = ["Codigo,Descripcion,Unidad,Cantidad,Precio Unitario,Total"]
    for item in data.get("items", []):
        lines.append(f"\"{item.get('codigo','')}\",\"{item.get('descripcion','')}\","
                     f"\"{item.get('unidad','')}\",{item.get('cantidad',0)},"
                     f"{item.get('precio_unitario',0)},{item.get('total',0)}")
    return "\n".join(lines)


# ── UI ────────────────────────────────────────────────────────────────────────
st.markdown("""
<div style='display:flex;align-items:baseline;gap:10px;margin-bottom:4px'>
  <span style='font-size:32px;font-weight:800;letter-spacing:-1px'>
    <span style='color:#e8521a'>Factura</span>AI
  </span>
  <span style='font-family:monospace;font-size:11px;background:#0a0a0f;color:#f5f2eb;padding:3px 8px;border-radius:3px;letter-spacing:1px'>BETA</span>
</div>
<p style='font-family:monospace;font-size:12px;color:#8a8680;margin-bottom:28px'>
  Lector inteligente de facturas · Powered by Claude
</p>
""", unsafe_allow_html=True)

# ── Sidebar: API Key ──────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuración")
    api_key = st.text_input("Anthropic API Key", type="password", placeholder="sk-ant-...")
    st.caption("Tu clave nunca se almacena. [Obtener clave →](https://console.anthropic.com/)")
    st.markdown("---")
    st.markdown("**Modos de extracción**")
    st.markdown("- **Completo** — toda la factura")
    st.markdown("- **Solo Items** — tabla de productos")
    st.markdown("- **Resumen** — datos clave")
    st.markdown("---")
    st.markdown("**Formatos soportados**")
    st.markdown("JPG · PNG · WEBP · PDF")

# ── Main columns ──────────────────────────────────────────────────────────────
col_left, col_right = st.columns([1, 1.6], gap="large")

with col_left:
    st.markdown('<div class="section-label">01 — Subir Factura</div>', unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "Arrastra o selecciona tu factura",
        type=["jpg", "jpeg", "png", "webp", "pdf"],
        label_visibility="collapsed",
    )

    if uploaded:
        if uploaded.type.startswith("image/"):
            st.image(uploaded, use_container_width=True)
        else:
            st.success(f"📋 PDF cargado: **{uploaded.name}**")

    mode_labels = {"Completo": "structured", "Solo Items": "items", "Resumen": "summary"}
    mode_sel = st.radio("Modo de extracción", list(mode_labels.keys()), horizontal=True, label_visibility="collapsed")
    mode = mode_labels[mode_sel]

    analyze_btn = st.button("🔍 Analizar Factura", disabled=not uploaded or not api_key)

    if not api_key:
        st.info("👈 Ingresa tu API Key en el panel lateral para comenzar.")

# ── Analysis ──────────────────────────────────────────────────────────────────
with col_right:
    st.markdown('<div class="section-label">02 — Datos Extraídos</div>', unsafe_allow_html=True)

    if analyze_btn and uploaded and api_key:
        with st.spinner("Claude está leyendo tu factura..."):
            try:
                file_bytes = uploaded.read()
                media_type = uploaded.type
                if media_type in ("image/jpg",):
                    media_type = "image/jpeg"

                data, raw_text = analyze_invoice(api_key, file_bytes, media_type, mode)
                st.session_state["result"] = data
                st.session_state["raw"] = raw_text
                st.success("✅ Factura analizada correctamente")
            except Exception as e:
                st.error(f"❌ Error: {e}")
                st.session_state.pop("result", None)

    if "result" in st.session_state:
        data = st.session_state["result"]

        # ── Export buttons ────────────────────────────────────────────────────
        ec1, ec2, ec3 = st.columns(3)
        with ec1:
            xlsx_bytes = build_excel(data)
            st.download_button("⬇ Excel", xlsx_bytes, "factura.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)
        with ec2:
            csv_str = build_csv(data)
            st.download_button("⬇ CSV", csv_str, "factura.csv", "text/csv", use_container_width=True)
        with ec3:
            st.download_button("⬇ JSON", st.session_state["raw"], "factura.json",
                               "application/json", use_container_width=True)

        # ── Info general ──────────────────────────────────────────────────────
        if mode in ("structured", "summary"):
            emisor = data.get("emisor", {})
            receptor = data.get("receptor", {})
            factura = data.get("factura", {})

            fields = [
                ("Emisor", emisor.get("nombre") or data.get("emisor","")),
                ("RUT Emisor", emisor.get("rut","")),
                ("Receptor", receptor.get("nombre") or data.get("receptor","")),
                ("RUT Receptor", receptor.get("rut","")),
                ("N° Factura", factura.get("numero") or data.get("numero_factura","")),
                ("Fecha", factura.get("fecha") or data.get("fecha","")),
                ("Cond. Venta", factura.get("condicion_venta","")),
                ("Vendedor", factura.get("vendedor","")),
            ]
            fields = [(k, v) for k, v in fields if v]

            if fields:
                st.markdown('<div class="section-label">Información General</div>', unsafe_allow_html=True)
                cols = st.columns(2)
                for i, (k, v) in enumerate(fields):
                    with cols[i % 2]:
                        st.markdown(f"""
                        <div class="info-card">
                            <div class="info-label">{k}</div>
                            <div class="info-value">{v}</div>
                        </div>""", unsafe_allow_html=True)

        # ── Items table ───────────────────────────────────────────────────────
        items = data.get("items", [])
        if items:
            st.markdown(f'<div class="section-label">Líneas de Detalle — {len(items)} items</div>', unsafe_allow_html=True)
            import pandas as pd
            df = pd.DataFrame(items)
            col_map = {"codigo": "Código", "descripcion": "Descripción", "unidad": "U/M",
                       "cantidad": "Cantidad", "precio_unitario": "P. Unitario", "descuento": "Dto.", "total": "Total"}
            df = df.rename(columns=col_map)
            for money_col in ["P. Unitario", "Total", "Dto."]:
                if money_col in df.columns:
                    df[money_col] = df[money_col].apply(lambda x: fmt_money(x) if x else "—")
            st.dataframe(df, use_container_width=True, hide_index=True)

        # ── Totals ────────────────────────────────────────────────────────────
        totales = data.get("totales", {})
        grand = totales.get("total") or data.get("total", 0)
        if grand:
            st.markdown('<div class="section-label">Totales</div>', unsafe_allow_html=True)
            rows_html = ""
            for label, key in [("Monto Neto", "neto"), ("Monto Exento", "exento"), ("IVA 19%", "iva")]:
                val = totales.get(key, 0)
                if val:
                    rows_html += f'<div class="total-row"><span>{label}</span><span style="font-family:monospace;font-weight:600">{fmt_money(val)}</span></div>'
            for otro in (totales.get("otros_impuestos") or []):
                rows_html += f'<div class="total-row"><span>{otro.get("nombre","Impuesto")}</span><span style="font-family:monospace;font-weight:600">{fmt_money(otro.get("monto",0))}</span></div>'

            st.markdown(f"""
            <div class="total-box">
                {rows_html}
                <div class="total-grand">
                    <span>TOTAL</span>
                    <span>{fmt_money(grand)}</span>
                </div>
            </div>""", unsafe_allow_html=True)

        # ── Notes ─────────────────────────────────────────────────────────────
        notas = data.get("notas","")
        if notas:
            st.markdown(f'<div class="note-box">⚠️ {notas}</div>', unsafe_allow_html=True)

        # ── Raw JSON expander ─────────────────────────────────────────────────
        with st.expander("Ver JSON raw"):
            st.code(st.session_state.get("raw",""), language="json")

    elif "result" not in st.session_state and not analyze_btn:
        st.markdown("""
        <div style='text-align:center;padding:80px 40px;color:#8a8680'>
            <div style='font-size:52px;margin-bottom:14px;opacity:0.3'>🔍</div>
            <div style='font-size:16px;font-weight:700;margin-bottom:8px'>Sin datos aún</div>
            <div style='font-family:monospace;font-size:12px;line-height:1.7'>
                Sube una foto o PDF de factura<br>e ingresa tu API Key para comenzar
            </div>
        </div>""", unsafe_allow_html=True)
