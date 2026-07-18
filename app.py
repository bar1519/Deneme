import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
from datetime import datetime
from openai import OpenAI
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer

st.set_page_config(
    page_title="Onay Mekanizmalı Çoklu Ajan Raporlama Sistemi",
    layout="wide",
)

st.title("Çoklu Ajan Raporlama Sistemi (Adım Adım Onaylı)")

NVIDIA_BASE_URL = "https://integrate.api.nvidia.com/v1"


def get_secret(name: str, default: str = "") -> str:
    try:
        value = st.secrets.get(name, default)
        if value is None:
            return default
        text = str(value).strip().strip('"').strip("'")
        if text.startswith("$") or text in ("", "None"):
            return default
        return text
    except Exception:
        return default


def get_secret_float(name: str, default: float) -> float:
    raw = get_secret(name, "")
    if not raw:
        return default
    try:
        return float(raw)
    except ValueError:
        return default


def get_secret_int(name: str, default: int) -> int:
    raw = get_secret(name, "")
    if not raw:
        return default
    try:
        return int(float(raw))
    except ValueError:
        return default


# Her ajan: kendi API key + model + parametreler (Secrets)
# 1: Llama 3.1 70B | 2: Mistral Large | 3: DeepSeek V4 Pro
AGENT_CONFIG = {
    1: {
        "name": "Llama 3.1 70B",
        "api_key": get_secret("AGENT1_API_KEY"),
        "model": get_secret("AGENT1_MODEL", "meta/llama-3.1-70b-instruct"),
        "temperature": get_secret_float("AGENT1_TEMPERATURE", 0.2),
        "top_p": get_secret_float("AGENT1_TOP_P", 0.7),
        "max_tokens": get_secret_int("AGENT1_MAX_TOKENS", 1024),
        "frequency_penalty": None,
        "presence_penalty": None,
        "extra_body": None,
    },
    2: {
        "name": "Mistral Large",
        "api_key": get_secret("AGENT2_API_KEY"),
        "model": get_secret(
            "AGENT2_MODEL", "mistralai/mistral-large-3-675b-instruct-2512"
        ),
        "temperature": get_secret_float("AGENT2_TEMPERATURE", 0.15),
        "top_p": get_secret_float("AGENT2_TOP_P", 1.0),
        "max_tokens": get_secret_int("AGENT2_MAX_TOKENS", 2048),
        "frequency_penalty": get_secret_float("AGENT2_FREQUENCY_PENALTY", 0.0),
        "presence_penalty": get_secret_float("AGENT2_PRESENCE_PENALTY", 0.0),
        "extra_body": None,
    },
    3: {
        "name": "DeepSeek V4 Pro",
        "api_key": get_secret("AGENT3_API_KEY"),
        "model": get_secret("AGENT3_MODEL", "deepseek-ai/deepseek-v4-pro"),
        "temperature": get_secret_float("AGENT3_TEMPERATURE", 1.0),
        "top_p": get_secret_float("AGENT3_TOP_P", 0.95),
        "max_tokens": get_secret_int("AGENT3_MAX_TOKENS", 16384),
        "frequency_penalty": None,
        "presence_penalty": None,
        "extra_body": {"chat_template_kwargs": {"thinking": False}},
    },
}

# --- YAN MENÜ ---
st.sidebar.header("Ajan / API Durumu")
st.sidebar.caption("Anahtarlar Streamlit Secrets'tan okunur.")
for no, cfg in AGENT_CONFIG.items():
    key_ok = (
        bool(cfg["api_key"])
        and cfg["api_key"].startswith("nvapi-")
        and len(cfg["api_key"]) > 10
    )
    st.sidebar.write(f"Ajan {no} ({cfg['name']}): {'key OK' if key_ok else 'key YOK'}")
    st.sidebar.caption(cfg["model"])


# --- SESSION STATE ---
defaults = {
    "current_step": 1,  # 1=upload, 2=ajan1, 3=ajan2, 4=ajan3, 5=final
    "raw_data": None,
    "agent1_output": None,
    "agent2_output": None,
    "agent3_output": None,
    "final_report": None,
    "finalized_at": None,  # 1 | 2 | 3
}
for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value

AGENT_PROMPTS = {
    1: (
        "Sen 1. Ajan (Veri Çıkarım ve Yapılandırma)sın. Görevin, sana sağlanan ham veri "
        "ve metin yığınlarını okumak, gereksiz tekrarları temizlemek ve veriyi mantıklı "
        "başlıklar altında yapılandırılmış teknik bir özet haline getirmektir. "
        "Yorum ekleme, sadece veriyi düzenle."
    ),
    2: (
        "Sen 2. Ajan (Teknik Analiz)sin. Sana hem ham kaynak veri hem de 1. ajanın "
        "yapılandırılmış çıktısı verilir. Eğilimleri, anormallikleri, kritik değerleri "
        "ve teknik çıkarımları belirleyerek detaylı bir analitik rapor taslağı oluştur. "
        "1. ajan çıktısını temel al, ham veriyle doğrula ve zenginleştir."
    ),
    3: (
        "Sen 3. Ajan (Nihai Rapor)sun. Sana ham kaynak veri, 1. ajan çıktısı ve 2. ajan "
        "analizi verilir. Bunları birleştirerek tutarlı, eksiksiz ve sunuma hazır nihai "
        "teknik raporu yaz. Çelişkileri çöz, tekrarları kaldır, net bölümler halinde sun."
    ),
}


def init_nvidia_client(agent_no: int = 1):
    cfg = AGENT_CONFIG[agent_no]
    api_key = cfg["api_key"]
    if not api_key or not api_key.startswith("nvapi-"):
        st.error(
            f"Ajan {agent_no} için geçerli API anahtarı yok. "
            f"Secrets'a AGENT{agent_no}_API_KEY = \"nvapi-...\" ekleyin."
        )
        return None
    return OpenAI(base_url=NVIDIA_BASE_URL, api_key=api_key)


def parse_excel(file):
    df = pd.read_excel(file)
    return df.to_string(index=False)


def parse_docx(file):
    doc = Document(io.BytesIO(file.read()))
    return "\n".join(para.text for para in doc.paragraphs)


def call_agent(client, agent_no: int, user_content: str, max_tokens=None):
    cfg = AGENT_CONFIG[agent_no]
    kwargs = {
        "model": cfg["model"],
        "messages": [
            {"role": "system", "content": AGENT_PROMPTS[agent_no]},
            {"role": "user", "content": user_content},
        ],
        "temperature": cfg["temperature"],
        "top_p": cfg["top_p"],
        "max_tokens": cfg["max_tokens"] if max_tokens is None else max_tokens,
        "stream": False,
    }
    if cfg.get("frequency_penalty") is not None:
        kwargs["frequency_penalty"] = cfg["frequency_penalty"]
    if cfg.get("presence_penalty") is not None:
        kwargs["presence_penalty"] = cfg["presence_penalty"]
    if cfg.get("extra_body"):
        kwargs["extra_body"] = cfg["extra_body"]

    completion = client.chat.completions.create(**kwargs)
    content = completion.choices[0].message.content
    return content if content is not None else ""


def build_agent2_prompt(raw_data: str, agent1_output: str) -> str:
    return (
        "=== HAM KAYNAK VERİ (1. ajana giden aynı veri) ===\n"
        f"{raw_data}\n\n"
        "=== 1. AJAN ÇIKTISI ===\n"
        f"{agent1_output}\n\n"
        "Yukarıdaki ham veri ile 1. ajan çıktısını birlikte kullanarak teknik analiz yap."
    )


def build_agent3_prompt(raw_data: str, agent1_output: str, agent2_output: str) -> str:
    return (
        "=== HAM KAYNAK VERİ (ilk veri) ===\n"
        f"{raw_data}\n\n"
        "=== 1. AJAN ÇIKTISI ===\n"
        f"{agent1_output}\n\n"
        "=== 2. AJAN ÇIKTISI ===\n"
        f"{agent2_output}\n\n"
        "Ham veri + 1. ajan + 2. ajan çıktılarını birleştirerek nihai raporu oluştur."
    )


def reset_pipeline():
    st.session_state.current_step = 1
    st.session_state.raw_data = None
    st.session_state.agent1_output = None
    st.session_state.agent2_output = None
    st.session_state.agent3_output = None
    st.session_state.final_report = None
    st.session_state.finalized_at = None


def finalize_with(output_text: str, agent_no: int):
    st.session_state.final_report = output_text
    st.session_state.finalized_at = agent_no
    st.session_state.current_step = 5
    st.rerun()


def _report_title(agent_no: int) -> str:
    return f"Teknik Rapor (Ajan {agent_no} çıktısı)"


def build_docx_bytes(report_text: str, agent_no: int) -> bytes:
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    title = doc.add_heading(_report_title(agent_no), level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    meta = doc.add_paragraph(
        f"Oluşturulma: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
    )
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("")

    for line in (report_text or "").splitlines():
        stripped = line.strip()
        if not stripped:
            doc.add_paragraph("")
            continue
        if stripped.startswith("### "):
            doc.add_heading(stripped[4:], level=3)
        elif stripped.startswith("## "):
            doc.add_heading(stripped[3:], level=2)
        elif stripped.startswith("# "):
            doc.add_heading(stripped[2:], level=1)
        else:
            p = doc.add_paragraph(stripped)
            for run in p.runs:
                run.font.size = Pt(11)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def _register_pdf_font() -> str:
    """Türkçe karakter destekli sistem fontu bulmaya çalışır; yoksa Helvetica."""
    candidates = [
        (r"C:\Windows\Fonts\arial.ttf", "ArialTR"),
        (r"C:\Windows\Fonts\calibri.ttf", "CalibriTR"),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "DejaVu"),
    ]
    for path, name in candidates:
        try:
            pdfmetrics.registerFont(TTFont(name, path))
            return name
        except Exception:
            continue
    return "Helvetica"


def build_pdf_bytes(report_text: str, agent_no: int) -> bytes:
    font_name = _register_pdf_font()
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=2 * cm,
        rightMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "ReportTitle",
        parent=styles["Heading1"],
        fontName=font_name,
        fontSize=16,
        spaceAfter=8,
        alignment=1,
    )
    meta_style = ParagraphStyle(
        "ReportMeta",
        parent=styles["Normal"],
        fontName=font_name,
        fontSize=9,
        spaceAfter=16,
        alignment=1,
    )
    body_style = ParagraphStyle(
        "ReportBody",
        parent=styles["Normal"],
        fontName=font_name,
        fontSize=11,
        leading=15,
        spaceAfter=6,
    )
    h1_style = ParagraphStyle(
        "ReportH1",
        parent=styles["Heading1"],
        fontName=font_name,
        fontSize=14,
        spaceBefore=12,
        spaceAfter=6,
    )
    h2_style = ParagraphStyle(
        "ReportH2",
        parent=styles["Heading2"],
        fontName=font_name,
        fontSize=12,
        spaceBefore=10,
        spaceAfter=4,
    )

    def esc(text: str) -> str:
        return (
            text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )

    story = [
        Paragraph(esc(_report_title(agent_no)), title_style),
        Paragraph(
            esc(f"Oluşturulma: {datetime.now().strftime('%d.%m.%Y %H:%M')}"),
            meta_style,
        ),
        Spacer(1, 0.3 * cm),
    ]

    for line in (report_text or "").splitlines():
        stripped = line.strip()
        if not stripped:
            story.append(Spacer(1, 0.25 * cm))
            continue
        if stripped.startswith("### "):
            story.append(Paragraph(esc(stripped[4:]), h2_style))
        elif stripped.startswith("## "):
            story.append(Paragraph(esc(stripped[3:]), h2_style))
        elif stripped.startswith("# "):
            story.append(Paragraph(esc(stripped[2:]), h1_style))
        else:
            story.append(Paragraph(esc(stripped), body_style))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def show_progress():
    labels = ["Yükleme", "1. Ajan", "2. Ajan", "3. Ajan", "Sonuç"]
    step = st.session_state.current_step
    cols = st.columns(5)
    for i, (col, label) in enumerate(zip(cols, labels), start=1):
        with col:
            if i < step:
                st.success(label)
            elif i == step:
                st.info(f"→ {label}")
            else:
                st.caption(label)


show_progress()

# --- ADIM 1: DOSYA YÜKLEME ---
if st.session_state.current_step == 1:
    st.header("Adım 1: Kaynak Dosyaların Yüklenmesi")
    st.caption(
        "Ham veri önce 1. ajana gider; aynı ham veri daha sonra 2. ve 3. ajana da "
        "bağlam olarak iletilir. Her ajan sonrası erken sonuçlandırma yapılabilir."
    )

    uploaded_files = st.file_uploader(
        "Analiz edilecek Excel (.xlsx) veya Word (.docx) dosyalarını seçin",
        type=["xlsx", "docx"],
        accept_multiple_files=True,
    )

    if uploaded_files and st.button("Dosyaları İşle ve 1. Ajanı Çalıştır"):
        client = init_nvidia_client(1)
        if client:
            with st.spinner("Dosyalar okunuyor, 1. ajan yapılandırıyor..."):
                combined = []
                for file in uploaded_files:
                    if file.name.endswith(".xlsx"):
                        content = parse_excel(file)
                        combined.append(f"\n--- {file.name} (Excel) ---\n{content}")
                    elif file.name.endswith(".docx"):
                        content = parse_docx(file)
                        combined.append(f"\n--- {file.name} (Word) ---\n{content}")

                raw_data = "\n".join(combined)
                st.session_state.raw_data = raw_data

                try:
                    st.session_state.agent1_output = call_agent(
                        client,
                        1,
                        f"Yapılandırılacak Ham Veri:\n{raw_data}",
                    )
                    st.session_state.current_step = 2
                    st.rerun()
                except Exception as e:
                    st.error(f"API çağrısı sırasında hata: {e}")

# --- ADIM 2: 1. AJAN ÇIKTISI ---
elif st.session_state.current_step == 2:
    st.header("Adım 2: 1. Ajan Çıktısı")
    st.info(
        "1. ajan ham veriyi yapılandırdı. İyi ise hemen sonuçlandırabilir veya "
        "2. ajana (ham veri + bu çıktı) gönderebilirsiniz."
    )

    with st.expander("Ham kaynak veri (önizleme)", expanded=False):
        st.text(st.session_state.raw_data[:4000] + ("..." if len(st.session_state.raw_data or "") > 4000 else ""))

    edited_a1 = st.text_area(
        "1. Ajan çıktısı (düzenlenebilir)",
        value=st.session_state.agent1_output or "",
        height=400,
        key="edit_agent1",
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("Baştan başla"):
            reset_pipeline()
            st.rerun()
    with c2:
        if st.button("Bu çıktıyla sonuçlandır", type="primary"):
            st.session_state.agent1_output = edited_a1
            finalize_with(edited_a1, 1)
    with c3:
        if st.button("Onayla → 2. Ajana gönder"):
            st.session_state.agent1_output = edited_a1
            client = init_nvidia_client(2)
            if client:
                with st.spinner("2. ajan analiz ediyor (ham veri + 1. ajan çıktısı)..."):
                    try:
                        prompt = build_agent2_prompt(
                            st.session_state.raw_data,
                            st.session_state.agent1_output,
                        )
                        st.session_state.agent2_output = call_agent(
                            client, 2, prompt
                        )
                        st.session_state.current_step = 3
                        st.rerun()
                    except Exception as e:
                        st.error(f"API çağrısı sırasında hata: {e}")

# --- ADIM 3: 2. AJAN ÇIKTISI ---
elif st.session_state.current_step == 3:
    st.header("Adım 3: 2. Ajan Çıktısı")
    st.info(
        "2. ajan ham veri + 1. ajan çıktısını analiz etti. "
        "Erken sonuçlandırabilir veya 3. ajana devam edebilirsiniz."
    )

    with st.expander("Girdi özeti", expanded=False):
        st.markdown("**Ham veri** (kısaltılmış)")
        st.text((st.session_state.raw_data or "")[:2000])
        st.markdown("**1. ajan çıktısı**")
        st.text(st.session_state.agent1_output or "")

    edited_a2 = st.text_area(
        "2. Ajan çıktısı (düzenlenebilir)",
        value=st.session_state.agent2_output or "",
        height=400,
        key="edit_agent2",
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("← 1. ajan çıktısına dön"):
            st.session_state.current_step = 2
            st.rerun()
    with c2:
        if st.button("Bu çıktıyla sonuçlandır", type="primary"):
            st.session_state.agent2_output = edited_a2
            finalize_with(edited_a2, 2)
    with c3:
        if st.button("Onayla → 3. Ajana gönder"):
            st.session_state.agent2_output = edited_a2
            client = init_nvidia_client(3)
            if client:
                with st.spinner(
                    "3. ajan nihai raporu yazıyor (ham + 1. ajan + 2. ajan)..."
                ):
                    try:
                        prompt = build_agent3_prompt(
                            st.session_state.raw_data,
                            st.session_state.agent1_output,
                            st.session_state.agent2_output,
                        )
                        st.session_state.agent3_output = call_agent(
                            client, 3, prompt
                        )
                        st.session_state.current_step = 4
                        st.rerun()
                    except Exception as e:
                        st.error(f"API çağrısı sırasında hata: {e}")

# --- ADIM 4: 3. AJAN ÇIKTISI ---
elif st.session_state.current_step == 4:
    st.header("Adım 4: 3. Ajan Çıktısı")
    st.warning(
        "3. ajan ham veri + 1. ajan + 2. ajan çıktılarını birleştirdi. "
        "Düzenleyip sonuçlandırın."
    )

    with st.expander("Girdi özeti", expanded=False):
        st.markdown("**Ham veri** (kısaltılmış)")
        st.text((st.session_state.raw_data or "")[:1500])
        st.markdown("**1. ajan**")
        st.text(st.session_state.agent1_output or "")
        st.markdown("**2. ajan**")
        st.text(st.session_state.agent2_output or "")

    edited_a3 = st.text_area(
        "3. Ajan çıktısı (düzenlenebilir)",
        value=st.session_state.agent3_output or "",
        height=400,
        key="edit_agent3",
    )

    c1, c2 = st.columns(2)
    with c1:
        if st.button("← 2. ajan çıktısına dön"):
            st.session_state.current_step = 3
            st.rerun()
    with c2:
        if st.button("Nihai rapor olarak sonuçlandır", type="primary"):
            st.session_state.agent3_output = edited_a3
            finalize_with(edited_a3, 3)

# --- ADIM 5: SONUÇ ---
elif st.session_state.current_step == 5:
    st.header("Nihai Sonuç")
    finished = st.session_state.finalized_at
    st.success(f"Rapor {finished}. ajan çıktısı üzerinden sonuçlandırıldı.")

    report_text = st.session_state.final_report or ""
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    base_name = f"rapor_ajan{finished}_{stamp}"

    st.text_area(
        "Nihai rapor",
        value=report_text,
        height=450,
        key="final_view",
    )

    st.subheader("İndir")
    d1, d2, d3 = st.columns(3)
    with d1:
        st.download_button(
            "TXT indir",
            data=report_text,
            file_name=f"{base_name}.txt",
            mime="text/plain",
            use_container_width=True,
        )
    with d2:
        st.download_button(
            "Word (.docx) indir",
            data=build_docx_bytes(report_text, finished),
            file_name=f"{base_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with d3:
        st.download_button(
            "PDF indir",
            data=build_pdf_bytes(report_text, finished),
            file_name=f"{base_name}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    if st.button("Yeni analiz başlat"):
        reset_pipeline()
        st.rerun()
