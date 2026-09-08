import streamlit as st
from groq import Groq
from ddgs import DDGS
import pypdf
import docx

st.set_page_config(page_title="Foxsay", page_icon="🦊", layout="centered")

st.markdown("""
<style>
    div.stButton > button,
    div[data-testid="stFormSubmitButton"] > button {
        background-color: #FF6B35 !important;
        color: white !important;
        border-radius: 20px !important;
        border: none !important;
        padding: 10px 24px !important;
        font-weight: bold !important;
    }
    .chat-bubble-user {
        background-color: #2b313e;
        padding: 12px 16px;
        border-radius: 15px;
        margin: 8px 0;
        color: white;
    }
    .chat-bubble-ai {
        background-color: #FF6B35;
        padding: 12px 16px;
        border-radius: 15px;
        margin: 8px 0;
        color: white;
    }
    .file-chip {
        background-color: #2b313e;
        padding: 6px 12px;
        border-radius: 10px;
        display: inline-block;
        margin-bottom: 8px;
        font-size: 0.85em;
        color: #FF6B35;
    }
</style>
""", unsafe_allow_html=True)

st.title("🦊 Foxsay")
st.caption("AI agent yang bisa jawab pertanyaan, search internet, dan baca file")

api_key = st.secrets["GROQ_API_KEY"]
client = Groq(api_key=api_key)

if "history" not in st.session_state:
    st.session_state.history = []
if "isi_file" not in st.session_state:
    st.session_state.isi_file = ""
if "nama_file" not in st.session_state:
    st.session_state.nama_file = ""
if "show_uploader" not in st.session_state:
    st.session_state.show_uploader = False

SYSTEM_STYLE = """Kamu adalah Foxsay, AI agent yang asik dan ekspresif. Jika ditanya siapa
namamu, jawab bahwa kamu adalah Foxsay.

Gaya bicaramu santai dan hidup, pakai kosakata khas berikut (bukan kata-kata umum
seperti "goks" atau "gaskeun" biasa):
- "vhom" = ekspresi kagum/heboh (pengganti "goks")
- "bhapp" = ajakan semangat/mulai (pengganti "gaskeun")
- "nyampe" atau "appruved" = ekspresi keren/oke
- "pun10" = tunggu sebentar

Selain istilah khusus di atas, tetap boleh pakai kata santai umum lain seperti "yap!",
"jujur", "wih", "dahlah", dll — tapi isinya tetap informatif dan berbobot, jangan sampai
dangkal cuma karena gayanya santai.

ATURAN KHUSUS UNTUK KATA "ALIAS": pakai MAKSIMAL 1 kali per jawaban, dan HANYA kalau
memang ada momen yang pas/related untuk itu, jangan dipaksakan di setiap jawaban.
Kalau dipakai, taruh di akhir jawaban sebagai semacam penutup/punchline, bukan di
tengah-tengah. Selalu tulis kapital: ALIAS. Contohnya: "Bandung kota metropolitan
terbesar ketiga, ALIAS kalau macet nomor 1 se-Jabar." Kalau nggak ada momen yang pas
untuk itu, nggak usah dipaksain pakai ALIAS sama sekali.

Kamu paham berbagai bahasa daerah Indonesia (Jawa, Sunda, Betawi, dll) kalau user
menggunakannya dalam pertanyaan, tapi kamu tetap menjawab pakai Bahasa Indonesia gaya
santai di atas, bukan ikut logat daerah."""

def baca_pdf(file):
    reader = pypdf.PdfReader(file)
    teks = ""
    for page in reader.pages:
        teks += page.extract_text() + "\n"
    return teks

def baca_docx(file):
    doc = docx.Document(file)
    return "\n".join([para.text for para in doc.paragraphs])

def cari_internet(query):
    try:
        hasil = DDGS().text(query, max_results=5)
        teks_hasil = ""
        for item in hasil:
            teks_hasil += f"- {item['title']}: {item['body']}\n"
        return teks_hasil if teks_hasil else "(tidak ada hasil pencarian)"
    except Exception:
        return "(pencarian gagal, jawab pakai pengetahuan umum saja)"

def agent(pertanyaan, konteks_file=""):
    if konteks_file:
        sumber_info = f"""Isi file yang diupload user ({st.session_state.nama_file}):
{konteks_file[:8000]}"""
    else:
        hasil_search = cari_internet(pertanyaan)
        sumber_info = f"""Info dari internet (kalau ada):
{hasil_search}"""

    prompt = f"""{SYSTEM_STYLE}

{sumber_info}

Pertanyaan: {pertanyaan}

Jawab dengan gaya di atas."""

    response = client.chat.completions.create(
        model="openai/gpt-oss-120b",
        messages=[{"role": "user", "content": prompt}]
    )
    return response.choices[0].message.content

# Riwayat chat ditampilkan dulu (terbaru di atas nanti setelah input)

if st.session_state.nama_file:
    st.markdown(f'<span class="file-chip">📎 {st.session_state.nama_file}</span>', unsafe_allow_html=True)

if st.session_state.show_uploader:
    uploaded_file = st.file_uploader("Upload PDF atau Word", type=["pdf", "docx"], label_visibility="collapsed")
    if uploaded_file is not None and uploaded_file.name != st.session_state.nama_file:
        with st.spinner("🦊 Foxsay lagi baca file..."):
            try:
                if uploaded_file.name.endswith(".pdf"):
                    isi = baca_pdf(uploaded_file)
                elif uploaded_file.name.endswith(".docx"):
                    isi = baca_docx(uploaded_file)
                else:
                    isi = ""
                st.session_state.isi_file = isi
                st.session_state.nama_file = uploaded_file.name
                st.session_state.show_uploader = False
                st.success(f"File '{uploaded_file.name}' berhasil dibaca!")
                st.rerun()
            except Exception as e:
                st.error(f"Gagal baca file: {e}")

if st.session_state.nama_file:
    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("Rangkum file ini"):
            with st.spinner("🦊 Foxsay lagi ngerangkum..."):
                jawaban = agent("Tolong rangkum isi file ini secara singkat dan jelas.", st.session_state.isi_file)
            st.session_state.history.append(("user", f"[Minta rangkuman: {st.session_state.nama_file}]"))
            st.session_state.history.append(("ai", jawaban))
            st.rerun()
    with col_b:
        if st.button("Hapus file"):
            st.session_state.isi_file = ""
            st.session_state.nama_file = ""
            st.rerun()

st.write("Tanya apa aja ke Foxsay:")
with st.form("tanya_form", clear_on_submit=True):
    col_plus, col_input, col_submit = st.columns([1, 5, 1.3])
    with col_plus:
        attach_clicked = st.form_submit_button("➕")
    with col_input:
        pertanyaan = st.text_input("Tanya apa aja ke Foxsay:", label_visibility="collapsed")
    with col_submit:
        submitted = st.form_submit_button("Tanya")

if attach_clicked:
    st.session_state.show_uploader = not st.session_state.show_uploader
    st.rerun()

if submitted and pertanyaan:
    with st.spinner("🦊 Foxsay lagi mikir..."):
        jawaban = agent(pertanyaan, st.session_state.isi_file)
    st.session_state.history.append(("user", pertanyaan))
    st.session_state.history.append(("ai", jawaban))

for role, teks in reversed(st.session_state.history):
    if role == "user":
        st.markdown(f'<div class="chat-bubble-user">🙋 {teks}</div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="chat-bubble-ai">🦊 {teks}</div>', unsafe_allow_html=True)
