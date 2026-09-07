import streamlit as st
from groq import Groq
from ddgs import DDGS

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
</style>
""", unsafe_allow_html=True)

st.title("🦊 Foxsay")
st.caption("AI agent yang bisa jawab pertanyaan pakai info terkini dari internet")

api_key = st.secrets["GROQ_API_KEY"]
client = Groq(api_key=api_key)

if "history" not in st.session_state:
    st.session_state.history = []

def cari_internet(query):
    try:
        hasil = DDGS().text(query, max_results=5)
        teks_hasil = ""
        for item in hasil:
            teks_hasil += f"- {item['title']}: {item['body']}\n"
        return teks_hasil if teks_hasil else "(tidak ada hasil pencarian)"
    except Exception:
        return "(pencarian gagal, jawab pakai pengetahuan umum saja)"

def agent(pertanyaan):
    hasil_search = cari_internet(pertanyaan)
    prompt = f"""Kamu adalah Foxsay, AI agent yang ramah.
Jika ditanya siapa namamu, jawab bahwa kamu adalah Foxsay.

Info dari internet (kalau ada):
{hasil_search}

Pertanyaan: {pertanyaan}

Jawab dengan jelas dan ringkas dalam Bahasa Indonesia."""

    response = client.chat.completions.create(
        model="openai/gpt-oss-120b",
        messages=[{"role": "user", "content": prompt}]
    )
    return response.choices[0].message.content

with st.form("tanya_form", clear_on_submit=True):
    pertanyaan = st.text_input("Tanya apa aja ke Foxsay:")
    submitted = st.form_submit_button("Tanya")

if submitted and pertanyaan:
    with st.spinner("🦊 Foxsay lagi mikir..."):
        jawaban = agent(pertanyaan)
    st.session_state.history.append(("user", pertanyaan))
    st.session_state.history.append(("ai", jawaban))

for role, teks in reversed(st.session_state.history):
    if role == "user":
        st.markdown(f'<div class="chat-bubble-user">🙋 {teks}</div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="chat-bubble-ai">🦊 {teks}</div>', unsafe_allow_html=True)