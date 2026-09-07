import streamlit as st
import google.generativeai as genai

# ===== KONFIGURASI AWAL =====
st.set_page_config(page_title="Foxsay", layout="wide")

st.title("🦊 Foxsay")
st.caption("Asisten Gen Z siap bantu lo!")

# ===== GAYA GEN Z =====
GAYA_GEN_Z = """
Kamu adalah Foxsay, asisten AI bergaya Gen Z yang santai dan gaul.
Gunakan ekspresi-ekspresi ini:
- Bhap!, Vhom!, Buset!, Goks! (kagum/kaget)
- Mungkeen (kaget/kagum)
- Santay (santai)
- Gaskeunclek, Gas (semangat/ayo)
- Mager (malas)
- Wkwk (ketawa)
- BTW, FYI (informasi tambahan)
- Ges (panggilan akrab)

TAPI tetap informatif dan sesuai fakta. Ekspresi hanya sebagai bumbu.
"""


genai.configure(api_key="ISI_API_KEY_GEMINI_")

# ===== MEMORY PERCAKAPAN =====
if "messages" not in st.session_state:
    st.session_state.messages = []

# Tampilkan history chat
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.write(msg["content"])

# ===== FUNGSI PANGGIL AI =====
def panggil_AI(prompt):
    model = genai.GenerativeModel('gemini-1.5-flash')
    response = model.generate_content(GAYA_GEN_Z + "\n\n" + prompt)
    return response.text

# ===== INPUT CHAT =====
if prompt := st.chat_input("Ada yang bisa dibantu, ges?"):
    # Simpan pertanyaan user
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.write(prompt)
    
    # Ambil history (5 percakapan terakhir)
    history = st.session_state.messages[-6:-1]
    konteks = "\n".join([f"{m['role']}: {m['content']}" for m in history])
    
    # Panggil AI
    with st.chat_message("assistant"):
        with st.spinner("Santuy ges, lagi mikir..."):
            jawaban = panggil_AI(f"Percakapan sebelumnya:\n{konteks}\n\nPertanyaan baru: {prompt}")
            st.write(jawaban)
    
    # Simpan jawaban AI
    st.session_state.messages.append({"role": "assistant", "content": jawaban})
