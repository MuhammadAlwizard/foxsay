import streamlit as st

# ===== KONFIGURASI AWAL =====
st.set_page_config(page_title="Foxsay", layout="wide")

st.title("🦊 Foxsay")
st.caption("Asisten Gen Z siap bantu lo!")

# ===== MEMORY PERCAKAPAN =====
if "messages" not in st.session_state:
    st.session_state.messages = []

# Tampilkan history chat
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.write(msg["content"])

# Input chat
if prompt := st.chat_input("Ada yang bisa dibantu, ges?"):
    # Simpan pertanyaan
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.write(prompt)
    
    # Jawaban sementara (pake teks statis dulu)
    with st.chat_message("assistant"):
        jawaban = f"Wih! Pertanyaan lo: '{prompt}' ya? Sabar dulu ges, masih diproses nih, bentar lagi gaskeun! 🚀"
        st.write(jawaban)
    
    # Simpan jawaban
    st.session_state.messages.append({"role": "assistant", "content": jawaban})
