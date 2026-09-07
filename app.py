import streamlit as st
from groq import Groq
from duckduckgo_search import DDGS
import pdfplumber
import docx
import pandas as pd
import speech_recognition as sr
from pydub import AudioSegment
import io

# ===== KONFIGURASI AWAL =====
st.set_page_config(page_title="Foxsay", page_icon="🦊", layout="wide")

st.title("🦊 Foxsay")
st.caption("Asisten Gen Z siap bantu lo!")

# ===== GAYA GEN Z =====
GAYA_GEN_Z = """
Kamu adalah Foxsay, asisten AI bergaya Gen Z yang santai dan gaul.
Gunakan ekspresi-ekspresi ini dalam setiap jawaban:
- Vhom, Bhap, OMG, Wadaw, Lah, Hah, Buset, Anjir, Wadidaw, Yailah, Walah (ekspresi kaget/heran)
- Wih, Beh, Wuih, Goks, Gokil (ekspresi kagum)
- Nah, Alias, Jujur (penegasan)
- Wkwk, Hehehe, Uhuy (ketawa/goda)
- Gas, Gaskeun, Uhuy (semangat/ayo)
- Santuy, Woles (santai)
- Mager, Males (malas)
- Ges, Woi, Heh (panggilan)
- Dahlah, Hadeh, Aduh, Duh (kesal/prihatin)
- Oalah, Ohhh, Hooo (paham)
- Ckckck, Ngehe, Bacot (heran/kesal)

TAPI tetap informatif dan sesuai fakta. Ekspresi hanya sebagai bumbu.
"""

# ===== API KEY (ambil dari Streamlit Secrets) =====
GROQ_API_KEY = st.secrets["GROQ_API_KEY"]
client = Groq(api_key=GROQ_API_KEY)

# ===== MEMORY PERCAKAPAN =====
if "messages" not in st.session_state:
    st.session_state.messages = []

# ===== FUNGSI PANGGIL GROQ =====
def panggil_AI(prompt):
    completion = client.chat.completions.create(
        model="llama3-70b-8192",
        messages=[
            {"role": "system", "content": GAYA_GEN_Z},
            {"role": "user", "content": prompt}
        ],
        temperature=0.7,
        max_tokens=1024
    )
    return completion.choices[0].message.content

# ===== FUNGSI SEARCH INTERNET =====
def cari_internet(query):
    try:
        with DDGS() as ddgs:
            results = list(ddgs.text(query, max_results=3))
            if not results:
                return "Nggak ada hasil, ges."
            return "\n\n".join([f"{r['title']}: {r['body']}" for r in results])
    except Exception as e:
        return f"Error: {e}"

# ===== SIDEBAR MENU =====
menu = st.sidebar.radio(
    "Pilih Fitur:",
    ["💬 Chat", "📄 PDF/Word", "📊 Excel", "🎤 Voice Note"]
)

# ============================================================
# ===== FITUR 1: CHAT + MEMORY + SEARCH =====
# ============================================================
if menu == "💬 Chat":
    # Tampilkan history chat
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.write(msg["content"])
    
    # Input chat
    if prompt := st.chat_input("Ada yang bisa dibantu, ges?"):
        # Simpan pertanyaan user
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.write(prompt)
        
        # Cek apakah perlu cari di internet
        if "cari" in prompt.lower() or "internet" in prompt.lower() or "info" in prompt.lower():
            with st.chat_message("assistant"):
                with st.spinner("Santuy ges, lagi cari di internet..."):
                    hasil_internet = cari_internet(prompt)
                    jawaban = panggil_AI(f"Pertanyaan: {prompt}\n\nHasil pencarian internet:\n{hasil_internet}")
                    st.write(jawaban)
        else:
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

# ============================================================
# ===== FITUR 2: PDF / WORD =====
# ============================================================
elif menu == "📄 PDF/Word":
    st.subheader("📄 Upload PDF atau Word")
    st.caption("Upload file PDF atau Word, nanti bisa dirangkum!")
    
    file = st.file_uploader("Upload filenya, ges!", type=["pdf", "docx"])
    
    if file is not None:
        teks = ""
        
        # Baca PDF
        if file.type == "application/pdf":
            with pdfplumber.open(file) as pdf:
                for page in pdf.pages:
                    teks += page.extract_text() or ""
        
        # Baca Word
        elif file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            doc = docx.Document(file)
            for para in doc.paragraphs:
                teks += para.text + "\n"
        
        # Tampilkan teks
        with st.expander("📝 Lihat isi file", expanded=True):
            st.text_area("Isi file:", teks, height=300)
        
        # Tombol rangkum
        if st.button("📝 Rangkumin Dong, Ges!"):
            if len(teks) > 5000:
                teks = teks[:5000] + "\n... (dipotong karena terlalu panjang)"
            
            with st.spinner("Wadaw! Lagi baca file..."):
                rangkuman = panggil_AI(f"Rangkum isi file ini dengan gaya ekspresif:\n\n{teks}")
                st.success("### Rangkuman:")
                st.write(rangkuman)

# ============================================================
# ===== FITUR 3: EXCEL =====
# ============================================================
elif menu == "📊 Excel":
    st.subheader("📊 Upload Excel")
    st.caption("Upload file Excel, nanti bisa dianalisis!")
    
    file = st.file_uploader("Upload Excelnya, ges!", type=["xlsx", "xls"])
    
    if file is not None:
        df = pd.read_excel(file)
        
        # Tampilkan tabel
        st.dataframe(df)
        st.caption(f"Jumlah baris: {len(df)} | Jumlah kolom: {len(df.columns)}")
        
        # Tombol analisis
        if st.button("🔍 Analisis Yuk, Ges!"):
            with st.spinner("Beh! Lagi ngitung..."):
                info = f"Kolom: {', '.join(df.columns)}\n\n5 data pertama:\n{df.head().to_string()}"
                analisis = panggil_AI(f"Analisis data Excel ini secara singkat (insight + rekomendasi):\n\n{info}")
                st.success("### Analisis:")
                st.write(analisis)

# ============================================================
# ===== FITUR 4: VOICE NOTE =====
# ============================================================
elif menu == "🎤 Voice Note":
    st.subheader("🎤 Rekam Suara")
    st.caption("Pencet tombol, bicara, stop. Gampang kan?")
    
    audio = st.audio_input("🎙️ Pencet terus sambil bicara, ya!")
    
    if audio is not None:
        st.audio(audio)
        
        with st.spinner("Hmmm... lagi dengerin suara lo..."):
            try:
                audio_bytes = audio.getvalue()
                audio_segment = AudioSegment.from_file(io.BytesIO(audio_bytes))
                
                with open("temp_audio.wav", "wb") as f:
                    audio_segment.export(f, format="wav")
                
                recognizer = sr.Recognizer()
                with sr.AudioFile("temp_audio.wav") as source:
                    audio_data = recognizer.record(source)
                    teks = recognizer.recognize_google(audio_data, language="id-ID")
                
                st.info(f"📝 Yang lo omongin: *{teks}*")
                
                with st.spinner("Vhom! Lagi nyusun jawaban..."):
                    jawaban = panggil_AI(f"User nanya: {teks}. Jawab dengan gaya Foxsay yang ekspresif!")
                    st.success("### Jawaban Foxsay:")
                    st.write(jawaban)
            
            except sr.UnknownValueError:
                st.error("Wadaw! Suaranya nggak jelas, coba ulangi ya!")
            except Exception as e:
                st.error(f"Bhap! Ada error: {e}")

# ============================================================
# ===== SIDEBAR: TOMBOL RESET =====
# ============================================================
if st.sidebar.button("🔄 Mulai Lagi, Ges!"):
    st.session_state.messages = []
    st.sidebar.success("History dibersihin! Gaskeun!")
