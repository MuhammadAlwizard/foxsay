import streamlit as st
from groq import Groq
from ddgs import DDGS
import pypdf
import docx
import pandas as pd
import plotly.express as px
from streamlit.errors import StreamlitSecretNotFoundError

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

try:
    api_key = st.secrets["GROQ_API_KEY"]
except (KeyError, StreamlitSecretNotFoundError):
    api_key = ""

client = Groq(api_key=api_key) if api_key else None
if client is None:
    st.warning("GROQ_API_KEY belum dikonfigurasi. Chat AI dan voice input akan aktif setelah secret ditambahkan.")

if "history" not in st.session_state:
    st.session_state.history = []
if "isi_file" not in st.session_state:
    st.session_state.isi_file = ""
if "nama_file" not in st.session_state:
    st.session_state.nama_file = ""
if "show_uploader" not in st.session_state:
    st.session_state.show_uploader = False
if "df_aktif" not in st.session_state:
    st.session_state.df_aktif = None
if "show_mic" not in st.session_state:
    st.session_state.show_mic = False
if "teks_transkrip" not in st.session_state:
    st.session_state.teks_transkrip = ""

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

ATURAN KHUSUS UNTUK KATA "ALIAS": jangan dipakai di setiap jawaban — ini fitur LANGKA yang
cuma muncul sesekali (kira-kira 1 dari 4-5 jawaban), HANYA kalau ada momen yang bener-bener
related dan lucu. Kalau nggak nemu momen yang pas, JANGAN pakai kata ALIAS sama sekali,
jawab normal aja tanpa maksain.

Kalau kamu pakai ALIAS, WAJIB diikuti punchline/kalimat lengkap setelahnya — jangan pernah
menaruh "ALIAS" sebagai kata penutup yang menggantung tanpa isi. FORMAT SALAH (dilarang):
"...jawabannya keren! ALIAS" (menggantung, tidak ada isi setelahnya).
FORMAT BENAR: "Bandung kota metropolitan terbesar ketiga, ALIAS kalau macet nomor 1
se-Jabar." — ALIAS diikuti kalimat utuh yang jadi punchline-nya, bukan berdiri sendiri di
akhir. Selalu tulis kapital: ALIAS.

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

def baca_excel(file):
    df = pd.read_excel(file)
    return df

def baca_csv(file):
    df = pd.read_csv(file)
    return df

def cari_kolom_tanggal(df):
    """Cari kolom tanggal secara aman tanpa mengubah dataframe asli."""
    kandidat = []
    for kolom in df.columns:
        nama = str(kolom).lower()
        seri = df[kolom]
        if pd.api.types.is_datetime64_any_dtype(seri):
            kandidat.append((kolom, seri))
            continue
        if not (pd.api.types.is_object_dtype(seri) or pd.api.types.is_string_dtype(seri)):
            continue
        try:
            hasil = pd.to_datetime(seri, errors="coerce")
            jumlah_valid = hasil.notna().sum()
            kata_tanggal = any(kata in nama for kata in ("date", "tanggal", "time", "waktu", "month", "year"))
            if jumlah_valid >= 2 and (kata_tanggal or jumlah_valid / max(len(seri), 1) >= 0.8):
                kandidat.append((kolom, hasil))
        except (ValueError, TypeError):
            continue
    return kandidat[0] if kandidat else (None, None)

def buat_insight_data(df, kolom_tanggal, kolom_numerik, kolom_kategori):
    """Buat insight singkat berbasis statistik dasar, tanpa memanggil AI."""
    insight = [f"Data berisi {len(df):,} baris dan {len(df.columns)} kolom.".replace(",", ".")]
    if kolom_numerik:
        kolom = kolom_numerik[0]
        seri = pd.to_numeric(df[kolom], errors="coerce").dropna()
        if not seri.empty:
            insight.append(f"Nilai rata-rata **{kolom}** adalah {seri.mean():,.2f}, dengan rentang {seri.min():,.2f} sampai {seri.max():,.2f}.")
    if kolom_kategori and kolom_numerik:
        kategori, nilai = kolom_kategori[0], kolom_numerik[0]
        ringkasan = df.groupby(kategori, dropna=False)[nilai].sum().sort_values(ascending=False)
        if not ringkasan.empty:
            insight.append(f"Kategori tertinggi berdasarkan total **{nilai}** adalah **{ringkasan.index[0]}**.")
    if kolom_tanggal:
        insight.append(f"Kolom **{kolom_tanggal}** dikenali sebagai tanggal sehingga grafik tren dapat dibuat.")
    return insight

def tampilkan_analisis_grafik(df):
    """Tampilkan tipe data, grafik otomatis, dan insight untuk CSV/XLSX."""
    if df is None or df.empty:
        st.warning("File data kosong, jadi belum ada grafik yang bisa dibuat.")
        return

    st.subheader("Analisis grafik otomatis")
    tipe_data = pd.DataFrame({"Kolom": df.columns, "Tipe data": [str(tipe) for tipe in df.dtypes]})
    st.caption("Kolom dan tipe data yang terbaca")
    st.dataframe(tipe_data, hide_index=True, use_container_width=True)

    kolom_numerik = list(df.select_dtypes(include="number").columns)
    kolom_kategori = [
        kolom for kolom in df.select_dtypes(include=["object", "category", "bool"]).columns
        if 1 < df[kolom].nunique(dropna=False) <= 30
    ]
    kolom_tanggal, tanggal = cari_kolom_tanggal(df)
    insight = buat_insight_data(df, kolom_tanggal, kolom_numerik, kolom_kategori)

    st.markdown("**Insight singkat**")
    for item in insight:
        st.markdown(f"- {item}")

    if kolom_tanggal and kolom_numerik:
        try:
            tren = df[[kolom_tanggal, kolom_numerik[0]]].copy()
            tren["__tanggal"] = tanggal
            tren["__nilai"] = pd.to_numeric(tren[kolom_numerik[0]], errors="coerce")
            tren = tren.dropna(subset=["__tanggal", "__nilai"]).sort_values("__tanggal")
            if not tren.empty:
                fig = px.line(tren, x="__tanggal", y="__nilai", markers=True, title=f"Tren {kolom_numerik[0]} dari waktu ke waktu")
                fig.update_layout(xaxis_title=kolom_tanggal, yaxis_title=kolom_numerik[0])
                st.plotly_chart(fig, use_container_width=True)
        except (ValueError, TypeError, KeyError) as error:
            st.warning(f"Grafik tren tidak dapat dibuat: {error}")
    elif kolom_tanggal:
        st.info("Kolom tanggal ditemukan, tetapi belum ada kolom angka untuk membuat grafik tren.")

    if kolom_kategori and kolom_numerik:
        try:
            kategori, nilai = kolom_kategori[0], kolom_numerik[0]
            perbandingan = df.groupby(kategori, dropna=False)[nilai].sum().reset_index().sort_values(nilai, ascending=False).head(20)
            fig = px.bar(perbandingan, x=kategori, y=nilai, title=f"Perbandingan {nilai} berdasarkan {kategori}")
            st.plotly_chart(fig, use_container_width=True)
        except (ValueError, TypeError, KeyError) as error:
            st.warning(f"Grafik perbandingan tidak dapat dibuat: {error}")
    elif kolom_kategori:
        st.info("Kolom kategori ditemukan, tetapi belum ada kolom angka untuk dibandingkan.")

    if kolom_numerik:
        try:
            nilai = kolom_numerik[0]
            fig = px.histogram(df, x=nilai, nbins=30, title=f"Distribusi {nilai}")
            st.plotly_chart(fig, use_container_width=True)
        except (ValueError, TypeError, KeyError) as error:
            st.warning(f"Grafik distribusi tidak dapat dibuat: {error}")

    if len(kolom_numerik) >= 2:
        try:
            korelasi = df[kolom_numerik].corr(numeric_only=True)
            fig = px.imshow(korelasi, text_auto=True, color_continuous_scale="RdBu_r", zmin=-1, zmax=1, title="Korelasi antar kolom angka")
            st.plotly_chart(fig, use_container_width=True)
        except (ValueError, TypeError) as error:
            st.warning(f"Grafik korelasi tidak dapat dibuat: {error}")
    else:
        st.info("Grafik korelasi membutuhkan minimal dua kolom angka.")

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
    if client is None:
        return "GROQ_API_KEY belum dikonfigurasi. Tambahkan secret tersebut untuk menggunakan chat AI Foxsay."

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
    uploaded_file = st.file_uploader("Upload PDF, Word, Excel, atau CSV", type=["pdf", "docx", "xlsx", "csv"], label_visibility="collapsed")
    if uploaded_file is not None and uploaded_file.name != st.session_state.nama_file:
        with st.spinner("🦊 Foxsay lagi baca file..."):
            try:
                st.session_state.df_aktif = None
                if uploaded_file.name.endswith(".pdf"):
                    isi = baca_pdf(uploaded_file)
                elif uploaded_file.name.endswith(".docx"):
                    isi = baca_docx(uploaded_file)
                elif uploaded_file.name.endswith(".xlsx"):
                    df = baca_excel(uploaded_file)
                    st.session_state.df_aktif = df
                    isi = df.to_string()
                elif uploaded_file.name.endswith(".csv"):
                    df = baca_csv(uploaded_file)
                    st.session_state.df_aktif = df
                    isi = df.to_string()
                else:
                    isi = ""
                st.session_state.isi_file = isi
                st.session_state.nama_file = uploaded_file.name
                st.session_state.show_uploader = False
                st.success(f"File '{uploaded_file.name}' berhasil dibaca!")
                st.rerun()
            except Exception as e:
                st.error(f"Gagal baca file: {e}")

if st.session_state.df_aktif is not None:
    st.dataframe(st.session_state.df_aktif, use_container_width=True)
    tampilkan_analisis_grafik(st.session_state.df_aktif)

if st.session_state.nama_file:
    label_tombol = "Analisis data ini" if st.session_state.df_aktif is not None else "Rangkum file ini"
    pesan_default = "Tolong analisis data ini: kasih insight menarik, pattern, atau hal penting yang perlu diperhatikan." if st.session_state.df_aktif is not None else "Tolong rangkum isi file ini secara singkat dan jelas."
    col_a, col_b = st.columns(2)
    with col_a:
        if st.button(label_tombol):
            with st.spinner("🦊 Foxsay lagi mikir..."):
                jawaban = agent(pesan_default, st.session_state.isi_file)
            st.session_state.history.append(("user", f"[Minta rangkuman: {st.session_state.nama_file}]"))
            st.session_state.history.append(("ai", jawaban))
            st.rerun()
    with col_b:
        if st.button("Hapus file"):
            st.session_state.isi_file = ""
            st.session_state.nama_file = ""
            st.session_state.df_aktif = None
            st.rerun()

st.write("Tanya apa aja ke Foxsay:")
kotak_utama = st.container(border=True)
with kotak_utama:
    with st.form("tanya_form", clear_on_submit=True):
        col_plus, col_mic, col_input, col_submit = st.columns([1, 1, 4, 1.3])
        with col_plus:
            attach_clicked = st.form_submit_button("➕")
        with col_mic:
            mic_clicked = st.form_submit_button("🎤")
        with col_input:
            pertanyaan = st.text_input("Tanya apa aja ke Foxsay:", label_visibility="collapsed")
        with col_submit:
            submitted = st.form_submit_button("Tanya")

    if st.session_state.show_mic:
        audio_value = st.audio_input("Rekam pesan suara kamu", label_visibility="collapsed")
        if audio_value is not None and client is not None:
            with st.spinner("🦊 Foxsay lagi dengerin..."):
                try:
                    transkrip = client.audio.transcriptions.create(
                        file=("voice.wav", audio_value.read()),
                        model="whisper-large-v3",
                        language="id"
                    )
                    st.session_state.teks_transkrip = transkrip.text
                except Exception as e:
                    st.error(f"Gagal transkrip suara: {e}")

        if st.session_state.teks_transkrip:
            st.info(f"📝 Hasil transkrip: \"{st.session_state.teks_transkrip}\"")
            col_x, col_y = st.columns(2)
            with col_x:
                if st.button("Kirim ke Foxsay"):
                    pertanyaan_vn = st.session_state.teks_transkrip
                    st.session_state.teks_transkrip = ""
                    st.session_state.show_mic = False
                    with st.spinner("🦊 Foxsay lagi mikir..."):
                        jawaban = agent(pertanyaan_vn, st.session_state.isi_file)
                    st.session_state.history.append(("user", f"🎤 {pertanyaan_vn}"))
                    st.session_state.history.append(("ai", jawaban))
                    st.rerun()
            with col_y:
                if st.button("Batal"):
                    st.session_state.teks_transkrip = ""
                    st.session_state.show_mic = False
                    st.rerun()

if attach_clicked:
    st.session_state.show_uploader = not st.session_state.show_uploader
    st.rerun()

if mic_clicked:
    st.session_state.show_mic = not st.session_state.show_mic
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
