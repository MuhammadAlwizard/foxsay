import streamlit as st
from groq import Groq
from ddgs import DDGS
import pypdf
import docx
from docx.document import Document as DocxDocument
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
import pandas as pd
import plotly.express as px
from html import escape
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
        white-space: pre-wrap;
        overflow-wrap: anywhere;
    }
    .chat-bubble-ai {
        background-color: #FF6B35;
        padding: 12px 16px;
        border-radius: 15px;
        margin: 8px 0;
        color: white;
        white-space: pre-wrap;
        overflow-wrap: anywhere;
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
if "mode_aktif" not in st.session_state:
    st.session_state.mode_aktif = "Chill"
if "excel_sheet_info" not in st.session_state:
    st.session_state.excel_sheet_info = ""
if "file_read_warning" not in st.session_state:
    st.session_state.file_read_warning = ""

MAX_HISTORY_STORED = 100
MAX_HISTORY_RENDER = 30

def tambah_riwayat(role, teks):
    """Simpan riwayat dengan batas agar session tidak tumbuh tanpa batas."""
    st.session_state.history.append((role, str(teks)))
    if len(st.session_state.history) > MAX_HISTORY_STORED:
        del st.session_state.history[:-MAX_HISTORY_STORED]

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

MODE_INSTRUCTIONS = {
    "Chill": "Jawab pertanyaan umum dengan gaya Foxsay yang santai, ekspresif, dan tetap informatif.",
    "Study": "Bantu user belajar. Jelaskan konsep secara bertahap, gunakan contoh sederhana, dan jika diminta buat rangkuman, flashcard, atau kuis. Jangan langsung memberi jawaban tugas tanpa penjelasan.",
    "Creator": "Bantu user membuat konten dengan cepat. Untuk ide atau script TikTok/Reels/Shorts, gunakan default yang masuk akal jika detail tidak disebutkan: TikTok, 30 detik, dan gaya santai. Selalu susun output dengan bagian: konsep, 3 hook, script per scene beserta estimasi waktu, arahan visual, voice-over, teks layar, caption, CTA, dan hashtag. Sesuaikan dengan detail yang diberikan user. Jangan mengarang fakta atau mengklaim tren terbaru tanpa sumber.",
    "Career": "Bantu user mempersiapkan karier. Fokus pada CV, portfolio, interview, personal branding, dan strategi pencarian kerja. Berikan saran yang konkret dan bisa langsung dipakai.",
}

def baca_pdf(file):
    file.seek(0)
    reader = pypdf.PdfReader(file)
    teks = ""
    for page in reader.pages:
        teks_halaman = page.extract_text() or ""
        if teks_halaman.strip():
            teks += teks_halaman + "\n"
    return teks

def iterasi_blok_docx(parent):
    """Iterasi paragraf dan tabel DOCX sesuai urutan kemunculannya."""
    if isinstance(parent, DocxDocument):
        parent_element = parent.element.body
    elif isinstance(parent, _Cell):
        parent_element = parent._tc
    else:
        raise TypeError("Parent DOCX tidak didukung")

    for child in parent_element.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def baca_docx(file):
    file.seek(0)
    doc = docx.Document(file)
    bagian = []
    for blok in iterasi_blok_docx(doc):
        if isinstance(blok, Paragraph):
            teks = blok.text.strip()
            if teks:
                bagian.append(teks)
        elif isinstance(blok, Table):
            baris = []
            for row in blok.rows:
                sel = [" ".join(cell.text.split()) for cell in row.cells]
                if any(sel):
                    baris.append(" | ".join(sel))
            if baris:
                bagian.append("[TABEL]\n" + "\n".join(baris) + "\n[/TABEL]")
    return "\n".join(bagian)

def daftar_sheet_excel(file):
    """Ambil nama sheet tanpa mengubah posisi baca file secara permanen."""
    file.seek(0)
    with pd.ExcelFile(file) as workbook:
        return workbook.sheet_names

def baca_excel(file, sheet_name=0):
    file.seek(0)
    df = pd.read_excel(file, sheet_name=sheet_name)
    return df

def baca_csv(file):
    file.seek(0)
    df = pd.read_csv(file)
    return df

def buat_payload_audio(audio_value):
    """Kirim bytes audio dengan nama dan MIME yang konsisten untuk Groq."""
    mime_to_extension = {
        "audio/flac": ".flac",
        "audio/mpeg": ".mp3",
        "audio/mp4": ".mp4",
        "audio/x-m4a": ".m4a",
        "audio/m4a": ".m4a",
        "audio/ogg": ".ogg",
        "audio/wav": ".wav",
        "audio/webm": ".webm",
    }
    extension_to_mime = {value: key for key, value in mime_to_extension.items()}
    mime = str(getattr(audio_value, "type", "") or "").lower()
    extension = mime_to_extension.get(mime)

    if extension is None:
        nama = str(getattr(audio_value, "name", "") or "").lower()
        suffix = "." + nama.rsplit(".", 1)[-1] if "." in nama else ""
        extension = suffix if suffix in extension_to_mime else ".wav"
        mime = extension_to_mime[extension]

    # Groq/OpenAI-compatible SDK mendokumentasikan tuple (filename, bytes).
    # Ekstensi pada filename tetap mencerminkan MIME agar format dikenali API.
    return (f"voice{extension}", audio_value.getvalue())

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

def format_statistik(nilai):
    """Format angka untuk konteks AI tanpa membuat nilai terlihat ambigu."""
    if pd.isna(nilai):
        return "NA"
    return f"{float(nilai):.4f}"

def buat_ringkasan_statistik(df):
    """Hitung ringkasan statistik yang menjadi sumber fakta untuk jawaban AI."""
    if df is None or df.empty:
        return "DATA KOSONG: tidak ada statistik yang dapat dihitung."

    bagian = [
        "STATISTIK TERVERIFIKASI DARI PANDAS",
        "Gunakan angka di bawah ini sebagai satu-satunya sumber statistik. Jangan menebak atau mengubah angka.",
        f"Jumlah baris: {len(df)}",
        f"Jumlah kolom: {len(df.columns)}",
    ]

    missing = df.isna().sum()
    missing = missing[missing > 0]
    if missing.empty:
        bagian.append("Data hilang: tidak ditemukan.")
    else:
        bagian.append("Data hilang per kolom:")
        for kolom, jumlah in missing.items():
            bagian.append(f"- {kolom}: {int(jumlah)} nilai hilang")

    numerik = df.select_dtypes(include="number")
    if numerik.empty:
        bagian.append("Kolom numerik: tidak ditemukan.")
    else:
        bagian.append("Statistik kolom numerik (count, mean, std, min, median, max):")
        deskripsi = numerik.describe().T
        for kolom, baris in deskripsi.iterrows():
            bagian.append(
                f"- {kolom}: count={format_statistik(baris['count'])}, "
                f"mean={format_statistik(baris['mean'])}, "
                f"std={format_statistik(baris['std'])}, "
                f"min={format_statistik(baris['min'])}, "
                f"median={format_statistik(baris['50%'])}, "
                f"max={format_statistik(baris['max'])}"
            )

        if len(numerik.columns) >= 2:
            korelasi = numerik.corr(numeric_only=True)
            pasangan = []
            for indeks, kolom_a in enumerate(korelasi.columns):
                for kolom_b in korelasi.columns[indeks + 1:]:
                    nilai = korelasi.loc[kolom_a, kolom_b]
                    if pd.notna(nilai):
                        pasangan.append((abs(nilai), kolom_a, kolom_b, nilai))
            pasangan.sort(reverse=True)
            if pasangan:
                bagian.append("Korelasi Pearson terkuat (hubungan, bukan sebab-akibat):")
                for _, kolom_a, kolom_b, nilai in pasangan[:10]:
                    bagian.append(f"- {kolom_a} dengan {kolom_b}: {format_statistik(nilai)}")

    kategori = df.select_dtypes(include=["object", "category", "bool"])
    if not kategori.empty:
        bagian.append("Kategori teratas per kolom kategorikal:")
        for kolom in kategori.columns:
            jumlah_unik = kategori[kolom].nunique(dropna=False)
            if 1 < jumlah_unik <= 30:
                frekuensi = kategori[kolom].value_counts(dropna=False).head(5)
                daftar = ", ".join(f"{repr(indeks)} ({jumlah})" for indeks, jumlah in frekuensi.items())
                bagian.append(f"- {kolom}: {daftar}")

    return "\n".join(bagian)

def buat_analisis_terstruktur(df):
    """Buat analisis bisnis terukur agar AI membedakan fakta dan dugaan."""
    hasil = {
        "Fakta dari data": [],
        "Interpretasi": [],
        "Hipotesis": [],
        "Rekomendasi": [],
    }
    if df is None or df.empty:
        hasil["Fakta dari data"].append("Data kosong; analisis lanjutan belum dapat dilakukan.")
        return hasil

    numerik = df.select_dtypes(include="number")
    kategori = df.select_dtypes(include=["object", "category", "bool"])

    missing = df.isna().sum()
    missing = missing[missing > 0]
    if missing.empty:
        hasil["Fakta dari data"].append("Tidak ditemukan nilai yang hilang.")
    else:
        for kolom, jumlah in missing.items():
            persentase = jumlah / len(df) * 100
            hasil["Fakta dari data"].append(
                f"Kolom {kolom} memiliki {int(jumlah)} nilai hilang ({persentase:.2f}% dari baris)."
            )
        hasil["Rekomendasi"].append(
            "Tentukan strategi untuk nilai hilang sebelum membuat model prediksi; jangan langsung menganggapnya acak."
        )

    # Nilai negatif hanya ditandai sebagai anomali. Maknanya tetap harus dikonfirmasi
    # berdasarkan definisi kolom, karena beberapa metrik seperti profit memang boleh negatif.
    anomali_negatif = []
    for kolom in numerik.columns:
        seri = pd.to_numeric(df[kolom], errors="coerce").dropna()
        negatif = seri[seri < 0]
        if not negatif.empty:
            anomali_negatif.append((kolom, len(negatif), negatif.min()))
    for kolom, jumlah, minimum in anomali_negatif:
        hasil["Fakta dari data"].append(
            f"Kolom {kolom} memiliki {jumlah} nilai negatif; nilai minimum {minimum:.4f}."
        )
        hasil["Hipotesis"].append(
            f"Nilai negatif pada {kolom} mungkin merupakan error input, tetapi perlu dikonfirmasi dari definisi data."
        )
        hasil["Rekomendasi"].append(
            f"Validasi nilai negatif pada {kolom} sebelum analisis lanjutan; jangan menghapusnya tanpa aturan bisnis."
        )

    # IQR digunakan untuk menemukan kandidat outlier, bukan untuk menyatakan bahwa
    # nilai tersebut pasti salah.
    kandidat_outlier = []
    for kolom in numerik.columns:
        seri = pd.to_numeric(df[kolom], errors="coerce").dropna()
        if len(seri) < 4:
            continue
        kuartil_1 = seri.quantile(0.25)
        kuartil_3 = seri.quantile(0.75)
        rentang_iqr = kuartil_3 - kuartil_1
        if rentang_iqr == 0:
            continue
        batas_bawah = kuartil_1 - 1.5 * rentang_iqr
        batas_atas = kuartil_3 + 1.5 * rentang_iqr
        jumlah = int(((seri < batas_bawah) | (seri > batas_atas)).sum())
        if jumlah:
            kandidat_outlier.append((jumlah, kolom, batas_bawah, batas_atas))
    for jumlah, kolom, batas_bawah, batas_atas in sorted(kandidat_outlier, reverse=True)[:5]:
        hasil["Fakta dari data"].append(
            f"Kolom {kolom} memiliki {jumlah} kandidat outlier menurut aturan IQR "
            f"(batas {batas_bawah:.4f} sampai {batas_atas:.4f})."
        )
        hasil["Interpretasi"].append(
            f"Kandidat outlier pada {kolom} perlu diperiksa; outlier belum tentu merupakan kesalahan data."
        )
        hasil["Rekomendasi"].append(
            f"Periksa baris sumber pada {kolom} dan bandingkan dengan konteks bisnis sebelum melakukan imputasi atau penghapusan."
        )

    # Cari perbedaan rata-rata antar kelompok kategorikal dengan jumlah sampel yang
    # cukup agar rekomendasi tidak hanya berdasarkan pengamatan visual.
    perbandingan = []
    for kolom_kategori in kategori.columns:
        jumlah_unik = kategori[kolom_kategori].nunique(dropna=False)
        if not 2 <= jumlah_unik <= 10:
            continue
        for kolom_nilai in numerik.columns:
            pasangan = df[[kolom_kategori, kolom_nilai]].copy()
            pasangan[kolom_nilai] = pd.to_numeric(pasangan[kolom_nilai], errors="coerce")
            grup = pasangan.groupby(kolom_kategori, dropna=False)[kolom_nilai].agg(["mean", "count"]).dropna(subset=["mean"])
            grup = grup[grup["count"] >= 2]
            if len(grup) < 2:
                continue
            tertinggi = grup["mean"].idxmax()
            terendah = grup["mean"].idxmin()
            nilai_tinggi = grup.loc[tertinggi, "mean"]
            nilai_terendah = grup.loc[terendah, "mean"]
            selisih = nilai_tinggi - nilai_terendah
            pembagi = max(abs(grup["mean"].mean()), 1e-9)
            perbandingan.append((abs(selisih) / pembagi, kolom_kategori, kolom_nilai, tertinggi, terendah, nilai_tinggi, nilai_terendah))

    for _, kolom_kategori, kolom_nilai, tertinggi, terendah, nilai_tinggi, nilai_terendah in sorted(perbandingan, reverse=True)[:5]:
        nama_tinggi = "(kosong)" if pd.isna(tertinggi) else str(tertinggi)
        nama_terendah = "(kosong)" if pd.isna(terendah) else str(terendah)
        hasil["Fakta dari data"].append(
            f"Rata-rata {kolom_nilai} tertinggi ada pada {kolom_kategori}={nama_tinggi} "
            f"({nilai_tinggi:.4f}); terendah pada {kolom_kategori}={nama_terendah} ({nilai_terendah:.4f})."
        )
        hasil["Interpretasi"].append(
            f"Perbedaan rata-rata {kolom_nilai} antar kelompok {kolom_kategori} terlihat pada data ini, "
            "tetapi belum membuktikan penyebabnya."
        )
        hasil["Hipotesis"].append(
            f"Perbedaan {kolom_nilai} mungkin berkaitan dengan karakteristik kelompok {kolom_kategori}; "
            "perlu diuji dengan variabel tambahan."
        )
        hasil["Rekomendasi"].append(
            f"Bandingkan ukuran sampel dan karakteristik tiap kelompok {kolom_kategori} sebelum membuat strategi khusus."
        )

    if not hasil["Interpretasi"] and numerik.empty:
        hasil["Interpretasi"].append("Belum ada kolom numerik untuk membuat perbandingan statistik.")
    return hasil

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
    missing = df.isna().sum()
    missing = missing[missing > 0]
    if not missing.empty:
        detail_missing = ", ".join(f"{kolom} ({int(jumlah)})" for kolom, jumlah in missing.items())
        insight.append(f"Ditemukan data hilang pada: **{detail_missing}**.")
    if len(kolom_numerik) >= 2:
        korelasi = df[kolom_numerik].corr(numeric_only=True)
        pasangan = []
        for indeks, kolom_a in enumerate(korelasi.columns):
            for kolom_b in korelasi.columns[indeks + 1:]:
                nilai = korelasi.loc[kolom_a, kolom_b]
                if pd.notna(nilai):
                    pasangan.append((abs(nilai), kolom_a, kolom_b, nilai))
        if pasangan:
            _, kolom_a, kolom_b, nilai = max(pasangan)
            insight.append(
                f"Hubungan numerik terkuat adalah **{kolom_a}** dan **{kolom_b}** "
                f"dengan korelasi {nilai:.2f}; ini bukan bukti sebab-akibat."
            )
    return insight

def tampilkan_grafik_interaktif(df):
    """Berikan kontrol manual untuk mengeksplorasi pasangan kolom yang dipilih user."""
    if df is None or df.empty or len(df.columns) == 0:
        return

    kolom_numerik = list(df.select_dtypes(include="number").columns)
    if not kolom_numerik:
        return

    with st.expander("Eksplorasi grafik manual"):
        st.caption("Pilih kolom sendiri untuk melihat pola yang belum tentu muncul pada grafik otomatis.")
        kolom_x, kolom_y, jenis_grafik = st.columns(3)
        with kolom_x:
            x = st.selectbox("Sumbu X", list(df.columns), key="grafik_manual_x")
        with kolom_y:
            pilihan_y = ["(tanpa sumbu Y)"] + kolom_numerik
            y = st.selectbox("Sumbu Y", pilihan_y, index=1, key="grafik_manual_y")
        with jenis_grafik:
            jenis = st.selectbox(
                "Jenis grafik",
                ["Bar", "Line", "Scatter", "Histogram"],
                key="grafik_manual_jenis",
            )

        data = df.copy()
        x_plot = x
        kolom_tanggal, tanggal = cari_kolom_tanggal(df)
        if x == kolom_tanggal:
            data["__x_grafik"] = tanggal
            x_plot = "__x_grafik"

        try:
            if jenis == "Histogram":
                data_histogram = data[[x_plot]].dropna()
                fig = px.histogram(data_histogram, x=x_plot, nbins=30, title=f"Distribusi {x}")
            elif y == "(tanpa sumbu Y)":
                st.info("Pilih kolom angka sebagai Sumbu Y untuk grafik ini.")
                return
            else:
                data[y] = pd.to_numeric(data[y], errors="coerce")
                data_plot = data[[x_plot, y]].dropna()
                if data_plot.empty:
                    st.warning("Tidak ada pasangan data valid untuk grafik yang dipilih.")
                    return

                if jenis == "Bar" and not pd.api.types.is_numeric_dtype(data_plot[x_plot]):
                    data_plot = (
                        data_plot.groupby(x_plot, dropna=False)[y]
                        .mean()
                        .reset_index()
                        .sort_values(y, ascending=False)
                        .head(30)
                    )

                if jenis == "Bar":
                    fig = px.bar(data_plot, x=x_plot, y=y, title=f"Rata-rata {y} berdasarkan {x}")
                elif jenis == "Line":
                    fig = px.line(data_plot.sort_values(x_plot), x=x_plot, y=y, markers=True, title=f"Tren {y} berdasarkan {x}")
                else:
                    fig = px.scatter(data_plot, x=x_plot, y=y, title=f"Hubungan {x} dan {y}")

            fig.update_layout(xaxis_title=x, yaxis_title=y if y != "(tanpa sumbu Y)" else "Jumlah")
            st.plotly_chart(fig, use_container_width=True)
        except (ValueError, TypeError, KeyError) as error:
            st.warning(f"Grafik manual tidak dapat dibuat: {error}")

def buat_grafik_laporan(df):
    """Buat salinan grafik otomatis untuk dimasukkan ke laporan HTML."""
    grafik = []
    if df is None or df.empty:
        return grafik

    kolom_numerik = list(df.select_dtypes(include="number").columns)
    kolom_kategori = [
        kolom for kolom in df.select_dtypes(include=["object", "category", "bool"]).columns
        if 1 < df[kolom].nunique(dropna=False) <= 30
    ]
    kolom_tanggal, tanggal = cari_kolom_tanggal(df)

    if kolom_tanggal and kolom_numerik:
        tren = df[[kolom_tanggal, kolom_numerik[0]]].copy()
        tren["__tanggal"] = tanggal
        tren["__nilai"] = pd.to_numeric(tren[kolom_numerik[0]], errors="coerce")
        tren = tren.dropna(subset=["__tanggal", "__nilai"]).sort_values("__tanggal")
        if not tren.empty:
            fig = px.line(
                tren,
                x="__tanggal",
                y="__nilai",
                markers=True,
                title=f"Tren {kolom_numerik[0]} dari waktu ke waktu",
            )
            fig.update_layout(xaxis_title=kolom_tanggal, yaxis_title=kolom_numerik[0])
            grafik.append(fig)

    if kolom_kategori and kolom_numerik:
        kategori, nilai = kolom_kategori[0], kolom_numerik[0]
        perbandingan = (
            df.groupby(kategori, dropna=False)[nilai]
            .sum()
            .reset_index()
            .sort_values(nilai, ascending=False)
            .head(20)
        )
        if not perbandingan.empty:
            fig = px.bar(
                perbandingan,
                x=kategori,
                y=nilai,
                title=f"Perbandingan {nilai} berdasarkan {kategori}",
            )
            grafik.append(fig)

    if kolom_numerik:
        nilai = kolom_numerik[0]
        fig = px.histogram(df, x=nilai, nbins=30, title=f"Distribusi {nilai}")
        grafik.append(fig)

    if len(kolom_numerik) >= 2:
        korelasi = df[kolom_numerik].corr(numeric_only=True)
        fig = px.imshow(
            korelasi,
            text_auto=True,
            color_continuous_scale="RdBu_r",
            zmin=-1,
            zmax=1,
            title="Korelasi antar kolom angka",
        )
        grafik.append(fig)

    return grafik

def buat_laporan_html(df, nama_file=""):
    """Buat laporan HTML mandiri yang bisa diunduh dan dibuka tanpa Streamlit."""
    statistik = df.select_dtypes(include="number").describe().T.reset_index()
    statistik = statistik.rename(columns={
        "index": "Kolom",
        "count": "Jumlah",
        "mean": "Rata-rata",
        "std": "Std",
        "min": "Minimum",
        "50%": "Median",
        "max": "Maksimum",
    })
    kolom_statistik = ["Kolom", "Jumlah", "Rata-rata", "Std", "Minimum", "Median", "Maksimum"]
    if not statistik.empty:
        tabel_statistik = statistik[kolom_statistik].round(4).to_html(index=False, classes="stats", border=0)
    else:
        tabel_statistik = "<p>Tidak ada kolom numerik.</p>"

    analisis = buat_analisis_terstruktur(df)
    bagian_analisis = []
    for judul, daftar in analisis.items():
        if daftar:
            isi = "".join(f"<li>{escape(str(item))}</li>" for item in daftar)
            bagian_analisis.append(f"<h3>{escape(judul)}</h3><ul>{isi}</ul>")

    grafik_html = []
    for indeks, fig in enumerate(buat_grafik_laporan(df)):
        grafik_html.append(
            fig.to_html(
                full_html=False,
                include_plotlyjs=True if indeks == 0 else False,
                config={"responsive": True, "displaylogo": False},
            )
        )

    judul_file = escape(nama_file or "data yang dianalisis")
    return f"""<!doctype html>
<html lang="id">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Laporan Analisis Foxsay</title>
<style>
body {{ font-family: Arial, sans-serif; max-width: 1100px; margin: 0 auto; padding: 32px; color: #252525; line-height: 1.55; }}
h1 {{ color: #e9572b; margin-bottom: 4px; }}
h2 {{ border-bottom: 2px solid #f0a07f; padding-bottom: 6px; margin-top: 32px; }}
h3 {{ color: #d94b24; margin-bottom: 4px; }}
.meta {{ color: #666; margin-bottom: 24px; }}
.stats {{ border-collapse: collapse; width: 100%; margin: 12px 0 24px; }}
.stats th, .stats td {{ border: 1px solid #ddd; padding: 8px; text-align: right; }}
.stats th:first-child, .stats td:first-child {{ text-align: left; }}
.stats th {{ background: #fff0ea; }}
.chart {{ margin: 24px 0; overflow-x: auto; }}
</style>
</head>
<body>
<h1>Foxsay — Laporan Analisis Data</h1>
<div class="meta">File: {judul_file} · {len(df):,} baris · {len(df.columns)} kolom</div>
<h2>Statistik Terverifikasi</h2>
{tabel_statistik}
<h2>Analisis Bisnis Terstruktur</h2>
{''.join(bagian_analisis) or '<p>Belum ada analisis lanjutan.</p>'}
<h2>Grafik</h2>
{''.join(f'<div class="chart">{grafik}</div>' for grafik in grafik_html) or '<p>Belum ada grafik yang dapat dibuat.</p>'}
<hr>
<p class="meta">Dibuat oleh Foxsay. Interpretasi dan hipotesis perlu dikonfirmasi dengan konteks bisnis.</p>
</body>
</html>"""

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

    if kolom_numerik:
        st.markdown("**Statistik terverifikasi**")
        statistik = df[kolom_numerik].describe().T.reset_index()
        statistik = statistik.rename(columns={
            "index": "Kolom",
            "count": "Jumlah",
            "mean": "Rata-rata",
            "std": "Std",
            "min": "Minimum",
            "50%": "Median",
            "max": "Maksimum",
        })
        kolom_tampil = ["Kolom", "Jumlah", "Rata-rata", "Std", "Minimum", "Median", "Maksimum"]
        st.dataframe(statistik[kolom_tampil].round(4), hide_index=True, use_container_width=True)

    analisis_terstruktur = buat_analisis_terstruktur(df)
    with st.expander("Analisis bisnis terstruktur", expanded=True):
        for judul, daftar in analisis_terstruktur.items():
            if daftar:
                st.markdown(f"**{judul}**")
                for item in daftar:
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

    tampilkan_grafik_interaktif(df)

    laporan_html = buat_laporan_html(df, st.session_state.nama_file)
    st.download_button(
        "Download laporan analisis (HTML)",
        data=laporan_html,
        file_name="foxsay-laporan-analisis.html",
        mime="text/html",
        key="download_laporan_html",
    )

def cari_internet(query):
    try:
        hasil = DDGS().text(query, max_results=5)
        teks_hasil = ""
        for item in hasil:
            teks_hasil += f"- {item['title']}: {item['body']}\n"
        return teks_hasil if teks_hasil else "(tidak ada hasil pencarian)"
    except Exception:
        return "(pencarian gagal, jawab pakai pengetahuan umum saja)"

def is_perintah_wrapped(teks):
    """Kenali perintah rahasia tanpa membuatnya muncul sebagai tombol di UI."""
    return " ".join(teks.strip().lower().split()) == "foxsay wrapped"

def buat_konteks_wrapped():
    """Siapkan riwayat sesi untuk perintah Foxsay Wrapped."""
    if not st.session_state.history:
        return ""

    pesan_user = [teks for role, teks in st.session_state.history if role == "user"]
    potongan_riwayat = []
    for role, teks in st.session_state.history[-40:]:
        nama_role = "USER" if role == "user" else "FOXSAY"
        potongan_riwayat.append(f"{nama_role}: {teks}")

    return f"""RIWAYAT SESI FOXSAY:
Jumlah pesan user: {len(pesan_user)}
Jumlah pesan AI: {sum(1 for role, _ in st.session_state.history if role == 'ai')}

Percakapan terakhir:
{chr(10).join(potongan_riwayat)}
"""

def agent(pertanyaan, konteks_file="", mode=None, special_mode=None):
    if client is None:
        return "GROQ_API_KEY belum dikonfigurasi. Tambahkan secret tersebut untuk menggunakan chat AI Foxsay."

    mode_aktif = mode or st.session_state.mode_aktif
    instruksi_mode = MODE_INSTRUCTIONS.get(mode_aktif, MODE_INSTRUCTIONS["Chill"])

    if special_mode == "wrapped":
        konteks_wrapped = buat_konteks_wrapped()
        if not konteks_wrapped:
            return "Belum ada cukup percakapan untuk dibuat Foxsay Wrapped. Ngobrol dulu, bhapp!"
        sumber_info = f"""{konteks_wrapped}

PERINTAH KHUSUS FOXSAY WRAPPED:
Buat rangkuman personal dari sesi ini berdasarkan riwayat di atas saja.
Jangan mengarang topik, angka, kebiasaan, atau kesimpulan yang tidak terlihat di riwayat.
"""
    elif st.session_state.df_aktif is not None:
        dataframe = st.session_state.df_aktif
        statistik_terverifikasi = buat_ringkasan_statistik(dataframe)
        analisis_terstruktur = buat_analisis_terstruktur(dataframe)
        analisis_terstruktur_teks = "\n".join(
            f"{judul}:\n- " + "\n- ".join(daftar)
            for judul, daftar in analisis_terstruktur.items()
            if daftar
        )
        sampel_data = dataframe.head(50).to_string(index=False)
        sumber_info = f"""Analisis data harus memakai statistik terverifikasi berikut:
{statistik_terverifikasi}

Analisis terstruktur:
{analisis_terstruktur_teks}

Sampel maksimal 50 baris untuk konteks tambahan:
{sampel_data}

ATURAN ANALISIS DATA:
- Jangan menghitung ulang atau menebak angka yang tidak ada di statistik terverifikasi.
- Jika data tidak cukup untuk menjawab, katakan dengan jelas.
- Korelasi menunjukkan hubungan, bukan sebab-akibat.
- Bedakan fakta statistik dari interpretasi atau dugaan.
- Jangan menyebut hipotesis sebagai fakta.
- Jangan mengklaim dampak bisnis, LTV, atau penyebab tanpa kolom/data pendukung.
- Sebutkan data hilang jika relevan.
"""
    elif konteks_file:
        sumber_info = f"""Isi file yang diupload user ({st.session_state.nama_file}):
{konteks_file[:8000]}"""
    else:
        hasil_search = cari_internet(pertanyaan)
        sumber_info = f"""Info dari internet (kalau ada):
{hasil_search}"""

    prompt = f"""{SYSTEM_STYLE}

MODE AKTIF: {mode_aktif}
Instruksi mode: {instruksi_mode}

{sumber_info}

    Pertanyaan: {pertanyaan}

Jawab dengan gaya di atas."""

    if special_mode == "wrapped":
        prompt += """

FORMAT FOXSAY WRAPPED:
- Judul yang playful.
- Jumlah pesan dan ringkasan topik utama.
- Vibe atau pola komunikasi user dengan bahasa yang tidak menghakimi.
- Momen atau tema yang paling sering muncul.
- 2-3 penghargaan lucu yang tetap berdasarkan percakapan.
- Satu rekomendasi atau tantangan kecil untuk sesi berikutnya.
Gunakan bahasa Indonesia santai. Tulis sebagai rangkuman sesi, bukan profil permanen user.
"""

    response = client.chat.completions.create(
        model="openai/gpt-oss-120b",
        messages=[{"role": "user", "content": prompt}]
    )
    return response.choices[0].message.content

def jawab_pengguna(pertanyaan, konteks_file=""):
    """Jalankan perintah rahasia atau teruskan pertanyaan ke agent biasa."""
    if is_perintah_wrapped(pertanyaan):
        return agent(
            "Buat Foxsay Wrapped dari sesi percakapan ini.",
            mode="Chill",
            special_mode="wrapped",
        )
    return agent(pertanyaan, konteks_file)

# Riwayat chat ditampilkan dulu (terbaru di atas nanti setelah input)

if st.session_state.nama_file:
    nama_file_aman = escape(str(st.session_state.nama_file))
    st.markdown(f'<span class="file-chip">📎 {nama_file_aman}</span>', unsafe_allow_html=True)

if st.session_state.file_read_warning:
    st.warning(st.session_state.file_read_warning)
if st.session_state.excel_sheet_info:
    st.info(st.session_state.excel_sheet_info)

if st.session_state.show_uploader:
    uploaded_file = st.file_uploader("Upload PDF, Word, Excel, atau CSV", type=["pdf", "docx", "xlsx", "csv"], label_visibility="collapsed")
    if uploaded_file is not None and uploaded_file.name != st.session_state.nama_file:
        with st.spinner("🦊 Foxsay lagi baca file..."):
            try:
                st.session_state.df_aktif = None
                st.session_state.file_read_warning = ""
                st.session_state.excel_sheet_info = ""
                if uploaded_file.name.endswith(".pdf"):
                    isi = baca_pdf(uploaded_file)
                    if not isi.strip():
                        st.session_state.file_read_warning = (
                            "PDF berhasil dibuka, tetapi tidak berisi teks yang bisa diekstrak. "
                            "Kemungkinan PDF berupa scan/gambar; OCR belum tersedia."
                        )
                elif uploaded_file.name.endswith(".docx"):
                    isi = baca_docx(uploaded_file)
                elif uploaded_file.name.endswith(".xlsx"):
                    nama_sheet = daftar_sheet_excel(uploaded_file)
                    if len(nama_sheet) > 1:
                        st.session_state.excel_sheet_info = (
                            f"Excel memiliki {len(nama_sheet)} sheet ({', '.join(nama_sheet)}). "
                            f"Foxsay saat ini membaca sheet pertama: {nama_sheet[0]}."
                        )
                    df = baca_excel(uploaded_file, sheet_name=0)
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
            tambah_riwayat("user", f"[Minta rangkuman: {st.session_state.nama_file}]")
            tambah_riwayat("ai", jawaban)
            st.rerun()
    with col_b:
        if st.button("Hapus file"):
            st.session_state.isi_file = ""
            st.session_state.nama_file = ""
            st.session_state.df_aktif = None
            st.session_state.file_read_warning = ""
            st.session_state.excel_sheet_info = ""
            st.rerun()

st.write("Tanya apa aja ke Foxsay:")
kotak_utama = st.container(border=True)
with kotak_utama:
    kolom_mode, kolom_info_mode = st.columns([1.5, 4.5])
    with kolom_mode:
        mode_aktif = st.selectbox("Mode Foxsay", list(MODE_INSTRUCTIONS.keys()), key="mode_aktif")
    with kolom_info_mode:
        st.caption("Chill adalah mode default. Pilih Study, Creator, atau Career sesuai kebutuhan.")

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
                    if getattr(audio_value, "size", 0) > 100 * 1024 * 1024:
                        raise ValueError("Ukuran audio melebihi batas Groq 100 MB.")
                    transkrip = client.audio.transcriptions.create(
                        file=buat_payload_audio(audio_value),
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
                        jawaban = jawab_pengguna(pertanyaan_vn, st.session_state.isi_file)
                    tambah_riwayat("user", f"🎤 {pertanyaan_vn}")
                    tambah_riwayat("ai", jawaban)
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
        jawaban = jawab_pengguna(pertanyaan, st.session_state.isi_file)
    tambah_riwayat("user", pertanyaan)
    tambah_riwayat("ai", jawaban)

if st.session_state.history:
    kolom_riwayat, kolom_clear = st.columns([4, 1])
    with kolom_riwayat:
        jumlah_tampil = min(len(st.session_state.history), MAX_HISTORY_RENDER)
        st.caption(f"Riwayat chat — menampilkan {jumlah_tampil} pesan terbaru")
    with kolom_clear:
        if st.button("Hapus chat", key="clear_chat"):
            st.session_state.history.clear()
            st.rerun()

for role, teks in reversed(st.session_state.history[-MAX_HISTORY_RENDER:]):
    teks_aman = escape(str(teks))
    if role == "user":
        st.markdown(f'<div class="chat-bubble-user">🙋 {teks_aman}</div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="chat-bubble-ai">🦊 {teks_aman}</div>', unsafe_allow_html=True)
