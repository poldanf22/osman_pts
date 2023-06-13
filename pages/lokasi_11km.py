import streamlit as st
import xlsxwriter
import tempfile
import pandas as pd

# menghilangkan hamburger
st.markdown("""
<style>
.css-1rs6os.edgvbvh3
{
    visibility:hidden;
}
.css-1lsmgbg.egzxvld0
{
    visibility:hidden;
}
</style>
""", unsafe_allow_html=True)

st.info("Jika melihat pesan error di paling bawah, silahkan refresh")
st.title("Olahan untuk Lokasi")
st.header("11 SMA KM")

col6 = st.container()

with col6:
    KELAS = st.selectbox(
        "KELAS",
        ("--Pilih Kelas--", "11 SMA"))

col7 = st.container()

with col7:
    SEMESTER = st.selectbox(
        "SEMESTER",
        ("--Pilih Semester--", "SEMESTER 1", "SEMESTER 2"))

col8 = st.container()

with col8:
    PENILAIAN = st.selectbox(
        "PENILAIAN",
        ("--Pilih Penilaian--", "PENILAIAN TENGAH SEMESTER", "PENILAIAN AKHIR SEMESTER", "PENILAIAN AKHIR TAHUN"))

col9 = st.container()

with col9:
    KURIKULUM = st.selectbox(
        "KURIKULUM",
        ("--Pilih Kurikulum--", "KM"))

TAHUN = st.text_input("Masukkan Tahun Ajaran", value="",
                      placeholder="contoh: 2022-2023")

col1, col2, col3, col4, col5, col6 = st.columns(
    6)

with col1:
    MTW = st.selectbox(
        "JML. SOAL MAW.",
        (15, 20, 25, 30, 35, 40, 45, 50))

with col2:
    MTP = st.selectbox(
        "JML. SOAL MAP.",
        (15, 20, 25, 30, 35, 40, 45, 50))

with col3:
    IND = st.selectbox(
        "JML. SOAL IND.",
        (15, 20, 25, 30, 35, 40, 45, 50))

with col4:
    ENG = st.selectbox(
        "JML. SOAL ENG.",
        (15, 20, 25, 30, 35, 40, 45, 50))

with col5:
    SEJ = st.selectbox(
        "JML. SOAL SEJ.",
        (15, 20, 25, 30, 35, 40, 45, 50))

with col6:
    GEO = st.selectbox(
        "JML. SOAL GEO.",
        (15, 20, 25, 30, 35, 40, 45, 50))

col7, col8, col9, col10, col11 = st.columns(5)

with col7:
    EKO = st.selectbox(
        "JML. SOAL EKO.",
        (15, 20, 25, 30, 35, 40, 45, 50))

with col8:
    SOS = st.selectbox(
        "JML. SOAL SOS.",
        (15, 20, 25, 30, 35, 40, 45, 50))

with col9:
    FIS = st.selectbox(
        "JML. SOAL FIS.",
        (15, 20, 25, 30, 35, 40, 45, 50))

with col10:
    KIM = st.selectbox(
        "JML. SOAL KIM.",
        (15, 20, 25, 30, 35, 40, 45, 50))

with col11:
    BIO = st.selectbox(
        "JML. SOAL BIO.",
        (15, 20, 25, 30, 35, 40, 45, 50))


JML_SOAL_MAW = MTW
JML_SOAL_MAP = MTP
JML_SOAL_IND = IND
JML_SOAL_ENG = ENG
JML_SOAL_SEJ = SEJ
JML_SOAL_GEO = GEO
JML_SOAL_EKO = EKO
JML_SOAL_SOS = SOS
JML_SOAL_FIS = FIS
JML_SOAL_KIM = KIM
JML_SOAL_BIO = BIO

kelas = KELAS
semester = SEMESTER
tahun = TAHUN
penilaian = PENILAIAN
kurikulum = KURIKULUM

uploaded_file = st.file_uploader(
    'Letakkan file excel NILAI STANDAR [LOKASI 101-160]', type='xlsx')

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    len_col = df.shape[1]

    r = df.shape[0]-5  # baris average
    s = df.shape[0]-4  # baris stdev
    t = df.shape[0]-3  # baris max
    u = df.shape[0]-2  # baris min

    # JUMLAH PESERTA
    peserta = df.iloc[r, len_col-283]

    # rata-rata jumlah benar
    rata_maw = df.iloc[r, len_col-48]
    rata_map = df.iloc[r, len_col-47]
    rata_ind = df.iloc[r, len_col-46]
    rata_eng = df.iloc[r, len_col-45]
    rata_sej = df.iloc[r, len_col-44]
    rata_geo = df.iloc[r, len_col-43]
    rata_eko = df.iloc[r, len_col-42]
    rata_sos = df.iloc[r, len_col-41]
    rata_fis = df.iloc[r, len_col-40]
    rata_kim = df.iloc[r, len_col-39]
    rata_bio = df.iloc[r, len_col-38]
    rata_jml = df.iloc[r, len_col-37]

    # rata-rata nilai standar
    rata_Smaw = df.iloc[r, len_col-25]
    rata_Smap = df.iloc[r, len_col-24]
    rata_Sind = df.iloc[r, len_col-23]
    rata_Seng = df.iloc[r, len_col-22]
    rata_Ssej = df.iloc[r, len_col-21]
    rata_Sgeo = df.iloc[r, len_col-20]
    rata_Seko = df.iloc[r, len_col-19]
    rata_Ssos = df.iloc[r, len_col-18]
    rata_Sfis = df.iloc[r, len_col-17]
    rata_Skim = df.iloc[r, len_col-16]
    rata_Sbio = df.iloc[r, len_col-15]
    rata_Sjml = df.iloc[r, len_col-14]

    # max jumlah benar
    max_maw = df.iloc[t, len_col-48]
    max_map = df.iloc[t, len_col-47]
    max_ind = df.iloc[t, len_col-46]
    max_eng = df.iloc[t, len_col-45]
    max_sej = df.iloc[t, len_col-44]
    max_geo = df.iloc[t, len_col-43]
    max_eko = df.iloc[t, len_col-42]
    max_sos = df.iloc[t, len_col-41]
    max_fis = df.iloc[t, len_col-40]
    max_kim = df.iloc[t, len_col-39]
    max_bio = df.iloc[t, len_col-38]
    max_jml = df.iloc[t, len_col-37]

    # max nilai standar
    max_Smaw = df.iloc[t, len_col-25]
    max_Smap = df.iloc[t, len_col-24]
    max_Sind = df.iloc[t, len_col-23]
    max_Seng = df.iloc[t, len_col-22]
    max_Ssej = df.iloc[t, len_col-21]
    max_Sgeo = df.iloc[t, len_col-20]
    max_Seko = df.iloc[t, len_col-19]
    max_Ssos = df.iloc[t, len_col-18]
    max_Sfis = df.iloc[t, len_col-17]
    max_Skim = df.iloc[t, len_col-16]
    max_Sbio = df.iloc[t, len_col-15]
    max_Sjml = df.iloc[t, len_col-14]

    # min jumlah benar
    min_maw = df.iloc[u, len_col-48]
    min_map = df.iloc[u, len_col-47]
    min_ind = df.iloc[u, len_col-46]
    min_eng = df.iloc[u, len_col-45]
    min_sej = df.iloc[u, len_col-44]
    min_geo = df.iloc[u, len_col-43]
    min_eko = df.iloc[u, len_col-42]
    min_sos = df.iloc[u, len_col-41]
    min_fis = df.iloc[u, len_col-40]
    min_kim = df.iloc[u, len_col-39]
    min_bio = df.iloc[u, len_col-38]
    min_jml = df.iloc[u, len_col-37]

    # min nilai standar
    min_Smaw = df.iloc[u, len_col-25]
    min_Smap = df.iloc[u, len_col-24]
    min_Sind = df.iloc[u, len_col-23]
    min_Seng = df.iloc[u, len_col-22]
    min_Ssej = df.iloc[u, len_col-21]
    min_Sgeo = df.iloc[u, len_col-20]
    min_Seko = df.iloc[u, len_col-19]
    min_Ssos = df.iloc[u, len_col-18]
    min_Sfis = df.iloc[u, len_col-17]
    min_Skim = df.iloc[u, len_col-16]
    min_Sbio = df.iloc[u, len_col-15]
    min_Sjml = df.iloc[u, len_col-14]

    data_jml_benar = {'BIDANG STUDI': ['MAT. WAJIB (MAW)', 'MAT. PEMINATAN (MAP)', 'B. INDONESIA (IND)', 'B. INGGRIS (ENG)', 'SEJARAH (SEJ)', 'GEOGRAFI (GEO)', 'EKONOMI (EKO)', 'SOSIOLOGI (SOS)', 'FISIKA (FIS)', 'KIMIA (KIM)', 'BIOLOGI(BIO)', 'JUMLAH (JML)'],
                      'TERENDAH': [min_maw, min_map, min_ind, min_eng, min_sej, min_geo, min_eko, min_sos, min_fis, min_kim, min_bio, min_jml],
                      'RATA-RATA': [rata_maw, rata_map, rata_ind, rata_eng, rata_sej, rata_geo, rata_eko, rata_sos, rata_fis, rata_kim, rata_bio, rata_jml],
                      'TERTINGGI': [max_maw, max_map, max_ind, max_eng, max_sej, max_geo, max_eko, max_sos, max_fis, max_kim, max_bio, max_jml]}

    jml_benar = pd.DataFrame(data_jml_benar)

    data_n_standar = {'BIDANG STUDI': ['MAT. WAJIB (MAW)', 'MAT. PEMINATAN (MAP)', 'B. INDONESIA (IND)', 'B. INGGRIS (ENG)', 'SEJARAH (SEJ)', 'GEOGRAFI (GEO)', 'EKONOMI (EKO)', 'SOSIOLOGI (SOS)', 'FISIKA (FIS)', 'KIMIA (KIM)', 'BIOLOGI(BIO)', 'JUMLAH (JML)'],
                      'TERENDAH': [min_Smaw, min_Smap, min_Sind, min_Seng, min_Ssej, min_Sgeo, min_Seko, min_Ssos, min_Sfis, min_Skim, min_Sbio, min_Sjml],
                      'RATA-RATA': [rata_Smaw, rata_Smap, rata_Sind, rata_Seng, rata_Ssej, rata_Sgeo, rata_Seko, rata_Ssos, rata_Sfis, rata_Skim, rata_Sbio, rata_Sjml],
                      'TERTINGGI': [max_Smaw, max_Smap, max_Sind, max_Seng, max_Ssej, max_Sgeo, max_Seko, max_Ssos, max_Sfis, max_Skim, max_Sbio, max_Sjml]}

    n_standar = pd.DataFrame(data_n_standar)

    data_jml_peserta = {'JUMLAH PESERTA': [peserta]}

    jml_peserta = pd.DataFrame(data_jml_peserta)

    data_jml_soal = {'BIDANG STUDI': ['MAW', 'MAP', 'IND', 'ENG', 'SEJ', 'GEO', 'EKO', 'SOS', 'FIS', 'KIM', 'BIO'],
                     'JUMLAH': [JML_SOAL_MAW, JML_SOAL_MAP, JML_SOAL_IND, JML_SOAL_ENG, JML_SOAL_SEJ, JML_SOAL_GEO, JML_SOAL_EKO, JML_SOAL_SOS, JML_SOAL_FIS, JML_SOAL_KIM, JML_SOAL_BIO]}

    jml_soal = pd.DataFrame(data_jml_soal)

    df = df[['LOKASI', 'RANK LOK.', 'RANK NAS.', 'NOMOR NF', 'NAMA SISWA', 'NAMA SEKOLAH', 'KELAS', 'MAW', 'MAP', 'IND',
             'ENG', 'SEJ', 'GEO', 'EKO', 'SOS', 'FIS', 'KIM', 'BIO', 'JML', 'S_MAW', 'S_MAP', 'S_IND', 'S_ENG', 'S_SEJ', 'S_GEO', 'S_EKO', 'S_SOS', 'S_FIS', 'S_KIM', 'S_BIO', 'S_JML']]

    # sort setiap lokasi
    sort101 = df[df['LOKASI'] == 101]
    sort102 = df[df['LOKASI'] == 102]
    sort103 = df[df['LOKASI'] == 103]
    # sort104 = df[df['LOKASI']==104]
    sort105 = df[df['LOKASI'] == 105]
    sort106 = df[df['LOKASI'] == 106]
    sort107 = df[df['LOKASI'] == 107]
    sort108 = df[df['LOKASI'] == 108]
    sort109 = df[df['LOKASI'] == 109]
    sort110 = df[df['LOKASI'] == 110]
    sort111 = df[df['LOKASI'] == 111]
    sort112 = df[df['LOKASI'] == 112]
    sort113 = df[df['LOKASI'] == 113]
    # sort114 = df[df['LOKASI']==114]
    sort115 = df[df['LOKASI'] == 115]
    sort116 = df[df['LOKASI'] == 116]
    sort117 = df[df['LOKASI'] == 117]
    sort118 = df[df['LOKASI'] == 118]
    sort119 = df[df['LOKASI'] == 119]
    sort120 = df[df['LOKASI'] == 120]
    sort121 = df[df['LOKASI'] == 121]
    sort122 = df[df['LOKASI'] == 122]
    sort123 = df[df['LOKASI'] == 123]
    sort124 = df[df['LOKASI'] == 124]
    sort125 = df[df['LOKASI'] == 125]
    sort126 = df[df['LOKASI'] == 126]
    sort127 = df[df['LOKASI'] == 127]
    sort128 = df[df['LOKASI'] == 128]
    sort129 = df[df['LOKASI'] == 129]
    sort130 = df[df['LOKASI'] == 130]
    sort131 = df[df['LOKASI'] == 131]
    sort132 = df[df['LOKASI'] == 132]
    sort133 = df[df['LOKASI'] == 133]
    sort134 = df[df['LOKASI'] == 134]
    sort135 = df[df['LOKASI'] == 135]
    sort136 = df[df['LOKASI'] == 136]
    sort137 = df[df['LOKASI'] == 137]
    sort138 = df[df['LOKASI'] == 138]
    # sort139 = df[df['LOKASI']==139]
    sort140 = df[df['LOKASI'] == 140]
    sort141 = df[df['LOKASI'] == 141]
    sort142 = df[df['LOKASI'] == 142]
    sort143 = df[df['LOKASI'] == 143]
    sort144 = df[df['LOKASI'] == 144]
    sort145 = df[df['LOKASI'] == 145]
    sort146 = df[df['LOKASI'] == 146]
    sort148 = df[df['LOKASI'] == 148]
    sort149 = df[df['LOKASI'] == 149]
    sort150 = df[df['LOKASI'] == 150]
    sort151 = df[df['LOKASI'] == 151]
    sort152 = df[df['LOKASI'] == 152]
    sort153 = df[df['LOKASI'] == 153]
    sort154 = df[df['LOKASI'] == 154]
    sort155 = df[df['LOKASI'] == 155]
    sort156 = df[df['LOKASI'] == 156]
    sort157 = df[df['LOKASI'] == 157]
    sort158 = df[df['LOKASI'] == 158]
    sort159 = df[df['LOKASI'] == 159]
    sort160 = df[df['LOKASI'] == 160]

    # 10 besar setiap lokasi
    # 101
    sort101_10 = sort101.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort101_10['LOKASI']
    sort101_10 = sort101_10.drop(
        sort101_10[(sort101_10['RANK LOK.'] > 10)].index)
    # 102
    sort102_10 = sort102.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort102_10['LOKASI']
    sort102_10 = sort102_10.drop(
        sort102_10[(sort102_10['RANK LOK.'] > 10)].index)
    # 103
    sort103_10 = sort103.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort103_10['LOKASI']
    sort103_10 = sort103_10.drop(
        sort103_10[(sort103_10['RANK LOK.'] > 10)].index)
    # # 104
    # sort104_10=sort104.sort_values(by=['RANK LOK.'], ascending=[True])
    # del sort104_10['LOKASI']
    # sort104_10=sort104_10.drop(sort104_10[(sort104_10['RANK LOK.']>10)].index)
    # 105
    sort105_10 = sort105.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort105_10['LOKASI']
    sort105_10 = sort105_10.drop(
        sort105_10[(sort105_10['RANK LOK.'] > 10)].index)
    # 106
    sort106_10 = sort106.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort106_10['LOKASI']
    sort106_10 = sort106_10.drop(
        sort106_10[(sort106_10['RANK LOK.'] > 10)].index)
    # 107
    sort107_10 = sort107.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort107_10['LOKASI']
    sort107_10 = sort107_10.drop(
        sort107_10[(sort107_10['RANK LOK.'] > 10)].index)
    # 108
    sort108_10 = sort108.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort108_10['LOKASI']
    sort108_10 = sort108_10.drop(
        sort108_10[(sort108_10['RANK LOK.'] > 10)].index)
    # 109
    sort109_10 = sort109.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort109_10['LOKASI']
    sort109_10 = sort109_10.drop(
        sort109_10[(sort109_10['RANK LOK.'] > 10)].index)
    # 110
    sort110_10 = sort110.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort110_10['LOKASI']
    sort110_10 = sort110_10.drop(
        sort110_10[(sort110_10['RANK LOK.'] > 10)].index)
    # 111
    sort111_10 = sort111.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort111_10['LOKASI']
    sort111_10 = sort111_10.drop(
        sort111_10[(sort111_10['RANK LOK.'] > 10)].index)
    # 112
    sort112_10 = sort112.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort112_10['LOKASI']
    sort112_10 = sort112_10.drop(
        sort112_10[(sort112_10['RANK LOK.'] > 10)].index)
    # 113
    sort113_10 = sort113.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort113_10['LOKASI']
    sort113_10 = sort113_10.drop(
        sort113_10[(sort113_10['RANK LOK.'] > 10)].index)
    # # 114
    # sort114_10=sort114.sort_values(by=['RANK LOK.'], ascending=[True])
    # del sort114_10['LOKASI']
    # sort114_10=sort114_10.drop(sort114_10[(sort114_10['RANK LOK.']>10)].index)
    # 115
    sort115_10 = sort115.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort115_10['LOKASI']
    sort115_10 = sort115_10.drop(
        sort115_10[(sort115_10['RANK LOK.'] > 10)].index)
    # 116
    sort116_10 = sort116.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort116_10['LOKASI']
    sort116_10 = sort116_10.drop(
        sort116_10[(sort116_10['RANK LOK.'] > 10)].index)
    # 117
    sort117_10 = sort117.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort117_10['LOKASI']
    sort117_10 = sort117_10.drop(
        sort117_10[(sort117_10['RANK LOK.'] > 10)].index)
    # 118
    sort118_10 = sort118.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort118_10['LOKASI']
    sort118_10 = sort118_10.drop(
        sort118_10[(sort118_10['RANK LOK.'] > 10)].index)
    # 119
    sort119_10 = sort119.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort119_10['LOKASI']
    sort119_10 = sort119_10.drop(
        sort119_10[(sort119_10['RANK LOK.'] > 10)].index)
    # 120
    sort120_10 = sort120.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort120_10['LOKASI']
    sort120_10 = sort120_10.drop(
        sort120_10[(sort120_10['RANK LOK.'] > 10)].index)
    # 121
    sort121_10 = sort121.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort121_10['LOKASI']
    sort121_10 = sort121_10.drop(
        sort121_10[(sort121_10['RANK LOK.'] > 10)].index)
    # 122
    sort122_10 = sort122.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort122_10['LOKASI']
    sort122_10 = sort122_10.drop(
        sort122_10[(sort122_10['RANK LOK.'] > 10)].index)
    # 123
    sort123_10 = sort123.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort123_10['LOKASI']
    sort123_10 = sort123_10.drop(
        sort123_10[(sort123_10['RANK LOK.'] > 10)].index)
    # 124
    sort124_10 = sort124.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort124_10['LOKASI']
    sort124_10 = sort124_10.drop(
        sort124_10[(sort124_10['RANK LOK.'] > 10)].index)
    # 125
    sort125_10 = sort125.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort125_10['LOKASI']
    sort125_10 = sort125_10.drop(
        sort125_10[(sort125_10['RANK LOK.'] > 10)].index)
    # 126
    sort126_10 = sort126.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort126_10['LOKASI']
    sort126_10 = sort126_10.drop(
        sort126_10[(sort126_10['RANK LOK.'] > 10)].index)
    # 127
    sort127_10 = sort127.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort127_10['LOKASI']
    sort127_10 = sort127_10.drop(
        sort127_10[(sort127_10['RANK LOK.'] > 10)].index)
    # 128
    sort128_10 = sort128.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort128_10['LOKASI']
    sort128_10 = sort128_10.drop(
        sort128_10[(sort128_10['RANK LOK.'] > 10)].index)
    # 129
    sort129_10 = sort129.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort129_10['LOKASI']
    sort129_10 = sort129_10.drop(
        sort129_10[(sort129_10['RANK LOK.'] > 10)].index)
    # 130
    sort130_10 = sort130.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort130_10['LOKASI']
    sort130_10 = sort130_10.drop(
        sort130_10[(sort130_10['RANK LOK.'] > 10)].index)
    # 131
    sort131_10 = sort131.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort131_10['LOKASI']
    sort131_10 = sort131_10.drop(
        sort131_10[(sort131_10['RANK LOK.'] > 10)].index)
    # 132
    sort132_10 = sort132.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort132_10['LOKASI']
    sort132_10 = sort132_10.drop(
        sort132_10[(sort132_10['RANK LOK.'] > 10)].index)
    # 133
    sort133_10 = sort133.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort133_10['LOKASI']
    sort133_10 = sort133_10.drop(
        sort133_10[(sort133_10['RANK LOK.'] > 10)].index)
    # 134
    sort134_10 = sort134.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort134_10['LOKASI']
    sort134_10 = sort134_10.drop(
        sort134_10[(sort134_10['RANK LOK.'] > 10)].index)
    # 135
    sort135_10 = sort135.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort135_10['LOKASI']
    sort135_10 = sort135_10.drop(
        sort135_10[(sort135_10['RANK LOK.'] > 10)].index)
    # 136
    sort136_10 = sort136.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort136_10['LOKASI']
    sort136_10 = sort136_10.drop(
        sort136_10[(sort136_10['RANK LOK.'] > 10)].index)
    # 137
    sort137_10 = sort137.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort137_10['LOKASI']
    sort137_10 = sort137_10.drop(
        sort137_10[(sort137_10['RANK LOK.'] > 10)].index)
    # 138
    sort138_10 = sort138.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort138_10['LOKASI']
    sort138_10 = sort138_10.drop(
        sort138_10[(sort138_10['RANK LOK.'] > 10)].index)
    # # 139
    # sort139_10=sort139.sort_values(by=['RANK LOK.'], ascending=[True])
    # del sort139_10['LOKASI']
    # sort139_10=sort139_10.drop(sort139_10[(sort139_10['RANK LOK.']>10)].index)
    # 140
    sort140_10 = sort140.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort140_10['LOKASI']
    sort140_10 = sort140_10.drop(
        sort140_10[(sort140_10['RANK LOK.'] > 10)].index)
    # 141
    sort141_10 = sort141.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort141_10['LOKASI']
    sort141_10 = sort141_10.drop(
        sort141_10[(sort141_10['RANK LOK.'] > 10)].index)
    # 142
    sort142_10 = sort142.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort142_10['LOKASI']
    sort142_10 = sort142_10.drop(
        sort142_10[(sort142_10['RANK LOK.'] > 10)].index)
    # 143
    sort143_10 = sort143.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort143_10['LOKASI']
    sort143_10 = sort143_10.drop(
        sort143_10[(sort143_10['RANK LOK.'] > 10)].index)
    # 144
    sort144_10 = sort144.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort144_10['LOKASI']
    sort144_10 = sort144_10.drop(
        sort144_10[(sort144_10['RANK LOK.'] > 10)].index)
    # 145
    sort145_10 = sort145.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort145_10['LOKASI']
    sort145_10 = sort145_10.drop(
        sort145_10[(sort145_10['RANK LOK.'] > 10)].index)
    # 146
    sort146_10 = sort146.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort146_10['LOKASI']
    sort146_10 = sort146_10.drop(
        sort146_10[(sort146_10['RANK LOK.'] > 10)].index)
    # 148
    sort148_10 = sort148.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort148_10['LOKASI']
    sort148_10 = sort148_10.drop(
        sort148_10[(sort148_10['RANK LOK.'] > 10)].index)
    # 149
    sort149_10 = sort149.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort149_10['LOKASI']
    sort149_10 = sort149_10.drop(
        sort149_10[(sort149_10['RANK LOK.'] > 10)].index)
    # 150
    sort150_10 = sort150.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort150_10['LOKASI']
    sort150_10 = sort150_10.drop(
        sort150_10[(sort150_10['RANK LOK.'] > 10)].index)
    # 151
    sort151_10 = sort151.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort151_10['LOKASI']
    sort151_10 = sort151_10.drop(
        sort151_10[(sort151_10['RANK LOK.'] > 10)].index)
    # 152
    sort152_10 = sort152.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort152_10['LOKASI']
    sort152_10 = sort152_10.drop(
        sort152_10[(sort152_10['RANK LOK.'] > 10)].index)
    # 153
    sort153_10 = sort153.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort153_10['LOKASI']
    sort153_10 = sort153_10.drop(
        sort153_10[(sort153_10['RANK LOK.'] > 10)].index)
    # 154
    sort154_10 = sort154.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort154_10['LOKASI']
    sort154_10 = sort154_10.drop(
        sort154_10[(sort154_10['RANK LOK.'] > 10)].index)
    # 155
    sort155_10 = sort155.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort155_10['LOKASI']
    sort155_10 = sort155_10.drop(
        sort155_10[(sort155_10['RANK LOK.'] > 10)].index)
    # 156
    sort156_10 = sort156.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort156_10['LOKASI']
    sort156_10 = sort156_10.drop(
        sort156_10[(sort156_10['RANK LOK.'] > 10)].index)
    # 157
    sort157_10 = sort157.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort157_10['LOKASI']
    sort157_10 = sort157_10.drop(
        sort157_10[(sort157_10['RANK LOK.'] > 10)].index)
    # 158
    sort158_10 = sort158.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort158_10['LOKASI']
    sort158_10 = sort158_10.drop(
        sort158_10[(sort158_10['RANK LOK.'] > 10)].index)
    # 159
    sort159_10 = sort159.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort159_10['LOKASI']
    sort159_10 = sort159_10.drop(
        sort159_10[(sort159_10['RANK LOK.'] > 10)].index)
    # 160
    sort160_10 = sort160.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort160_10['LOKASI']
    sort160_10 = sort160_10.drop(
        sort160_10[(sort160_10['RANK LOK.'] > 10)].index)

    # All 101
    sort101 = sort101.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort101['LOKASI']
    # All 102
    sort102 = sort102.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort102['LOKASI']
    # All 103
    sort103 = sort103.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort103['LOKASI']
    # # All 104
    # sort104=sort104.sort_values(by=['NAMA SISWA'], ascending=[True])
    # del sort104['LOKASI']
    # All 105
    sort105 = sort105.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort105['LOKASI']
    # All 106
    sort106 = sort106.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort106['LOKASI']
    # All 107
    sort107 = sort107.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort107['LOKASI']
    # All 108
    sort108 = sort108.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort108['LOKASI']
    # All 109
    sort109 = sort109.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort109['LOKASI']
    # All 110
    sort110 = sort110.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort110['LOKASI']
    # All 111
    sort111 = sort111.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort111['LOKASI']
    # All 112
    sort112 = sort112.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort112['LOKASI']
    # All 113
    sort113 = sort113.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort113['LOKASI']
    # # All 114
    # sort114=sort114.sort_values(by=['NAMA SISWA'], ascending=[True])
    # del sort114['LOKASI']
    # All 115
    sort115 = sort115.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort115['LOKASI']
    # All 116
    sort116 = sort116.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort116['LOKASI']
    # All 117
    sort117 = sort117.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort117['LOKASI']
    # All 118
    sort118 = sort118.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort118['LOKASI']
    # All 119
    sort119 = sort119.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort119['LOKASI']
    # All 120
    sort120 = sort120.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort120['LOKASI']
    # All 121
    sort121 = sort121.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort121['LOKASI']
    # All 122
    sort122 = sort122.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort122['LOKASI']
    # All 123
    sort123 = sort123.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort123['LOKASI']
    # All 124
    sort124 = sort124.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort124['LOKASI']
    # All 125
    sort125 = sort125.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort125['LOKASI']
    # All 126
    sort126 = sort126.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort126['LOKASI']
    # All 127
    sort127 = sort127.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort127['LOKASI']
    # All 128
    sort128 = sort128.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort128['LOKASI']
    # All 129
    sort129 = sort129.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort129['LOKASI']
    # All 130
    sort130 = sort130.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort130['LOKASI']
    # All 131
    sort131 = sort131.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort131['LOKASI']
    # All 132
    sort132 = sort132.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort132['LOKASI']
    # All 133
    sort133 = sort133.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort133['LOKASI']
    # All 134
    sort134 = sort134.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort134['LOKASI']
    # All 135
    sort135 = sort135.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort135['LOKASI']
    # All 136
    sort136 = sort136.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort136['LOKASI']
    # All 137
    sort137 = sort137.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort137['LOKASI']
    # All 138
    sort138 = sort138.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort138['LOKASI']
    # # All 139
    # sort139=sort139.sort_values(by=['NAMA SISWA'], ascending=[True])
    # del sort139['LOKASI']
    # All 140
    sort140 = sort140.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort140['LOKASI']
    # All 141
    sort141 = sort141.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort141['LOKASI']
    # All 142
    sort142 = sort142.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort142['LOKASI']
    # All 143
    sort143 = sort143.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort143['LOKASI']
    # All 144
    sort144 = sort144.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort144['LOKASI']
    # All 145
    sort145 = sort145.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort145['LOKASI']
    # All 146
    sort146 = sort146.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort146['LOKASI']
    # All 148
    sort148 = sort148.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort148['LOKASI']
    # All 149
    sort149 = sort149.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort149['LOKASI']
    # All 150
    sort150 = sort150.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort150['LOKASI']
    # All 151
    sort151 = sort151.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort151['LOKASI']
    # All 152
    sort152 = sort152.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort152['LOKASI']
    # All 153
    sort153 = sort153.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort153['LOKASI']
    # All 154
    sort154 = sort154.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort154['LOKASI']
    # All 155
    sort155 = sort155.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort155['LOKASI']
    # All 156
    sort156 = sort156.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort156['LOKASI']
    # All 157
    sort157 = sort157.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort157['LOKASI']
    # All 158
    sort158 = sort158.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort158['LOKASI']
    # All 159
    sort159 = sort159.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort159['LOKASI']
    # All 160
    sort160 = sort160.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort160['LOKASI']

    # jumlah row
    # 101
    row101_10 = sort101_10.shape[0]
    row101 = sort101.shape[0]
    # 102
    row102_10 = sort102_10.shape[0]
    row102 = sort102.shape[0]
    # 103
    row103_10 = sort103_10.shape[0]
    row103 = sort103.shape[0]
    # # 104
    # row104_10=sort104_10.shape[0]
    # row104=sort104.shape[0]
    # 105
    row105_10 = sort105_10.shape[0]
    row105 = sort105.shape[0]
    # 106
    row106_10 = sort106_10.shape[0]
    row106 = sort106.shape[0]
    # 107
    row107_10 = sort107_10.shape[0]
    row107 = sort107.shape[0]
    # 108
    row108_10 = sort108_10.shape[0]
    row108 = sort108.shape[0]
    # 109
    row109_10 = sort109_10.shape[0]
    row109 = sort109.shape[0]
    # 110
    row110_10 = sort110_10.shape[0]
    row110 = sort110.shape[0]
    # 111
    row111_10 = sort111_10.shape[0]
    row111 = sort111.shape[0]
    # 112
    row112_10 = sort112_10.shape[0]
    row112 = sort112.shape[0]
    # 113
    row113_10 = sort113_10.shape[0]
    row113 = sort113.shape[0]
    # # 114
    # row114_10=sort114_10.shape[0]
    # row114=sort114.shape[0]
    # 115
    row115_10 = sort115_10.shape[0]
    row115 = sort115.shape[0]
    # 116
    row116_10 = sort116_10.shape[0]
    row116 = sort116.shape[0]
    # 117
    row117_10 = sort117_10.shape[0]
    row117 = sort117.shape[0]
    # 118
    row118_10 = sort118_10.shape[0]
    row118 = sort118.shape[0]
    # 119
    row119_10 = sort119_10.shape[0]
    row119 = sort119.shape[0]
    # 120
    row120_10 = sort120_10.shape[0]
    row120 = sort120.shape[0]
    # 121
    row121_10 = sort121_10.shape[0]
    row121 = sort121.shape[0]
    # 122
    row122_10 = sort122_10.shape[0]
    row122 = sort122.shape[0]
    # 123
    row123_10 = sort123_10.shape[0]
    row123 = sort123.shape[0]
    # 124
    row124_10 = sort124_10.shape[0]
    row124 = sort124.shape[0]
    # 125
    row125_10 = sort125_10.shape[0]
    row125 = sort125.shape[0]
    # 126
    row126_10 = sort126_10.shape[0]
    row126 = sort126.shape[0]
    # 127
    row127_10 = sort127_10.shape[0]
    row127 = sort127.shape[0]
    # 128
    row128_10 = sort128_10.shape[0]
    row128 = sort128.shape[0]
    # 129
    row129_10 = sort129_10.shape[0]
    row129 = sort129.shape[0]
    # 130
    row130_10 = sort130_10.shape[0]
    row130 = sort130.shape[0]
    # 131
    row131_10 = sort131_10.shape[0]
    row131 = sort131.shape[0]
    # 132
    row132_10 = sort132_10.shape[0]
    row132 = sort132.shape[0]
    # 133
    row133_10 = sort133_10.shape[0]
    row133 = sort133.shape[0]
    # 134
    row134_10 = sort134_10.shape[0]
    row134 = sort134.shape[0]
    # 135
    row135_10 = sort135_10.shape[0]
    row135 = sort135.shape[0]
    # 136
    row136_10 = sort136_10.shape[0]
    row136 = sort136.shape[0]
    # 137
    row137_10 = sort137_10.shape[0]
    row137 = sort137.shape[0]
    # 138
    row138_10 = sort138_10.shape[0]
    row138 = sort138.shape[0]
    # # 139
    # row139_10=sort139_10.shape[0]
    # row139=sort139.shape[0]
    # 140
    row140_10 = sort140_10.shape[0]
    row140 = sort140.shape[0]
    # 141
    row141_10 = sort141_10.shape[0]
    row141 = sort141.shape[0]
    # 142
    row142_10 = sort142_10.shape[0]
    row142 = sort142.shape[0]
    # 143
    row143_10 = sort143_10.shape[0]
    row143 = sort143.shape[0]
    # 144
    row144_10 = sort144_10.shape[0]
    row144 = sort144.shape[0]
    # 145
    row145_10 = sort145_10.shape[0]
    row145 = sort145.shape[0]
    # 146
    row146_10 = sort146_10.shape[0]
    row146 = sort146.shape[0]
    # 148
    row148_10 = sort148_10.shape[0]
    row148 = sort148.shape[0]
    # 149
    row149_10 = sort149_10.shape[0]
    row149 = sort149.shape[0]
    # 150
    row150_10 = sort150_10.shape[0]
    row150 = sort150.shape[0]
    # 151
    row151_10 = sort151_10.shape[0]
    row151 = sort151.shape[0]
    # 152
    row152_10 = sort152_10.shape[0]
    row152 = sort152.shape[0]
    # 153
    row153_10 = sort153_10.shape[0]
    row153 = sort153.shape[0]
    # 154
    row154_10 = sort154_10.shape[0]
    row154 = sort154.shape[0]
    # 155
    row155_10 = sort155_10.shape[0]
    row155 = sort155.shape[0]
    # 156
    row156_10 = sort156_10.shape[0]
    row156 = sort156.shape[0]
    # 157
    row157_10 = sort157_10.shape[0]
    row157 = sort157.shape[0]
    # 158
    row158_10 = sort158_10.shape[0]
    row158 = sort158.shape[0]
    # 159
    row159_10 = sort159_10.shape[0]
    row159 = sort159.shape[0]
    # 160
    row160_10 = sort160_10.shape[0]
    row160 = sort160.shape[0]

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    file_name = f"{kelas}_{penilaian}_{semester}_lokasi_101_160.xlsx"
    file_path = tempfile.gettempdir() + '/' + file_name

    # Menyimpan file Excel
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object cover.
    jml_benar.to_excel(writer, sheet_name='cover',
                       startrow=10,
                       startcol=0,
                       index=False,
                       )

    # Convert the dataframe to an XlsxWriter Excel object cover.
    n_standar.to_excel(writer, sheet_name='cover',
                       startrow=27,
                       startcol=0,
                       index=False,
                       header=False)

    # Convert the dataframe to an XlsxWriter Excel object cover.
    jml_peserta.to_excel(writer, sheet_name='cover',
                         startrow=27,
                         startcol=5,
                         index=False,
                         header=False)

    # Convert the dataframe to an XlsxWriter Excel object cover.
    jml_soal.to_excel(writer, sheet_name='cover',
                      startrow=13,
                      startcol=5,
                      index=False,
                      header=False)

    # 101
    # Convert the dataframe to an XlsxWriter Excel object.
    sort101_10.to_excel(writer, sheet_name='101',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort101.to_excel(writer, sheet_name='101',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 102
    # Convert the dataframe to an XlsxWriter Excel object.
    sort102_10.to_excel(writer, sheet_name='102',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort102.to_excel(writer, sheet_name='102',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 103
    # Convert the dataframe to an XlsxWriter Excel object.
    sort103_10.to_excel(writer, sheet_name='103',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort103.to_excel(writer, sheet_name='103',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # # 104
    # # Convert the dataframe to an XlsxWriter Excel object.
    # sort104_10.to_excel(writer, sheet_name='104',
    #                startrow = 5,
    #                startcol = 0,
    #                index = False,
    #                header = False)
    # # Convert the dataframe to an XlsxWriter Excel object.
    # sort104.to_excel(writer, sheet_name='104',
    #                startrow = 22,
    #                startcol = 0,
    #                index = False,
    #                header = False)
    # 105
    # Convert the dataframe to an XlsxWriter Excel object.
    sort105_10.to_excel(writer, sheet_name='105',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort105.to_excel(writer, sheet_name='105',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 106
    # Convert the dataframe to an XlsxWriter Excel object.
    sort106_10.to_excel(writer, sheet_name='106',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort106.to_excel(writer, sheet_name='106',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 107
    # Convert the dataframe to an XlsxWriter Excel object.
    sort107_10.to_excel(writer, sheet_name='107',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort107.to_excel(writer, sheet_name='107',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 108
    # Convert the dataframe to an XlsxWriter Excel object.
    sort108_10.to_excel(writer, sheet_name='108',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort108.to_excel(writer, sheet_name='108',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 109
    # Convert the dataframe to an XlsxWriter Excel object.
    sort109_10.to_excel(writer, sheet_name='109',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort109.to_excel(writer, sheet_name='109',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 110
    # Convert the dataframe to an XlsxWriter Excel object.
    sort110_10.to_excel(writer, sheet_name='110',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort110.to_excel(writer, sheet_name='110',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 111
    # Convert the dataframe to an XlsxWriter Excel object.
    sort111_10.to_excel(writer, sheet_name='111',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort111.to_excel(writer, sheet_name='111',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 112
    # Convert the dataframe to an XlsxWriter Excel object.
    sort112_10.to_excel(writer, sheet_name='112',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort112.to_excel(writer, sheet_name='112',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 113
    # Convert the dataframe to an XlsxWriter Excel object.
    sort113_10.to_excel(writer, sheet_name='113',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort113.to_excel(writer, sheet_name='113',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # # 114
    # # Convert the dataframe to an XlsxWriter Excel object.
    # sort114_10.to_excel(writer, sheet_name='114',
    #                startrow = 5,
    #                startcol = 0,
    #                index = False,
    #                header = False)
    # # Convert the dataframe to an XlsxWriter Excel object.
    # sort114.to_excel(writer, sheet_name='114',
    #                startrow = 22,
    #                startcol = 0,
    #                index = False,
    #                header = False)
    # 115
    # Convert the dataframe to an XlsxWriter Excel object.
    sort115_10.to_excel(writer, sheet_name='115',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort115.to_excel(writer, sheet_name='115',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 116
    # Convert the dataframe to an XlsxWriter Excel object.
    sort116_10.to_excel(writer, sheet_name='116',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort116.to_excel(writer, sheet_name='116',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 117
    # Convert the dataframe to an XlsxWriter Excel object.
    sort117_10.to_excel(writer, sheet_name='117',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort117.to_excel(writer, sheet_name='117',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 118
    # Convert the dataframe to an XlsxWriter Excel object.
    sort118_10.to_excel(writer, sheet_name='118',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort118.to_excel(writer, sheet_name='118',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 119
    # Convert the dataframe to an XlsxWriter Excel object.
    sort119_10.to_excel(writer, sheet_name='119',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort119.to_excel(writer, sheet_name='119',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 120
    # Convert the dataframe to an XlsxWriter Excel object.
    sort120_10.to_excel(writer, sheet_name='120',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort120.to_excel(writer, sheet_name='120',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 121
    # Convert the dataframe to an XlsxWriter Excel object.
    sort121_10.to_excel(writer, sheet_name='121',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort121.to_excel(writer, sheet_name='121',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 122
    # Convert the dataframe to an XlsxWriter Excel object.
    sort122_10.to_excel(writer, sheet_name='122',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort122.to_excel(writer, sheet_name='122',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 123
    # Convert the dataframe to an XlsxWriter Excel object.
    sort123_10.to_excel(writer, sheet_name='123',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort123.to_excel(writer, sheet_name='123',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 124
    # Convert the dataframe to an XlsxWriter Excel object.
    sort124_10.to_excel(writer, sheet_name='124',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort124.to_excel(writer, sheet_name='124',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 125
    # Convert the dataframe to an XlsxWriter Excel object.
    sort125_10.to_excel(writer, sheet_name='125',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort125.to_excel(writer, sheet_name='125',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 126
    # Convert the dataframe to an XlsxWriter Excel object.
    sort126_10.to_excel(writer, sheet_name='126',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort126.to_excel(writer, sheet_name='126',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 127
    # Convert the dataframe to an XlsxWriter Excel object.
    sort127_10.to_excel(writer, sheet_name='127',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort127.to_excel(writer, sheet_name='127',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 128
    # Convert the dataframe to an XlsxWriter Excel object.
    sort128_10.to_excel(writer, sheet_name='128',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort128.to_excel(writer, sheet_name='128',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 129
    # Convert the dataframe to an XlsxWriter Excel object.
    sort129_10.to_excel(writer, sheet_name='129',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort129.to_excel(writer, sheet_name='129',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 130
    # Convert the dataframe to an XlsxWriter Excel object.
    sort130_10.to_excel(writer, sheet_name='130',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort130.to_excel(writer, sheet_name='130',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 131
    # Convert the dataframe to an XlsxWriter Excel object.
    sort131_10.to_excel(writer, sheet_name='131',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort131.to_excel(writer, sheet_name='131',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 132
    # Convert the dataframe to an XlsxWriter Excel object.
    sort132_10.to_excel(writer, sheet_name='132',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort132.to_excel(writer, sheet_name='132',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 133
    # Convert the dataframe to an XlsxWriter Excel object.
    sort133_10.to_excel(writer, sheet_name='133',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort133.to_excel(writer, sheet_name='133',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 134
    # Convert the dataframe to an XlsxWriter Excel object.
    sort134_10.to_excel(writer, sheet_name='134',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort134.to_excel(writer, sheet_name='134',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 135
    # Convert the dataframe to an XlsxWriter Excel object.
    sort135_10.to_excel(writer, sheet_name='135',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort135.to_excel(writer, sheet_name='135',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 136
    # Convert the dataframe to an XlsxWriter Excel object.
    sort136_10.to_excel(writer, sheet_name='136',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort136.to_excel(writer, sheet_name='136',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 137
    # Convert the dataframe to an XlsxWriter Excel object.
    sort137_10.to_excel(writer, sheet_name='137',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort137.to_excel(writer, sheet_name='137',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 138
    # Convert the dataframe to an XlsxWriter Excel object.
    sort138_10.to_excel(writer, sheet_name='138',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort138.to_excel(writer, sheet_name='138',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # # 139
    # # Convert the dataframe to an XlsxWriter Excel object.
    # sort139_10.to_excel(writer, sheet_name='139',
    #                startrow = 5,
    #                startcol = 0,
    #                index = False,
    #                header = False)
    # # Convert the dataframe to an XlsxWriter Excel object.
    # sort139.to_excel(writer, sheet_name='139',
    #                startrow = 22,
    #                startcol = 0,
    #                index = False,
    #                header = False)
    # 140
    # Convert the dataframe to an XlsxWriter Excel object.
    sort140_10.to_excel(writer, sheet_name='140',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort140.to_excel(writer, sheet_name='140',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 141
    # Convert the dataframe to an XlsxWriter Excel object.
    sort141_10.to_excel(writer, sheet_name='141',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort141.to_excel(writer, sheet_name='141',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 142
    # Convert the dataframe to an XlsxWriter Excel object.
    sort142_10.to_excel(writer, sheet_name='142',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort142.to_excel(writer, sheet_name='142',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 143
    # Convert the dataframe to an XlsxWriter Excel object.
    sort143_10.to_excel(writer, sheet_name='143',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort143.to_excel(writer, sheet_name='143',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 144
    # Convert the dataframe to an XlsxWriter Excel object.
    sort144_10.to_excel(writer, sheet_name='144',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort144.to_excel(writer, sheet_name='144',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 145
    # Convert the dataframe to an XlsxWriter Excel object.
    sort145_10.to_excel(writer, sheet_name='145',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort145.to_excel(writer, sheet_name='145',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 146
    # Convert the dataframe to an XlsxWriter Excel object.
    sort146_10.to_excel(writer, sheet_name='146',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort146.to_excel(writer, sheet_name='146',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 148
    # Convert the dataframe to an XlsxWriter Excel object.
    sort148_10.to_excel(writer, sheet_name='148',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort148.to_excel(writer, sheet_name='148',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 149
    # Convert the dataframe to an XlsxWriter Excel object.
    sort149_10.to_excel(writer, sheet_name='149',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort149.to_excel(writer, sheet_name='149',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 150
    # Convert the dataframe to an XlsxWriter Excel object.
    sort150_10.to_excel(writer, sheet_name='150',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort150.to_excel(writer, sheet_name='150',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 151
    # Convert the dataframe to an XlsxWriter Excel object.
    sort151_10.to_excel(writer, sheet_name='151',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort151.to_excel(writer, sheet_name='151',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 152
    # Convert the dataframe to an XlsxWriter Excel object.
    sort152_10.to_excel(writer, sheet_name='152',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort152.to_excel(writer, sheet_name='152',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 153
    # Convert the dataframe to an XlsxWriter Excel object.
    sort153_10.to_excel(writer, sheet_name='153',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort153.to_excel(writer, sheet_name='153',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 154
    # Convert the dataframe to an XlsxWriter Excel object.
    sort154_10.to_excel(writer, sheet_name='154',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort154.to_excel(writer, sheet_name='154',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 155
    # Convert the dataframe to an XlsxWriter Excel object.
    sort155_10.to_excel(writer, sheet_name='155',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort155.to_excel(writer, sheet_name='155',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 156
    # Convert the dataframe to an XlsxWriter Excel object.
    sort156_10.to_excel(writer, sheet_name='156',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort156.to_excel(writer, sheet_name='156',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 157
    # Convert the dataframe to an XlsxWriter Excel object.
    sort157_10.to_excel(writer, sheet_name='157',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort157.to_excel(writer, sheet_name='157',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 158
    # Convert the dataframe to an XlsxWriter Excel object.
    sort158_10.to_excel(writer, sheet_name='158',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort158.to_excel(writer, sheet_name='158',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 159
    # Convert the dataframe to an XlsxWriter Excel object.
    sort159_10.to_excel(writer, sheet_name='159',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort159.to_excel(writer, sheet_name='159',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 160
    # Convert the dataframe to an XlsxWriter Excel object.
    sort160_10.to_excel(writer, sheet_name='160',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort160.to_excel(writer, sheet_name='160',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)

    # Get the xlsxwriter objects from the dataframe writer object.
    workbook = writer.book

    # membuat worksheet baru
    worksheetcover = writer.sheets['cover']
    worksheet101 = writer.sheets['101']
    worksheet102 = writer.sheets['102']
    worksheet103 = writer.sheets['103']
    # worksheet104 = writer.sheets['104']
    worksheet105 = writer.sheets['105']
    worksheet106 = writer.sheets['106']
    worksheet107 = writer.sheets['107']
    worksheet108 = writer.sheets['108']
    worksheet109 = writer.sheets['109']
    worksheet110 = writer.sheets['110']
    worksheet111 = writer.sheets['111']
    worksheet112 = writer.sheets['112']
    worksheet113 = writer.sheets['113']
    # worksheet114 = writer.sheets['114']
    worksheet115 = writer.sheets['115']
    worksheet116 = writer.sheets['116']
    worksheet117 = writer.sheets['117']
    worksheet118 = writer.sheets['118']
    worksheet119 = writer.sheets['119']
    worksheet120 = writer.sheets['120']
    worksheet121 = writer.sheets['121']
    worksheet122 = writer.sheets['122']
    worksheet123 = writer.sheets['123']
    worksheet124 = writer.sheets['124']
    worksheet125 = writer.sheets['125']
    worksheet126 = writer.sheets['126']
    worksheet127 = writer.sheets['127']
    worksheet128 = writer.sheets['128']
    worksheet129 = writer.sheets['129']
    worksheet130 = writer.sheets['130']
    worksheet131 = writer.sheets['131']
    worksheet132 = writer.sheets['132']
    worksheet133 = writer.sheets['133']
    worksheet134 = writer.sheets['134']
    worksheet135 = writer.sheets['135']
    worksheet136 = writer.sheets['136']
    worksheet137 = writer.sheets['137']
    worksheet138 = writer.sheets['138']
    # worksheet139 = writer.sheets['139']
    worksheet140 = writer.sheets['140']
    worksheet141 = writer.sheets['141']
    worksheet142 = writer.sheets['142']
    worksheet143 = writer.sheets['143']
    worksheet144 = writer.sheets['144']
    worksheet145 = writer.sheets['145']
    worksheet146 = writer.sheets['146']
    worksheet148 = writer.sheets['148']
    worksheet149 = writer.sheets['149']
    worksheet150 = writer.sheets['150']
    worksheet151 = writer.sheets['151']
    worksheet152 = writer.sheets['152']
    worksheet153 = writer.sheets['153']
    worksheet154 = writer.sheets['154']
    worksheet155 = writer.sheets['155']
    worksheet156 = writer.sheets['156']
    worksheet157 = writer.sheets['157']
    worksheet158 = writer.sheets['158']
    worksheet159 = writer.sheets['159']
    worksheet160 = writer.sheets['160']

    # format workbook
    titleCover = workbook.add_format({
        'bold': 1,
        'border': 0,
        'align': 'left',
        'valign': 'vcenter',
        'font_color': '#00058E',
        'font_size': 52,
        'font_name': 'Arial Black'})
    sub_titleCover = workbook.add_format({
        'bold': 0,
        'border': 0,
        'align': 'left',
        'valign': 'vcenter',
        'font_size': 27,
        'font_name': 'Arial Unicode MS'})
    headerCover = workbook.add_format({
        'bold': 1,
        'border': 0,
        'align': 'left',
        'valign': 'vcenter',
        'font_size': 24,
        'font_name': 'Arial Rounded MT Bold'})
    sub_headerCover = workbook.add_format({
        'bold': 0,
        'border': 0,
        'align': 'left',
        'valign': 'vcenter',
        'font_size': 16,
        'font_name': 'Arial'})
    sub_header1Cover = workbook.add_format({
        'bold': 1,
        'border': 0,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 20,
        'font_name': 'Arial'})
    kelasCover = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 40,
        'font_name': 'Arial Rounded MT Bold'})
    borderCover = workbook.add_format({
        'bottom': 1,
        'top': 1,
        'left': 1,
        'right': 1})
    centerCover = workbook.add_format({
        'align': 'center',
        'font_size': 12,
        'font_name': 'Arial'})
    center1Cover = workbook.add_format({
        'align': 'center',
        'font_size': 20,
        'font_name': 'Arial'})
    bodyCover = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 12,
        'font_name': 'Arial',
        'bg_color': 'FFF684'})
    center = workbook.add_format({
        'align': 'center',
        'font_size': 10,
        'font_name': 'Arial'})
    left = workbook.add_format({
        'align': 'left',
        'font_size': 10,
        'font_name': 'Arial'})
    title = workbook.add_format({
        'bold': 1,
        'border': 0,
        'align': 'center',
        'valign': 'vcenter',
        'font_color': '#00058E',
        'font_size': 12,
        'font_name': 'Arial'})
    sub_title = workbook.add_format({
        'bold': 1,
        'border': 0,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 12,
        'font_name': 'Arial'})
    subTitle = workbook.add_format({
        'bold': 1,
        'border': 0,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 14,
        'font_name': 'Arial'})
    header = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Arial',
        'bg_color': 'FFF684'})
    body = workbook.add_format({
        'bold': 0,
        'border': 1,
        'align': 'center',
        'font_size': 10,
        'font_name': 'Arial',
        'bg_color': 'FFF684'})
    border = workbook.add_format({
        'bottom': 1,
        'top': 1,
        'left': 1,
        'right': 1})

    # worksheet cover
    # sampai baris 19, dari kolom 1, mulai dari baris 12, sampai kolom 4
    worksheetcover.conditional_format(22, 0, 11, 3,
                                      {'type': 'no_errors', 'format': borderCover})

    worksheetcover.insert_image('F1', r'logo nf.jpg')

    worksheetcover.merge_range('A10:A11', 'BIDANG STUDI', bodyCover)
    worksheetcover.merge_range('B10:B11', 'TERENDAH', bodyCover)
    worksheetcover.merge_range('C10:C11', 'RATA-RATA', bodyCover)
    worksheetcover.merge_range('D10:D11', 'TERTINGGI', bodyCover)
    worksheetcover.merge_range('A26:A27', 'BIDANG STUDI', bodyCover)
    worksheetcover.merge_range('B26:B27', 'TERENDAH', bodyCover)
    worksheetcover.merge_range('C26:C27', 'RATA-RATA', bodyCover)
    worksheetcover.merge_range('D26:D27', 'TERTINGGI', bodyCover)
    worksheetcover.write('F13', 'BIDANG STUDI', bodyCover)
    worksheetcover.merge_range('F26:F27', 'JUMLAH', sub_header1Cover)
    worksheetcover.merge_range('F29:F30', 'PESERTA', sub_header1Cover)
    worksheetcover.write('G13', 'JUMLAH', bodyCover)
    worksheetcover.set_column('A:A', 25.71, centerCover)
    worksheetcover.set_column('B:B', 15, centerCover)
    worksheetcover.set_column('C:C', 15, centerCover)
    worksheetcover.set_column('D:D', 15, centerCover)
    worksheetcover.set_column('F:F', 25.71, centerCover)
    worksheetcover.set_column('G:G', 13, centerCover)
    worksheetcover.merge_range('A1:F3', 'DAFTAR NILAI', titleCover)
    worksheetcover.merge_range('A4:F5', fr'{penilaian}', sub_titleCover)
    worksheetcover.merge_range(
        'A6:F7', fr'{semester} TAHUN {tahun} ({kurikulum})', headerCover)
    worksheetcover.write('A9', 'JUMLAH BENAR', sub_headerCover)
    worksheetcover.write('A25', 'NILAI STANDAR', sub_headerCover)
    worksheetcover.merge_range('F8:G9', fr'{kelas}', kelasCover)
    worksheetcover.merge_range('F11:G12', 'JUMLAH SOAL', sub_header1Cover)

    # sampai baris 39, dari kolom 1, mulai dari baris 26, sampai kolom 4
    # nilai standar
    worksheetcover.conditional_format(38, 0, 27, 3,
                                      {'type': 'no_errors', 'format': borderCover})

    # jumlah soal
    worksheetcover.conditional_format(23, 6, 13, 5,
                                      {'type': 'no_errors', 'format': borderCover})

    # value jml peserta
    worksheetcover.conditional_format(27, 5, 27, 5,
                                      {'type': 'no_errors', 'format': borderCover})

    # worksheet 101
    worksheet101.insert_image('A1', r'logo resmi nf.jpg')

    worksheet101.set_column('A:A', 7, center)
    worksheet101.set_column('B:B', 6, center)
    worksheet101.set_column('C:C', 18.14, center)
    worksheet101.set_column('D:D', 25, left)
    worksheet101.set_column('E:E', 13.14, left)
    worksheet101.set_column('F:F', 8.57, center)
    worksheet101.set_column('G:AD', 5, center)
    worksheet101.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF TAMAN MARGASATWA', title)
    worksheet101.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet101.write('A5', 'LOKASI', header)
    worksheet101.write('B5', 'TOTAL', header)
    worksheet101.merge_range('A4:B4', 'RANK', header)
    worksheet101.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet101.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet101.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet101.merge_range('F4:F5', 'KELAS', header)
    worksheet101.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet101.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet101.write('G5', 'MAW', body)
    worksheet101.write('H5', 'MAP', body)
    worksheet101.write('I5', 'IND', body)
    worksheet101.write('J5', 'ENG', body)
    worksheet101.write('K5', 'SEJ', body)
    worksheet101.write('L5', 'GEO', body)
    worksheet101.write('M5', 'EKO', body)
    worksheet101.write('N5', 'SOS', body)
    worksheet101.write('O5', 'FIS', body)
    worksheet101.write('P5', 'KIM', body)
    worksheet101.write('Q5', 'BIO', body)
    worksheet101.write('R5', 'JML', body)
    worksheet101.write('S5', 'MAW', body)
    worksheet101.write('T5', 'MAP', body)
    worksheet101.write('U5', 'IND', body)
    worksheet101.write('V5', 'ENG', body)
    worksheet101.write('W5', 'SEJ', body)
    worksheet101.write('X5', 'GEO', body)
    worksheet101.write('Y5', 'EKO', body)
    worksheet101.write('Z5', 'SOS', body)
    worksheet101.write('AA5', 'FIS', body)
    worksheet101.write('AB5', 'KIM', body)
    worksheet101.write('AC5', 'BIO', body)
    worksheet101.write('AD5', 'JML', body)

    worksheet101.conditional_format(5, 0, row101_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet101.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF TAMAN MARGASATWA', title)
    worksheet101.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet101.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet101.write('A22', 'LOKASI', header)
    worksheet101.write('B22', 'TOTAL', header)
    worksheet101.merge_range('A21:B21', 'RANK', header)
    worksheet101.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet101.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet101.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet101.merge_range('F21:F22', 'KELAS', header)
    worksheet101.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet101.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet101.write('G22', 'MAW', body)
    worksheet101.write('H22', 'MAP', body)
    worksheet101.write('I22', 'IND', body)
    worksheet101.write('J22', 'ENG', body)
    worksheet101.write('K22', 'SEJ', body)
    worksheet101.write('L22', 'GEO', body)
    worksheet101.write('M22', 'EKO', body)
    worksheet101.write('N22', 'SOS', body)
    worksheet101.write('O22', 'FIS', body)
    worksheet101.write('P22', 'KIM', body)
    worksheet101.write('Q22', 'BIO', body)
    worksheet101.write('R22', 'JML', body)
    worksheet101.write('S22', 'MAW', body)
    worksheet101.write('T22', 'MAP', body)
    worksheet101.write('U22', 'IND', body)
    worksheet101.write('V22', 'ENG', body)
    worksheet101.write('W22', 'SEJ', body)
    worksheet101.write('X22', 'GEO', body)
    worksheet101.write('Y22', 'EKO', body)
    worksheet101.write('Z22', 'SOS', body)
    worksheet101.write('AA22', 'FIS', body)
    worksheet101.write('AB22', 'KIM', body)
    worksheet101.write('AC22', 'BIO', body)
    worksheet101.write('AD22', 'JML', body)

    worksheet101.conditional_format(22, 0, row101+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 102
    worksheet102.insert_image('A1', r'logo resmi nf.jpg')

    worksheet102.set_column('A:A', 7, center)
    worksheet102.set_column('B:B', 6, center)
    worksheet102.set_column('C:C', 18.14, center)
    worksheet102.set_column('D:D', 25, left)
    worksheet102.set_column('E:E', 13.14, left)
    worksheet102.set_column('F:F', 8.57, center)
    worksheet102.set_column('G:AD', 5, center)
    worksheet102.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CEMPAKA', title)
    worksheet102.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet102.write('A5', 'LOKASI', header)
    worksheet102.write('B5', 'TOTAL', header)
    worksheet102.merge_range('A4:B4', 'RANK', header)
    worksheet102.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet102.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet102.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet102.merge_range('F4:F5', 'KELAS', header)
    worksheet102.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet102.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet102.write('G5', 'MAW', body)
    worksheet102.write('H5', 'MAP', body)
    worksheet102.write('I5', 'IND', body)
    worksheet102.write('J5', 'ENG', body)
    worksheet102.write('K5', 'SEJ', body)
    worksheet102.write('L5', 'GEO', body)
    worksheet102.write('M5', 'EKO', body)
    worksheet102.write('N5', 'SOS', body)
    worksheet102.write('O5', 'FIS', body)
    worksheet102.write('P5', 'KIM', body)
    worksheet102.write('Q5', 'BIO', body)
    worksheet102.write('R5', 'JML', body)
    worksheet102.write('S5', 'MAW', body)
    worksheet102.write('T5', 'MAP', body)
    worksheet102.write('U5', 'IND', body)
    worksheet102.write('V5', 'ENG', body)
    worksheet102.write('W5', 'SEJ', body)
    worksheet102.write('X5', 'GEO', body)
    worksheet102.write('Y5', 'EKO', body)
    worksheet102.write('Z5', 'SOS', body)
    worksheet102.write('AA5', 'FIS', body)
    worksheet102.write('AB5', 'KIM', body)
    worksheet102.write('AC5', 'BIO', body)
    worksheet102.write('AD5', 'JML', body)

    worksheet102.conditional_format(5, 0, row102_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet102.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF CEMPAKA', title)
    worksheet102.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet102.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet102.write('A22', 'LOKASI', header)
    worksheet102.write('B22', 'TOTAL', header)
    worksheet102.merge_range('A21:B21', 'RANK', header)
    worksheet102.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet102.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet102.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet102.merge_range('F21:F22', 'KELAS', header)
    worksheet102.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet102.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet102.write('G22', 'MAW', body)
    worksheet102.write('H22', 'MAP', body)
    worksheet102.write('I22', 'IND', body)
    worksheet102.write('J22', 'ENG', body)
    worksheet102.write('K22', 'SEJ', body)
    worksheet102.write('L22', 'GEO', body)
    worksheet102.write('M22', 'EKO', body)
    worksheet102.write('N22', 'SOS', body)
    worksheet102.write('O22', 'FIS', body)
    worksheet102.write('P22', 'KIM', body)
    worksheet102.write('Q22', 'BIO', body)
    worksheet102.write('R22', 'JML', body)
    worksheet102.write('S22', 'MAW', body)
    worksheet102.write('T22', 'MAP', body)
    worksheet102.write('U22', 'IND', body)
    worksheet102.write('V22', 'ENG', body)
    worksheet102.write('W22', 'SEJ', body)
    worksheet102.write('X22', 'GEO', body)
    worksheet102.write('Y22', 'EKO', body)
    worksheet102.write('Z22', 'SOS', body)
    worksheet102.write('AA22', 'FIS', body)
    worksheet102.write('AB22', 'KIM', body)
    worksheet102.write('AC22', 'BIO', body)
    worksheet102.write('AD22', 'JML', body)

    worksheet102.conditional_format(22, 0, row102+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 103
    worksheet103.insert_image('A1', r'logo resmi nf.jpg')

    worksheet103.set_column('A:A', 7, center)
    worksheet103.set_column('B:B', 6, center)
    worksheet103.set_column('C:C', 18.14, center)
    worksheet103.set_column('D:D', 25, left)
    worksheet103.set_column('E:E', 13.14, left)
    worksheet103.set_column('F:F', 8.57, center)
    worksheet103.set_column('G:AD', 5, center)
    worksheet103.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PANGKALAN JATI', title)
    worksheet103.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet103.write('A5', 'LOKASI', header)
    worksheet103.write('B5', 'TOTAL', header)
    worksheet103.merge_range('A4:B4', 'RANK', header)
    worksheet103.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet103.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet103.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet103.merge_range('F4:F5', 'KELAS', header)
    worksheet103.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet103.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet103.write('G5', 'MAW', body)
    worksheet103.write('H5', 'MAP', body)
    worksheet103.write('I5', 'IND', body)
    worksheet103.write('J5', 'ENG', body)
    worksheet103.write('K5', 'SEJ', body)
    worksheet103.write('L5', 'GEO', body)
    worksheet103.write('M5', 'EKO', body)
    worksheet103.write('N5', 'SOS', body)
    worksheet103.write('O5', 'FIS', body)
    worksheet103.write('P5', 'KIM', body)
    worksheet103.write('Q5', 'BIO', body)
    worksheet103.write('R5', 'JML', body)
    worksheet103.write('S5', 'MAW', body)
    worksheet103.write('T5', 'MAP', body)
    worksheet103.write('U5', 'IND', body)
    worksheet103.write('V5', 'ENG', body)
    worksheet103.write('W5', 'SEJ', body)
    worksheet103.write('X5', 'GEO', body)
    worksheet103.write('Y5', 'EKO', body)
    worksheet103.write('Z5', 'SOS', body)
    worksheet103.write('AA5', 'FIS', body)
    worksheet103.write('AB5', 'KIM', body)
    worksheet103.write('AC5', 'BIO', body)
    worksheet103.write('AD5', 'JML', body)

    worksheet103.conditional_format(5, 0, row103_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet103.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF PANGKALAN JATI', title)
    worksheet103.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet103.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet103.write('A22', 'LOKASI', header)
    worksheet103.write('B22', 'TOTAL', header)
    worksheet103.merge_range('A21:B21', 'RANK', header)
    worksheet103.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet103.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet103.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet103.merge_range('F21:F22', 'KELAS', header)
    worksheet103.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet103.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet103.write('G22', 'MAW', body)
    worksheet103.write('H22', 'MAP', body)
    worksheet103.write('I22', 'IND', body)
    worksheet103.write('J22', 'ENG', body)
    worksheet103.write('K22', 'SEJ', body)
    worksheet103.write('L22', 'GEO', body)
    worksheet103.write('M22', 'EKO', body)
    worksheet103.write('N22', 'SOS', body)
    worksheet103.write('O22', 'FIS', body)
    worksheet103.write('P22', 'KIM', body)
    worksheet103.write('Q22', 'BIO', body)
    worksheet103.write('R22', 'JML', body)
    worksheet103.write('S22', 'MAW', body)
    worksheet103.write('T22', 'MAP', body)
    worksheet103.write('U22', 'IND', body)
    worksheet103.write('V22', 'ENG', body)
    worksheet103.write('W22', 'SEJ', body)
    worksheet103.write('X22', 'GEO', body)
    worksheet103.write('Y22', 'EKO', body)
    worksheet103.write('Z22', 'SOS', body)
    worksheet103.write('AA22', 'FIS', body)
    worksheet103.write('AB22', 'KIM', body)
    worksheet103.write('AC22', 'BIO', body)
    worksheet103.write('AD22', 'JML', body)

    worksheet103.conditional_format(22, 0, row103+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # # worksheet 104
    # worksheet104.insert_image('A1',r'logo resmi nf.jpg')

    # worksheet104.set_column('A:A', 7, center)
    # worksheet104.set_column('B:B', 6, center)
    # worksheet104.set_column('C:C', 18.14, center)
    # worksheet104.set_column('D:D', 25, left)
    # worksheet104.set_column('E:E', 13.14, left)
    # worksheet104.set_column('F:F', 8.57, center)
    # worksheet104.set_column('G:AD', 5, center)
    # worksheet104.merge_range('A1:V1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KENARI', title)
    # worksheet104.merge_range('A2:V2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    # worksheet104.write('A5', 'LOKASI', header)
    # worksheet104.write('B5', 'TOTAL', header)
    # worksheet104.merge_range('A4:B4', 'RANK', header)
    # worksheet104.merge_range('C4:C5', 'NOMOR NF', header)
    # worksheet104.merge_range('D4:D5', 'NAMA SISWA', header)
    # worksheet104.merge_range('E4:E5', 'SEKOLAH', header)
    # worksheet104.merge_range('F4:F5', 'KELAS', header)
    # worksheet104.merge_range('G4:R4', 'JUMLAH BENAR', header)
    # worksheet104.merge_range('S4:AD4', 'NILAI STANDAR', header)
    # worksheet104.write('G5', 'MAW', body)
    # worksheet104.write('H5', 'MAP', body)
    # worksheet104.write('I5', 'IND', body)
    # worksheet104.write('J5', 'ENG', body)
    # worksheet104.write('K5', 'SEJ', body)
    # worksheet104.write('L5', 'GEO', body)
    # worksheet104.write('M5', 'EKO', body)
    # worksheet104.write('N5', 'SOS', body)
    # worksheet104.write('O5', 'FIS', body)
    # worksheet104.write('P5', 'KIM', body)
    # worksheet104.write('Q5', 'BIO', body)
    # worksheet104.write('R5', 'JML', body)
    # worksheet104.write('S5', 'MAW', body)
    # worksheet104.write('T5', 'MAP', body)
    # worksheet104.write('U5', 'IND', body)
    # worksheet104.write('V5', 'ENG', body)
    # worksheet104.write('W5', 'SEJ', body)
    # worksheet104.write('X5', 'GEO', body)
    # worksheet104.write('Y5', 'EKO', body)
    # worksheet104.write('Z5', 'SOS', body)
    # worksheet104.write('AA5', 'FIS', body)
    # worksheet104.write('AB5', 'KIM', body)
    # worksheet104.write('AC5', 'BIO', body)
    # worksheet104.write('AD5', 'JML', body)

    # worksheet104.conditional_format(5,0,row104_10+4,21,
    #                              {'type': 'no_errors', 'format': border})

    # worksheet104.merge_range('A17:V17', fr'KELAS {kelas} - LOKASI NF KENARI', title)
    # worksheet104.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    # worksheet104.merge_range('A19:V19', fr'{semester} TAHUN {tahun}', sub_title)
    # worksheet104.write('A22', 'LOKASI', header)
    # worksheet104.write('B22', 'TOTAL', header)
    # worksheet104.merge_range('A21:B21', 'RANK', header)
    # worksheet104.merge_range('C21:C22', 'NOMOR NF', header)
    # worksheet104.merge_range('D21:D22', 'NAMA SISWA', header)
    # worksheet104.merge_range('E21:E22', 'SEKOLAH', header)
    # worksheet104.merge_range('F21:F22', 'KELAS', header)
    # worksheet104.merge_range('G21:N21', 'JUMLAH BENAR', header)
    # worksheet104.merge_range('O21:V21', 'NILAI STANDAR', header)
    # worksheet104.write('G22', 'MAT', body)
    # worksheet104.write('H22', 'IND', body)
    # worksheet104.write('I22', 'ENG', body)
    # worksheet104.write('J22', 'SEJ', body)
    # worksheet104.write('K22', 'GEO', body)
    # worksheet104.write('L22', 'SOS', body)
    # worksheet104.write('M22', 'EKO', body)
    # worksheet104.write('N22', 'JML', body)
    # worksheet104.write('O22', 'MAT', body)
    # worksheet104.write('P22', 'IND', body)
    # worksheet104.write('Q22', 'ENG', body)
    # worksheet104.write('R22', 'SEJ', body)
    # worksheet104.write('S22', 'GEO', body)
    # worksheet104.write('T22', 'SOS', body)
    # worksheet104.write('U22', 'EKO', body)
    # worksheet104.write('V22', 'JML', body)

    # worksheet104.conditional_format(22,0,row104+21,21,
    #                              {'type': 'no_errors', 'format': border})

    # worksheet 105
    worksheet105.insert_image('A1', r'logo resmi nf.jpg')

    worksheet105.set_column('A:A', 7, center)
    worksheet105.set_column('B:B', 6, center)
    worksheet105.set_column('C:C', 18.14, center)
    worksheet105.set_column('D:D', 25, left)
    worksheet105.set_column('E:E', 13.14, left)
    worksheet105.set_column('F:F', 8.57, center)
    worksheet105.set_column('G:AD', 5, center)
    worksheet105.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BUARAN', title)
    worksheet105.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet105.write('A5', 'LOKASI', header)
    worksheet105.write('B5', 'TOTAL', header)
    worksheet105.merge_range('A4:B4', 'RANK', header)
    worksheet105.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet105.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet105.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet105.merge_range('F4:F5', 'KELAS', header)
    worksheet105.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet105.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet105.write('G5', 'MAW', body)
    worksheet105.write('H5', 'MAP', body)
    worksheet105.write('I5', 'IND', body)
    worksheet105.write('J5', 'ENG', body)
    worksheet105.write('K5', 'SEJ', body)
    worksheet105.write('L5', 'GEO', body)
    worksheet105.write('M5', 'EKO', body)
    worksheet105.write('N5', 'SOS', body)
    worksheet105.write('O5', 'FIS', body)
    worksheet105.write('P5', 'KIM', body)
    worksheet105.write('Q5', 'BIO', body)
    worksheet105.write('R5', 'JML', body)
    worksheet105.write('S5', 'MAW', body)
    worksheet105.write('T5', 'MAP', body)
    worksheet105.write('U5', 'IND', body)
    worksheet105.write('V5', 'ENG', body)
    worksheet105.write('W5', 'SEJ', body)
    worksheet105.write('X5', 'GEO', body)
    worksheet105.write('Y5', 'EKO', body)
    worksheet105.write('Z5', 'SOS', body)
    worksheet105.write('AA5', 'FIS', body)
    worksheet105.write('AB5', 'KIM', body)
    worksheet105.write('AC5', 'BIO', body)
    worksheet105.write('AD5', 'JML', body)

    worksheet105.conditional_format(5, 0, row105_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet105.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF BUARAN', title)
    worksheet105.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet105.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet105.write('A22', 'LOKASI', header)
    worksheet105.write('B22', 'TOTAL', header)
    worksheet105.merge_range('A21:B21', 'RANK', header)
    worksheet105.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet105.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet105.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet105.merge_range('F21:F22', 'KELAS', header)
    worksheet105.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet105.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet105.write('G22', 'MAW', body)
    worksheet105.write('H22', 'MAP', body)
    worksheet105.write('I22', 'IND', body)
    worksheet105.write('J22', 'ENG', body)
    worksheet105.write('K22', 'SEJ', body)
    worksheet105.write('L22', 'GEO', body)
    worksheet105.write('M22', 'EKO', body)
    worksheet105.write('N22', 'SOS', body)
    worksheet105.write('O22', 'FIS', body)
    worksheet105.write('P22', 'KIM', body)
    worksheet105.write('Q22', 'BIO', body)
    worksheet105.write('R22', 'JML', body)
    worksheet105.write('S22', 'MAW', body)
    worksheet105.write('T22', 'MAP', body)
    worksheet105.write('U22', 'IND', body)
    worksheet105.write('V22', 'ENG', body)
    worksheet105.write('W22', 'SEJ', body)
    worksheet105.write('X22', 'GEO', body)
    worksheet105.write('Y22', 'EKO', body)
    worksheet105.write('Z22', 'SOS', body)
    worksheet105.write('AA22', 'FIS', body)
    worksheet105.write('AB22', 'KIM', body)
    worksheet105.write('AC22', 'BIO', body)
    worksheet105.write('AD22', 'JML', body)

    worksheet105.conditional_format(22, 0, row105+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 106
    worksheet106.insert_image('A1', r'logo resmi nf.jpg')

    worksheet106.set_column('A:A', 7, center)
    worksheet106.set_column('B:B', 6, center)
    worksheet106.set_column('C:C', 18.14, center)
    worksheet106.set_column('D:D', 25, left)
    worksheet106.set_column('E:E', 13.14, left)
    worksheet106.set_column('F:F', 8.57, center)
    worksheet106.set_column('G:AD', 5, center)
    worksheet106.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF HEK-KRAMAT JATI', title)
    worksheet106.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet106.write('A5', 'LOKASI', header)
    worksheet106.write('B5', 'TOTAL', header)
    worksheet106.merge_range('A4:B4', 'RANK', header)
    worksheet106.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet106.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet106.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet106.merge_range('F4:F5', 'KELAS', header)
    worksheet106.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet106.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet106.write('G5', 'MAW', body)
    worksheet106.write('H5', 'MAP', body)
    worksheet106.write('I5', 'IND', body)
    worksheet106.write('J5', 'ENG', body)
    worksheet106.write('K5', 'SEJ', body)
    worksheet106.write('L5', 'GEO', body)
    worksheet106.write('M5', 'EKO', body)
    worksheet106.write('N5', 'SOS', body)
    worksheet106.write('O5', 'FIS', body)
    worksheet106.write('P5', 'KIM', body)
    worksheet106.write('Q5', 'BIO', body)
    worksheet106.write('R5', 'JML', body)
    worksheet106.write('S5', 'MAW', body)
    worksheet106.write('T5', 'MAP', body)
    worksheet106.write('U5', 'IND', body)
    worksheet106.write('V5', 'ENG', body)
    worksheet106.write('W5', 'SEJ', body)
    worksheet106.write('X5', 'GEO', body)
    worksheet106.write('Y5', 'EKO', body)
    worksheet106.write('Z5', 'SOS', body)
    worksheet106.write('AA5', 'FIS', body)
    worksheet106.write('AB5', 'KIM', body)
    worksheet106.write('AC5', 'BIO', body)
    worksheet106.write('AD5', 'JML', body)

    worksheet106.conditional_format(5, 0, row106_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet106.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF HEK-KRAMAT JATI', title)
    worksheet106.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet106.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet106.write('A22', 'LOKASI', header)
    worksheet106.write('B22', 'TOTAL', header)
    worksheet106.merge_range('A21:B21', 'RANK', header)
    worksheet106.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet106.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet106.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet106.merge_range('F21:F22', 'KELAS', header)
    worksheet106.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet106.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet106.write('G22', 'MAW', body)
    worksheet106.write('H22', 'MAP', body)
    worksheet106.write('I22', 'IND', body)
    worksheet106.write('J22', 'ENG', body)
    worksheet106.write('K22', 'SEJ', body)
    worksheet106.write('L22', 'GEO', body)
    worksheet106.write('M22', 'EKO', body)
    worksheet106.write('N22', 'SOS', body)
    worksheet106.write('O22', 'FIS', body)
    worksheet106.write('P22', 'KIM', body)
    worksheet106.write('Q22', 'BIO', body)
    worksheet106.write('R22', 'JML', body)
    worksheet106.write('S22', 'MAW', body)
    worksheet106.write('T22', 'MAP', body)
    worksheet106.write('U22', 'IND', body)
    worksheet106.write('V22', 'ENG', body)
    worksheet106.write('W22', 'SEJ', body)
    worksheet106.write('X22', 'GEO', body)
    worksheet106.write('Y22', 'EKO', body)
    worksheet106.write('Z22', 'SOS', body)
    worksheet106.write('AA22', 'FIS', body)
    worksheet106.write('AB22', 'KIM', body)
    worksheet106.write('AC22', 'BIO', body)
    worksheet106.write('AD22', 'JML', body)

    worksheet106.conditional_format(22, 0, row106+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 107
    worksheet107.insert_image('A1', r'logo resmi nf.jpg')

    worksheet107.set_column('A:A', 7, center)
    worksheet107.set_column('B:B', 6, center)
    worksheet107.set_column('C:C', 18.14, center)
    worksheet107.set_column('D:D', 25, left)
    worksheet107.set_column('E:E', 13.14, left)
    worksheet107.set_column('F:F', 8.57, center)
    worksheet107.set_column('G:AD', 5, center)
    worksheet107.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MAMPANG', title)
    worksheet107.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet107.write('A5', 'LOKASI', header)
    worksheet107.write('B5', 'TOTAL', header)
    worksheet107.merge_range('A4:B4', 'RANK', header)
    worksheet107.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet107.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet107.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet107.merge_range('F4:F5', 'KELAS', header)
    worksheet107.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet107.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet107.write('G5', 'MAW', body)
    worksheet107.write('H5', 'MAP', body)
    worksheet107.write('I5', 'IND', body)
    worksheet107.write('J5', 'ENG', body)
    worksheet107.write('K5', 'SEJ', body)
    worksheet107.write('L5', 'GEO', body)
    worksheet107.write('M5', 'EKO', body)
    worksheet107.write('N5', 'SOS', body)
    worksheet107.write('O5', 'FIS', body)
    worksheet107.write('P5', 'KIM', body)
    worksheet107.write('Q5', 'BIO', body)
    worksheet107.write('R5', 'JML', body)
    worksheet107.write('S5', 'MAW', body)
    worksheet107.write('T5', 'MAP', body)
    worksheet107.write('U5', 'IND', body)
    worksheet107.write('V5', 'ENG', body)
    worksheet107.write('W5', 'SEJ', body)
    worksheet107.write('X5', 'GEO', body)
    worksheet107.write('Y5', 'EKO', body)
    worksheet107.write('Z5', 'SOS', body)
    worksheet107.write('AA5', 'FIS', body)
    worksheet107.write('AB5', 'KIM', body)
    worksheet107.write('AC5', 'BIO', body)
    worksheet107.write('AD5', 'JML', body)

    worksheet107.conditional_format(5, 0, row107_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet107.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF MAMPANG', title)
    worksheet107.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet107.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet107.write('A22', 'LOKASI', header)
    worksheet107.write('B22', 'TOTAL', header)
    worksheet107.merge_range('A21:B21', 'RANK', header)
    worksheet107.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet107.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet107.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet107.merge_range('F21:F22', 'KELAS', header)
    worksheet107.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet107.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet107.write('G22', 'MAW', body)
    worksheet107.write('H22', 'MAP', body)
    worksheet107.write('I22', 'IND', body)
    worksheet107.write('J22', 'ENG', body)
    worksheet107.write('K22', 'SEJ', body)
    worksheet107.write('L22', 'GEO', body)
    worksheet107.write('M22', 'EKO', body)
    worksheet107.write('N22', 'SOS', body)
    worksheet107.write('O22', 'FIS', body)
    worksheet107.write('P22', 'KIM', body)
    worksheet107.write('Q22', 'BIO', body)
    worksheet107.write('R22', 'JML', body)
    worksheet107.write('S22', 'MAW', body)
    worksheet107.write('T22', 'MAP', body)
    worksheet107.write('U22', 'IND', body)
    worksheet107.write('V22', 'ENG', body)
    worksheet107.write('W22', 'SEJ', body)
    worksheet107.write('X22', 'GEO', body)
    worksheet107.write('Y22', 'EKO', body)
    worksheet107.write('Z22', 'SOS', body)
    worksheet107.write('AA22', 'FIS', body)
    worksheet107.write('AB22', 'KIM', body)
    worksheet107.write('AC22', 'BIO', body)
    worksheet107.write('AD22', 'JML', body)

    worksheet107.conditional_format(22, 0, row107+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 108
    worksheet108.insert_image('A1', r'logo resmi nf.jpg')

    worksheet108.set_column('A:A', 7, center)
    worksheet108.set_column('B:B', 6, center)
    worksheet108.set_column('C:C', 18.14, center)
    worksheet108.set_column('D:D', 25, left)
    worksheet108.set_column('E:E', 13.14, left)
    worksheet108.set_column('F:F', 8.57, center)
    worksheet108.set_column('G:AD', 5, center)
    worksheet108.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PALMERAH', title)
    worksheet108.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet108.write('A5', 'LOKASI', header)
    worksheet108.write('B5', 'TOTAL', header)
    worksheet108.merge_range('A4:B4', 'RANK', header)
    worksheet108.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet108.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet108.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet108.merge_range('F4:F5', 'KELAS', header)
    worksheet108.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet108.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet108.write('G5', 'MAW', body)
    worksheet108.write('H5', 'MAP', body)
    worksheet108.write('I5', 'IND', body)
    worksheet108.write('J5', 'ENG', body)
    worksheet108.write('K5', 'SEJ', body)
    worksheet108.write('L5', 'GEO', body)
    worksheet108.write('M5', 'EKO', body)
    worksheet108.write('N5', 'SOS', body)
    worksheet108.write('O5', 'FIS', body)
    worksheet108.write('P5', 'KIM', body)
    worksheet108.write('Q5', 'BIO', body)
    worksheet108.write('R5', 'JML', body)
    worksheet108.write('S5', 'MAW', body)
    worksheet108.write('T5', 'MAP', body)
    worksheet108.write('U5', 'IND', body)
    worksheet108.write('V5', 'ENG', body)
    worksheet108.write('W5', 'SEJ', body)
    worksheet108.write('X5', 'GEO', body)
    worksheet108.write('Y5', 'EKO', body)
    worksheet108.write('Z5', 'SOS', body)
    worksheet108.write('AA5', 'FIS', body)
    worksheet108.write('AB5', 'KIM', body)
    worksheet108.write('AC5', 'BIO', body)
    worksheet108.write('AD5', 'JML', body)

    worksheet108.conditional_format(5, 0, row108_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet108.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF PALMERAH', title)
    worksheet108.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet108.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet108.write('A22', 'LOKASI', header)
    worksheet108.write('B22', 'TOTAL', header)
    worksheet108.merge_range('A21:B21', 'RANK', header)
    worksheet108.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet108.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet108.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet108.merge_range('F21:F22', 'KELAS', header)
    worksheet108.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet108.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet108.write('G22', 'MAW', body)
    worksheet108.write('H22', 'MAP', body)
    worksheet108.write('I22', 'IND', body)
    worksheet108.write('J22', 'ENG', body)
    worksheet108.write('K22', 'SEJ', body)
    worksheet108.write('L22', 'GEO', body)
    worksheet108.write('M22', 'EKO', body)
    worksheet108.write('N22', 'SOS', body)
    worksheet108.write('O22', 'FIS', body)
    worksheet108.write('P22', 'KIM', body)
    worksheet108.write('Q22', 'BIO', body)
    worksheet108.write('R22', 'JML', body)
    worksheet108.write('S22', 'MAW', body)
    worksheet108.write('T22', 'MAP', body)
    worksheet108.write('U22', 'IND', body)
    worksheet108.write('V22', 'ENG', body)
    worksheet108.write('W22', 'SEJ', body)
    worksheet108.write('X22', 'GEO', body)
    worksheet108.write('Y22', 'EKO', body)
    worksheet108.write('Z22', 'SOS', body)
    worksheet108.write('AA22', 'FIS', body)
    worksheet108.write('AB22', 'KIM', body)
    worksheet108.write('AC22', 'BIO', body)
    worksheet108.write('AD22', 'JML', body)

    worksheet108.conditional_format(22, 0, row108+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 109
    worksheet109.insert_image('A1', r'logo resmi nf.jpg')

    worksheet109.set_column('A:A', 7, center)
    worksheet109.set_column('B:B', 6, center)
    worksheet109.set_column('C:C', 18.14, center)
    worksheet109.set_column('D:D', 25, left)
    worksheet109.set_column('E:E', 13.14, left)
    worksheet109.set_column('F:F', 8.57, center)
    worksheet109.set_column('G:AD', 5, center)
    worksheet109.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PASAR MINGGU', title)
    worksheet109.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet109.write('A5', 'LOKASI', header)
    worksheet109.write('B5', 'TOTAL', header)
    worksheet109.merge_range('A4:B4', 'RANK', header)
    worksheet109.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet109.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet109.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet109.merge_range('F4:F5', 'KELAS', header)
    worksheet109.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet109.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet109.write('G5', 'MAW', body)
    worksheet109.write('H5', 'MAP', body)
    worksheet109.write('I5', 'IND', body)
    worksheet109.write('J5', 'ENG', body)
    worksheet109.write('K5', 'SEJ', body)
    worksheet109.write('L5', 'GEO', body)
    worksheet109.write('M5', 'EKO', body)
    worksheet109.write('N5', 'SOS', body)
    worksheet109.write('O5', 'FIS', body)
    worksheet109.write('P5', 'KIM', body)
    worksheet109.write('Q5', 'BIO', body)
    worksheet109.write('R5', 'JML', body)
    worksheet109.write('S5', 'MAW', body)
    worksheet109.write('T5', 'MAP', body)
    worksheet109.write('U5', 'IND', body)
    worksheet109.write('V5', 'ENG', body)
    worksheet109.write('W5', 'SEJ', body)
    worksheet109.write('X5', 'GEO', body)
    worksheet109.write('Y5', 'EKO', body)
    worksheet109.write('Z5', 'SOS', body)
    worksheet109.write('AA5', 'FIS', body)
    worksheet109.write('AB5', 'KIM', body)
    worksheet109.write('AC5', 'BIO', body)
    worksheet109.write('AD5', 'JML', body)

    worksheet109.conditional_format(5, 0, row109_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet109.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF PASAR MINGGU', title)
    worksheet109.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet109.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet109.write('A22', 'LOKASI', header)
    worksheet109.write('B22', 'TOTAL', header)
    worksheet109.merge_range('A21:B21', 'RANK', header)
    worksheet109.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet109.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet109.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet109.merge_range('F21:F22', 'KELAS', header)
    worksheet109.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet109.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet109.write('G22', 'MAW', body)
    worksheet109.write('H22', 'MAP', body)
    worksheet109.write('I22', 'IND', body)
    worksheet109.write('J22', 'ENG', body)
    worksheet109.write('K22', 'SEJ', body)
    worksheet109.write('L22', 'GEO', body)
    worksheet109.write('M22', 'EKO', body)
    worksheet109.write('N22', 'SOS', body)
    worksheet109.write('O22', 'FIS', body)
    worksheet109.write('P22', 'KIM', body)
    worksheet109.write('Q22', 'BIO', body)
    worksheet109.write('R22', 'JML', body)
    worksheet109.write('S22', 'MAW', body)
    worksheet109.write('T22', 'MAP', body)
    worksheet109.write('U22', 'IND', body)
    worksheet109.write('V22', 'ENG', body)
    worksheet109.write('W22', 'SEJ', body)
    worksheet109.write('X22', 'GEO', body)
    worksheet109.write('Y22', 'EKO', body)
    worksheet109.write('Z22', 'SOS', body)
    worksheet109.write('AA22', 'FIS', body)
    worksheet109.write('AB22', 'KIM', body)
    worksheet109.write('AC22', 'BIO', body)
    worksheet109.write('AD22', 'JML', body)

    worksheet109.conditional_format(22, 0, row109+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 110
    worksheet110.insert_image('A1', r'logo resmi nf.jpg')

    worksheet110.set_column('A:A', 7, center)
    worksheet110.set_column('B:B', 6, center)
    worksheet110.set_column('C:C', 18.14, center)
    worksheet110.set_column('D:D', 25, left)
    worksheet110.set_column('E:E', 13.14, left)
    worksheet110.set_column('F:F', 8.57, center)
    worksheet110.set_column('G:AD', 5, center)
    worksheet110.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BINTARO', title)
    worksheet110.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet110.write('A5', 'LOKASI', header)
    worksheet110.write('B5', 'TOTAL', header)
    worksheet110.merge_range('A4:B4', 'RANK', header)
    worksheet110.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet110.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet110.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet110.merge_range('F4:F5', 'KELAS', header)
    worksheet110.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet110.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet110.write('G5', 'MAW', body)
    worksheet110.write('H5', 'MAP', body)
    worksheet110.write('I5', 'IND', body)
    worksheet110.write('J5', 'ENG', body)
    worksheet110.write('K5', 'SEJ', body)
    worksheet110.write('L5', 'GEO', body)
    worksheet110.write('M5', 'EKO', body)
    worksheet110.write('N5', 'SOS', body)
    worksheet110.write('O5', 'FIS', body)
    worksheet110.write('P5', 'KIM', body)
    worksheet110.write('Q5', 'BIO', body)
    worksheet110.write('R5', 'JML', body)
    worksheet110.write('S5', 'MAW', body)
    worksheet110.write('T5', 'MAP', body)
    worksheet110.write('U5', 'IND', body)
    worksheet110.write('V5', 'ENG', body)
    worksheet110.write('W5', 'SEJ', body)
    worksheet110.write('X5', 'GEO', body)
    worksheet110.write('Y5', 'EKO', body)
    worksheet110.write('Z5', 'SOS', body)
    worksheet110.write('AA5', 'FIS', body)
    worksheet110.write('AB5', 'KIM', body)
    worksheet110.write('AC5', 'BIO', body)
    worksheet110.write('AD5', 'JML', body)

    worksheet110.conditional_format(5, 0, row110_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet110.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF BINTARO', title)
    worksheet110.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet110.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet110.write('A22', 'LOKASI', header)
    worksheet110.write('B22', 'TOTAL', header)
    worksheet110.merge_range('A21:B21', 'RANK', header)
    worksheet110.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet110.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet110.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet110.merge_range('F21:F22', 'KELAS', header)
    worksheet110.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet110.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet110.write('G22', 'MAW', body)
    worksheet110.write('H22', 'MAP', body)
    worksheet110.write('I22', 'IND', body)
    worksheet110.write('J22', 'ENG', body)
    worksheet110.write('K22', 'SEJ', body)
    worksheet110.write('L22', 'GEO', body)
    worksheet110.write('M22', 'EKO', body)
    worksheet110.write('N22', 'SOS', body)
    worksheet110.write('O22', 'FIS', body)
    worksheet110.write('P22', 'KIM', body)
    worksheet110.write('Q22', 'BIO', body)
    worksheet110.write('R22', 'JML', body)
    worksheet110.write('S22', 'MAW', body)
    worksheet110.write('T22', 'MAP', body)
    worksheet110.write('U22', 'IND', body)
    worksheet110.write('V22', 'ENG', body)
    worksheet110.write('W22', 'SEJ', body)
    worksheet110.write('X22', 'GEO', body)
    worksheet110.write('Y22', 'EKO', body)
    worksheet110.write('Z22', 'SOS', body)
    worksheet110.write('AA22', 'FIS', body)
    worksheet110.write('AB22', 'KIM', body)
    worksheet110.write('AC22', 'BIO', body)
    worksheet110.write('AD22', 'JML', body)

    worksheet110.conditional_format(22, 0, row110+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 111
    worksheet111.insert_image('A1', r'logo resmi nf.jpg')

    worksheet111.set_column('A:A', 7, center)
    worksheet111.set_column('B:B', 6, center)
    worksheet111.set_column('C:C', 18.14, center)
    worksheet111.set_column('D:D', 25, left)
    worksheet111.set_column('E:E', 13.14, left)
    worksheet111.set_column('F:F', 8.57, center)
    worksheet111.set_column('G:AD', 5, center)
    worksheet111.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF LAMPIRI', title)
    worksheet111.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet111.write('A5', 'LOKASI', header)
    worksheet111.write('B5', 'TOTAL', header)
    worksheet111.merge_range('A4:B4', 'RANK', header)
    worksheet111.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet111.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet111.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet111.merge_range('F4:F5', 'KELAS', header)
    worksheet111.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet111.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet111.write('G5', 'MAW', body)
    worksheet111.write('H5', 'MAP', body)
    worksheet111.write('I5', 'IND', body)
    worksheet111.write('J5', 'ENG', body)
    worksheet111.write('K5', 'SEJ', body)
    worksheet111.write('L5', 'GEO', body)
    worksheet111.write('M5', 'EKO', body)
    worksheet111.write('N5', 'SOS', body)
    worksheet111.write('O5', 'FIS', body)
    worksheet111.write('P5', 'KIM', body)
    worksheet111.write('Q5', 'BIO', body)
    worksheet111.write('R5', 'JML', body)
    worksheet111.write('S5', 'MAW', body)
    worksheet111.write('T5', 'MAP', body)
    worksheet111.write('U5', 'IND', body)
    worksheet111.write('V5', 'ENG', body)
    worksheet111.write('W5', 'SEJ', body)
    worksheet111.write('X5', 'GEO', body)
    worksheet111.write('Y5', 'EKO', body)
    worksheet111.write('Z5', 'SOS', body)
    worksheet111.write('AA5', 'FIS', body)
    worksheet111.write('AB5', 'KIM', body)
    worksheet111.write('AC5', 'BIO', body)
    worksheet111.write('AD5', 'JML', body)

    worksheet111.conditional_format(5, 0, row111_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet111.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF LAMPIRI', title)
    worksheet111.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet111.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet111.write('A22', 'LOKASI', header)
    worksheet111.write('B22', 'TOTAL', header)
    worksheet111.merge_range('A21:B21', 'RANK', header)
    worksheet111.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet111.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet111.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet111.merge_range('F21:F22', 'KELAS', header)
    worksheet111.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet111.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet111.write('G22', 'MAW', body)
    worksheet111.write('H22', 'MAP', body)
    worksheet111.write('I22', 'IND', body)
    worksheet111.write('J22', 'ENG', body)
    worksheet111.write('K22', 'SEJ', body)
    worksheet111.write('L22', 'GEO', body)
    worksheet111.write('M22', 'EKO', body)
    worksheet111.write('N22', 'SOS', body)
    worksheet111.write('O22', 'FIS', body)
    worksheet111.write('P22', 'KIM', body)
    worksheet111.write('Q22', 'BIO', body)
    worksheet111.write('R22', 'JML', body)
    worksheet111.write('S22', 'MAW', body)
    worksheet111.write('T22', 'MAP', body)
    worksheet111.write('U22', 'IND', body)
    worksheet111.write('V22', 'ENG', body)
    worksheet111.write('W22', 'SEJ', body)
    worksheet111.write('X22', 'GEO', body)
    worksheet111.write('Y22', 'EKO', body)
    worksheet111.write('Z22', 'SOS', body)
    worksheet111.write('AA22', 'FIS', body)
    worksheet111.write('AB22', 'KIM', body)
    worksheet111.write('AC22', 'BIO', body)
    worksheet111.write('AD22', 'JML', body)

    worksheet111.conditional_format(22, 0, row111+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 112
    worksheet112.insert_image('A1', r'logo resmi nf.jpg')

    worksheet112.set_column('A:A', 7, center)
    worksheet112.set_column('B:B', 6, center)
    worksheet112.set_column('C:C', 18.14, center)
    worksheet112.set_column('D:D', 25, left)
    worksheet112.set_column('E:E', 13.14, left)
    worksheet112.set_column('F:F', 8.57, center)
    worksheet112.set_column('G:AD', 5, center)
    worksheet112.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PONDOK BAMBU', title)
    worksheet112.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet112.write('A5', 'LOKASI', header)
    worksheet112.write('B5', 'TOTAL', header)
    worksheet112.merge_range('A4:B4', 'RANK', header)
    worksheet112.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet112.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet112.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet112.merge_range('F4:F5', 'KELAS', header)
    worksheet112.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet112.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet112.write('G5', 'MAW', body)
    worksheet112.write('H5', 'MAP', body)
    worksheet112.write('I5', 'IND', body)
    worksheet112.write('J5', 'ENG', body)
    worksheet112.write('K5', 'SEJ', body)
    worksheet112.write('L5', 'GEO', body)
    worksheet112.write('M5', 'EKO', body)
    worksheet112.write('N5', 'SOS', body)
    worksheet112.write('O5', 'FIS', body)
    worksheet112.write('P5', 'KIM', body)
    worksheet112.write('Q5', 'BIO', body)
    worksheet112.write('R5', 'JML', body)
    worksheet112.write('S5', 'MAW', body)
    worksheet112.write('T5', 'MAP', body)
    worksheet112.write('U5', 'IND', body)
    worksheet112.write('V5', 'ENG', body)
    worksheet112.write('W5', 'SEJ', body)
    worksheet112.write('X5', 'GEO', body)
    worksheet112.write('Y5', 'EKO', body)
    worksheet112.write('Z5', 'SOS', body)
    worksheet112.write('AA5', 'FIS', body)
    worksheet112.write('AB5', 'KIM', body)
    worksheet112.write('AC5', 'BIO', body)
    worksheet112.write('AD5', 'JML', body)

    worksheet112.conditional_format(5, 0, row112_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet112.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF PONDOK BAMBU', title)
    worksheet112.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet112.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet112.write('A22', 'LOKASI', header)
    worksheet112.write('B22', 'TOTAL', header)
    worksheet112.merge_range('A21:B21', 'RANK', header)
    worksheet112.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet112.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet112.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet112.merge_range('F21:F22', 'KELAS', header)
    worksheet112.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet112.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet112.write('G22', 'MAW', body)
    worksheet112.write('H22', 'MAP', body)
    worksheet112.write('I22', 'IND', body)
    worksheet112.write('J22', 'ENG', body)
    worksheet112.write('K22', 'SEJ', body)
    worksheet112.write('L22', 'GEO', body)
    worksheet112.write('M22', 'EKO', body)
    worksheet112.write('N22', 'SOS', body)
    worksheet112.write('O22', 'FIS', body)
    worksheet112.write('P22', 'KIM', body)
    worksheet112.write('Q22', 'BIO', body)
    worksheet112.write('R22', 'JML', body)
    worksheet112.write('S22', 'MAW', body)
    worksheet112.write('T22', 'MAP', body)
    worksheet112.write('U22', 'IND', body)
    worksheet112.write('V22', 'ENG', body)
    worksheet112.write('W22', 'SEJ', body)
    worksheet112.write('X22', 'GEO', body)
    worksheet112.write('Y22', 'EKO', body)
    worksheet112.write('Z22', 'SOS', body)
    worksheet112.write('AA22', 'FIS', body)
    worksheet112.write('AB22', 'KIM', body)
    worksheet112.write('AC22', 'BIO', body)
    worksheet112.write('AD22', 'JML', body)

    worksheet112.conditional_format(22, 0, row112+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 113
    worksheet113.insert_image('A1', r'logo resmi nf.jpg')

    worksheet113.set_column('A:A', 7, center)
    worksheet113.set_column('B:B', 6, center)
    worksheet113.set_column('C:C', 18.14, center)
    worksheet113.set_column('D:D', 25, left)
    worksheet113.set_column('E:E', 13.14, left)
    worksheet113.set_column('F:F', 8.57, center)
    worksheet113.set_column('G:AD', 5, center)
    worksheet113.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF RAWA BADAK', title)
    worksheet113.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet113.write('A5', 'LOKASI', header)
    worksheet113.write('B5', 'TOTAL', header)
    worksheet113.merge_range('A4:B4', 'RANK', header)
    worksheet113.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet113.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet113.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet113.merge_range('F4:F5', 'KELAS', header)
    worksheet113.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet113.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet113.write('G5', 'MAW', body)
    worksheet113.write('H5', 'MAP', body)
    worksheet113.write('I5', 'IND', body)
    worksheet113.write('J5', 'ENG', body)
    worksheet113.write('K5', 'SEJ', body)
    worksheet113.write('L5', 'GEO', body)
    worksheet113.write('M5', 'EKO', body)
    worksheet113.write('N5', 'SOS', body)
    worksheet113.write('O5', 'FIS', body)
    worksheet113.write('P5', 'KIM', body)
    worksheet113.write('Q5', 'BIO', body)
    worksheet113.write('R5', 'JML', body)
    worksheet113.write('S5', 'MAW', body)
    worksheet113.write('T5', 'MAP', body)
    worksheet113.write('U5', 'IND', body)
    worksheet113.write('V5', 'ENG', body)
    worksheet113.write('W5', 'SEJ', body)
    worksheet113.write('X5', 'GEO', body)
    worksheet113.write('Y5', 'EKO', body)
    worksheet113.write('Z5', 'SOS', body)
    worksheet113.write('AA5', 'FIS', body)
    worksheet113.write('AB5', 'KIM', body)
    worksheet113.write('AC5', 'BIO', body)
    worksheet113.write('AD5', 'JML', body)

    worksheet113.conditional_format(5, 0, row113_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet113.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF RAWA BADAK', title)
    worksheet113.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet113.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet113.write('A22', 'LOKASI', header)
    worksheet113.write('B22', 'TOTAL', header)
    worksheet113.merge_range('A21:B21', 'RANK', header)
    worksheet113.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet113.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet113.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet113.merge_range('F21:F22', 'KELAS', header)
    worksheet113.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet113.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet113.write('G22', 'MAW', body)
    worksheet113.write('H22', 'MAP', body)
    worksheet113.write('I22', 'IND', body)
    worksheet113.write('J22', 'ENG', body)
    worksheet113.write('K22', 'SEJ', body)
    worksheet113.write('L22', 'GEO', body)
    worksheet113.write('M22', 'EKO', body)
    worksheet113.write('N22', 'SOS', body)
    worksheet113.write('O22', 'FIS', body)
    worksheet113.write('P22', 'KIM', body)
    worksheet113.write('Q22', 'BIO', body)
    worksheet113.write('R22', 'JML', body)
    worksheet113.write('S22', 'MAW', body)
    worksheet113.write('T22', 'MAP', body)
    worksheet113.write('U22', 'IND', body)
    worksheet113.write('V22', 'ENG', body)
    worksheet113.write('W22', 'SEJ', body)
    worksheet113.write('X22', 'GEO', body)
    worksheet113.write('Y22', 'EKO', body)
    worksheet113.write('Z22', 'SOS', body)
    worksheet113.write('AA22', 'FIS', body)
    worksheet113.write('AB22', 'KIM', body)
    worksheet113.write('AC22', 'BIO', body)
    worksheet113.write('AD22', 'JML', body)

    worksheet113.conditional_format(22, 0, row113+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # # worksheet 114
    # worksheet114.insert_image('A1',r'logo resmi nf.jpg')

    # worksheet114.set_column('A:A', 7, center)
    # worksheet114.set_column('B:B', 6, center)
    # worksheet114.set_column('C:C', 18.14, center)
    # worksheet114.set_column('D:D', 25, left)
    # worksheet114.set_column('E:E', 13.14, left)
    # worksheet114.set_column('F:F', 8.57, center)
    # worksheet114.set_column('G:AD', 5, center)
    # worksheet114.merge_range('A1:V1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PASAR REBO', title)
    # worksheet114.merge_range('A2:V2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    # worksheet114.write('A5', 'LOKASI', header)
    # worksheet114.write('B5', 'TOTAL', header)
    # worksheet114.merge_range('A4:B4', 'RANK', header)
    # worksheet114.merge_range('C4:C5', 'NOMOR NF', header)
    # worksheet114.merge_range('D4:D5', 'NAMA SISWA', header)
    # worksheet114.merge_range('E4:E5', 'SEKOLAH', header)
    # worksheet114.merge_range('F4:F5', 'KELAS', header)
    # worksheet114.merge_range('G4:R4', 'JUMLAH BENAR', header)
    # worksheet114.merge_range('S4:AD4', 'NILAI STANDAR', header)
    # worksheet114.write('G5', 'MAW', body)
    # worksheet114.write('H5', 'MAP', body)
    # worksheet114.write('I5', 'IND', body)
    # worksheet114.write('J5', 'ENG', body)
    # worksheet114.write('K5', 'SEJ', body)
    # worksheet114.write('L5', 'GEO', body)
    # worksheet114.write('M5', 'EKO', body)
    # worksheet114.write('N5', 'SOS', body)
    # worksheet114.write('O5', 'FIS', body)
    # worksheet114.write('P5', 'KIM', body)
    # worksheet114.write('Q5', 'BIO', body)
    # worksheet114.write('R5', 'JML', body)
    # worksheet114.write('S5', 'MAW', body)
    # worksheet114.write('T5', 'MAP', body)
    # worksheet114.write('U5', 'IND', body)
    # worksheet114.write('V5', 'ENG', body)
    # worksheet114.write('W5', 'SEJ', body)
    # worksheet114.write('X5', 'GEO', body)
    # worksheet114.write('Y5', 'EKO', body)
    # worksheet114.write('Z5', 'SOS', body)
    # worksheet114.write('AA5', 'FIS', body)
    # worksheet114.write('AB5', 'KIM', body)
    # worksheet114.write('AC5', 'BIO', body)
    # worksheet114.write('AD5', 'JML', body)

    # worksheet114.conditional_format(5,0,row114_10+4,21,
    #                              {'type': 'no_errors', 'format': border})

    # worksheet114.merge_range('A17:V17', fr'KELAS {kelas} - LOKASI NF PASAR REBO', title)
    # worksheet114.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    # worksheet114.merge_range('A19:V19', fr'{semester} TAHUN {tahun}', sub_title)
    # worksheet114.write('A22', 'LOKASI', header)
    # worksheet114.write('B22', 'TOTAL', header)
    # worksheet114.merge_range('A21:B21', 'RANK', header)
    # worksheet114.merge_range('C21:C22', 'NOMOR NF', header)
    # worksheet114.merge_range('D21:D22', 'NAMA SISWA', header)
    # worksheet114.merge_range('E21:E22', 'SEKOLAH', header)
    # worksheet114.merge_range('F21:F22', 'KELAS', header)
    # worksheet114.merge_range('G21:N21', 'JUMLAH BENAR', header)
    # worksheet114.merge_range('O21:V21', 'NILAI STANDAR', header)
    # worksheet114.write('G22', 'MAT', body)
    # worksheet114.write('H22', 'IND', body)
    # worksheet114.write('I22', 'ENG', body)
    # worksheet114.write('J22', 'SEJ', body)
    # worksheet114.write('K22', 'GEO', body)
    # worksheet114.write('L22', 'SOS', body)
    # worksheet114.write('M22', 'EKO', body)
    # worksheet114.write('N22', 'JML', body)
    # worksheet114.write('O22', 'MAT', body)
    # worksheet114.write('P22', 'IND', body)
    # worksheet114.write('Q22', 'ENG', body)
    # worksheet114.write('R22', 'SEJ', body)
    # worksheet114.write('S22', 'GEO', body)
    # worksheet114.write('T22', 'SOS', body)
    # worksheet114.write('U22', 'EKO', body)
    # worksheet114.write('V22', 'JML', body)

    # worksheet114.conditional_format(22,0,row114+21,21,
    #                              {'type': 'no_errors', 'format': border})
    # worksheet 115
    worksheet115.insert_image('A1', r'logo resmi nf.jpg')

    worksheet115.set_column('A:A', 7, center)
    worksheet115.set_column('B:B', 6, center)
    worksheet115.set_column('C:C', 18.14, center)
    worksheet115.set_column('D:D', 25, left)
    worksheet115.set_column('E:E', 13.14, left)
    worksheet115.set_column('F:F', 8.57, center)
    worksheet115.set_column('G:AD', 5, center)
    worksheet115.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF RAWAMANGUN', title)
    worksheet115.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet115.write('A5', 'LOKASI', header)
    worksheet115.write('B5', 'TOTAL', header)
    worksheet115.merge_range('A4:B4', 'RANK', header)
    worksheet115.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet115.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet115.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet115.merge_range('F4:F5', 'KELAS', header)
    worksheet115.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet115.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet115.write('G5', 'MAW', body)
    worksheet115.write('H5', 'MAP', body)
    worksheet115.write('I5', 'IND', body)
    worksheet115.write('J5', 'ENG', body)
    worksheet115.write('K5', 'SEJ', body)
    worksheet115.write('L5', 'GEO', body)
    worksheet115.write('M5', 'EKO', body)
    worksheet115.write('N5', 'SOS', body)
    worksheet115.write('O5', 'FIS', body)
    worksheet115.write('P5', 'KIM', body)
    worksheet115.write('Q5', 'BIO', body)
    worksheet115.write('R5', 'JML', body)
    worksheet115.write('S5', 'MAW', body)
    worksheet115.write('T5', 'MAP', body)
    worksheet115.write('U5', 'IND', body)
    worksheet115.write('V5', 'ENG', body)
    worksheet115.write('W5', 'SEJ', body)
    worksheet115.write('X5', 'GEO', body)
    worksheet115.write('Y5', 'EKO', body)
    worksheet115.write('Z5', 'SOS', body)
    worksheet115.write('AA5', 'FIS', body)
    worksheet115.write('AB5', 'KIM', body)
    worksheet115.write('AC5', 'BIO', body)
    worksheet115.write('AD5', 'JML', body)

    worksheet115.conditional_format(5, 0, row115_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet115.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF RAWAMANGUN', title)
    worksheet115.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet115.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet115.write('A22', 'LOKASI', header)
    worksheet115.write('B22', 'TOTAL', header)
    worksheet115.merge_range('A21:B21', 'RANK', header)
    worksheet115.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet115.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet115.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet115.merge_range('F21:F22', 'KELAS', header)
    worksheet115.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet115.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet115.write('G22', 'MAW', body)
    worksheet115.write('H22', 'MAP', body)
    worksheet115.write('I22', 'IND', body)
    worksheet115.write('J22', 'ENG', body)
    worksheet115.write('K22', 'SEJ', body)
    worksheet115.write('L22', 'GEO', body)
    worksheet115.write('M22', 'EKO', body)
    worksheet115.write('N22', 'SOS', body)
    worksheet115.write('O22', 'FIS', body)
    worksheet115.write('P22', 'KIM', body)
    worksheet115.write('Q22', 'BIO', body)
    worksheet115.write('R22', 'JML', body)
    worksheet115.write('S22', 'MAW', body)
    worksheet115.write('T22', 'MAP', body)
    worksheet115.write('U22', 'IND', body)
    worksheet115.write('V22', 'ENG', body)
    worksheet115.write('W22', 'SEJ', body)
    worksheet115.write('X22', 'GEO', body)
    worksheet115.write('Y22', 'EKO', body)
    worksheet115.write('Z22', 'SOS', body)
    worksheet115.write('AA22', 'FIS', body)
    worksheet115.write('AB22', 'KIM', body)
    worksheet115.write('AC22', 'BIO', body)
    worksheet115.write('AD22', 'JML', body)

    worksheet115.conditional_format(22, 0, row115+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 116
    worksheet116.insert_image('A1', r'logo resmi nf.jpg')

    worksheet116.set_column('A:A', 7, center)
    worksheet116.set_column('B:B', 6, center)
    worksheet116.set_column('C:C', 18.14, center)
    worksheet116.set_column('D:D', 25, left)
    worksheet116.set_column('E:E', 13.14, left)
    worksheet116.set_column('F:F', 8.57, center)
    worksheet116.set_column('G:AD', 5, center)
    worksheet116.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIRACAS', title)
    worksheet116.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet116.write('A5', 'LOKASI', header)
    worksheet116.write('B5', 'TOTAL', header)
    worksheet116.merge_range('A4:B4', 'RANK', header)
    worksheet116.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet116.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet116.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet116.merge_range('F4:F5', 'KELAS', header)
    worksheet116.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet116.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet116.write('G5', 'MAW', body)
    worksheet116.write('H5', 'MAP', body)
    worksheet116.write('I5', 'IND', body)
    worksheet116.write('J5', 'ENG', body)
    worksheet116.write('K5', 'SEJ', body)
    worksheet116.write('L5', 'GEO', body)
    worksheet116.write('M5', 'EKO', body)
    worksheet116.write('N5', 'SOS', body)
    worksheet116.write('O5', 'FIS', body)
    worksheet116.write('P5', 'KIM', body)
    worksheet116.write('Q5', 'BIO', body)
    worksheet116.write('R5', 'JML', body)
    worksheet116.write('S5', 'MAW', body)
    worksheet116.write('T5', 'MAP', body)
    worksheet116.write('U5', 'IND', body)
    worksheet116.write('V5', 'ENG', body)
    worksheet116.write('W5', 'SEJ', body)
    worksheet116.write('X5', 'GEO', body)
    worksheet116.write('Y5', 'EKO', body)
    worksheet116.write('Z5', 'SOS', body)
    worksheet116.write('AA5', 'FIS', body)
    worksheet116.write('AB5', 'KIM', body)
    worksheet116.write('AC5', 'BIO', body)
    worksheet116.write('AD5', 'JML', body)

    worksheet116.conditional_format(5, 0, row116_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet116.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF CIRACAS', title)
    worksheet116.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet116.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet116.write('A22', 'LOKASI', header)
    worksheet116.write('B22', 'TOTAL', header)
    worksheet116.merge_range('A21:B21', 'RANK', header)
    worksheet116.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet116.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet116.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet116.merge_range('F21:F22', 'KELAS', header)
    worksheet116.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet116.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet116.write('G22', 'MAW', body)
    worksheet116.write('H22', 'MAP', body)
    worksheet116.write('I22', 'IND', body)
    worksheet116.write('J22', 'ENG', body)
    worksheet116.write('K22', 'SEJ', body)
    worksheet116.write('L22', 'GEO', body)
    worksheet116.write('M22', 'EKO', body)
    worksheet116.write('N22', 'SOS', body)
    worksheet116.write('O22', 'FIS', body)
    worksheet116.write('P22', 'KIM', body)
    worksheet116.write('Q22', 'BIO', body)
    worksheet116.write('R22', 'JML', body)
    worksheet116.write('S22', 'MAW', body)
    worksheet116.write('T22', 'MAP', body)
    worksheet116.write('U22', 'IND', body)
    worksheet116.write('V22', 'ENG', body)
    worksheet116.write('W22', 'SEJ', body)
    worksheet116.write('X22', 'GEO', body)
    worksheet116.write('Y22', 'EKO', body)
    worksheet116.write('Z22', 'SOS', body)
    worksheet116.write('AA22', 'FIS', body)
    worksheet116.write('AB22', 'KIM', body)
    worksheet116.write('AC22', 'BIO', body)
    worksheet116.write('AD22', 'JML', body)

    worksheet116.conditional_format(22, 0, row116+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 117
    worksheet117.insert_image('A1', r'logo resmi nf.jpg')

    worksheet117.set_column('A:A', 7, center)
    worksheet117.set_column('B:B', 6, center)
    worksheet117.set_column('C:C', 18.14, center)
    worksheet117.set_column('D:D', 25, left)
    worksheet117.set_column('E:E', 13.14, left)
    worksheet117.set_column('F:F', 8.57, center)
    worksheet117.set_column('G:AD', 5, center)
    worksheet117.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KAMPUNG MELAYU', title)
    worksheet117.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet117.write('A5', 'LOKASI', header)
    worksheet117.write('B5', 'TOTAL', header)
    worksheet117.merge_range('A4:B4', 'RANK', header)
    worksheet117.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet117.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet117.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet117.merge_range('F4:F5', 'KELAS', header)
    worksheet117.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet117.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet117.write('G5', 'MAW', body)
    worksheet117.write('H5', 'MAP', body)
    worksheet117.write('I5', 'IND', body)
    worksheet117.write('J5', 'ENG', body)
    worksheet117.write('K5', 'SEJ', body)
    worksheet117.write('L5', 'GEO', body)
    worksheet117.write('M5', 'EKO', body)
    worksheet117.write('N5', 'SOS', body)
    worksheet117.write('O5', 'FIS', body)
    worksheet117.write('P5', 'KIM', body)
    worksheet117.write('Q5', 'BIO', body)
    worksheet117.write('R5', 'JML', body)
    worksheet117.write('S5', 'MAW', body)
    worksheet117.write('T5', 'MAP', body)
    worksheet117.write('U5', 'IND', body)
    worksheet117.write('V5', 'ENG', body)
    worksheet117.write('W5', 'SEJ', body)
    worksheet117.write('X5', 'GEO', body)
    worksheet117.write('Y5', 'EKO', body)
    worksheet117.write('Z5', 'SOS', body)
    worksheet117.write('AA5', 'FIS', body)
    worksheet117.write('AB5', 'KIM', body)
    worksheet117.write('AC5', 'BIO', body)
    worksheet117.write('AD5', 'JML', body)

    worksheet117.conditional_format(5, 0, row117_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet117.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF KAMPUNG MELAYU', title)
    worksheet117.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet117.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet117.write('A22', 'LOKASI', header)
    worksheet117.write('B22', 'TOTAL', header)
    worksheet117.merge_range('A21:B21', 'RANK', header)
    worksheet117.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet117.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet117.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet117.merge_range('F21:F22', 'KELAS', header)
    worksheet117.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet117.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet117.write('G22', 'MAW', body)
    worksheet117.write('H22', 'MAP', body)
    worksheet117.write('I22', 'IND', body)
    worksheet117.write('J22', 'ENG', body)
    worksheet117.write('K22', 'SEJ', body)
    worksheet117.write('L22', 'GEO', body)
    worksheet117.write('M22', 'EKO', body)
    worksheet117.write('N22', 'SOS', body)
    worksheet117.write('O22', 'FIS', body)
    worksheet117.write('P22', 'KIM', body)
    worksheet117.write('Q22', 'BIO', body)
    worksheet117.write('R22', 'JML', body)
    worksheet117.write('S22', 'MAW', body)
    worksheet117.write('T22', 'MAP', body)
    worksheet117.write('U22', 'IND', body)
    worksheet117.write('V22', 'ENG', body)
    worksheet117.write('W22', 'SEJ', body)
    worksheet117.write('X22', 'GEO', body)
    worksheet117.write('Y22', 'EKO', body)
    worksheet117.write('Z22', 'SOS', body)
    worksheet117.write('AA22', 'FIS', body)
    worksheet117.write('AB22', 'KIM', body)
    worksheet117.write('AC22', 'BIO', body)
    worksheet117.write('AD22', 'JML', body)

    worksheet117.conditional_format(22, 0, row117+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 118
    worksheet118.insert_image('A1', r'logo resmi nf.jpg')

    worksheet118.set_column('A:A', 7, center)
    worksheet118.set_column('B:B', 6, center)
    worksheet118.set_column('C:C', 18.14, center)
    worksheet118.set_column('D:D', 25, left)
    worksheet118.set_column('E:E', 13.14, left)
    worksheet118.set_column('F:F', 8.57, center)
    worksheet118.set_column('G:AD', 5, center)
    worksheet118.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF AKSES UI', title)
    worksheet118.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet118.write('A5', 'LOKASI', header)
    worksheet118.write('B5', 'TOTAL', header)
    worksheet118.merge_range('A4:B4', 'RANK', header)
    worksheet118.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet118.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet118.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet118.merge_range('F4:F5', 'KELAS', header)
    worksheet118.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet118.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet118.write('G5', 'MAW', body)
    worksheet118.write('H5', 'MAP', body)
    worksheet118.write('I5', 'IND', body)
    worksheet118.write('J5', 'ENG', body)
    worksheet118.write('K5', 'SEJ', body)
    worksheet118.write('L5', 'GEO', body)
    worksheet118.write('M5', 'EKO', body)
    worksheet118.write('N5', 'SOS', body)
    worksheet118.write('O5', 'FIS', body)
    worksheet118.write('P5', 'KIM', body)
    worksheet118.write('Q5', 'BIO', body)
    worksheet118.write('R5', 'JML', body)
    worksheet118.write('S5', 'MAW', body)
    worksheet118.write('T5', 'MAP', body)
    worksheet118.write('U5', 'IND', body)
    worksheet118.write('V5', 'ENG', body)
    worksheet118.write('W5', 'SEJ', body)
    worksheet118.write('X5', 'GEO', body)
    worksheet118.write('Y5', 'EKO', body)
    worksheet118.write('Z5', 'SOS', body)
    worksheet118.write('AA5', 'FIS', body)
    worksheet118.write('AB5', 'KIM', body)
    worksheet118.write('AC5', 'BIO', body)
    worksheet118.write('AD5', 'JML', body)

    worksheet118.conditional_format(5, 0, row118_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet118.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF AKSES UI', title)
    worksheet118.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet118.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet118.write('A22', 'LOKASI', header)
    worksheet118.write('B22', 'TOTAL', header)
    worksheet118.merge_range('A21:B21', 'RANK', header)
    worksheet118.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet118.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet118.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet118.merge_range('F21:F22', 'KELAS', header)
    worksheet118.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet118.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet118.write('G22', 'MAW', body)
    worksheet118.write('H22', 'MAP', body)
    worksheet118.write('I22', 'IND', body)
    worksheet118.write('J22', 'ENG', body)
    worksheet118.write('K22', 'SEJ', body)
    worksheet118.write('L22', 'GEO', body)
    worksheet118.write('M22', 'EKO', body)
    worksheet118.write('N22', 'SOS', body)
    worksheet118.write('O22', 'FIS', body)
    worksheet118.write('P22', 'KIM', body)
    worksheet118.write('Q22', 'BIO', body)
    worksheet118.write('R22', 'JML', body)
    worksheet118.write('S22', 'MAW', body)
    worksheet118.write('T22', 'MAP', body)
    worksheet118.write('U22', 'IND', body)
    worksheet118.write('V22', 'ENG', body)
    worksheet118.write('W22', 'SEJ', body)
    worksheet118.write('X22', 'GEO', body)
    worksheet118.write('Y22', 'EKO', body)
    worksheet118.write('Z22', 'SOS', body)
    worksheet118.write('AA22', 'FIS', body)
    worksheet118.write('AB22', 'KIM', body)
    worksheet118.write('AC22', 'BIO', body)
    worksheet118.write('AD22', 'JML', body)

    worksheet118.conditional_format(22, 0, row118+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 119
    worksheet119.insert_image('A1', r'logo resmi nf.jpg')

    worksheet119.set_column('A:A', 7, center)
    worksheet119.set_column('B:B', 6, center)
    worksheet119.set_column('C:C', 18.14, center)
    worksheet119.set_column('D:D', 25, left)
    worksheet119.set_column('E:E', 13.14, left)
    worksheet119.set_column('F:F', 8.57, center)
    worksheet119.set_column('G:AD', 5, center)
    worksheet119.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF JATIMEKAR', title)
    worksheet119.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet119.write('A5', 'LOKASI', header)
    worksheet119.write('B5', 'TOTAL', header)
    worksheet119.merge_range('A4:B4', 'RANK', header)
    worksheet119.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet119.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet119.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet119.merge_range('F4:F5', 'KELAS', header)
    worksheet119.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet119.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet119.write('G5', 'MAW', body)
    worksheet119.write('H5', 'MAP', body)
    worksheet119.write('I5', 'IND', body)
    worksheet119.write('J5', 'ENG', body)
    worksheet119.write('K5', 'SEJ', body)
    worksheet119.write('L5', 'GEO', body)
    worksheet119.write('M5', 'EKO', body)
    worksheet119.write('N5', 'SOS', body)
    worksheet119.write('O5', 'FIS', body)
    worksheet119.write('P5', 'KIM', body)
    worksheet119.write('Q5', 'BIO', body)
    worksheet119.write('R5', 'JML', body)
    worksheet119.write('S5', 'MAW', body)
    worksheet119.write('T5', 'MAP', body)
    worksheet119.write('U5', 'IND', body)
    worksheet119.write('V5', 'ENG', body)
    worksheet119.write('W5', 'SEJ', body)
    worksheet119.write('X5', 'GEO', body)
    worksheet119.write('Y5', 'EKO', body)
    worksheet119.write('Z5', 'SOS', body)
    worksheet119.write('AA5', 'FIS', body)
    worksheet119.write('AB5', 'KIM', body)
    worksheet119.write('AC5', 'BIO', body)
    worksheet119.write('AD5', 'JML', body)

    worksheet119.conditional_format(5, 0, row119_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet119.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF JATIMEKAR', title)
    worksheet119.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet119.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet119.write('A22', 'LOKASI', header)
    worksheet119.write('B22', 'TOTAL', header)
    worksheet119.merge_range('A21:B21', 'RANK', header)
    worksheet119.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet119.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet119.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet119.merge_range('F21:F22', 'KELAS', header)
    worksheet119.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet119.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet119.write('G22', 'MAW', body)
    worksheet119.write('H22', 'MAP', body)
    worksheet119.write('I22', 'IND', body)
    worksheet119.write('J22', 'ENG', body)
    worksheet119.write('K22', 'SEJ', body)
    worksheet119.write('L22', 'GEO', body)
    worksheet119.write('M22', 'EKO', body)
    worksheet119.write('N22', 'SOS', body)
    worksheet119.write('O22', 'FIS', body)
    worksheet119.write('P22', 'KIM', body)
    worksheet119.write('Q22', 'BIO', body)
    worksheet119.write('R22', 'JML', body)
    worksheet119.write('S22', 'MAW', body)
    worksheet119.write('T22', 'MAP', body)
    worksheet119.write('U22', 'IND', body)
    worksheet119.write('V22', 'ENG', body)
    worksheet119.write('W22', 'SEJ', body)
    worksheet119.write('X22', 'GEO', body)
    worksheet119.write('Y22', 'EKO', body)
    worksheet119.write('Z22', 'SOS', body)
    worksheet119.write('AA22', 'FIS', body)
    worksheet119.write('AB22', 'KIM', body)
    worksheet119.write('AC22', 'BIO', body)
    worksheet119.write('AD22', 'JML', body)

    worksheet119.conditional_format(22, 0, row119+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 120
    worksheet120.insert_image('A1', r'logo resmi nf.jpg')

    worksheet120.set_column('A:A', 7, center)
    worksheet120.set_column('B:B', 6, center)
    worksheet120.set_column('C:C', 18.14, center)
    worksheet120.set_column('D:D', 25, left)
    worksheet120.set_column('E:E', 13.14, left)
    worksheet120.set_column('F:F', 8.57, center)
    worksheet120.set_column('G:AD', 5, center)
    worksheet120.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF RAWALUMBU', title)
    worksheet120.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet120.write('A5', 'LOKASI', header)
    worksheet120.write('B5', 'TOTAL', header)
    worksheet120.merge_range('A4:B4', 'RANK', header)
    worksheet120.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet120.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet120.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet120.merge_range('F4:F5', 'KELAS', header)
    worksheet120.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet120.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet120.write('G5', 'MAW', body)
    worksheet120.write('H5', 'MAP', body)
    worksheet120.write('I5', 'IND', body)
    worksheet120.write('J5', 'ENG', body)
    worksheet120.write('K5', 'SEJ', body)
    worksheet120.write('L5', 'GEO', body)
    worksheet120.write('M5', 'EKO', body)
    worksheet120.write('N5', 'SOS', body)
    worksheet120.write('O5', 'FIS', body)
    worksheet120.write('P5', 'KIM', body)
    worksheet120.write('Q5', 'BIO', body)
    worksheet120.write('R5', 'JML', body)
    worksheet120.write('S5', 'MAW', body)
    worksheet120.write('T5', 'MAP', body)
    worksheet120.write('U5', 'IND', body)
    worksheet120.write('V5', 'ENG', body)
    worksheet120.write('W5', 'SEJ', body)
    worksheet120.write('X5', 'GEO', body)
    worksheet120.write('Y5', 'EKO', body)
    worksheet120.write('Z5', 'SOS', body)
    worksheet120.write('AA5', 'FIS', body)
    worksheet120.write('AB5', 'KIM', body)
    worksheet120.write('AC5', 'BIO', body)
    worksheet120.write('AD5', 'JML', body)

    worksheet120.conditional_format(5, 0, row120_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet120.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF RAWALUMBU', title)
    worksheet120.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet120.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet120.write('A22', 'LOKASI', header)
    worksheet120.write('B22', 'TOTAL', header)
    worksheet120.merge_range('A21:B21', 'RANK', header)
    worksheet120.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet120.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet120.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet120.merge_range('F21:F22', 'KELAS', header)
    worksheet120.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet120.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet120.write('G22', 'MAW', body)
    worksheet120.write('H22', 'MAP', body)
    worksheet120.write('I22', 'IND', body)
    worksheet120.write('J22', 'ENG', body)
    worksheet120.write('K22', 'SEJ', body)
    worksheet120.write('L22', 'GEO', body)
    worksheet120.write('M22', 'EKO', body)
    worksheet120.write('N22', 'SOS', body)
    worksheet120.write('O22', 'FIS', body)
    worksheet120.write('P22', 'KIM', body)
    worksheet120.write('Q22', 'BIO', body)
    worksheet120.write('R22', 'JML', body)
    worksheet120.write('S22', 'MAW', body)
    worksheet120.write('T22', 'MAP', body)
    worksheet120.write('U22', 'IND', body)
    worksheet120.write('V22', 'ENG', body)
    worksheet120.write('W22', 'SEJ', body)
    worksheet120.write('X22', 'GEO', body)
    worksheet120.write('Y22', 'EKO', body)
    worksheet120.write('Z22', 'SOS', body)
    worksheet120.write('AA22', 'FIS', body)
    worksheet120.write('AB22', 'KIM', body)
    worksheet120.write('AC22', 'BIO', body)
    worksheet120.write('AD22', 'JML', body)

    worksheet120.conditional_format(22, 0, row120+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 121
    worksheet121.insert_image('A1', r'logo resmi nf.jpg')

    worksheet121.set_column('A:A', 7, center)
    worksheet121.set_column('B:B', 6, center)
    worksheet121.set_column('C:C', 18.14, center)
    worksheet121.set_column('D:D', 25, left)
    worksheet121.set_column('E:E', 13.14, left)
    worksheet121.set_column('F:F', 8.57, center)
    worksheet121.set_column('G:AD', 5, center)
    worksheet121.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF TAMAN HARAPAN BARU', title)
    worksheet121.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet121.write('A5', 'LOKASI', header)
    worksheet121.write('B5', 'TOTAL', header)
    worksheet121.merge_range('A4:B4', 'RANK', header)
    worksheet121.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet121.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet121.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet121.merge_range('F4:F5', 'KELAS', header)
    worksheet121.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet121.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet121.write('G5', 'MAW', body)
    worksheet121.write('H5', 'MAP', body)
    worksheet121.write('I5', 'IND', body)
    worksheet121.write('J5', 'ENG', body)
    worksheet121.write('K5', 'SEJ', body)
    worksheet121.write('L5', 'GEO', body)
    worksheet121.write('M5', 'EKO', body)
    worksheet121.write('N5', 'SOS', body)
    worksheet121.write('O5', 'FIS', body)
    worksheet121.write('P5', 'KIM', body)
    worksheet121.write('Q5', 'BIO', body)
    worksheet121.write('R5', 'JML', body)
    worksheet121.write('S5', 'MAW', body)
    worksheet121.write('T5', 'MAP', body)
    worksheet121.write('U5', 'IND', body)
    worksheet121.write('V5', 'ENG', body)
    worksheet121.write('W5', 'SEJ', body)
    worksheet121.write('X5', 'GEO', body)
    worksheet121.write('Y5', 'EKO', body)
    worksheet121.write('Z5', 'SOS', body)
    worksheet121.write('AA5', 'FIS', body)
    worksheet121.write('AB5', 'KIM', body)
    worksheet121.write('AC5', 'BIO', body)
    worksheet121.write('AD5', 'JML', body)

    worksheet121.conditional_format(5, 0, row121_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet121.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF TAMAN HARAPAN BARU', title)
    worksheet121.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet121.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet121.write('A22', 'LOKASI', header)
    worksheet121.write('B22', 'TOTAL', header)
    worksheet121.merge_range('A21:B21', 'RANK', header)
    worksheet121.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet121.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet121.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet121.merge_range('F21:F22', 'KELAS', header)
    worksheet121.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet121.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet121.write('G22', 'MAW', body)
    worksheet121.write('H22', 'MAP', body)
    worksheet121.write('I22', 'IND', body)
    worksheet121.write('J22', 'ENG', body)
    worksheet121.write('K22', 'SEJ', body)
    worksheet121.write('L22', 'GEO', body)
    worksheet121.write('M22', 'EKO', body)
    worksheet121.write('N22', 'SOS', body)
    worksheet121.write('O22', 'FIS', body)
    worksheet121.write('P22', 'KIM', body)
    worksheet121.write('Q22', 'BIO', body)
    worksheet121.write('R22', 'JML', body)
    worksheet121.write('S22', 'MAW', body)
    worksheet121.write('T22', 'MAP', body)
    worksheet121.write('U22', 'IND', body)
    worksheet121.write('V22', 'ENG', body)
    worksheet121.write('W22', 'SEJ', body)
    worksheet121.write('X22', 'GEO', body)
    worksheet121.write('Y22', 'EKO', body)
    worksheet121.write('Z22', 'SOS', body)
    worksheet121.write('AA22', 'FIS', body)
    worksheet121.write('AB22', 'KIM', body)
    worksheet121.write('AC22', 'BIO', body)
    worksheet121.write('AD22', 'JML', body)

    worksheet121.conditional_format(22, 0, row121+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 122
    worksheet122.insert_image('A1', r'logo resmi nf.jpg')

    worksheet122.set_column('A:A', 7, center)
    worksheet122.set_column('B:B', 6, center)
    worksheet122.set_column('C:C', 18.14, center)
    worksheet122.set_column('D:D', 25, left)
    worksheet122.set_column('E:E', 13.14, left)
    worksheet122.set_column('F:F', 8.57, center)
    worksheet122.set_column('G:AD', 5, center)
    worksheet122.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF VILA NUSA INDAH', title)
    worksheet122.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet122.write('A5', 'LOKASI', header)
    worksheet122.write('B5', 'TOTAL', header)
    worksheet122.merge_range('A4:B4', 'RANK', header)
    worksheet122.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet122.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet122.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet122.merge_range('F4:F5', 'KELAS', header)
    worksheet122.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet122.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet122.write('G5', 'MAW', body)
    worksheet122.write('H5', 'MAP', body)
    worksheet122.write('I5', 'IND', body)
    worksheet122.write('J5', 'ENG', body)
    worksheet122.write('K5', 'SEJ', body)
    worksheet122.write('L5', 'GEO', body)
    worksheet122.write('M5', 'EKO', body)
    worksheet122.write('N5', 'SOS', body)
    worksheet122.write('O5', 'FIS', body)
    worksheet122.write('P5', 'KIM', body)
    worksheet122.write('Q5', 'BIO', body)
    worksheet122.write('R5', 'JML', body)
    worksheet122.write('S5', 'MAW', body)
    worksheet122.write('T5', 'MAP', body)
    worksheet122.write('U5', 'IND', body)
    worksheet122.write('V5', 'ENG', body)
    worksheet122.write('W5', 'SEJ', body)
    worksheet122.write('X5', 'GEO', body)
    worksheet122.write('Y5', 'EKO', body)
    worksheet122.write('Z5', 'SOS', body)
    worksheet122.write('AA5', 'FIS', body)
    worksheet122.write('AB5', 'KIM', body)
    worksheet122.write('AC5', 'BIO', body)
    worksheet122.write('AD5', 'JML', body)

    worksheet122.conditional_format(5, 0, row122_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet122.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF VILA NUSA INDAH', title)
    worksheet122.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet122.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet122.write('A22', 'LOKASI', header)
    worksheet122.write('B22', 'TOTAL', header)
    worksheet122.merge_range('A21:B21', 'RANK', header)
    worksheet122.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet122.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet122.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet122.merge_range('F21:F22', 'KELAS', header)
    worksheet122.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet122.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet122.write('G22', 'MAW', body)
    worksheet122.write('H22', 'MAP', body)
    worksheet122.write('I22', 'IND', body)
    worksheet122.write('J22', 'ENG', body)
    worksheet122.write('K22', 'SEJ', body)
    worksheet122.write('L22', 'GEO', body)
    worksheet122.write('M22', 'EKO', body)
    worksheet122.write('N22', 'SOS', body)
    worksheet122.write('O22', 'FIS', body)
    worksheet122.write('P22', 'KIM', body)
    worksheet122.write('Q22', 'BIO', body)
    worksheet122.write('R22', 'JML', body)
    worksheet122.write('S22', 'MAW', body)
    worksheet122.write('T22', 'MAP', body)
    worksheet122.write('U22', 'IND', body)
    worksheet122.write('V22', 'ENG', body)
    worksheet122.write('W22', 'SEJ', body)
    worksheet122.write('X22', 'GEO', body)
    worksheet122.write('Y22', 'EKO', body)
    worksheet122.write('Z22', 'SOS', body)
    worksheet122.write('AA22', 'FIS', body)
    worksheet122.write('AB22', 'KIM', body)
    worksheet122.write('AC22', 'BIO', body)
    worksheet122.write('AD22', 'JML', body)

    worksheet122.conditional_format(22, 0, row122+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 123
    worksheet123.insert_image('A1', r'logo resmi nf.jpg')

    worksheet123.set_column('A:A', 7, center)
    worksheet123.set_column('B:B', 6, center)
    worksheet123.set_column('C:C', 18.14, center)
    worksheet123.set_column('D:D', 25, left)
    worksheet123.set_column('E:E', 13.14, left)
    worksheet123.set_column('F:F', 8.57, center)
    worksheet123.set_column('G:AD', 5, center)
    worksheet123.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF JATIWARNA', title)
    worksheet123.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet123.write('A5', 'LOKASI', header)
    worksheet123.write('B5', 'TOTAL', header)
    worksheet123.merge_range('A4:B4', 'RANK', header)
    worksheet123.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet123.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet123.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet123.merge_range('F4:F5', 'KELAS', header)
    worksheet123.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet123.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet123.write('G5', 'MAW', body)
    worksheet123.write('H5', 'MAP', body)
    worksheet123.write('I5', 'IND', body)
    worksheet123.write('J5', 'ENG', body)
    worksheet123.write('K5', 'SEJ', body)
    worksheet123.write('L5', 'GEO', body)
    worksheet123.write('M5', 'EKO', body)
    worksheet123.write('N5', 'SOS', body)
    worksheet123.write('O5', 'FIS', body)
    worksheet123.write('P5', 'KIM', body)
    worksheet123.write('Q5', 'BIO', body)
    worksheet123.write('R5', 'JML', body)
    worksheet123.write('S5', 'MAW', body)
    worksheet123.write('T5', 'MAP', body)
    worksheet123.write('U5', 'IND', body)
    worksheet123.write('V5', 'ENG', body)
    worksheet123.write('W5', 'SEJ', body)
    worksheet123.write('X5', 'GEO', body)
    worksheet123.write('Y5', 'EKO', body)
    worksheet123.write('Z5', 'SOS', body)
    worksheet123.write('AA5', 'FIS', body)
    worksheet123.write('AB5', 'KIM', body)
    worksheet123.write('AC5', 'BIO', body)
    worksheet123.write('AD5', 'JML', body)

    worksheet123.conditional_format(5, 0, row123_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet123.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF JATIWARNA', title)
    worksheet123.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet123.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet123.write('A22', 'LOKASI', header)
    worksheet123.write('B22', 'TOTAL', header)
    worksheet123.merge_range('A21:B21', 'RANK', header)
    worksheet123.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet123.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet123.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet123.merge_range('F21:F22', 'KELAS', header)
    worksheet123.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet123.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet123.write('G22', 'MAW', body)
    worksheet123.write('H22', 'MAP', body)
    worksheet123.write('I22', 'IND', body)
    worksheet123.write('J22', 'ENG', body)
    worksheet123.write('K22', 'SEJ', body)
    worksheet123.write('L22', 'GEO', body)
    worksheet123.write('M22', 'EKO', body)
    worksheet123.write('N22', 'SOS', body)
    worksheet123.write('O22', 'FIS', body)
    worksheet123.write('P22', 'KIM', body)
    worksheet123.write('Q22', 'BIO', body)
    worksheet123.write('R22', 'JML', body)
    worksheet123.write('S22', 'MAW', body)
    worksheet123.write('T22', 'MAP', body)
    worksheet123.write('U22', 'IND', body)
    worksheet123.write('V22', 'ENG', body)
    worksheet123.write('W22', 'SEJ', body)
    worksheet123.write('X22', 'GEO', body)
    worksheet123.write('Y22', 'EKO', body)
    worksheet123.write('Z22', 'SOS', body)
    worksheet123.write('AA22', 'FIS', body)
    worksheet123.write('AB22', 'KIM', body)
    worksheet123.write('AC22', 'BIO', body)
    worksheet123.write('AD22', 'JML', body)

    worksheet123.conditional_format(22, 0, row123+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 124
    worksheet124.insert_image('A1', r'logo resmi nf.jpg')

    worksheet124.set_column('A:A', 7, center)
    worksheet124.set_column('B:B', 6, center)
    worksheet124.set_column('C:C', 18.14, center)
    worksheet124.set_column('D:D', 25, left)
    worksheet124.set_column('E:E', 13.14, left)
    worksheet124.set_column('F:F', 8.57, center)
    worksheet124.set_column('G:AD', 5, center)
    worksheet124.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF TAMBUN', title)
    worksheet124.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet124.write('A5', 'LOKASI', header)
    worksheet124.write('B5', 'TOTAL', header)
    worksheet124.merge_range('A4:B4', 'RANK', header)
    worksheet124.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet124.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet124.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet124.merge_range('F4:F5', 'KELAS', header)
    worksheet124.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet124.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet124.write('G5', 'MAW', body)
    worksheet124.write('H5', 'MAP', body)
    worksheet124.write('I5', 'IND', body)
    worksheet124.write('J5', 'ENG', body)
    worksheet124.write('K5', 'SEJ', body)
    worksheet124.write('L5', 'GEO', body)
    worksheet124.write('M5', 'EKO', body)
    worksheet124.write('N5', 'SOS', body)
    worksheet124.write('O5', 'FIS', body)
    worksheet124.write('P5', 'KIM', body)
    worksheet124.write('Q5', 'BIO', body)
    worksheet124.write('R5', 'JML', body)
    worksheet124.write('S5', 'MAW', body)
    worksheet124.write('T5', 'MAP', body)
    worksheet124.write('U5', 'IND', body)
    worksheet124.write('V5', 'ENG', body)
    worksheet124.write('W5', 'SEJ', body)
    worksheet124.write('X5', 'GEO', body)
    worksheet124.write('Y5', 'EKO', body)
    worksheet124.write('Z5', 'SOS', body)
    worksheet124.write('AA5', 'FIS', body)
    worksheet124.write('AB5', 'KIM', body)
    worksheet124.write('AC5', 'BIO', body)
    worksheet124.write('AD5', 'JML', body)

    worksheet124.conditional_format(5, 0, row124_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet124.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF TAMBUN', title)
    worksheet124.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet124.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet124.write('A22', 'LOKASI', header)
    worksheet124.write('B22', 'TOTAL', header)
    worksheet124.merge_range('A21:B21', 'RANK', header)
    worksheet124.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet124.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet124.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet124.merge_range('F21:F22', 'KELAS', header)
    worksheet124.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet124.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet124.write('G22', 'MAW', body)
    worksheet124.write('H22', 'MAP', body)
    worksheet124.write('I22', 'IND', body)
    worksheet124.write('J22', 'ENG', body)
    worksheet124.write('K22', 'SEJ', body)
    worksheet124.write('L22', 'GEO', body)
    worksheet124.write('M22', 'EKO', body)
    worksheet124.write('N22', 'SOS', body)
    worksheet124.write('O22', 'FIS', body)
    worksheet124.write('P22', 'KIM', body)
    worksheet124.write('Q22', 'BIO', body)
    worksheet124.write('R22', 'JML', body)
    worksheet124.write('S22', 'MAW', body)
    worksheet124.write('T22', 'MAP', body)
    worksheet124.write('U22', 'IND', body)
    worksheet124.write('V22', 'ENG', body)
    worksheet124.write('W22', 'SEJ', body)
    worksheet124.write('X22', 'GEO', body)
    worksheet124.write('Y22', 'EKO', body)
    worksheet124.write('Z22', 'SOS', body)
    worksheet124.write('AA22', 'FIS', body)
    worksheet124.write('AB22', 'KIM', body)
    worksheet124.write('AC22', 'BIO', body)
    worksheet124.write('AD22', 'JML', body)

    worksheet124.conditional_format(22, 0, row124+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 125
    worksheet125.insert_image('A1', r'logo resmi nf.jpg')

    worksheet125.set_column('A:A', 7, center)
    worksheet125.set_column('B:B', 6, center)
    worksheet125.set_column('C:C', 18.14, center)
    worksheet125.set_column('D:D', 25, left)
    worksheet125.set_column('E:E', 13.14, left)
    worksheet125.set_column('F:F', 8.57, center)
    worksheet125.set_column('G:AD', 5, center)
    worksheet125.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF DAAN MOGOT', title)
    worksheet125.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet125.write('A5', 'LOKASI', header)
    worksheet125.write('B5', 'TOTAL', header)
    worksheet125.merge_range('A4:B4', 'RANK', header)
    worksheet125.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet125.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet125.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet125.merge_range('F4:F5', 'KELAS', header)
    worksheet125.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet125.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet125.write('G5', 'MAW', body)
    worksheet125.write('H5', 'MAP', body)
    worksheet125.write('I5', 'IND', body)
    worksheet125.write('J5', 'ENG', body)
    worksheet125.write('K5', 'SEJ', body)
    worksheet125.write('L5', 'GEO', body)
    worksheet125.write('M5', 'EKO', body)
    worksheet125.write('N5', 'SOS', body)
    worksheet125.write('O5', 'FIS', body)
    worksheet125.write('P5', 'KIM', body)
    worksheet125.write('Q5', 'BIO', body)
    worksheet125.write('R5', 'JML', body)
    worksheet125.write('S5', 'MAW', body)
    worksheet125.write('T5', 'MAP', body)
    worksheet125.write('U5', 'IND', body)
    worksheet125.write('V5', 'ENG', body)
    worksheet125.write('W5', 'SEJ', body)
    worksheet125.write('X5', 'GEO', body)
    worksheet125.write('Y5', 'EKO', body)
    worksheet125.write('Z5', 'SOS', body)
    worksheet125.write('AA5', 'FIS', body)
    worksheet125.write('AB5', 'KIM', body)
    worksheet125.write('AC5', 'BIO', body)
    worksheet125.write('AD5', 'JML', body)

    worksheet125.conditional_format(5, 0, row125_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet125.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF DAAN MOGOT', title)
    worksheet125.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet125.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet125.write('A22', 'LOKASI', header)
    worksheet125.write('B22', 'TOTAL', header)
    worksheet125.merge_range('A21:B21', 'RANK', header)
    worksheet125.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet125.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet125.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet125.merge_range('F21:F22', 'KELAS', header)
    worksheet125.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet125.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet125.write('G22', 'MAW', body)
    worksheet125.write('H22', 'MAP', body)
    worksheet125.write('I22', 'IND', body)
    worksheet125.write('J22', 'ENG', body)
    worksheet125.write('K22', 'SEJ', body)
    worksheet125.write('L22', 'GEO', body)
    worksheet125.write('M22', 'EKO', body)
    worksheet125.write('N22', 'SOS', body)
    worksheet125.write('O22', 'FIS', body)
    worksheet125.write('P22', 'KIM', body)
    worksheet125.write('Q22', 'BIO', body)
    worksheet125.write('R22', 'JML', body)
    worksheet125.write('S22', 'MAW', body)
    worksheet125.write('T22', 'MAP', body)
    worksheet125.write('U22', 'IND', body)
    worksheet125.write('V22', 'ENG', body)
    worksheet125.write('W22', 'SEJ', body)
    worksheet125.write('X22', 'GEO', body)
    worksheet125.write('Y22', 'EKO', body)
    worksheet125.write('Z22', 'SOS', body)
    worksheet125.write('AA22', 'FIS', body)
    worksheet125.write('AB22', 'KIM', body)
    worksheet125.write('AC22', 'BIO', body)
    worksheet125.write('AD22', 'JML', body)

    worksheet125.conditional_format(22, 0, row125+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 126
    worksheet126.insert_image('A1', r'logo resmi nf.jpg')

    worksheet126.set_column('A:A', 7, center)
    worksheet126.set_column('B:B', 6, center)
    worksheet126.set_column('C:C', 18.14, center)
    worksheet126.set_column('D:D', 25, left)
    worksheet126.set_column('E:E', 13.14, left)
    worksheet126.set_column('F:F', 8.57, center)
    worksheet126.set_column('G:AD', 5, center)
    worksheet126.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIBUBUR', title)
    worksheet126.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet126.write('A5', 'LOKASI', header)
    worksheet126.write('B5', 'TOTAL', header)
    worksheet126.merge_range('A4:B4', 'RANK', header)
    worksheet126.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet126.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet126.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet126.merge_range('F4:F5', 'KELAS', header)
    worksheet126.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet126.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet126.write('G5', 'MAW', body)
    worksheet126.write('H5', 'MAP', body)
    worksheet126.write('I5', 'IND', body)
    worksheet126.write('J5', 'ENG', body)
    worksheet126.write('K5', 'SEJ', body)
    worksheet126.write('L5', 'GEO', body)
    worksheet126.write('M5', 'EKO', body)
    worksheet126.write('N5', 'SOS', body)
    worksheet126.write('O5', 'FIS', body)
    worksheet126.write('P5', 'KIM', body)
    worksheet126.write('Q5', 'BIO', body)
    worksheet126.write('R5', 'JML', body)
    worksheet126.write('S5', 'MAW', body)
    worksheet126.write('T5', 'MAP', body)
    worksheet126.write('U5', 'IND', body)
    worksheet126.write('V5', 'ENG', body)
    worksheet126.write('W5', 'SEJ', body)
    worksheet126.write('X5', 'GEO', body)
    worksheet126.write('Y5', 'EKO', body)
    worksheet126.write('Z5', 'SOS', body)
    worksheet126.write('AA5', 'FIS', body)
    worksheet126.write('AB5', 'KIM', body)
    worksheet126.write('AC5', 'BIO', body)
    worksheet126.write('AD5', 'JML', body)

    worksheet126.conditional_format(5, 0, row126_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet126.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF CIBUBUR', title)
    worksheet126.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet126.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet126.write('A22', 'LOKASI', header)
    worksheet126.write('B22', 'TOTAL', header)
    worksheet126.merge_range('A21:B21', 'RANK', header)
    worksheet126.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet126.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet126.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet126.merge_range('F21:F22', 'KELAS', header)
    worksheet126.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet126.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet126.write('G22', 'MAW', body)
    worksheet126.write('H22', 'MAP', body)
    worksheet126.write('I22', 'IND', body)
    worksheet126.write('J22', 'ENG', body)
    worksheet126.write('K22', 'SEJ', body)
    worksheet126.write('L22', 'GEO', body)
    worksheet126.write('M22', 'EKO', body)
    worksheet126.write('N22', 'SOS', body)
    worksheet126.write('O22', 'FIS', body)
    worksheet126.write('P22', 'KIM', body)
    worksheet126.write('Q22', 'BIO', body)
    worksheet126.write('R22', 'JML', body)
    worksheet126.write('S22', 'MAW', body)
    worksheet126.write('T22', 'MAP', body)
    worksheet126.write('U22', 'IND', body)
    worksheet126.write('V22', 'ENG', body)
    worksheet126.write('W22', 'SEJ', body)
    worksheet126.write('X22', 'GEO', body)
    worksheet126.write('Y22', 'EKO', body)
    worksheet126.write('Z22', 'SOS', body)
    worksheet126.write('AA22', 'FIS', body)
    worksheet126.write('AB22', 'KIM', body)
    worksheet126.write('AC22', 'BIO', body)
    worksheet126.write('AD22', 'JML', body)

    worksheet126.conditional_format(22, 0, row126+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 127
    worksheet127.insert_image('A1', r'logo resmi nf.jpg')

    worksheet127.set_column('A:A', 7, center)
    worksheet127.set_column('B:B', 6, center)
    worksheet127.set_column('C:C', 18.14, center)
    worksheet127.set_column('D:D', 25, left)
    worksheet127.set_column('E:E', 13.14, left)
    worksheet127.set_column('F:F', 8.57, center)
    worksheet127.set_column('G:AD', 5, center)
    worksheet127.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CENGKARENG', title)
    worksheet127.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet127.write('A5', 'LOKASI', header)
    worksheet127.write('B5', 'TOTAL', header)
    worksheet127.merge_range('A4:B4', 'RANK', header)
    worksheet127.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet127.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet127.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet127.merge_range('F4:F5', 'KELAS', header)
    worksheet127.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet127.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet127.write('G5', 'MAW', body)
    worksheet127.write('H5', 'MAP', body)
    worksheet127.write('I5', 'IND', body)
    worksheet127.write('J5', 'ENG', body)
    worksheet127.write('K5', 'SEJ', body)
    worksheet127.write('L5', 'GEO', body)
    worksheet127.write('M5', 'EKO', body)
    worksheet127.write('N5', 'SOS', body)
    worksheet127.write('O5', 'FIS', body)
    worksheet127.write('P5', 'KIM', body)
    worksheet127.write('Q5', 'BIO', body)
    worksheet127.write('R5', 'JML', body)
    worksheet127.write('S5', 'MAW', body)
    worksheet127.write('T5', 'MAP', body)
    worksheet127.write('U5', 'IND', body)
    worksheet127.write('V5', 'ENG', body)
    worksheet127.write('W5', 'SEJ', body)
    worksheet127.write('X5', 'GEO', body)
    worksheet127.write('Y5', 'EKO', body)
    worksheet127.write('Z5', 'SOS', body)
    worksheet127.write('AA5', 'FIS', body)
    worksheet127.write('AB5', 'KIM', body)
    worksheet127.write('AC5', 'BIO', body)
    worksheet127.write('AD5', 'JML', body)

    worksheet127.conditional_format(5, 0, row127_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet127.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF CENGKARENG', title)
    worksheet127.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet127.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet127.write('A22', 'LOKASI', header)
    worksheet127.write('B22', 'TOTAL', header)
    worksheet127.merge_range('A21:B21', 'RANK', header)
    worksheet127.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet127.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet127.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet127.merge_range('F21:F22', 'KELAS', header)
    worksheet127.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet127.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet127.write('G22', 'MAW', body)
    worksheet127.write('H22', 'MAP', body)
    worksheet127.write('I22', 'IND', body)
    worksheet127.write('J22', 'ENG', body)
    worksheet127.write('K22', 'SEJ', body)
    worksheet127.write('L22', 'GEO', body)
    worksheet127.write('M22', 'EKO', body)
    worksheet127.write('N22', 'SOS', body)
    worksheet127.write('O22', 'FIS', body)
    worksheet127.write('P22', 'KIM', body)
    worksheet127.write('Q22', 'BIO', body)
    worksheet127.write('R22', 'JML', body)
    worksheet127.write('S22', 'MAW', body)
    worksheet127.write('T22', 'MAP', body)
    worksheet127.write('U22', 'IND', body)
    worksheet127.write('V22', 'ENG', body)
    worksheet127.write('W22', 'SEJ', body)
    worksheet127.write('X22', 'GEO', body)
    worksheet127.write('Y22', 'EKO', body)
    worksheet127.write('Z22', 'SOS', body)
    worksheet127.write('AA22', 'FIS', body)
    worksheet127.write('AB22', 'KIM', body)
    worksheet127.write('AC22', 'BIO', body)
    worksheet127.write('AD22', 'JML', body)

    worksheet127.conditional_format(22, 0, row127+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 128
    worksheet128.insert_image('A1', r'logo resmi nf.jpg')

    worksheet128.set_column('A:A', 7, center)
    worksheet128.set_column('B:B', 6, center)
    worksheet128.set_column('C:C', 18.14, center)
    worksheet128.set_column('D:D', 25, left)
    worksheet128.set_column('E:E', 13.14, left)
    worksheet128.set_column('F:F', 8.57, center)
    worksheet128.set_column('G:AD', 5, center)
    worksheet128.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PETUKANGAN', title)
    worksheet128.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet128.write('A5', 'LOKASI', header)
    worksheet128.write('B5', 'TOTAL', header)
    worksheet128.merge_range('A4:B4', 'RANK', header)
    worksheet128.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet128.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet128.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet128.merge_range('F4:F5', 'KELAS', header)
    worksheet128.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet128.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet128.write('G5', 'MAW', body)
    worksheet128.write('H5', 'MAP', body)
    worksheet128.write('I5', 'IND', body)
    worksheet128.write('J5', 'ENG', body)
    worksheet128.write('K5', 'SEJ', body)
    worksheet128.write('L5', 'GEO', body)
    worksheet128.write('M5', 'EKO', body)
    worksheet128.write('N5', 'SOS', body)
    worksheet128.write('O5', 'FIS', body)
    worksheet128.write('P5', 'KIM', body)
    worksheet128.write('Q5', 'BIO', body)
    worksheet128.write('R5', 'JML', body)
    worksheet128.write('S5', 'MAW', body)
    worksheet128.write('T5', 'MAP', body)
    worksheet128.write('U5', 'IND', body)
    worksheet128.write('V5', 'ENG', body)
    worksheet128.write('W5', 'SEJ', body)
    worksheet128.write('X5', 'GEO', body)
    worksheet128.write('Y5', 'EKO', body)
    worksheet128.write('Z5', 'SOS', body)
    worksheet128.write('AA5', 'FIS', body)
    worksheet128.write('AB5', 'KIM', body)
    worksheet128.write('AC5', 'BIO', body)
    worksheet128.write('AD5', 'JML', body)

    worksheet128.conditional_format(5, 0, row128_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet128.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF PETUKANGAN', title)
    worksheet128.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet128.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet128.write('A22', 'LOKASI', header)
    worksheet128.write('B22', 'TOTAL', header)
    worksheet128.merge_range('A21:B21', 'RANK', header)
    worksheet128.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet128.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet128.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet128.merge_range('F21:F22', 'KELAS', header)
    worksheet128.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet128.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet128.write('G22', 'MAW', body)
    worksheet128.write('H22', 'MAP', body)
    worksheet128.write('I22', 'IND', body)
    worksheet128.write('J22', 'ENG', body)
    worksheet128.write('K22', 'SEJ', body)
    worksheet128.write('L22', 'GEO', body)
    worksheet128.write('M22', 'EKO', body)
    worksheet128.write('N22', 'SOS', body)
    worksheet128.write('O22', 'FIS', body)
    worksheet128.write('P22', 'KIM', body)
    worksheet128.write('Q22', 'BIO', body)
    worksheet128.write('R22', 'JML', body)
    worksheet128.write('S22', 'MAW', body)
    worksheet128.write('T22', 'MAP', body)
    worksheet128.write('U22', 'IND', body)
    worksheet128.write('V22', 'ENG', body)
    worksheet128.write('W22', 'SEJ', body)
    worksheet128.write('X22', 'GEO', body)
    worksheet128.write('Y22', 'EKO', body)
    worksheet128.write('Z22', 'SOS', body)
    worksheet128.write('AA22', 'FIS', body)
    worksheet128.write('AB22', 'KIM', body)
    worksheet128.write('AC22', 'BIO', body)
    worksheet128.write('AD22', 'JML', body)

    worksheet128.conditional_format(22, 0, row128+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 129
    worksheet129.insert_image('A1', r'logo resmi nf.jpg')

    worksheet129.set_column('A:A', 7, center)
    worksheet129.set_column('B:B', 6, center)
    worksheet129.set_column('C:C', 18.14, center)
    worksheet129.set_column('D:D', 25, left)
    worksheet129.set_column('E:E', 13.14, left)
    worksheet129.set_column('F:F', 8.57, center)
    worksheet129.set_column('G:AD', 5, center)
    worksheet129.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MERUYA UTARA', title)
    worksheet129.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet129.write('A5', 'LOKASI', header)
    worksheet129.write('B5', 'TOTAL', header)
    worksheet129.merge_range('A4:B4', 'RANK', header)
    worksheet129.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet129.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet129.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet129.merge_range('F4:F5', 'KELAS', header)
    worksheet129.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet129.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet129.write('G5', 'MAW', body)
    worksheet129.write('H5', 'MAP', body)
    worksheet129.write('I5', 'IND', body)
    worksheet129.write('J5', 'ENG', body)
    worksheet129.write('K5', 'SEJ', body)
    worksheet129.write('L5', 'GEO', body)
    worksheet129.write('M5', 'EKO', body)
    worksheet129.write('N5', 'SOS', body)
    worksheet129.write('O5', 'FIS', body)
    worksheet129.write('P5', 'KIM', body)
    worksheet129.write('Q5', 'BIO', body)
    worksheet129.write('R5', 'JML', body)
    worksheet129.write('S5', 'MAW', body)
    worksheet129.write('T5', 'MAP', body)
    worksheet129.write('U5', 'IND', body)
    worksheet129.write('V5', 'ENG', body)
    worksheet129.write('W5', 'SEJ', body)
    worksheet129.write('X5', 'GEO', body)
    worksheet129.write('Y5', 'EKO', body)
    worksheet129.write('Z5', 'SOS', body)
    worksheet129.write('AA5', 'FIS', body)
    worksheet129.write('AB5', 'KIM', body)
    worksheet129.write('AC5', 'BIO', body)
    worksheet129.write('AD5', 'JML', body)

    worksheet129.conditional_format(5, 0, row129_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet129.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF MERUYA UTARA', title)
    worksheet129.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet129.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet129.write('A22', 'LOKASI', header)
    worksheet129.write('B22', 'TOTAL', header)
    worksheet129.merge_range('A21:B21', 'RANK', header)
    worksheet129.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet129.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet129.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet129.merge_range('F21:F22', 'KELAS', header)
    worksheet129.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet129.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet129.write('G22', 'MAW', body)
    worksheet129.write('H22', 'MAP', body)
    worksheet129.write('I22', 'IND', body)
    worksheet129.write('J22', 'ENG', body)
    worksheet129.write('K22', 'SEJ', body)
    worksheet129.write('L22', 'GEO', body)
    worksheet129.write('M22', 'EKO', body)
    worksheet129.write('N22', 'SOS', body)
    worksheet129.write('O22', 'FIS', body)
    worksheet129.write('P22', 'KIM', body)
    worksheet129.write('Q22', 'BIO', body)
    worksheet129.write('R22', 'JML', body)
    worksheet129.write('S22', 'MAW', body)
    worksheet129.write('T22', 'MAP', body)
    worksheet129.write('U22', 'IND', body)
    worksheet129.write('V22', 'ENG', body)
    worksheet129.write('W22', 'SEJ', body)
    worksheet129.write('X22', 'GEO', body)
    worksheet129.write('Y22', 'EKO', body)
    worksheet129.write('Z22', 'SOS', body)
    worksheet129.write('AA22', 'FIS', body)
    worksheet129.write('AB22', 'KIM', body)
    worksheet129.write('AC22', 'BIO', body)
    worksheet129.write('AD22', 'JML', body)

    worksheet129.conditional_format(22, 0, row129+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 130
    worksheet130.insert_image('A1', r'logo resmi nf.jpg')

    worksheet130.set_column('A:A', 7, center)
    worksheet130.set_column('B:B', 6, center)
    worksheet130.set_column('C:C', 18.14, center)
    worksheet130.set_column('D:D', 25, left)
    worksheet130.set_column('E:E', 13.14, left)
    worksheet130.set_column('F:F', 8.57, center)
    worksheet130.set_column('G:AD', 5, center)
    worksheet130.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BINTARA', title)
    worksheet130.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet130.write('A5', 'LOKASI', header)
    worksheet130.write('B5', 'TOTAL', header)
    worksheet130.merge_range('A4:B4', 'RANK', header)
    worksheet130.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet130.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet130.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet130.merge_range('F4:F5', 'KELAS', header)
    worksheet130.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet130.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet130.write('G5', 'MAW', body)
    worksheet130.write('H5', 'MAP', body)
    worksheet130.write('I5', 'IND', body)
    worksheet130.write('J5', 'ENG', body)
    worksheet130.write('K5', 'SEJ', body)
    worksheet130.write('L5', 'GEO', body)
    worksheet130.write('M5', 'EKO', body)
    worksheet130.write('N5', 'SOS', body)
    worksheet130.write('O5', 'FIS', body)
    worksheet130.write('P5', 'KIM', body)
    worksheet130.write('Q5', 'BIO', body)
    worksheet130.write('R5', 'JML', body)
    worksheet130.write('S5', 'MAW', body)
    worksheet130.write('T5', 'MAP', body)
    worksheet130.write('U5', 'IND', body)
    worksheet130.write('V5', 'ENG', body)
    worksheet130.write('W5', 'SEJ', body)
    worksheet130.write('X5', 'GEO', body)
    worksheet130.write('Y5', 'EKO', body)
    worksheet130.write('Z5', 'SOS', body)
    worksheet130.write('AA5', 'FIS', body)
    worksheet130.write('AB5', 'KIM', body)
    worksheet130.write('AC5', 'BIO', body)
    worksheet130.write('AD5', 'JML', body)

    worksheet130.conditional_format(5, 0, row130_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet130.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF BINTARA', title)
    worksheet130.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet130.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet130.write('A22', 'LOKASI', header)
    worksheet130.write('B22', 'TOTAL', header)
    worksheet130.merge_range('A21:B21', 'RANK', header)
    worksheet130.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet130.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet130.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet130.merge_range('F21:F22', 'KELAS', header)
    worksheet130.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet130.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet130.write('G22', 'MAW', body)
    worksheet130.write('H22', 'MAP', body)
    worksheet130.write('I22', 'IND', body)
    worksheet130.write('J22', 'ENG', body)
    worksheet130.write('K22', 'SEJ', body)
    worksheet130.write('L22', 'GEO', body)
    worksheet130.write('M22', 'EKO', body)
    worksheet130.write('N22', 'SOS', body)
    worksheet130.write('O22', 'FIS', body)
    worksheet130.write('P22', 'KIM', body)
    worksheet130.write('Q22', 'BIO', body)
    worksheet130.write('R22', 'JML', body)
    worksheet130.write('S22', 'MAW', body)
    worksheet130.write('T22', 'MAP', body)
    worksheet130.write('U22', 'IND', body)
    worksheet130.write('V22', 'ENG', body)
    worksheet130.write('W22', 'SEJ', body)
    worksheet130.write('X22', 'GEO', body)
    worksheet130.write('Y22', 'EKO', body)
    worksheet130.write('Z22', 'SOS', body)
    worksheet130.write('AA22', 'FIS', body)
    worksheet130.write('AB22', 'KIM', body)
    worksheet130.write('AC22', 'BIO', body)
    worksheet130.write('AD22', 'JML', body)

    worksheet130.conditional_format(22, 0, row130+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 131
    worksheet131.insert_image('A1', r'logo resmi nf.jpg')

    worksheet131.set_column('A:A', 7, center)
    worksheet131.set_column('B:B', 6, center)
    worksheet131.set_column('C:C', 18.14, center)
    worksheet131.set_column('D:D', 25, left)
    worksheet131.set_column('E:E', 13.14, left)
    worksheet131.set_column('F:F', 8.57, center)
    worksheet131.set_column('G:AD', 5, center)
    worksheet131.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MALANG', title)
    worksheet131.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet131.write('A5', 'LOKASI', header)
    worksheet131.write('B5', 'TOTAL', header)
    worksheet131.merge_range('A4:B4', 'RANK', header)
    worksheet131.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet131.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet131.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet131.merge_range('F4:F5', 'KELAS', header)
    worksheet131.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet131.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet131.write('G5', 'MAW', body)
    worksheet131.write('H5', 'MAP', body)
    worksheet131.write('I5', 'IND', body)
    worksheet131.write('J5', 'ENG', body)
    worksheet131.write('K5', 'SEJ', body)
    worksheet131.write('L5', 'GEO', body)
    worksheet131.write('M5', 'EKO', body)
    worksheet131.write('N5', 'SOS', body)
    worksheet131.write('O5', 'FIS', body)
    worksheet131.write('P5', 'KIM', body)
    worksheet131.write('Q5', 'BIO', body)
    worksheet131.write('R5', 'JML', body)
    worksheet131.write('S5', 'MAW', body)
    worksheet131.write('T5', 'MAP', body)
    worksheet131.write('U5', 'IND', body)
    worksheet131.write('V5', 'ENG', body)
    worksheet131.write('W5', 'SEJ', body)
    worksheet131.write('X5', 'GEO', body)
    worksheet131.write('Y5', 'EKO', body)
    worksheet131.write('Z5', 'SOS', body)
    worksheet131.write('AA5', 'FIS', body)
    worksheet131.write('AB5', 'KIM', body)
    worksheet131.write('AC5', 'BIO', body)
    worksheet131.write('AD5', 'JML', body)

    worksheet131.conditional_format(5, 0, row131_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet131.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF MALANG', title)
    worksheet131.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet131.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet131.write('A22', 'LOKASI', header)
    worksheet131.write('B22', 'TOTAL', header)
    worksheet131.merge_range('A21:B21', 'RANK', header)
    worksheet131.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet131.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet131.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet131.merge_range('F21:F22', 'KELAS', header)
    worksheet131.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet131.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet131.write('G22', 'MAW', body)
    worksheet131.write('H22', 'MAP', body)
    worksheet131.write('I22', 'IND', body)
    worksheet131.write('J22', 'ENG', body)
    worksheet131.write('K22', 'SEJ', body)
    worksheet131.write('L22', 'GEO', body)
    worksheet131.write('M22', 'EKO', body)
    worksheet131.write('N22', 'SOS', body)
    worksheet131.write('O22', 'FIS', body)
    worksheet131.write('P22', 'KIM', body)
    worksheet131.write('Q22', 'BIO', body)
    worksheet131.write('R22', 'JML', body)
    worksheet131.write('S22', 'MAW', body)
    worksheet131.write('T22', 'MAP', body)
    worksheet131.write('U22', 'IND', body)
    worksheet131.write('V22', 'ENG', body)
    worksheet131.write('W22', 'SEJ', body)
    worksheet131.write('X22', 'GEO', body)
    worksheet131.write('Y22', 'EKO', body)
    worksheet131.write('Z22', 'SOS', body)
    worksheet131.write('AA22', 'FIS', body)
    worksheet131.write('AB22', 'KIM', body)
    worksheet131.write('AC22', 'BIO', body)
    worksheet131.write('AD22', 'JML', body)

    worksheet131.conditional_format(22, 0, row131+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 132
    worksheet132.insert_image('A1', r'logo resmi nf.jpg')

    worksheet132.set_column('A:A', 7, center)
    worksheet132.set_column('B:B', 6, center)
    worksheet132.set_column('C:C', 18.14, center)
    worksheet132.set_column('D:D', 25, left)
    worksheet132.set_column('E:E', 13.14, left)
    worksheet132.set_column('F:F', 8.57, center)
    worksheet132.set_column('G:AD', 5, center)
    worksheet132.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MEDAN BARU', title)
    worksheet132.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet132.write('A5', 'LOKASI', header)
    worksheet132.write('B5', 'TOTAL', header)
    worksheet132.merge_range('A4:B4', 'RANK', header)
    worksheet132.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet132.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet132.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet132.merge_range('F4:F5', 'KELAS', header)
    worksheet132.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet132.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet132.write('G5', 'MAW', body)
    worksheet132.write('H5', 'MAP', body)
    worksheet132.write('I5', 'IND', body)
    worksheet132.write('J5', 'ENG', body)
    worksheet132.write('K5', 'SEJ', body)
    worksheet132.write('L5', 'GEO', body)
    worksheet132.write('M5', 'EKO', body)
    worksheet132.write('N5', 'SOS', body)
    worksheet132.write('O5', 'FIS', body)
    worksheet132.write('P5', 'KIM', body)
    worksheet132.write('Q5', 'BIO', body)
    worksheet132.write('R5', 'JML', body)
    worksheet132.write('S5', 'MAW', body)
    worksheet132.write('T5', 'MAP', body)
    worksheet132.write('U5', 'IND', body)
    worksheet132.write('V5', 'ENG', body)
    worksheet132.write('W5', 'SEJ', body)
    worksheet132.write('X5', 'GEO', body)
    worksheet132.write('Y5', 'EKO', body)
    worksheet132.write('Z5', 'SOS', body)
    worksheet132.write('AA5', 'FIS', body)
    worksheet132.write('AB5', 'KIM', body)
    worksheet132.write('AC5', 'BIO', body)
    worksheet132.write('AD5', 'JML', body)

    worksheet132.conditional_format(5, 0, row132_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet132.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF MEDAN BARU', title)
    worksheet132.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet132.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet132.write('A22', 'LOKASI', header)
    worksheet132.write('B22', 'TOTAL', header)
    worksheet132.merge_range('A21:B21', 'RANK', header)
    worksheet132.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet132.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet132.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet132.merge_range('F21:F22', 'KELAS', header)
    worksheet132.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet132.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet132.write('G22', 'MAW', body)
    worksheet132.write('H22', 'MAP', body)
    worksheet132.write('I22', 'IND', body)
    worksheet132.write('J22', 'ENG', body)
    worksheet132.write('K22', 'SEJ', body)
    worksheet132.write('L22', 'GEO', body)
    worksheet132.write('M22', 'EKO', body)
    worksheet132.write('N22', 'SOS', body)
    worksheet132.write('O22', 'FIS', body)
    worksheet132.write('P22', 'KIM', body)
    worksheet132.write('Q22', 'BIO', body)
    worksheet132.write('R22', 'JML', body)
    worksheet132.write('S22', 'MAW', body)
    worksheet132.write('T22', 'MAP', body)
    worksheet132.write('U22', 'IND', body)
    worksheet132.write('V22', 'ENG', body)
    worksheet132.write('W22', 'SEJ', body)
    worksheet132.write('X22', 'GEO', body)
    worksheet132.write('Y22', 'EKO', body)
    worksheet132.write('Z22', 'SOS', body)
    worksheet132.write('AA22', 'FIS', body)
    worksheet132.write('AB22', 'KIM', body)
    worksheet132.write('AC22', 'BIO', body)
    worksheet132.write('AD22', 'JML', body)

    worksheet132.conditional_format(22, 0, row132+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 133
    worksheet133.insert_image('A1', r'logo resmi nf.jpg')

    worksheet133.set_column('A:A', 7, center)
    worksheet133.set_column('B:B', 6, center)
    worksheet133.set_column('C:C', 18.14, center)
    worksheet133.set_column('D:D', 25, left)
    worksheet133.set_column('E:E', 13.14, left)
    worksheet133.set_column('F:F', 8.57, center)
    worksheet133.set_column('G:AD', 5, center)
    worksheet133.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MEDAN HELVETIA', title)
    worksheet133.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet133.write('A5', 'LOKASI', header)
    worksheet133.write('B5', 'TOTAL', header)
    worksheet133.merge_range('A4:B4', 'RANK', header)
    worksheet133.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet133.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet133.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet133.merge_range('F4:F5', 'KELAS', header)
    worksheet133.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet133.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet133.write('G5', 'MAW', body)
    worksheet133.write('H5', 'MAP', body)
    worksheet133.write('I5', 'IND', body)
    worksheet133.write('J5', 'ENG', body)
    worksheet133.write('K5', 'SEJ', body)
    worksheet133.write('L5', 'GEO', body)
    worksheet133.write('M5', 'EKO', body)
    worksheet133.write('N5', 'SOS', body)
    worksheet133.write('O5', 'FIS', body)
    worksheet133.write('P5', 'KIM', body)
    worksheet133.write('Q5', 'BIO', body)
    worksheet133.write('R5', 'JML', body)
    worksheet133.write('S5', 'MAW', body)
    worksheet133.write('T5', 'MAP', body)
    worksheet133.write('U5', 'IND', body)
    worksheet133.write('V5', 'ENG', body)
    worksheet133.write('W5', 'SEJ', body)
    worksheet133.write('X5', 'GEO', body)
    worksheet133.write('Y5', 'EKO', body)
    worksheet133.write('Z5', 'SOS', body)
    worksheet133.write('AA5', 'FIS', body)
    worksheet133.write('AB5', 'KIM', body)
    worksheet133.write('AC5', 'BIO', body)
    worksheet133.write('AD5', 'JML', body)

    worksheet133.conditional_format(5, 0, row133_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet133.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF MEDAN HELVETIA', title)
    worksheet133.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet133.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet133.write('A22', 'LOKASI', header)
    worksheet133.write('B22', 'TOTAL', header)
    worksheet133.merge_range('A21:B21', 'RANK', header)
    worksheet133.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet133.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet133.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet133.merge_range('F21:F22', 'KELAS', header)
    worksheet133.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet133.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet133.write('G22', 'MAW', body)
    worksheet133.write('H22', 'MAP', body)
    worksheet133.write('I22', 'IND', body)
    worksheet133.write('J22', 'ENG', body)
    worksheet133.write('K22', 'SEJ', body)
    worksheet133.write('L22', 'GEO', body)
    worksheet133.write('M22', 'EKO', body)
    worksheet133.write('N22', 'SOS', body)
    worksheet133.write('O22', 'FIS', body)
    worksheet133.write('P22', 'KIM', body)
    worksheet133.write('Q22', 'BIO', body)
    worksheet133.write('R22', 'JML', body)
    worksheet133.write('S22', 'MAW', body)
    worksheet133.write('T22', 'MAP', body)
    worksheet133.write('U22', 'IND', body)
    worksheet133.write('V22', 'ENG', body)
    worksheet133.write('W22', 'SEJ', body)
    worksheet133.write('X22', 'GEO', body)
    worksheet133.write('Y22', 'EKO', body)
    worksheet133.write('Z22', 'SOS', body)
    worksheet133.write('AA22', 'FIS', body)
    worksheet133.write('AB22', 'KIM', body)
    worksheet133.write('AC22', 'BIO', body)
    worksheet133.write('AD22', 'JML', body)

    worksheet133.conditional_format(22, 0, row133+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 134
    worksheet134.insert_image('A1', r'logo resmi nf.jpg')

    worksheet134.set_column('A:A', 7, center)
    worksheet134.set_column('B:B', 6, center)
    worksheet134.set_column('C:C', 18.14, center)
    worksheet134.set_column('D:D', 25, left)
    worksheet134.set_column('E:E', 13.14, left)
    worksheet134.set_column('F:F', 8.57, center)
    worksheet134.set_column('G:AD', 5, center)
    worksheet134.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIHANJUANG', title)
    worksheet134.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet134.write('A5', 'LOKASI', header)
    worksheet134.write('B5', 'TOTAL', header)
    worksheet134.merge_range('A4:B4', 'RANK', header)
    worksheet134.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet134.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet134.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet134.merge_range('F4:F5', 'KELAS', header)
    worksheet134.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet134.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet134.write('G5', 'MAW', body)
    worksheet134.write('H5', 'MAP', body)
    worksheet134.write('I5', 'IND', body)
    worksheet134.write('J5', 'ENG', body)
    worksheet134.write('K5', 'SEJ', body)
    worksheet134.write('L5', 'GEO', body)
    worksheet134.write('M5', 'EKO', body)
    worksheet134.write('N5', 'SOS', body)
    worksheet134.write('O5', 'FIS', body)
    worksheet134.write('P5', 'KIM', body)
    worksheet134.write('Q5', 'BIO', body)
    worksheet134.write('R5', 'JML', body)
    worksheet134.write('S5', 'MAW', body)
    worksheet134.write('T5', 'MAP', body)
    worksheet134.write('U5', 'IND', body)
    worksheet134.write('V5', 'ENG', body)
    worksheet134.write('W5', 'SEJ', body)
    worksheet134.write('X5', 'GEO', body)
    worksheet134.write('Y5', 'EKO', body)
    worksheet134.write('Z5', 'SOS', body)
    worksheet134.write('AA5', 'FIS', body)
    worksheet134.write('AB5', 'KIM', body)
    worksheet134.write('AC5', 'BIO', body)
    worksheet134.write('AD5', 'JML', body)

    worksheet134.conditional_format(5, 0, row134_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet134.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF CIHANJUANG', title)
    worksheet134.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet134.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet134.write('A22', 'LOKASI', header)
    worksheet134.write('B22', 'TOTAL', header)
    worksheet134.merge_range('A21:B21', 'RANK', header)
    worksheet134.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet134.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet134.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet134.merge_range('F21:F22', 'KELAS', header)
    worksheet134.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet134.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet134.write('G22', 'MAW', body)
    worksheet134.write('H22', 'MAP', body)
    worksheet134.write('I22', 'IND', body)
    worksheet134.write('J22', 'ENG', body)
    worksheet134.write('K22', 'SEJ', body)
    worksheet134.write('L22', 'GEO', body)
    worksheet134.write('M22', 'EKO', body)
    worksheet134.write('N22', 'SOS', body)
    worksheet134.write('O22', 'FIS', body)
    worksheet134.write('P22', 'KIM', body)
    worksheet134.write('Q22', 'BIO', body)
    worksheet134.write('R22', 'JML', body)
    worksheet134.write('S22', 'MAW', body)
    worksheet134.write('T22', 'MAP', body)
    worksheet134.write('U22', 'IND', body)
    worksheet134.write('V22', 'ENG', body)
    worksheet134.write('W22', 'SEJ', body)
    worksheet134.write('X22', 'GEO', body)
    worksheet134.write('Y22', 'EKO', body)
    worksheet134.write('Z22', 'SOS', body)
    worksheet134.write('AA22', 'FIS', body)
    worksheet134.write('AB22', 'KIM', body)
    worksheet134.write('AC22', 'BIO', body)
    worksheet134.write('AD22', 'JML', body)

    worksheet134.conditional_format(22, 0, row134+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 135
    worksheet135.insert_image('A1', r'logo resmi nf.jpg')

    worksheet135.set_column('A:A', 7, center)
    worksheet135.set_column('B:B', 6, center)
    worksheet135.set_column('C:C', 18.14, center)
    worksheet135.set_column('D:D', 25, left)
    worksheet135.set_column('E:E', 13.14, left)
    worksheet135.set_column('F:F', 8.57, center)
    worksheet135.set_column('G:AD', 5, center)
    worksheet135.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BUAH BATU', title)
    worksheet135.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet135.write('A5', 'LOKASI', header)
    worksheet135.write('B5', 'TOTAL', header)
    worksheet135.merge_range('A4:B4', 'RANK', header)
    worksheet135.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet135.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet135.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet135.merge_range('F4:F5', 'KELAS', header)
    worksheet135.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet135.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet135.write('G5', 'MAW', body)
    worksheet135.write('H5', 'MAP', body)
    worksheet135.write('I5', 'IND', body)
    worksheet135.write('J5', 'ENG', body)
    worksheet135.write('K5', 'SEJ', body)
    worksheet135.write('L5', 'GEO', body)
    worksheet135.write('M5', 'EKO', body)
    worksheet135.write('N5', 'SOS', body)
    worksheet135.write('O5', 'FIS', body)
    worksheet135.write('P5', 'KIM', body)
    worksheet135.write('Q5', 'BIO', body)
    worksheet135.write('R5', 'JML', body)
    worksheet135.write('S5', 'MAW', body)
    worksheet135.write('T5', 'MAP', body)
    worksheet135.write('U5', 'IND', body)
    worksheet135.write('V5', 'ENG', body)
    worksheet135.write('W5', 'SEJ', body)
    worksheet135.write('X5', 'GEO', body)
    worksheet135.write('Y5', 'EKO', body)
    worksheet135.write('Z5', 'SOS', body)
    worksheet135.write('AA5', 'FIS', body)
    worksheet135.write('AB5', 'KIM', body)
    worksheet135.write('AC5', 'BIO', body)
    worksheet135.write('AD5', 'JML', body)

    worksheet135.conditional_format(5, 0, row135_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet135.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF BUAH BATU', title)
    worksheet135.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet135.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet135.write('A22', 'LOKASI', header)
    worksheet135.write('B22', 'TOTAL', header)
    worksheet135.merge_range('A21:B21', 'RANK', header)
    worksheet135.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet135.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet135.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet135.merge_range('F21:F22', 'KELAS', header)
    worksheet135.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet135.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet135.write('G22', 'MAW', body)
    worksheet135.write('H22', 'MAP', body)
    worksheet135.write('I22', 'IND', body)
    worksheet135.write('J22', 'ENG', body)
    worksheet135.write('K22', 'SEJ', body)
    worksheet135.write('L22', 'GEO', body)
    worksheet135.write('M22', 'EKO', body)
    worksheet135.write('N22', 'SOS', body)
    worksheet135.write('O22', 'FIS', body)
    worksheet135.write('P22', 'KIM', body)
    worksheet135.write('Q22', 'BIO', body)
    worksheet135.write('R22', 'JML', body)
    worksheet135.write('S22', 'MAW', body)
    worksheet135.write('T22', 'MAP', body)
    worksheet135.write('U22', 'IND', body)
    worksheet135.write('V22', 'ENG', body)
    worksheet135.write('W22', 'SEJ', body)
    worksheet135.write('X22', 'GEO', body)
    worksheet135.write('Y22', 'EKO', body)
    worksheet135.write('Z22', 'SOS', body)
    worksheet135.write('AA22', 'FIS', body)
    worksheet135.write('AB22', 'KIM', body)
    worksheet135.write('AC22', 'BIO', body)
    worksheet135.write('AD22', 'JML', body)

    worksheet135.conditional_format(22, 0, row135+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 136
    worksheet136.insert_image('A1', r'logo resmi nf.jpg')

    worksheet136.set_column('A:A', 7, center)
    worksheet136.set_column('B:B', 6, center)
    worksheet136.set_column('C:C', 18.14, center)
    worksheet136.set_column('D:D', 25, left)
    worksheet136.set_column('E:E', 13.14, left)
    worksheet136.set_column('F:F', 8.57, center)
    worksheet136.set_column('G:AD', 5, center)
    worksheet136.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SUMBAWA', title)
    worksheet136.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet136.write('A5', 'LOKASI', header)
    worksheet136.write('B5', 'TOTAL', header)
    worksheet136.merge_range('A4:B4', 'RANK', header)
    worksheet136.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet136.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet136.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet136.merge_range('F4:F5', 'KELAS', header)
    worksheet136.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet136.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet136.write('G5', 'MAW', body)
    worksheet136.write('H5', 'MAP', body)
    worksheet136.write('I5', 'IND', body)
    worksheet136.write('J5', 'ENG', body)
    worksheet136.write('K5', 'SEJ', body)
    worksheet136.write('L5', 'GEO', body)
    worksheet136.write('M5', 'EKO', body)
    worksheet136.write('N5', 'SOS', body)
    worksheet136.write('O5', 'FIS', body)
    worksheet136.write('P5', 'KIM', body)
    worksheet136.write('Q5', 'BIO', body)
    worksheet136.write('R5', 'JML', body)
    worksheet136.write('S5', 'MAW', body)
    worksheet136.write('T5', 'MAP', body)
    worksheet136.write('U5', 'IND', body)
    worksheet136.write('V5', 'ENG', body)
    worksheet136.write('W5', 'SEJ', body)
    worksheet136.write('X5', 'GEO', body)
    worksheet136.write('Y5', 'EKO', body)
    worksheet136.write('Z5', 'SOS', body)
    worksheet136.write('AA5', 'FIS', body)
    worksheet136.write('AB5', 'KIM', body)
    worksheet136.write('AC5', 'BIO', body)
    worksheet136.write('AD5', 'JML', body)

    worksheet136.conditional_format(5, 0, row136_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet136.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF SUMBAWA', title)
    worksheet136.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet136.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet136.write('A22', 'LOKASI', header)
    worksheet136.write('B22', 'TOTAL', header)
    worksheet136.merge_range('A21:B21', 'RANK', header)
    worksheet136.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet136.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet136.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet136.merge_range('F21:F22', 'KELAS', header)
    worksheet136.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet136.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet136.write('G22', 'MAW', body)
    worksheet136.write('H22', 'MAP', body)
    worksheet136.write('I22', 'IND', body)
    worksheet136.write('J22', 'ENG', body)
    worksheet136.write('K22', 'SEJ', body)
    worksheet136.write('L22', 'GEO', body)
    worksheet136.write('M22', 'EKO', body)
    worksheet136.write('N22', 'SOS', body)
    worksheet136.write('O22', 'FIS', body)
    worksheet136.write('P22', 'KIM', body)
    worksheet136.write('Q22', 'BIO', body)
    worksheet136.write('R22', 'JML', body)
    worksheet136.write('S22', 'MAW', body)
    worksheet136.write('T22', 'MAP', body)
    worksheet136.write('U22', 'IND', body)
    worksheet136.write('V22', 'ENG', body)
    worksheet136.write('W22', 'SEJ', body)
    worksheet136.write('X22', 'GEO', body)
    worksheet136.write('Y22', 'EKO', body)
    worksheet136.write('Z22', 'SOS', body)
    worksheet136.write('AA22', 'FIS', body)
    worksheet136.write('AB22', 'KIM', body)
    worksheet136.write('AC22', 'BIO', body)
    worksheet136.write('AD22', 'JML', body)

    worksheet136.conditional_format(22, 0, row136+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 137
    worksheet137.insert_image('A1', r'logo resmi nf.jpg')

    worksheet137.set_column('A:A', 7, center)
    worksheet137.set_column('B:B', 6, center)
    worksheet137.set_column('C:C', 18.14, center)
    worksheet137.set_column('D:D', 25, left)
    worksheet137.set_column('E:E', 13.14, left)
    worksheet137.set_column('F:F', 8.57, center)
    worksheet137.set_column('G:AD', 5, center)
    worksheet137.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF UJUNG BERUNG', title)
    worksheet137.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet137.write('A5', 'LOKASI', header)
    worksheet137.write('B5', 'TOTAL', header)
    worksheet137.merge_range('A4:B4', 'RANK', header)
    worksheet137.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet137.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet137.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet137.merge_range('F4:F5', 'KELAS', header)
    worksheet137.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet137.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet137.write('G5', 'MAW', body)
    worksheet137.write('H5', 'MAP', body)
    worksheet137.write('I5', 'IND', body)
    worksheet137.write('J5', 'ENG', body)
    worksheet137.write('K5', 'SEJ', body)
    worksheet137.write('L5', 'GEO', body)
    worksheet137.write('M5', 'EKO', body)
    worksheet137.write('N5', 'SOS', body)
    worksheet137.write('O5', 'FIS', body)
    worksheet137.write('P5', 'KIM', body)
    worksheet137.write('Q5', 'BIO', body)
    worksheet137.write('R5', 'JML', body)
    worksheet137.write('S5', 'MAW', body)
    worksheet137.write('T5', 'MAP', body)
    worksheet137.write('U5', 'IND', body)
    worksheet137.write('V5', 'ENG', body)
    worksheet137.write('W5', 'SEJ', body)
    worksheet137.write('X5', 'GEO', body)
    worksheet137.write('Y5', 'EKO', body)
    worksheet137.write('Z5', 'SOS', body)
    worksheet137.write('AA5', 'FIS', body)
    worksheet137.write('AB5', 'KIM', body)
    worksheet137.write('AC5', 'BIO', body)
    worksheet137.write('AD5', 'JML', body)

    worksheet137.conditional_format(5, 0, row137_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet137.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF UJUNG BERUNG', title)
    worksheet137.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet137.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet137.write('A22', 'LOKASI', header)
    worksheet137.write('B22', 'TOTAL', header)
    worksheet137.merge_range('A21:B21', 'RANK', header)
    worksheet137.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet137.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet137.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet137.merge_range('F21:F22', 'KELAS', header)
    worksheet137.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet137.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet137.write('G22', 'MAW', body)
    worksheet137.write('H22', 'MAP', body)
    worksheet137.write('I22', 'IND', body)
    worksheet137.write('J22', 'ENG', body)
    worksheet137.write('K22', 'SEJ', body)
    worksheet137.write('L22', 'GEO', body)
    worksheet137.write('M22', 'EKO', body)
    worksheet137.write('N22', 'SOS', body)
    worksheet137.write('O22', 'FIS', body)
    worksheet137.write('P22', 'KIM', body)
    worksheet137.write('Q22', 'BIO', body)
    worksheet137.write('R22', 'JML', body)
    worksheet137.write('S22', 'MAW', body)
    worksheet137.write('T22', 'MAP', body)
    worksheet137.write('U22', 'IND', body)
    worksheet137.write('V22', 'ENG', body)
    worksheet137.write('W22', 'SEJ', body)
    worksheet137.write('X22', 'GEO', body)
    worksheet137.write('Y22', 'EKO', body)
    worksheet137.write('Z22', 'SOS', body)
    worksheet137.write('AA22', 'FIS', body)
    worksheet137.write('AB22', 'KIM', body)
    worksheet137.write('AC22', 'BIO', body)
    worksheet137.write('AD22', 'JML', body)

    worksheet137.conditional_format(22, 0, row137+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 138
    worksheet138.insert_image('A1', r'logo resmi nf.jpg')

    worksheet138.set_column('A:A', 7, center)
    worksheet138.set_column('B:B', 6, center)
    worksheet138.set_column('C:C', 18.14, center)
    worksheet138.set_column('D:D', 25, left)
    worksheet138.set_column('E:E', 13.14, left)
    worksheet138.set_column('F:F', 8.57, center)
    worksheet138.set_column('G:AD', 5, center)
    worksheet138.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SANGKURIANG', title)
    worksheet138.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet138.write('A5', 'LOKASI', header)
    worksheet138.write('B5', 'TOTAL', header)
    worksheet138.merge_range('A4:B4', 'RANK', header)
    worksheet138.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet138.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet138.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet138.merge_range('F4:F5', 'KELAS', header)
    worksheet138.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet138.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet138.write('G5', 'MAW', body)
    worksheet138.write('H5', 'MAP', body)
    worksheet138.write('I5', 'IND', body)
    worksheet138.write('J5', 'ENG', body)
    worksheet138.write('K5', 'SEJ', body)
    worksheet138.write('L5', 'GEO', body)
    worksheet138.write('M5', 'EKO', body)
    worksheet138.write('N5', 'SOS', body)
    worksheet138.write('O5', 'FIS', body)
    worksheet138.write('P5', 'KIM', body)
    worksheet138.write('Q5', 'BIO', body)
    worksheet138.write('R5', 'JML', body)
    worksheet138.write('S5', 'MAW', body)
    worksheet138.write('T5', 'MAP', body)
    worksheet138.write('U5', 'IND', body)
    worksheet138.write('V5', 'ENG', body)
    worksheet138.write('W5', 'SEJ', body)
    worksheet138.write('X5', 'GEO', body)
    worksheet138.write('Y5', 'EKO', body)
    worksheet138.write('Z5', 'SOS', body)
    worksheet138.write('AA5', 'FIS', body)
    worksheet138.write('AB5', 'KIM', body)
    worksheet138.write('AC5', 'BIO', body)
    worksheet138.write('AD5', 'JML', body)

    worksheet138.conditional_format(5, 0, row138_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet138.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF SANGKURIANG', title)
    worksheet138.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet138.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet138.write('A22', 'LOKASI', header)
    worksheet138.write('B22', 'TOTAL', header)
    worksheet138.merge_range('A21:B21', 'RANK', header)
    worksheet138.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet138.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet138.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet138.merge_range('F21:F22', 'KELAS', header)
    worksheet138.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet138.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet138.write('G22', 'MAW', body)
    worksheet138.write('H22', 'MAP', body)
    worksheet138.write('I22', 'IND', body)
    worksheet138.write('J22', 'ENG', body)
    worksheet138.write('K22', 'SEJ', body)
    worksheet138.write('L22', 'GEO', body)
    worksheet138.write('M22', 'EKO', body)
    worksheet138.write('N22', 'SOS', body)
    worksheet138.write('O22', 'FIS', body)
    worksheet138.write('P22', 'KIM', body)
    worksheet138.write('Q22', 'BIO', body)
    worksheet138.write('R22', 'JML', body)
    worksheet138.write('S22', 'MAW', body)
    worksheet138.write('T22', 'MAP', body)
    worksheet138.write('U22', 'IND', body)
    worksheet138.write('V22', 'ENG', body)
    worksheet138.write('W22', 'SEJ', body)
    worksheet138.write('X22', 'GEO', body)
    worksheet138.write('Y22', 'EKO', body)
    worksheet138.write('Z22', 'SOS', body)
    worksheet138.write('AA22', 'FIS', body)
    worksheet138.write('AB22', 'KIM', body)
    worksheet138.write('AC22', 'BIO', body)
    worksheet138.write('AD22', 'JML', body)

    worksheet138.conditional_format(22, 0, row138+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # # worksheet 139
    # worksheet139.insert_image('A1',r'logo resmi nf.jpg')

    # worksheet139.set_column('A:A', 7, center)
    # worksheet139.set_column('B:B', 6, center)
    # worksheet139.set_column('C:C', 18.14, center)
    # worksheet139.set_column('D:D', 25, left)
    # worksheet139.set_column('E:E', 13.14, left)
    # worksheet139.set_column('F:F', 8.57, center)
    # worksheet139.set_column('G:AD', 5, center)
    # worksheet139.merge_range('A1:V1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SARIJADI', title)
    # worksheet139.merge_range('A2:V2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    # worksheet139.write('A5', 'LOKASI', header)
    # worksheet139.write('B5', 'TOTAL', header)
    # worksheet139.merge_range('A4:B4', 'RANK', header)
    # worksheet139.merge_range('C4:C5', 'NOMOR NF', header)
    # worksheet139.merge_range('D4:D5', 'NAMA SISWA', header)
    # worksheet139.merge_range('E4:E5', 'SEKOLAH', header)
    # worksheet139.merge_range('F4:F5', 'KELAS', header)
    # worksheet139.merge_range('G4:R4', 'JUMLAH BENAR', header)
    # worksheet139.merge_range('S4:AD4', 'NILAI STANDAR', header)
    # worksheet139.write('G5', 'MAW', body)
    # worksheet139.write('H5', 'MAP', body)
    # worksheet139.write('I5', 'IND', body)
    # worksheet139.write('J5', 'ENG', body)
    # worksheet139.write('K5', 'SEJ', body)
    # worksheet139.write('L5', 'GEO', body)
    # worksheet139.write('M5', 'EKO', body)
    # worksheet139.write('N5', 'SOS', body)
    # worksheet139.write('O5', 'FIS', body)
    # worksheet139.write('P5', 'KIM', body)
    # worksheet139.write('Q5', 'BIO', body)
    # worksheet139.write('R5', 'JML', body)
    # worksheet139.write('S5', 'MAW', body)
    # worksheet139.write('T5', 'MAP', body)
    # worksheet139.write('U5', 'IND', body)
    # worksheet139.write('V5', 'ENG', body)
    # worksheet139.write('W5', 'SEJ', body)
    # worksheet139.write('X5', 'GEO', body)
    # worksheet139.write('Y5', 'EKO', body)
    # worksheet139.write('Z5', 'SOS', body)
    # worksheet139.write('AA5', 'FIS', body)
    # worksheet139.write('AB5', 'KIM', body)
    # worksheet139.write('AC5', 'BIO', body)
    # worksheet139.write('AD5', 'JML', body)

    # worksheet139.conditional_format(5,0,row139_10+4,21,
    #                              {'type': 'no_errors', 'format': border})

    # worksheet139.merge_range('A17:V17', fr'KELAS {kelas} - LOKASI NF SARIJADI', title)
    # worksheet139.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    # worksheet139.merge_range('A19:V19', fr'{semester} TAHUN {tahun}', sub_title)
    # worksheet139.write('A22', 'LOKASI', header)
    # worksheet139.write('B22', 'TOTAL', header)
    # worksheet139.merge_range('A21:B21', 'RANK', header)
    # worksheet139.merge_range('C21:C22', 'NOMOR NF', header)
    # worksheet139.merge_range('D21:D22', 'NAMA SISWA', header)
    # worksheet139.merge_range('E21:E22', 'SEKOLAH', header)
    # worksheet139.merge_range('F21:F22', 'KELAS', header)
    # worksheet139.merge_range('G21:N21', 'JUMLAH BENAR', header)
    # worksheet139.merge_range('O21:V21', 'NILAI STANDAR', header)
    # worksheet139.write('G22', 'MAT', body)
    # worksheet139.write('H22', 'IND', body)
    # worksheet139.write('I22', 'ENG', body)
    # worksheet139.write('J22', 'SEJ', body)
    # worksheet139.write('K22', 'GEO', body)
    # worksheet139.write('L22', 'SOS', body)
    # worksheet139.write('M22', 'EKO', body)
    # worksheet139.write('N22', 'JML', body)
    # worksheet139.write('O22', 'MAT', body)
    # worksheet139.write('P22', 'IND', body)
    # worksheet139.write('Q22', 'ENG', body)
    # worksheet139.write('R22', 'SEJ', body)
    # worksheet139.write('S22', 'GEO', body)
    # worksheet139.write('T22', 'SOS', body)
    # worksheet139.write('U22', 'EKO', body)
    # worksheet139.write('V22', 'JML', body)

    # worksheet139.conditional_format(22,0,row139+21,21,
    #                              {'type': 'no_errors', 'format': border})

    # worksheet 140
    worksheet140.insert_image('A1', r'logo resmi nf.jpg')

    worksheet140.set_column('A:A', 7, center)
    worksheet140.set_column('B:B', 6, center)
    worksheet140.set_column('C:C', 18.14, center)
    worksheet140.set_column('D:D', 25, left)
    worksheet140.set_column('E:E', 13.14, left)
    worksheet140.set_column('F:F', 8.57, center)
    worksheet140.set_column('G:AD', 5, center)
    worksheet140.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KARAWACI', title)
    worksheet140.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet140.write('A5', 'LOKASI', header)
    worksheet140.write('B5', 'TOTAL', header)
    worksheet140.merge_range('A4:B4', 'RANK', header)
    worksheet140.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet140.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet140.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet140.merge_range('F4:F5', 'KELAS', header)
    worksheet140.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet140.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet140.write('G5', 'MAW', body)
    worksheet140.write('H5', 'MAP', body)
    worksheet140.write('I5', 'IND', body)
    worksheet140.write('J5', 'ENG', body)
    worksheet140.write('K5', 'SEJ', body)
    worksheet140.write('L5', 'GEO', body)
    worksheet140.write('M5', 'EKO', body)
    worksheet140.write('N5', 'SOS', body)
    worksheet140.write('O5', 'FIS', body)
    worksheet140.write('P5', 'KIM', body)
    worksheet140.write('Q5', 'BIO', body)
    worksheet140.write('R5', 'JML', body)
    worksheet140.write('S5', 'MAW', body)
    worksheet140.write('T5', 'MAP', body)
    worksheet140.write('U5', 'IND', body)
    worksheet140.write('V5', 'ENG', body)
    worksheet140.write('W5', 'SEJ', body)
    worksheet140.write('X5', 'GEO', body)
    worksheet140.write('Y5', 'EKO', body)
    worksheet140.write('Z5', 'SOS', body)
    worksheet140.write('AA5', 'FIS', body)
    worksheet140.write('AB5', 'KIM', body)
    worksheet140.write('AC5', 'BIO', body)
    worksheet140.write('AD5', 'JML', body)

    worksheet140.conditional_format(5, 0, row140_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet140.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF KARAWACI', title)
    worksheet140.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet140.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet140.write('A22', 'LOKASI', header)
    worksheet140.write('B22', 'TOTAL', header)
    worksheet140.merge_range('A21:B21', 'RANK', header)
    worksheet140.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet140.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet140.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet140.merge_range('F21:F22', 'KELAS', header)
    worksheet140.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet140.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet140.write('G22', 'MAW', body)
    worksheet140.write('H22', 'MAP', body)
    worksheet140.write('I22', 'IND', body)
    worksheet140.write('J22', 'ENG', body)
    worksheet140.write('K22', 'SEJ', body)
    worksheet140.write('L22', 'GEO', body)
    worksheet140.write('M22', 'EKO', body)
    worksheet140.write('N22', 'SOS', body)
    worksheet140.write('O22', 'FIS', body)
    worksheet140.write('P22', 'KIM', body)
    worksheet140.write('Q22', 'BIO', body)
    worksheet140.write('R22', 'JML', body)
    worksheet140.write('S22', 'MAW', body)
    worksheet140.write('T22', 'MAP', body)
    worksheet140.write('U22', 'IND', body)
    worksheet140.write('V22', 'ENG', body)
    worksheet140.write('W22', 'SEJ', body)
    worksheet140.write('X22', 'GEO', body)
    worksheet140.write('Y22', 'EKO', body)
    worksheet140.write('Z22', 'SOS', body)
    worksheet140.write('AA22', 'FIS', body)
    worksheet140.write('AB22', 'KIM', body)
    worksheet140.write('AC22', 'BIO', body)
    worksheet140.write('AD22', 'JML', body)

    worksheet140.conditional_format(22, 0, row140+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 141
    worksheet141.insert_image('A1', r'logo resmi nf.jpg')

    worksheet141.set_column('A:A', 7, center)
    worksheet141.set_column('B:B', 6, center)
    worksheet141.set_column('C:C', 18.14, center)
    worksheet141.set_column('D:D', 25, left)
    worksheet141.set_column('E:E', 13.14, left)
    worksheet141.set_column('F:F', 8.57, center)
    worksheet141.set_column('G:AD', 5, center)
    worksheet141.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF VETERAN TANGERANG', title)
    worksheet141.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet141.write('A5', 'LOKASI', header)
    worksheet141.write('B5', 'TOTAL', header)
    worksheet141.merge_range('A4:B4', 'RANK', header)
    worksheet141.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet141.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet141.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet141.merge_range('F4:F5', 'KELAS', header)
    worksheet141.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet141.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet141.write('G5', 'MAW', body)
    worksheet141.write('H5', 'MAP', body)
    worksheet141.write('I5', 'IND', body)
    worksheet141.write('J5', 'ENG', body)
    worksheet141.write('K5', 'SEJ', body)
    worksheet141.write('L5', 'GEO', body)
    worksheet141.write('M5', 'EKO', body)
    worksheet141.write('N5', 'SOS', body)
    worksheet141.write('O5', 'FIS', body)
    worksheet141.write('P5', 'KIM', body)
    worksheet141.write('Q5', 'BIO', body)
    worksheet141.write('R5', 'JML', body)
    worksheet141.write('S5', 'MAW', body)
    worksheet141.write('T5', 'MAP', body)
    worksheet141.write('U5', 'IND', body)
    worksheet141.write('V5', 'ENG', body)
    worksheet141.write('W5', 'SEJ', body)
    worksheet141.write('X5', 'GEO', body)
    worksheet141.write('Y5', 'EKO', body)
    worksheet141.write('Z5', 'SOS', body)
    worksheet141.write('AA5', 'FIS', body)
    worksheet141.write('AB5', 'KIM', body)
    worksheet141.write('AC5', 'BIO', body)
    worksheet141.write('AD5', 'JML', body)

    worksheet141.conditional_format(5, 0, row141_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet141.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF VETERAN TANGERANG', title)
    worksheet141.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet141.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet141.write('A22', 'LOKASI', header)
    worksheet141.write('B22', 'TOTAL', header)
    worksheet141.merge_range('A21:B21', 'RANK', header)
    worksheet141.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet141.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet141.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet141.merge_range('F21:F22', 'KELAS', header)
    worksheet141.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet141.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet141.write('G22', 'MAW', body)
    worksheet141.write('H22', 'MAP', body)
    worksheet141.write('I22', 'IND', body)
    worksheet141.write('J22', 'ENG', body)
    worksheet141.write('K22', 'SEJ', body)
    worksheet141.write('L22', 'GEO', body)
    worksheet141.write('M22', 'EKO', body)
    worksheet141.write('N22', 'SOS', body)
    worksheet141.write('O22', 'FIS', body)
    worksheet141.write('P22', 'KIM', body)
    worksheet141.write('Q22', 'BIO', body)
    worksheet141.write('R22', 'JML', body)
    worksheet141.write('S22', 'MAW', body)
    worksheet141.write('T22', 'MAP', body)
    worksheet141.write('U22', 'IND', body)
    worksheet141.write('V22', 'ENG', body)
    worksheet141.write('W22', 'SEJ', body)
    worksheet141.write('X22', 'GEO', body)
    worksheet141.write('Y22', 'EKO', body)
    worksheet141.write('Z22', 'SOS', body)
    worksheet141.write('AA22', 'FIS', body)
    worksheet141.write('AB22', 'KIM', body)
    worksheet141.write('AC22', 'BIO', body)
    worksheet141.write('AD22', 'JML', body)

    worksheet141.conditional_format(22, 0, row141+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 142
    worksheet142.insert_image('A1', r'logo resmi nf.jpg')

    worksheet142.set_column('A:A', 7, center)
    worksheet142.set_column('B:B', 6, center)
    worksheet142.set_column('C:C', 18.14, center)
    worksheet142.set_column('D:D', 25, left)
    worksheet142.set_column('E:E', 13.14, left)
    worksheet142.set_column('F:F', 8.57, center)
    worksheet142.set_column('G:AD', 5, center)
    worksheet142.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PERUMNAS 2 TANGERANG', title)
    worksheet142.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet142.write('A5', 'LOKASI', header)
    worksheet142.write('B5', 'TOTAL', header)
    worksheet142.merge_range('A4:B4', 'RANK', header)
    worksheet142.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet142.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet142.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet142.merge_range('F4:F5', 'KELAS', header)
    worksheet142.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet142.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet142.write('G5', 'MAW', body)
    worksheet142.write('H5', 'MAP', body)
    worksheet142.write('I5', 'IND', body)
    worksheet142.write('J5', 'ENG', body)
    worksheet142.write('K5', 'SEJ', body)
    worksheet142.write('L5', 'GEO', body)
    worksheet142.write('M5', 'EKO', body)
    worksheet142.write('N5', 'SOS', body)
    worksheet142.write('O5', 'FIS', body)
    worksheet142.write('P5', 'KIM', body)
    worksheet142.write('Q5', 'BIO', body)
    worksheet142.write('R5', 'JML', body)
    worksheet142.write('S5', 'MAW', body)
    worksheet142.write('T5', 'MAP', body)
    worksheet142.write('U5', 'IND', body)
    worksheet142.write('V5', 'ENG', body)
    worksheet142.write('W5', 'SEJ', body)
    worksheet142.write('X5', 'GEO', body)
    worksheet142.write('Y5', 'EKO', body)
    worksheet142.write('Z5', 'SOS', body)
    worksheet142.write('AA5', 'FIS', body)
    worksheet142.write('AB5', 'KIM', body)
    worksheet142.write('AC5', 'BIO', body)
    worksheet142.write('AD5', 'JML', body)

    worksheet142.conditional_format(5, 0, row142_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet142.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF PERUMNAS 2 TANGERANG', title)
    worksheet142.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet142.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet142.write('A22', 'LOKASI', header)
    worksheet142.write('B22', 'TOTAL', header)
    worksheet142.merge_range('A21:B21', 'RANK', header)
    worksheet142.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet142.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet142.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet142.merge_range('F21:F22', 'KELAS', header)
    worksheet142.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet142.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet142.write('G22', 'MAW', body)
    worksheet142.write('H22', 'MAP', body)
    worksheet142.write('I22', 'IND', body)
    worksheet142.write('J22', 'ENG', body)
    worksheet142.write('K22', 'SEJ', body)
    worksheet142.write('L22', 'GEO', body)
    worksheet142.write('M22', 'EKO', body)
    worksheet142.write('N22', 'SOS', body)
    worksheet142.write('O22', 'FIS', body)
    worksheet142.write('P22', 'KIM', body)
    worksheet142.write('Q22', 'BIO', body)
    worksheet142.write('R22', 'JML', body)
    worksheet142.write('S22', 'MAW', body)
    worksheet142.write('T22', 'MAP', body)
    worksheet142.write('U22', 'IND', body)
    worksheet142.write('V22', 'ENG', body)
    worksheet142.write('W22', 'SEJ', body)
    worksheet142.write('X22', 'GEO', body)
    worksheet142.write('Y22', 'EKO', body)
    worksheet142.write('Z22', 'SOS', body)
    worksheet142.write('AA22', 'FIS', body)
    worksheet142.write('AB22', 'KIM', body)
    worksheet142.write('AC22', 'BIO', body)
    worksheet142.write('AD22', 'JML', body)

    worksheet142.conditional_format(22, 0, row142+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 143
    worksheet143.insert_image('A1', r'logo resmi nf.jpg')

    worksheet143.set_column('A:A', 7, center)
    worksheet143.set_column('B:B', 6, center)
    worksheet143.set_column('C:C', 18.14, center)
    worksheet143.set_column('D:D', 25, left)
    worksheet143.set_column('E:E', 13.14, left)
    worksheet143.set_column('F:F', 8.57, center)
    worksheet143.set_column('G:AD', 5, center)
    worksheet143.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KAYURINGIN', title)
    worksheet143.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet143.write('A5', 'LOKASI', header)
    worksheet143.write('B5', 'TOTAL', header)
    worksheet143.merge_range('A4:B4', 'RANK', header)
    worksheet143.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet143.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet143.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet143.merge_range('F4:F5', 'KELAS', header)
    worksheet143.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet143.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet143.write('G5', 'MAW', body)
    worksheet143.write('H5', 'MAP', body)
    worksheet143.write('I5', 'IND', body)
    worksheet143.write('J5', 'ENG', body)
    worksheet143.write('K5', 'SEJ', body)
    worksheet143.write('L5', 'GEO', body)
    worksheet143.write('M5', 'EKO', body)
    worksheet143.write('N5', 'SOS', body)
    worksheet143.write('O5', 'FIS', body)
    worksheet143.write('P5', 'KIM', body)
    worksheet143.write('Q5', 'BIO', body)
    worksheet143.write('R5', 'JML', body)
    worksheet143.write('S5', 'MAW', body)
    worksheet143.write('T5', 'MAP', body)
    worksheet143.write('U5', 'IND', body)
    worksheet143.write('V5', 'ENG', body)
    worksheet143.write('W5', 'SEJ', body)
    worksheet143.write('X5', 'GEO', body)
    worksheet143.write('Y5', 'EKO', body)
    worksheet143.write('Z5', 'SOS', body)
    worksheet143.write('AA5', 'FIS', body)
    worksheet143.write('AB5', 'KIM', body)
    worksheet143.write('AC5', 'BIO', body)
    worksheet143.write('AD5', 'JML', body)

    worksheet143.conditional_format(5, 0, row143_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet143.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF KAYURINGIN', title)
    worksheet143.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet143.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet143.write('A22', 'LOKASI', header)
    worksheet143.write('B22', 'TOTAL', header)
    worksheet143.merge_range('A21:B21', 'RANK', header)
    worksheet143.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet143.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet143.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet143.merge_range('F21:F22', 'KELAS', header)
    worksheet143.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet143.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet143.write('G22', 'MAW', body)
    worksheet143.write('H22', 'MAP', body)
    worksheet143.write('I22', 'IND', body)
    worksheet143.write('J22', 'ENG', body)
    worksheet143.write('K22', 'SEJ', body)
    worksheet143.write('L22', 'GEO', body)
    worksheet143.write('M22', 'EKO', body)
    worksheet143.write('N22', 'SOS', body)
    worksheet143.write('O22', 'FIS', body)
    worksheet143.write('P22', 'KIM', body)
    worksheet143.write('Q22', 'BIO', body)
    worksheet143.write('R22', 'JML', body)
    worksheet143.write('S22', 'MAW', body)
    worksheet143.write('T22', 'MAP', body)
    worksheet143.write('U22', 'IND', body)
    worksheet143.write('V22', 'ENG', body)
    worksheet143.write('W22', 'SEJ', body)
    worksheet143.write('X22', 'GEO', body)
    worksheet143.write('Y22', 'EKO', body)
    worksheet143.write('Z22', 'SOS', body)
    worksheet143.write('AA22', 'FIS', body)
    worksheet143.write('AB22', 'KIM', body)
    worksheet143.write('AC22', 'BIO', body)
    worksheet143.write('AD22', 'JML', body)

    worksheet143.conditional_format(22, 0, row143+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 144
    worksheet144.insert_image('A1', r'logo resmi nf.jpg')

    worksheet144.set_column('A:A', 7, center)
    worksheet144.set_column('B:B', 6, center)
    worksheet144.set_column('C:C', 18.14, center)
    worksheet144.set_column('D:D', 25, left)
    worksheet144.set_column('E:E', 13.14, left)
    worksheet144.set_column('F:F', 8.57, center)
    worksheet144.set_column('G:AD', 5, center)
    worksheet144.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF AGUS SALIM', title)
    worksheet144.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet144.write('A5', 'LOKASI', header)
    worksheet144.write('B5', 'TOTAL', header)
    worksheet144.merge_range('A4:B4', 'RANK', header)
    worksheet144.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet144.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet144.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet144.merge_range('F4:F5', 'KELAS', header)
    worksheet144.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet144.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet144.write('G5', 'MAW', body)
    worksheet144.write('H5', 'MAP', body)
    worksheet144.write('I5', 'IND', body)
    worksheet144.write('J5', 'ENG', body)
    worksheet144.write('K5', 'SEJ', body)
    worksheet144.write('L5', 'GEO', body)
    worksheet144.write('M5', 'EKO', body)
    worksheet144.write('N5', 'SOS', body)
    worksheet144.write('O5', 'FIS', body)
    worksheet144.write('P5', 'KIM', body)
    worksheet144.write('Q5', 'BIO', body)
    worksheet144.write('R5', 'JML', body)
    worksheet144.write('S5', 'MAW', body)
    worksheet144.write('T5', 'MAP', body)
    worksheet144.write('U5', 'IND', body)
    worksheet144.write('V5', 'ENG', body)
    worksheet144.write('W5', 'SEJ', body)
    worksheet144.write('X5', 'GEO', body)
    worksheet144.write('Y5', 'EKO', body)
    worksheet144.write('Z5', 'SOS', body)
    worksheet144.write('AA5', 'FIS', body)
    worksheet144.write('AB5', 'KIM', body)
    worksheet144.write('AC5', 'BIO', body)
    worksheet144.write('AD5', 'JML', body)

    worksheet144.conditional_format(5, 0, row144_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet144.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF AGUS SALIM', title)
    worksheet144.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet144.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet144.write('A22', 'LOKASI', header)
    worksheet144.write('B22', 'TOTAL', header)
    worksheet144.merge_range('A21:B21', 'RANK', header)
    worksheet144.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet144.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet144.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet144.merge_range('F21:F22', 'KELAS', header)
    worksheet144.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet144.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet144.write('G22', 'MAW', body)
    worksheet144.write('H22', 'MAP', body)
    worksheet144.write('I22', 'IND', body)
    worksheet144.write('J22', 'ENG', body)
    worksheet144.write('K22', 'SEJ', body)
    worksheet144.write('L22', 'GEO', body)
    worksheet144.write('M22', 'EKO', body)
    worksheet144.write('N22', 'SOS', body)
    worksheet144.write('O22', 'FIS', body)
    worksheet144.write('P22', 'KIM', body)
    worksheet144.write('Q22', 'BIO', body)
    worksheet144.write('R22', 'JML', body)
    worksheet144.write('S22', 'MAW', body)
    worksheet144.write('T22', 'MAP', body)
    worksheet144.write('U22', 'IND', body)
    worksheet144.write('V22', 'ENG', body)
    worksheet144.write('W22', 'SEJ', body)
    worksheet144.write('X22', 'GEO', body)
    worksheet144.write('Y22', 'EKO', body)
    worksheet144.write('Z22', 'SOS', body)
    worksheet144.write('AA22', 'FIS', body)
    worksheet144.write('AB22', 'KIM', body)
    worksheet144.write('AC22', 'BIO', body)
    worksheet144.write('AD22', 'JML', body)

    worksheet144.conditional_format(22, 0, row144+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 145
    worksheet145.insert_image('A1', r'logo resmi nf.jpg')

    worksheet145.set_column('A:A', 7, center)
    worksheet145.set_column('B:B', 6, center)
    worksheet145.set_column('C:C', 18.14, center)
    worksheet145.set_column('D:D', 25, left)
    worksheet145.set_column('E:E', 13.14, left)
    worksheet145.set_column('F:F', 8.57, center)
    worksheet145.set_column('G:AD', 5, center)
    worksheet145.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SUMERU', title)
    worksheet145.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet145.write('A5', 'LOKASI', header)
    worksheet145.write('B5', 'TOTAL', header)
    worksheet145.merge_range('A4:B4', 'RANK', header)
    worksheet145.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet145.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet145.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet145.merge_range('F4:F5', 'KELAS', header)
    worksheet145.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet145.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet145.write('G5', 'MAW', body)
    worksheet145.write('H5', 'MAP', body)
    worksheet145.write('I5', 'IND', body)
    worksheet145.write('J5', 'ENG', body)
    worksheet145.write('K5', 'SEJ', body)
    worksheet145.write('L5', 'GEO', body)
    worksheet145.write('M5', 'EKO', body)
    worksheet145.write('N5', 'SOS', body)
    worksheet145.write('O5', 'FIS', body)
    worksheet145.write('P5', 'KIM', body)
    worksheet145.write('Q5', 'BIO', body)
    worksheet145.write('R5', 'JML', body)
    worksheet145.write('S5', 'MAW', body)
    worksheet145.write('T5', 'MAP', body)
    worksheet145.write('U5', 'IND', body)
    worksheet145.write('V5', 'ENG', body)
    worksheet145.write('W5', 'SEJ', body)
    worksheet145.write('X5', 'GEO', body)
    worksheet145.write('Y5', 'EKO', body)
    worksheet145.write('Z5', 'SOS', body)
    worksheet145.write('AA5', 'FIS', body)
    worksheet145.write('AB5', 'KIM', body)
    worksheet145.write('AC5', 'BIO', body)
    worksheet145.write('AD5', 'JML', body)

    worksheet145.conditional_format(5, 0, row145_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet145.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF SUMERU', title)
    worksheet145.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet145.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet145.write('A22', 'LOKASI', header)
    worksheet145.write('B22', 'TOTAL', header)
    worksheet145.merge_range('A21:B21', 'RANK', header)
    worksheet145.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet145.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet145.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet145.merge_range('F21:F22', 'KELAS', header)
    worksheet145.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet145.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet145.write('G22', 'MAW', body)
    worksheet145.write('H22', 'MAP', body)
    worksheet145.write('I22', 'IND', body)
    worksheet145.write('J22', 'ENG', body)
    worksheet145.write('K22', 'SEJ', body)
    worksheet145.write('L22', 'GEO', body)
    worksheet145.write('M22', 'EKO', body)
    worksheet145.write('N22', 'SOS', body)
    worksheet145.write('O22', 'FIS', body)
    worksheet145.write('P22', 'KIM', body)
    worksheet145.write('Q22', 'BIO', body)
    worksheet145.write('R22', 'JML', body)
    worksheet145.write('S22', 'MAW', body)
    worksheet145.write('T22', 'MAP', body)
    worksheet145.write('U22', 'IND', body)
    worksheet145.write('V22', 'ENG', body)
    worksheet145.write('W22', 'SEJ', body)
    worksheet145.write('X22', 'GEO', body)
    worksheet145.write('Y22', 'EKO', body)
    worksheet145.write('Z22', 'SOS', body)
    worksheet145.write('AA22', 'FIS', body)
    worksheet145.write('AB22', 'KIM', body)
    worksheet145.write('AC22', 'BIO', body)
    worksheet145.write('AD22', 'JML', body)

    worksheet145.conditional_format(22, 0, row145+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 146
    worksheet146.insert_image('A1', r'logo resmi nf.jpg')

    worksheet146.set_column('A:A', 7, center)
    worksheet146.set_column('B:B', 6, center)
    worksheet146.set_column('C:C', 18.14, center)
    worksheet146.set_column('D:D', 25, left)
    worksheet146.set_column('E:E', 13.14, left)
    worksheet146.set_column('F:F', 8.57, center)
    worksheet146.set_column('G:AD', 5, center)
    worksheet146.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIKEAS', title)
    worksheet146.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet146.write('A5', 'LOKASI', header)
    worksheet146.write('B5', 'TOTAL', header)
    worksheet146.merge_range('A4:B4', 'RANK', header)
    worksheet146.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet146.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet146.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet146.merge_range('F4:F5', 'KELAS', header)
    worksheet146.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet146.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet146.write('G5', 'MAW', body)
    worksheet146.write('H5', 'MAP', body)
    worksheet146.write('I5', 'IND', body)
    worksheet146.write('J5', 'ENG', body)
    worksheet146.write('K5', 'SEJ', body)
    worksheet146.write('L5', 'GEO', body)
    worksheet146.write('M5', 'EKO', body)
    worksheet146.write('N5', 'SOS', body)
    worksheet146.write('O5', 'FIS', body)
    worksheet146.write('P5', 'KIM', body)
    worksheet146.write('Q5', 'BIO', body)
    worksheet146.write('R5', 'JML', body)
    worksheet146.write('S5', 'MAW', body)
    worksheet146.write('T5', 'MAP', body)
    worksheet146.write('U5', 'IND', body)
    worksheet146.write('V5', 'ENG', body)
    worksheet146.write('W5', 'SEJ', body)
    worksheet146.write('X5', 'GEO', body)
    worksheet146.write('Y5', 'EKO', body)
    worksheet146.write('Z5', 'SOS', body)
    worksheet146.write('AA5', 'FIS', body)
    worksheet146.write('AB5', 'KIM', body)
    worksheet146.write('AC5', 'BIO', body)
    worksheet146.write('AD5', 'JML', body)

    worksheet146.conditional_format(5, 0, row146_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet146.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF CIKEAS', title)
    worksheet146.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet146.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet146.write('A22', 'LOKASI', header)
    worksheet146.write('B22', 'TOTAL', header)
    worksheet146.merge_range('A21:B21', 'RANK', header)
    worksheet146.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet146.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet146.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet146.merge_range('F21:F22', 'KELAS', header)
    worksheet146.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet146.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet146.write('G22', 'MAW', body)
    worksheet146.write('H22', 'MAP', body)
    worksheet146.write('I22', 'IND', body)
    worksheet146.write('J22', 'ENG', body)
    worksheet146.write('K22', 'SEJ', body)
    worksheet146.write('L22', 'GEO', body)
    worksheet146.write('M22', 'EKO', body)
    worksheet146.write('N22', 'SOS', body)
    worksheet146.write('O22', 'FIS', body)
    worksheet146.write('P22', 'KIM', body)
    worksheet146.write('Q22', 'BIO', body)
    worksheet146.write('R22', 'JML', body)
    worksheet146.write('S22', 'MAW', body)
    worksheet146.write('T22', 'MAP', body)
    worksheet146.write('U22', 'IND', body)
    worksheet146.write('V22', 'ENG', body)
    worksheet146.write('W22', 'SEJ', body)
    worksheet146.write('X22', 'GEO', body)
    worksheet146.write('Y22', 'EKO', body)
    worksheet146.write('Z22', 'SOS', body)
    worksheet146.write('AA22', 'FIS', body)
    worksheet146.write('AB22', 'KIM', body)
    worksheet146.write('AC22', 'BIO', body)
    worksheet146.write('AD22', 'JML', body)

    worksheet146.conditional_format(22, 0, row146+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 148
    worksheet148.insert_image('A1', r'logo resmi nf.jpg')

    worksheet148.set_column('A:A', 7, center)
    worksheet148.set_column('B:B', 6, center)
    worksheet148.set_column('C:C', 18.14, center)
    worksheet148.set_column('D:D', 25, left)
    worksheet148.set_column('E:E', 13.14, left)
    worksheet148.set_column('F:F', 8.57, center)
    worksheet148.set_column('G:AD', 5, center)
    worksheet148.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIJAWA MASJID', title)
    worksheet148.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet148.write('A5', 'LOKASI', header)
    worksheet148.write('B5', 'TOTAL', header)
    worksheet148.merge_range('A4:B4', 'RANK', header)
    worksheet148.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet148.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet148.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet148.merge_range('F4:F5', 'KELAS', header)
    worksheet148.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet148.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet148.write('G5', 'MAW', body)
    worksheet148.write('H5', 'MAP', body)
    worksheet148.write('I5', 'IND', body)
    worksheet148.write('J5', 'ENG', body)
    worksheet148.write('K5', 'SEJ', body)
    worksheet148.write('L5', 'GEO', body)
    worksheet148.write('M5', 'EKO', body)
    worksheet148.write('N5', 'SOS', body)
    worksheet148.write('O5', 'FIS', body)
    worksheet148.write('P5', 'KIM', body)
    worksheet148.write('Q5', 'BIO', body)
    worksheet148.write('R5', 'JML', body)
    worksheet148.write('S5', 'MAW', body)
    worksheet148.write('T5', 'MAP', body)
    worksheet148.write('U5', 'IND', body)
    worksheet148.write('V5', 'ENG', body)
    worksheet148.write('W5', 'SEJ', body)
    worksheet148.write('X5', 'GEO', body)
    worksheet148.write('Y5', 'EKO', body)
    worksheet148.write('Z5', 'SOS', body)
    worksheet148.write('AA5', 'FIS', body)
    worksheet148.write('AB5', 'KIM', body)
    worksheet148.write('AC5', 'BIO', body)
    worksheet148.write('AD5', 'JML', body)

    worksheet148.conditional_format(5, 0, row148_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet148.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF CIJAWA MASJID', title)
    worksheet148.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet148.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet148.write('A22', 'LOKASI', header)
    worksheet148.write('B22', 'TOTAL', header)
    worksheet148.merge_range('A21:B21', 'RANK', header)
    worksheet148.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet148.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet148.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet148.merge_range('F21:F22', 'KELAS', header)
    worksheet148.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet148.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet148.write('G22', 'MAW', body)
    worksheet148.write('H22', 'MAP', body)
    worksheet148.write('I22', 'IND', body)
    worksheet148.write('J22', 'ENG', body)
    worksheet148.write('K22', 'SEJ', body)
    worksheet148.write('L22', 'GEO', body)
    worksheet148.write('M22', 'EKO', body)
    worksheet148.write('N22', 'SOS', body)
    worksheet148.write('O22', 'FIS', body)
    worksheet148.write('P22', 'KIM', body)
    worksheet148.write('Q22', 'BIO', body)
    worksheet148.write('R22', 'JML', body)
    worksheet148.write('S22', 'MAW', body)
    worksheet148.write('T22', 'MAP', body)
    worksheet148.write('U22', 'IND', body)
    worksheet148.write('V22', 'ENG', body)
    worksheet148.write('W22', 'SEJ', body)
    worksheet148.write('X22', 'GEO', body)
    worksheet148.write('Y22', 'EKO', body)
    worksheet148.write('Z22', 'SOS', body)
    worksheet148.write('AA22', 'FIS', body)
    worksheet148.write('AB22', 'KIM', body)
    worksheet148.write('AC22', 'BIO', body)
    worksheet148.write('AD22', 'JML', body)

    worksheet148.conditional_format(22, 0, row148+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 149
    worksheet149.insert_image('A1', r'logo resmi nf.jpg')

    worksheet149.set_column('A:A', 7, center)
    worksheet149.set_column('B:B', 6, center)
    worksheet149.set_column('C:C', 18.14, center)
    worksheet149.set_column('D:D', 25, left)
    worksheet149.set_column('E:E', 13.14, left)
    worksheet149.set_column('F:F', 8.57, center)
    worksheet149.set_column('G:AD', 5, center)
    worksheet149.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PALEDANG', title)
    worksheet149.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet149.write('A5', 'LOKASI', header)
    worksheet149.write('B5', 'TOTAL', header)
    worksheet149.merge_range('A4:B4', 'RANK', header)
    worksheet149.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet149.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet149.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet149.merge_range('F4:F5', 'KELAS', header)
    worksheet149.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet149.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet149.write('G5', 'MAW', body)
    worksheet149.write('H5', 'MAP', body)
    worksheet149.write('I5', 'IND', body)
    worksheet149.write('J5', 'ENG', body)
    worksheet149.write('K5', 'SEJ', body)
    worksheet149.write('L5', 'GEO', body)
    worksheet149.write('M5', 'EKO', body)
    worksheet149.write('N5', 'SOS', body)
    worksheet149.write('O5', 'FIS', body)
    worksheet149.write('P5', 'KIM', body)
    worksheet149.write('Q5', 'BIO', body)
    worksheet149.write('R5', 'JML', body)
    worksheet149.write('S5', 'MAW', body)
    worksheet149.write('T5', 'MAP', body)
    worksheet149.write('U5', 'IND', body)
    worksheet149.write('V5', 'ENG', body)
    worksheet149.write('W5', 'SEJ', body)
    worksheet149.write('X5', 'GEO', body)
    worksheet149.write('Y5', 'EKO', body)
    worksheet149.write('Z5', 'SOS', body)
    worksheet149.write('AA5', 'FIS', body)
    worksheet149.write('AB5', 'KIM', body)
    worksheet149.write('AC5', 'BIO', body)
    worksheet149.write('AD5', 'JML', body)

    worksheet149.conditional_format(5, 0, row149_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet149.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF PALEDANG', title)
    worksheet149.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet149.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet149.write('A22', 'LOKASI', header)
    worksheet149.write('B22', 'TOTAL', header)
    worksheet149.merge_range('A21:B21', 'RANK', header)
    worksheet149.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet149.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet149.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet149.merge_range('F21:F22', 'KELAS', header)
    worksheet149.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet149.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet149.write('G22', 'MAW', body)
    worksheet149.write('H22', 'MAP', body)
    worksheet149.write('I22', 'IND', body)
    worksheet149.write('J22', 'ENG', body)
    worksheet149.write('K22', 'SEJ', body)
    worksheet149.write('L22', 'GEO', body)
    worksheet149.write('M22', 'EKO', body)
    worksheet149.write('N22', 'SOS', body)
    worksheet149.write('O22', 'FIS', body)
    worksheet149.write('P22', 'KIM', body)
    worksheet149.write('Q22', 'BIO', body)
    worksheet149.write('R22', 'JML', body)
    worksheet149.write('S22', 'MAW', body)
    worksheet149.write('T22', 'MAP', body)
    worksheet149.write('U22', 'IND', body)
    worksheet149.write('V22', 'ENG', body)
    worksheet149.write('W22', 'SEJ', body)
    worksheet149.write('X22', 'GEO', body)
    worksheet149.write('Y22', 'EKO', body)
    worksheet149.write('Z22', 'SOS', body)
    worksheet149.write('AA22', 'FIS', body)
    worksheet149.write('AB22', 'KIM', body)
    worksheet149.write('AC22', 'BIO', body)
    worksheet149.write('AD22', 'JML', body)

    worksheet149.conditional_format(22, 0, row149+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 150
    worksheet150.insert_image('A1', r'logo resmi nf.jpg')

    worksheet150.set_column('A:A', 7, center)
    worksheet150.set_column('B:B', 6, center)
    worksheet150.set_column('C:C', 18.14, center)
    worksheet150.set_column('D:D', 25, left)
    worksheet150.set_column('E:E', 13.14, left)
    worksheet150.set_column('F:F', 8.57, center)
    worksheet150.set_column('G:AD', 5, center)
    worksheet150.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF GEDONG KUNING', title)
    worksheet150.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet150.write('A5', 'LOKASI', header)
    worksheet150.write('B5', 'TOTAL', header)
    worksheet150.merge_range('A4:B4', 'RANK', header)
    worksheet150.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet150.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet150.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet150.merge_range('F4:F5', 'KELAS', header)
    worksheet150.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet150.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet150.write('G5', 'MAW', body)
    worksheet150.write('H5', 'MAP', body)
    worksheet150.write('I5', 'IND', body)
    worksheet150.write('J5', 'ENG', body)
    worksheet150.write('K5', 'SEJ', body)
    worksheet150.write('L5', 'GEO', body)
    worksheet150.write('M5', 'EKO', body)
    worksheet150.write('N5', 'SOS', body)
    worksheet150.write('O5', 'FIS', body)
    worksheet150.write('P5', 'KIM', body)
    worksheet150.write('Q5', 'BIO', body)
    worksheet150.write('R5', 'JML', body)
    worksheet150.write('S5', 'MAW', body)
    worksheet150.write('T5', 'MAP', body)
    worksheet150.write('U5', 'IND', body)
    worksheet150.write('V5', 'ENG', body)
    worksheet150.write('W5', 'SEJ', body)
    worksheet150.write('X5', 'GEO', body)
    worksheet150.write('Y5', 'EKO', body)
    worksheet150.write('Z5', 'SOS', body)
    worksheet150.write('AA5', 'FIS', body)
    worksheet150.write('AB5', 'KIM', body)
    worksheet150.write('AC5', 'BIO', body)
    worksheet150.write('AD5', 'JML', body)

    worksheet150.conditional_format(5, 0, row150_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet150.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF GEDONG KUNING', title)
    worksheet150.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet150.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet150.write('A22', 'LOKASI', header)
    worksheet150.write('B22', 'TOTAL', header)
    worksheet150.merge_range('A21:B21', 'RANK', header)
    worksheet150.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet150.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet150.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet150.merge_range('F21:F22', 'KELAS', header)
    worksheet150.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet150.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet150.write('G22', 'MAW', body)
    worksheet150.write('H22', 'MAP', body)
    worksheet150.write('I22', 'IND', body)
    worksheet150.write('J22', 'ENG', body)
    worksheet150.write('K22', 'SEJ', body)
    worksheet150.write('L22', 'GEO', body)
    worksheet150.write('M22', 'EKO', body)
    worksheet150.write('N22', 'SOS', body)
    worksheet150.write('O22', 'FIS', body)
    worksheet150.write('P22', 'KIM', body)
    worksheet150.write('Q22', 'BIO', body)
    worksheet150.write('R22', 'JML', body)
    worksheet150.write('S22', 'MAW', body)
    worksheet150.write('T22', 'MAP', body)
    worksheet150.write('U22', 'IND', body)
    worksheet150.write('V22', 'ENG', body)
    worksheet150.write('W22', 'SEJ', body)
    worksheet150.write('X22', 'GEO', body)
    worksheet150.write('Y22', 'EKO', body)
    worksheet150.write('Z22', 'SOS', body)
    worksheet150.write('AA22', 'FIS', body)
    worksheet150.write('AB22', 'KIM', body)
    worksheet150.write('AC22', 'BIO', body)
    worksheet150.write('AD22', 'JML', body)

    worksheet150.conditional_format(22, 0, row150+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 151
    worksheet151.insert_image('A1', r'logo resmi nf.jpg')

    worksheet151.set_column('A:A', 7, center)
    worksheet151.set_column('B:B', 6, center)
    worksheet151.set_column('C:C', 18.14, center)
    worksheet151.set_column('D:D', 25, left)
    worksheet151.set_column('E:E', 13.14, left)
    worksheet151.set_column('F:F', 8.57, center)
    worksheet151.set_column('G:AD', 5, center)
    worksheet151.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF JATIWARINGIN', title)
    worksheet151.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet151.write('A5', 'LOKASI', header)
    worksheet151.write('B5', 'TOTAL', header)
    worksheet151.merge_range('A4:B4', 'RANK', header)
    worksheet151.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet151.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet151.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet151.merge_range('F4:F5', 'KELAS', header)
    worksheet151.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet151.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet151.write('G5', 'MAW', body)
    worksheet151.write('H5', 'MAP', body)
    worksheet151.write('I5', 'IND', body)
    worksheet151.write('J5', 'ENG', body)
    worksheet151.write('K5', 'SEJ', body)
    worksheet151.write('L5', 'GEO', body)
    worksheet151.write('M5', 'EKO', body)
    worksheet151.write('N5', 'SOS', body)
    worksheet151.write('O5', 'FIS', body)
    worksheet151.write('P5', 'KIM', body)
    worksheet151.write('Q5', 'BIO', body)
    worksheet151.write('R5', 'JML', body)
    worksheet151.write('S5', 'MAW', body)
    worksheet151.write('T5', 'MAP', body)
    worksheet151.write('U5', 'IND', body)
    worksheet151.write('V5', 'ENG', body)
    worksheet151.write('W5', 'SEJ', body)
    worksheet151.write('X5', 'GEO', body)
    worksheet151.write('Y5', 'EKO', body)
    worksheet151.write('Z5', 'SOS', body)
    worksheet151.write('AA5', 'FIS', body)
    worksheet151.write('AB5', 'KIM', body)
    worksheet151.write('AC5', 'BIO', body)
    worksheet151.write('AD5', 'JML', body)

    worksheet151.conditional_format(5, 0, row151_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet151.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF JATIWARINGIN', title)
    worksheet151.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet151.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet151.write('A22', 'LOKASI', header)
    worksheet151.write('B22', 'TOTAL', header)
    worksheet151.merge_range('A21:B21', 'RANK', header)
    worksheet151.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet151.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet151.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet151.merge_range('F21:F22', 'KELAS', header)
    worksheet151.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet151.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet151.write('G22', 'MAW', body)
    worksheet151.write('H22', 'MAP', body)
    worksheet151.write('I22', 'IND', body)
    worksheet151.write('J22', 'ENG', body)
    worksheet151.write('K22', 'SEJ', body)
    worksheet151.write('L22', 'GEO', body)
    worksheet151.write('M22', 'EKO', body)
    worksheet151.write('N22', 'SOS', body)
    worksheet151.write('O22', 'FIS', body)
    worksheet151.write('P22', 'KIM', body)
    worksheet151.write('Q22', 'BIO', body)
    worksheet151.write('R22', 'JML', body)
    worksheet151.write('S22', 'MAW', body)
    worksheet151.write('T22', 'MAP', body)
    worksheet151.write('U22', 'IND', body)
    worksheet151.write('V22', 'ENG', body)
    worksheet151.write('W22', 'SEJ', body)
    worksheet151.write('X22', 'GEO', body)
    worksheet151.write('Y22', 'EKO', body)
    worksheet151.write('Z22', 'SOS', body)
    worksheet151.write('AA22', 'FIS', body)
    worksheet151.write('AB22', 'KIM', body)
    worksheet151.write('AC22', 'BIO', body)
    worksheet151.write('AD22', 'JML', body)

    worksheet151.conditional_format(22, 0, row151+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 152
    worksheet152.insert_image('A1', r'logo resmi nf.jpg')

    worksheet152.set_column('A:A', 7, center)
    worksheet152.set_column('B:B', 6, center)
    worksheet152.set_column('C:C', 18.14, center)
    worksheet152.set_column('D:D', 25, left)
    worksheet152.set_column('E:E', 13.14, left)
    worksheet152.set_column('F:F', 8.57, center)
    worksheet152.set_column('G:AD', 5, center)
    worksheet152.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CILEDUG', title)
    worksheet152.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet152.write('A5', 'LOKASI', header)
    worksheet152.write('B5', 'TOTAL', header)
    worksheet152.merge_range('A4:B4', 'RANK', header)
    worksheet152.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet152.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet152.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet152.merge_range('F4:F5', 'KELAS', header)
    worksheet152.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet152.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet152.write('G5', 'MAW', body)
    worksheet152.write('H5', 'MAP', body)
    worksheet152.write('I5', 'IND', body)
    worksheet152.write('J5', 'ENG', body)
    worksheet152.write('K5', 'SEJ', body)
    worksheet152.write('L5', 'GEO', body)
    worksheet152.write('M5', 'EKO', body)
    worksheet152.write('N5', 'SOS', body)
    worksheet152.write('O5', 'FIS', body)
    worksheet152.write('P5', 'KIM', body)
    worksheet152.write('Q5', 'BIO', body)
    worksheet152.write('R5', 'JML', body)
    worksheet152.write('S5', 'MAW', body)
    worksheet152.write('T5', 'MAP', body)
    worksheet152.write('U5', 'IND', body)
    worksheet152.write('V5', 'ENG', body)
    worksheet152.write('W5', 'SEJ', body)
    worksheet152.write('X5', 'GEO', body)
    worksheet152.write('Y5', 'EKO', body)
    worksheet152.write('Z5', 'SOS', body)
    worksheet152.write('AA5', 'FIS', body)
    worksheet152.write('AB5', 'KIM', body)
    worksheet152.write('AC5', 'BIO', body)
    worksheet152.write('AD5', 'JML', body)

    worksheet152.conditional_format(5, 0, row152_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet152.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF CILEDUG', title)
    worksheet152.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet152.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet152.write('A22', 'LOKASI', header)
    worksheet152.write('B22', 'TOTAL', header)
    worksheet152.merge_range('A21:B21', 'RANK', header)
    worksheet152.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet152.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet152.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet152.merge_range('F21:F22', 'KELAS', header)
    worksheet152.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet152.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet152.write('G22', 'MAW', body)
    worksheet152.write('H22', 'MAP', body)
    worksheet152.write('I22', 'IND', body)
    worksheet152.write('J22', 'ENG', body)
    worksheet152.write('K22', 'SEJ', body)
    worksheet152.write('L22', 'GEO', body)
    worksheet152.write('M22', 'EKO', body)
    worksheet152.write('N22', 'SOS', body)
    worksheet152.write('O22', 'FIS', body)
    worksheet152.write('P22', 'KIM', body)
    worksheet152.write('Q22', 'BIO', body)
    worksheet152.write('R22', 'JML', body)
    worksheet152.write('S22', 'MAW', body)
    worksheet152.write('T22', 'MAP', body)
    worksheet152.write('U22', 'IND', body)
    worksheet152.write('V22', 'ENG', body)
    worksheet152.write('W22', 'SEJ', body)
    worksheet152.write('X22', 'GEO', body)
    worksheet152.write('Y22', 'EKO', body)
    worksheet152.write('Z22', 'SOS', body)
    worksheet152.write('AA22', 'FIS', body)
    worksheet152.write('AB22', 'KIM', body)
    worksheet152.write('AC22', 'BIO', body)
    worksheet152.write('AD22', 'JML', body)

    worksheet152.conditional_format(22, 0, row152+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 153
    worksheet153.insert_image('A1', r'logo resmi nf.jpg')

    worksheet153.set_column('A:A', 7, center)
    worksheet153.set_column('B:B', 6, center)
    worksheet153.set_column('C:C', 18.14, center)
    worksheet153.set_column('D:D', 25, left)
    worksheet153.set_column('E:E', 13.14, left)
    worksheet153.set_column('F:F', 8.57, center)
    worksheet153.set_column('G:AD', 5, center)
    worksheet153.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KRANGGAN', title)
    worksheet153.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet153.write('A5', 'LOKASI', header)
    worksheet153.write('B5', 'TOTAL', header)
    worksheet153.merge_range('A4:B4', 'RANK', header)
    worksheet153.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet153.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet153.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet153.merge_range('F4:F5', 'KELAS', header)
    worksheet153.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet153.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet153.write('G5', 'MAW', body)
    worksheet153.write('H5', 'MAP', body)
    worksheet153.write('I5', 'IND', body)
    worksheet153.write('J5', 'ENG', body)
    worksheet153.write('K5', 'SEJ', body)
    worksheet153.write('L5', 'GEO', body)
    worksheet153.write('M5', 'EKO', body)
    worksheet153.write('N5', 'SOS', body)
    worksheet153.write('O5', 'FIS', body)
    worksheet153.write('P5', 'KIM', body)
    worksheet153.write('Q5', 'BIO', body)
    worksheet153.write('R5', 'JML', body)
    worksheet153.write('S5', 'MAW', body)
    worksheet153.write('T5', 'MAP', body)
    worksheet153.write('U5', 'IND', body)
    worksheet153.write('V5', 'ENG', body)
    worksheet153.write('W5', 'SEJ', body)
    worksheet153.write('X5', 'GEO', body)
    worksheet153.write('Y5', 'EKO', body)
    worksheet153.write('Z5', 'SOS', body)
    worksheet153.write('AA5', 'FIS', body)
    worksheet153.write('AB5', 'KIM', body)
    worksheet153.write('AC5', 'BIO', body)
    worksheet153.write('AD5', 'JML', body)

    worksheet153.conditional_format(5, 0, row153_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet153.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF KRANGGAN', title)
    worksheet153.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet153.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet153.write('A22', 'LOKASI', header)
    worksheet153.write('B22', 'TOTAL', header)
    worksheet153.merge_range('A21:B21', 'RANK', header)
    worksheet153.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet153.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet153.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet153.merge_range('F21:F22', 'KELAS', header)
    worksheet153.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet153.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet153.write('G22', 'MAW', body)
    worksheet153.write('H22', 'MAP', body)
    worksheet153.write('I22', 'IND', body)
    worksheet153.write('J22', 'ENG', body)
    worksheet153.write('K22', 'SEJ', body)
    worksheet153.write('L22', 'GEO', body)
    worksheet153.write('M22', 'EKO', body)
    worksheet153.write('N22', 'SOS', body)
    worksheet153.write('O22', 'FIS', body)
    worksheet153.write('P22', 'KIM', body)
    worksheet153.write('Q22', 'BIO', body)
    worksheet153.write('R22', 'JML', body)
    worksheet153.write('S22', 'MAW', body)
    worksheet153.write('T22', 'MAP', body)
    worksheet153.write('U22', 'IND', body)
    worksheet153.write('V22', 'ENG', body)
    worksheet153.write('W22', 'SEJ', body)
    worksheet153.write('X22', 'GEO', body)
    worksheet153.write('Y22', 'EKO', body)
    worksheet153.write('Z22', 'SOS', body)
    worksheet153.write('AA22', 'FIS', body)
    worksheet153.write('AB22', 'KIM', body)
    worksheet153.write('AC22', 'BIO', body)
    worksheet153.write('AD22', 'JML', body)

    worksheet153.conditional_format(22, 0, row153+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 154
    worksheet154.insert_image('A1', r'logo resmi nf.jpg')

    worksheet154.set_column('A:A', 7, center)
    worksheet154.set_column('B:B', 6, center)
    worksheet154.set_column('C:C', 18.14, center)
    worksheet154.set_column('D:D', 25, left)
    worksheet154.set_column('E:E', 13.14, left)
    worksheet154.set_column('F:F', 8.57, center)
    worksheet154.set_column('G:AD', 5, center)
    worksheet154.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MUSTIKA JAYA', title)
    worksheet154.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet154.write('A5', 'LOKASI', header)
    worksheet154.write('B5', 'TOTAL', header)
    worksheet154.merge_range('A4:B4', 'RANK', header)
    worksheet154.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet154.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet154.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet154.merge_range('F4:F5', 'KELAS', header)
    worksheet154.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet154.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet154.write('G5', 'MAW', body)
    worksheet154.write('H5', 'MAP', body)
    worksheet154.write('I5', 'IND', body)
    worksheet154.write('J5', 'ENG', body)
    worksheet154.write('K5', 'SEJ', body)
    worksheet154.write('L5', 'GEO', body)
    worksheet154.write('M5', 'EKO', body)
    worksheet154.write('N5', 'SOS', body)
    worksheet154.write('O5', 'FIS', body)
    worksheet154.write('P5', 'KIM', body)
    worksheet154.write('Q5', 'BIO', body)
    worksheet154.write('R5', 'JML', body)
    worksheet154.write('S5', 'MAW', body)
    worksheet154.write('T5', 'MAP', body)
    worksheet154.write('U5', 'IND', body)
    worksheet154.write('V5', 'ENG', body)
    worksheet154.write('W5', 'SEJ', body)
    worksheet154.write('X5', 'GEO', body)
    worksheet154.write('Y5', 'EKO', body)
    worksheet154.write('Z5', 'SOS', body)
    worksheet154.write('AA5', 'FIS', body)
    worksheet154.write('AB5', 'KIM', body)
    worksheet154.write('AC5', 'BIO', body)
    worksheet154.write('AD5', 'JML', body)

    worksheet154.conditional_format(5, 0, row154_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet154.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF MUSTIKA JAYA', title)
    worksheet154.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet154.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet154.write('A22', 'LOKASI', header)
    worksheet154.write('B22', 'TOTAL', header)
    worksheet154.merge_range('A21:B21', 'RANK', header)
    worksheet154.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet154.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet154.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet154.merge_range('F21:F22', 'KELAS', header)
    worksheet154.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet154.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet154.write('G22', 'MAW', body)
    worksheet154.write('H22', 'MAP', body)
    worksheet154.write('I22', 'IND', body)
    worksheet154.write('J22', 'ENG', body)
    worksheet154.write('K22', 'SEJ', body)
    worksheet154.write('L22', 'GEO', body)
    worksheet154.write('M22', 'EKO', body)
    worksheet154.write('N22', 'SOS', body)
    worksheet154.write('O22', 'FIS', body)
    worksheet154.write('P22', 'KIM', body)
    worksheet154.write('Q22', 'BIO', body)
    worksheet154.write('R22', 'JML', body)
    worksheet154.write('S22', 'MAW', body)
    worksheet154.write('T22', 'MAP', body)
    worksheet154.write('U22', 'IND', body)
    worksheet154.write('V22', 'ENG', body)
    worksheet154.write('W22', 'SEJ', body)
    worksheet154.write('X22', 'GEO', body)
    worksheet154.write('Y22', 'EKO', body)
    worksheet154.write('Z22', 'SOS', body)
    worksheet154.write('AA22', 'FIS', body)
    worksheet154.write('AB22', 'KIM', body)
    worksheet154.write('AC22', 'BIO', body)
    worksheet154.write('AD22', 'JML', body)

    worksheet154.conditional_format(22, 0, row154+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 155
    worksheet155.insert_image('A1', r'logo resmi nf.jpg')

    worksheet155.set_column('A:A', 7, center)
    worksheet155.set_column('B:B', 6, center)
    worksheet155.set_column('C:C', 18.14, center)
    worksheet155.set_column('D:D', 25, left)
    worksheet155.set_column('E:E', 13.14, left)
    worksheet155.set_column('F:F', 8.57, center)
    worksheet155.set_column('G:AD', 5, center)
    worksheet155.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF ALEXINDO', title)
    worksheet155.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet155.write('A5', 'LOKASI', header)
    worksheet155.write('B5', 'TOTAL', header)
    worksheet155.merge_range('A4:B4', 'RANK', header)
    worksheet155.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet155.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet155.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet155.merge_range('F4:F5', 'KELAS', header)
    worksheet155.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet155.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet155.write('G5', 'MAW', body)
    worksheet155.write('H5', 'MAP', body)
    worksheet155.write('I5', 'IND', body)
    worksheet155.write('J5', 'ENG', body)
    worksheet155.write('K5', 'SEJ', body)
    worksheet155.write('L5', 'GEO', body)
    worksheet155.write('M5', 'EKO', body)
    worksheet155.write('N5', 'SOS', body)
    worksheet155.write('O5', 'FIS', body)
    worksheet155.write('P5', 'KIM', body)
    worksheet155.write('Q5', 'BIO', body)
    worksheet155.write('R5', 'JML', body)
    worksheet155.write('S5', 'MAW', body)
    worksheet155.write('T5', 'MAP', body)
    worksheet155.write('U5', 'IND', body)
    worksheet155.write('V5', 'ENG', body)
    worksheet155.write('W5', 'SEJ', body)
    worksheet155.write('X5', 'GEO', body)
    worksheet155.write('Y5', 'EKO', body)
    worksheet155.write('Z5', 'SOS', body)
    worksheet155.write('AA5', 'FIS', body)
    worksheet155.write('AB5', 'KIM', body)
    worksheet155.write('AC5', 'BIO', body)
    worksheet155.write('AD5', 'JML', body)

    worksheet155.conditional_format(5, 0, row155_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet155.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF ALEXINDO', title)
    worksheet155.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet155.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet155.write('A22', 'LOKASI', header)
    worksheet155.write('B22', 'TOTAL', header)
    worksheet155.merge_range('A21:B21', 'RANK', header)
    worksheet155.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet155.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet155.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet155.merge_range('F21:F22', 'KELAS', header)
    worksheet155.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet155.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet155.write('G22', 'MAW', body)
    worksheet155.write('H22', 'MAP', body)
    worksheet155.write('I22', 'IND', body)
    worksheet155.write('J22', 'ENG', body)
    worksheet155.write('K22', 'SEJ', body)
    worksheet155.write('L22', 'GEO', body)
    worksheet155.write('M22', 'EKO', body)
    worksheet155.write('N22', 'SOS', body)
    worksheet155.write('O22', 'FIS', body)
    worksheet155.write('P22', 'KIM', body)
    worksheet155.write('Q22', 'BIO', body)
    worksheet155.write('R22', 'JML', body)
    worksheet155.write('S22', 'MAW', body)
    worksheet155.write('T22', 'MAP', body)
    worksheet155.write('U22', 'IND', body)
    worksheet155.write('V22', 'ENG', body)
    worksheet155.write('W22', 'SEJ', body)
    worksheet155.write('X22', 'GEO', body)
    worksheet155.write('Y22', 'EKO', body)
    worksheet155.write('Z22', 'SOS', body)
    worksheet155.write('AA22', 'FIS', body)
    worksheet155.write('AB22', 'KIM', body)
    worksheet155.write('AC22', 'BIO', body)
    worksheet155.write('AD22', 'JML', body)

    worksheet155.conditional_format(22, 0, row155+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 156
    worksheet156.insert_image('A1', r'logo resmi nf.jpg')

    worksheet156.set_column('A:A', 7, center)
    worksheet156.set_column('B:B', 6, center)
    worksheet156.set_column('C:C', 18.14, center)
    worksheet156.set_column('D:D', 25, left)
    worksheet156.set_column('E:E', 13.14, left)
    worksheet156.set_column('F:F', 8.57, center)
    worksheet156.set_column('G:AD', 5, center)
    worksheet156.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIBITUNG', title)
    worksheet156.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet156.write('A5', 'LOKASI', header)
    worksheet156.write('B5', 'TOTAL', header)
    worksheet156.merge_range('A4:B4', 'RANK', header)
    worksheet156.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet156.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet156.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet156.merge_range('F4:F5', 'KELAS', header)
    worksheet156.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet156.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet156.write('G5', 'MAW', body)
    worksheet156.write('H5', 'MAP', body)
    worksheet156.write('I5', 'IND', body)
    worksheet156.write('J5', 'ENG', body)
    worksheet156.write('K5', 'SEJ', body)
    worksheet156.write('L5', 'GEO', body)
    worksheet156.write('M5', 'EKO', body)
    worksheet156.write('N5', 'SOS', body)
    worksheet156.write('O5', 'FIS', body)
    worksheet156.write('P5', 'KIM', body)
    worksheet156.write('Q5', 'BIO', body)
    worksheet156.write('R5', 'JML', body)
    worksheet156.write('S5', 'MAW', body)
    worksheet156.write('T5', 'MAP', body)
    worksheet156.write('U5', 'IND', body)
    worksheet156.write('V5', 'ENG', body)
    worksheet156.write('W5', 'SEJ', body)
    worksheet156.write('X5', 'GEO', body)
    worksheet156.write('Y5', 'EKO', body)
    worksheet156.write('Z5', 'SOS', body)
    worksheet156.write('AA5', 'FIS', body)
    worksheet156.write('AB5', 'KIM', body)
    worksheet156.write('AC5', 'BIO', body)
    worksheet156.write('AD5', 'JML', body)

    worksheet156.conditional_format(5, 0, row156_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet156.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF CIBITUNG', title)
    worksheet156.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet156.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet156.write('A22', 'LOKASI', header)
    worksheet156.write('B22', 'TOTAL', header)
    worksheet156.merge_range('A21:B21', 'RANK', header)
    worksheet156.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet156.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet156.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet156.merge_range('F21:F22', 'KELAS', header)
    worksheet156.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet156.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet156.write('G22', 'MAW', body)
    worksheet156.write('H22', 'MAP', body)
    worksheet156.write('I22', 'IND', body)
    worksheet156.write('J22', 'ENG', body)
    worksheet156.write('K22', 'SEJ', body)
    worksheet156.write('L22', 'GEO', body)
    worksheet156.write('M22', 'EKO', body)
    worksheet156.write('N22', 'SOS', body)
    worksheet156.write('O22', 'FIS', body)
    worksheet156.write('P22', 'KIM', body)
    worksheet156.write('Q22', 'BIO', body)
    worksheet156.write('R22', 'JML', body)
    worksheet156.write('S22', 'MAW', body)
    worksheet156.write('T22', 'MAP', body)
    worksheet156.write('U22', 'IND', body)
    worksheet156.write('V22', 'ENG', body)
    worksheet156.write('W22', 'SEJ', body)
    worksheet156.write('X22', 'GEO', body)
    worksheet156.write('Y22', 'EKO', body)
    worksheet156.write('Z22', 'SOS', body)
    worksheet156.write('AA22', 'FIS', body)
    worksheet156.write('AB22', 'KIM', body)
    worksheet156.write('AC22', 'BIO', body)
    worksheet156.write('AD22', 'JML', body)

    worksheet156.conditional_format(22, 0, row156+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 157
    worksheet157.insert_image('A1', r'logo resmi nf.jpg')

    worksheet157.set_column('A:A', 7, center)
    worksheet157.set_column('B:B', 6, center)
    worksheet157.set_column('C:C', 18.14, center)
    worksheet157.set_column('D:D', 25, left)
    worksheet157.set_column('E:E', 13.14, left)
    worksheet157.set_column('F:F', 8.57, center)
    worksheet157.set_column('G:AD', 5, center)
    worksheet157.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KRAMAT JAYA', title)
    worksheet157.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet157.write('A5', 'LOKASI', header)
    worksheet157.write('B5', 'TOTAL', header)
    worksheet157.merge_range('A4:B4', 'RANK', header)
    worksheet157.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet157.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet157.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet157.merge_range('F4:F5', 'KELAS', header)
    worksheet157.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet157.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet157.write('G5', 'MAW', body)
    worksheet157.write('H5', 'MAP', body)
    worksheet157.write('I5', 'IND', body)
    worksheet157.write('J5', 'ENG', body)
    worksheet157.write('K5', 'SEJ', body)
    worksheet157.write('L5', 'GEO', body)
    worksheet157.write('M5', 'EKO', body)
    worksheet157.write('N5', 'SOS', body)
    worksheet157.write('O5', 'FIS', body)
    worksheet157.write('P5', 'KIM', body)
    worksheet157.write('Q5', 'BIO', body)
    worksheet157.write('R5', 'JML', body)
    worksheet157.write('S5', 'MAW', body)
    worksheet157.write('T5', 'MAP', body)
    worksheet157.write('U5', 'IND', body)
    worksheet157.write('V5', 'ENG', body)
    worksheet157.write('W5', 'SEJ', body)
    worksheet157.write('X5', 'GEO', body)
    worksheet157.write('Y5', 'EKO', body)
    worksheet157.write('Z5', 'SOS', body)
    worksheet157.write('AA5', 'FIS', body)
    worksheet157.write('AB5', 'KIM', body)
    worksheet157.write('AC5', 'BIO', body)
    worksheet157.write('AD5', 'JML', body)

    worksheet157.conditional_format(5, 0, row157_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet157.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF KRAMAT JAYA', title)
    worksheet157.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet157.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet157.write('A22', 'LOKASI', header)
    worksheet157.write('B22', 'TOTAL', header)
    worksheet157.merge_range('A21:B21', 'RANK', header)
    worksheet157.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet157.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet157.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet157.merge_range('F21:F22', 'KELAS', header)
    worksheet157.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet157.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet157.write('G22', 'MAW', body)
    worksheet157.write('H22', 'MAP', body)
    worksheet157.write('I22', 'IND', body)
    worksheet157.write('J22', 'ENG', body)
    worksheet157.write('K22', 'SEJ', body)
    worksheet157.write('L22', 'GEO', body)
    worksheet157.write('M22', 'EKO', body)
    worksheet157.write('N22', 'SOS', body)
    worksheet157.write('O22', 'FIS', body)
    worksheet157.write('P22', 'KIM', body)
    worksheet157.write('Q22', 'BIO', body)
    worksheet157.write('R22', 'JML', body)
    worksheet157.write('S22', 'MAW', body)
    worksheet157.write('T22', 'MAP', body)
    worksheet157.write('U22', 'IND', body)
    worksheet157.write('V22', 'ENG', body)
    worksheet157.write('W22', 'SEJ', body)
    worksheet157.write('X22', 'GEO', body)
    worksheet157.write('Y22', 'EKO', body)
    worksheet157.write('Z22', 'SOS', body)
    worksheet157.write('AA22', 'FIS', body)
    worksheet157.write('AB22', 'KIM', body)
    worksheet157.write('AC22', 'BIO', body)
    worksheet157.write('AD22', 'JML', body)

    worksheet157.conditional_format(22, 0, row157+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 158
    worksheet158.insert_image('A1', r'logo resmi nf.jpg')

    worksheet158.set_column('A:A', 7, center)
    worksheet158.set_column('B:B', 6, center)
    worksheet158.set_column('C:C', 18.14, center)
    worksheet158.set_column('D:D', 25, left)
    worksheet158.set_column('E:E', 13.14, left)
    worksheet158.set_column('F:F', 8.57, center)
    worksheet158.set_column('G:AD', 5, center)
    worksheet158.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PONDOK GEDE', title)
    worksheet158.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet158.write('A5', 'LOKASI', header)
    worksheet158.write('B5', 'TOTAL', header)
    worksheet158.merge_range('A4:B4', 'RANK', header)
    worksheet158.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet158.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet158.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet158.merge_range('F4:F5', 'KELAS', header)
    worksheet158.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet158.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet158.write('G5', 'MAW', body)
    worksheet158.write('H5', 'MAP', body)
    worksheet158.write('I5', 'IND', body)
    worksheet158.write('J5', 'ENG', body)
    worksheet158.write('K5', 'SEJ', body)
    worksheet158.write('L5', 'GEO', body)
    worksheet158.write('M5', 'EKO', body)
    worksheet158.write('N5', 'SOS', body)
    worksheet158.write('O5', 'FIS', body)
    worksheet158.write('P5', 'KIM', body)
    worksheet158.write('Q5', 'BIO', body)
    worksheet158.write('R5', 'JML', body)
    worksheet158.write('S5', 'MAW', body)
    worksheet158.write('T5', 'MAP', body)
    worksheet158.write('U5', 'IND', body)
    worksheet158.write('V5', 'ENG', body)
    worksheet158.write('W5', 'SEJ', body)
    worksheet158.write('X5', 'GEO', body)
    worksheet158.write('Y5', 'EKO', body)
    worksheet158.write('Z5', 'SOS', body)
    worksheet158.write('AA5', 'FIS', body)
    worksheet158.write('AB5', 'KIM', body)
    worksheet158.write('AC5', 'BIO', body)
    worksheet158.write('AD5', 'JML', body)

    worksheet158.conditional_format(5, 0, row158_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet158.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF PONDOK GEDE', title)
    worksheet158.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet158.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet158.write('A22', 'LOKASI', header)
    worksheet158.write('B22', 'TOTAL', header)
    worksheet158.merge_range('A21:B21', 'RANK', header)
    worksheet158.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet158.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet158.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet158.merge_range('F21:F22', 'KELAS', header)
    worksheet158.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet158.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet158.write('G22', 'MAW', body)
    worksheet158.write('H22', 'MAP', body)
    worksheet158.write('I22', 'IND', body)
    worksheet158.write('J22', 'ENG', body)
    worksheet158.write('K22', 'SEJ', body)
    worksheet158.write('L22', 'GEO', body)
    worksheet158.write('M22', 'EKO', body)
    worksheet158.write('N22', 'SOS', body)
    worksheet158.write('O22', 'FIS', body)
    worksheet158.write('P22', 'KIM', body)
    worksheet158.write('Q22', 'BIO', body)
    worksheet158.write('R22', 'JML', body)
    worksheet158.write('S22', 'MAW', body)
    worksheet158.write('T22', 'MAP', body)
    worksheet158.write('U22', 'IND', body)
    worksheet158.write('V22', 'ENG', body)
    worksheet158.write('W22', 'SEJ', body)
    worksheet158.write('X22', 'GEO', body)
    worksheet158.write('Y22', 'EKO', body)
    worksheet158.write('Z22', 'SOS', body)
    worksheet158.write('AA22', 'FIS', body)
    worksheet158.write('AB22', 'KIM', body)
    worksheet158.write('AC22', 'BIO', body)
    worksheet158.write('AD22', 'JML', body)

    worksheet158.conditional_format(22, 0, row158+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 159
    worksheet159.insert_image('A1', r'logo resmi nf.jpg')

    worksheet159.set_column('A:A', 7, center)
    worksheet159.set_column('B:B', 6, center)
    worksheet159.set_column('C:C', 18.14, center)
    worksheet159.set_column('D:D', 25, left)
    worksheet159.set_column('E:E', 13.14, left)
    worksheet159.set_column('F:F', 8.57, center)
    worksheet159.set_column('G:AD', 5, center)
    worksheet159.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF GALAXY', title)
    worksheet159.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet159.write('A5', 'LOKASI', header)
    worksheet159.write('B5', 'TOTAL', header)
    worksheet159.merge_range('A4:B4', 'RANK', header)
    worksheet159.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet159.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet159.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet159.merge_range('F4:F5', 'KELAS', header)
    worksheet159.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet159.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet159.write('G5', 'MAW', body)
    worksheet159.write('H5', 'MAP', body)
    worksheet159.write('I5', 'IND', body)
    worksheet159.write('J5', 'ENG', body)
    worksheet159.write('K5', 'SEJ', body)
    worksheet159.write('L5', 'GEO', body)
    worksheet159.write('M5', 'EKO', body)
    worksheet159.write('N5', 'SOS', body)
    worksheet159.write('O5', 'FIS', body)
    worksheet159.write('P5', 'KIM', body)
    worksheet159.write('Q5', 'BIO', body)
    worksheet159.write('R5', 'JML', body)
    worksheet159.write('S5', 'MAW', body)
    worksheet159.write('T5', 'MAP', body)
    worksheet159.write('U5', 'IND', body)
    worksheet159.write('V5', 'ENG', body)
    worksheet159.write('W5', 'SEJ', body)
    worksheet159.write('X5', 'GEO', body)
    worksheet159.write('Y5', 'EKO', body)
    worksheet159.write('Z5', 'SOS', body)
    worksheet159.write('AA5', 'FIS', body)
    worksheet159.write('AB5', 'KIM', body)
    worksheet159.write('AC5', 'BIO', body)
    worksheet159.write('AD5', 'JML', body)

    worksheet159.conditional_format(5, 0, row159_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet159.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF GALAXY', title)
    worksheet159.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet159.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet159.write('A22', 'LOKASI', header)
    worksheet159.write('B22', 'TOTAL', header)
    worksheet159.merge_range('A21:B21', 'RANK', header)
    worksheet159.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet159.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet159.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet159.merge_range('F21:F22', 'KELAS', header)
    worksheet159.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet159.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet159.write('G22', 'MAW', body)
    worksheet159.write('H22', 'MAP', body)
    worksheet159.write('I22', 'IND', body)
    worksheet159.write('J22', 'ENG', body)
    worksheet159.write('K22', 'SEJ', body)
    worksheet159.write('L22', 'GEO', body)
    worksheet159.write('M22', 'EKO', body)
    worksheet159.write('N22', 'SOS', body)
    worksheet159.write('O22', 'FIS', body)
    worksheet159.write('P22', 'KIM', body)
    worksheet159.write('Q22', 'BIO', body)
    worksheet159.write('R22', 'JML', body)
    worksheet159.write('S22', 'MAW', body)
    worksheet159.write('T22', 'MAP', body)
    worksheet159.write('U22', 'IND', body)
    worksheet159.write('V22', 'ENG', body)
    worksheet159.write('W22', 'SEJ', body)
    worksheet159.write('X22', 'GEO', body)
    worksheet159.write('Y22', 'EKO', body)
    worksheet159.write('Z22', 'SOS', body)
    worksheet159.write('AA22', 'FIS', body)
    worksheet159.write('AB22', 'KIM', body)
    worksheet159.write('AC22', 'BIO', body)
    worksheet159.write('AD22', 'JML', body)

    worksheet159.conditional_format(22, 0, row159+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 160
    worksheet160.insert_image('A1', r'logo resmi nf.jpg')

    worksheet160.set_column('A:A', 7, center)
    worksheet160.set_column('B:B', 6, center)
    worksheet160.set_column('C:C', 18.14, center)
    worksheet160.set_column('D:D', 25, left)
    worksheet160.set_column('E:E', 13.14, left)
    worksheet160.set_column('F:F', 8.57, center)
    worksheet160.set_column('G:AD', 5, center)
    worksheet160.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIGANJUR', title)
    worksheet160.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet160.write('A5', 'LOKASI', header)
    worksheet160.write('B5', 'TOTAL', header)
    worksheet160.merge_range('A4:B4', 'RANK', header)
    worksheet160.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet160.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet160.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet160.merge_range('F4:F5', 'KELAS', header)
    worksheet160.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet160.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet160.write('G5', 'MAW', body)
    worksheet160.write('H5', 'MAP', body)
    worksheet160.write('I5', 'IND', body)
    worksheet160.write('J5', 'ENG', body)
    worksheet160.write('K5', 'SEJ', body)
    worksheet160.write('L5', 'GEO', body)
    worksheet160.write('M5', 'EKO', body)
    worksheet160.write('N5', 'SOS', body)
    worksheet160.write('O5', 'FIS', body)
    worksheet160.write('P5', 'KIM', body)
    worksheet160.write('Q5', 'BIO', body)
    worksheet160.write('R5', 'JML', body)
    worksheet160.write('S5', 'MAW', body)
    worksheet160.write('T5', 'MAP', body)
    worksheet160.write('U5', 'IND', body)
    worksheet160.write('V5', 'ENG', body)
    worksheet160.write('W5', 'SEJ', body)
    worksheet160.write('X5', 'GEO', body)
    worksheet160.write('Y5', 'EKO', body)
    worksheet160.write('Z5', 'SOS', body)
    worksheet160.write('AA5', 'FIS', body)
    worksheet160.write('AB5', 'KIM', body)
    worksheet160.write('AC5', 'BIO', body)
    worksheet160.write('AD5', 'JML', body)

    worksheet160.conditional_format(5, 0, row160_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet160.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF CIGANJUR', title)
    worksheet160.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet160.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet160.write('A22', 'LOKASI', header)
    worksheet160.write('B22', 'TOTAL', header)
    worksheet160.merge_range('A21:B21', 'RANK', header)
    worksheet160.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet160.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet160.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet160.merge_range('F21:F22', 'KELAS', header)
    worksheet160.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet160.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet160.write('G22', 'MAW', body)
    worksheet160.write('H22', 'MAP', body)
    worksheet160.write('I22', 'IND', body)
    worksheet160.write('J22', 'ENG', body)
    worksheet160.write('K22', 'SEJ', body)
    worksheet160.write('L22', 'GEO', body)
    worksheet160.write('M22', 'EKO', body)
    worksheet160.write('N22', 'SOS', body)
    worksheet160.write('O22', 'FIS', body)
    worksheet160.write('P22', 'KIM', body)
    worksheet160.write('Q22', 'BIO', body)
    worksheet160.write('R22', 'JML', body)
    worksheet160.write('S22', 'MAW', body)
    worksheet160.write('T22', 'MAP', body)
    worksheet160.write('U22', 'IND', body)
    worksheet160.write('V22', 'ENG', body)
    worksheet160.write('W22', 'SEJ', body)
    worksheet160.write('X22', 'GEO', body)
    worksheet160.write('Y22', 'EKO', body)
    worksheet160.write('Z22', 'SOS', body)
    worksheet160.write('AA22', 'FIS', body)
    worksheet160.write('AB22', 'KIM', body)
    worksheet160.write('AC22', 'BIO', body)
    worksheet160.write('AD22', 'JML', body)

    worksheet160.conditional_format(22, 0, row160+21, 21,
                                    {'type': 'no_errors', 'format': border})

    workbook.close()
    st.success("File siap diunduh!")

    # Tombol unduh file
    with open(file_path, "rb") as f:
        bytes_data = f.read()
    st.download_button(label="Unduh File", data=bytes_data,
                       file_name=file_name)
