import streamlit as st
import pandas as pd
import tempfile
import xlsxwriter

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
st.header("PPLS-RONIN IPS")

col6 = st.container()

with col6:
    KELAS = st.selectbox(
        "KELAS",
        ("--Pilih Kelas--", "PPLS IPS", "RONIN IPS"))

col7 = st.container()

with col7:
    SEMESTER = st.selectbox(
        "SEMESTER",
        ("--Pilih Semester--", "SEMESTER 1", "SEMESTER 2"))

col8 = st.container()

with col8:
    PENILAIAN = st.selectbox(
        "PENILAIAN",
        ("--Pilih Penilaian--", "PENILAIAN TENGAH SEMESTER", "PENILAIAN AKHIR SEMESTER", "PENILAIAN AKHIR TAHUN", "TES EVALUASI"))

TAHUN = st.text_input("Masukkan Tahun Ajaran", value="",
                      placeholder="contoh: 2022-2023")

col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    MTK = st.selectbox(
        "JML. SOAL MAT.",
        ("--Pilih--", 25, 30, 35, 40, 45))

with col2:
    SEJ = st.selectbox(
        "JML. SOAL SEJ.",
        ("--Pilih--", 25, 30, 35, 40, 45))

with col3:
    GEO = st.selectbox(
        "JML. SOAL GEO.",
        ("--Pilih--", 25, 30, 35, 40, 45))

with col4:
    EKO = st.selectbox(
        "JML. SOAL EKO.",
        ("--Pilih--", 25, 30, 35, 40, 45))

with col5:
    SOS = st.selectbox(
        "JML. SOAL SOS.",
        ("--Pilih--", 25, 30, 35, 40, 45))

JML_SOAL_MAT = MTK
JML_SOAL_SEJ = SEJ
JML_SOAL_GEO = GEO
JML_SOAL_EKO = EKO
JML_SOAL_SOS = SOS

kelas = KELAS.replace(" ", "")
semester = SEMESTER
tahun = TAHUN
penilaian = PENILAIAN

uploaded_file = st.file_uploader(
    'Letakkan file excel NILAI STANDAR [LOKASI 101-160]', type='xlsx')

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # 89
    r = df.shape[0]-5
    # 90
    s = df.shape[0]-4
    # 91
    t = df.shape[0]-3
    # 92
    u = df.shape[0]-2

    # JUMLAH PESERTA
    peserta = df.iloc[r, 23]

    # rata-rata jumlah benar
    rata_mat = df.iloc[r, 156]
    rata_sej = df.iloc[r, 157]
    rata_geo = df.iloc[r, 158]
    rata_eko = df.iloc[r, 159]
    rata_sos = df.iloc[r, 160]
    rata_jml = df.iloc[r, 161]

    # rata-rata nilai standar
    rata_Smat = df.iloc[t, 167]
    rata_Ssej = df.iloc[t, 168]
    rata_Sgeo = df.iloc[t, 169]
    rata_Seko = df.iloc[t, 170]
    rata_Ssos = df.iloc[t, 171]
    rata_Sjml = df.iloc[t, 172]

    max_mat = df.iloc[t, 156]
    max_sej = df.iloc[t, 157]
    max_geo = df.iloc[t, 158]
    max_eko = df.iloc[t, 159]
    max_sos = df.iloc[t, 160]
    max_jml = df.iloc[t, 161]

    # max nilai standar
    max_Smat = df.iloc[r, 167]
    max_Ssej = df.iloc[r, 168]
    max_Sgeo = df.iloc[r, 169]
    max_Seko = df.iloc[r, 170]
    max_Ssos = df.iloc[r, 171]
    max_Sjml = df.iloc[r, 172]

    # min jumlah benar
    min_mat = df.iloc[u, 156]
    min_sej = df.iloc[u, 157]
    min_geo = df.iloc[u, 158]
    min_eko = df.iloc[u, 159]
    min_sos = df.iloc[u, 160]
    min_jml = df.iloc[u, 161]

    # min nilai standar
    min_Smat = df.iloc[s, 167]
    min_Ssej = df.iloc[s, 168]
    min_Sgeo = df.iloc[s, 169]
    min_Seko = df.iloc[s, 170]
    min_Ssos = df.iloc[s, 171]
    min_Sjml = df.iloc[s, 172]

    data_jml_benar = {'BIDANG STUDI': ['MATEMATIKA (MAT)', 'SEJARAH (SEJ)', 'GEOGRAFI (GEO)', 'EKONOMI (EKO)', 'SOSIOLOGI (S0S)', 'JUMLAH (JML)'],
                      'TERENDAH': [min_mat, min_sej, min_geo, min_eko, min_sos, min_jml],
                      'RATA-RATA': [rata_mat, rata_sej, rata_geo, rata_eko, rata_sos, rata_jml],
                      'TERTINGGI': [max_mat, max_sej, max_geo, max_eko, max_sos, max_jml]}

    jml_benar = pd.DataFrame(data_jml_benar)

    data_n_standar = {'BIDANG STUDI': ['MATEMATIKA (MAT)', 'SEJARAH (SEJ)', 'GEOGRAFI (GEO)', 'EKONOMI (EKO)', 'SOSIOLOGI (S0S)', 'JUMLAH (JML)'],
                      'TERENDAH': [min_Smat, min_Ssej, min_Sgeo, min_Seko, min_Ssos, min_Sjml],
                      'RATA-RATA': [rata_Smat, rata_Ssej, rata_Sgeo, rata_Seko, rata_Ssos, rata_Sjml],
                      'TERTINGGI': [max_Smat, max_Ssej, max_Sgeo, max_Seko, max_Ssos, max_Sjml]}

    n_standar = pd.DataFrame(data_n_standar)

    data_jml_peserta = {'JUMLAH PESERTA': [peserta]}

    jml_peserta = pd.DataFrame(data_jml_peserta)

    data_jml_soal = {'BIDANG STUDI': ['MAT', 'SEJ', 'GEO', 'EKO', 'SOS'],
                     'JUMLAH': [JML_SOAL_MAT, JML_SOAL_SEJ, JML_SOAL_GEO, JML_SOAL_EKO, JML_SOAL_SOS]}

    jml_soal = pd.DataFrame(data_jml_soal)

    df = df[['LOKASI', 'RANK LOK.', 'RANK NAS.', 'NOMOR NF', 'NAMA SISWA', 'NAMA SEKOLAH', 'KELAS',
             'MAT', 'SEJ', 'GEO', 'EKO', 'SOS', 'JML', 'S_MAT', 'S_SEJ', 'S_GEO', 'S_EKO', 'S_SOS', 'S_JML']]

    # sort setiap lokasi
    sort101 = df[df['LOKASI'] == 101]
    sort102 = df[df['LOKASI'] == 102]
    sort103 = df[df['LOKASI'] == 103]
    sort104 = df[df['LOKASI'] == 104]
    sort105 = df[df['LOKASI'] == 105]
    sort106 = df[df['LOKASI'] == 106]
    sort107 = df[df['LOKASI'] == 107]
    sort108 = df[df['LOKASI'] == 108]
    sort109 = df[df['LOKASI'] == 109]
    sort110 = df[df['LOKASI'] == 110]
    sort111 = df[df['LOKASI'] == 111]
    sort112 = df[df['LOKASI'] == 112]
    sort113 = df[df['LOKASI'] == 113]
    sort114 = df[df['LOKASI'] == 114]
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
    sort139 = df[df['LOKASI'] == 139]
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
    # 104
    sort104_10 = sort104.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort104_10['LOKASI']
    sort104_10 = sort104_10.drop(
        sort104_10[(sort104_10['RANK LOK.'] > 10)].index)
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
    # 114
    sort114_10 = sort114.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort114_10['LOKASI']
    sort114_10 = sort114_10.drop(
        sort114_10[(sort114_10['RANK LOK.'] > 10)].index)
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
    # 139
    sort139_10 = sort139.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort139_10['LOKASI']
    sort139_10 = sort139_10.drop(
        sort139_10[(sort139_10['RANK LOK.'] > 10)].index)
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
    # All 104
    sort104 = sort104.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort104['LOKASI']
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
    # All 114
    sort114 = sort114.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort114['LOKASI']
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
    # All 139
    sort139 = sort139.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort139['LOKASI']
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
    # 104
    row104_10 = sort104_10.shape[0]
    row104 = sort104.shape[0]
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
    # 114
    row114_10 = sort114_10.shape[0]
    row114 = sort114.shape[0]
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
    # 139
    row139_10 = sort139_10.shape[0]
    row139 = sort139.shape[0]
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
    # Path file hasil penyimpanan
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
                       startrow=21,
                       startcol=0,
                       index=False,
                       header=False)

    # Convert the dataframe to an XlsxWriter Excel object cover.
    jml_peserta.to_excel(writer, sheet_name='cover',
                         startrow=21,
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
    # 104
    # Convert the dataframe to an XlsxWriter Excel object.
    sort104_10.to_excel(writer, sheet_name='104',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort104.to_excel(writer, sheet_name='104',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
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
    # 114
    # Convert the dataframe to an XlsxWriter Excel object.
    sort114_10.to_excel(writer, sheet_name='114',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort114.to_excel(writer, sheet_name='114',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
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
    # 139
    # Convert the dataframe to an XlsxWriter Excel object.
    sort139_10.to_excel(writer, sheet_name='139',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort139.to_excel(writer, sheet_name='139',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
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
    worksheet104 = writer.sheets['104']
    worksheet105 = writer.sheets['105']
    worksheet106 = writer.sheets['106']
    worksheet107 = writer.sheets['107']
    worksheet108 = writer.sheets['108']
    worksheet109 = writer.sheets['109']
    worksheet110 = writer.sheets['110']
    worksheet111 = writer.sheets['111']
    worksheet112 = writer.sheets['112']
    worksheet113 = writer.sheets['113']
    worksheet114 = writer.sheets['114']
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
    worksheet139 = writer.sheets['139']
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
    worksheetcover.conditional_format(16, 0, 11, 3,
                                      {'type': 'no_errors', 'format': borderCover})

    worksheetcover.insert_image('F1', r'logo nf.jpg')

    worksheetcover.merge_range('A10:A11', 'BIDANG STUDI', bodyCover)
    worksheetcover.merge_range('B10:B11', 'TERENDAH', bodyCover)
    worksheetcover.merge_range('C10:C11', 'RATA-RATA', bodyCover)
    worksheetcover.merge_range('D10:D11', 'TERTINGGI', bodyCover)
    worksheetcover.merge_range('A20:A21', 'BIDANG STUDI', bodyCover)
    worksheetcover.merge_range('B20:B21', 'TERENDAH', bodyCover)
    worksheetcover.merge_range('C20:C21', 'RATA-RATA', bodyCover)
    worksheetcover.merge_range('D20:D21', 'TERTINGGI', bodyCover)
    worksheetcover.write('F13', 'BIDANG STUDI', bodyCover)
    worksheetcover.merge_range('F20:F21', 'JUMLAH', sub_header1Cover)
    worksheetcover.merge_range('F23:F24', 'PESERTA', sub_header1Cover)
    worksheetcover.write('G13', 'JUMLAH', bodyCover)
    worksheetcover.set_column('A:A', 25.71, centerCover)
    worksheetcover.set_column('B:B', 15, centerCover)
    worksheetcover.set_column('C:C', 15, centerCover)
    worksheetcover.set_column('D:D', 15, centerCover)
    worksheetcover.set_column('F:F', 25.71, centerCover)
    worksheetcover.set_column('G:G', 13, centerCover)
    worksheetcover.merge_range('A1:F3', 'DAFTAR NILAI', titleCover)
    worksheetcover.merge_range(
        'A4:F5', fr'{penilaian}', sub_titleCover)
    worksheetcover.merge_range(
        'A6:F7', fr'{semester} TAHUN {tahun}', headerCover)
    worksheetcover.write('A9', 'JUMLAH BENAR', sub_headerCover)
    worksheetcover.write('A19', 'NILAI STANDAR', sub_headerCover)
    worksheetcover.merge_range('F8:G9', fr'{kelas}', kelasCover)
    worksheetcover.merge_range('F11:G12', 'JUMLAH SOAL', sub_header1Cover)

    worksheetcover.conditional_format(26, 0, 21, 3,
                                      {'type': 'no_errors', 'format': borderCover})

    worksheetcover.conditional_format(17, 6, 13, 5,
                                      {'type': 'no_errors', 'format': borderCover})

    worksheetcover.conditional_format(21, 5, 21, 5,
                                      {'type': 'no_errors', 'format': borderCover})

    # worksheet 101
    worksheet101.insert_image('A1', r'logo resmi nf.jpg')

    worksheet101.set_column('A:A', 7, center)
    worksheet101.set_column('B:B', 6, center)
    worksheet101.set_column('C:C', 18.14, center)
    worksheet101.set_column('D:D', 25, left)
    worksheet101.set_column('E:E', 13.14, left)
    worksheet101.set_column('F:F', 8.57, center)
    worksheet101.set_column('G:R', 5, center)
    worksheet101.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF TAMAN MARGASATWA', title)
    worksheet101.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet101.write('A5', 'LOKASI', header)
    worksheet101.write('B5', 'TOTAL', header)
    worksheet101.merge_range('A4:B4', 'RANK', header)
    worksheet101.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet101.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet101.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet101.merge_range('F4:F5', 'KELAS', header)
    worksheet101.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet101.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet101.write('G5', 'MAT', body)
    worksheet101.write('H5', 'SEJ', body)
    worksheet101.write('I5', 'GEO', body)
    worksheet101.write('J5', 'EKO', body)
    worksheet101.write('K5', 'SOS', body)
    worksheet101.write('L5', 'JML', body)
    worksheet101.write('M5', 'MAT', body)
    worksheet101.write('N5', 'SEJ', body)
    worksheet101.write('O5', 'GEO', body)
    worksheet101.write('P5', 'EKO', body)
    worksheet101.write('Q5', 'SOS', body)
    worksheet101.write('R5', 'JML', body)

    worksheet101.conditional_format(5, 0, row101_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet101.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF TAMAN MARGASATWA', title)
    worksheet101.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet101.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet101.write('A22', 'LOKASI', header)
    worksheet101.write('B22', 'TOTAL', header)
    worksheet101.merge_range('A21:B21', 'RANK', header)
    worksheet101.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet101.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet101.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet101.merge_range('F21:F22', 'KELAS', header)
    worksheet101.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet101.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet101.write('G22', 'MAT', body)
    worksheet101.write('H22', 'SEJ', body)
    worksheet101.write('I22', 'GEO', body)
    worksheet101.write('J22', 'EKO', body)
    worksheet101.write('K22', 'SOS', body)
    worksheet101.write('L22', 'JML', body)
    worksheet101.write('M22', 'MAT', body)
    worksheet101.write('N22', 'SEJ', body)
    worksheet101.write('O22', 'GEO', body)
    worksheet101.write('P22', 'EKO', body)
    worksheet101.write('Q22', 'SOS', body)
    worksheet101.write('R22', 'JML', body)

    worksheet101.conditional_format(22, 0, row101+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 102
    worksheet102.insert_image('A1', r'logo resmi nf.jpg')

    worksheet102.set_column('A:A', 7, center)
    worksheet102.set_column('B:B', 6, center)
    worksheet102.set_column('C:C', 18.14, center)
    worksheet102.set_column('D:D', 25, left)
    worksheet102.set_column('E:E', 13.14, left)
    worksheet102.set_column('F:F', 8.57, center)
    worksheet102.set_column('G:R', 5, center)
    worksheet102.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CEMPAKA', title)
    worksheet102.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet102.write('A5', 'LOKASI', header)
    worksheet102.write('B5', 'TOTAL', header)
    worksheet102.merge_range('A4:B4', 'RANK', header)
    worksheet102.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet102.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet102.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet102.merge_range('F4:F5', 'KELAS', header)
    worksheet102.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet102.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet102.write('G5', 'MAT', body)
    worksheet102.write('H5', 'SEJ', body)
    worksheet102.write('I5', 'GEO', body)
    worksheet102.write('J5', 'EKO', body)
    worksheet102.write('K5', 'SOS', body)
    worksheet102.write('L5', 'JML', body)
    worksheet102.write('M5', 'MAT', body)
    worksheet102.write('N5', 'SEJ', body)
    worksheet102.write('O5', 'GEO', body)
    worksheet102.write('P5', 'EKO', body)
    worksheet102.write('Q5', 'SOS', body)
    worksheet102.write('R5', 'JML', body)

    worksheet102.conditional_format(5, 0, row102_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet102.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CEMPAKA', title)
    worksheet102.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet102.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet102.write('A22', 'LOKASI', header)
    worksheet102.write('B22', 'TOTAL', header)
    worksheet102.merge_range('A21:B21', 'RANK', header)
    worksheet102.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet102.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet102.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet102.merge_range('F21:F22', 'KELAS', header)
    worksheet102.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet102.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet102.write('G22', 'MAT', body)
    worksheet102.write('H22', 'SEJ', body)
    worksheet102.write('I22', 'GEO', body)
    worksheet102.write('J22', 'EKO', body)
    worksheet102.write('K22', 'SOS', body)
    worksheet102.write('L22', 'JML', body)
    worksheet102.write('M22', 'MAT', body)
    worksheet102.write('N22', 'SEJ', body)
    worksheet102.write('O22', 'GEO', body)
    worksheet102.write('P22', 'EKO', body)
    worksheet102.write('Q22', 'SOS', body)
    worksheet102.write('R22', 'JML', body)

    worksheet102.conditional_format(22, 0, row102+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 103
    worksheet103.insert_image('A1', r'logo resmi nf.jpg')

    worksheet103.set_column('A:A', 7, center)
    worksheet103.set_column('B:B', 6, center)
    worksheet103.set_column('C:C', 18.14, center)
    worksheet103.set_column('D:D', 25, left)
    worksheet103.set_column('E:E', 13.14, left)
    worksheet103.set_column('F:F', 8.57, center)
    worksheet103.set_column('G:R', 5, center)
    worksheet103.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PANGKALAN JATI', title)
    worksheet103.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet103.write('A5', 'LOKASI', header)
    worksheet103.write('B5', 'TOTAL', header)
    worksheet103.merge_range('A4:B4', 'RANK', header)
    worksheet103.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet103.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet103.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet103.merge_range('F4:F5', 'KELAS', header)
    worksheet103.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet103.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet103.write('G5', 'MAT', body)
    worksheet103.write('H5', 'SEJ', body)
    worksheet103.write('I5', 'GEO', body)
    worksheet103.write('J5', 'EKO', body)
    worksheet103.write('K5', 'SOS', body)
    worksheet103.write('L5', 'JML', body)
    worksheet103.write('M5', 'MAT', body)
    worksheet103.write('N5', 'SEJ', body)
    worksheet103.write('O5', 'GEO', body)
    worksheet103.write('P5', 'EKO', body)
    worksheet103.write('Q5', 'SOS', body)
    worksheet103.write('R5', 'JML', body)

    worksheet103.conditional_format(5, 0, row103_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet103.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PANGKALAN JATI', title)
    worksheet103.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet103.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet103.write('A22', 'LOKASI', header)
    worksheet103.write('B22', 'TOTAL', header)
    worksheet103.merge_range('A21:B21', 'RANK', header)
    worksheet103.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet103.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet103.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet103.merge_range('F21:F22', 'KELAS', header)
    worksheet103.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet103.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet103.write('G22', 'MAT', body)
    worksheet103.write('H22', 'SEJ', body)
    worksheet103.write('I22', 'GEO', body)
    worksheet103.write('J22', 'EKO', body)
    worksheet103.write('K22', 'SOS', body)
    worksheet103.write('L22', 'JML', body)
    worksheet103.write('M22', 'MAT', body)
    worksheet103.write('N22', 'SEJ', body)
    worksheet103.write('O22', 'GEO', body)
    worksheet103.write('P22', 'EKO', body)
    worksheet103.write('Q22', 'SOS', body)
    worksheet103.write('R22', 'JML', body)

    worksheet103.conditional_format(22, 0, row103+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 104
    worksheet104.insert_image('A1', r'logo resmi nf.jpg')

    worksheet104.set_column('A:A', 7, center)
    worksheet104.set_column('B:B', 6, center)
    worksheet104.set_column('C:C', 18.14, center)
    worksheet104.set_column('D:D', 25, left)
    worksheet104.set_column('E:E', 13.14, left)
    worksheet104.set_column('F:F', 8.57, center)
    worksheet104.set_column('G:R', 5, center)
    worksheet104.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KENARI', title)
    worksheet104.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet104.write('A5', 'LOKASI', header)
    worksheet104.write('B5', 'TOTAL', header)
    worksheet104.merge_range('A4:B4', 'RANK', header)
    worksheet104.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet104.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet104.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet104.merge_range('F4:F5', 'KELAS', header)
    worksheet104.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet104.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet104.write('G5', 'MAT', body)
    worksheet104.write('H5', 'SEJ', body)
    worksheet104.write('I5', 'GEO', body)
    worksheet104.write('J5', 'EKO', body)
    worksheet104.write('K5', 'SOS', body)
    worksheet104.write('L5', 'JML', body)
    worksheet104.write('M5', 'MAT', body)
    worksheet104.write('N5', 'SEJ', body)
    worksheet104.write('O5', 'GEO', body)
    worksheet104.write('P5', 'EKO', body)
    worksheet104.write('Q5', 'SOS', body)
    worksheet104.write('R5', 'JML', body)

    worksheet104.conditional_format(5, 0, row104_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet104.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF KENARI', title)
    worksheet104.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet104.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet104.write('A22', 'LOKASI', header)
    worksheet104.write('B22', 'TOTAL', header)
    worksheet104.merge_range('A21:B21', 'RANK', header)
    worksheet104.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet104.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet104.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet104.merge_range('F21:F22', 'KELAS', header)
    worksheet104.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet104.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet104.write('G22', 'MAT', body)
    worksheet104.write('H22', 'SEJ', body)
    worksheet104.write('I22', 'GEO', body)
    worksheet104.write('J22', 'EKO', body)
    worksheet104.write('K22', 'SOS', body)
    worksheet104.write('L22', 'JML', body)
    worksheet104.write('M22', 'MAT', body)
    worksheet104.write('N22', 'SEJ', body)
    worksheet104.write('O22', 'GEO', body)
    worksheet104.write('P22', 'EKO', body)
    worksheet104.write('Q22', 'SOS', body)
    worksheet104.write('R22', 'JML', body)

    worksheet104.conditional_format(22, 0, row104+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 105
    worksheet105.insert_image('A1', r'logo resmi nf.jpg')

    worksheet105.set_column('A:A', 7, center)
    worksheet105.set_column('B:B', 6, center)
    worksheet105.set_column('C:C', 18.14, center)
    worksheet105.set_column('D:D', 25, left)
    worksheet105.set_column('E:E', 13.14, left)
    worksheet105.set_column('F:F', 8.57, center)
    worksheet105.set_column('G:R', 5, center)
    worksheet105.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BUARAN', title)
    worksheet105.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet105.write('A5', 'LOKASI', header)
    worksheet105.write('B5', 'TOTAL', header)
    worksheet105.merge_range('A4:B4', 'RANK', header)
    worksheet105.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet105.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet105.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet105.merge_range('F4:F5', 'KELAS', header)
    worksheet105.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet105.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet105.write('G5', 'MAT', body)
    worksheet105.write('H5', 'SEJ', body)
    worksheet105.write('I5', 'GEO', body)
    worksheet105.write('J5', 'EKO', body)
    worksheet105.write('K5', 'SOS', body)
    worksheet105.write('L5', 'JML', body)
    worksheet105.write('M5', 'MAT', body)
    worksheet105.write('N5', 'SEJ', body)
    worksheet105.write('O5', 'GEO', body)
    worksheet105.write('P5', 'EKO', body)
    worksheet105.write('Q5', 'SOS', body)
    worksheet105.write('R5', 'JML', body)

    worksheet105.conditional_format(5, 0, row105_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet105.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF BUARAN', title)
    worksheet105.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet105.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet105.write('A22', 'LOKASI', header)
    worksheet105.write('B22', 'TOTAL', header)
    worksheet105.merge_range('A21:B21', 'RANK', header)
    worksheet105.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet105.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet105.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet105.merge_range('F21:F22', 'KELAS', header)
    worksheet105.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet105.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet105.write('G22', 'MAT', body)
    worksheet105.write('H22', 'SEJ', body)
    worksheet105.write('I22', 'GEO', body)
    worksheet105.write('J22', 'EKO', body)
    worksheet105.write('K22', 'SOS', body)
    worksheet105.write('L22', 'JML', body)
    worksheet105.write('M22', 'MAT', body)
    worksheet105.write('N22', 'SEJ', body)
    worksheet105.write('O22', 'GEO', body)
    worksheet105.write('P22', 'EKO', body)
    worksheet105.write('Q22', 'SOS', body)
    worksheet105.write('R22', 'JML', body)

    worksheet105.conditional_format(22, 0, row105+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 106
    worksheet106.insert_image('A1', r'logo resmi nf.jpg')

    worksheet106.set_column('A:A', 7, center)
    worksheet106.set_column('B:B', 6, center)
    worksheet106.set_column('C:C', 18.14, center)
    worksheet106.set_column('D:D', 25, left)
    worksheet106.set_column('E:E', 13.14, left)
    worksheet106.set_column('F:F', 8.57, center)
    worksheet106.set_column('G:R', 5, center)
    worksheet106.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF HEK-KRAMAT JATI', title)
    worksheet106.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet106.write('A5', 'LOKASI', header)
    worksheet106.write('B5', 'TOTAL', header)
    worksheet106.merge_range('A4:B4', 'RANK', header)
    worksheet106.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet106.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet106.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet106.merge_range('F4:F5', 'KELAS', header)
    worksheet106.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet106.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet106.write('G5', 'MAT', body)
    worksheet106.write('H5', 'SEJ', body)
    worksheet106.write('I5', 'GEO', body)
    worksheet106.write('J5', 'EKO', body)
    worksheet106.write('K5', 'SOS', body)
    worksheet106.write('L5', 'JML', body)
    worksheet106.write('M5', 'MAT', body)
    worksheet106.write('N5', 'SEJ', body)
    worksheet106.write('O5', 'GEO', body)
    worksheet106.write('P5', 'EKO', body)
    worksheet106.write('Q5', 'SOS', body)
    worksheet106.write('R5', 'JML', body)

    worksheet106.conditional_format(5, 0, row106_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet106.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF HEK-KRAMAT JATI', title)
    worksheet106.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet106.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet106.write('A22', 'LOKASI', header)
    worksheet106.write('B22', 'TOTAL', header)
    worksheet106.merge_range('A21:B21', 'RANK', header)
    worksheet106.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet106.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet106.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet106.merge_range('F21:F22', 'KELAS', header)
    worksheet106.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet106.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet106.write('G22', 'MAT', body)
    worksheet106.write('H22', 'SEJ', body)
    worksheet106.write('I22', 'GEO', body)
    worksheet106.write('J22', 'EKO', body)
    worksheet106.write('K22', 'SOS', body)
    worksheet106.write('L22', 'JML', body)
    worksheet106.write('M22', 'MAT', body)
    worksheet106.write('N22', 'SEJ', body)
    worksheet106.write('O22', 'GEO', body)
    worksheet106.write('P22', 'EKO', body)
    worksheet106.write('Q22', 'SOS', body)
    worksheet106.write('R22', 'JML', body)

    worksheet106.conditional_format(22, 0, row106+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 107
    worksheet107.insert_image('A1', r'logo resmi nf.jpg')

    worksheet107.set_column('A:A', 7, center)
    worksheet107.set_column('B:B', 6, center)
    worksheet107.set_column('C:C', 18.14, center)
    worksheet107.set_column('D:D', 25, left)
    worksheet107.set_column('E:E', 13.14, left)
    worksheet107.set_column('F:F', 8.57, center)
    worksheet107.set_column('G:R', 5, center)
    worksheet107.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MAMPANG', title)
    worksheet107.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet107.write('A5', 'LOKASI', header)
    worksheet107.write('B5', 'TOTAL', header)
    worksheet107.merge_range('A4:B4', 'RANK', header)
    worksheet107.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet107.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet107.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet107.merge_range('F4:F5', 'KELAS', header)
    worksheet107.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet107.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet107.write('G5', 'MAT', body)
    worksheet107.write('H5', 'SEJ', body)
    worksheet107.write('I5', 'GEO', body)
    worksheet107.write('J5', 'EKO', body)
    worksheet107.write('K5', 'SOS', body)
    worksheet107.write('L5', 'JML', body)
    worksheet107.write('M5', 'MAT', body)
    worksheet107.write('N5', 'SEJ', body)
    worksheet107.write('O5', 'GEO', body)
    worksheet107.write('P5', 'EKO', body)
    worksheet107.write('Q5', 'SOS', body)
    worksheet107.write('R5', 'JML', body)

    worksheet107.conditional_format(5, 0, row107_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet107.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF MAMPANG', title)
    worksheet107.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet107.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet107.write('A22', 'LOKASI', header)
    worksheet107.write('B22', 'TOTAL', header)
    worksheet107.merge_range('A21:B21', 'RANK', header)
    worksheet107.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet107.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet107.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet107.merge_range('F21:F22', 'KELAS', header)
    worksheet107.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet107.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet107.write('G22', 'MAT', body)
    worksheet107.write('H22', 'SEJ', body)
    worksheet107.write('I22', 'GEO', body)
    worksheet107.write('J22', 'EKO', body)
    worksheet107.write('K22', 'SOS', body)
    worksheet107.write('L22', 'JML', body)
    worksheet107.write('M22', 'MAT', body)
    worksheet107.write('N22', 'SEJ', body)
    worksheet107.write('O22', 'GEO', body)
    worksheet107.write('P22', 'EKO', body)
    worksheet107.write('Q22', 'SOS', body)
    worksheet107.write('R22', 'JML', body)

    worksheet107.conditional_format(22, 0, row107+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 108
    worksheet108.insert_image('A1', r'logo resmi nf.jpg')

    worksheet108.set_column('A:A', 7, center)
    worksheet108.set_column('B:B', 6, center)
    worksheet108.set_column('C:C', 18.14, center)
    worksheet108.set_column('D:D', 25, left)
    worksheet108.set_column('E:E', 13.14, left)
    worksheet108.set_column('F:F', 8.57, center)
    worksheet108.set_column('G:R', 5, center)
    worksheet108.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PALMERAH', title)
    worksheet108.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet108.write('A5', 'LOKASI', header)
    worksheet108.write('B5', 'TOTAL', header)
    worksheet108.merge_range('A4:B4', 'RANK', header)
    worksheet108.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet108.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet108.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet108.merge_range('F4:F5', 'KELAS', header)
    worksheet108.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet108.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet108.write('G5', 'MAT', body)
    worksheet108.write('H5', 'SEJ', body)
    worksheet108.write('I5', 'GEO', body)
    worksheet108.write('J5', 'EKO', body)
    worksheet108.write('K5', 'SOS', body)
    worksheet108.write('L5', 'JML', body)
    worksheet108.write('M5', 'MAT', body)
    worksheet108.write('N5', 'SEJ', body)
    worksheet108.write('O5', 'GEO', body)
    worksheet108.write('P5', 'EKO', body)
    worksheet108.write('Q5', 'SOS', body)
    worksheet108.write('R5', 'JML', body)

    worksheet108.conditional_format(5, 0, row108_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet108.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PALMERAH', title)
    worksheet108.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet108.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet108.write('A22', 'LOKASI', header)
    worksheet108.write('B22', 'TOTAL', header)
    worksheet108.merge_range('A21:B21', 'RANK', header)
    worksheet108.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet108.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet108.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet108.merge_range('F21:F22', 'KELAS', header)
    worksheet108.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet108.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet108.write('G22', 'MAT', body)
    worksheet108.write('H22', 'SEJ', body)
    worksheet108.write('I22', 'GEO', body)
    worksheet108.write('J22', 'EKO', body)
    worksheet108.write('K22', 'SOS', body)
    worksheet108.write('L22', 'JML', body)
    worksheet108.write('M22', 'MAT', body)
    worksheet108.write('N22', 'SEJ', body)
    worksheet108.write('O22', 'GEO', body)
    worksheet108.write('P22', 'EKO', body)
    worksheet108.write('Q22', 'SOS', body)
    worksheet108.write('R22', 'JML', body)

    worksheet108.conditional_format(22, 0, row108+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 109
    worksheet109.insert_image('A1', r'logo resmi nf.jpg')

    worksheet109.set_column('A:A', 7, center)
    worksheet109.set_column('B:B', 6, center)
    worksheet109.set_column('C:C', 18.14, center)
    worksheet109.set_column('D:D', 25, left)
    worksheet109.set_column('E:E', 13.14, left)
    worksheet109.set_column('F:F', 8.57, center)
    worksheet109.set_column('G:R', 5, center)
    worksheet109.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PASAR MINGGU', title)
    worksheet109.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet109.write('A5', 'LOKASI', header)
    worksheet109.write('B5', 'TOTAL', header)
    worksheet109.merge_range('A4:B4', 'RANK', header)
    worksheet109.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet109.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet109.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet109.merge_range('F4:F5', 'KELAS', header)
    worksheet109.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet109.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet109.write('G5', 'MAT', body)
    worksheet109.write('H5', 'SEJ', body)
    worksheet109.write('I5', 'GEO', body)
    worksheet109.write('J5', 'EKO', body)
    worksheet109.write('K5', 'SOS', body)
    worksheet109.write('L5', 'JML', body)
    worksheet109.write('M5', 'MAT', body)
    worksheet109.write('N5', 'SEJ', body)
    worksheet109.write('O5', 'GEO', body)
    worksheet109.write('P5', 'EKO', body)
    worksheet109.write('Q5', 'SOS', body)
    worksheet109.write('R5', 'JML', body)

    worksheet109.conditional_format(5, 0, row109_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet109.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PASAR MINGGU', title)
    worksheet109.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet109.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet109.write('A22', 'LOKASI', header)
    worksheet109.write('B22', 'TOTAL', header)
    worksheet109.merge_range('A21:B21', 'RANK', header)
    worksheet109.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet109.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet109.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet109.merge_range('F21:F22', 'KELAS', header)
    worksheet109.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet109.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet109.write('G22', 'MAT', body)
    worksheet109.write('H22', 'SEJ', body)
    worksheet109.write('I22', 'GEO', body)
    worksheet109.write('J22', 'EKO', body)
    worksheet109.write('K22', 'SOS', body)
    worksheet109.write('L22', 'JML', body)
    worksheet109.write('M22', 'MAT', body)
    worksheet109.write('N22', 'SEJ', body)
    worksheet109.write('O22', 'GEO', body)
    worksheet109.write('P22', 'EKO', body)
    worksheet109.write('Q22', 'SOS', body)
    worksheet109.write('R22', 'JML', body)

    worksheet109.conditional_format(22, 0, row109+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 110
    worksheet110.insert_image('A1', r'logo resmi nf.jpg')

    worksheet110.set_column('A:A', 7, center)
    worksheet110.set_column('B:B', 6, center)
    worksheet110.set_column('C:C', 18.14, center)
    worksheet110.set_column('D:D', 25, left)
    worksheet110.set_column('E:E', 13.14, left)
    worksheet110.set_column('F:F', 8.57, center)
    worksheet110.set_column('G:R', 5, center)
    worksheet110.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BINTARO', title)
    worksheet110.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet110.write('A5', 'LOKASI', header)
    worksheet110.write('B5', 'TOTAL', header)
    worksheet110.merge_range('A4:B4', 'RANK', header)
    worksheet110.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet110.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet110.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet110.merge_range('F4:F5', 'KELAS', header)
    worksheet110.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet110.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet110.write('G5', 'MAT', body)
    worksheet110.write('H5', 'SEJ', body)
    worksheet110.write('I5', 'GEO', body)
    worksheet110.write('J5', 'EKO', body)
    worksheet110.write('K5', 'SOS', body)
    worksheet110.write('L5', 'JML', body)
    worksheet110.write('M5', 'MAT', body)
    worksheet110.write('N5', 'SEJ', body)
    worksheet110.write('O5', 'GEO', body)
    worksheet110.write('P5', 'EKO', body)
    worksheet110.write('Q5', 'SOS', body)
    worksheet110.write('R5', 'JML', body)

    worksheet110.conditional_format(5, 0, row110_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet110.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF BINTARO', title)
    worksheet110.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet110.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet110.write('A22', 'LOKASI', header)
    worksheet110.write('B22', 'TOTAL', header)
    worksheet110.merge_range('A21:B21', 'RANK', header)
    worksheet110.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet110.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet110.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet110.merge_range('F21:F22', 'KELAS', header)
    worksheet110.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet110.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet110.write('G22', 'MAT', body)
    worksheet110.write('H22', 'SEJ', body)
    worksheet110.write('I22', 'GEO', body)
    worksheet110.write('J22', 'EKO', body)
    worksheet110.write('K22', 'SOS', body)
    worksheet110.write('L22', 'JML', body)
    worksheet110.write('M22', 'MAT', body)
    worksheet110.write('N22', 'SEJ', body)
    worksheet110.write('O22', 'GEO', body)
    worksheet110.write('P22', 'EKO', body)
    worksheet110.write('Q22', 'SOS', body)
    worksheet110.write('R22', 'JML', body)

    worksheet110.conditional_format(22, 0, row110+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 111
    worksheet111.insert_image('A1', r'logo resmi nf.jpg')

    worksheet111.set_column('A:A', 7, center)
    worksheet111.set_column('B:B', 6, center)
    worksheet111.set_column('C:C', 18.14, center)
    worksheet111.set_column('D:D', 25, left)
    worksheet111.set_column('E:E', 13.14, left)
    worksheet111.set_column('F:F', 8.57, center)
    worksheet111.set_column('G:R', 5, center)
    worksheet111.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF LAMPIRI', title)
    worksheet111.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet111.write('A5', 'LOKASI', header)
    worksheet111.write('B5', 'TOTAL', header)
    worksheet111.merge_range('A4:B4', 'RANK', header)
    worksheet111.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet111.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet111.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet111.merge_range('F4:F5', 'KELAS', header)
    worksheet111.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet111.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet111.write('G5', 'MAT', body)
    worksheet111.write('H5', 'SEJ', body)
    worksheet111.write('I5', 'GEO', body)
    worksheet111.write('J5', 'EKO', body)
    worksheet111.write('K5', 'SOS', body)
    worksheet111.write('L5', 'JML', body)
    worksheet111.write('M5', 'MAT', body)
    worksheet111.write('N5', 'SEJ', body)
    worksheet111.write('O5', 'GEO', body)
    worksheet111.write('P5', 'EKO', body)
    worksheet111.write('Q5', 'SOS', body)
    worksheet111.write('R5', 'JML', body)

    worksheet111.conditional_format(5, 0, row111_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet111.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF LAMPIRI', title)
    worksheet111.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet111.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet111.write('A22', 'LOKASI', header)
    worksheet111.write('B22', 'TOTAL', header)
    worksheet111.merge_range('A21:B21', 'RANK', header)
    worksheet111.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet111.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet111.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet111.merge_range('F21:F22', 'KELAS', header)
    worksheet111.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet111.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet111.write('G22', 'MAT', body)
    worksheet111.write('H22', 'SEJ', body)
    worksheet111.write('I22', 'GEO', body)
    worksheet111.write('J22', 'EKO', body)
    worksheet111.write('K22', 'SOS', body)
    worksheet111.write('L22', 'JML', body)
    worksheet111.write('M22', 'MAT', body)
    worksheet111.write('N22', 'SEJ', body)
    worksheet111.write('O22', 'GEO', body)
    worksheet111.write('P22', 'EKO', body)
    worksheet111.write('Q22', 'SOS', body)
    worksheet111.write('R22', 'JML', body)

    worksheet111.conditional_format(22, 0, row111+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 112
    worksheet112.insert_image('A1', r'logo resmi nf.jpg')

    worksheet112.set_column('A:A', 7, center)
    worksheet112.set_column('B:B', 6, center)
    worksheet112.set_column('C:C', 18.14, center)
    worksheet112.set_column('D:D', 25, left)
    worksheet112.set_column('E:E', 13.14, left)
    worksheet112.set_column('F:F', 8.57, center)
    worksheet112.set_column('G:R', 5, center)
    worksheet112.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PONDOK BAMBU', title)
    worksheet112.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet112.write('A5', 'LOKASI', header)
    worksheet112.write('B5', 'TOTAL', header)
    worksheet112.merge_range('A4:B4', 'RANK', header)
    worksheet112.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet112.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet112.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet112.merge_range('F4:F5', 'KELAS', header)
    worksheet112.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet112.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet112.write('G5', 'MAT', body)
    worksheet112.write('H5', 'SEJ', body)
    worksheet112.write('I5', 'GEO', body)
    worksheet112.write('J5', 'EKO', body)
    worksheet112.write('K5', 'SOS', body)
    worksheet112.write('L5', 'JML', body)
    worksheet112.write('M5', 'MAT', body)
    worksheet112.write('N5', 'SEJ', body)
    worksheet112.write('O5', 'GEO', body)
    worksheet112.write('P5', 'EKO', body)
    worksheet112.write('Q5', 'SOS', body)
    worksheet112.write('R5', 'JML', body)

    worksheet112.conditional_format(5, 0, row112_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet112.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PONDOK BAMBU', title)
    worksheet112.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet112.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet112.write('A22', 'LOKASI', header)
    worksheet112.write('B22', 'TOTAL', header)
    worksheet112.merge_range('A21:B21', 'RANK', header)
    worksheet112.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet112.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet112.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet112.merge_range('F21:F22', 'KELAS', header)
    worksheet112.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet112.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet112.write('G22', 'MAT', body)
    worksheet112.write('H22', 'SEJ', body)
    worksheet112.write('I22', 'GEO', body)
    worksheet112.write('J22', 'EKO', body)
    worksheet112.write('K22', 'SOS', body)
    worksheet112.write('L22', 'JML', body)
    worksheet112.write('M22', 'MAT', body)
    worksheet112.write('N22', 'SEJ', body)
    worksheet112.write('O22', 'GEO', body)
    worksheet112.write('P22', 'EKO', body)
    worksheet112.write('Q22', 'SOS', body)
    worksheet112.write('R22', 'JML', body)

    worksheet112.conditional_format(22, 0, row112+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 113
    worksheet113.insert_image('A1', r'logo resmi nf.jpg')

    worksheet113.set_column('A:A', 7, center)
    worksheet113.set_column('B:B', 6, center)
    worksheet113.set_column('C:C', 18.14, center)
    worksheet113.set_column('D:D', 25, left)
    worksheet113.set_column('E:E', 13.14, left)
    worksheet113.set_column('F:F', 8.57, center)
    worksheet113.set_column('G:R', 5, center)
    worksheet113.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF RAWA BADAK', title)
    worksheet113.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet113.write('A5', 'LOKASI', header)
    worksheet113.write('B5', 'TOTAL', header)
    worksheet113.merge_range('A4:B4', 'RANK', header)
    worksheet113.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet113.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet113.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet113.merge_range('F4:F5', 'KELAS', header)
    worksheet113.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet113.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet113.write('G5', 'MAT', body)
    worksheet113.write('H5', 'SEJ', body)
    worksheet113.write('I5', 'GEO', body)
    worksheet113.write('J5', 'EKO', body)
    worksheet113.write('K5', 'SOS', body)
    worksheet113.write('L5', 'JML', body)
    worksheet113.write('M5', 'MAT', body)
    worksheet113.write('N5', 'SEJ', body)
    worksheet113.write('O5', 'GEO', body)
    worksheet113.write('P5', 'EKO', body)
    worksheet113.write('Q5', 'SOS', body)
    worksheet113.write('R5', 'JML', body)

    worksheet113.conditional_format(5, 0, row113_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet113.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF RAWA BADAK', title)
    worksheet113.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet113.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet113.write('A22', 'LOKASI', header)
    worksheet113.write('B22', 'TOTAL', header)
    worksheet113.merge_range('A21:B21', 'RANK', header)
    worksheet113.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet113.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet113.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet113.merge_range('F21:F22', 'KELAS', header)
    worksheet113.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet113.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet113.write('G22', 'MAT', body)
    worksheet113.write('H22', 'SEJ', body)
    worksheet113.write('I22', 'GEO', body)
    worksheet113.write('J22', 'EKO', body)
    worksheet113.write('K22', 'SOS', body)
    worksheet113.write('L22', 'JML', body)
    worksheet113.write('M22', 'MAT', body)
    worksheet113.write('N22', 'SEJ', body)
    worksheet113.write('O22', 'GEO', body)
    worksheet113.write('P22', 'EKO', body)
    worksheet113.write('Q22', 'SOS', body)
    worksheet113.write('R22', 'JML', body)

    worksheet113.conditional_format(22, 0, row113+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 114
    worksheet114.insert_image('A1', r'logo resmi nf.jpg')

    worksheet114.set_column('A:A', 7, center)
    worksheet114.set_column('B:B', 6, center)
    worksheet114.set_column('C:C', 18.14, center)
    worksheet114.set_column('D:D', 25, left)
    worksheet114.set_column('E:E', 13.14, left)
    worksheet114.set_column('F:F', 8.57, center)
    worksheet114.set_column('G:R', 5, center)
    worksheet114.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PASAR REBO', title)
    worksheet114.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet114.write('A5', 'LOKASI', header)
    worksheet114.write('B5', 'TOTAL', header)
    worksheet114.merge_range('A4:B4', 'RANK', header)
    worksheet114.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet114.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet114.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet114.merge_range('F4:F5', 'KELAS', header)
    worksheet114.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet114.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet114.write('G5', 'MAT', body)
    worksheet114.write('H5', 'SEJ', body)
    worksheet114.write('I5', 'GEO', body)
    worksheet114.write('J5', 'EKO', body)
    worksheet114.write('K5', 'SOS', body)
    worksheet114.write('L5', 'JML', body)
    worksheet114.write('M5', 'MAT', body)
    worksheet114.write('N5', 'SEJ', body)
    worksheet114.write('O5', 'GEO', body)
    worksheet114.write('P5', 'EKO', body)
    worksheet114.write('Q5', 'SOS', body)
    worksheet114.write('R5', 'JML', body)

    worksheet114.conditional_format(5, 0, row114_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet114.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PASAR REBO', title)
    worksheet114.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet114.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet114.write('A22', 'LOKASI', header)
    worksheet114.write('B22', 'TOTAL', header)
    worksheet114.merge_range('A21:B21', 'RANK', header)
    worksheet114.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet114.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet114.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet114.merge_range('F21:F22', 'KELAS', header)
    worksheet114.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet114.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet114.write('G22', 'MAT', body)
    worksheet114.write('H22', 'SEJ', body)
    worksheet114.write('I22', 'GEO', body)
    worksheet114.write('J22', 'EKO', body)
    worksheet114.write('K22', 'SOS', body)
    worksheet114.write('L22', 'JML', body)
    worksheet114.write('M22', 'MAT', body)
    worksheet114.write('N22', 'SEJ', body)
    worksheet114.write('O22', 'GEO', body)
    worksheet114.write('P22', 'EKO', body)
    worksheet114.write('Q22', 'SOS', body)
    worksheet114.write('R22', 'JML', body)

    worksheet114.conditional_format(22, 0, row114+21, 17,
                                    {'type': 'no_errors', 'format': border})
    # worksheet 115
    worksheet115.insert_image('A1', r'logo resmi nf.jpg')

    worksheet115.set_column('A:A', 7, center)
    worksheet115.set_column('B:B', 6, center)
    worksheet115.set_column('C:C', 18.14, center)
    worksheet115.set_column('D:D', 25, left)
    worksheet115.set_column('E:E', 13.14, left)
    worksheet115.set_column('F:F', 8.57, center)
    worksheet115.set_column('G:R', 5, center)
    worksheet115.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF RAWAMANGUN', title)
    worksheet115.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet115.write('A5', 'LOKASI', header)
    worksheet115.write('B5', 'TOTAL', header)
    worksheet115.merge_range('A4:B4', 'RANK', header)
    worksheet115.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet115.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet115.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet115.merge_range('F4:F5', 'KELAS', header)
    worksheet115.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet115.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet115.write('G5', 'MAT', body)
    worksheet115.write('H5', 'SEJ', body)
    worksheet115.write('I5', 'GEO', body)
    worksheet115.write('J5', 'EKO', body)
    worksheet115.write('K5', 'SOS', body)
    worksheet115.write('L5', 'JML', body)
    worksheet115.write('M5', 'MAT', body)
    worksheet115.write('N5', 'SEJ', body)
    worksheet115.write('O5', 'GEO', body)
    worksheet115.write('P5', 'EKO', body)
    worksheet115.write('Q5', 'SOS', body)
    worksheet115.write('R5', 'JML', body)

    worksheet115.conditional_format(5, 0, row115_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet115.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF RAWAMANGUN', title)
    worksheet115.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet115.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet115.write('A22', 'LOKASI', header)
    worksheet115.write('B22', 'TOTAL', header)
    worksheet115.merge_range('A21:B21', 'RANK', header)
    worksheet115.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet115.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet115.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet115.merge_range('F21:F22', 'KELAS', header)
    worksheet115.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet115.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet115.write('G22', 'MAT', body)
    worksheet115.write('H22', 'SEJ', body)
    worksheet115.write('I22', 'GEO', body)
    worksheet115.write('J22', 'EKO', body)
    worksheet115.write('K22', 'SOS', body)
    worksheet115.write('L22', 'JML', body)
    worksheet115.write('M22', 'MAT', body)
    worksheet115.write('N22', 'SEJ', body)
    worksheet115.write('O22', 'GEO', body)
    worksheet115.write('P22', 'EKO', body)
    worksheet115.write('Q22', 'SOS', body)
    worksheet115.write('R22', 'JML', body)

    worksheet115.conditional_format(22, 0, row115+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 116
    worksheet116.insert_image('A1', r'logo resmi nf.jpg')

    worksheet116.set_column('A:A', 7, center)
    worksheet116.set_column('B:B', 6, center)
    worksheet116.set_column('C:C', 18.14, center)
    worksheet116.set_column('D:D', 25, left)
    worksheet116.set_column('E:E', 13.14, left)
    worksheet116.set_column('F:F', 8.57, center)
    worksheet116.set_column('G:R', 5, center)
    worksheet116.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIRACAS', title)
    worksheet116.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet116.write('A5', 'LOKASI', header)
    worksheet116.write('B5', 'TOTAL', header)
    worksheet116.merge_range('A4:B4', 'RANK', header)
    worksheet116.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet116.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet116.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet116.merge_range('F4:F5', 'KELAS', header)
    worksheet116.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet116.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet116.write('G5', 'MAT', body)
    worksheet116.write('H5', 'SEJ', body)
    worksheet116.write('I5', 'GEO', body)
    worksheet116.write('J5', 'EKO', body)
    worksheet116.write('K5', 'SOS', body)
    worksheet116.write('L5', 'JML', body)
    worksheet116.write('M5', 'MAT', body)
    worksheet116.write('N5', 'SEJ', body)
    worksheet116.write('O5', 'GEO', body)
    worksheet116.write('P5', 'EKO', body)
    worksheet116.write('Q5', 'SOS', body)
    worksheet116.write('R5', 'JML', body)

    worksheet116.conditional_format(5, 0, row116_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet116.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CIRACAS', title)
    worksheet116.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet116.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet116.write('A22', 'LOKASI', header)
    worksheet116.write('B22', 'TOTAL', header)
    worksheet116.merge_range('A21:B21', 'RANK', header)
    worksheet116.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet116.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet116.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet116.merge_range('F21:F22', 'KELAS', header)
    worksheet116.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet116.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet116.write('G22', 'MAT', body)
    worksheet116.write('H22', 'SEJ', body)
    worksheet116.write('I22', 'GEO', body)
    worksheet116.write('J22', 'EKO', body)
    worksheet116.write('K22', 'SOS', body)
    worksheet116.write('L22', 'JML', body)
    worksheet116.write('M22', 'MAT', body)
    worksheet116.write('N22', 'SEJ', body)
    worksheet116.write('O22', 'GEO', body)
    worksheet116.write('P22', 'EKO', body)
    worksheet116.write('Q22', 'SOS', body)
    worksheet116.write('R22', 'JML', body)

    worksheet116.conditional_format(22, 0, row116+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 117
    worksheet117.insert_image('A1', r'logo resmi nf.jpg')

    worksheet117.set_column('A:A', 7, center)
    worksheet117.set_column('B:B', 6, center)
    worksheet117.set_column('C:C', 18.14, center)
    worksheet117.set_column('D:D', 25, left)
    worksheet117.set_column('E:E', 13.14, left)
    worksheet117.set_column('F:F', 8.57, center)
    worksheet117.set_column('G:R', 5, center)
    worksheet117.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KAMPUNG MELAYU', title)
    worksheet117.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet117.write('A5', 'LOKASI', header)
    worksheet117.write('B5', 'TOTAL', header)
    worksheet117.merge_range('A4:B4', 'RANK', header)
    worksheet117.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet117.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet117.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet117.merge_range('F4:F5', 'KELAS', header)
    worksheet117.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet117.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet117.write('G5', 'MAT', body)
    worksheet117.write('H5', 'SEJ', body)
    worksheet117.write('I5', 'GEO', body)
    worksheet117.write('J5', 'EKO', body)
    worksheet117.write('K5', 'SOS', body)
    worksheet117.write('L5', 'JML', body)
    worksheet117.write('M5', 'MAT', body)
    worksheet117.write('N5', 'SEJ', body)
    worksheet117.write('O5', 'GEO', body)
    worksheet117.write('P5', 'EKO', body)
    worksheet117.write('Q5', 'SOS', body)
    worksheet117.write('R5', 'JML', body)

    worksheet117.conditional_format(5, 0, row117_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet117.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF KAMPUNG MELAYU', title)
    worksheet117.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet117.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet117.write('A22', 'LOKASI', header)
    worksheet117.write('B22', 'TOTAL', header)
    worksheet117.merge_range('A21:B21', 'RANK', header)
    worksheet117.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet117.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet117.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet117.merge_range('F21:F22', 'KELAS', header)
    worksheet117.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet117.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet117.write('G22', 'MAT', body)
    worksheet117.write('H22', 'SEJ', body)
    worksheet117.write('I22', 'GEO', body)
    worksheet117.write('J22', 'EKO', body)
    worksheet117.write('K22', 'SOS', body)
    worksheet117.write('L22', 'JML', body)
    worksheet117.write('M22', 'MAT', body)
    worksheet117.write('N22', 'SEJ', body)
    worksheet117.write('O22', 'GEO', body)
    worksheet117.write('P22', 'EKO', body)
    worksheet117.write('Q22', 'SOS', body)
    worksheet117.write('R22', 'JML', body)

    worksheet117.conditional_format(22, 0, row117+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 118
    worksheet118.insert_image('A1', r'logo resmi nf.jpg')

    worksheet118.set_column('A:A', 7, center)
    worksheet118.set_column('B:B', 6, center)
    worksheet118.set_column('C:C', 18.14, center)
    worksheet118.set_column('D:D', 25, left)
    worksheet118.set_column('E:E', 13.14, left)
    worksheet118.set_column('F:F', 8.57, center)
    worksheet118.set_column('G:R', 5, center)
    worksheet118.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF AKSES UI', title)
    worksheet118.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet118.write('A5', 'LOKASI', header)
    worksheet118.write('B5', 'TOTAL', header)
    worksheet118.merge_range('A4:B4', 'RANK', header)
    worksheet118.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet118.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet118.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet118.merge_range('F4:F5', 'KELAS', header)
    worksheet118.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet118.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet118.write('G5', 'MAT', body)
    worksheet118.write('H5', 'SEJ', body)
    worksheet118.write('I5', 'GEO', body)
    worksheet118.write('J5', 'EKO', body)
    worksheet118.write('K5', 'SOS', body)
    worksheet118.write('L5', 'JML', body)
    worksheet118.write('M5', 'MAT', body)
    worksheet118.write('N5', 'SEJ', body)
    worksheet118.write('O5', 'GEO', body)
    worksheet118.write('P5', 'EKO', body)
    worksheet118.write('Q5', 'SOS', body)
    worksheet118.write('R5', 'JML', body)

    worksheet118.conditional_format(5, 0, row118_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet118.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF AKSES UI', title)
    worksheet118.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet118.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet118.write('A22', 'LOKASI', header)
    worksheet118.write('B22', 'TOTAL', header)
    worksheet118.merge_range('A21:B21', 'RANK', header)
    worksheet118.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet118.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet118.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet118.merge_range('F21:F22', 'KELAS', header)
    worksheet118.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet118.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet118.write('G22', 'MAT', body)
    worksheet118.write('H22', 'SEJ', body)
    worksheet118.write('I22', 'GEO', body)
    worksheet118.write('J22', 'EKO', body)
    worksheet118.write('K22', 'SOS', body)
    worksheet118.write('L22', 'JML', body)
    worksheet118.write('M22', 'MAT', body)
    worksheet118.write('N22', 'SEJ', body)
    worksheet118.write('O22', 'GEO', body)
    worksheet118.write('P22', 'EKO', body)
    worksheet118.write('Q22', 'SOS', body)
    worksheet118.write('R22', 'JML', body)

    worksheet118.conditional_format(22, 0, row118+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 119
    worksheet119.insert_image('A1', r'logo resmi nf.jpg')

    worksheet119.set_column('A:A', 7, center)
    worksheet119.set_column('B:B', 6, center)
    worksheet119.set_column('C:C', 18.14, center)
    worksheet119.set_column('D:D', 25, left)
    worksheet119.set_column('E:E', 13.14, left)
    worksheet119.set_column('F:F', 8.57, center)
    worksheet119.set_column('G:R', 5, center)
    worksheet119.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF JATIMEKAR', title)
    worksheet119.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet119.write('A5', 'LOKASI', header)
    worksheet119.write('B5', 'TOTAL', header)
    worksheet119.merge_range('A4:B4', 'RANK', header)
    worksheet119.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet119.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet119.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet119.merge_range('F4:F5', 'KELAS', header)
    worksheet119.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet119.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet119.write('G5', 'MAT', body)
    worksheet119.write('H5', 'SEJ', body)
    worksheet119.write('I5', 'GEO', body)
    worksheet119.write('J5', 'EKO', body)
    worksheet119.write('K5', 'SOS', body)
    worksheet119.write('L5', 'JML', body)
    worksheet119.write('M5', 'MAT', body)
    worksheet119.write('N5', 'SEJ', body)
    worksheet119.write('O5', 'GEO', body)
    worksheet119.write('P5', 'EKO', body)
    worksheet119.write('Q5', 'SOS', body)
    worksheet119.write('R5', 'JML', body)

    worksheet119.conditional_format(5, 0, row119_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet119.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF JATIMEKAR', title)
    worksheet119.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet119.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet119.write('A22', 'LOKASI', header)
    worksheet119.write('B22', 'TOTAL', header)
    worksheet119.merge_range('A21:B21', 'RANK', header)
    worksheet119.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet119.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet119.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet119.merge_range('F21:F22', 'KELAS', header)
    worksheet119.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet119.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet119.write('G22', 'MAT', body)
    worksheet119.write('H22', 'SEJ', body)
    worksheet119.write('I22', 'GEO', body)
    worksheet119.write('J22', 'EKO', body)
    worksheet119.write('K22', 'SOS', body)
    worksheet119.write('L22', 'JML', body)
    worksheet119.write('M22', 'MAT', body)
    worksheet119.write('N22', 'SEJ', body)
    worksheet119.write('O22', 'GEO', body)
    worksheet119.write('P22', 'EKO', body)
    worksheet119.write('Q22', 'SOS', body)
    worksheet119.write('R22', 'JML', body)

    worksheet119.conditional_format(22, 0, row119+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 120
    worksheet120.insert_image('A1', r'logo resmi nf.jpg')

    worksheet120.set_column('A:A', 7, center)
    worksheet120.set_column('B:B', 6, center)
    worksheet120.set_column('C:C', 18.14, center)
    worksheet120.set_column('D:D', 25, left)
    worksheet120.set_column('E:E', 13.14, left)
    worksheet120.set_column('F:F', 8.57, center)
    worksheet120.set_column('G:R', 5, center)
    worksheet120.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF RAWALUMBU', title)
    worksheet120.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet120.write('A5', 'LOKASI', header)
    worksheet120.write('B5', 'TOTAL', header)
    worksheet120.merge_range('A4:B4', 'RANK', header)
    worksheet120.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet120.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet120.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet120.merge_range('F4:F5', 'KELAS', header)
    worksheet120.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet120.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet120.write('G5', 'MAT', body)
    worksheet120.write('H5', 'SEJ', body)
    worksheet120.write('I5', 'GEO', body)
    worksheet120.write('J5', 'EKO', body)
    worksheet120.write('K5', 'SOS', body)
    worksheet120.write('L5', 'JML', body)
    worksheet120.write('M5', 'MAT', body)
    worksheet120.write('N5', 'SEJ', body)
    worksheet120.write('O5', 'GEO', body)
    worksheet120.write('P5', 'EKO', body)
    worksheet120.write('Q5', 'SOS', body)
    worksheet120.write('R5', 'JML', body)

    worksheet120.conditional_format(5, 0, row120_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet120.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF RAWALUMBU', title)
    worksheet120.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet120.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet120.write('A22', 'LOKASI', header)
    worksheet120.write('B22', 'TOTAL', header)
    worksheet120.merge_range('A21:B21', 'RANK', header)
    worksheet120.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet120.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet120.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet120.merge_range('F21:F22', 'KELAS', header)
    worksheet120.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet120.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet120.write('G22', 'MAT', body)
    worksheet120.write('H22', 'SEJ', body)
    worksheet120.write('I22', 'GEO', body)
    worksheet120.write('J22', 'EKO', body)
    worksheet120.write('K22', 'SOS', body)
    worksheet120.write('L22', 'JML', body)
    worksheet120.write('M22', 'MAT', body)
    worksheet120.write('N22', 'SEJ', body)
    worksheet120.write('O22', 'GEO', body)
    worksheet120.write('P22', 'EKO', body)
    worksheet120.write('Q22', 'SOS', body)
    worksheet120.write('R22', 'JML', body)

    worksheet120.conditional_format(22, 0, row120+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 121
    worksheet121.insert_image('A1', r'logo resmi nf.jpg')

    worksheet121.set_column('A:A', 7, center)
    worksheet121.set_column('B:B', 6, center)
    worksheet121.set_column('C:C', 18.14, center)
    worksheet121.set_column('D:D', 25, left)
    worksheet121.set_column('E:E', 13.14, left)
    worksheet121.set_column('F:F', 8.57, center)
    worksheet121.set_column('G:R', 5, center)
    worksheet121.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF TAMAN HARAPAN BARU', title)
    worksheet121.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet121.write('A5', 'LOKASI', header)
    worksheet121.write('B5', 'TOTAL', header)
    worksheet121.merge_range('A4:B4', 'RANK', header)
    worksheet121.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet121.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet121.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet121.merge_range('F4:F5', 'KELAS', header)
    worksheet121.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet121.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet121.write('G5', 'MAT', body)
    worksheet121.write('H5', 'SEJ', body)
    worksheet121.write('I5', 'GEO', body)
    worksheet121.write('J5', 'EKO', body)
    worksheet121.write('K5', 'SOS', body)
    worksheet121.write('L5', 'JML', body)
    worksheet121.write('M5', 'MAT', body)
    worksheet121.write('N5', 'SEJ', body)
    worksheet121.write('O5', 'GEO', body)
    worksheet121.write('P5', 'EKO', body)
    worksheet121.write('Q5', 'SOS', body)
    worksheet121.write('R5', 'JML', body)

    worksheet121.conditional_format(5, 0, row121_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet121.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF TAMAN HARAPAN BARU', title)
    worksheet121.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet121.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet121.write('A22', 'LOKASI', header)
    worksheet121.write('B22', 'TOTAL', header)
    worksheet121.merge_range('A21:B21', 'RANK', header)
    worksheet121.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet121.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet121.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet121.merge_range('F21:F22', 'KELAS', header)
    worksheet121.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet121.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet121.write('G22', 'MAT', body)
    worksheet121.write('H22', 'SEJ', body)
    worksheet121.write('I22', 'GEO', body)
    worksheet121.write('J22', 'EKO', body)
    worksheet121.write('K22', 'SOS', body)
    worksheet121.write('L22', 'JML', body)
    worksheet121.write('M22', 'MAT', body)
    worksheet121.write('N22', 'SEJ', body)
    worksheet121.write('O22', 'GEO', body)
    worksheet121.write('P22', 'EKO', body)
    worksheet121.write('Q22', 'SOS', body)
    worksheet121.write('R22', 'JML', body)

    worksheet121.conditional_format(22, 0, row121+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 122
    worksheet122.insert_image('A1', r'logo resmi nf.jpg')

    worksheet122.set_column('A:A', 7, center)
    worksheet122.set_column('B:B', 6, center)
    worksheet122.set_column('C:C', 18.14, center)
    worksheet122.set_column('D:D', 25, left)
    worksheet122.set_column('E:E', 13.14, left)
    worksheet122.set_column('F:F', 8.57, center)
    worksheet122.set_column('G:R', 5, center)
    worksheet122.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF VILA NUSA INDAH', title)
    worksheet122.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet122.write('A5', 'LOKASI', header)
    worksheet122.write('B5', 'TOTAL', header)
    worksheet122.merge_range('A4:B4', 'RANK', header)
    worksheet122.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet122.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet122.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet122.merge_range('F4:F5', 'KELAS', header)
    worksheet122.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet122.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet122.write('G5', 'MAT', body)
    worksheet122.write('H5', 'SEJ', body)
    worksheet122.write('I5', 'GEO', body)
    worksheet122.write('J5', 'EKO', body)
    worksheet122.write('K5', 'SOS', body)
    worksheet122.write('L5', 'JML', body)
    worksheet122.write('M5', 'MAT', body)
    worksheet122.write('N5', 'SEJ', body)
    worksheet122.write('O5', 'GEO', body)
    worksheet122.write('P5', 'EKO', body)
    worksheet122.write('Q5', 'SOS', body)
    worksheet122.write('R5', 'JML', body)

    worksheet122.conditional_format(5, 0, row122_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet122.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF VILA NUSA INDAH', title)
    worksheet122.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet122.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet122.write('A22', 'LOKASI', header)
    worksheet122.write('B22', 'TOTAL', header)
    worksheet122.merge_range('A21:B21', 'RANK', header)
    worksheet122.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet122.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet122.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet122.merge_range('F21:F22', 'KELAS', header)
    worksheet122.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet122.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet122.write('G22', 'MAT', body)
    worksheet122.write('H22', 'SEJ', body)
    worksheet122.write('I22', 'GEO', body)
    worksheet122.write('J22', 'EKO', body)
    worksheet122.write('K22', 'SOS', body)
    worksheet122.write('L22', 'JML', body)
    worksheet122.write('M22', 'MAT', body)
    worksheet122.write('N22', 'SEJ', body)
    worksheet122.write('O22', 'GEO', body)
    worksheet122.write('P22', 'EKO', body)
    worksheet122.write('Q22', 'SOS', body)
    worksheet122.write('R22', 'JML', body)

    worksheet122.conditional_format(22, 0, row122+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 123
    worksheet123.insert_image('A1', r'logo resmi nf.jpg')

    worksheet123.set_column('A:A', 7, center)
    worksheet123.set_column('B:B', 6, center)
    worksheet123.set_column('C:C', 18.14, center)
    worksheet123.set_column('D:D', 25, left)
    worksheet123.set_column('E:E', 13.14, left)
    worksheet123.set_column('F:F', 8.57, center)
    worksheet123.set_column('G:R', 5, center)
    worksheet123.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF JATIWARNA', title)
    worksheet123.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet123.write('A5', 'LOKASI', header)
    worksheet123.write('B5', 'TOTAL', header)
    worksheet123.merge_range('A4:B4', 'RANK', header)
    worksheet123.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet123.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet123.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet123.merge_range('F4:F5', 'KELAS', header)
    worksheet123.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet123.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet123.write('G5', 'MAT', body)
    worksheet123.write('H5', 'SEJ', body)
    worksheet123.write('I5', 'GEO', body)
    worksheet123.write('J5', 'EKO', body)
    worksheet123.write('K5', 'SOS', body)
    worksheet123.write('L5', 'JML', body)
    worksheet123.write('M5', 'MAT', body)
    worksheet123.write('N5', 'SEJ', body)
    worksheet123.write('O5', 'GEO', body)
    worksheet123.write('P5', 'EKO', body)
    worksheet123.write('Q5', 'SOS', body)
    worksheet123.write('R5', 'JML', body)

    worksheet123.conditional_format(5, 0, row123_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet123.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF JATIWARNA', title)
    worksheet123.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet123.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet123.write('A22', 'LOKASI', header)
    worksheet123.write('B22', 'TOTAL', header)
    worksheet123.merge_range('A21:B21', 'RANK', header)
    worksheet123.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet123.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet123.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet123.merge_range('F21:F22', 'KELAS', header)
    worksheet123.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet123.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet123.write('G22', 'MAT', body)
    worksheet123.write('H22', 'SEJ', body)
    worksheet123.write('I22', 'GEO', body)
    worksheet123.write('J22', 'EKO', body)
    worksheet123.write('K22', 'SOS', body)
    worksheet123.write('L22', 'JML', body)
    worksheet123.write('M22', 'MAT', body)
    worksheet123.write('N22', 'SEJ', body)
    worksheet123.write('O22', 'GEO', body)
    worksheet123.write('P22', 'EKO', body)
    worksheet123.write('Q22', 'SOS', body)
    worksheet123.write('R22', 'JML', body)

    worksheet123.conditional_format(22, 0, row123+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 124
    worksheet124.insert_image('A1', r'logo resmi nf.jpg')

    worksheet124.set_column('A:A', 7, center)
    worksheet124.set_column('B:B', 6, center)
    worksheet124.set_column('C:C', 18.14, center)
    worksheet124.set_column('D:D', 25, left)
    worksheet124.set_column('E:E', 13.14, left)
    worksheet124.set_column('F:F', 8.57, center)
    worksheet124.set_column('G:R', 5, center)
    worksheet124.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF TAMBUN', title)
    worksheet124.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet124.write('A5', 'LOKASI', header)
    worksheet124.write('B5', 'TOTAL', header)
    worksheet124.merge_range('A4:B4', 'RANK', header)
    worksheet124.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet124.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet124.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet124.merge_range('F4:F5', 'KELAS', header)
    worksheet124.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet124.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet124.write('G5', 'MAT', body)
    worksheet124.write('H5', 'SEJ', body)
    worksheet124.write('I5', 'GEO', body)
    worksheet124.write('J5', 'EKO', body)
    worksheet124.write('K5', 'SOS', body)
    worksheet124.write('L5', 'JML', body)
    worksheet124.write('M5', 'MAT', body)
    worksheet124.write('N5', 'SEJ', body)
    worksheet124.write('O5', 'GEO', body)
    worksheet124.write('P5', 'EKO', body)
    worksheet124.write('Q5', 'SOS', body)
    worksheet124.write('R5', 'JML', body)

    worksheet124.conditional_format(5, 0, row124_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet124.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF TAMBUN', title)
    worksheet124.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet124.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet124.write('A22', 'LOKASI', header)
    worksheet124.write('B22', 'TOTAL', header)
    worksheet124.merge_range('A21:B21', 'RANK', header)
    worksheet124.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet124.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet124.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet124.merge_range('F21:F22', 'KELAS', header)
    worksheet124.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet124.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet124.write('G22', 'MAT', body)
    worksheet124.write('H22', 'SEJ', body)
    worksheet124.write('I22', 'GEO', body)
    worksheet124.write('J22', 'EKO', body)
    worksheet124.write('K22', 'SOS', body)
    worksheet124.write('L22', 'JML', body)
    worksheet124.write('M22', 'MAT', body)
    worksheet124.write('N22', 'SEJ', body)
    worksheet124.write('O22', 'GEO', body)
    worksheet124.write('P22', 'EKO', body)
    worksheet124.write('Q22', 'SOS', body)
    worksheet124.write('R22', 'JML', body)

    worksheet124.conditional_format(22, 0, row124+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 125
    worksheet125.insert_image('A1', r'logo resmi nf.jpg')

    worksheet125.set_column('A:A', 7, center)
    worksheet125.set_column('B:B', 6, center)
    worksheet125.set_column('C:C', 18.14, center)
    worksheet125.set_column('D:D', 25, left)
    worksheet125.set_column('E:E', 13.14, left)
    worksheet125.set_column('F:F', 8.57, center)
    worksheet125.set_column('G:R', 5, center)
    worksheet125.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF DAAN MOGOT', title)
    worksheet125.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet125.write('A5', 'LOKASI', header)
    worksheet125.write('B5', 'TOTAL', header)
    worksheet125.merge_range('A4:B4', 'RANK', header)
    worksheet125.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet125.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet125.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet125.merge_range('F4:F5', 'KELAS', header)
    worksheet125.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet125.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet125.write('G5', 'MAT', body)
    worksheet125.write('H5', 'SEJ', body)
    worksheet125.write('I5', 'GEO', body)
    worksheet125.write('J5', 'EKO', body)
    worksheet125.write('K5', 'SOS', body)
    worksheet125.write('L5', 'JML', body)
    worksheet125.write('M5', 'MAT', body)
    worksheet125.write('N5', 'SEJ', body)
    worksheet125.write('O5', 'GEO', body)
    worksheet125.write('P5', 'EKO', body)
    worksheet125.write('Q5', 'SOS', body)
    worksheet125.write('R5', 'JML', body)

    worksheet125.conditional_format(5, 0, row125_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet125.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF DAAN MOGOT', title)
    worksheet125.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet125.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet125.write('A22', 'LOKASI', header)
    worksheet125.write('B22', 'TOTAL', header)
    worksheet125.merge_range('A21:B21', 'RANK', header)
    worksheet125.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet125.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet125.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet125.merge_range('F21:F22', 'KELAS', header)
    worksheet125.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet125.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet125.write('G22', 'MAT', body)
    worksheet125.write('H22', 'SEJ', body)
    worksheet125.write('I22', 'GEO', body)
    worksheet125.write('J22', 'EKO', body)
    worksheet125.write('K22', 'SOS', body)
    worksheet125.write('L22', 'JML', body)
    worksheet125.write('M22', 'MAT', body)
    worksheet125.write('N22', 'SEJ', body)
    worksheet125.write('O22', 'GEO', body)
    worksheet125.write('P22', 'EKO', body)
    worksheet125.write('Q22', 'SOS', body)
    worksheet125.write('R22', 'JML', body)

    worksheet125.conditional_format(22, 0, row125+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 126
    worksheet126.insert_image('A1', r'logo resmi nf.jpg')

    worksheet126.set_column('A:A', 7, center)
    worksheet126.set_column('B:B', 6, center)
    worksheet126.set_column('C:C', 18.14, center)
    worksheet126.set_column('D:D', 25, left)
    worksheet126.set_column('E:E', 13.14, left)
    worksheet126.set_column('F:F', 8.57, center)
    worksheet126.set_column('G:R', 5, center)
    worksheet126.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIBUBUR', title)
    worksheet126.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet126.write('A5', 'LOKASI', header)
    worksheet126.write('B5', 'TOTAL', header)
    worksheet126.merge_range('A4:B4', 'RANK', header)
    worksheet126.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet126.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet126.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet126.merge_range('F4:F5', 'KELAS', header)
    worksheet126.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet126.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet126.write('G5', 'MAT', body)
    worksheet126.write('H5', 'SEJ', body)
    worksheet126.write('I5', 'GEO', body)
    worksheet126.write('J5', 'EKO', body)
    worksheet126.write('K5', 'SOS', body)
    worksheet126.write('L5', 'JML', body)
    worksheet126.write('M5', 'MAT', body)
    worksheet126.write('N5', 'SEJ', body)
    worksheet126.write('O5', 'GEO', body)
    worksheet126.write('P5', 'EKO', body)
    worksheet126.write('Q5', 'SOS', body)
    worksheet126.write('R5', 'JML', body)

    worksheet126.conditional_format(5, 0, row126_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet126.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CIBUBUR', title)
    worksheet126.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet126.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet126.write('A22', 'LOKASI', header)
    worksheet126.write('B22', 'TOTAL', header)
    worksheet126.merge_range('A21:B21', 'RANK', header)
    worksheet126.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet126.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet126.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet126.merge_range('F21:F22', 'KELAS', header)
    worksheet126.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet126.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet126.write('G22', 'MAT', body)
    worksheet126.write('H22', 'SEJ', body)
    worksheet126.write('I22', 'GEO', body)
    worksheet126.write('J22', 'EKO', body)
    worksheet126.write('K22', 'SOS', body)
    worksheet126.write('L22', 'JML', body)
    worksheet126.write('M22', 'MAT', body)
    worksheet126.write('N22', 'SEJ', body)
    worksheet126.write('O22', 'GEO', body)
    worksheet126.write('P22', 'EKO', body)
    worksheet126.write('Q22', 'SOS', body)
    worksheet126.write('R22', 'JML', body)

    worksheet126.conditional_format(22, 0, row126+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 127
    worksheet127.insert_image('A1', r'logo resmi nf.jpg')

    worksheet127.set_column('A:A', 7, center)
    worksheet127.set_column('B:B', 6, center)
    worksheet127.set_column('C:C', 18.14, center)
    worksheet127.set_column('D:D', 25, left)
    worksheet127.set_column('E:E', 13.14, left)
    worksheet127.set_column('F:F', 8.57, center)
    worksheet127.set_column('G:R', 5, center)
    worksheet127.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CENGKARENG', title)
    worksheet127.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet127.write('A5', 'LOKASI', header)
    worksheet127.write('B5', 'TOTAL', header)
    worksheet127.merge_range('A4:B4', 'RANK', header)
    worksheet127.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet127.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet127.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet127.merge_range('F4:F5', 'KELAS', header)
    worksheet127.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet127.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet127.write('G5', 'MAT', body)
    worksheet127.write('H5', 'SEJ', body)
    worksheet127.write('I5', 'GEO', body)
    worksheet127.write('J5', 'EKO', body)
    worksheet127.write('K5', 'SOS', body)
    worksheet127.write('L5', 'JML', body)
    worksheet127.write('M5', 'MAT', body)
    worksheet127.write('N5', 'SEJ', body)
    worksheet127.write('O5', 'GEO', body)
    worksheet127.write('P5', 'EKO', body)
    worksheet127.write('Q5', 'SOS', body)
    worksheet127.write('R5', 'JML', body)

    worksheet127.conditional_format(5, 0, row127_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet127.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CENGKARENG', title)
    worksheet127.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet127.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet127.write('A22', 'LOKASI', header)
    worksheet127.write('B22', 'TOTAL', header)
    worksheet127.merge_range('A21:B21', 'RANK', header)
    worksheet127.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet127.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet127.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet127.merge_range('F21:F22', 'KELAS', header)
    worksheet127.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet127.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet127.write('G22', 'MAT', body)
    worksheet127.write('H22', 'SEJ', body)
    worksheet127.write('I22', 'GEO', body)
    worksheet127.write('J22', 'EKO', body)
    worksheet127.write('K22', 'SOS', body)
    worksheet127.write('L22', 'JML', body)
    worksheet127.write('M22', 'MAT', body)
    worksheet127.write('N22', 'SEJ', body)
    worksheet127.write('O22', 'GEO', body)
    worksheet127.write('P22', 'EKO', body)
    worksheet127.write('Q22', 'SOS', body)
    worksheet127.write('R22', 'JML', body)

    worksheet127.conditional_format(22, 0, row127+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 128
    worksheet128.insert_image('A1', r'logo resmi nf.jpg')

    worksheet128.set_column('A:A', 7, center)
    worksheet128.set_column('B:B', 6, center)
    worksheet128.set_column('C:C', 18.14, center)
    worksheet128.set_column('D:D', 25, left)
    worksheet128.set_column('E:E', 13.14, left)
    worksheet128.set_column('F:F', 8.57, center)
    worksheet128.set_column('G:R', 5, center)
    worksheet128.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PETUKANGAN', title)
    worksheet128.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet128.write('A5', 'LOKASI', header)
    worksheet128.write('B5', 'TOTAL', header)
    worksheet128.merge_range('A4:B4', 'RANK', header)
    worksheet128.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet128.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet128.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet128.merge_range('F4:F5', 'KELAS', header)
    worksheet128.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet128.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet128.write('G5', 'MAT', body)
    worksheet128.write('H5', 'SEJ', body)
    worksheet128.write('I5', 'GEO', body)
    worksheet128.write('J5', 'EKO', body)
    worksheet128.write('K5', 'SOS', body)
    worksheet128.write('L5', 'JML', body)
    worksheet128.write('M5', 'MAT', body)
    worksheet128.write('N5', 'SEJ', body)
    worksheet128.write('O5', 'GEO', body)
    worksheet128.write('P5', 'EKO', body)
    worksheet128.write('Q5', 'SOS', body)
    worksheet128.write('R5', 'JML', body)

    worksheet128.conditional_format(5, 0, row128_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet128.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PETUKANGAN', title)
    worksheet128.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet128.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet128.write('A22', 'LOKASI', header)
    worksheet128.write('B22', 'TOTAL', header)
    worksheet128.merge_range('A21:B21', 'RANK', header)
    worksheet128.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet128.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet128.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet128.merge_range('F21:F22', 'KELAS', header)
    worksheet128.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet128.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet128.write('G22', 'MAT', body)
    worksheet128.write('H22', 'SEJ', body)
    worksheet128.write('I22', 'GEO', body)
    worksheet128.write('J22', 'EKO', body)
    worksheet128.write('K22', 'SOS', body)
    worksheet128.write('L22', 'JML', body)
    worksheet128.write('M22', 'MAT', body)
    worksheet128.write('N22', 'SEJ', body)
    worksheet128.write('O22', 'GEO', body)
    worksheet128.write('P22', 'EKO', body)
    worksheet128.write('Q22', 'SOS', body)
    worksheet128.write('R22', 'JML', body)

    worksheet128.conditional_format(22, 0, row128+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 129
    worksheet129.insert_image('A1', r'logo resmi nf.jpg')

    worksheet129.set_column('A:A', 7, center)
    worksheet129.set_column('B:B', 6, center)
    worksheet129.set_column('C:C', 18.14, center)
    worksheet129.set_column('D:D', 25, left)
    worksheet129.set_column('E:E', 13.14, left)
    worksheet129.set_column('F:F', 8.57, center)
    worksheet129.set_column('G:R', 5, center)
    worksheet129.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MERUYA UTARA', title)
    worksheet129.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet129.write('A5', 'LOKASI', header)
    worksheet129.write('B5', 'TOTAL', header)
    worksheet129.merge_range('A4:B4', 'RANK', header)
    worksheet129.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet129.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet129.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet129.merge_range('F4:F5', 'KELAS', header)
    worksheet129.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet129.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet129.write('G5', 'MAT', body)
    worksheet129.write('H5', 'SEJ', body)
    worksheet129.write('I5', 'GEO', body)
    worksheet129.write('J5', 'EKO', body)
    worksheet129.write('K5', 'SOS', body)
    worksheet129.write('L5', 'JML', body)
    worksheet129.write('M5', 'MAT', body)
    worksheet129.write('N5', 'SEJ', body)
    worksheet129.write('O5', 'GEO', body)
    worksheet129.write('P5', 'EKO', body)
    worksheet129.write('Q5', 'SOS', body)
    worksheet129.write('R5', 'JML', body)

    worksheet129.conditional_format(5, 0, row129_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet129.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF MERUYA UTARA', title)
    worksheet129.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet129.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet129.write('A22', 'LOKASI', header)
    worksheet129.write('B22', 'TOTAL', header)
    worksheet129.merge_range('A21:B21', 'RANK', header)
    worksheet129.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet129.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet129.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet129.merge_range('F21:F22', 'KELAS', header)
    worksheet129.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet129.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet129.write('G22', 'MAT', body)
    worksheet129.write('H22', 'SEJ', body)
    worksheet129.write('I22', 'GEO', body)
    worksheet129.write('J22', 'EKO', body)
    worksheet129.write('K22', 'SOS', body)
    worksheet129.write('L22', 'JML', body)
    worksheet129.write('M22', 'MAT', body)
    worksheet129.write('N22', 'SEJ', body)
    worksheet129.write('O22', 'GEO', body)
    worksheet129.write('P22', 'EKO', body)
    worksheet129.write('Q22', 'SOS', body)
    worksheet129.write('R22', 'JML', body)

    worksheet129.conditional_format(22, 0, row129+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 130
    worksheet130.insert_image('A1', r'logo resmi nf.jpg')

    worksheet130.set_column('A:A', 7, center)
    worksheet130.set_column('B:B', 6, center)
    worksheet130.set_column('C:C', 18.14, center)
    worksheet130.set_column('D:D', 25, left)
    worksheet130.set_column('E:E', 13.14, left)
    worksheet130.set_column('F:F', 8.57, center)
    worksheet130.set_column('G:R', 5, center)
    worksheet130.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BINTARA', title)
    worksheet130.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet130.write('A5', 'LOKASI', header)
    worksheet130.write('B5', 'TOTAL', header)
    worksheet130.merge_range('A4:B4', 'RANK', header)
    worksheet130.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet130.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet130.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet130.merge_range('F4:F5', 'KELAS', header)
    worksheet130.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet130.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet130.write('G5', 'MAT', body)
    worksheet130.write('H5', 'SEJ', body)
    worksheet130.write('I5', 'GEO', body)
    worksheet130.write('J5', 'EKO', body)
    worksheet130.write('K5', 'SOS', body)
    worksheet130.write('L5', 'JML', body)
    worksheet130.write('M5', 'MAT', body)
    worksheet130.write('N5', 'SEJ', body)
    worksheet130.write('O5', 'GEO', body)
    worksheet130.write('P5', 'EKO', body)
    worksheet130.write('Q5', 'SOS', body)
    worksheet130.write('R5', 'JML', body)

    worksheet130.conditional_format(5, 0, row130_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet130.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF BINTARA', title)
    worksheet130.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet130.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet130.write('A22', 'LOKASI', header)
    worksheet130.write('B22', 'TOTAL', header)
    worksheet130.merge_range('A21:B21', 'RANK', header)
    worksheet130.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet130.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet130.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet130.merge_range('F21:F22', 'KELAS', header)
    worksheet130.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet130.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet130.write('G22', 'MAT', body)
    worksheet130.write('H22', 'SEJ', body)
    worksheet130.write('I22', 'GEO', body)
    worksheet130.write('J22', 'EKO', body)
    worksheet130.write('K22', 'SOS', body)
    worksheet130.write('L22', 'JML', body)
    worksheet130.write('M22', 'MAT', body)
    worksheet130.write('N22', 'SEJ', body)
    worksheet130.write('O22', 'GEO', body)
    worksheet130.write('P22', 'EKO', body)
    worksheet130.write('Q22', 'SOS', body)
    worksheet130.write('R22', 'JML', body)

    worksheet130.conditional_format(22, 0, row130+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 131
    worksheet131.insert_image('A1', r'logo resmi nf.jpg')

    worksheet131.set_column('A:A', 7, center)
    worksheet131.set_column('B:B', 6, center)
    worksheet131.set_column('C:C', 18.14, center)
    worksheet131.set_column('D:D', 25, left)
    worksheet131.set_column('E:E', 13.14, left)
    worksheet131.set_column('F:F', 8.57, center)
    worksheet131.set_column('G:R', 5, center)
    worksheet131.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MALANG', title)
    worksheet131.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet131.write('A5', 'LOKASI', header)
    worksheet131.write('B5', 'TOTAL', header)
    worksheet131.merge_range('A4:B4', 'RANK', header)
    worksheet131.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet131.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet131.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet131.merge_range('F4:F5', 'KELAS', header)
    worksheet131.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet131.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet131.write('G5', 'MAT', body)
    worksheet131.write('H5', 'SEJ', body)
    worksheet131.write('I5', 'GEO', body)
    worksheet131.write('J5', 'EKO', body)
    worksheet131.write('K5', 'SOS', body)
    worksheet131.write('L5', 'JML', body)
    worksheet131.write('M5', 'MAT', body)
    worksheet131.write('N5', 'SEJ', body)
    worksheet131.write('O5', 'GEO', body)
    worksheet131.write('P5', 'EKO', body)
    worksheet131.write('Q5', 'SOS', body)
    worksheet131.write('R5', 'JML', body)

    worksheet131.conditional_format(5, 0, row131_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet131.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF MALANG', title)
    worksheet131.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet131.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet131.write('A22', 'LOKASI', header)
    worksheet131.write('B22', 'TOTAL', header)
    worksheet131.merge_range('A21:B21', 'RANK', header)
    worksheet131.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet131.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet131.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet131.merge_range('F21:F22', 'KELAS', header)
    worksheet131.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet131.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet131.write('G22', 'MAT', body)
    worksheet131.write('H22', 'SEJ', body)
    worksheet131.write('I22', 'GEO', body)
    worksheet131.write('J22', 'EKO', body)
    worksheet131.write('K22', 'SOS', body)
    worksheet131.write('L22', 'JML', body)
    worksheet131.write('M22', 'MAT', body)
    worksheet131.write('N22', 'SEJ', body)
    worksheet131.write('O22', 'GEO', body)
    worksheet131.write('P22', 'EKO', body)
    worksheet131.write('Q22', 'SOS', body)
    worksheet131.write('R22', 'JML', body)

    worksheet131.conditional_format(22, 0, row131+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 132
    worksheet132.insert_image('A1', r'logo resmi nf.jpg')

    worksheet132.set_column('A:A', 7, center)
    worksheet132.set_column('B:B', 6, center)
    worksheet132.set_column('C:C', 18.14, center)
    worksheet132.set_column('D:D', 25, left)
    worksheet132.set_column('E:E', 13.14, left)
    worksheet132.set_column('F:F', 8.57, center)
    worksheet132.set_column('G:R', 5, center)
    worksheet132.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MEDAN BARU', title)
    worksheet132.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet132.write('A5', 'LOKASI', header)
    worksheet132.write('B5', 'TOTAL', header)
    worksheet132.merge_range('A4:B4', 'RANK', header)
    worksheet132.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet132.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet132.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet132.merge_range('F4:F5', 'KELAS', header)
    worksheet132.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet132.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet132.write('G5', 'MAT', body)
    worksheet132.write('H5', 'SEJ', body)
    worksheet132.write('I5', 'GEO', body)
    worksheet132.write('J5', 'EKO', body)
    worksheet132.write('K5', 'SOS', body)
    worksheet132.write('L5', 'JML', body)
    worksheet132.write('M5', 'MAT', body)
    worksheet132.write('N5', 'SEJ', body)
    worksheet132.write('O5', 'GEO', body)
    worksheet132.write('P5', 'EKO', body)
    worksheet132.write('Q5', 'SOS', body)
    worksheet132.write('R5', 'JML', body)

    worksheet132.conditional_format(5, 0, row132_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet132.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF MEDAN BARU', title)
    worksheet132.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet132.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet132.write('A22', 'LOKASI', header)
    worksheet132.write('B22', 'TOTAL', header)
    worksheet132.merge_range('A21:B21', 'RANK', header)
    worksheet132.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet132.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet132.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet132.merge_range('F21:F22', 'KELAS', header)
    worksheet132.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet132.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet132.write('G22', 'MAT', body)
    worksheet132.write('H22', 'SEJ', body)
    worksheet132.write('I22', 'GEO', body)
    worksheet132.write('J22', 'EKO', body)
    worksheet132.write('K22', 'SOS', body)
    worksheet132.write('L22', 'JML', body)
    worksheet132.write('M22', 'MAT', body)
    worksheet132.write('N22', 'SEJ', body)
    worksheet132.write('O22', 'GEO', body)
    worksheet132.write('P22', 'EKO', body)
    worksheet132.write('Q22', 'SOS', body)
    worksheet132.write('R22', 'JML', body)

    worksheet132.conditional_format(22, 0, row132+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 133
    worksheet133.insert_image('A1', r'logo resmi nf.jpg')

    worksheet133.set_column('A:A', 7, center)
    worksheet133.set_column('B:B', 6, center)
    worksheet133.set_column('C:C', 18.14, center)
    worksheet133.set_column('D:D', 25, left)
    worksheet133.set_column('E:E', 13.14, left)
    worksheet133.set_column('F:F', 8.57, center)
    worksheet133.set_column('G:R', 5, center)
    worksheet133.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MEDAN HELVETIA', title)
    worksheet133.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet133.write('A5', 'LOKASI', header)
    worksheet133.write('B5', 'TOTAL', header)
    worksheet133.merge_range('A4:B4', 'RANK', header)
    worksheet133.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet133.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet133.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet133.merge_range('F4:F5', 'KELAS', header)
    worksheet133.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet133.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet133.write('G5', 'MAT', body)
    worksheet133.write('H5', 'SEJ', body)
    worksheet133.write('I5', 'GEO', body)
    worksheet133.write('J5', 'EKO', body)
    worksheet133.write('K5', 'SOS', body)
    worksheet133.write('L5', 'JML', body)
    worksheet133.write('M5', 'MAT', body)
    worksheet133.write('N5', 'SEJ', body)
    worksheet133.write('O5', 'GEO', body)
    worksheet133.write('P5', 'EKO', body)
    worksheet133.write('Q5', 'SOS', body)
    worksheet133.write('R5', 'JML', body)

    worksheet133.conditional_format(5, 0, row133_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet133.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF MEDAN HELVETIA', title)
    worksheet133.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet133.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet133.write('A22', 'LOKASI', header)
    worksheet133.write('B22', 'TOTAL', header)
    worksheet133.merge_range('A21:B21', 'RANK', header)
    worksheet133.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet133.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet133.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet133.merge_range('F21:F22', 'KELAS', header)
    worksheet133.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet133.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet133.write('G22', 'MAT', body)
    worksheet133.write('H22', 'SEJ', body)
    worksheet133.write('I22', 'GEO', body)
    worksheet133.write('J22', 'EKO', body)
    worksheet133.write('K22', 'SOS', body)
    worksheet133.write('L22', 'JML', body)
    worksheet133.write('M22', 'MAT', body)
    worksheet133.write('N22', 'SEJ', body)
    worksheet133.write('O22', 'GEO', body)
    worksheet133.write('P22', 'EKO', body)
    worksheet133.write('Q22', 'SOS', body)
    worksheet133.write('R22', 'JML', body)

    worksheet133.conditional_format(22, 0, row133+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 134
    worksheet134.insert_image('A1', r'logo resmi nf.jpg')

    worksheet134.set_column('A:A', 7, center)
    worksheet134.set_column('B:B', 6, center)
    worksheet134.set_column('C:C', 18.14, center)
    worksheet134.set_column('D:D', 25, left)
    worksheet134.set_column('E:E', 13.14, left)
    worksheet134.set_column('F:F', 8.57, center)
    worksheet134.set_column('G:R', 5, center)
    worksheet134.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIHANJUANG', title)
    worksheet134.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet134.write('A5', 'LOKASI', header)
    worksheet134.write('B5', 'TOTAL', header)
    worksheet134.merge_range('A4:B4', 'RANK', header)
    worksheet134.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet134.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet134.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet134.merge_range('F4:F5', 'KELAS', header)
    worksheet134.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet134.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet134.write('G5', 'MAT', body)
    worksheet134.write('H5', 'SEJ', body)
    worksheet134.write('I5', 'GEO', body)
    worksheet134.write('J5', 'EKO', body)
    worksheet134.write('K5', 'SOS', body)
    worksheet134.write('L5', 'JML', body)
    worksheet134.write('M5', 'MAT', body)
    worksheet134.write('N5', 'SEJ', body)
    worksheet134.write('O5', 'GEO', body)
    worksheet134.write('P5', 'EKO', body)
    worksheet134.write('Q5', 'SOS', body)
    worksheet134.write('R5', 'JML', body)

    worksheet134.conditional_format(5, 0, row134_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet134.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CIHANJUANG', title)
    worksheet134.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet134.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet134.write('A22', 'LOKASI', header)
    worksheet134.write('B22', 'TOTAL', header)
    worksheet134.merge_range('A21:B21', 'RANK', header)
    worksheet134.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet134.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet134.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet134.merge_range('F21:F22', 'KELAS', header)
    worksheet134.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet134.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet134.write('G22', 'MAT', body)
    worksheet134.write('H22', 'SEJ', body)
    worksheet134.write('I22', 'GEO', body)
    worksheet134.write('J22', 'EKO', body)
    worksheet134.write('K22', 'SOS', body)
    worksheet134.write('L22', 'JML', body)
    worksheet134.write('M22', 'MAT', body)
    worksheet134.write('N22', 'SEJ', body)
    worksheet134.write('O22', 'GEO', body)
    worksheet134.write('P22', 'EKO', body)
    worksheet134.write('Q22', 'SOS', body)
    worksheet134.write('R22', 'JML', body)

    worksheet134.conditional_format(22, 0, row134+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 135
    worksheet135.insert_image('A1', r'logo resmi nf.jpg')

    worksheet135.set_column('A:A', 7, center)
    worksheet135.set_column('B:B', 6, center)
    worksheet135.set_column('C:C', 18.14, center)
    worksheet135.set_column('D:D', 25, left)
    worksheet135.set_column('E:E', 13.14, left)
    worksheet135.set_column('F:F', 8.57, center)
    worksheet135.set_column('G:R', 5, center)
    worksheet135.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BUAH BATU', title)
    worksheet135.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet135.write('A5', 'LOKASI', header)
    worksheet135.write('B5', 'TOTAL', header)
    worksheet135.merge_range('A4:B4', 'RANK', header)
    worksheet135.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet135.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet135.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet135.merge_range('F4:F5', 'KELAS', header)
    worksheet135.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet135.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet135.write('G5', 'MAT', body)
    worksheet135.write('H5', 'SEJ', body)
    worksheet135.write('I5', 'GEO', body)
    worksheet135.write('J5', 'EKO', body)
    worksheet135.write('K5', 'SOS', body)
    worksheet135.write('L5', 'JML', body)
    worksheet135.write('M5', 'MAT', body)
    worksheet135.write('N5', 'SEJ', body)
    worksheet135.write('O5', 'GEO', body)
    worksheet135.write('P5', 'EKO', body)
    worksheet135.write('Q5', 'SOS', body)
    worksheet135.write('R5', 'JML', body)

    worksheet135.conditional_format(5, 0, row135_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet135.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF BUAH BATU', title)
    worksheet135.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet135.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet135.write('A22', 'LOKASI', header)
    worksheet135.write('B22', 'TOTAL', header)
    worksheet135.merge_range('A21:B21', 'RANK', header)
    worksheet135.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet135.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet135.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet135.merge_range('F21:F22', 'KELAS', header)
    worksheet135.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet135.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet135.write('G22', 'MAT', body)
    worksheet135.write('H22', 'SEJ', body)
    worksheet135.write('I22', 'GEO', body)
    worksheet135.write('J22', 'EKO', body)
    worksheet135.write('K22', 'SOS', body)
    worksheet135.write('L22', 'JML', body)
    worksheet135.write('M22', 'MAT', body)
    worksheet135.write('N22', 'SEJ', body)
    worksheet135.write('O22', 'GEO', body)
    worksheet135.write('P22', 'EKO', body)
    worksheet135.write('Q22', 'SOS', body)
    worksheet135.write('R22', 'JML', body)

    worksheet135.conditional_format(22, 0, row135+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 136
    worksheet136.insert_image('A1', r'logo resmi nf.jpg')

    worksheet136.set_column('A:A', 7, center)
    worksheet136.set_column('B:B', 6, center)
    worksheet136.set_column('C:C', 18.14, center)
    worksheet136.set_column('D:D', 25, left)
    worksheet136.set_column('E:E', 13.14, left)
    worksheet136.set_column('F:F', 8.57, center)
    worksheet136.set_column('G:R', 5, center)
    worksheet136.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SUMBAWA', title)
    worksheet136.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet136.write('A5', 'LOKASI', header)
    worksheet136.write('B5', 'TOTAL', header)
    worksheet136.merge_range('A4:B4', 'RANK', header)
    worksheet136.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet136.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet136.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet136.merge_range('F4:F5', 'KELAS', header)
    worksheet136.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet136.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet136.write('G5', 'MAT', body)
    worksheet136.write('H5', 'SEJ', body)
    worksheet136.write('I5', 'GEO', body)
    worksheet136.write('J5', 'EKO', body)
    worksheet136.write('K5', 'SOS', body)
    worksheet136.write('L5', 'JML', body)
    worksheet136.write('M5', 'MAT', body)
    worksheet136.write('N5', 'SEJ', body)
    worksheet136.write('O5', 'GEO', body)
    worksheet136.write('P5', 'EKO', body)
    worksheet136.write('Q5', 'SOS', body)
    worksheet136.write('R5', 'JML', body)

    worksheet136.conditional_format(5, 0, row136_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet136.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF SUMBAWA', title)
    worksheet136.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet136.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet136.write('A22', 'LOKASI', header)
    worksheet136.write('B22', 'TOTAL', header)
    worksheet136.merge_range('A21:B21', 'RANK', header)
    worksheet136.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet136.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet136.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet136.merge_range('F21:F22', 'KELAS', header)
    worksheet136.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet136.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet136.write('G22', 'MAT', body)
    worksheet136.write('H22', 'SEJ', body)
    worksheet136.write('I22', 'GEO', body)
    worksheet136.write('J22', 'EKO', body)
    worksheet136.write('K22', 'SOS', body)
    worksheet136.write('L22', 'JML', body)
    worksheet136.write('M22', 'MAT', body)
    worksheet136.write('N22', 'SEJ', body)
    worksheet136.write('O22', 'GEO', body)
    worksheet136.write('P22', 'EKO', body)
    worksheet136.write('Q22', 'SOS', body)
    worksheet136.write('R22', 'JML', body)

    worksheet136.conditional_format(22, 0, row136+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 137
    worksheet137.insert_image('A1', r'logo resmi nf.jpg')

    worksheet137.set_column('A:A', 7, center)
    worksheet137.set_column('B:B', 6, center)
    worksheet137.set_column('C:C', 18.14, center)
    worksheet137.set_column('D:D', 25, left)
    worksheet137.set_column('E:E', 13.14, left)
    worksheet137.set_column('F:F', 8.57, center)
    worksheet137.set_column('G:R', 5, center)
    worksheet137.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF UJUNG BERUNG', title)
    worksheet137.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet137.write('A5', 'LOKASI', header)
    worksheet137.write('B5', 'TOTAL', header)
    worksheet137.merge_range('A4:B4', 'RANK', header)
    worksheet137.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet137.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet137.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet137.merge_range('F4:F5', 'KELAS', header)
    worksheet137.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet137.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet137.write('G5', 'MAT', body)
    worksheet137.write('H5', 'SEJ', body)
    worksheet137.write('I5', 'GEO', body)
    worksheet137.write('J5', 'EKO', body)
    worksheet137.write('K5', 'SOS', body)
    worksheet137.write('L5', 'JML', body)
    worksheet137.write('M5', 'MAT', body)
    worksheet137.write('N5', 'SEJ', body)
    worksheet137.write('O5', 'GEO', body)
    worksheet137.write('P5', 'EKO', body)
    worksheet137.write('Q5', 'SOS', body)
    worksheet137.write('R5', 'JML', body)

    worksheet137.conditional_format(5, 0, row137_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet137.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF UJUNG BERUNG', title)
    worksheet137.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet137.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet137.write('A22', 'LOKASI', header)
    worksheet137.write('B22', 'TOTAL', header)
    worksheet137.merge_range('A21:B21', 'RANK', header)
    worksheet137.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet137.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet137.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet137.merge_range('F21:F22', 'KELAS', header)
    worksheet137.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet137.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet137.write('G22', 'MAT', body)
    worksheet137.write('H22', 'SEJ', body)
    worksheet137.write('I22', 'GEO', body)
    worksheet137.write('J22', 'EKO', body)
    worksheet137.write('K22', 'SOS', body)
    worksheet137.write('L22', 'JML', body)
    worksheet137.write('M22', 'MAT', body)
    worksheet137.write('N22', 'SEJ', body)
    worksheet137.write('O22', 'GEO', body)
    worksheet137.write('P22', 'EKO', body)
    worksheet137.write('Q22', 'SOS', body)
    worksheet137.write('R22', 'JML', body)

    worksheet137.conditional_format(22, 0, row137+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 138
    worksheet138.insert_image('A1', r'logo resmi nf.jpg')

    worksheet138.set_column('A:A', 7, center)
    worksheet138.set_column('B:B', 6, center)
    worksheet138.set_column('C:C', 18.14, center)
    worksheet138.set_column('D:D', 25, left)
    worksheet138.set_column('E:E', 13.14, left)
    worksheet138.set_column('F:F', 8.57, center)
    worksheet138.set_column('G:R', 5, center)
    worksheet138.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SANGKURIANG', title)
    worksheet138.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet138.write('A5', 'LOKASI', header)
    worksheet138.write('B5', 'TOTAL', header)
    worksheet138.merge_range('A4:B4', 'RANK', header)
    worksheet138.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet138.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet138.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet138.merge_range('F4:F5', 'KELAS', header)
    worksheet138.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet138.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet138.write('G5', 'MAT', body)
    worksheet138.write('H5', 'SEJ', body)
    worksheet138.write('I5', 'GEO', body)
    worksheet138.write('J5', 'EKO', body)
    worksheet138.write('K5', 'SOS', body)
    worksheet138.write('L5', 'JML', body)
    worksheet138.write('M5', 'MAT', body)
    worksheet138.write('N5', 'SEJ', body)
    worksheet138.write('O5', 'GEO', body)
    worksheet138.write('P5', 'EKO', body)
    worksheet138.write('Q5', 'SOS', body)
    worksheet138.write('R5', 'JML', body)

    worksheet138.conditional_format(5, 0, row138_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet138.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF SANGKURIANG', title)
    worksheet138.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet138.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet138.write('A22', 'LOKASI', header)
    worksheet138.write('B22', 'TOTAL', header)
    worksheet138.merge_range('A21:B21', 'RANK', header)
    worksheet138.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet138.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet138.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet138.merge_range('F21:F22', 'KELAS', header)
    worksheet138.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet138.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet138.write('G22', 'MAT', body)
    worksheet138.write('H22', 'SEJ', body)
    worksheet138.write('I22', 'GEO', body)
    worksheet138.write('J22', 'EKO', body)
    worksheet138.write('K22', 'SOS', body)
    worksheet138.write('L22', 'JML', body)
    worksheet138.write('M22', 'MAT', body)
    worksheet138.write('N22', 'SEJ', body)
    worksheet138.write('O22', 'GEO', body)
    worksheet138.write('P22', 'EKO', body)
    worksheet138.write('Q22', 'SOS', body)
    worksheet138.write('R22', 'JML', body)

    worksheet138.conditional_format(22, 0, row138+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 139
    worksheet139.insert_image('A1', r'logo resmi nf.jpg')

    worksheet139.set_column('A:A', 7, center)
    worksheet139.set_column('B:B', 6, center)
    worksheet139.set_column('C:C', 18.14, center)
    worksheet139.set_column('D:D', 25, left)
    worksheet139.set_column('E:E', 13.14, left)
    worksheet139.set_column('F:F', 8.57, center)
    worksheet139.set_column('G:R', 5, center)
    worksheet139.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SARIJADI', title)
    worksheet139.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet139.write('A5', 'LOKASI', header)
    worksheet139.write('B5', 'TOTAL', header)
    worksheet139.merge_range('A4:B4', 'RANK', header)
    worksheet139.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet139.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet139.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet139.merge_range('F4:F5', 'KELAS', header)
    worksheet139.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet139.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet139.write('G5', 'MAT', body)
    worksheet139.write('H5', 'SEJ', body)
    worksheet139.write('I5', 'GEO', body)
    worksheet139.write('J5', 'EKO', body)
    worksheet139.write('K5', 'SOS', body)
    worksheet139.write('L5', 'JML', body)
    worksheet139.write('M5', 'MAT', body)
    worksheet139.write('N5', 'SEJ', body)
    worksheet139.write('O5', 'GEO', body)
    worksheet139.write('P5', 'EKO', body)
    worksheet139.write('Q5', 'SOS', body)
    worksheet139.write('R5', 'JML', body)

    worksheet139.conditional_format(5, 0, row139_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet139.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF SARIJADI', title)
    worksheet139.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet139.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet139.write('A22', 'LOKASI', header)
    worksheet139.write('B22', 'TOTAL', header)
    worksheet139.merge_range('A21:B21', 'RANK', header)
    worksheet139.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet139.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet139.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet139.merge_range('F21:F22', 'KELAS', header)
    worksheet139.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet139.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet139.write('G22', 'MAT', body)
    worksheet139.write('H22', 'SEJ', body)
    worksheet139.write('I22', 'GEO', body)
    worksheet139.write('J22', 'EKO', body)
    worksheet139.write('K22', 'SOS', body)
    worksheet139.write('L22', 'JML', body)
    worksheet139.write('M22', 'MAT', body)
    worksheet139.write('N22', 'SEJ', body)
    worksheet139.write('O22', 'GEO', body)
    worksheet139.write('P22', 'EKO', body)
    worksheet139.write('Q22', 'SOS', body)
    worksheet139.write('R22', 'JML', body)

    worksheet139.conditional_format(22, 0, row139+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 140
    worksheet140.insert_image('A1', r'logo resmi nf.jpg')

    worksheet140.set_column('A:A', 7, center)
    worksheet140.set_column('B:B', 6, center)
    worksheet140.set_column('C:C', 18.14, center)
    worksheet140.set_column('D:D', 25, left)
    worksheet140.set_column('E:E', 13.14, left)
    worksheet140.set_column('F:F', 8.57, center)
    worksheet140.set_column('G:R', 5, center)
    worksheet140.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KARAWACI', title)
    worksheet140.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet140.write('A5', 'LOKASI', header)
    worksheet140.write('B5', 'TOTAL', header)
    worksheet140.merge_range('A4:B4', 'RANK', header)
    worksheet140.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet140.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet140.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet140.merge_range('F4:F5', 'KELAS', header)
    worksheet140.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet140.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet140.write('G5', 'MAT', body)
    worksheet140.write('H5', 'SEJ', body)
    worksheet140.write('I5', 'GEO', body)
    worksheet140.write('J5', 'EKO', body)
    worksheet140.write('K5', 'SOS', body)
    worksheet140.write('L5', 'JML', body)
    worksheet140.write('M5', 'MAT', body)
    worksheet140.write('N5', 'SEJ', body)
    worksheet140.write('O5', 'GEO', body)
    worksheet140.write('P5', 'EKO', body)
    worksheet140.write('Q5', 'SOS', body)
    worksheet140.write('R5', 'JML', body)

    worksheet140.conditional_format(5, 0, row140_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet140.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF KARAWACI', title)
    worksheet140.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet140.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet140.write('A22', 'LOKASI', header)
    worksheet140.write('B22', 'TOTAL', header)
    worksheet140.merge_range('A21:B21', 'RANK', header)
    worksheet140.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet140.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet140.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet140.merge_range('F21:F22', 'KELAS', header)
    worksheet140.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet140.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet140.write('G22', 'MAT', body)
    worksheet140.write('H22', 'SEJ', body)
    worksheet140.write('I22', 'GEO', body)
    worksheet140.write('J22', 'EKO', body)
    worksheet140.write('K22', 'SOS', body)
    worksheet140.write('L22', 'JML', body)
    worksheet140.write('M22', 'MAT', body)
    worksheet140.write('N22', 'SEJ', body)
    worksheet140.write('O22', 'GEO', body)
    worksheet140.write('P22', 'EKO', body)
    worksheet140.write('Q22', 'SOS', body)
    worksheet140.write('R22', 'JML', body)

    worksheet140.conditional_format(22, 0, row140+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 141
    worksheet141.insert_image('A1', r'logo resmi nf.jpg')

    worksheet141.set_column('A:A', 7, center)
    worksheet141.set_column('B:B', 6, center)
    worksheet141.set_column('C:C', 18.14, center)
    worksheet141.set_column('D:D', 25, left)
    worksheet141.set_column('E:E', 13.14, left)
    worksheet141.set_column('F:F', 8.57, center)
    worksheet141.set_column('G:R', 5, center)
    worksheet141.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF VETERAN TANGERANG', title)
    worksheet141.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet141.write('A5', 'LOKASI', header)
    worksheet141.write('B5', 'TOTAL', header)
    worksheet141.merge_range('A4:B4', 'RANK', header)
    worksheet141.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet141.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet141.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet141.merge_range('F4:F5', 'KELAS', header)
    worksheet141.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet141.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet141.write('G5', 'MAT', body)
    worksheet141.write('H5', 'SEJ', body)
    worksheet141.write('I5', 'GEO', body)
    worksheet141.write('J5', 'EKO', body)
    worksheet141.write('K5', 'SOS', body)
    worksheet141.write('L5', 'JML', body)
    worksheet141.write('M5', 'MAT', body)
    worksheet141.write('N5', 'SEJ', body)
    worksheet141.write('O5', 'GEO', body)
    worksheet141.write('P5', 'EKO', body)
    worksheet141.write('Q5', 'SOS', body)
    worksheet141.write('R5', 'JML', body)

    worksheet141.conditional_format(5, 0, row141_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet141.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF VETERAN TANGERANG', title)
    worksheet141.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet141.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet141.write('A22', 'LOKASI', header)
    worksheet141.write('B22', 'TOTAL', header)
    worksheet141.merge_range('A21:B21', 'RANK', header)
    worksheet141.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet141.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet141.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet141.merge_range('F21:F22', 'KELAS', header)
    worksheet141.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet141.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet141.write('G22', 'MAT', body)
    worksheet141.write('H22', 'SEJ', body)
    worksheet141.write('I22', 'GEO', body)
    worksheet141.write('J22', 'EKO', body)
    worksheet141.write('K22', 'SOS', body)
    worksheet141.write('L22', 'JML', body)
    worksheet141.write('M22', 'MAT', body)
    worksheet141.write('N22', 'SEJ', body)
    worksheet141.write('O22', 'GEO', body)
    worksheet141.write('P22', 'EKO', body)
    worksheet141.write('Q22', 'SOS', body)
    worksheet141.write('R22', 'JML', body)

    worksheet141.conditional_format(22, 0, row141+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 142
    worksheet142.insert_image('A1', r'logo resmi nf.jpg')

    worksheet142.set_column('A:A', 7, center)
    worksheet142.set_column('B:B', 6, center)
    worksheet142.set_column('C:C', 18.14, center)
    worksheet142.set_column('D:D', 25, left)
    worksheet142.set_column('E:E', 13.14, left)
    worksheet142.set_column('F:F', 8.57, center)
    worksheet142.set_column('G:R', 5, center)
    worksheet142.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PERUMNAS 2 TANGERANG', title)
    worksheet142.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet142.write('A5', 'LOKASI', header)
    worksheet142.write('B5', 'TOTAL', header)
    worksheet142.merge_range('A4:B4', 'RANK', header)
    worksheet142.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet142.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet142.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet142.merge_range('F4:F5', 'KELAS', header)
    worksheet142.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet142.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet142.write('G5', 'MAT', body)
    worksheet142.write('H5', 'SEJ', body)
    worksheet142.write('I5', 'GEO', body)
    worksheet142.write('J5', 'EKO', body)
    worksheet142.write('K5', 'SOS', body)
    worksheet142.write('L5', 'JML', body)
    worksheet142.write('M5', 'MAT', body)
    worksheet142.write('N5', 'SEJ', body)
    worksheet142.write('O5', 'GEO', body)
    worksheet142.write('P5', 'EKO', body)
    worksheet142.write('Q5', 'SOS', body)
    worksheet142.write('R5', 'JML', body)

    worksheet142.conditional_format(5, 0, row142_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet142.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PERUMNAS 2 TANGERANG', title)
    worksheet142.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet142.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet142.write('A22', 'LOKASI', header)
    worksheet142.write('B22', 'TOTAL', header)
    worksheet142.merge_range('A21:B21', 'RANK', header)
    worksheet142.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet142.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet142.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet142.merge_range('F21:F22', 'KELAS', header)
    worksheet142.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet142.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet142.write('G22', 'MAT', body)
    worksheet142.write('H22', 'SEJ', body)
    worksheet142.write('I22', 'GEO', body)
    worksheet142.write('J22', 'EKO', body)
    worksheet142.write('K22', 'SOS', body)
    worksheet142.write('L22', 'JML', body)
    worksheet142.write('M22', 'MAT', body)
    worksheet142.write('N22', 'SEJ', body)
    worksheet142.write('O22', 'GEO', body)
    worksheet142.write('P22', 'EKO', body)
    worksheet142.write('Q22', 'SOS', body)
    worksheet142.write('R22', 'JML', body)

    worksheet142.conditional_format(22, 0, row142+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 143
    worksheet143.insert_image('A1', r'logo resmi nf.jpg')

    worksheet143.set_column('A:A', 7, center)
    worksheet143.set_column('B:B', 6, center)
    worksheet143.set_column('C:C', 18.14, center)
    worksheet143.set_column('D:D', 25, left)
    worksheet143.set_column('E:E', 13.14, left)
    worksheet143.set_column('F:F', 8.57, center)
    worksheet143.set_column('G:R', 5, center)
    worksheet143.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KAYURINGIN', title)
    worksheet143.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet143.write('A5', 'LOKASI', header)
    worksheet143.write('B5', 'TOTAL', header)
    worksheet143.merge_range('A4:B4', 'RANK', header)
    worksheet143.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet143.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet143.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet143.merge_range('F4:F5', 'KELAS', header)
    worksheet143.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet143.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet143.write('G5', 'MAT', body)
    worksheet143.write('H5', 'SEJ', body)
    worksheet143.write('I5', 'GEO', body)
    worksheet143.write('J5', 'EKO', body)
    worksheet143.write('K5', 'SOS', body)
    worksheet143.write('L5', 'JML', body)
    worksheet143.write('M5', 'MAT', body)
    worksheet143.write('N5', 'SEJ', body)
    worksheet143.write('O5', 'GEO', body)
    worksheet143.write('P5', 'EKO', body)
    worksheet143.write('Q5', 'SOS', body)
    worksheet143.write('R5', 'JML', body)

    worksheet143.conditional_format(5, 0, row143_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet143.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF KAYURINGIN', title)
    worksheet143.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet143.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet143.write('A22', 'LOKASI', header)
    worksheet143.write('B22', 'TOTAL', header)
    worksheet143.merge_range('A21:B21', 'RANK', header)
    worksheet143.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet143.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet143.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet143.merge_range('F21:F22', 'KELAS', header)
    worksheet143.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet143.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet143.write('G22', 'MAT', body)
    worksheet143.write('H22', 'SEJ', body)
    worksheet143.write('I22', 'GEO', body)
    worksheet143.write('J22', 'EKO', body)
    worksheet143.write('K22', 'SOS', body)
    worksheet143.write('L22', 'JML', body)
    worksheet143.write('M22', 'MAT', body)
    worksheet143.write('N22', 'SEJ', body)
    worksheet143.write('O22', 'GEO', body)
    worksheet143.write('P22', 'EKO', body)
    worksheet143.write('Q22', 'SOS', body)
    worksheet143.write('R22', 'JML', body)

    worksheet143.conditional_format(22, 0, row143+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 144
    worksheet144.insert_image('A1', r'logo resmi nf.jpg')

    worksheet144.set_column('A:A', 7, center)
    worksheet144.set_column('B:B', 6, center)
    worksheet144.set_column('C:C', 18.14, center)
    worksheet144.set_column('D:D', 25, left)
    worksheet144.set_column('E:E', 13.14, left)
    worksheet144.set_column('F:F', 8.57, center)
    worksheet144.set_column('G:R', 5, center)
    worksheet144.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF AGUS SALIM', title)
    worksheet144.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet144.write('A5', 'LOKASI', header)
    worksheet144.write('B5', 'TOTAL', header)
    worksheet144.merge_range('A4:B4', 'RANK', header)
    worksheet144.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet144.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet144.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet144.merge_range('F4:F5', 'KELAS', header)
    worksheet144.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet144.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet144.write('G5', 'MAT', body)
    worksheet144.write('H5', 'SEJ', body)
    worksheet144.write('I5', 'GEO', body)
    worksheet144.write('J5', 'EKO', body)
    worksheet144.write('K5', 'SOS', body)
    worksheet144.write('L5', 'JML', body)
    worksheet144.write('M5', 'MAT', body)
    worksheet144.write('N5', 'SEJ', body)
    worksheet144.write('O5', 'GEO', body)
    worksheet144.write('P5', 'EKO', body)
    worksheet144.write('Q5', 'SOS', body)
    worksheet144.write('R5', 'JML', body)

    worksheet144.conditional_format(5, 0, row144_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet144.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF AGUS SALIM', title)
    worksheet144.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet144.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet144.write('A22', 'LOKASI', header)
    worksheet144.write('B22', 'TOTAL', header)
    worksheet144.merge_range('A21:B21', 'RANK', header)
    worksheet144.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet144.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet144.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet144.merge_range('F21:F22', 'KELAS', header)
    worksheet144.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet144.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet144.write('G22', 'MAT', body)
    worksheet144.write('H22', 'SEJ', body)
    worksheet144.write('I22', 'GEO', body)
    worksheet144.write('J22', 'EKO', body)
    worksheet144.write('K22', 'SOS', body)
    worksheet144.write('L22', 'JML', body)
    worksheet144.write('M22', 'MAT', body)
    worksheet144.write('N22', 'SEJ', body)
    worksheet144.write('O22', 'GEO', body)
    worksheet144.write('P22', 'EKO', body)
    worksheet144.write('Q22', 'SOS', body)
    worksheet144.write('R22', 'JML', body)

    worksheet144.conditional_format(22, 0, row144+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 145
    worksheet145.insert_image('A1', r'logo resmi nf.jpg')

    worksheet145.set_column('A:A', 7, center)
    worksheet145.set_column('B:B', 6, center)
    worksheet145.set_column('C:C', 18.14, center)
    worksheet145.set_column('D:D', 25, left)
    worksheet145.set_column('E:E', 13.14, left)
    worksheet145.set_column('F:F', 8.57, center)
    worksheet145.set_column('G:R', 5, center)
    worksheet145.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SUMERU', title)
    worksheet145.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet145.write('A5', 'LOKASI', header)
    worksheet145.write('B5', 'TOTAL', header)
    worksheet145.merge_range('A4:B4', 'RANK', header)
    worksheet145.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet145.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet145.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet145.merge_range('F4:F5', 'KELAS', header)
    worksheet145.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet145.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet145.write('G5', 'MAT', body)
    worksheet145.write('H5', 'SEJ', body)
    worksheet145.write('I5', 'GEO', body)
    worksheet145.write('J5', 'EKO', body)
    worksheet145.write('K5', 'SOS', body)
    worksheet145.write('L5', 'JML', body)
    worksheet145.write('M5', 'MAT', body)
    worksheet145.write('N5', 'SEJ', body)
    worksheet145.write('O5', 'GEO', body)
    worksheet145.write('P5', 'EKO', body)
    worksheet145.write('Q5', 'SOS', body)
    worksheet145.write('R5', 'JML', body)

    worksheet145.conditional_format(5, 0, row145_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet145.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF SUMERU', title)
    worksheet145.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet145.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet145.write('A22', 'LOKASI', header)
    worksheet145.write('B22', 'TOTAL', header)
    worksheet145.merge_range('A21:B21', 'RANK', header)
    worksheet145.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet145.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet145.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet145.merge_range('F21:F22', 'KELAS', header)
    worksheet145.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet145.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet145.write('G22', 'MAT', body)
    worksheet145.write('H22', 'SEJ', body)
    worksheet145.write('I22', 'GEO', body)
    worksheet145.write('J22', 'EKO', body)
    worksheet145.write('K22', 'SOS', body)
    worksheet145.write('L22', 'JML', body)
    worksheet145.write('M22', 'MAT', body)
    worksheet145.write('N22', 'SEJ', body)
    worksheet145.write('O22', 'GEO', body)
    worksheet145.write('P22', 'EKO', body)
    worksheet145.write('Q22', 'SOS', body)
    worksheet145.write('R22', 'JML', body)

    worksheet145.conditional_format(22, 0, row145+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 146
    worksheet146.insert_image('A1', r'logo resmi nf.jpg')

    worksheet146.set_column('A:A', 7, center)
    worksheet146.set_column('B:B', 6, center)
    worksheet146.set_column('C:C', 18.14, center)
    worksheet146.set_column('D:D', 25, left)
    worksheet146.set_column('E:E', 13.14, left)
    worksheet146.set_column('F:F', 8.57, center)
    worksheet146.set_column('G:R', 5, center)
    worksheet146.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIKEAS', title)
    worksheet146.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet146.write('A5', 'LOKASI', header)
    worksheet146.write('B5', 'TOTAL', header)
    worksheet146.merge_range('A4:B4', 'RANK', header)
    worksheet146.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet146.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet146.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet146.merge_range('F4:F5', 'KELAS', header)
    worksheet146.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet146.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet146.write('G5', 'MAT', body)
    worksheet146.write('H5', 'SEJ', body)
    worksheet146.write('I5', 'GEO', body)
    worksheet146.write('J5', 'EKO', body)
    worksheet146.write('K5', 'SOS', body)
    worksheet146.write('L5', 'JML', body)
    worksheet146.write('M5', 'MAT', body)
    worksheet146.write('N5', 'SEJ', body)
    worksheet146.write('O5', 'GEO', body)
    worksheet146.write('P5', 'EKO', body)
    worksheet146.write('Q5', 'SOS', body)
    worksheet146.write('R5', 'JML', body)

    worksheet146.conditional_format(5, 0, row146_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet146.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CIKEAS', title)
    worksheet146.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet146.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet146.write('A22', 'LOKASI', header)
    worksheet146.write('B22', 'TOTAL', header)
    worksheet146.merge_range('A21:B21', 'RANK', header)
    worksheet146.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet146.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet146.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet146.merge_range('F21:F22', 'KELAS', header)
    worksheet146.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet146.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet146.write('G22', 'MAT', body)
    worksheet146.write('H22', 'SEJ', body)
    worksheet146.write('I22', 'GEO', body)
    worksheet146.write('J22', 'EKO', body)
    worksheet146.write('K22', 'SOS', body)
    worksheet146.write('L22', 'JML', body)
    worksheet146.write('M22', 'MAT', body)
    worksheet146.write('N22', 'SEJ', body)
    worksheet146.write('O22', 'GEO', body)
    worksheet146.write('P22', 'EKO', body)
    worksheet146.write('Q22', 'SOS', body)
    worksheet146.write('R22', 'JML', body)

    worksheet146.conditional_format(22, 0, row146+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 148
    worksheet148.insert_image('A1', r'logo resmi nf.jpg')

    worksheet148.set_column('A:A', 7, center)
    worksheet148.set_column('B:B', 6, center)
    worksheet148.set_column('C:C', 18.14, center)
    worksheet148.set_column('D:D', 25, left)
    worksheet148.set_column('E:E', 13.14, left)
    worksheet148.set_column('F:F', 8.57, center)
    worksheet148.set_column('G:R', 5, center)
    worksheet148.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIJAWA MASJID', title)
    worksheet148.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet148.write('A5', 'LOKASI', header)
    worksheet148.write('B5', 'TOTAL', header)
    worksheet148.merge_range('A4:B4', 'RANK', header)
    worksheet148.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet148.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet148.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet148.merge_range('F4:F5', 'KELAS', header)
    worksheet148.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet148.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet148.write('G5', 'MAT', body)
    worksheet148.write('H5', 'SEJ', body)
    worksheet148.write('I5', 'GEO', body)
    worksheet148.write('J5', 'EKO', body)
    worksheet148.write('K5', 'SOS', body)
    worksheet148.write('L5', 'JML', body)
    worksheet148.write('M5', 'MAT', body)
    worksheet148.write('N5', 'SEJ', body)
    worksheet148.write('O5', 'GEO', body)
    worksheet148.write('P5', 'EKO', body)
    worksheet148.write('Q5', 'SOS', body)
    worksheet148.write('R5', 'JML', body)

    worksheet148.conditional_format(5, 0, row148_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet148.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CIJAWA MASJID', title)
    worksheet148.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet148.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet148.write('A22', 'LOKASI', header)
    worksheet148.write('B22', 'TOTAL', header)
    worksheet148.merge_range('A21:B21', 'RANK', header)
    worksheet148.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet148.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet148.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet148.merge_range('F21:F22', 'KELAS', header)
    worksheet148.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet148.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet148.write('G22', 'MAT', body)
    worksheet148.write('H22', 'SEJ', body)
    worksheet148.write('I22', 'GEO', body)
    worksheet148.write('J22', 'EKO', body)
    worksheet148.write('K22', 'SOS', body)
    worksheet148.write('L22', 'JML', body)
    worksheet148.write('M22', 'MAT', body)
    worksheet148.write('N22', 'SEJ', body)
    worksheet148.write('O22', 'GEO', body)
    worksheet148.write('P22', 'EKO', body)
    worksheet148.write('Q22', 'SOS', body)
    worksheet148.write('R22', 'JML', body)

    worksheet148.conditional_format(22, 0, row148+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 149
    worksheet149.insert_image('A1', r'logo resmi nf.jpg')

    worksheet149.set_column('A:A', 7, center)
    worksheet149.set_column('B:B', 6, center)
    worksheet149.set_column('C:C', 18.14, center)
    worksheet149.set_column('D:D', 25, left)
    worksheet149.set_column('E:E', 13.14, left)
    worksheet149.set_column('F:F', 8.57, center)
    worksheet149.set_column('G:R', 5, center)
    worksheet149.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PALEDANG', title)
    worksheet149.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet149.write('A5', 'LOKASI', header)
    worksheet149.write('B5', 'TOTAL', header)
    worksheet149.merge_range('A4:B4', 'RANK', header)
    worksheet149.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet149.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet149.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet149.merge_range('F4:F5', 'KELAS', header)
    worksheet149.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet149.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet149.write('G5', 'MAT', body)
    worksheet149.write('H5', 'SEJ', body)
    worksheet149.write('I5', 'GEO', body)
    worksheet149.write('J5', 'EKO', body)
    worksheet149.write('K5', 'SOS', body)
    worksheet149.write('L5', 'JML', body)
    worksheet149.write('M5', 'MAT', body)
    worksheet149.write('N5', 'SEJ', body)
    worksheet149.write('O5', 'GEO', body)
    worksheet149.write('P5', 'EKO', body)
    worksheet149.write('Q5', 'SOS', body)
    worksheet149.write('R5', 'JML', body)

    worksheet149.conditional_format(5, 0, row149_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet149.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PALEDANG', title)
    worksheet149.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet149.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet149.write('A22', 'LOKASI', header)
    worksheet149.write('B22', 'TOTAL', header)
    worksheet149.merge_range('A21:B21', 'RANK', header)
    worksheet149.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet149.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet149.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet149.merge_range('F21:F22', 'KELAS', header)
    worksheet149.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet149.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet149.write('G22', 'MAT', body)
    worksheet149.write('H22', 'SEJ', body)
    worksheet149.write('I22', 'GEO', body)
    worksheet149.write('J22', 'EKO', body)
    worksheet149.write('K22', 'SOS', body)
    worksheet149.write('L22', 'JML', body)
    worksheet149.write('M22', 'MAT', body)
    worksheet149.write('N22', 'SEJ', body)
    worksheet149.write('O22', 'GEO', body)
    worksheet149.write('P22', 'EKO', body)
    worksheet149.write('Q22', 'SOS', body)
    worksheet149.write('R22', 'JML', body)

    worksheet149.conditional_format(22, 0, row149+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 150
    worksheet150.insert_image('A1', r'logo resmi nf.jpg')

    worksheet150.set_column('A:A', 7, center)
    worksheet150.set_column('B:B', 6, center)
    worksheet150.set_column('C:C', 18.14, center)
    worksheet150.set_column('D:D', 25, left)
    worksheet150.set_column('E:E', 13.14, left)
    worksheet150.set_column('F:F', 8.57, center)
    worksheet150.set_column('G:R', 5, center)
    worksheet150.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF GEDONG KUNING', title)
    worksheet150.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet150.write('A5', 'LOKASI', header)
    worksheet150.write('B5', 'TOTAL', header)
    worksheet150.merge_range('A4:B4', 'RANK', header)
    worksheet150.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet150.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet150.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet150.merge_range('F4:F5', 'KELAS', header)
    worksheet150.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet150.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet150.write('G5', 'MAT', body)
    worksheet150.write('H5', 'SEJ', body)
    worksheet150.write('I5', 'GEO', body)
    worksheet150.write('J5', 'EKO', body)
    worksheet150.write('K5', 'SOS', body)
    worksheet150.write('L5', 'JML', body)
    worksheet150.write('M5', 'MAT', body)
    worksheet150.write('N5', 'SEJ', body)
    worksheet150.write('O5', 'GEO', body)
    worksheet150.write('P5', 'EKO', body)
    worksheet150.write('Q5', 'SOS', body)
    worksheet150.write('R5', 'JML', body)

    worksheet150.conditional_format(5, 0, row150_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet150.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF GEDONG KUNING', title)
    worksheet150.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet150.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet150.write('A22', 'LOKASI', header)
    worksheet150.write('B22', 'TOTAL', header)
    worksheet150.merge_range('A21:B21', 'RANK', header)
    worksheet150.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet150.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet150.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet150.merge_range('F21:F22', 'KELAS', header)
    worksheet150.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet150.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet150.write('G22', 'MAT', body)
    worksheet150.write('H22', 'SEJ', body)
    worksheet150.write('I22', 'GEO', body)
    worksheet150.write('J22', 'EKO', body)
    worksheet150.write('K22', 'SOS', body)
    worksheet150.write('L22', 'JML', body)
    worksheet150.write('M22', 'MAT', body)
    worksheet150.write('N22', 'SEJ', body)
    worksheet150.write('O22', 'GEO', body)
    worksheet150.write('P22', 'EKO', body)
    worksheet150.write('Q22', 'SOS', body)
    worksheet150.write('R22', 'JML', body)

    worksheet150.conditional_format(22, 0, row150+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 151
    worksheet151.insert_image('A1', r'logo resmi nf.jpg')

    worksheet151.set_column('A:A', 7, center)
    worksheet151.set_column('B:B', 6, center)
    worksheet151.set_column('C:C', 18.14, center)
    worksheet151.set_column('D:D', 25, left)
    worksheet151.set_column('E:E', 13.14, left)
    worksheet151.set_column('F:F', 8.57, center)
    worksheet151.set_column('G:R', 5, center)
    worksheet151.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF JATIWARINGIN', title)
    worksheet151.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet151.write('A5', 'LOKASI', header)
    worksheet151.write('B5', 'TOTAL', header)
    worksheet151.merge_range('A4:B4', 'RANK', header)
    worksheet151.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet151.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet151.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet151.merge_range('F4:F5', 'KELAS', header)
    worksheet151.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet151.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet151.write('G5', 'MAT', body)
    worksheet151.write('H5', 'SEJ', body)
    worksheet151.write('I5', 'GEO', body)
    worksheet151.write('J5', 'EKO', body)
    worksheet151.write('K5', 'SOS', body)
    worksheet151.write('L5', 'JML', body)
    worksheet151.write('M5', 'MAT', body)
    worksheet151.write('N5', 'SEJ', body)
    worksheet151.write('O5', 'GEO', body)
    worksheet151.write('P5', 'EKO', body)
    worksheet151.write('Q5', 'SOS', body)
    worksheet151.write('R5', 'JML', body)

    worksheet151.conditional_format(5, 0, row151_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet151.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF JATIWARINGIN', title)
    worksheet151.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet151.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet151.write('A22', 'LOKASI', header)
    worksheet151.write('B22', 'TOTAL', header)
    worksheet151.merge_range('A21:B21', 'RANK', header)
    worksheet151.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet151.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet151.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet151.merge_range('F21:F22', 'KELAS', header)
    worksheet151.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet151.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet151.write('G22', 'MAT', body)
    worksheet151.write('H22', 'SEJ', body)
    worksheet151.write('I22', 'GEO', body)
    worksheet151.write('J22', 'EKO', body)
    worksheet151.write('K22', 'SOS', body)
    worksheet151.write('L22', 'JML', body)
    worksheet151.write('M22', 'MAT', body)
    worksheet151.write('N22', 'SEJ', body)
    worksheet151.write('O22', 'GEO', body)
    worksheet151.write('P22', 'EKO', body)
    worksheet151.write('Q22', 'SOS', body)
    worksheet151.write('R22', 'JML', body)

    worksheet151.conditional_format(22, 0, row151+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 152
    worksheet152.insert_image('A1', r'logo resmi nf.jpg')

    worksheet152.set_column('A:A', 7, center)
    worksheet152.set_column('B:B', 6, center)
    worksheet152.set_column('C:C', 18.14, center)
    worksheet152.set_column('D:D', 25, left)
    worksheet152.set_column('E:E', 13.14, left)
    worksheet152.set_column('F:F', 8.57, center)
    worksheet152.set_column('G:R', 5, center)
    worksheet152.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CILEDUG', title)
    worksheet152.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet152.write('A5', 'LOKASI', header)
    worksheet152.write('B5', 'TOTAL', header)
    worksheet152.merge_range('A4:B4', 'RANK', header)
    worksheet152.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet152.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet152.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet152.merge_range('F4:F5', 'KELAS', header)
    worksheet152.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet152.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet152.write('G5', 'MAT', body)
    worksheet152.write('H5', 'SEJ', body)
    worksheet152.write('I5', 'GEO', body)
    worksheet152.write('J5', 'EKO', body)
    worksheet152.write('K5', 'SOS', body)
    worksheet152.write('L5', 'JML', body)
    worksheet152.write('M5', 'MAT', body)
    worksheet152.write('N5', 'SEJ', body)
    worksheet152.write('O5', 'GEO', body)
    worksheet152.write('P5', 'EKO', body)
    worksheet152.write('Q5', 'SOS', body)
    worksheet152.write('R5', 'JML', body)

    worksheet152.conditional_format(5, 0, row152_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet152.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CILEDUG', title)
    worksheet152.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet152.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet152.write('A22', 'LOKASI', header)
    worksheet152.write('B22', 'TOTAL', header)
    worksheet152.merge_range('A21:B21', 'RANK', header)
    worksheet152.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet152.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet152.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet152.merge_range('F21:F22', 'KELAS', header)
    worksheet152.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet152.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet152.write('G22', 'MAT', body)
    worksheet152.write('H22', 'SEJ', body)
    worksheet152.write('I22', 'GEO', body)
    worksheet152.write('J22', 'EKO', body)
    worksheet152.write('K22', 'SOS', body)
    worksheet152.write('L22', 'JML', body)
    worksheet152.write('M22', 'MAT', body)
    worksheet152.write('N22', 'SEJ', body)
    worksheet152.write('O22', 'GEO', body)
    worksheet152.write('P22', 'EKO', body)
    worksheet152.write('Q22', 'SOS', body)
    worksheet152.write('R22', 'JML', body)

    worksheet152.conditional_format(22, 0, row152+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 153
    worksheet153.insert_image('A1', r'logo resmi nf.jpg')

    worksheet153.set_column('A:A', 7, center)
    worksheet153.set_column('B:B', 6, center)
    worksheet153.set_column('C:C', 18.14, center)
    worksheet153.set_column('D:D', 25, left)
    worksheet153.set_column('E:E', 13.14, left)
    worksheet153.set_column('F:F', 8.57, center)
    worksheet153.set_column('G:R', 5, center)
    worksheet153.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KRANGGAN', title)
    worksheet153.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet153.write('A5', 'LOKASI', header)
    worksheet153.write('B5', 'TOTAL', header)
    worksheet153.merge_range('A4:B4', 'RANK', header)
    worksheet153.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet153.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet153.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet153.merge_range('F4:F5', 'KELAS', header)
    worksheet153.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet153.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet153.write('G5', 'MAT', body)
    worksheet153.write('H5', 'SEJ', body)
    worksheet153.write('I5', 'GEO', body)
    worksheet153.write('J5', 'EKO', body)
    worksheet153.write('K5', 'SOS', body)
    worksheet153.write('L5', 'JML', body)
    worksheet153.write('M5', 'MAT', body)
    worksheet153.write('N5', 'SEJ', body)
    worksheet153.write('O5', 'GEO', body)
    worksheet153.write('P5', 'EKO', body)
    worksheet153.write('Q5', 'SOS', body)
    worksheet153.write('R5', 'JML', body)

    worksheet153.conditional_format(5, 0, row153_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet153.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF KRANGGAN', title)
    worksheet153.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet153.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet153.write('A22', 'LOKASI', header)
    worksheet153.write('B22', 'TOTAL', header)
    worksheet153.merge_range('A21:B21', 'RANK', header)
    worksheet153.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet153.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet153.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet153.merge_range('F21:F22', 'KELAS', header)
    worksheet153.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet153.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet153.write('G22', 'MAT', body)
    worksheet153.write('H22', 'SEJ', body)
    worksheet153.write('I22', 'GEO', body)
    worksheet153.write('J22', 'EKO', body)
    worksheet153.write('K22', 'SOS', body)
    worksheet153.write('L22', 'JML', body)
    worksheet153.write('M22', 'MAT', body)
    worksheet153.write('N22', 'SEJ', body)
    worksheet153.write('O22', 'GEO', body)
    worksheet153.write('P22', 'EKO', body)
    worksheet153.write('Q22', 'SOS', body)
    worksheet153.write('R22', 'JML', body)

    worksheet153.conditional_format(22, 0, row153+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 154
    worksheet154.insert_image('A1', r'logo resmi nf.jpg')

    worksheet154.set_column('A:A', 7, center)
    worksheet154.set_column('B:B', 6, center)
    worksheet154.set_column('C:C', 18.14, center)
    worksheet154.set_column('D:D', 25, left)
    worksheet154.set_column('E:E', 13.14, left)
    worksheet154.set_column('F:F', 8.57, center)
    worksheet154.set_column('G:R', 5, center)
    worksheet154.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MUSTIKA JAYA', title)
    worksheet154.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet154.write('A5', 'LOKASI', header)
    worksheet154.write('B5', 'TOTAL', header)
    worksheet154.merge_range('A4:B4', 'RANK', header)
    worksheet154.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet154.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet154.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet154.merge_range('F4:F5', 'KELAS', header)
    worksheet154.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet154.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet154.write('G5', 'MAT', body)
    worksheet154.write('H5', 'SEJ', body)
    worksheet154.write('I5', 'GEO', body)
    worksheet154.write('J5', 'EKO', body)
    worksheet154.write('K5', 'SOS', body)
    worksheet154.write('L5', 'JML', body)
    worksheet154.write('M5', 'MAT', body)
    worksheet154.write('N5', 'SEJ', body)
    worksheet154.write('O5', 'GEO', body)
    worksheet154.write('P5', 'EKO', body)
    worksheet154.write('Q5', 'SOS', body)
    worksheet154.write('R5', 'JML', body)

    worksheet154.conditional_format(5, 0, row154_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet154.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF MUSTIKA JAYA', title)
    worksheet154.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet154.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet154.write('A22', 'LOKASI', header)
    worksheet154.write('B22', 'TOTAL', header)
    worksheet154.merge_range('A21:B21', 'RANK', header)
    worksheet154.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet154.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet154.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet154.merge_range('F21:F22', 'KELAS', header)
    worksheet154.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet154.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet154.write('G22', 'MAT', body)
    worksheet154.write('H22', 'SEJ', body)
    worksheet154.write('I22', 'GEO', body)
    worksheet154.write('J22', 'EKO', body)
    worksheet154.write('K22', 'SOS', body)
    worksheet154.write('L22', 'JML', body)
    worksheet154.write('M22', 'MAT', body)
    worksheet154.write('N22', 'SEJ', body)
    worksheet154.write('O22', 'GEO', body)
    worksheet154.write('P22', 'EKO', body)
    worksheet154.write('Q22', 'SOS', body)
    worksheet154.write('R22', 'JML', body)

    worksheet154.conditional_format(22, 0, row154+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 155
    worksheet155.insert_image('A1', r'logo resmi nf.jpg')

    worksheet155.set_column('A:A', 7, center)
    worksheet155.set_column('B:B', 6, center)
    worksheet155.set_column('C:C', 18.14, center)
    worksheet155.set_column('D:D', 25, left)
    worksheet155.set_column('E:E', 13.14, left)
    worksheet155.set_column('F:F', 8.57, center)
    worksheet155.set_column('G:R', 5, center)
    worksheet155.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF ALEXINDO', title)
    worksheet155.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet155.write('A5', 'LOKASI', header)
    worksheet155.write('B5', 'TOTAL', header)
    worksheet155.merge_range('A4:B4', 'RANK', header)
    worksheet155.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet155.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet155.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet155.merge_range('F4:F5', 'KELAS', header)
    worksheet155.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet155.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet155.write('G5', 'MAT', body)
    worksheet155.write('H5', 'SEJ', body)
    worksheet155.write('I5', 'GEO', body)
    worksheet155.write('J5', 'EKO', body)
    worksheet155.write('K5', 'SOS', body)
    worksheet155.write('L5', 'JML', body)
    worksheet155.write('M5', 'MAT', body)
    worksheet155.write('N5', 'SEJ', body)
    worksheet155.write('O5', 'GEO', body)
    worksheet155.write('P5', 'EKO', body)
    worksheet155.write('Q5', 'SOS', body)
    worksheet155.write('R5', 'JML', body)

    worksheet155.conditional_format(5, 0, row155_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet155.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF ALEXINDO', title)
    worksheet155.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet155.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet155.write('A22', 'LOKASI', header)
    worksheet155.write('B22', 'TOTAL', header)
    worksheet155.merge_range('A21:B21', 'RANK', header)
    worksheet155.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet155.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet155.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet155.merge_range('F21:F22', 'KELAS', header)
    worksheet155.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet155.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet155.write('G22', 'MAT', body)
    worksheet155.write('H22', 'SEJ', body)
    worksheet155.write('I22', 'GEO', body)
    worksheet155.write('J22', 'EKO', body)
    worksheet155.write('K22', 'SOS', body)
    worksheet155.write('L22', 'JML', body)
    worksheet155.write('M22', 'MAT', body)
    worksheet155.write('N22', 'SEJ', body)
    worksheet155.write('O22', 'GEO', body)
    worksheet155.write('P22', 'EKO', body)
    worksheet155.write('Q22', 'SOS', body)
    worksheet155.write('R22', 'JML', body)

    worksheet155.conditional_format(22, 0, row155+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 156
    worksheet156.insert_image('A1', r'logo resmi nf.jpg')

    worksheet156.set_column('A:A', 7, center)
    worksheet156.set_column('B:B', 6, center)
    worksheet156.set_column('C:C', 18.14, center)
    worksheet156.set_column('D:D', 25, left)
    worksheet156.set_column('E:E', 13.14, left)
    worksheet156.set_column('F:F', 8.57, center)
    worksheet156.set_column('G:R', 5, center)
    worksheet156.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIBITUNG', title)
    worksheet156.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet156.write('A5', 'LOKASI', header)
    worksheet156.write('B5', 'TOTAL', header)
    worksheet156.merge_range('A4:B4', 'RANK', header)
    worksheet156.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet156.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet156.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet156.merge_range('F4:F5', 'KELAS', header)
    worksheet156.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet156.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet156.write('G5', 'MAT', body)
    worksheet156.write('H5', 'SEJ', body)
    worksheet156.write('I5', 'GEO', body)
    worksheet156.write('J5', 'EKO', body)
    worksheet156.write('K5', 'SOS', body)
    worksheet156.write('L5', 'JML', body)
    worksheet156.write('M5', 'MAT', body)
    worksheet156.write('N5', 'SEJ', body)
    worksheet156.write('O5', 'GEO', body)
    worksheet156.write('P5', 'EKO', body)
    worksheet156.write('Q5', 'SOS', body)
    worksheet156.write('R5', 'JML', body)

    worksheet156.conditional_format(5, 0, row156_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet156.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CIBITUNG', title)
    worksheet156.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet156.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet156.write('A22', 'LOKASI', header)
    worksheet156.write('B22', 'TOTAL', header)
    worksheet156.merge_range('A21:B21', 'RANK', header)
    worksheet156.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet156.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet156.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet156.merge_range('F21:F22', 'KELAS', header)
    worksheet156.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet156.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet156.write('G22', 'MAT', body)
    worksheet156.write('H22', 'SEJ', body)
    worksheet156.write('I22', 'GEO', body)
    worksheet156.write('J22', 'EKO', body)
    worksheet156.write('K22', 'SOS', body)
    worksheet156.write('L22', 'JML', body)
    worksheet156.write('M22', 'MAT', body)
    worksheet156.write('N22', 'SEJ', body)
    worksheet156.write('O22', 'GEO', body)
    worksheet156.write('P22', 'EKO', body)
    worksheet156.write('Q22', 'SOS', body)
    worksheet156.write('R22', 'JML', body)

    worksheet156.conditional_format(22, 0, row156+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 157
    worksheet157.insert_image('A1', r'logo resmi nf.jpg')

    worksheet157.set_column('A:A', 7, center)
    worksheet157.set_column('B:B', 6, center)
    worksheet157.set_column('C:C', 18.14, center)
    worksheet157.set_column('D:D', 25, left)
    worksheet157.set_column('E:E', 13.14, left)
    worksheet157.set_column('F:F', 8.57, center)
    worksheet157.set_column('G:R', 5, center)
    worksheet157.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KRAMAT JAYA', title)
    worksheet157.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet157.write('A5', 'LOKASI', header)
    worksheet157.write('B5', 'TOTAL', header)
    worksheet157.merge_range('A4:B4', 'RANK', header)
    worksheet157.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet157.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet157.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet157.merge_range('F4:F5', 'KELAS', header)
    worksheet157.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet157.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet157.write('G5', 'MAT', body)
    worksheet157.write('H5', 'SEJ', body)
    worksheet157.write('I5', 'GEO', body)
    worksheet157.write('J5', 'EKO', body)
    worksheet157.write('K5', 'SOS', body)
    worksheet157.write('L5', 'JML', body)
    worksheet157.write('M5', 'MAT', body)
    worksheet157.write('N5', 'SEJ', body)
    worksheet157.write('O5', 'GEO', body)
    worksheet157.write('P5', 'EKO', body)
    worksheet157.write('Q5', 'SOS', body)
    worksheet157.write('R5', 'JML', body)

    worksheet157.conditional_format(5, 0, row157_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet157.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF KRAMAT JAYA', title)
    worksheet157.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet157.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet157.write('A22', 'LOKASI', header)
    worksheet157.write('B22', 'TOTAL', header)
    worksheet157.merge_range('A21:B21', 'RANK', header)
    worksheet157.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet157.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet157.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet157.merge_range('F21:F22', 'KELAS', header)
    worksheet157.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet157.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet157.write('G22', 'MAT', body)
    worksheet157.write('H22', 'SEJ', body)
    worksheet157.write('I22', 'GEO', body)
    worksheet157.write('J22', 'EKO', body)
    worksheet157.write('K22', 'SOS', body)
    worksheet157.write('L22', 'JML', body)
    worksheet157.write('M22', 'MAT', body)
    worksheet157.write('N22', 'SEJ', body)
    worksheet157.write('O22', 'GEO', body)
    worksheet157.write('P22', 'EKO', body)
    worksheet157.write('Q22', 'SOS', body)
    worksheet157.write('R22', 'JML', body)

    worksheet157.conditional_format(22, 0, row157+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 158
    worksheet158.insert_image('A1', r'logo resmi nf.jpg')

    worksheet158.set_column('A:A', 7, center)
    worksheet158.set_column('B:B', 6, center)
    worksheet158.set_column('C:C', 18.14, center)
    worksheet158.set_column('D:D', 25, left)
    worksheet158.set_column('E:E', 13.14, left)
    worksheet158.set_column('F:F', 8.57, center)
    worksheet158.set_column('G:R', 5, center)
    worksheet158.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PONDOK GEDE', title)
    worksheet158.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet158.write('A5', 'LOKASI', header)
    worksheet158.write('B5', 'TOTAL', header)
    worksheet158.merge_range('A4:B4', 'RANK', header)
    worksheet158.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet158.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet158.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet158.merge_range('F4:F5', 'KELAS', header)
    worksheet158.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet158.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet158.write('G5', 'MAT', body)
    worksheet158.write('H5', 'SEJ', body)
    worksheet158.write('I5', 'GEO', body)
    worksheet158.write('J5', 'EKO', body)
    worksheet158.write('K5', 'SOS', body)
    worksheet158.write('L5', 'JML', body)
    worksheet158.write('M5', 'MAT', body)
    worksheet158.write('N5', 'SEJ', body)
    worksheet158.write('O5', 'GEO', body)
    worksheet158.write('P5', 'EKO', body)
    worksheet158.write('Q5', 'SOS', body)
    worksheet158.write('R5', 'JML', body)

    worksheet158.conditional_format(5, 0, row158_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet158.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PONDOK GEDE', title)
    worksheet158.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet158.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet158.write('A22', 'LOKASI', header)
    worksheet158.write('B22', 'TOTAL', header)
    worksheet158.merge_range('A21:B21', 'RANK', header)
    worksheet158.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet158.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet158.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet158.merge_range('F21:F22', 'KELAS', header)
    worksheet158.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet158.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet158.write('G22', 'MAT', body)
    worksheet158.write('H22', 'SEJ', body)
    worksheet158.write('I22', 'GEO', body)
    worksheet158.write('J22', 'EKO', body)
    worksheet158.write('K22', 'SOS', body)
    worksheet158.write('L22', 'JML', body)
    worksheet158.write('M22', 'MAT', body)
    worksheet158.write('N22', 'SEJ', body)
    worksheet158.write('O22', 'GEO', body)
    worksheet158.write('P22', 'EKO', body)
    worksheet158.write('Q22', 'SOS', body)
    worksheet158.write('R22', 'JML', body)

    worksheet158.conditional_format(22, 0, row158+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 159
    worksheet159.insert_image('A1', r'logo resmi nf.jpg')

    worksheet159.set_column('A:A', 7, center)
    worksheet159.set_column('B:B', 6, center)
    worksheet159.set_column('C:C', 18.14, center)
    worksheet159.set_column('D:D', 25, left)
    worksheet159.set_column('E:E', 13.14, left)
    worksheet159.set_column('F:F', 8.57, center)
    worksheet159.set_column('G:R', 5, center)
    worksheet159.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF GALAXY', title)
    worksheet159.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet159.write('A5', 'LOKASI', header)
    worksheet159.write('B5', 'TOTAL', header)
    worksheet159.merge_range('A4:B4', 'RANK', header)
    worksheet159.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet159.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet159.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet159.merge_range('F4:F5', 'KELAS', header)
    worksheet159.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet159.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet159.write('G5', 'MAT', body)
    worksheet159.write('H5', 'SEJ', body)
    worksheet159.write('I5', 'GEO', body)
    worksheet159.write('J5', 'EKO', body)
    worksheet159.write('K5', 'SOS', body)
    worksheet159.write('L5', 'JML', body)
    worksheet159.write('M5', 'MAT', body)
    worksheet159.write('N5', 'SEJ', body)
    worksheet159.write('O5', 'GEO', body)
    worksheet159.write('P5', 'EKO', body)
    worksheet159.write('Q5', 'SOS', body)
    worksheet159.write('R5', 'JML', body)

    worksheet159.conditional_format(5, 0, row159_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet159.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF GALAXY', title)
    worksheet159.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet159.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet159.write('A22', 'LOKASI', header)
    worksheet159.write('B22', 'TOTAL', header)
    worksheet159.merge_range('A21:B21', 'RANK', header)
    worksheet159.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet159.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet159.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet159.merge_range('F21:F22', 'KELAS', header)
    worksheet159.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet159.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet159.write('G22', 'MAT', body)
    worksheet159.write('H22', 'SEJ', body)
    worksheet159.write('I22', 'GEO', body)
    worksheet159.write('J22', 'EKO', body)
    worksheet159.write('K22', 'SOS', body)
    worksheet159.write('L22', 'JML', body)
    worksheet159.write('M22', 'MAT', body)
    worksheet159.write('N22', 'SEJ', body)
    worksheet159.write('O22', 'GEO', body)
    worksheet159.write('P22', 'EKO', body)
    worksheet159.write('Q22', 'SOS', body)
    worksheet159.write('R22', 'JML', body)

    worksheet159.conditional_format(22, 0, row159+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 160
    worksheet160.insert_image('A1', r'logo resmi nf.jpg')

    worksheet160.set_column('A:A', 7, center)
    worksheet160.set_column('B:B', 6, center)
    worksheet160.set_column('C:C', 18.14, center)
    worksheet160.set_column('D:D', 25, left)
    worksheet160.set_column('E:E', 13.14, left)
    worksheet160.set_column('F:F', 8.57, center)
    worksheet160.set_column('G:R', 5, center)
    worksheet160.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIGANJUR', title)
    worksheet160.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet160.write('A5', 'LOKASI', header)
    worksheet160.write('B5', 'TOTAL', header)
    worksheet160.merge_range('A4:B4', 'RANK', header)
    worksheet160.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet160.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet160.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet160.merge_range('F4:F5', 'KELAS', header)
    worksheet160.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet160.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet160.write('G5', 'MAT', body)
    worksheet160.write('H5', 'SEJ', body)
    worksheet160.write('I5', 'GEO', body)
    worksheet160.write('J5', 'EKO', body)
    worksheet160.write('K5', 'SOS', body)
    worksheet160.write('L5', 'JML', body)
    worksheet160.write('M5', 'MAT', body)
    worksheet160.write('N5', 'SEJ', body)
    worksheet160.write('O5', 'GEO', body)
    worksheet160.write('P5', 'EKO', body)
    worksheet160.write('Q5', 'SOS', body)
    worksheet160.write('R5', 'JML', body)

    worksheet160.conditional_format(5, 0, row160_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet160.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CIGANJUR', title)
    worksheet160.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet160.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet160.write('A22', 'LOKASI', header)
    worksheet160.write('B22', 'TOTAL', header)
    worksheet160.merge_range('A21:B21', 'RANK', header)
    worksheet160.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet160.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet160.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet160.merge_range('F21:F22', 'KELAS', header)
    worksheet160.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet160.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet160.write('G22', 'MAT', body)
    worksheet160.write('H22', 'SEJ', body)
    worksheet160.write('I22', 'GEO', body)
    worksheet160.write('J22', 'EKO', body)
    worksheet160.write('K22', 'SOS', body)
    worksheet160.write('L22', 'JML', body)
    worksheet160.write('M22', 'MAT', body)
    worksheet160.write('N22', 'SEJ', body)
    worksheet160.write('O22', 'GEO', body)
    worksheet160.write('P22', 'EKO', body)
    worksheet160.write('Q22', 'SOS', body)
    worksheet160.write('R22', 'JML', body)

    worksheet160.conditional_format(22, 0, row160+21, 17,
                                    {'type': 'no_errors', 'format': border})

    workbook.close()
    st.success("File siap diunduh!")

    # Tombol unduh file
    with open(file_path, "rb") as f:
        bytes_data = f.read()
    st.download_button(label="Unduh File", data=bytes_data,
                       file_name=file_name)

uploaded_file = st.file_uploader(
    'Letakkan file excel NILAI STANDAR [LOKASI 161-236]', type='xlsx')

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # 89
    r = df.shape[0]-5
    # 90
    s = df.shape[0]-4
    # 91
    t = df.shape[0]-3
    # 92
    u = df.shape[0]-2

    # JUMLAH PESERTA
    peserta = df.iloc[r, 23]

    # rata-rata jumlah benar
    rata_mat = df.iloc[r, 156]
    rata_sej = df.iloc[r, 157]
    rata_geo = df.iloc[r, 158]
    rata_eko = df.iloc[r, 159]
    rata_sos = df.iloc[r, 160]
    rata_jml = df.iloc[r, 161]

    # rata-rata nilai standar
    rata_Smat = df.iloc[t, 167]
    rata_Ssej = df.iloc[t, 168]
    rata_Sgeo = df.iloc[t, 169]
    rata_Seko = df.iloc[t, 170]
    rata_Ssos = df.iloc[t, 171]
    rata_Sjml = df.iloc[t, 172]

    max_mat = df.iloc[t, 156]
    max_sej = df.iloc[t, 157]
    max_geo = df.iloc[t, 158]
    max_eko = df.iloc[t, 159]
    max_sos = df.iloc[t, 160]
    max_jml = df.iloc[t, 161]

    # max nilai standar
    max_Smat = df.iloc[r, 167]
    max_Ssej = df.iloc[r, 168]
    max_Sgeo = df.iloc[r, 169]
    max_Seko = df.iloc[r, 170]
    max_Ssos = df.iloc[r, 171]
    max_Sjml = df.iloc[r, 172]

    # min jumlah benar
    min_mat = df.iloc[u, 156]
    min_sej = df.iloc[u, 157]
    min_geo = df.iloc[u, 158]
    min_eko = df.iloc[u, 159]
    min_sos = df.iloc[u, 160]
    min_jml = df.iloc[u, 161]

    # min nilai standar
    min_Smat = df.iloc[s, 167]
    min_Ssej = df.iloc[s, 168]
    min_Sgeo = df.iloc[s, 169]
    min_Seko = df.iloc[s, 170]
    min_Ssos = df.iloc[s, 171]
    min_Sjml = df.iloc[s, 172]

    data_jml_benar = {'BIDANG STUDI': ['MATEMATIKA (MAT)', 'SEJARAH (SEJ)', 'GEOGRAFI (GEO)', 'EKONOMI (EKO)', 'SOSIOLOGI (SOS)', 'JUMLAH (JML)'],
                      'TERENDAH': [min_mat, min_sej, min_geo, min_eko, min_sos, min_jml],
                      'RATA-RATA': [rata_mat, rata_sej, rata_geo, rata_eko, rata_sos, rata_jml],
                      'TERTINGGI': [max_mat, max_sej, max_geo, max_eko, max_sos, max_jml]}

    jml_benar = pd.DataFrame(data_jml_benar)

    data_n_standar = {'BIDANG STUDI': ['MATEMATIKA (MAT)', 'SEJARAH (SEJ)', 'GEOGRAFI (GEO)', 'EKONOMI (EKO)', 'SOSIOLOGI (SOS)', 'JUMLAH (JML)'],
                      'TERENDAH': [min_Smat, min_Ssej, min_Sgeo, min_Seko, min_Ssos, min_Sjml],
                      'RATA-RATA': [rata_Smat, rata_Ssej, rata_Sgeo, rata_Seko, rata_Ssos, rata_Sjml],
                      'TERTINGGI': [max_Smat, max_Ssej, max_Sgeo, max_Seko, max_Ssos, max_Sjml]}

    n_standar = pd.DataFrame(data_n_standar)

    data_jml_peserta = {'JUMLAH PESERTA': [peserta]}

    jml_peserta = pd.DataFrame(data_jml_peserta)

    data_jml_soal = {'BIDANG STUDI': ['MAT', 'SEJ', 'GEO', 'EKO', 'SOS'],
                     'JUMLAH': [JML_SOAL_MAT, JML_SOAL_SEJ, JML_SOAL_GEO, JML_SOAL_EKO, JML_SOAL_SOS]}

    jml_soal = pd.DataFrame(data_jml_soal)

    df = df[['LOKASI', 'RANK LOK.', 'RANK NAS.', 'NOMOR NF', 'NAMA SISWA', 'NAMA SEKOLAH', 'KELAS',
             'MAT', 'SEJ', 'GEO', 'EKO', 'SOS', 'JML', 'S_MAT', 'S_SEJ', 'S_GEO', 'S_EKO', 'S_SOS', 'S_JML']]

    # sort setiap lokasi
    sort161 = df[df['LOKASI'] == 161]
    sort162 = df[df['LOKASI'] == 162]
    sort163 = df[df['LOKASI'] == 163]
    sort164 = df[df['LOKASI'] == 164]
    sort165 = df[df['LOKASI'] == 165]
    sort166 = df[df['LOKASI'] == 166]
    sort167 = df[df['LOKASI'] == 167]
    sort168 = df[df['LOKASI'] == 168]
    sort169 = df[df['LOKASI'] == 169]
    sort171 = df[df['LOKASI'] == 171]
    sort173 = df[df['LOKASI'] == 173]
    sort174 = df[df['LOKASI'] == 174]
    sort175 = df[df['LOKASI'] == 175]
    sort176 = df[df['LOKASI'] == 176]
    sort177 = df[df['LOKASI'] == 177]
    sort178 = df[df['LOKASI'] == 178]
    sort179 = df[df['LOKASI'] == 179]
    sort180 = df[df['LOKASI'] == 180]
    sort181 = df[df['LOKASI'] == 181]
    sort182 = df[df['LOKASI'] == 182]
    sort183 = df[df['LOKASI'] == 183]
    sort184 = df[df['LOKASI'] == 184]
    sort185 = df[df['LOKASI'] == 185]
    sort186 = df[df['LOKASI'] == 186]
    sort187 = df[df['LOKASI'] == 187]
    sort188 = df[df['LOKASI'] == 188]
    sort189 = df[df['LOKASI'] == 189]
    sort190 = df[df['LOKASI'] == 190]
    sort191 = df[df['LOKASI'] == 191]
    sort192 = df[df['LOKASI'] == 192]
    sort193 = df[df['LOKASI'] == 193]
    sort194 = df[df['LOKASI'] == 194]
    sort195 = df[df['LOKASI'] == 195]
    sort196 = df[df['LOKASI'] == 196]
    sort197 = df[df['LOKASI'] == 197]
    sort198 = df[df['LOKASI'] == 198]
    sort199 = df[df['LOKASI'] == 199]
    sort201 = df[df['LOKASI'] == 201]
    sort202 = df[df['LOKASI'] == 202]
    sort203 = df[df['LOKASI'] == 203]
    sort210 = df[df['LOKASI'] == 210]
    sort211 = df[df['LOKASI'] == 211]
    sort212 = df[df['LOKASI'] == 212]
    sort216 = df[df['LOKASI'] == 216]
    sort217 = df[df['LOKASI'] == 217]
    sort218 = df[df['LOKASI'] == 218]
    sort219 = df[df['LOKASI'] == 219]
    sort220 = df[df['LOKASI'] == 220]
    sort222 = df[df['LOKASI'] == 222]
    sort226 = df[df['LOKASI'] == 226]
    sort227 = df[df['LOKASI'] == 227]
    sort228 = df[df['LOKASI'] == 228]
    sort229 = df[df['LOKASI'] == 229]
    sort230 = df[df['LOKASI'] == 230]
    sort231 = df[df['LOKASI'] == 231]
    sort233 = df[df['LOKASI'] == 233]
    sort234 = df[df['LOKASI'] == 234]
    sort235 = df[df['LOKASI'] == 235]
    sort236 = df[df['LOKASI'] == 236]

    # 10 besar setiap lokasi
    # 161
    sort161_10 = sort161.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort161_10['LOKASI']
    sort161_10 = sort161_10.drop(
        sort161_10[(sort161_10['RANK LOK.'] > 10)].index)
    # 162
    sort162_10 = sort162.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort162_10['LOKASI']
    sort162_10 = sort162_10.drop(
        sort162_10[(sort162_10['RANK LOK.'] > 10)].index)
    # 163
    sort163_10 = sort163.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort163_10['LOKASI']
    sort163_10 = sort163_10.drop(
        sort163_10[(sort163_10['RANK LOK.'] > 10)].index)
    # 164
    sort164_10 = sort164.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort164_10['LOKASI']
    sort164_10 = sort164_10.drop(
        sort164_10[(sort164_10['RANK LOK.'] > 10)].index)
    # 165
    sort165_10 = sort165.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort165_10['LOKASI']
    sort165_10 = sort165_10.drop(
        sort165_10[(sort165_10['RANK LOK.'] > 10)].index)
    # 166
    sort166_10 = sort166.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort166_10['LOKASI']
    sort166_10 = sort166_10.drop(
        sort166_10[(sort166_10['RANK LOK.'] > 10)].index)
    # 167
    sort167_10 = sort167.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort167_10['LOKASI']
    sort167_10 = sort167_10.drop(
        sort167_10[(sort167_10['RANK LOK.'] > 10)].index)
    # 168
    sort168_10 = sort168.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort168_10['LOKASI']
    sort168_10 = sort168_10.drop(
        sort168_10[(sort168_10['RANK LOK.'] > 10)].index)
    # 169
    sort169_10 = sort169.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort169_10['LOKASI']
    sort169_10 = sort169_10.drop(
        sort169_10[(sort169_10['RANK LOK.'] > 10)].index)
    # 171
    sort171_10 = sort171.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort171_10['LOKASI']
    sort171_10 = sort171_10.drop(
        sort171_10[(sort171_10['RANK LOK.'] > 10)].index)
    # 173
    sort173_10 = sort173.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort173_10['LOKASI']
    sort173_10 = sort173_10.drop(
        sort173_10[(sort173_10['RANK LOK.'] > 10)].index)
    # 174
    sort174_10 = sort174.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort174_10['LOKASI']
    sort174_10 = sort174_10.drop(
        sort174_10[(sort174_10['RANK LOK.'] > 10)].index)
    # 175
    sort175_10 = sort175.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort175_10['LOKASI']
    sort175_10 = sort175_10.drop(
        sort175_10[(sort175_10['RANK LOK.'] > 10)].index)
    # 176
    sort176_10 = sort176.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort176_10['LOKASI']
    sort176_10 = sort176_10.drop(
        sort176_10[(sort176_10['RANK LOK.'] > 10)].index)
    # 177
    sort177_10 = sort177.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort177_10['LOKASI']
    sort177_10 = sort177_10.drop(
        sort177_10[(sort177_10['RANK LOK.'] > 10)].index)
    # 178
    sort178_10 = sort178.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort178_10['LOKASI']
    sort178_10 = sort178_10.drop(
        sort178_10[(sort178_10['RANK LOK.'] > 10)].index)
    # 179
    sort179_10 = sort179.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort179_10['LOKASI']
    sort179_10 = sort179_10.drop(
        sort179_10[(sort179_10['RANK LOK.'] > 10)].index)
    # 180
    sort180_10 = sort180.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort180_10['LOKASI']
    sort180_10 = sort180_10.drop(
        sort180_10[(sort180_10['RANK LOK.'] > 10)].index)
    # 181
    sort181_10 = sort181.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort181_10['LOKASI']
    sort181_10 = sort181_10.drop(
        sort181_10[(sort181_10['RANK LOK.'] > 10)].index)
    # 182
    sort182_10 = sort182.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort182_10['LOKASI']
    sort182_10 = sort182_10.drop(
        sort182_10[(sort182_10['RANK LOK.'] > 10)].index)
    # 183
    sort183_10 = sort183.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort183_10['LOKASI']
    sort183_10 = sort183_10.drop(
        sort183_10[(sort183_10['RANK LOK.'] > 10)].index)
    # 184
    sort184_10 = sort184.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort184_10['LOKASI']
    sort184_10 = sort184_10.drop(
        sort184_10[(sort184_10['RANK LOK.'] > 10)].index)
    # 185
    sort185_10 = sort185.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort185_10['LOKASI']
    sort185_10 = sort185_10.drop(
        sort185_10[(sort185_10['RANK LOK.'] > 10)].index)
    # 186
    sort186_10 = sort186.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort186_10['LOKASI']
    sort186_10 = sort186_10.drop(
        sort186_10[(sort186_10['RANK LOK.'] > 10)].index)
    # 187
    sort187_10 = sort187.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort187_10['LOKASI']
    sort187_10 = sort187_10.drop(
        sort187_10[(sort187_10['RANK LOK.'] > 10)].index)
    # 188
    sort188_10 = sort188.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort188_10['LOKASI']
    sort188_10 = sort188_10.drop(
        sort188_10[(sort188_10['RANK LOK.'] > 10)].index)
    # 189
    sort189_10 = sort189.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort189_10['LOKASI']
    sort189_10 = sort189_10.drop(
        sort189_10[(sort189_10['RANK LOK.'] > 10)].index)
    # 190
    sort190_10 = sort190.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort190_10['LOKASI']
    sort190_10 = sort190_10.drop(
        sort190_10[(sort190_10['RANK LOK.'] > 10)].index)
    # 191
    sort191_10 = sort191.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort191_10['LOKASI']
    sort191_10 = sort191_10.drop(
        sort191_10[(sort191_10['RANK LOK.'] > 10)].index)
    # 192
    sort192_10 = sort192.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort192_10['LOKASI']
    sort192_10 = sort192_10.drop(
        sort192_10[(sort192_10['RANK LOK.'] > 10)].index)
    # 193
    sort193_10 = sort193.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort193_10['LOKASI']
    sort193_10 = sort193_10.drop(
        sort193_10[(sort193_10['RANK LOK.'] > 10)].index)
    # 194
    sort194_10 = sort194.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort194_10['LOKASI']
    sort194_10 = sort194_10.drop(
        sort194_10[(sort194_10['RANK LOK.'] > 10)].index)
    # 195
    sort195_10 = sort195.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort195_10['LOKASI']
    sort195_10 = sort195_10.drop(
        sort195_10[(sort195_10['RANK LOK.'] > 10)].index)
    # 196
    sort196_10 = sort196.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort196_10['LOKASI']
    sort196_10 = sort196_10.drop(
        sort196_10[(sort196_10['RANK LOK.'] > 10)].index)
    # 197
    sort197_10 = sort197.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort197_10['LOKASI']
    sort197_10 = sort197_10.drop(
        sort197_10[(sort197_10['RANK LOK.'] > 10)].index)
    # 198
    sort198_10 = sort198.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort198_10['LOKASI']
    sort198_10 = sort198_10.drop(
        sort198_10[(sort198_10['RANK LOK.'] > 10)].index)
    # 199
    sort199_10 = sort199.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort199_10['LOKASI']
    sort199_10 = sort199_10.drop(
        sort199_10[(sort199_10['RANK LOK.'] > 10)].index)
    # 201
    sort201_10 = sort201.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort201_10['LOKASI']
    sort201_10 = sort201_10.drop(
        sort201_10[(sort201_10['RANK LOK.'] > 10)].index)
    # 202
    sort202_10 = sort202.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort202_10['LOKASI']
    sort202_10 = sort202_10.drop(
        sort202_10[(sort202_10['RANK LOK.'] > 10)].index)
    # 203
    sort203_10 = sort203.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort203_10['LOKASI']
    sort203_10 = sort203_10.drop(
        sort203_10[(sort203_10['RANK LOK.'] > 10)].index)
    # 210
    sort210_10 = sort210.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort210_10['LOKASI']
    sort210_10 = sort210_10.drop(
        sort210_10[(sort210_10['RANK LOK.'] > 10)].index)
    # 211
    sort211_10 = sort211.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort211_10['LOKASI']
    sort211_10 = sort211_10.drop(
        sort211_10[(sort211_10['RANK LOK.'] > 10)].index)
    # 212
    sort212_10 = sort212.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort212_10['LOKASI']
    sort212_10 = sort212_10.drop(
        sort212_10[(sort212_10['RANK LOK.'] > 10)].index)
    # 216
    sort216_10 = sort216.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort216_10['LOKASI']
    sort216_10 = sort216_10.drop(
        sort216_10[(sort216_10['RANK LOK.'] > 10)].index)
    # 217
    sort217_10 = sort217.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort217_10['LOKASI']
    sort217_10 = sort217_10.drop(
        sort217_10[(sort217_10['RANK LOK.'] > 10)].index)
    # 218
    sort218_10 = sort218.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort218_10['LOKASI']
    sort218_10 = sort218_10.drop(
        sort218_10[(sort218_10['RANK LOK.'] > 10)].index)
    # 219
    sort219_10 = sort219.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort219_10['LOKASI']
    sort219_10 = sort219_10.drop(
        sort219_10[(sort219_10['RANK LOK.'] > 10)].index)
    # 220
    sort220_10 = sort220.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort220_10['LOKASI']
    sort220_10 = sort220_10.drop(
        sort220_10[(sort220_10['RANK LOK.'] > 10)].index)
    # 222
    sort222_10 = sort222.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort222_10['LOKASI']
    sort222_10 = sort222_10.drop(
        sort222_10[(sort222_10['RANK LOK.'] > 10)].index)
    # 226
    sort226_10 = sort226.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort226_10['LOKASI']
    sort226_10 = sort226_10.drop(
        sort226_10[(sort226_10['RANK LOK.'] > 10)].index)
    # 227
    sort227_10 = sort227.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort227_10['LOKASI']
    sort227_10 = sort227_10.drop(
        sort227_10[(sort227_10['RANK LOK.'] > 10)].index)
    # 228
    sort228_10 = sort228.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort228_10['LOKASI']
    sort228_10 = sort228_10.drop(
        sort228_10[(sort228_10['RANK LOK.'] > 10)].index)
    # 229
    sort229_10 = sort229.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort229_10['LOKASI']
    sort229_10 = sort229_10.drop(
        sort229_10[(sort229_10['RANK LOK.'] > 10)].index)
    # 230
    sort230_10 = sort230.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort230_10['LOKASI']
    sort230_10 = sort230_10.drop(
        sort230_10[(sort230_10['RANK LOK.'] > 10)].index)
    # 231
    sort231_10 = sort231.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort231_10['LOKASI']
    sort231_10 = sort231_10.drop(
        sort231_10[(sort231_10['RANK LOK.'] > 10)].index)
    # 233
    sort233_10 = sort233.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort233_10['LOKASI']
    sort233_10 = sort233_10.drop(
        sort233_10[(sort233_10['RANK LOK.'] > 10)].index)
    # 234
    sort234_10 = sort234.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort234_10['LOKASI']
    sort234_10 = sort234_10.drop(
        sort234_10[(sort234_10['RANK LOK.'] > 10)].index)
    # 235
    sort235_10 = sort235.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort235_10['LOKASI']
    sort235_10 = sort235_10.drop(
        sort235_10[(sort235_10['RANK LOK.'] > 10)].index)
    # 236
    sort236_10 = sort236.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort236_10['LOKASI']
    sort236_10 = sort236_10.drop(
        sort236_10[(sort236_10['RANK LOK.'] > 10)].index)

    # All 161
    sort161 = sort161.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort161['LOKASI']
    # All 162
    sort162 = sort162.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort162['LOKASI']
    # All 163
    sort163 = sort163.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort163['LOKASI']
    # All 164
    sort164 = sort164.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort164['LOKASI']
    # All 165
    sort165 = sort165.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort165['LOKASI']
    # All 166
    sort166 = sort166.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort166['LOKASI']
    # All 167
    sort167 = sort167.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort167['LOKASI']
    # All 168
    sort168 = sort168.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort168['LOKASI']
    # All 169
    sort169 = sort169.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort169['LOKASI']
    # All 171
    sort171 = sort171.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort171['LOKASI']
    # All 173
    sort173 = sort173.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort173['LOKASI']
    # All 174
    sort174 = sort174.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort174['LOKASI']
    # All 175
    sort175 = sort175.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort175['LOKASI']
    # All 176
    sort176 = sort176.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort176['LOKASI']
    # All 177
    sort177 = sort177.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort177['LOKASI']
    # All 178
    sort178 = sort178.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort178['LOKASI']
    # All 179
    sort179 = sort179.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort179['LOKASI']
    # All 180
    sort180 = sort180.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort180['LOKASI']
    # All 181
    sort181 = sort181.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort181['LOKASI']
    # All 182
    sort182 = sort182.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort182['LOKASI']
    # All 183
    sort183 = sort183.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort183['LOKASI']
    # All 184
    sort184 = sort184.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort184['LOKASI']
    # All 185
    sort185 = sort185.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort185['LOKASI']
    # All 186
    sort186 = sort186.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort186['LOKASI']
    # All 187
    sort187 = sort187.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort187['LOKASI']
    # All 188
    sort188 = sort188.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort188['LOKASI']
    # All 189
    sort189 = sort189.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort189['LOKASI']
    # All 190
    sort190 = sort190.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort190['LOKASI']
    # All 191
    sort191 = sort191.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort191['LOKASI']
    # All 192
    sort192 = sort192.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort192['LOKASI']
    # All 193
    sort193 = sort193.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort193['LOKASI']
    # All 194
    sort194 = sort194.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort194['LOKASI']
    # All 195
    sort195 = sort195.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort195['LOKASI']
    # All 196
    sort196 = sort196.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort196['LOKASI']
    # All 197
    sort197 = sort197.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort197['LOKASI']
    # All 198
    sort198 = sort198.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort198['LOKASI']
    # All 199
    sort199 = sort199.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort199['LOKASI']
    # All 201
    sort201 = sort201.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort201['LOKASI']
    # All 202
    sort202 = sort202.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort202['LOKASI']
    # All 203
    sort203 = sort203.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort203['LOKASI']
    # All 210
    sort210 = sort210.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort210['LOKASI']
    # All 211
    sort211 = sort211.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort211['LOKASI']
    # All 212
    sort212 = sort212.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort212['LOKASI']
    # All 216
    sort216 = sort216.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort216['LOKASI']
    # All 217
    sort217 = sort217.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort217['LOKASI']
    # All 218
    sort218 = sort218.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort218['LOKASI']
    # All 219
    sort219 = sort219.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort219['LOKASI']
    # All 220
    sort220 = sort220.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort220['LOKASI']
    # All 222
    sort222 = sort222.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort222['LOKASI']
    # All 226
    sort226 = sort226.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort226['LOKASI']
    # All 227
    sort227 = sort227.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort227['LOKASI']
    # All 228
    sort228 = sort228.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort228['LOKASI']
    # All 229
    sort229 = sort229.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort229['LOKASI']
    # All 230
    sort230 = sort230.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort230['LOKASI']
    # All 231
    sort231 = sort231.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort231['LOKASI']
    # All 233
    sort233 = sort233.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort233['LOKASI']
    # All 234
    sort234 = sort234.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort234['LOKASI']
    # All 235
    sort235 = sort235.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort235['LOKASI']
    # All 236
    sort236 = sort236.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort236['LOKASI']

    # jumlah row
    # 161
    row161_10 = sort161_10.shape[0]
    row161 = sort161.shape[0]
    # 162
    row162_10 = sort162_10.shape[0]
    row162 = sort162.shape[0]
    # 163
    row163_10 = sort163_10.shape[0]
    row163 = sort163.shape[0]
    # 164
    row164_10 = sort164_10.shape[0]
    row164 = sort164.shape[0]
    # 165
    row165_10 = sort165_10.shape[0]
    row165 = sort165.shape[0]
    # 166
    row166_10 = sort166_10.shape[0]
    row166 = sort166.shape[0]
    # 167
    row167_10 = sort167_10.shape[0]
    row167 = sort167.shape[0]
    # 168
    row168_10 = sort168_10.shape[0]
    row168 = sort168.shape[0]
    # 169
    row169_10 = sort169_10.shape[0]
    row169 = sort169.shape[0]
    # 171
    row171_10 = sort171_10.shape[0]
    row171 = sort171.shape[0]
    # 173
    row173_10 = sort173_10.shape[0]
    row173 = sort173.shape[0]
    # 174
    row174_10 = sort174_10.shape[0]
    row174 = sort174.shape[0]
    # 175
    row175_10 = sort175_10.shape[0]
    row175 = sort175.shape[0]
    # 176
    row176_10 = sort176_10.shape[0]
    row176 = sort176.shape[0]
    # 177
    row177_10 = sort177_10.shape[0]
    row177 = sort177.shape[0]
    # 178
    row178_10 = sort178_10.shape[0]
    row178 = sort178.shape[0]
    # 179
    row179_10 = sort179_10.shape[0]
    row179 = sort179.shape[0]
    # 180
    row180_10 = sort180_10.shape[0]
    row180 = sort180.shape[0]
    # 181
    row181_10 = sort181_10.shape[0]
    row181 = sort181.shape[0]
    # 182
    row182_10 = sort182_10.shape[0]
    row182 = sort182.shape[0]
    # 183
    row183_10 = sort183_10.shape[0]
    row183 = sort183.shape[0]
    # 184
    row184_10 = sort184_10.shape[0]
    row184 = sort184.shape[0]
    # 185
    row185_10 = sort185_10.shape[0]
    row185 = sort185.shape[0]
    # 186
    row186_10 = sort186_10.shape[0]
    row186 = sort186.shape[0]
    # 187
    row187_10 = sort187_10.shape[0]
    row187 = sort187.shape[0]
    # 188
    row188_10 = sort188_10.shape[0]
    row188 = sort188.shape[0]
    # 189
    row189_10 = sort189_10.shape[0]
    row189 = sort189.shape[0]
    # 190
    row190_10 = sort190_10.shape[0]
    row190 = sort190.shape[0]
    # 191
    row191_10 = sort191_10.shape[0]
    row191 = sort191.shape[0]
    # 192
    row192_10 = sort192_10.shape[0]
    row192 = sort192.shape[0]
    # 193
    row193_10 = sort193_10.shape[0]
    row193 = sort193.shape[0]
    # 194
    row194_10 = sort194_10.shape[0]
    row194 = sort194.shape[0]
    # 195
    row195_10 = sort195_10.shape[0]
    row195 = sort195.shape[0]
    # 196
    row196_10 = sort196_10.shape[0]
    row196 = sort196.shape[0]
    # 197
    row197_10 = sort197_10.shape[0]
    row197 = sort197.shape[0]
    # 198
    row198_10 = sort198_10.shape[0]
    row198 = sort198.shape[0]
    # 199
    row199_10 = sort199_10.shape[0]
    row199 = sort199.shape[0]
    # 201
    row201_10 = sort201_10.shape[0]
    row201 = sort201.shape[0]
    # 202
    row202_10 = sort202_10.shape[0]
    row202 = sort202.shape[0]
    # 203
    row203_10 = sort203_10.shape[0]
    row203 = sort203.shape[0]
    # 210
    row210_10 = sort210_10.shape[0]
    row210 = sort210.shape[0]
    # 211
    row211_10 = sort211_10.shape[0]
    row211 = sort211.shape[0]
    # 212
    row212_10 = sort212_10.shape[0]
    row212 = sort212.shape[0]
    # 216
    row216_10 = sort216_10.shape[0]
    row216 = sort216.shape[0]
    # 217
    row217_10 = sort217_10.shape[0]
    row217 = sort217.shape[0]
    # 218
    row218_10 = sort218_10.shape[0]
    row218 = sort218.shape[0]
    # 219
    row219_10 = sort219_10.shape[0]
    row219 = sort219.shape[0]
    # 220
    row220_10 = sort220_10.shape[0]
    row220 = sort220.shape[0]
    # 222
    row222_10 = sort222_10.shape[0]
    row222 = sort222.shape[0]
    # 226
    row226_10 = sort226_10.shape[0]
    row226 = sort226.shape[0]
    # 227
    row227_10 = sort227_10.shape[0]
    row227 = sort227.shape[0]
    # 228
    row228_10 = sort228_10.shape[0]
    row228 = sort228.shape[0]
    # 229
    row229_10 = sort229_10.shape[0]
    row229 = sort229.shape[0]
    # 230
    row230_10 = sort230_10.shape[0]
    row230 = sort230.shape[0]
    # 231
    row231_10 = sort231_10.shape[0]
    row231 = sort231.shape[0]
    # 233
    row233_10 = sort233_10.shape[0]
    row233 = sort233.shape[0]
    # 234
    row234_10 = sort234_10.shape[0]
    row234 = sort234.shape[0]
    # 235
    row235_10 = sort235_10.shape[0]
    row235 = sort235.shape[0]
    # 236
    row236_10 = sort236_10.shape[0]
    row236 = sort236.shape[0]

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    # Path file hasil penyimpanan
    file_name = f"{kelas}_{penilaian}_{semester}_lokasi_161_236.xlsx"
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
                       startrow=21,
                       startcol=0,
                       index=False,
                       header=False)

    # Convert the dataframe to an XlsxWriter Excel object cover.
    jml_peserta.to_excel(writer, sheet_name='cover',
                         startrow=21,
                         startcol=5,
                         index=False,
                         header=False)

    # Convert the dataframe to an XlsxWriter Excel object cover.
    jml_soal.to_excel(writer, sheet_name='cover',
                      startrow=13,
                      startcol=5,
                      index=False,
                      header=False)

    # 161
    # Convert the dataframe to an XlsxWriter Excel object.
    sort161_10.to_excel(writer, sheet_name='161',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort161.to_excel(writer, sheet_name='161',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 162
    # Convert the dataframe to an XlsxWriter Excel object.
    sort162_10.to_excel(writer, sheet_name='162',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort162.to_excel(writer, sheet_name='162',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 163
    # Convert the dataframe to an XlsxWriter Excel object.
    sort163_10.to_excel(writer, sheet_name='163',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort163.to_excel(writer, sheet_name='163',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 164
    # Convert the dataframe to an XlsxWriter Excel object.
    sort164_10.to_excel(writer, sheet_name='164',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort164.to_excel(writer, sheet_name='164',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 165
    # Convert the dataframe to an XlsxWriter Excel object.
    sort165_10.to_excel(writer, sheet_name='165',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort165.to_excel(writer, sheet_name='165',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 166
    # Convert the dataframe to an XlsxWriter Excel object.
    sort166_10.to_excel(writer, sheet_name='166',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort166.to_excel(writer, sheet_name='166',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 167
    # Convert the dataframe to an XlsxWriter Excel object.
    sort167_10.to_excel(writer, sheet_name='167',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort167.to_excel(writer, sheet_name='167',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 168
    # Convert the dataframe to an XlsxWriter Excel object.
    sort168_10.to_excel(writer, sheet_name='168',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort168.to_excel(writer, sheet_name='168',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 169
    # Convert the dataframe to an XlsxWriter Excel object.
    sort169_10.to_excel(writer, sheet_name='169',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort169.to_excel(writer, sheet_name='169',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 171
    # Convert the dataframe to an XlsxWriter Excel object.
    sort171_10.to_excel(writer, sheet_name='171',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort171.to_excel(writer, sheet_name='171',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 173
    # Convert the dataframe to an XlsxWriter Excel object.
    sort173_10.to_excel(writer, sheet_name='173',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort173.to_excel(writer, sheet_name='173',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 174
    # Convert the dataframe to an XlsxWriter Excel object.
    sort174_10.to_excel(writer, sheet_name='174',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort174.to_excel(writer, sheet_name='174',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 175
    # Convert the dataframe to an XlsxWriter Excel object.
    sort175_10.to_excel(writer, sheet_name='175',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort175.to_excel(writer, sheet_name='175',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 176
    # Convert the dataframe to an XlsxWriter Excel object.
    sort176_10.to_excel(writer, sheet_name='176',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort176.to_excel(writer, sheet_name='176',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 177
    # Convert the dataframe to an XlsxWriter Excel object.
    sort177_10.to_excel(writer, sheet_name='177',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort177.to_excel(writer, sheet_name='177',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 178
    # Convert the dataframe to an XlsxWriter Excel object.
    sort178_10.to_excel(writer, sheet_name='178',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort178.to_excel(writer, sheet_name='178',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 179
    # Convert the dataframe to an XlsxWriter Excel object.
    sort179_10.to_excel(writer, sheet_name='179',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort179.to_excel(writer, sheet_name='179',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 180
    # Convert the dataframe to an XlsxWriter Excel object.
    sort180_10.to_excel(writer, sheet_name='180',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort180.to_excel(writer, sheet_name='180',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 181
    # Convert the dataframe to an XlsxWriter Excel object.
    sort181_10.to_excel(writer, sheet_name='181',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort181.to_excel(writer, sheet_name='181',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 182
    # Convert the dataframe to an XlsxWriter Excel object.
    sort182_10.to_excel(writer, sheet_name='182',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort182.to_excel(writer, sheet_name='182',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 183
    # Convert the dataframe to an XlsxWriter Excel object.
    sort183_10.to_excel(writer, sheet_name='183',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort183.to_excel(writer, sheet_name='183',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 184
    # Convert the dataframe to an XlsxWriter Excel object.
    sort184_10.to_excel(writer, sheet_name='184',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort184.to_excel(writer, sheet_name='184',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 185
    # Convert the dataframe to an XlsxWriter Excel object.
    sort185_10.to_excel(writer, sheet_name='185',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort185.to_excel(writer, sheet_name='185',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 186
    # Convert the dataframe to an XlsxWriter Excel object.
    sort186_10.to_excel(writer, sheet_name='186',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort186.to_excel(writer, sheet_name='186',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 187
    # Convert the dataframe to an XlsxWriter Excel object.
    sort187_10.to_excel(writer, sheet_name='187',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort187.to_excel(writer, sheet_name='187',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 188
    # Convert the dataframe to an XlsxWriter Excel object.
    sort188_10.to_excel(writer, sheet_name='188',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort188.to_excel(writer, sheet_name='188',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 189
    # Convert the dataframe to an XlsxWriter Excel object.
    sort189_10.to_excel(writer, sheet_name='189',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort189.to_excel(writer, sheet_name='189',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 190
    # Convert the dataframe to an XlsxWriter Excel object.
    sort190_10.to_excel(writer, sheet_name='190',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort190.to_excel(writer, sheet_name='190',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 191
    # Convert the dataframe to an XlsxWriter Excel object.
    sort191_10.to_excel(writer, sheet_name='191',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort191.to_excel(writer, sheet_name='191',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 192
    # Convert the dataframe to an XlsxWriter Excel object.
    sort192_10.to_excel(writer, sheet_name='192',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort192.to_excel(writer, sheet_name='192',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 193
    # Convert the dataframe to an XlsxWriter Excel object.
    sort193_10.to_excel(writer, sheet_name='193',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort193.to_excel(writer, sheet_name='193',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 194
    # Convert the dataframe to an XlsxWriter Excel object.
    sort194_10.to_excel(writer, sheet_name='194',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort194.to_excel(writer, sheet_name='194',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 195
    # Convert the dataframe to an XlsxWriter Excel object.
    sort195_10.to_excel(writer, sheet_name='195',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort195.to_excel(writer, sheet_name='195',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 196
    # Convert the dataframe to an XlsxWriter Excel object.
    sort196_10.to_excel(writer, sheet_name='196',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort196.to_excel(writer, sheet_name='196',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 197
    # Convert the dataframe to an XlsxWriter Excel object.
    sort197_10.to_excel(writer, sheet_name='197',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort197.to_excel(writer, sheet_name='197',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 198
    # Convert the dataframe to an XlsxWriter Excel object.
    sort198_10.to_excel(writer, sheet_name='198',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort198.to_excel(writer, sheet_name='198',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 199
    # Convert the dataframe to an XlsxWriter Excel object.
    sort199_10.to_excel(writer, sheet_name='199',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort199.to_excel(writer, sheet_name='199',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 201
    # Convert the dataframe to an XlsxWriter Excel object.
    sort201_10.to_excel(writer, sheet_name='201',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort201.to_excel(writer, sheet_name='201',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 202
    # Convert the dataframe to an XlsxWriter Excel object.
    sort202_10.to_excel(writer, sheet_name='202',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort202.to_excel(writer, sheet_name='202',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 203
    # Convert the dataframe to an XlsxWriter Excel object.
    sort203_10.to_excel(writer, sheet_name='203',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort203.to_excel(writer, sheet_name='203',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 210
    # Convert the dataframe to an XlsxWriter Excel object.
    sort210_10.to_excel(writer, sheet_name='210',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort210.to_excel(writer, sheet_name='210',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 211
    # Convert the dataframe to an XlsxWriter Excel object.
    sort211_10.to_excel(writer, sheet_name='211',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort211.to_excel(writer, sheet_name='211',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 212
    # Convert the dataframe to an XlsxWriter Excel object.
    sort212_10.to_excel(writer, sheet_name='212',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort212.to_excel(writer, sheet_name='212',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 216
    # Convert the dataframe to an XlsxWriter Excel object.
    sort216_10.to_excel(writer, sheet_name='216',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort216.to_excel(writer, sheet_name='216',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 217
    # Convert the dataframe to an XlsxWriter Excel object.
    sort217_10.to_excel(writer, sheet_name='217',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort217.to_excel(writer, sheet_name='217',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 218
    # Convert the dataframe to an XlsxWriter Excel object.
    sort218_10.to_excel(writer, sheet_name='218',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort218.to_excel(writer, sheet_name='218',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 219
    # Convert the dataframe to an XlsxWriter Excel object.
    sort219_10.to_excel(writer, sheet_name='219',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort219.to_excel(writer, sheet_name='219',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 220
    # Convert the dataframe to an XlsxWriter Excel object.
    sort220_10.to_excel(writer, sheet_name='220',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort220.to_excel(writer, sheet_name='220',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 222
    # Convert the dataframe to an XlsxWriter Excel object.
    sort222_10.to_excel(writer, sheet_name='222',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort222.to_excel(writer, sheet_name='222',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 226
    # Convert the dataframe to an XlsxWriter Excel object.
    sort226_10.to_excel(writer, sheet_name='226',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort226.to_excel(writer, sheet_name='226',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 227
    # Convert the dataframe to an XlsxWriter Excel object.
    sort227_10.to_excel(writer, sheet_name='227',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort227.to_excel(writer, sheet_name='227',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 228
    # Convert the dataframe to an XlsxWriter Excel object.
    sort228_10.to_excel(writer, sheet_name='228',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort228.to_excel(writer, sheet_name='228',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 229
    # Convert the dataframe to an XlsxWriter Excel object.
    sort229_10.to_excel(writer, sheet_name='229',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort229.to_excel(writer, sheet_name='229',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 230
    # Convert the dataframe to an XlsxWriter Excel object.
    sort230_10.to_excel(writer, sheet_name='230',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort230.to_excel(writer, sheet_name='230',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 231
    # Convert the dataframe to an XlsxWriter Excel object.
    sort231_10.to_excel(writer, sheet_name='231',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort231.to_excel(writer, sheet_name='231',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 233
    # Convert the dataframe to an XlsxWriter Excel object.
    sort233_10.to_excel(writer, sheet_name='233',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort233.to_excel(writer, sheet_name='233',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 234
    # Convert the dataframe to an XlsxWriter Excel object.
    sort234_10.to_excel(writer, sheet_name='234',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort234.to_excel(writer, sheet_name='234',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 235
    # Convert the dataframe to an XlsxWriter Excel object.
    sort235_10.to_excel(writer, sheet_name='235',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort235.to_excel(writer, sheet_name='235',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 236
    # Convert the dataframe to an XlsxWriter Excel object.
    sort236_10.to_excel(writer, sheet_name='236',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort236.to_excel(writer, sheet_name='236',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)

    # Get the xlsxwriter objects from the dataframe writer object.
    workbook = writer.book

    # membuat worksheet baru
    worksheetcover = writer.sheets['cover']
    worksheet161 = writer.sheets['161']
    worksheet162 = writer.sheets['162']
    worksheet163 = writer.sheets['163']
    worksheet164 = writer.sheets['164']
    worksheet165 = writer.sheets['165']
    worksheet166 = writer.sheets['166']
    worksheet167 = writer.sheets['167']
    worksheet168 = writer.sheets['168']
    worksheet169 = writer.sheets['169']
    worksheet171 = writer.sheets['171']
    worksheet173 = writer.sheets['173']
    worksheet174 = writer.sheets['174']
    worksheet175 = writer.sheets['175']
    worksheet176 = writer.sheets['176']
    worksheet177 = writer.sheets['177']
    worksheet178 = writer.sheets['178']
    worksheet179 = writer.sheets['179']
    worksheet180 = writer.sheets['180']
    worksheet181 = writer.sheets['181']
    worksheet182 = writer.sheets['182']
    worksheet183 = writer.sheets['183']
    worksheet184 = writer.sheets['184']
    worksheet185 = writer.sheets['185']
    worksheet186 = writer.sheets['186']
    worksheet187 = writer.sheets['187']
    worksheet188 = writer.sheets['188']
    worksheet189 = writer.sheets['189']
    worksheet190 = writer.sheets['190']
    worksheet191 = writer.sheets['191']
    worksheet192 = writer.sheets['192']
    worksheet193 = writer.sheets['193']
    worksheet194 = writer.sheets['194']
    worksheet195 = writer.sheets['195']
    worksheet196 = writer.sheets['196']
    worksheet197 = writer.sheets['197']
    worksheet198 = writer.sheets['198']
    worksheet199 = writer.sheets['199']
    worksheet201 = writer.sheets['201']
    worksheet202 = writer.sheets['202']
    worksheet203 = writer.sheets['203']
    worksheet210 = writer.sheets['210']
    worksheet211 = writer.sheets['211']
    worksheet212 = writer.sheets['212']
    worksheet216 = writer.sheets['216']
    worksheet217 = writer.sheets['217']
    worksheet218 = writer.sheets['218']
    worksheet219 = writer.sheets['219']
    worksheet220 = writer.sheets['220']
    worksheet222 = writer.sheets['222']
    worksheet226 = writer.sheets['226']
    worksheet227 = writer.sheets['227']
    worksheet228 = writer.sheets['228']
    worksheet229 = writer.sheets['229']
    worksheet230 = writer.sheets['230']
    worksheet231 = writer.sheets['231']
    worksheet233 = writer.sheets['233']
    worksheet234 = writer.sheets['234']
    worksheet235 = writer.sheets['235']
    worksheet236 = writer.sheets['236']

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
        'font_size': 35,
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
    worksheetcover.conditional_format(16, 0, 11, 3,
                                      {'type': 'no_errors', 'format': borderCover})

    worksheetcover.insert_image('F1', r'logo nf.jpg')

    worksheetcover.merge_range('A10:A11', 'BIDANG STUDI', bodyCover)
    worksheetcover.merge_range('B10:B11', 'TERENDAH', bodyCover)
    worksheetcover.merge_range('C10:C11', 'RATA-RATA', bodyCover)
    worksheetcover.merge_range('D10:D11', 'TERTINGGI', bodyCover)
    worksheetcover.merge_range('A20:A21', 'BIDANG STUDI', bodyCover)
    worksheetcover.merge_range('B20:B21', 'TERENDAH', bodyCover)
    worksheetcover.merge_range('C20:C21', 'RATA-RATA', bodyCover)
    worksheetcover.merge_range('D20:D21', 'TERTINGGI', bodyCover)
    worksheetcover.write('F13', 'BIDANG STUDI', bodyCover)
    worksheetcover.merge_range('F20:F21', 'JUMLAH', sub_header1Cover)
    worksheetcover.merge_range('F23:F24', 'PESERTA', sub_header1Cover)
    worksheetcover.write('G13', 'JUMLAH', bodyCover)
    worksheetcover.set_column('A:A', 25.71, centerCover)
    worksheetcover.set_column('B:B', 15, centerCover)
    worksheetcover.set_column('C:C', 15, centerCover)
    worksheetcover.set_column('D:D', 15, centerCover)
    worksheetcover.set_column('F:F', 25.71, centerCover)
    worksheetcover.set_column('G:G', 13, centerCover)
    worksheetcover.merge_range('A1:F3', 'DAFTAR NILAI', titleCover)
    worksheetcover.merge_range(
        'A4:F5', fr'{penilaian}', sub_titleCover)
    worksheetcover.merge_range(
        'A6:F7', fr'{semester} TAHUN {tahun}', headerCover)
    worksheetcover.write('A9', 'JUMLAH BENAR', sub_headerCover)
    worksheetcover.write('A19', 'NILAI STANDAR', sub_headerCover)
    worksheetcover.merge_range('F8:G9', fr'{kelas}', kelasCover)
    worksheetcover.merge_range('F11:G12', 'JUMLAH SOAL', sub_header1Cover)

    worksheetcover.conditional_format(26, 0, 21, 3,
                                      {'type': 'no_errors', 'format': borderCover})

    worksheetcover.conditional_format(17, 6, 13, 5,
                                      {'type': 'no_errors', 'format': borderCover})

    worksheetcover.conditional_format(21, 5, 21, 5,
                                      {'type': 'no_errors', 'format': borderCover})

    # worksheet 161
    worksheet161.insert_image('A1', r'logo resmi nf.jpg')

    worksheet161.set_column('A:A', 7, center)
    worksheet161.set_column('B:B', 6, center)
    worksheet161.set_column('C:C', 18.14, center)
    worksheet161.set_column('D:D', 25, left)
    worksheet161.set_column('E:E', 13.14, left)
    worksheet161.set_column('F:F', 8.57, center)
    worksheet161.set_column('G:R', 5, center)
    worksheet161.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BENHIL', title)
    worksheet161.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet161.write('A5', 'LOKASI', header)
    worksheet161.write('B5', 'TOTAL', header)
    worksheet161.merge_range('A4:B4', 'RANK', header)
    worksheet161.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet161.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet161.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet161.merge_range('F4:F5', 'KELAS', header)
    worksheet161.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet161.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet161.write('G5', 'MAT', body)
    worksheet161.write('H5', 'SEJ', body)
    worksheet161.write('I5', 'GEO', body)
    worksheet161.write('J5', 'EKO', body)
    worksheet161.write('K5', 'SOS', body)
    worksheet161.write('L5', 'JML', body)
    worksheet161.write('M5', 'MAT', body)
    worksheet161.write('N5', 'SEJ', body)
    worksheet161.write('O5', 'GEO', body)
    worksheet161.write('P5', 'EKO', body)
    worksheet161.write('Q5', 'SOS', body)
    worksheet161.write('R5', 'JML', body)

    worksheet161.conditional_format(5, 0, row161_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet161.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF BENHIL', title)
    worksheet161.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet161.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet161.write('A22', 'LOKASI', header)
    worksheet161.write('B22', 'TOTAL', header)
    worksheet161.merge_range('A21:B21', 'RANK', header)
    worksheet161.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet161.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet161.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet161.merge_range('F21:F22', 'KELAS', header)
    worksheet161.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet161.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet161.write('G22', 'MAT', body)
    worksheet161.write('H22', 'SEJ', body)
    worksheet161.write('I22', 'GEO', body)
    worksheet161.write('J22', 'EKO', body)
    worksheet161.write('K22', 'SOS', body)
    worksheet161.write('L22', 'JML', body)
    worksheet161.write('M22', 'MAT', body)
    worksheet161.write('N22', 'SEJ', body)
    worksheet161.write('O22', 'GEO', body)
    worksheet161.write('P22', 'EKO', body)
    worksheet161.write('Q22', 'SOS', body)
    worksheet161.write('R22', 'JML', body)

    worksheet161.conditional_format(22, 0, row161+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 162
    worksheet162.insert_image('A1', r'logo resmi nf.jpg')

    worksheet162.set_column('A:A', 7, center)
    worksheet162.set_column('B:B', 6, center)
    worksheet162.set_column('C:C', 18.14, center)
    worksheet162.set_column('D:D', 25, left)
    worksheet162.set_column('E:E', 13.14, left)
    worksheet162.set_column('F:F', 8.57, center)
    worksheet162.set_column('G:R', 5, center)
    worksheet162.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PUNTI KAYU', title)
    worksheet162.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet162.write('A5', 'LOKASI', header)
    worksheet162.write('B5', 'TOTAL', header)
    worksheet162.merge_range('A4:B4', 'RANK', header)
    worksheet162.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet162.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet162.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet162.merge_range('F4:F5', 'KELAS', header)
    worksheet162.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet162.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet162.write('G5', 'MAT', body)
    worksheet162.write('H5', 'SEJ', body)
    worksheet162.write('I5', 'GEO', body)
    worksheet162.write('J5', 'EKO', body)
    worksheet162.write('K5', 'SOS', body)
    worksheet162.write('L5', 'JML', body)
    worksheet162.write('M5', 'MAT', body)
    worksheet162.write('N5', 'SEJ', body)
    worksheet162.write('O5', 'GEO', body)
    worksheet162.write('P5', 'EKO', body)
    worksheet162.write('Q5', 'SOS', body)
    worksheet162.write('R5', 'JML', body)

    worksheet162.conditional_format(5, 0, row162_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet162.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PUNTI KAYU', title)
    worksheet162.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet162.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet162.write('A22', 'LOKASI', header)
    worksheet162.write('B22', 'TOTAL', header)
    worksheet162.merge_range('A21:B21', 'RANK', header)
    worksheet162.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet162.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet162.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet162.merge_range('F21:F22', 'KELAS', header)
    worksheet162.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet162.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet162.write('G22', 'MAT', body)
    worksheet162.write('H22', 'SEJ', body)
    worksheet162.write('I22', 'GEO', body)
    worksheet162.write('J22', 'EKO', body)
    worksheet162.write('K22', 'SOS', body)
    worksheet162.write('L22', 'JML', body)
    worksheet162.write('M22', 'MAT', body)
    worksheet162.write('N22', 'SEJ', body)
    worksheet162.write('O22', 'GEO', body)
    worksheet162.write('P22', 'EKO', body)
    worksheet162.write('Q22', 'SOS', body)
    worksheet162.write('R22', 'JML', body)

    worksheet162.conditional_format(22, 0, row162+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 163
    worksheet163.insert_image('A1', r'logo resmi nf.jpg')

    worksheet163.set_column('A:A', 7, center)
    worksheet163.set_column('B:B', 6, center)
    worksheet163.set_column('C:C', 18.14, center)
    worksheet163.set_column('D:D', 25, left)
    worksheet163.set_column('E:E', 13.14, left)
    worksheet163.set_column('F:F', 8.57, center)
    worksheet163.set_column('G:R', 5, center)
    worksheet163.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SUKAMTO', title)
    worksheet163.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet163.write('A5', 'LOKASI', header)
    worksheet163.write('B5', 'TOTAL', header)
    worksheet163.merge_range('A4:B4', 'RANK', header)
    worksheet163.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet163.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet163.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet163.merge_range('F4:F5', 'KELAS', header)
    worksheet163.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet163.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet163.write('G5', 'MAT', body)
    worksheet163.write('H5', 'SEJ', body)
    worksheet163.write('I5', 'GEO', body)
    worksheet163.write('J5', 'EKO', body)
    worksheet163.write('K5', 'SOS', body)
    worksheet163.write('L5', 'JML', body)
    worksheet163.write('M5', 'MAT', body)
    worksheet163.write('N5', 'SEJ', body)
    worksheet163.write('O5', 'GEO', body)
    worksheet163.write('P5', 'EKO', body)
    worksheet163.write('Q5', 'SOS', body)
    worksheet163.write('R5', 'JML', body)

    worksheet163.conditional_format(5, 0, row163_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet163.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF SUKAMTO', title)
    worksheet163.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet163.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet163.write('A22', 'LOKASI', header)
    worksheet163.write('B22', 'TOTAL', header)
    worksheet163.merge_range('A21:B21', 'RANK', header)
    worksheet163.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet163.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet163.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet163.merge_range('F21:F22', 'KELAS', header)
    worksheet163.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet163.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet163.write('G22', 'MAT', body)
    worksheet163.write('H22', 'SEJ', body)
    worksheet163.write('I22', 'GEO', body)
    worksheet163.write('J22', 'EKO', body)
    worksheet163.write('K22', 'SOS', body)
    worksheet163.write('L22', 'JML', body)
    worksheet163.write('M22', 'MAT', body)
    worksheet163.write('N22', 'SEJ', body)
    worksheet163.write('O22', 'GEO', body)
    worksheet163.write('P22', 'EKO', body)
    worksheet163.write('Q22', 'SOS', body)
    worksheet163.write('R22', 'JML', body)

    worksheet163.conditional_format(22, 0, row163+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 164
    worksheet164.insert_image('A1', r'logo resmi nf.jpg')

    worksheet164.set_column('A:A', 7, center)
    worksheet164.set_column('B:B', 6, center)
    worksheet164.set_column('C:C', 18.14, center)
    worksheet164.set_column('D:D', 25, left)
    worksheet164.set_column('E:E', 13.14, left)
    worksheet164.set_column('F:F', 8.57, center)
    worksheet164.set_column('G:R', 5, center)
    worksheet164.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BUKIT BESAR', title)
    worksheet164.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet164.write('A5', 'LOKASI', header)
    worksheet164.write('B5', 'TOTAL', header)
    worksheet164.merge_range('A4:B4', 'RANK', header)
    worksheet164.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet164.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet164.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet164.merge_range('F4:F5', 'KELAS', header)
    worksheet164.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet164.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet164.write('G5', 'MAT', body)
    worksheet164.write('H5', 'SEJ', body)
    worksheet164.write('I5', 'GEO', body)
    worksheet164.write('J5', 'EKO', body)
    worksheet164.write('K5', 'SOS', body)
    worksheet164.write('L5', 'JML', body)
    worksheet164.write('M5', 'MAT', body)
    worksheet164.write('N5', 'SEJ', body)
    worksheet164.write('O5', 'GEO', body)
    worksheet164.write('P5', 'EKO', body)
    worksheet164.write('Q5', 'SOS', body)
    worksheet164.write('R5', 'JML', body)

    worksheet164.conditional_format(5, 0, row164_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet164.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF BUKIT BESAR', title)
    worksheet164.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet164.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet164.write('A22', 'LOKASI', header)
    worksheet164.write('B22', 'TOTAL', header)
    worksheet164.merge_range('A21:B21', 'RANK', header)
    worksheet164.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet164.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet164.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet164.merge_range('F21:F22', 'KELAS', header)
    worksheet164.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet164.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet164.write('G22', 'MAT', body)
    worksheet164.write('H22', 'SEJ', body)
    worksheet164.write('I22', 'GEO', body)
    worksheet164.write('J22', 'EKO', body)
    worksheet164.write('K22', 'SOS', body)
    worksheet164.write('L22', 'JML', body)
    worksheet164.write('M22', 'MAT', body)
    worksheet164.write('N22', 'SEJ', body)
    worksheet164.write('O22', 'GEO', body)
    worksheet164.write('P22', 'EKO', body)
    worksheet164.write('Q22', 'SOS', body)
    worksheet164.write('R22', 'JML', body)

    worksheet164.conditional_format(22, 0, row164+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 165
    worksheet165.insert_image('A1', r'logo resmi nf.jpg')

    worksheet165.set_column('A:A', 7, center)
    worksheet165.set_column('B:B', 6, center)
    worksheet165.set_column('C:C', 18.14, center)
    worksheet165.set_column('D:D', 25, left)
    worksheet165.set_column('E:E', 13.14, left)
    worksheet165.set_column('F:F', 8.57, center)
    worksheet165.set_column('G:R', 5, center)
    worksheet165.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SUDIRMAN', title)
    worksheet165.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet165.write('A5', 'LOKASI', header)
    worksheet165.write('B5', 'TOTAL', header)
    worksheet165.merge_range('A4:B4', 'RANK', header)
    worksheet165.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet165.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet165.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet165.merge_range('F4:F5', 'KELAS', header)
    worksheet165.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet165.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet165.write('G5', 'MAT', body)
    worksheet165.write('H5', 'SEJ', body)
    worksheet165.write('I5', 'GEO', body)
    worksheet165.write('J5', 'EKO', body)
    worksheet165.write('K5', 'SOS', body)
    worksheet165.write('L5', 'JML', body)
    worksheet165.write('M5', 'MAT', body)
    worksheet165.write('N5', 'SEJ', body)
    worksheet165.write('O5', 'GEO', body)
    worksheet165.write('P5', 'EKO', body)
    worksheet165.write('Q5', 'SOS', body)
    worksheet165.write('R5', 'JML', body)

    worksheet165.conditional_format(5, 0, row165_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet165.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF SUDIRMAN', title)
    worksheet165.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet165.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet165.write('A22', 'LOKASI', header)
    worksheet165.write('B22', 'TOTAL', header)
    worksheet165.merge_range('A21:B21', 'RANK', header)
    worksheet165.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet165.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet165.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet165.merge_range('F21:F22', 'KELAS', header)
    worksheet165.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet165.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet165.write('G22', 'MAT', body)
    worksheet165.write('H22', 'SEJ', body)
    worksheet165.write('I22', 'GEO', body)
    worksheet165.write('J22', 'EKO', body)
    worksheet165.write('K22', 'SOS', body)
    worksheet165.write('L22', 'JML', body)
    worksheet165.write('M22', 'MAT', body)
    worksheet165.write('N22', 'SEJ', body)
    worksheet165.write('O22', 'GEO', body)
    worksheet165.write('P22', 'EKO', body)
    worksheet165.write('Q22', 'SOS', body)
    worksheet165.write('R22', 'JML', body)

    worksheet165.conditional_format(22, 0, row165+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 166
    worksheet166.insert_image('A1', r'logo resmi nf.jpg')

    worksheet166.set_column('A:A', 7, center)
    worksheet166.set_column('B:B', 6, center)
    worksheet166.set_column('C:C', 18.14, center)
    worksheet166.set_column('D:D', 25, left)
    worksheet166.set_column('E:E', 13.14, left)
    worksheet166.set_column('F:F', 8.57, center)
    worksheet166.set_column('G:R', 5, center)
    worksheet166.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIMANGGU', title)
    worksheet166.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet166.write('A5', 'LOKASI', header)
    worksheet166.write('B5', 'TOTAL', header)
    worksheet166.merge_range('A4:B4', 'RANK', header)
    worksheet166.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet166.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet166.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet166.merge_range('F4:F5', 'KELAS', header)
    worksheet166.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet166.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet166.write('G5', 'MAT', body)
    worksheet166.write('H5', 'SEJ', body)
    worksheet166.write('I5', 'GEO', body)
    worksheet166.write('J5', 'EKO', body)
    worksheet166.write('K5', 'SOS', body)
    worksheet166.write('L5', 'JML', body)
    worksheet166.write('M5', 'MAT', body)
    worksheet166.write('N5', 'SEJ', body)
    worksheet166.write('O5', 'GEO', body)
    worksheet166.write('P5', 'EKO', body)
    worksheet166.write('Q5', 'SOS', body)
    worksheet166.write('R5', 'JML', body)

    worksheet166.conditional_format(5, 0, row166_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet166.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CIMANGGU', title)
    worksheet166.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet166.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet166.write('A22', 'LOKASI', header)
    worksheet166.write('B22', 'TOTAL', header)
    worksheet166.merge_range('A21:B21', 'RANK', header)
    worksheet166.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet166.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet166.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet166.merge_range('F21:F22', 'KELAS', header)
    worksheet166.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet166.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet166.write('G22', 'MAT', body)
    worksheet166.write('H22', 'SEJ', body)
    worksheet166.write('I22', 'GEO', body)
    worksheet166.write('J22', 'EKO', body)
    worksheet166.write('K22', 'SOS', body)
    worksheet166.write('L22', 'JML', body)
    worksheet166.write('M22', 'MAT', body)
    worksheet166.write('N22', 'SEJ', body)
    worksheet166.write('O22', 'GEO', body)
    worksheet166.write('P22', 'EKO', body)
    worksheet166.write('Q22', 'SOS', body)
    worksheet166.write('R22', 'JML', body)

    worksheet166.conditional_format(22, 0, row166+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 167
    worksheet167.insert_image('A1', r'logo resmi nf.jpg')

    worksheet167.set_column('A:A', 7, center)
    worksheet167.set_column('B:B', 6, center)
    worksheet167.set_column('C:C', 18.14, center)
    worksheet167.set_column('D:D', 25, left)
    worksheet167.set_column('E:E', 13.14, left)
    worksheet167.set_column('F:F', 8.57, center)
    worksheet167.set_column('G:R', 5, center)
    worksheet167.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CILANGKAP', title)
    worksheet167.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet167.write('A5', 'LOKASI', header)
    worksheet167.write('B5', 'TOTAL', header)
    worksheet167.merge_range('A4:B4', 'RANK', header)
    worksheet167.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet167.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet167.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet167.merge_range('F4:F5', 'KELAS', header)
    worksheet167.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet167.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet167.write('G5', 'MAT', body)
    worksheet167.write('H5', 'SEJ', body)
    worksheet167.write('I5', 'GEO', body)
    worksheet167.write('J5', 'EKO', body)
    worksheet167.write('K5', 'SOS', body)
    worksheet167.write('L5', 'JML', body)
    worksheet167.write('M5', 'MAT', body)
    worksheet167.write('N5', 'SEJ', body)
    worksheet167.write('O5', 'GEO', body)
    worksheet167.write('P5', 'EKO', body)
    worksheet167.write('Q5', 'SOS', body)
    worksheet167.write('R5', 'JML', body)

    worksheet167.conditional_format(5, 0, row167_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet167.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CILANGKAP', title)
    worksheet167.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet167.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet167.write('A22', 'LOKASI', header)
    worksheet167.write('B22', 'TOTAL', header)
    worksheet167.merge_range('A21:B21', 'RANK', header)
    worksheet167.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet167.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet167.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet167.merge_range('F21:F22', 'KELAS', header)
    worksheet167.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet167.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet167.write('G22', 'MAT', body)
    worksheet167.write('H22', 'SEJ', body)
    worksheet167.write('I22', 'GEO', body)
    worksheet167.write('J22', 'EKO', body)
    worksheet167.write('K22', 'SOS', body)
    worksheet167.write('L22', 'JML', body)
    worksheet167.write('M22', 'MAT', body)
    worksheet167.write('N22', 'SEJ', body)
    worksheet167.write('O22', 'GEO', body)
    worksheet167.write('P22', 'EKO', body)
    worksheet167.write('Q22', 'SOS', body)
    worksheet167.write('R22', 'JML', body)

    worksheet167.conditional_format(22, 0, row167+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 168
    worksheet168.insert_image('A1', r'logo resmi nf.jpg')

    worksheet168.set_column('A:A', 7, center)
    worksheet168.set_column('B:B', 6, center)
    worksheet168.set_column('C:C', 18.14, center)
    worksheet168.set_column('D:D', 25, left)
    worksheet168.set_column('E:E', 13.14, left)
    worksheet168.set_column('F:F', 8.57, center)
    worksheet168.set_column('G:R', 5, center)
    worksheet168.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF HALIM', title)
    worksheet168.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet168.write('A5', 'LOKASI', header)
    worksheet168.write('B5', 'TOTAL', header)
    worksheet168.merge_range('A4:B4', 'RANK', header)
    worksheet168.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet168.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet168.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet168.merge_range('F4:F5', 'KELAS', header)
    worksheet168.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet168.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet168.write('G5', 'MAT', body)
    worksheet168.write('H5', 'SEJ', body)
    worksheet168.write('I5', 'GEO', body)
    worksheet168.write('J5', 'EKO', body)
    worksheet168.write('K5', 'SOS', body)
    worksheet168.write('L5', 'JML', body)
    worksheet168.write('M5', 'MAT', body)
    worksheet168.write('N5', 'SEJ', body)
    worksheet168.write('O5', 'GEO', body)
    worksheet168.write('P5', 'EKO', body)
    worksheet168.write('Q5', 'SOS', body)
    worksheet168.write('R5', 'JML', body)

    worksheet168.conditional_format(5, 0, row168_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet168.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF HALIM', title)
    worksheet168.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet168.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet168.write('A22', 'LOKASI', header)
    worksheet168.write('B22', 'TOTAL', header)
    worksheet168.merge_range('A21:B21', 'RANK', header)
    worksheet168.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet168.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet168.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet168.merge_range('F21:F22', 'KELAS', header)
    worksheet168.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet168.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet168.write('G22', 'MAT', body)
    worksheet168.write('H22', 'SEJ', body)
    worksheet168.write('I22', 'GEO', body)
    worksheet168.write('J22', 'EKO', body)
    worksheet168.write('K22', 'SOS', body)
    worksheet168.write('L22', 'JML', body)
    worksheet168.write('M22', 'MAT', body)
    worksheet168.write('N22', 'SEJ', body)
    worksheet168.write('O22', 'GEO', body)
    worksheet168.write('P22', 'EKO', body)
    worksheet168.write('Q22', 'SOS', body)
    worksheet168.write('R22', 'JML', body)

    worksheet168.conditional_format(22, 0, row168+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 169
    worksheet169.insert_image('A1', r'logo resmi nf.jpg')

    worksheet169.set_column('A:A', 7, center)
    worksheet169.set_column('B:B', 6, center)
    worksheet169.set_column('C:C', 18.14, center)
    worksheet169.set_column('D:D', 25, left)
    worksheet169.set_column('E:E', 13.14, left)
    worksheet169.set_column('F:F', 8.57, center)
    worksheet169.set_column('G:R', 5, center)
    worksheet169.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF TANAH MERDEKA', title)
    worksheet169.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet169.write('A5', 'LOKASI', header)
    worksheet169.write('B5', 'TOTAL', header)
    worksheet169.merge_range('A4:B4', 'RANK', header)
    worksheet169.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet169.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet169.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet169.merge_range('F4:F5', 'KELAS', header)
    worksheet169.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet169.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet169.write('G5', 'MAT', body)
    worksheet169.write('H5', 'SEJ', body)
    worksheet169.write('I5', 'GEO', body)
    worksheet169.write('J5', 'EKO', body)
    worksheet169.write('K5', 'SOS', body)
    worksheet169.write('L5', 'JML', body)
    worksheet169.write('M5', 'MAT', body)
    worksheet169.write('N5', 'SEJ', body)
    worksheet169.write('O5', 'GEO', body)
    worksheet169.write('P5', 'EKO', body)
    worksheet169.write('Q5', 'SOS', body)
    worksheet169.write('R5', 'JML', body)

    worksheet169.conditional_format(5, 0, row169_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet169.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF TANAH MERDEKA', title)
    worksheet169.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet169.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet169.write('A22', 'LOKASI', header)
    worksheet169.write('B22', 'TOTAL', header)
    worksheet169.merge_range('A21:B21', 'RANK', header)
    worksheet169.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet169.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet169.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet169.merge_range('F21:F22', 'KELAS', header)
    worksheet169.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet169.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet169.write('G22', 'MAT', body)
    worksheet169.write('H22', 'SEJ', body)
    worksheet169.write('I22', 'GEO', body)
    worksheet169.write('J22', 'EKO', body)
    worksheet169.write('K22', 'SOS', body)
    worksheet169.write('L22', 'JML', body)
    worksheet169.write('M22', 'MAT', body)
    worksheet169.write('N22', 'SEJ', body)
    worksheet169.write('O22', 'GEO', body)
    worksheet169.write('P22', 'EKO', body)
    worksheet169.write('Q22', 'SOS', body)
    worksheet169.write('R22', 'JML', body)

    worksheet169.conditional_format(22, 0, row169+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 171
    worksheet171.insert_image('A1', r'logo resmi nf.jpg')

    worksheet171.set_column('A:A', 7, center)
    worksheet171.set_column('B:B', 6, center)
    worksheet171.set_column('C:C', 18.14, center)
    worksheet171.set_column('D:D', 25, left)
    worksheet171.set_column('E:E', 13.14, left)
    worksheet171.set_column('F:F', 8.57, center)
    worksheet171.set_column('G:R', 5, center)
    worksheet171.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIPUTAT', title)
    worksheet171.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet171.write('A5', 'LOKASI', header)
    worksheet171.write('B5', 'TOTAL', header)
    worksheet171.merge_range('A4:B4', 'RANK', header)
    worksheet171.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet171.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet171.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet171.merge_range('F4:F5', 'KELAS', header)
    worksheet171.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet171.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet171.write('G5', 'MAT', body)
    worksheet171.write('H5', 'SEJ', body)
    worksheet171.write('I5', 'GEO', body)
    worksheet171.write('J5', 'EKO', body)
    worksheet171.write('K5', 'SOS', body)
    worksheet171.write('L5', 'JML', body)
    worksheet171.write('M5', 'MAT', body)
    worksheet171.write('N5', 'SEJ', body)
    worksheet171.write('O5', 'GEO', body)
    worksheet171.write('P5', 'EKO', body)
    worksheet171.write('Q5', 'SOS', body)
    worksheet171.write('R5', 'JML', body)

    worksheet171.conditional_format(5, 0, row171_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet171.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CIPUTAT', title)
    worksheet171.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet171.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet171.write('A22', 'LOKASI', header)
    worksheet171.write('B22', 'TOTAL', header)
    worksheet171.merge_range('A21:B21', 'RANK', header)
    worksheet171.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet171.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet171.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet171.merge_range('F21:F22', 'KELAS', header)
    worksheet171.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet171.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet171.write('G22', 'MAT', body)
    worksheet171.write('H22', 'SEJ', body)
    worksheet171.write('I22', 'GEO', body)
    worksheet171.write('J22', 'EKO', body)
    worksheet171.write('K22', 'SOS', body)
    worksheet171.write('L22', 'JML', body)
    worksheet171.write('M22', 'MAT', body)
    worksheet171.write('N22', 'SEJ', body)
    worksheet171.write('O22', 'GEO', body)
    worksheet171.write('P22', 'EKO', body)
    worksheet171.write('Q22', 'SOS', body)
    worksheet171.write('R22', 'JML', body)

    worksheet171.conditional_format(22, 0, row171+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 173
    worksheet173.insert_image('A1', r'logo resmi nf.jpg')

    worksheet173.set_column('A:A', 7, center)
    worksheet173.set_column('B:B', 6, center)
    worksheet173.set_column('C:C', 18.14, center)
    worksheet173.set_column('D:D', 25, left)
    worksheet173.set_column('E:E', 13.14, left)
    worksheet173.set_column('F:F', 8.57, center)
    worksheet173.set_column('G:R', 5, center)
    worksheet173.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SALEMBA', title)
    worksheet173.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet173.write('A5', 'LOKASI', header)
    worksheet173.write('B5', 'TOTAL', header)
    worksheet173.merge_range('A4:B4', 'RANK', header)
    worksheet173.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet173.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet173.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet173.merge_range('F4:F5', 'KELAS', header)
    worksheet173.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet173.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet173.write('G5', 'MAT', body)
    worksheet173.write('H5', 'SEJ', body)
    worksheet173.write('I5', 'GEO', body)
    worksheet173.write('J5', 'EKO', body)
    worksheet173.write('K5', 'SOS', body)
    worksheet173.write('L5', 'JML', body)
    worksheet173.write('M5', 'MAT', body)
    worksheet173.write('N5', 'SEJ', body)
    worksheet173.write('O5', 'GEO', body)
    worksheet173.write('P5', 'EKO', body)
    worksheet173.write('Q5', 'SOS', body)
    worksheet173.write('R5', 'JML', body)

    worksheet173.conditional_format(5, 0, row173_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet173.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF SALEMBA', title)
    worksheet173.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet173.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet173.write('A22', 'LOKASI', header)
    worksheet173.write('B22', 'TOTAL', header)
    worksheet173.merge_range('A21:B21', 'RANK', header)
    worksheet173.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet173.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet173.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet173.merge_range('F21:F22', 'KELAS', header)
    worksheet173.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet173.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet173.write('G22', 'MAT', body)
    worksheet173.write('H22', 'SEJ', body)
    worksheet173.write('I22', 'GEO', body)
    worksheet173.write('J22', 'EKO', body)
    worksheet173.write('K22', 'SOS', body)
    worksheet173.write('L22', 'JML', body)
    worksheet173.write('M22', 'MAT', body)
    worksheet173.write('N22', 'SEJ', body)
    worksheet173.write('O22', 'GEO', body)
    worksheet173.write('P22', 'EKO', body)
    worksheet173.write('Q22', 'SOS', body)
    worksheet173.write('R22', 'JML', body)

    worksheet173.conditional_format(22, 0, row173+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 174
    worksheet174.insert_image('A1', r'logo resmi nf.jpg')

    worksheet174.set_column('A:A', 7, center)
    worksheet174.set_column('B:B', 6, center)
    worksheet174.set_column('C:C', 18.14, center)
    worksheet174.set_column('D:D', 25, left)
    worksheet174.set_column('E:E', 13.14, left)
    worksheet174.set_column('F:F', 8.57, center)
    worksheet174.set_column('G:R', 5, center)
    worksheet174.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIPINANG', title)
    worksheet174.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet174.write('A5', 'LOKASI', header)
    worksheet174.write('B5', 'TOTAL', header)
    worksheet174.merge_range('A4:B4', 'RANK', header)
    worksheet174.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet174.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet174.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet174.merge_range('F4:F5', 'KELAS', header)
    worksheet174.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet174.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet174.write('G5', 'MAT', body)
    worksheet174.write('H5', 'SEJ', body)
    worksheet174.write('I5', 'GEO', body)
    worksheet174.write('J5', 'EKO', body)
    worksheet174.write('K5', 'SOS', body)
    worksheet174.write('L5', 'JML', body)
    worksheet174.write('M5', 'MAT', body)
    worksheet174.write('N5', 'SEJ', body)
    worksheet174.write('O5', 'GEO', body)
    worksheet174.write('P5', 'EKO', body)
    worksheet174.write('Q5', 'SOS', body)
    worksheet174.write('R5', 'JML', body)

    worksheet174.conditional_format(5, 0, row174_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet174.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CIPINANG', title)
    worksheet174.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet174.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet174.write('A22', 'LOKASI', header)
    worksheet174.write('B22', 'TOTAL', header)
    worksheet174.merge_range('A21:B21', 'RANK', header)
    worksheet174.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet174.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet174.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet174.merge_range('F21:F22', 'KELAS', header)
    worksheet174.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet174.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet174.write('G22', 'MAT', body)
    worksheet174.write('H22', 'SEJ', body)
    worksheet174.write('I22', 'GEO', body)
    worksheet174.write('J22', 'EKO', body)
    worksheet174.write('K22', 'SOS', body)
    worksheet174.write('L22', 'JML', body)
    worksheet174.write('M22', 'MAT', body)
    worksheet174.write('N22', 'SEJ', body)
    worksheet174.write('O22', 'GEO', body)
    worksheet174.write('P22', 'EKO', body)
    worksheet174.write('Q22', 'SOS', body)
    worksheet174.write('R22', 'JML', body)

    worksheet174.conditional_format(22, 0, row174+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 175
    worksheet175.insert_image('A1', r'logo resmi nf.jpg')

    worksheet175.set_column('A:A', 7, center)
    worksheet175.set_column('B:B', 6, center)
    worksheet175.set_column('C:C', 18.14, center)
    worksheet175.set_column('D:D', 25, left)
    worksheet175.set_column('E:E', 13.14, left)
    worksheet175.set_column('F:F', 8.57, center)
    worksheet175.set_column('G:R', 5, center)
    worksheet175.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KRAMAT ASEM', title)
    worksheet175.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet175.write('A5', 'LOKASI', header)
    worksheet175.write('B5', 'TOTAL', header)
    worksheet175.merge_range('A4:B4', 'RANK', header)
    worksheet175.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet175.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet175.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet175.merge_range('F4:F5', 'KELAS', header)
    worksheet175.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet175.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet175.write('G5', 'MAT', body)
    worksheet175.write('H5', 'SEJ', body)
    worksheet175.write('I5', 'GEO', body)
    worksheet175.write('J5', 'EKO', body)
    worksheet175.write('K5', 'SOS', body)
    worksheet175.write('L5', 'JML', body)
    worksheet175.write('M5', 'MAT', body)
    worksheet175.write('N5', 'SEJ', body)
    worksheet175.write('O5', 'GEO', body)
    worksheet175.write('P5', 'EKO', body)
    worksheet175.write('Q5', 'SOS', body)
    worksheet175.write('R5', 'JML', body)

    worksheet175.conditional_format(5, 0, row175_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet175.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF KRAMAT ASEM', title)
    worksheet175.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet175.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet175.write('A22', 'LOKASI', header)
    worksheet175.write('B22', 'TOTAL', header)
    worksheet175.merge_range('A21:B21', 'RANK', header)
    worksheet175.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet175.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet175.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet175.merge_range('F21:F22', 'KELAS', header)
    worksheet175.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet175.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet175.write('G22', 'MAT', body)
    worksheet175.write('H22', 'SEJ', body)
    worksheet175.write('I22', 'GEO', body)
    worksheet175.write('J22', 'EKO', body)
    worksheet175.write('K22', 'SOS', body)
    worksheet175.write('L22', 'JML', body)
    worksheet175.write('M22', 'MAT', body)
    worksheet175.write('N22', 'SEJ', body)
    worksheet175.write('O22', 'GEO', body)
    worksheet175.write('P22', 'EKO', body)
    worksheet175.write('Q22', 'SOS', body)
    worksheet175.write('R22', 'JML', body)

    worksheet175.conditional_format(22, 0, row175+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 176
    worksheet176.insert_image('A1', r'logo resmi nf.jpg')

    worksheet176.set_column('A:A', 7, center)
    worksheet176.set_column('B:B', 6, center)
    worksheet176.set_column('C:C', 18.14, center)
    worksheet176.set_column('D:D', 25, left)
    worksheet176.set_column('E:E', 13.14, left)
    worksheet176.set_column('F:F', 8.57, center)
    worksheet176.set_column('G:R', 5, center)
    worksheet176.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PANGKALAN ASEM', title)
    worksheet176.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet176.write('A5', 'LOKASI', header)
    worksheet176.write('B5', 'TOTAL', header)
    worksheet176.merge_range('A4:B4', 'RANK', header)
    worksheet176.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet176.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet176.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet176.merge_range('F4:F5', 'KELAS', header)
    worksheet176.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet176.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet176.write('G5', 'MAT', body)
    worksheet176.write('H5', 'SEJ', body)
    worksheet176.write('I5', 'GEO', body)
    worksheet176.write('J5', 'EKO', body)
    worksheet176.write('K5', 'SOS', body)
    worksheet176.write('L5', 'JML', body)
    worksheet176.write('M5', 'MAT', body)
    worksheet176.write('N5', 'SEJ', body)
    worksheet176.write('O5', 'GEO', body)
    worksheet176.write('P5', 'EKO', body)
    worksheet176.write('Q5', 'SOS', body)
    worksheet176.write('R5', 'JML', body)

    worksheet176.conditional_format(5, 0, row176_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet176.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PANGKALAN ASEM', title)
    worksheet176.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet176.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet176.write('A22', 'LOKASI', header)
    worksheet176.write('B22', 'TOTAL', header)
    worksheet176.merge_range('A21:B21', 'RANK', header)
    worksheet176.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet176.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet176.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet176.merge_range('F21:F22', 'KELAS', header)
    worksheet176.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet176.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet176.write('G22', 'MAT', body)
    worksheet176.write('H22', 'SEJ', body)
    worksheet176.write('I22', 'GEO', body)
    worksheet176.write('J22', 'EKO', body)
    worksheet176.write('K22', 'SOS', body)
    worksheet176.write('L22', 'JML', body)
    worksheet176.write('M22', 'MAT', body)
    worksheet176.write('N22', 'SEJ', body)
    worksheet176.write('O22', 'GEO', body)
    worksheet176.write('P22', 'EKO', body)
    worksheet176.write('Q22', 'SOS', body)
    worksheet176.write('R22', 'JML', body)

    worksheet176.conditional_format(22, 0, row176+21, 17,
                                    {'type': 'no_errors', 'format': border})
    # worksheet 177
    worksheet177.insert_image('A1', r'logo resmi nf.jpg')

    worksheet177.set_column('A:A', 7, center)
    worksheet177.set_column('B:B', 6, center)
    worksheet177.set_column('C:C', 18.14, center)
    worksheet177.set_column('D:D', 25, left)
    worksheet177.set_column('E:E', 13.14, left)
    worksheet177.set_column('F:F', 8.57, center)
    worksheet177.set_column('G:R', 5, center)
    worksheet177.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PAMULANG 2', title)
    worksheet177.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet177.write('A5', 'LOKASI', header)
    worksheet177.write('B5', 'TOTAL', header)
    worksheet177.merge_range('A4:B4', 'RANK', header)
    worksheet177.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet177.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet177.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet177.merge_range('F4:F5', 'KELAS', header)
    worksheet177.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet177.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet177.write('G5', 'MAT', body)
    worksheet177.write('H5', 'SEJ', body)
    worksheet177.write('I5', 'GEO', body)
    worksheet177.write('J5', 'EKO', body)
    worksheet177.write('K5', 'SOS', body)
    worksheet177.write('L5', 'JML', body)
    worksheet177.write('M5', 'MAT', body)
    worksheet177.write('N5', 'SEJ', body)
    worksheet177.write('O5', 'GEO', body)
    worksheet177.write('P5', 'EKO', body)
    worksheet177.write('Q5', 'SOS', body)
    worksheet177.write('R5', 'JML', body)

    worksheet177.conditional_format(5, 0, row177_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet177.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PAMULANG 2', title)
    worksheet177.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet177.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet177.write('A22', 'LOKASI', header)
    worksheet177.write('B22', 'TOTAL', header)
    worksheet177.merge_range('A21:B21', 'RANK', header)
    worksheet177.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet177.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet177.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet177.merge_range('F21:F22', 'KELAS', header)
    worksheet177.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet177.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet177.write('G22', 'MAT', body)
    worksheet177.write('H22', 'SEJ', body)
    worksheet177.write('I22', 'GEO', body)
    worksheet177.write('J22', 'EKO', body)
    worksheet177.write('K22', 'SOS', body)
    worksheet177.write('L22', 'JML', body)
    worksheet177.write('M22', 'MAT', body)
    worksheet177.write('N22', 'SEJ', body)
    worksheet177.write('O22', 'GEO', body)
    worksheet177.write('P22', 'EKO', body)
    worksheet177.write('Q22', 'SOS', body)
    worksheet177.write('R22', 'JML', body)

    worksheet177.conditional_format(22, 0, row177+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 178
    worksheet178.insert_image('A1', r'logo resmi nf.jpg')

    worksheet178.set_column('A:A', 7, center)
    worksheet178.set_column('B:B', 6, center)
    worksheet178.set_column('C:C', 18.14, center)
    worksheet178.set_column('D:D', 25, left)
    worksheet178.set_column('E:E', 13.14, left)
    worksheet178.set_column('F:F', 8.57, center)
    worksheet178.set_column('G:R', 5, center)
    worksheet178.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PURI BETA LARANGAN', title)
    worksheet178.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet178.write('A5', 'LOKASI', header)
    worksheet178.write('B5', 'TOTAL', header)
    worksheet178.merge_range('A4:B4', 'RANK', header)
    worksheet178.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet178.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet178.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet178.merge_range('F4:F5', 'KELAS', header)
    worksheet178.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet178.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet178.write('G5', 'MAT', body)
    worksheet178.write('H5', 'SEJ', body)
    worksheet178.write('I5', 'GEO', body)
    worksheet178.write('J5', 'EKO', body)
    worksheet178.write('K5', 'SOS', body)
    worksheet178.write('L5', 'JML', body)
    worksheet178.write('M5', 'MAT', body)
    worksheet178.write('N5', 'SEJ', body)
    worksheet178.write('O5', 'GEO', body)
    worksheet178.write('P5', 'EKO', body)
    worksheet178.write('Q5', 'SOS', body)
    worksheet178.write('R5', 'JML', body)

    worksheet178.conditional_format(5, 0, row178_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet178.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PURI BETA LARANGAN', title)
    worksheet178.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet178.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet178.write('A22', 'LOKASI', header)
    worksheet178.write('B22', 'TOTAL', header)
    worksheet178.merge_range('A21:B21', 'RANK', header)
    worksheet178.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet178.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet178.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet178.merge_range('F21:F22', 'KELAS', header)
    worksheet178.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet178.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet178.write('G22', 'MAT', body)
    worksheet178.write('H22', 'SEJ', body)
    worksheet178.write('I22', 'GEO', body)
    worksheet178.write('J22', 'EKO', body)
    worksheet178.write('K22', 'SOS', body)
    worksheet178.write('L22', 'JML', body)
    worksheet178.write('M22', 'MAT', body)
    worksheet178.write('N22', 'SEJ', body)
    worksheet178.write('O22', 'GEO', body)
    worksheet178.write('P22', 'EKO', body)
    worksheet178.write('Q22', 'SOS', body)
    worksheet178.write('R22', 'JML', body)

    worksheet178.conditional_format(22, 0, row178+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 179
    worksheet179.insert_image('A1', r'logo resmi nf.jpg')

    worksheet179.set_column('A:A', 7, center)
    worksheet179.set_column('B:B', 6, center)
    worksheet179.set_column('C:C', 18.14, center)
    worksheet179.set_column('D:D', 25, left)
    worksheet179.set_column('E:E', 13.14, left)
    worksheet179.set_column('F:F', 8.57, center)
    worksheet179.set_column('G:R', 5, center)
    worksheet179.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CEGER', title)
    worksheet179.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet179.write('A5', 'LOKASI', header)
    worksheet179.write('B5', 'TOTAL', header)
    worksheet179.merge_range('A4:B4', 'RANK', header)
    worksheet179.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet179.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet179.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet179.merge_range('F4:F5', 'KELAS', header)
    worksheet179.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet179.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet179.write('G5', 'MAT', body)
    worksheet179.write('H5', 'SEJ', body)
    worksheet179.write('I5', 'GEO', body)
    worksheet179.write('J5', 'EKO', body)
    worksheet179.write('K5', 'SOS', body)
    worksheet179.write('L5', 'JML', body)
    worksheet179.write('M5', 'MAT', body)
    worksheet179.write('N5', 'SEJ', body)
    worksheet179.write('O5', 'GEO', body)
    worksheet179.write('P5', 'EKO', body)
    worksheet179.write('Q5', 'SOS', body)
    worksheet179.write('R5', 'JML', body)

    worksheet179.conditional_format(5, 0, row179_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet179.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CEGER', title)
    worksheet179.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet179.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet179.write('A22', 'LOKASI', header)
    worksheet179.write('B22', 'TOTAL', header)
    worksheet179.merge_range('A21:B21', 'RANK', header)
    worksheet179.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet179.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet179.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet179.merge_range('F21:F22', 'KELAS', header)
    worksheet179.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet179.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet179.write('G22', 'MAT', body)
    worksheet179.write('H22', 'SEJ', body)
    worksheet179.write('I22', 'GEO', body)
    worksheet179.write('J22', 'EKO', body)
    worksheet179.write('K22', 'SOS', body)
    worksheet179.write('L22', 'JML', body)
    worksheet179.write('M22', 'MAT', body)
    worksheet179.write('N22', 'SEJ', body)
    worksheet179.write('O22', 'GEO', body)
    worksheet179.write('P22', 'EKO', body)
    worksheet179.write('Q22', 'SOS', body)
    worksheet179.write('R22', 'JML', body)

    worksheet179.conditional_format(22, 0, row179+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 180
    worksheet180.insert_image('A1', r'logo resmi nf.jpg')

    worksheet180.set_column('A:A', 7, center)
    worksheet180.set_column('B:B', 6, center)
    worksheet180.set_column('C:C', 18.14, center)
    worksheet180.set_column('D:D', 25, left)
    worksheet180.set_column('E:E', 13.14, left)
    worksheet180.set_column('F:F', 8.57, center)
    worksheet180.set_column('G:R', 5, center)
    worksheet180.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SMA KOMPLEK', title)
    worksheet180.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet180.write('A5', 'LOKASI', header)
    worksheet180.write('B5', 'TOTAL', header)
    worksheet180.merge_range('A4:B4', 'RANK', header)
    worksheet180.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet180.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet180.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet180.merge_range('F4:F5', 'KELAS', header)
    worksheet180.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet180.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet180.write('G5', 'MAT', body)
    worksheet180.write('H5', 'SEJ', body)
    worksheet180.write('I5', 'GEO', body)
    worksheet180.write('J5', 'EKO', body)
    worksheet180.write('K5', 'SOS', body)
    worksheet180.write('L5', 'JML', body)
    worksheet180.write('M5', 'MAT', body)
    worksheet180.write('N5', 'SEJ', body)
    worksheet180.write('O5', 'GEO', body)
    worksheet180.write('P5', 'EKO', body)
    worksheet180.write('Q5', 'SOS', body)
    worksheet180.write('R5', 'JML', body)

    worksheet180.conditional_format(5, 0, row180_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet180.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF SMA KOMPLEK', title)
    worksheet180.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet180.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet180.write('A22', 'LOKASI', header)
    worksheet180.write('B22', 'TOTAL', header)
    worksheet180.merge_range('A21:B21', 'RANK', header)
    worksheet180.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet180.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet180.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet180.merge_range('F21:F22', 'KELAS', header)
    worksheet180.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet180.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet180.write('G22', 'MAT', body)
    worksheet180.write('H22', 'SEJ', body)
    worksheet180.write('I22', 'GEO', body)
    worksheet180.write('J22', 'EKO', body)
    worksheet180.write('K22', 'SOS', body)
    worksheet180.write('L22', 'JML', body)
    worksheet180.write('M22', 'MAT', body)
    worksheet180.write('N22', 'SEJ', body)
    worksheet180.write('O22', 'GEO', body)
    worksheet180.write('P22', 'EKO', body)
    worksheet180.write('Q22', 'SOS', body)
    worksheet180.write('R22', 'JML', body)

    worksheet180.conditional_format(22, 0, row180+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 181
    worksheet181.insert_image('A1', r'logo resmi nf.jpg')

    worksheet181.set_column('A:A', 7, center)
    worksheet181.set_column('B:B', 6, center)
    worksheet181.set_column('C:C', 18.14, center)
    worksheet181.set_column('D:D', 25, left)
    worksheet181.set_column('E:E', 13.14, left)
    worksheet181.set_column('F:F', 8.57, center)
    worksheet181.set_column('G:R', 5, center)
    worksheet181.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF GAYUNGSARI', title)
    worksheet181.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet181.write('A5', 'LOKASI', header)
    worksheet181.write('B5', 'TOTAL', header)
    worksheet181.merge_range('A4:B4', 'RANK', header)
    worksheet181.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet181.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet181.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet181.merge_range('F4:F5', 'KELAS', header)
    worksheet181.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet181.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet181.write('G5', 'MAT', body)
    worksheet181.write('H5', 'SEJ', body)
    worksheet181.write('I5', 'GEO', body)
    worksheet181.write('J5', 'EKO', body)
    worksheet181.write('K5', 'SOS', body)
    worksheet181.write('L5', 'JML', body)
    worksheet181.write('M5', 'MAT', body)
    worksheet181.write('N5', 'SEJ', body)
    worksheet181.write('O5', 'GEO', body)
    worksheet181.write('P5', 'EKO', body)
    worksheet181.write('Q5', 'SOS', body)
    worksheet181.write('R5', 'JML', body)

    worksheet181.conditional_format(5, 0, row181_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet181.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF GAYUNGSARI', title)
    worksheet181.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet181.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet181.write('A22', 'LOKASI', header)
    worksheet181.write('B22', 'TOTAL', header)
    worksheet181.merge_range('A21:B21', 'RANK', header)
    worksheet181.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet181.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet181.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet181.merge_range('F21:F22', 'KELAS', header)
    worksheet181.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet181.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet181.write('G22', 'MAT', body)
    worksheet181.write('H22', 'SEJ', body)
    worksheet181.write('I22', 'GEO', body)
    worksheet181.write('J22', 'EKO', body)
    worksheet181.write('K22', 'SOS', body)
    worksheet181.write('L22', 'JML', body)
    worksheet181.write('M22', 'MAT', body)
    worksheet181.write('N22', 'SEJ', body)
    worksheet181.write('O22', 'GEO', body)
    worksheet181.write('P22', 'EKO', body)
    worksheet181.write('Q22', 'SOS', body)
    worksheet181.write('R22', 'JML', body)

    worksheet181.conditional_format(22, 0, row181+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 182
    worksheet182.insert_image('A1', r'logo resmi nf.jpg')

    worksheet182.set_column('A:A', 7, center)
    worksheet182.set_column('B:B', 6, center)
    worksheet182.set_column('C:C', 18.14, center)
    worksheet182.set_column('D:D', 25, left)
    worksheet182.set_column('E:E', 13.14, left)
    worksheet182.set_column('F:F', 8.57, center)
    worksheet182.set_column('G:R', 5, center)
    worksheet182.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF TUPAREV', title)
    worksheet182.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet182.write('A5', 'LOKASI', header)
    worksheet182.write('B5', 'TOTAL', header)
    worksheet182.merge_range('A4:B4', 'RANK', header)
    worksheet182.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet182.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet182.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet182.merge_range('F4:F5', 'KELAS', header)
    worksheet182.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet182.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet182.write('G5', 'MAT', body)
    worksheet182.write('H5', 'SEJ', body)
    worksheet182.write('I5', 'GEO', body)
    worksheet182.write('J5', 'EKO', body)
    worksheet182.write('K5', 'SOS', body)
    worksheet182.write('L5', 'JML', body)
    worksheet182.write('M5', 'MAT', body)
    worksheet182.write('N5', 'SEJ', body)
    worksheet182.write('O5', 'GEO', body)
    worksheet182.write('P5', 'EKO', body)
    worksheet182.write('Q5', 'SOS', body)
    worksheet182.write('R5', 'JML', body)

    worksheet182.conditional_format(5, 0, row182_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet182.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF TUPAREV', title)
    worksheet182.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet182.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet182.write('A22', 'LOKASI', header)
    worksheet182.write('B22', 'TOTAL', header)
    worksheet182.merge_range('A21:B21', 'RANK', header)
    worksheet182.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet182.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet182.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet182.merge_range('F21:F22', 'KELAS', header)
    worksheet182.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet182.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet182.write('G22', 'MAT', body)
    worksheet182.write('H22', 'SEJ', body)
    worksheet182.write('I22', 'GEO', body)
    worksheet182.write('J22', 'EKO', body)
    worksheet182.write('K22', 'SOS', body)
    worksheet182.write('L22', 'JML', body)
    worksheet182.write('M22', 'MAT', body)
    worksheet182.write('N22', 'SEJ', body)
    worksheet182.write('O22', 'GEO', body)
    worksheet182.write('P22', 'EKO', body)
    worksheet182.write('Q22', 'SOS', body)
    worksheet182.write('R22', 'JML', body)

    worksheet182.conditional_format(22, 0, row182+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 183
    worksheet183.insert_image('A1', r'logo resmi nf.jpg')

    worksheet183.set_column('A:A', 7, center)
    worksheet183.set_column('B:B', 6, center)
    worksheet183.set_column('C:C', 18.14, center)
    worksheet183.set_column('D:D', 25, left)
    worksheet183.set_column('E:E', 13.14, left)
    worksheet183.set_column('F:F', 8.57, center)
    worksheet183.set_column('G:R', 5, center)
    worksheet183.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PERUMNAS KLENDER', title)
    worksheet183.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet183.write('A5', 'LOKASI', header)
    worksheet183.write('B5', 'TOTAL', header)
    worksheet183.merge_range('A4:B4', 'RANK', header)
    worksheet183.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet183.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet183.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet183.merge_range('F4:F5', 'KELAS', header)
    worksheet183.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet183.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet183.write('G5', 'MAT', body)
    worksheet183.write('H5', 'SEJ', body)
    worksheet183.write('I5', 'GEO', body)
    worksheet183.write('J5', 'EKO', body)
    worksheet183.write('K5', 'SOS', body)
    worksheet183.write('L5', 'JML', body)
    worksheet183.write('M5', 'MAT', body)
    worksheet183.write('N5', 'SEJ', body)
    worksheet183.write('O5', 'GEO', body)
    worksheet183.write('P5', 'EKO', body)
    worksheet183.write('Q5', 'SOS', body)
    worksheet183.write('R5', 'JML', body)

    worksheet183.conditional_format(5, 0, row183_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet183.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PERUMNAS KLENDER', title)
    worksheet183.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet183.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet183.write('A22', 'LOKASI', header)
    worksheet183.write('B22', 'TOTAL', header)
    worksheet183.merge_range('A21:B21', 'RANK', header)
    worksheet183.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet183.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet183.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet183.merge_range('F21:F22', 'KELAS', header)
    worksheet183.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet183.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet183.write('G22', 'MAT', body)
    worksheet183.write('H22', 'SEJ', body)
    worksheet183.write('I22', 'GEO', body)
    worksheet183.write('J22', 'EKO', body)
    worksheet183.write('K22', 'SOS', body)
    worksheet183.write('L22', 'JML', body)
    worksheet183.write('M22', 'MAT', body)
    worksheet183.write('N22', 'SEJ', body)
    worksheet183.write('O22', 'GEO', body)
    worksheet183.write('P22', 'EKO', body)
    worksheet183.write('Q22', 'SOS', body)
    worksheet183.write('R22', 'JML', body)

    worksheet183.conditional_format(22, 0, row183+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 184
    worksheet184.insert_image('A1', r'logo resmi nf.jpg')

    worksheet184.set_column('A:A', 7, center)
    worksheet184.set_column('B:B', 6, center)
    worksheet184.set_column('C:C', 18.14, center)
    worksheet184.set_column('D:D', 25, left)
    worksheet184.set_column('E:E', 13.14, left)
    worksheet184.set_column('F:F', 8.57, center)
    worksheet184.set_column('G:R', 5, center)
    worksheet184.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KARANG TENGAH', title)
    worksheet184.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet184.write('A5', 'LOKASI', header)
    worksheet184.write('B5', 'TOTAL', header)
    worksheet184.merge_range('A4:B4', 'RANK', header)
    worksheet184.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet184.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet184.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet184.merge_range('F4:F5', 'KELAS', header)
    worksheet184.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet184.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet184.write('G5', 'MAT', body)
    worksheet184.write('H5', 'SEJ', body)
    worksheet184.write('I5', 'GEO', body)
    worksheet184.write('J5', 'EKO', body)
    worksheet184.write('K5', 'SOS', body)
    worksheet184.write('L5', 'JML', body)
    worksheet184.write('M5', 'MAT', body)
    worksheet184.write('N5', 'SEJ', body)
    worksheet184.write('O5', 'GEO', body)
    worksheet184.write('P5', 'EKO', body)
    worksheet184.write('Q5', 'SOS', body)
    worksheet184.write('R5', 'JML', body)

    worksheet184.conditional_format(5, 0, row184_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet184.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF KARANG TENGAH', title)
    worksheet184.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet184.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet184.write('A22', 'LOKASI', header)
    worksheet184.write('B22', 'TOTAL', header)
    worksheet184.merge_range('A21:B21', 'RANK', header)
    worksheet184.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet184.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet184.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet184.merge_range('F21:F22', 'KELAS', header)
    worksheet184.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet184.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet184.write('G22', 'MAT', body)
    worksheet184.write('H22', 'SEJ', body)
    worksheet184.write('I22', 'GEO', body)
    worksheet184.write('J22', 'EKO', body)
    worksheet184.write('K22', 'SOS', body)
    worksheet184.write('L22', 'JML', body)
    worksheet184.write('M22', 'MAT', body)
    worksheet184.write('N22', 'SEJ', body)
    worksheet184.write('O22', 'GEO', body)
    worksheet184.write('P22', 'EKO', body)
    worksheet184.write('Q22', 'SOS', body)
    worksheet184.write('R22', 'JML', body)

    worksheet184.conditional_format(22, 0, row184+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 185
    worksheet185.insert_image('A1', r'logo resmi nf.jpg')

    worksheet185.set_column('A:A', 7, center)
    worksheet185.set_column('B:B', 6, center)
    worksheet185.set_column('C:C', 18.14, center)
    worksheet185.set_column('D:D', 25, left)
    worksheet185.set_column('E:E', 13.14, left)
    worksheet185.set_column('F:F', 8.57, center)
    worksheet185.set_column('G:R', 5, center)
    worksheet185.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SIMPANG TIGA', title)
    worksheet185.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet185.write('A5', 'LOKASI', header)
    worksheet185.write('B5', 'TOTAL', header)
    worksheet185.merge_range('A4:B4', 'RANK', header)
    worksheet185.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet185.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet185.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet185.merge_range('F4:F5', 'KELAS', header)
    worksheet185.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet185.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet185.write('G5', 'MAT', body)
    worksheet185.write('H5', 'SEJ', body)
    worksheet185.write('I5', 'GEO', body)
    worksheet185.write('J5', 'EKO', body)
    worksheet185.write('K5', 'SOS', body)
    worksheet185.write('L5', 'JML', body)
    worksheet185.write('M5', 'MAT', body)
    worksheet185.write('N5', 'SEJ', body)
    worksheet185.write('O5', 'GEO', body)
    worksheet185.write('P5', 'EKO', body)
    worksheet185.write('Q5', 'SOS', body)
    worksheet185.write('R5', 'JML', body)

    worksheet185.conditional_format(5, 0, row185_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet185.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF SIMPANG TIGA', title)
    worksheet185.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet185.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet185.write('A22', 'LOKASI', header)
    worksheet185.write('B22', 'TOTAL', header)
    worksheet185.merge_range('A21:B21', 'RANK', header)
    worksheet185.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet185.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet185.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet185.merge_range('F21:F22', 'KELAS', header)
    worksheet185.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet185.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet185.write('G22', 'MAT', body)
    worksheet185.write('H22', 'SEJ', body)
    worksheet185.write('I22', 'GEO', body)
    worksheet185.write('J22', 'EKO', body)
    worksheet185.write('K22', 'SOS', body)
    worksheet185.write('L22', 'JML', body)
    worksheet185.write('M22', 'MAT', body)
    worksheet185.write('N22', 'SEJ', body)
    worksheet185.write('O22', 'GEO', body)
    worksheet185.write('P22', 'EKO', body)
    worksheet185.write('Q22', 'SOS', body)
    worksheet185.write('R22', 'JML', body)

    worksheet185.conditional_format(22, 0, row185+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 186
    worksheet186.insert_image('A1', r'logo resmi nf.jpg')

    worksheet186.set_column('A:A', 7, center)
    worksheet186.set_column('B:B', 6, center)
    worksheet186.set_column('C:C', 18.14, center)
    worksheet186.set_column('D:D', 25, left)
    worksheet186.set_column('E:E', 13.14, left)
    worksheet186.set_column('F:F', 8.57, center)
    worksheet186.set_column('G:R', 5, center)
    worksheet186.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF RUKO PCI', title)
    worksheet186.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet186.write('A5', 'LOKASI', header)
    worksheet186.write('B5', 'TOTAL', header)
    worksheet186.merge_range('A4:B4', 'RANK', header)
    worksheet186.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet186.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet186.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet186.merge_range('F4:F5', 'KELAS', header)
    worksheet186.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet186.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet186.write('G5', 'MAT', body)
    worksheet186.write('H5', 'SEJ', body)
    worksheet186.write('I5', 'GEO', body)
    worksheet186.write('J5', 'EKO', body)
    worksheet186.write('K5', 'SOS', body)
    worksheet186.write('L5', 'JML', body)
    worksheet186.write('M5', 'MAT', body)
    worksheet186.write('N5', 'SEJ', body)
    worksheet186.write('O5', 'GEO', body)
    worksheet186.write('P5', 'EKO', body)
    worksheet186.write('Q5', 'SOS', body)
    worksheet186.write('R5', 'JML', body)

    worksheet186.conditional_format(5, 0, row186_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet186.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF RUKO PCI', title)
    worksheet186.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet186.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet186.write('A22', 'LOKASI', header)
    worksheet186.write('B22', 'TOTAL', header)
    worksheet186.merge_range('A21:B21', 'RANK', header)
    worksheet186.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet186.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet186.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet186.merge_range('F21:F22', 'KELAS', header)
    worksheet186.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet186.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet186.write('G22', 'MAT', body)
    worksheet186.write('H22', 'SEJ', body)
    worksheet186.write('I22', 'GEO', body)
    worksheet186.write('J22', 'EKO', body)
    worksheet186.write('K22', 'SOS', body)
    worksheet186.write('L22', 'JML', body)
    worksheet186.write('M22', 'MAT', body)
    worksheet186.write('N22', 'SEJ', body)
    worksheet186.write('O22', 'GEO', body)
    worksheet186.write('P22', 'EKO', body)
    worksheet186.write('Q22', 'SOS', body)
    worksheet186.write('R22', 'JML', body)

    worksheet186.conditional_format(22, 0, row186+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 187
    worksheet187.insert_image('A1', r'logo resmi nf.jpg')

    worksheet187.set_column('A:A', 7, center)
    worksheet187.set_column('B:B', 6, center)
    worksheet187.set_column('C:C', 18.14, center)
    worksheet187.set_column('D:D', 25, left)
    worksheet187.set_column('E:E', 13.14, left)
    worksheet187.set_column('F:F', 8.57, center)
    worksheet187.set_column('G:R', 5, center)
    worksheet187.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KRAMATWATU', title)
    worksheet187.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet187.write('A5', 'LOKASI', header)
    worksheet187.write('B5', 'TOTAL', header)
    worksheet187.merge_range('A4:B4', 'RANK', header)
    worksheet187.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet187.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet187.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet187.merge_range('F4:F5', 'KELAS', header)
    worksheet187.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet187.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet187.write('G5', 'MAT', body)
    worksheet187.write('H5', 'SEJ', body)
    worksheet187.write('I5', 'GEO', body)
    worksheet187.write('J5', 'EKO', body)
    worksheet187.write('K5', 'SOS', body)
    worksheet187.write('L5', 'JML', body)
    worksheet187.write('M5', 'MAT', body)
    worksheet187.write('N5', 'SEJ', body)
    worksheet187.write('O5', 'GEO', body)
    worksheet187.write('P5', 'EKO', body)
    worksheet187.write('Q5', 'SOS', body)
    worksheet187.write('R5', 'JML', body)

    worksheet187.conditional_format(5, 0, row187_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet187.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF KRAMATWATU', title)
    worksheet187.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet187.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet187.write('A22', 'LOKASI', header)
    worksheet187.write('B22', 'TOTAL', header)
    worksheet187.merge_range('A21:B21', 'RANK', header)
    worksheet187.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet187.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet187.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet187.merge_range('F21:F22', 'KELAS', header)
    worksheet187.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet187.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet187.write('G22', 'MAT', body)
    worksheet187.write('H22', 'SEJ', body)
    worksheet187.write('I22', 'GEO', body)
    worksheet187.write('J22', 'EKO', body)
    worksheet187.write('K22', 'SOS', body)
    worksheet187.write('L22', 'JML', body)
    worksheet187.write('M22', 'MAT', body)
    worksheet187.write('N22', 'SEJ', body)
    worksheet187.write('O22', 'GEO', body)
    worksheet187.write('P22', 'EKO', body)
    worksheet187.write('Q22', 'SOS', body)
    worksheet187.write('R22', 'JML', body)

    worksheet187.conditional_format(22, 0, row187+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 188
    worksheet188.insert_image('A1', r'logo resmi nf.jpg')

    worksheet188.set_column('A:A', 7, center)
    worksheet188.set_column('B:B', 6, center)
    worksheet188.set_column('C:C', 18.14, center)
    worksheet188.set_column('D:D', 25, left)
    worksheet188.set_column('E:E', 13.14, left)
    worksheet188.set_column('F:F', 8.57, center)
    worksheet188.set_column('G:R', 5, center)
    worksheet188.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KEPANDEAN', title)
    worksheet188.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet188.write('A5', 'LOKASI', header)
    worksheet188.write('B5', 'TOTAL', header)
    worksheet188.merge_range('A4:B4', 'RANK', header)
    worksheet188.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet188.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet188.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet188.merge_range('F4:F5', 'KELAS', header)
    worksheet188.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet188.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet188.write('G5', 'MAT', body)
    worksheet188.write('H5', 'SEJ', body)
    worksheet188.write('I5', 'GEO', body)
    worksheet188.write('J5', 'EKO', body)
    worksheet188.write('K5', 'SOS', body)
    worksheet188.write('L5', 'JML', body)
    worksheet188.write('M5', 'MAT', body)
    worksheet188.write('N5', 'SEJ', body)
    worksheet188.write('O5', 'GEO', body)
    worksheet188.write('P5', 'EKO', body)
    worksheet188.write('Q5', 'SOS', body)
    worksheet188.write('R5', 'JML', body)

    worksheet188.conditional_format(5, 0, row188_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet188.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF KEPANDEAN', title)
    worksheet188.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet188.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet188.write('A22', 'LOKASI', header)
    worksheet188.write('B22', 'TOTAL', header)
    worksheet188.merge_range('A21:B21', 'RANK', header)
    worksheet188.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet188.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet188.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet188.merge_range('F21:F22', 'KELAS', header)
    worksheet188.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet188.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet188.write('G22', 'MAT', body)
    worksheet188.write('H22', 'SEJ', body)
    worksheet188.write('I22', 'GEO', body)
    worksheet188.write('J22', 'EKO', body)
    worksheet188.write('K22', 'SOS', body)
    worksheet188.write('L22', 'JML', body)
    worksheet188.write('M22', 'MAT', body)
    worksheet188.write('N22', 'SEJ', body)
    worksheet188.write('O22', 'GEO', body)
    worksheet188.write('P22', 'EKO', body)
    worksheet188.write('Q22', 'SOS', body)
    worksheet188.write('R22', 'JML', body)

    worksheet188.conditional_format(22, 0, row188+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 189
    worksheet189.insert_image('A1', r'logo resmi nf.jpg')

    worksheet189.set_column('A:A', 7, center)
    worksheet189.set_column('B:B', 6, center)
    worksheet189.set_column('C:C', 18.14, center)
    worksheet189.set_column('D:D', 25, left)
    worksheet189.set_column('E:E', 13.14, left)
    worksheet189.set_column('F:F', 8.57, center)
    worksheet189.set_column('G:R', 5, center)
    worksheet189.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PINANG', title)
    worksheet189.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet189.write('A5', 'LOKASI', header)
    worksheet189.write('B5', 'TOTAL', header)
    worksheet189.merge_range('A4:B4', 'RANK', header)
    worksheet189.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet189.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet189.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet189.merge_range('F4:F5', 'KELAS', header)
    worksheet189.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet189.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet189.write('G5', 'MAT', body)
    worksheet189.write('H5', 'SEJ', body)
    worksheet189.write('I5', 'GEO', body)
    worksheet189.write('J5', 'EKO', body)
    worksheet189.write('K5', 'SOS', body)
    worksheet189.write('L5', 'JML', body)
    worksheet189.write('M5', 'MAT', body)
    worksheet189.write('N5', 'SEJ', body)
    worksheet189.write('O5', 'GEO', body)
    worksheet189.write('P5', 'EKO', body)
    worksheet189.write('Q5', 'SOS', body)
    worksheet189.write('R5', 'JML', body)

    worksheet189.conditional_format(5, 0, row189_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet189.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PINANG', title)
    worksheet189.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet189.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet189.write('A22', 'LOKASI', header)
    worksheet189.write('B22', 'TOTAL', header)
    worksheet189.merge_range('A21:B21', 'RANK', header)
    worksheet189.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet189.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet189.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet189.merge_range('F21:F22', 'KELAS', header)
    worksheet189.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet189.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet189.write('G22', 'MAT', body)
    worksheet189.write('H22', 'SEJ', body)
    worksheet189.write('I22', 'GEO', body)
    worksheet189.write('J22', 'EKO', body)
    worksheet189.write('K22', 'SOS', body)
    worksheet189.write('L22', 'JML', body)
    worksheet189.write('M22', 'MAT', body)
    worksheet189.write('N22', 'SEJ', body)
    worksheet189.write('O22', 'GEO', body)
    worksheet189.write('P22', 'EKO', body)
    worksheet189.write('Q22', 'SOS', body)
    worksheet189.write('R22', 'JML', body)

    worksheet189.conditional_format(22, 0, row189+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 190
    worksheet190.insert_image('A1', r'logo resmi nf.jpg')

    worksheet190.set_column('A:A', 7, center)
    worksheet190.set_column('B:B', 6, center)
    worksheet190.set_column('C:C', 18.14, center)
    worksheet190.set_column('D:D', 25, left)
    worksheet190.set_column('E:E', 13.14, left)
    worksheet190.set_column('F:F', 8.57, center)
    worksheet190.set_column('G:R', 5, center)
    worksheet190.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BOJONG GEDE', title)
    worksheet190.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet190.write('A5', 'LOKASI', header)
    worksheet190.write('B5', 'TOTAL', header)
    worksheet190.merge_range('A4:B4', 'RANK', header)
    worksheet190.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet190.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet190.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet190.merge_range('F4:F5', 'KELAS', header)
    worksheet190.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet190.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet190.write('G5', 'MAT', body)
    worksheet190.write('H5', 'SEJ', body)
    worksheet190.write('I5', 'GEO', body)
    worksheet190.write('J5', 'EKO', body)
    worksheet190.write('K5', 'SOS', body)
    worksheet190.write('L5', 'JML', body)
    worksheet190.write('M5', 'MAT', body)
    worksheet190.write('N5', 'SEJ', body)
    worksheet190.write('O5', 'GEO', body)
    worksheet190.write('P5', 'EKO', body)
    worksheet190.write('Q5', 'SOS', body)
    worksheet190.write('R5', 'JML', body)

    worksheet190.conditional_format(5, 0, row190_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet190.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF BOJONG GEDE', title)
    worksheet190.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet190.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet190.write('A22', 'LOKASI', header)
    worksheet190.write('B22', 'TOTAL', header)
    worksheet190.merge_range('A21:B21', 'RANK', header)
    worksheet190.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet190.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet190.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet190.merge_range('F21:F22', 'KELAS', header)
    worksheet190.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet190.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet190.write('G22', 'MAT', body)
    worksheet190.write('H22', 'SEJ', body)
    worksheet190.write('I22', 'GEO', body)
    worksheet190.write('J22', 'EKO', body)
    worksheet190.write('K22', 'SOS', body)
    worksheet190.write('L22', 'JML', body)
    worksheet190.write('M22', 'MAT', body)
    worksheet190.write('N22', 'SEJ', body)
    worksheet190.write('O22', 'GEO', body)
    worksheet190.write('P22', 'EKO', body)
    worksheet190.write('Q22', 'SOS', body)
    worksheet190.write('R22', 'JML', body)

    worksheet190.conditional_format(22, 0, row190+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 191
    worksheet191.insert_image('A1', r'logo resmi nf.jpg')

    worksheet191.set_column('A:A', 7, center)
    worksheet191.set_column('B:B', 6, center)
    worksheet191.set_column('C:C', 18.14, center)
    worksheet191.set_column('D:D', 25, left)
    worksheet191.set_column('E:E', 13.14, left)
    worksheet191.set_column('F:F', 8.57, center)
    worksheet191.set_column('G:R', 5, center)
    worksheet191.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF POMAD', title)
    worksheet191.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet191.write('A5', 'LOKASI', header)
    worksheet191.write('B5', 'TOTAL', header)
    worksheet191.merge_range('A4:B4', 'RANK', header)
    worksheet191.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet191.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet191.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet191.merge_range('F4:F5', 'KELAS', header)
    worksheet191.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet191.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet191.write('G5', 'MAT', body)
    worksheet191.write('H5', 'SEJ', body)
    worksheet191.write('I5', 'GEO', body)
    worksheet191.write('J5', 'EKO', body)
    worksheet191.write('K5', 'SOS', body)
    worksheet191.write('L5', 'JML', body)
    worksheet191.write('M5', 'MAT', body)
    worksheet191.write('N5', 'SEJ', body)
    worksheet191.write('O5', 'GEO', body)
    worksheet191.write('P5', 'EKO', body)
    worksheet191.write('Q5', 'SOS', body)
    worksheet191.write('R5', 'JML', body)

    worksheet191.conditional_format(5, 0, row191_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet191.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF POMAD', title)
    worksheet191.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet191.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet191.write('A22', 'LOKASI', header)
    worksheet191.write('B22', 'TOTAL', header)
    worksheet191.merge_range('A21:B21', 'RANK', header)
    worksheet191.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet191.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet191.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet191.merge_range('F21:F22', 'KELAS', header)
    worksheet191.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet191.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet191.write('G22', 'MAT', body)
    worksheet191.write('H22', 'SEJ', body)
    worksheet191.write('I22', 'GEO', body)
    worksheet191.write('J22', 'EKO', body)
    worksheet191.write('K22', 'SOS', body)
    worksheet191.write('L22', 'JML', body)
    worksheet191.write('M22', 'MAT', body)
    worksheet191.write('N22', 'SEJ', body)
    worksheet191.write('O22', 'GEO', body)
    worksheet191.write('P22', 'EKO', body)
    worksheet191.write('Q22', 'SOS', body)
    worksheet191.write('R22', 'JML', body)

    worksheet191.conditional_format(22, 0, row191+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 192
    worksheet192.insert_image('A1', r'logo resmi nf.jpg')

    worksheet192.set_column('A:A', 7, center)
    worksheet192.set_column('B:B', 6, center)
    worksheet192.set_column('C:C', 18.14, center)
    worksheet192.set_column('D:D', 25, left)
    worksheet192.set_column('E:E', 13.14, left)
    worksheet192.set_column('F:F', 8.57, center)
    worksheet192.set_column('G:R', 5, center)
    worksheet192.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CONDET', title)
    worksheet192.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet192.write('A5', 'LOKASI', header)
    worksheet192.write('B5', 'TOTAL', header)
    worksheet192.merge_range('A4:B4', 'RANK', header)
    worksheet192.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet192.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet192.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet192.merge_range('F4:F5', 'KELAS', header)
    worksheet192.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet192.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet192.write('G5', 'MAT', body)
    worksheet192.write('H5', 'SEJ', body)
    worksheet192.write('I5', 'GEO', body)
    worksheet192.write('J5', 'EKO', body)
    worksheet192.write('K5', 'SOS', body)
    worksheet192.write('L5', 'JML', body)
    worksheet192.write('M5', 'MAT', body)
    worksheet192.write('N5', 'SEJ', body)
    worksheet192.write('O5', 'GEO', body)
    worksheet192.write('P5', 'EKO', body)
    worksheet192.write('Q5', 'SOS', body)
    worksheet192.write('R5', 'JML', body)

    worksheet192.conditional_format(5, 0, row192_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet192.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CONDET', title)
    worksheet192.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet192.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet192.write('A22', 'LOKASI', header)
    worksheet192.write('B22', 'TOTAL', header)
    worksheet192.merge_range('A21:B21', 'RANK', header)
    worksheet192.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet192.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet192.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet192.merge_range('F21:F22', 'KELAS', header)
    worksheet192.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet192.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet192.write('G22', 'MAT', body)
    worksheet192.write('H22', 'SEJ', body)
    worksheet192.write('I22', 'GEO', body)
    worksheet192.write('J22', 'EKO', body)
    worksheet192.write('K22', 'SOS', body)
    worksheet192.write('L22', 'JML', body)
    worksheet192.write('M22', 'MAT', body)
    worksheet192.write('N22', 'SEJ', body)
    worksheet192.write('O22', 'GEO', body)
    worksheet192.write('P22', 'EKO', body)
    worksheet192.write('Q22', 'SOS', body)
    worksheet192.write('R22', 'JML', body)

    worksheet192.conditional_format(22, 0, row192+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 193
    worksheet193.insert_image('A1', r'logo resmi nf.jpg')

    worksheet193.set_column('A:A', 7, center)
    worksheet193.set_column('B:B', 6, center)
    worksheet193.set_column('C:C', 18.14, center)
    worksheet193.set_column('D:D', 25, left)
    worksheet193.set_column('E:E', 13.14, left)
    worksheet193.set_column('F:F', 8.57, center)
    worksheet193.set_column('G:R', 5, center)
    worksheet193.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF JOMBANG', title)
    worksheet193.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet193.write('A5', 'LOKASI', header)
    worksheet193.write('B5', 'TOTAL', header)
    worksheet193.merge_range('A4:B4', 'RANK', header)
    worksheet193.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet193.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet193.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet193.merge_range('F4:F5', 'KELAS', header)
    worksheet193.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet193.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet193.write('G5', 'MAT', body)
    worksheet193.write('H5', 'SEJ', body)
    worksheet193.write('I5', 'GEO', body)
    worksheet193.write('J5', 'EKO', body)
    worksheet193.write('K5', 'SOS', body)
    worksheet193.write('L5', 'JML', body)
    worksheet193.write('M5', 'MAT', body)
    worksheet193.write('N5', 'SEJ', body)
    worksheet193.write('O5', 'GEO', body)
    worksheet193.write('P5', 'EKO', body)
    worksheet193.write('Q5', 'SOS', body)
    worksheet193.write('R5', 'JML', body)

    worksheet193.conditional_format(5, 0, row193_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet193.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF JOMBANG', title)
    worksheet193.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet193.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet193.write('A22', 'LOKASI', header)
    worksheet193.write('B22', 'TOTAL', header)
    worksheet193.merge_range('A21:B21', 'RANK', header)
    worksheet193.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet193.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet193.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet193.merge_range('F21:F22', 'KELAS', header)
    worksheet193.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet193.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet193.write('G22', 'MAT', body)
    worksheet193.write('H22', 'SEJ', body)
    worksheet193.write('I22', 'GEO', body)
    worksheet193.write('J22', 'EKO', body)
    worksheet193.write('K22', 'SOS', body)
    worksheet193.write('L22', 'JML', body)
    worksheet193.write('M22', 'MAT', body)
    worksheet193.write('N22', 'SEJ', body)
    worksheet193.write('O22', 'GEO', body)
    worksheet193.write('P22', 'EKO', body)
    worksheet193.write('Q22', 'SOS', body)
    worksheet193.write('R22', 'JML', body)

    worksheet193.conditional_format(22, 0, row193+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 194
    worksheet194.insert_image('A1', r'logo resmi nf.jpg')

    worksheet194.set_column('A:A', 7, center)
    worksheet194.set_column('B:B', 6, center)
    worksheet194.set_column('C:C', 18.14, center)
    worksheet194.set_column('D:D', 25, left)
    worksheet194.set_column('E:E', 13.14, left)
    worksheet194.set_column('F:F', 8.57, center)
    worksheet194.set_column('G:R', 5, center)
    worksheet194.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KEMAYORAN', title)
    worksheet194.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet194.write('A5', 'LOKASI', header)
    worksheet194.write('B5', 'TOTAL', header)
    worksheet194.merge_range('A4:B4', 'RANK', header)
    worksheet194.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet194.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet194.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet194.merge_range('F4:F5', 'KELAS', header)
    worksheet194.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet194.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet194.write('G5', 'MAT', body)
    worksheet194.write('H5', 'SEJ', body)
    worksheet194.write('I5', 'GEO', body)
    worksheet194.write('J5', 'EKO', body)
    worksheet194.write('K5', 'SOS', body)
    worksheet194.write('L5', 'JML', body)
    worksheet194.write('M5', 'MAT', body)
    worksheet194.write('N5', 'SEJ', body)
    worksheet194.write('O5', 'GEO', body)
    worksheet194.write('P5', 'EKO', body)
    worksheet194.write('Q5', 'SOS', body)
    worksheet194.write('R5', 'JML', body)

    worksheet194.conditional_format(5, 0, row194_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet194.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF KEMAYORAN', title)
    worksheet194.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet194.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet194.write('A22', 'LOKASI', header)
    worksheet194.write('B22', 'TOTAL', header)
    worksheet194.merge_range('A21:B21', 'RANK', header)
    worksheet194.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet194.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet194.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet194.merge_range('F21:F22', 'KELAS', header)
    worksheet194.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet194.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet194.write('G22', 'MAT', body)
    worksheet194.write('H22', 'SEJ', body)
    worksheet194.write('I22', 'GEO', body)
    worksheet194.write('J22', 'EKO', body)
    worksheet194.write('K22', 'SOS', body)
    worksheet194.write('L22', 'JML', body)
    worksheet194.write('M22', 'MAT', body)
    worksheet194.write('N22', 'SEJ', body)
    worksheet194.write('O22', 'GEO', body)
    worksheet194.write('P22', 'EKO', body)
    worksheet194.write('Q22', 'SOS', body)
    worksheet194.write('R22', 'JML', body)

    worksheet194.conditional_format(22, 0, row194+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 195
    worksheet195.insert_image('A1', r'logo resmi nf.jpg')

    worksheet195.set_column('A:A', 7, center)
    worksheet195.set_column('B:B', 6, center)
    worksheet195.set_column('C:C', 18.14, center)
    worksheet195.set_column('D:D', 25, left)
    worksheet195.set_column('E:E', 13.14, left)
    worksheet195.set_column('F:F', 8.57, center)
    worksheet195.set_column('G:R', 5, center)
    worksheet195.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KALISARI', title)
    worksheet195.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet195.write('A5', 'LOKASI', header)
    worksheet195.write('B5', 'TOTAL', header)
    worksheet195.merge_range('A4:B4', 'RANK', header)
    worksheet195.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet195.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet195.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet195.merge_range('F4:F5', 'KELAS', header)
    worksheet195.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet195.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet195.write('G5', 'MAT', body)
    worksheet195.write('H5', 'SEJ', body)
    worksheet195.write('I5', 'GEO', body)
    worksheet195.write('J5', 'EKO', body)
    worksheet195.write('K5', 'SOS', body)
    worksheet195.write('L5', 'JML', body)
    worksheet195.write('M5', 'MAT', body)
    worksheet195.write('N5', 'SEJ', body)
    worksheet195.write('O5', 'GEO', body)
    worksheet195.write('P5', 'EKO', body)
    worksheet195.write('Q5', 'SOS', body)
    worksheet195.write('R5', 'JML', body)

    worksheet195.conditional_format(5, 0, row195_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet195.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF KALISARI', title)
    worksheet195.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet195.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet195.write('A22', 'LOKASI', header)
    worksheet195.write('B22', 'TOTAL', header)
    worksheet195.merge_range('A21:B21', 'RANK', header)
    worksheet195.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet195.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet195.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet195.merge_range('F21:F22', 'KELAS', header)
    worksheet195.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet195.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet195.write('G22', 'MAT', body)
    worksheet195.write('H22', 'SEJ', body)
    worksheet195.write('I22', 'GEO', body)
    worksheet195.write('J22', 'EKO', body)
    worksheet195.write('K22', 'SOS', body)
    worksheet195.write('L22', 'JML', body)
    worksheet195.write('M22', 'MAT', body)
    worksheet195.write('N22', 'SEJ', body)
    worksheet195.write('O22', 'GEO', body)
    worksheet195.write('P22', 'EKO', body)
    worksheet195.write('Q22', 'SOS', body)
    worksheet195.write('R22', 'JML', body)

    worksheet195.conditional_format(22, 0, row195+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 196
    worksheet196.insert_image('A1', r'logo resmi nf.jpg')

    worksheet196.set_column('A:A', 7, center)
    worksheet196.set_column('B:B', 6, center)
    worksheet196.set_column('C:C', 18.14, center)
    worksheet196.set_column('D:D', 25, left)
    worksheet196.set_column('E:E', 13.14, left)
    worksheet196.set_column('F:F', 8.57, center)
    worksheet196.set_column('G:R', 5, center)
    worksheet196.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PAMULANG 1', title)
    worksheet196.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet196.write('A5', 'LOKASI', header)
    worksheet196.write('B5', 'TOTAL', header)
    worksheet196.merge_range('A4:B4', 'RANK', header)
    worksheet196.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet196.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet196.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet196.merge_range('F4:F5', 'KELAS', header)
    worksheet196.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet196.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet196.write('G5', 'MAT', body)
    worksheet196.write('H5', 'SEJ', body)
    worksheet196.write('I5', 'GEO', body)
    worksheet196.write('J5', 'EKO', body)
    worksheet196.write('K5', 'SOS', body)
    worksheet196.write('L5', 'JML', body)
    worksheet196.write('M5', 'MAT', body)
    worksheet196.write('N5', 'SEJ', body)
    worksheet196.write('O5', 'GEO', body)
    worksheet196.write('P5', 'EKO', body)
    worksheet196.write('Q5', 'SOS', body)
    worksheet196.write('R5', 'JML', body)

    worksheet196.conditional_format(5, 0, row196_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet196.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PAMULANG 1', title)
    worksheet196.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet196.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet196.write('A22', 'LOKASI', header)
    worksheet196.write('B22', 'TOTAL', header)
    worksheet196.merge_range('A21:B21', 'RANK', header)
    worksheet196.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet196.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet196.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet196.merge_range('F21:F22', 'KELAS', header)
    worksheet196.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet196.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet196.write('G22', 'MAT', body)
    worksheet196.write('H22', 'SEJ', body)
    worksheet196.write('I22', 'GEO', body)
    worksheet196.write('J22', 'EKO', body)
    worksheet196.write('K22', 'SOS', body)
    worksheet196.write('L22', 'JML', body)
    worksheet196.write('M22', 'MAT', body)
    worksheet196.write('N22', 'SEJ', body)
    worksheet196.write('O22', 'GEO', body)
    worksheet196.write('P22', 'EKO', body)
    worksheet196.write('Q22', 'SOS', body)
    worksheet196.write('R22', 'JML', body)

    worksheet196.conditional_format(22, 0, row196+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 197
    worksheet197.insert_image('A1', r'logo resmi nf.jpg')

    worksheet197.set_column('A:A', 7, center)
    worksheet197.set_column('B:B', 6, center)
    worksheet197.set_column('C:C', 18.14, center)
    worksheet197.set_column('D:D', 25, left)
    worksheet197.set_column('E:E', 13.14, left)
    worksheet197.set_column('F:F', 8.57, center)
    worksheet197.set_column('G:R', 5, center)
    worksheet197.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PANDEGLANG BARU', title)
    worksheet197.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet197.write('A5', 'LOKASI', header)
    worksheet197.write('B5', 'TOTAL', header)
    worksheet197.merge_range('A4:B4', 'RANK', header)
    worksheet197.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet197.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet197.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet197.merge_range('F4:F5', 'KELAS', header)
    worksheet197.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet197.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet197.write('G5', 'MAT', body)
    worksheet197.write('H5', 'SEJ', body)
    worksheet197.write('I5', 'GEO', body)
    worksheet197.write('J5', 'EKO', body)
    worksheet197.write('K5', 'SOS', body)
    worksheet197.write('L5', 'JML', body)
    worksheet197.write('M5', 'MAT', body)
    worksheet197.write('N5', 'SEJ', body)
    worksheet197.write('O5', 'GEO', body)
    worksheet197.write('P5', 'EKO', body)
    worksheet197.write('Q5', 'SOS', body)
    worksheet197.write('R5', 'JML', body)

    worksheet197.conditional_format(5, 0, row197_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet197.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PANDEGLANG BARU', title)
    worksheet197.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet197.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet197.write('A22', 'LOKASI', header)
    worksheet197.write('B22', 'TOTAL', header)
    worksheet197.merge_range('A21:B21', 'RANK', header)
    worksheet197.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet197.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet197.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet197.merge_range('F21:F22', 'KELAS', header)
    worksheet197.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet197.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet197.write('G22', 'MAT', body)
    worksheet197.write('H22', 'SEJ', body)
    worksheet197.write('I22', 'GEO', body)
    worksheet197.write('J22', 'EKO', body)
    worksheet197.write('K22', 'SOS', body)
    worksheet197.write('L22', 'JML', body)
    worksheet197.write('M22', 'MAT', body)
    worksheet197.write('N22', 'SEJ', body)
    worksheet197.write('O22', 'GEO', body)
    worksheet197.write('P22', 'EKO', body)
    worksheet197.write('Q22', 'SOS', body)
    worksheet197.write('R22', 'JML', body)

    worksheet197.conditional_format(22, 0, row197+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 198
    worksheet198.insert_image('A1', r'logo resmi nf.jpg')

    worksheet198.set_column('A:A', 7, center)
    worksheet198.set_column('B:B', 6, center)
    worksheet198.set_column('C:C', 18.14, center)
    worksheet198.set_column('D:D', 25, left)
    worksheet198.set_column('E:E', 13.14, left)
    worksheet198.set_column('F:F', 8.57, center)
    worksheet198.set_column('G:R', 5, center)
    worksheet198.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF RUNGKUT', title)
    worksheet198.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet198.write('A5', 'LOKASI', header)
    worksheet198.write('B5', 'TOTAL', header)
    worksheet198.merge_range('A4:B4', 'RANK', header)
    worksheet198.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet198.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet198.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet198.merge_range('F4:F5', 'KELAS', header)
    worksheet198.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet198.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet198.write('G5', 'MAT', body)
    worksheet198.write('H5', 'SEJ', body)
    worksheet198.write('I5', 'GEO', body)
    worksheet198.write('J5', 'EKO', body)
    worksheet198.write('K5', 'SOS', body)
    worksheet198.write('L5', 'JML', body)
    worksheet198.write('M5', 'MAT', body)
    worksheet198.write('N5', 'SEJ', body)
    worksheet198.write('O5', 'GEO', body)
    worksheet198.write('P5', 'EKO', body)
    worksheet198.write('Q5', 'SOS', body)
    worksheet198.write('R5', 'JML', body)

    worksheet198.conditional_format(5, 0, row198_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet198.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF RUNGKUT', title)
    worksheet198.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet198.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet198.write('A22', 'LOKASI', header)
    worksheet198.write('B22', 'TOTAL', header)
    worksheet198.merge_range('A21:B21', 'RANK', header)
    worksheet198.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet198.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet198.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet198.merge_range('F21:F22', 'KELAS', header)
    worksheet198.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet198.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet198.write('G22', 'MAT', body)
    worksheet198.write('H22', 'SEJ', body)
    worksheet198.write('I22', 'GEO', body)
    worksheet198.write('J22', 'EKO', body)
    worksheet198.write('K22', 'SOS', body)
    worksheet198.write('L22', 'JML', body)
    worksheet198.write('M22', 'MAT', body)
    worksheet198.write('N22', 'SEJ', body)
    worksheet198.write('O22', 'GEO', body)
    worksheet198.write('P22', 'EKO', body)
    worksheet198.write('Q22', 'SOS', body)
    worksheet198.write('R22', 'JML', body)

    worksheet198.conditional_format(22, 0, row198+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 199
    worksheet199.insert_image('A1', r'logo resmi nf.jpg')

    worksheet199.set_column('A:A', 7, center)
    worksheet199.set_column('B:B', 6, center)
    worksheet199.set_column('C:C', 18.14, center)
    worksheet199.set_column('D:D', 25, left)
    worksheet199.set_column('E:E', 13.14, left)
    worksheet199.set_column('F:F', 8.57, center)
    worksheet199.set_column('G:R', 5, center)
    worksheet199.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIOMAS', title)
    worksheet199.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet199.write('A5', 'LOKASI', header)
    worksheet199.write('B5', 'TOTAL', header)
    worksheet199.merge_range('A4:B4', 'RANK', header)
    worksheet199.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet199.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet199.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet199.merge_range('F4:F5', 'KELAS', header)
    worksheet199.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet199.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet199.write('G5', 'MAT', body)
    worksheet199.write('H5', 'SEJ', body)
    worksheet199.write('I5', 'GEO', body)
    worksheet199.write('J5', 'EKO', body)
    worksheet199.write('K5', 'SOS', body)
    worksheet199.write('L5', 'JML', body)
    worksheet199.write('M5', 'MAT', body)
    worksheet199.write('N5', 'SEJ', body)
    worksheet199.write('O5', 'GEO', body)
    worksheet199.write('P5', 'EKO', body)
    worksheet199.write('Q5', 'SOS', body)
    worksheet199.write('R5', 'JML', body)

    worksheet199.conditional_format(5, 0, row199_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet199.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CIOMAS', title)
    worksheet199.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet199.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet199.write('A22', 'LOKASI', header)
    worksheet199.write('B22', 'TOTAL', header)
    worksheet199.merge_range('A21:B21', 'RANK', header)
    worksheet199.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet199.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet199.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet199.merge_range('F21:F22', 'KELAS', header)
    worksheet199.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet199.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet199.write('G22', 'MAT', body)
    worksheet199.write('H22', 'SEJ', body)
    worksheet199.write('I22', 'GEO', body)
    worksheet199.write('J22', 'EKO', body)
    worksheet199.write('K22', 'SOS', body)
    worksheet199.write('L22', 'JML', body)
    worksheet199.write('M22', 'MAT', body)
    worksheet199.write('N22', 'SEJ', body)
    worksheet199.write('O22', 'GEO', body)
    worksheet199.write('P22', 'EKO', body)
    worksheet199.write('Q22', 'SOS', body)
    worksheet199.write('R22', 'JML', body)

    worksheet199.conditional_format(22, 0, row199+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 201
    worksheet201.insert_image('A1', r'logo resmi nf.jpg')

    worksheet201.set_column('A:A', 7, center)
    worksheet201.set_column('B:B', 6, center)
    worksheet201.set_column('C:C', 18.14, center)
    worksheet201.set_column('D:D', 25, left)
    worksheet201.set_column('E:E', 13.14, left)
    worksheet201.set_column('F:F', 8.57, center)
    worksheet201.set_column('G:R', 5, center)
    worksheet201.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SUNTER JAYA', title)
    worksheet201.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet201.write('A5', 'LOKASI', header)
    worksheet201.write('B5', 'TOTAL', header)
    worksheet201.merge_range('A4:B4', 'RANK', header)
    worksheet201.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet201.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet201.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet201.merge_range('F4:F5', 'KELAS', header)
    worksheet201.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet201.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet201.write('G5', 'MAT', body)
    worksheet201.write('H5', 'SEJ', body)
    worksheet201.write('I5', 'GEO', body)
    worksheet201.write('J5', 'EKO', body)
    worksheet201.write('K5', 'SOS', body)
    worksheet201.write('L5', 'JML', body)
    worksheet201.write('M5', 'MAT', body)
    worksheet201.write('N5', 'SEJ', body)
    worksheet201.write('O5', 'GEO', body)
    worksheet201.write('P5', 'EKO', body)
    worksheet201.write('Q5', 'SOS', body)
    worksheet201.write('R5', 'JML', body)

    worksheet201.conditional_format(5, 0, row201_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet201.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF SUNTER JAYA', title)
    worksheet201.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet201.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet201.write('A22', 'LOKASI', header)
    worksheet201.write('B22', 'TOTAL', header)
    worksheet201.merge_range('A21:B21', 'RANK', header)
    worksheet201.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet201.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet201.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet201.merge_range('F21:F22', 'KELAS', header)
    worksheet201.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet201.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet201.write('G22', 'MAT', body)
    worksheet201.write('H22', 'SEJ', body)
    worksheet201.write('I22', 'GEO', body)
    worksheet201.write('J22', 'EKO', body)
    worksheet201.write('K22', 'SOS', body)
    worksheet201.write('L22', 'JML', body)
    worksheet201.write('M22', 'MAT', body)
    worksheet201.write('N22', 'SEJ', body)
    worksheet201.write('O22', 'GEO', body)
    worksheet201.write('P22', 'EKO', body)
    worksheet201.write('Q22', 'SOS', body)
    worksheet201.write('R22', 'JML', body)

    worksheet201.conditional_format(22, 0, row201+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 202
    worksheet202.insert_image('A1', r'logo resmi nf.jpg')

    worksheet202.set_column('A:A', 7, center)
    worksheet202.set_column('B:B', 6, center)
    worksheet202.set_column('C:C', 18.14, center)
    worksheet202.set_column('D:D', 25, left)
    worksheet202.set_column('E:E', 13.14, left)
    worksheet202.set_column('F:F', 8.57, center)
    worksheet202.set_column('G:R', 5, center)
    worksheet202.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PENGGILINGAN', title)
    worksheet202.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet202.write('A5', 'LOKASI', header)
    worksheet202.write('B5', 'TOTAL', header)
    worksheet202.merge_range('A4:B4', 'RANK', header)
    worksheet202.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet202.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet202.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet202.merge_range('F4:F5', 'KELAS', header)
    worksheet202.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet202.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet202.write('G5', 'MAT', body)
    worksheet202.write('H5', 'SEJ', body)
    worksheet202.write('I5', 'GEO', body)
    worksheet202.write('J5', 'EKO', body)
    worksheet202.write('K5', 'SOS', body)
    worksheet202.write('L5', 'JML', body)
    worksheet202.write('M5', 'MAT', body)
    worksheet202.write('N5', 'SEJ', body)
    worksheet202.write('O5', 'GEO', body)
    worksheet202.write('P5', 'EKO', body)
    worksheet202.write('Q5', 'SOS', body)
    worksheet202.write('R5', 'JML', body)

    worksheet202.conditional_format(5, 0, row202_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet202.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PENGGILINGAN', title)
    worksheet202.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet202.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet202.write('A22', 'LOKASI', header)
    worksheet202.write('B22', 'TOTAL', header)
    worksheet202.merge_range('A21:B21', 'RANK', header)
    worksheet202.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet202.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet202.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet202.merge_range('F21:F22', 'KELAS', header)
    worksheet202.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet202.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet202.write('G22', 'MAT', body)
    worksheet202.write('H22', 'SEJ', body)
    worksheet202.write('I22', 'GEO', body)
    worksheet202.write('J22', 'EKO', body)
    worksheet202.write('K22', 'SOS', body)
    worksheet202.write('L22', 'JML', body)
    worksheet202.write('M22', 'MAT', body)
    worksheet202.write('N22', 'SEJ', body)
    worksheet202.write('O22', 'GEO', body)
    worksheet202.write('P22', 'EKO', body)
    worksheet202.write('Q22', 'SOS', body)
    worksheet202.write('R22', 'JML', body)

    worksheet202.conditional_format(22, 0, row202+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 203
    worksheet203.insert_image('A1', r'logo resmi nf.jpg')

    worksheet203.set_column('A:A', 7, center)
    worksheet203.set_column('B:B', 6, center)
    worksheet203.set_column('C:C', 18.14, center)
    worksheet203.set_column('D:D', 25, left)
    worksheet203.set_column('E:E', 13.14, left)
    worksheet203.set_column('F:F', 8.57, center)
    worksheet203.set_column('G:R', 5, center)
    worksheet203.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF GORONTALO', title)
    worksheet203.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet203.write('A5', 'LOKASI', header)
    worksheet203.write('B5', 'TOTAL', header)
    worksheet203.merge_range('A4:B4', 'RANK', header)
    worksheet203.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet203.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet203.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet203.merge_range('F4:F5', 'KELAS', header)
    worksheet203.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet203.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet203.write('G5', 'MAT', body)
    worksheet203.write('H5', 'SEJ', body)
    worksheet203.write('I5', 'GEO', body)
    worksheet203.write('J5', 'EKO', body)
    worksheet203.write('K5', 'SOS', body)
    worksheet203.write('L5', 'JML', body)
    worksheet203.write('M5', 'MAT', body)
    worksheet203.write('N5', 'SEJ', body)
    worksheet203.write('O5', 'GEO', body)
    worksheet203.write('P5', 'EKO', body)
    worksheet203.write('Q5', 'SOS', body)
    worksheet203.write('R5', 'JML', body)

    worksheet203.conditional_format(5, 0, row203_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet203.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF GORONTALO', title)
    worksheet203.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet203.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet203.write('A22', 'LOKASI', header)
    worksheet203.write('B22', 'TOTAL', header)
    worksheet203.merge_range('A21:B21', 'RANK', header)
    worksheet203.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet203.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet203.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet203.merge_range('F21:F22', 'KELAS', header)
    worksheet203.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet203.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet203.write('G22', 'MAT', body)
    worksheet203.write('H22', 'SEJ', body)
    worksheet203.write('I22', 'GEO', body)
    worksheet203.write('J22', 'EKO', body)
    worksheet203.write('K22', 'SOS', body)
    worksheet203.write('L22', 'JML', body)
    worksheet203.write('M22', 'MAT', body)
    worksheet203.write('N22', 'SEJ', body)
    worksheet203.write('O22', 'GEO', body)
    worksheet203.write('P22', 'EKO', body)
    worksheet203.write('Q22', 'SOS', body)
    worksheet203.write('R22', 'JML', body)

    worksheet203.conditional_format(22, 0, row203+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 210
    worksheet210.insert_image('A1', r'logo resmi nf.jpg')

    worksheet210.set_column('A:A', 7, center)
    worksheet210.set_column('B:B', 6, center)
    worksheet210.set_column('C:C', 18.14, center)
    worksheet210.set_column('D:D', 25, left)
    worksheet210.set_column('E:E', 13.14, left)
    worksheet210.set_column('F:F', 8.57, center)
    worksheet210.set_column('G:R', 5, center)
    worksheet210.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KOPO', title)
    worksheet210.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet210.write('A5', 'LOKASI', header)
    worksheet210.write('B5', 'TOTAL', header)
    worksheet210.merge_range('A4:B4', 'RANK', header)
    worksheet210.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet210.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet210.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet210.merge_range('F4:F5', 'KELAS', header)
    worksheet210.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet210.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet210.write('G5', 'MAT', body)
    worksheet210.write('H5', 'SEJ', body)
    worksheet210.write('I5', 'GEO', body)
    worksheet210.write('J5', 'EKO', body)
    worksheet210.write('K5', 'SOS', body)
    worksheet210.write('L5', 'JML', body)
    worksheet210.write('M5', 'MAT', body)
    worksheet210.write('N5', 'SEJ', body)
    worksheet210.write('O5', 'GEO', body)
    worksheet210.write('P5', 'EKO', body)
    worksheet210.write('Q5', 'SOS', body)
    worksheet210.write('R5', 'JML', body)

    worksheet210.conditional_format(5, 0, row210_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet210.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF KOPO', title)
    worksheet210.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet210.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet210.write('A22', 'LOKASI', header)
    worksheet210.write('B22', 'TOTAL', header)
    worksheet210.merge_range('A21:B21', 'RANK', header)
    worksheet210.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet210.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet210.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet210.merge_range('F21:F22', 'KELAS', header)
    worksheet210.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet210.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet210.write('G22', 'MAT', body)
    worksheet210.write('H22', 'SEJ', body)
    worksheet210.write('I22', 'GEO', body)
    worksheet210.write('J22', 'EKO', body)
    worksheet210.write('K22', 'SOS', body)
    worksheet210.write('L22', 'JML', body)
    worksheet210.write('M22', 'MAT', body)
    worksheet210.write('N22', 'SEJ', body)
    worksheet210.write('O22', 'GEO', body)
    worksheet210.write('P22', 'EKO', body)
    worksheet210.write('Q22', 'SOS', body)
    worksheet210.write('R22', 'JML', body)

    worksheet210.conditional_format(22, 0, row210+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 211
    worksheet211.insert_image('A1', r'logo resmi nf.jpg')

    worksheet211.set_column('A:A', 7, center)
    worksheet211.set_column('B:B', 6, center)
    worksheet211.set_column('C:C', 18.14, center)
    worksheet211.set_column('D:D', 25, left)
    worksheet211.set_column('E:E', 13.14, left)
    worksheet211.set_column('F:F', 8.57, center)
    worksheet211.set_column('G:R', 5, center)
    worksheet211.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF VILLA INDAH PERMAI', title)
    worksheet211.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet211.write('A5', 'LOKASI', header)
    worksheet211.write('B5', 'TOTAL', header)
    worksheet211.merge_range('A4:B4', 'RANK', header)
    worksheet211.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet211.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet211.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet211.merge_range('F4:F5', 'KELAS', header)
    worksheet211.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet211.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet211.write('G5', 'MAT', body)
    worksheet211.write('H5', 'SEJ', body)
    worksheet211.write('I5', 'GEO', body)
    worksheet211.write('J5', 'EKO', body)
    worksheet211.write('K5', 'SOS', body)
    worksheet211.write('L5', 'JML', body)
    worksheet211.write('M5', 'MAT', body)
    worksheet211.write('N5', 'SEJ', body)
    worksheet211.write('O5', 'GEO', body)
    worksheet211.write('P5', 'EKO', body)
    worksheet211.write('Q5', 'SOS', body)
    worksheet211.write('R5', 'JML', body)

    worksheet211.conditional_format(5, 0, row211_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet211.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF VILLA INDAH PERMAI', title)
    worksheet211.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet211.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet211.write('A22', 'LOKASI', header)
    worksheet211.write('B22', 'TOTAL', header)
    worksheet211.merge_range('A21:B21', 'RANK', header)
    worksheet211.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet211.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet211.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet211.merge_range('F21:F22', 'KELAS', header)
    worksheet211.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet211.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet211.write('G22', 'MAT', body)
    worksheet211.write('H22', 'SEJ', body)
    worksheet211.write('I22', 'GEO', body)
    worksheet211.write('J22', 'EKO', body)
    worksheet211.write('K22', 'SOS', body)
    worksheet211.write('L22', 'JML', body)
    worksheet211.write('M22', 'MAT', body)
    worksheet211.write('N22', 'SEJ', body)
    worksheet211.write('O22', 'GEO', body)
    worksheet211.write('P22', 'EKO', body)
    worksheet211.write('Q22', 'SOS', body)
    worksheet211.write('R22', 'JML', body)

    worksheet211.conditional_format(22, 0, row211+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 212
    worksheet212.insert_image('A1', r'logo resmi nf.jpg')

    worksheet212.set_column('A:A', 7, center)
    worksheet212.set_column('B:B', 6, center)
    worksheet212.set_column('C:C', 18.14, center)
    worksheet212.set_column('D:D', 25, left)
    worksheet212.set_column('E:E', 13.14, left)
    worksheet212.set_column('F:F', 8.57, center)
    worksheet212.set_column('G:R', 5, center)
    worksheet212.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CISAUK', title)
    worksheet212.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet212.write('A5', 'LOKASI', header)
    worksheet212.write('B5', 'TOTAL', header)
    worksheet212.merge_range('A4:B4', 'RANK', header)
    worksheet212.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet212.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet212.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet212.merge_range('F4:F5', 'KELAS', header)
    worksheet212.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet212.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet212.write('G5', 'MAT', body)
    worksheet212.write('H5', 'SEJ', body)
    worksheet212.write('I5', 'GEO', body)
    worksheet212.write('J5', 'EKO', body)
    worksheet212.write('K5', 'SOS', body)
    worksheet212.write('L5', 'JML', body)
    worksheet212.write('M5', 'MAT', body)
    worksheet212.write('N5', 'SEJ', body)
    worksheet212.write('O5', 'GEO', body)
    worksheet212.write('P5', 'EKO', body)
    worksheet212.write('Q5', 'SOS', body)
    worksheet212.write('R5', 'JML', body)

    worksheet212.conditional_format(5, 0, row212_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet212.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CISAUK', title)
    worksheet212.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet212.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet212.write('A22', 'LOKASI', header)
    worksheet212.write('B22', 'TOTAL', header)
    worksheet212.merge_range('A21:B21', 'RANK', header)
    worksheet212.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet212.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet212.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet212.merge_range('F21:F22', 'KELAS', header)
    worksheet212.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet212.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet212.write('G22', 'MAT', body)
    worksheet212.write('H22', 'SEJ', body)
    worksheet212.write('I22', 'GEO', body)
    worksheet212.write('J22', 'EKO', body)
    worksheet212.write('K22', 'SOS', body)
    worksheet212.write('L22', 'JML', body)
    worksheet212.write('M22', 'MAT', body)
    worksheet212.write('N22', 'SEJ', body)
    worksheet212.write('O22', 'GEO', body)
    worksheet212.write('P22', 'EKO', body)
    worksheet212.write('Q22', 'SOS', body)
    worksheet212.write('R22', 'JML', body)

    worksheet212.conditional_format(22, 0, row212+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 216
    worksheet216.insert_image('A1', r'logo resmi nf.jpg')

    worksheet216.set_column('A:A', 7, center)
    worksheet216.set_column('B:B', 6, center)
    worksheet216.set_column('C:C', 18.14, center)
    worksheet216.set_column('D:D', 25, left)
    worksheet216.set_column('E:E', 13.14, left)
    worksheet216.set_column('F:F', 8.57, center)
    worksheet216.set_column('G:R', 5, center)
    worksheet216.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF RADIO DALAM', title)
    worksheet216.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet216.write('A5', 'LOKASI', header)
    worksheet216.write('B5', 'TOTAL', header)
    worksheet216.merge_range('A4:B4', 'RANK', header)
    worksheet216.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet216.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet216.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet216.merge_range('F4:F5', 'KELAS', header)
    worksheet216.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet216.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet216.write('G5', 'MAT', body)
    worksheet216.write('H5', 'SEJ', body)
    worksheet216.write('I5', 'GEO', body)
    worksheet216.write('J5', 'EKO', body)
    worksheet216.write('K5', 'SOS', body)
    worksheet216.write('L5', 'JML', body)
    worksheet216.write('M5', 'MAT', body)
    worksheet216.write('N5', 'SEJ', body)
    worksheet216.write('O5', 'GEO', body)
    worksheet216.write('P5', 'EKO', body)
    worksheet216.write('Q5', 'SOS', body)
    worksheet216.write('R5', 'JML', body)

    worksheet216.conditional_format(5, 0, row216_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet216.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF RADIO DALAM', title)
    worksheet216.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet216.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet216.write('A22', 'LOKASI', header)
    worksheet216.write('B22', 'TOTAL', header)
    worksheet216.merge_range('A21:B21', 'RANK', header)
    worksheet216.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet216.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet216.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet216.merge_range('F21:F22', 'KELAS', header)
    worksheet216.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet216.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet216.write('G22', 'MAT', body)
    worksheet216.write('H22', 'SEJ', body)
    worksheet216.write('I22', 'GEO', body)
    worksheet216.write('J22', 'EKO', body)
    worksheet216.write('K22', 'SOS', body)
    worksheet216.write('L22', 'JML', body)
    worksheet216.write('M22', 'MAT', body)
    worksheet216.write('N22', 'SEJ', body)
    worksheet216.write('O22', 'GEO', body)
    worksheet216.write('P22', 'EKO', body)
    worksheet216.write('Q22', 'SOS', body)
    worksheet216.write('R22', 'JML', body)

    worksheet216.conditional_format(22, 0, row216+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 217
    worksheet217.insert_image('A1', r'logo resmi nf.jpg')

    worksheet217.set_column('A:A', 7, center)
    worksheet217.set_column('B:B', 6, center)
    worksheet217.set_column('C:C', 18.14, center)
    worksheet217.set_column('D:D', 25, left)
    worksheet217.set_column('E:E', 13.14, left)
    worksheet217.set_column('F:F', 8.57, center)
    worksheet217.set_column('G:R', 5, center)
    worksheet217.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KALIBATA CITY', title)
    worksheet217.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet217.write('A5', 'LOKASI', header)
    worksheet217.write('B5', 'TOTAL', header)
    worksheet217.merge_range('A4:B4', 'RANK', header)
    worksheet217.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet217.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet217.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet217.merge_range('F4:F5', 'KELAS', header)
    worksheet217.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet217.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet217.write('G5', 'MAT', body)
    worksheet217.write('H5', 'SEJ', body)
    worksheet217.write('I5', 'GEO', body)
    worksheet217.write('J5', 'EKO', body)
    worksheet217.write('K5', 'SOS', body)
    worksheet217.write('L5', 'JML', body)
    worksheet217.write('M5', 'MAT', body)
    worksheet217.write('N5', 'SEJ', body)
    worksheet217.write('O5', 'GEO', body)
    worksheet217.write('P5', 'EKO', body)
    worksheet217.write('Q5', 'SOS', body)
    worksheet217.write('R5', 'JML', body)

    worksheet217.conditional_format(5, 0, row217_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet217.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF KALIBATA CITY', title)
    worksheet217.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet217.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet217.write('A22', 'LOKASI', header)
    worksheet217.write('B22', 'TOTAL', header)
    worksheet217.merge_range('A21:B21', 'RANK', header)
    worksheet217.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet217.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet217.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet217.merge_range('F21:F22', 'KELAS', header)
    worksheet217.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet217.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet217.write('G22', 'MAT', body)
    worksheet217.write('H22', 'SEJ', body)
    worksheet217.write('I22', 'GEO', body)
    worksheet217.write('J22', 'EKO', body)
    worksheet217.write('K22', 'SOS', body)
    worksheet217.write('L22', 'JML', body)
    worksheet217.write('M22', 'MAT', body)
    worksheet217.write('N22', 'SEJ', body)
    worksheet217.write('O22', 'GEO', body)
    worksheet217.write('P22', 'EKO', body)
    worksheet217.write('Q22', 'SOS', body)
    worksheet217.write('R22', 'JML', body)

    worksheet217.conditional_format(22, 0, row217+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 218
    worksheet218.insert_image('A1', r'logo resmi nf.jpg')

    worksheet218.set_column('A:A', 7, center)
    worksheet218.set_column('B:B', 6, center)
    worksheet218.set_column('C:C', 18.14, center)
    worksheet218.set_column('D:D', 25, left)
    worksheet218.set_column('E:E', 13.14, left)
    worksheet218.set_column('F:F', 8.57, center)
    worksheet218.set_column('G:R', 5, center)
    worksheet218.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF JAGAKARSA', title)
    worksheet218.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet218.write('A5', 'LOKASI', header)
    worksheet218.write('B5', 'TOTAL', header)
    worksheet218.merge_range('A4:B4', 'RANK', header)
    worksheet218.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet218.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet218.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet218.merge_range('F4:F5', 'KELAS', header)
    worksheet218.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet218.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet218.write('G5', 'MAT', body)
    worksheet218.write('H5', 'SEJ', body)
    worksheet218.write('I5', 'GEO', body)
    worksheet218.write('J5', 'EKO', body)
    worksheet218.write('K5', 'SOS', body)
    worksheet218.write('L5', 'JML', body)
    worksheet218.write('M5', 'MAT', body)
    worksheet218.write('N5', 'SEJ', body)
    worksheet218.write('O5', 'GEO', body)
    worksheet218.write('P5', 'EKO', body)
    worksheet218.write('Q5', 'SOS', body)
    worksheet218.write('R5', 'JML', body)

    worksheet218.conditional_format(5, 0, row218_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet218.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF JAGAKARSA', title)
    worksheet218.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet218.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet218.write('A22', 'LOKASI', header)
    worksheet218.write('B22', 'TOTAL', header)
    worksheet218.merge_range('A21:B21', 'RANK', header)
    worksheet218.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet218.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet218.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet218.merge_range('F21:F22', 'KELAS', header)
    worksheet218.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet218.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet218.write('G22', 'MAT', body)
    worksheet218.write('H22', 'SEJ', body)
    worksheet218.write('I22', 'GEO', body)
    worksheet218.write('J22', 'EKO', body)
    worksheet218.write('K22', 'SOS', body)
    worksheet218.write('L22', 'JML', body)
    worksheet218.write('M22', 'MAT', body)
    worksheet218.write('N22', 'SEJ', body)
    worksheet218.write('O22', 'GEO', body)
    worksheet218.write('P22', 'EKO', body)
    worksheet218.write('Q22', 'SOS', body)
    worksheet218.write('R22', 'JML', body)

    worksheet218.conditional_format(22, 0, row218+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 219
    worksheet219.insert_image('A1', r'logo resmi nf.jpg')

    worksheet219.set_column('A:A', 7, center)
    worksheet219.set_column('B:B', 6, center)
    worksheet219.set_column('C:C', 18.14, center)
    worksheet219.set_column('D:D', 25, left)
    worksheet219.set_column('E:E', 13.14, left)
    worksheet219.set_column('F:F', 8.57, center)
    worksheet219.set_column('G:R', 5, center)
    worksheet219.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF TEBET', title)
    worksheet219.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet219.write('A5', 'LOKASI', header)
    worksheet219.write('B5', 'TOTAL', header)
    worksheet219.merge_range('A4:B4', 'RANK', header)
    worksheet219.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet219.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet219.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet219.merge_range('F4:F5', 'KELAS', header)
    worksheet219.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet219.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet219.write('G5', 'MAT', body)
    worksheet219.write('H5', 'SEJ', body)
    worksheet219.write('I5', 'GEO', body)
    worksheet219.write('J5', 'EKO', body)
    worksheet219.write('K5', 'SOS', body)
    worksheet219.write('L5', 'JML', body)
    worksheet219.write('M5', 'MAT', body)
    worksheet219.write('N5', 'SEJ', body)
    worksheet219.write('O5', 'GEO', body)
    worksheet219.write('P5', 'EKO', body)
    worksheet219.write('Q5', 'SOS', body)
    worksheet219.write('R5', 'JML', body)

    worksheet219.conditional_format(5, 0, row219_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet219.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF TEBET', title)
    worksheet219.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet219.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet219.write('A22', 'LOKASI', header)
    worksheet219.write('B22', 'TOTAL', header)
    worksheet219.merge_range('A21:B21', 'RANK', header)
    worksheet219.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet219.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet219.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet219.merge_range('F21:F22', 'KELAS', header)
    worksheet219.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet219.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet219.write('G22', 'MAT', body)
    worksheet219.write('H22', 'SEJ', body)
    worksheet219.write('I22', 'GEO', body)
    worksheet219.write('J22', 'EKO', body)
    worksheet219.write('K22', 'SOS', body)
    worksheet219.write('L22', 'JML', body)
    worksheet219.write('M22', 'MAT', body)
    worksheet219.write('N22', 'SEJ', body)
    worksheet219.write('O22', 'GEO', body)
    worksheet219.write('P22', 'EKO', body)
    worksheet219.write('Q22', 'SOS', body)
    worksheet219.write('R22', 'JML', body)

    worksheet219.conditional_format(22, 0, row219+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 220
    worksheet220.insert_image('A1', r'logo resmi nf.jpg')

    worksheet220.set_column('A:A', 7, center)
    worksheet220.set_column('B:B', 6, center)
    worksheet220.set_column('C:C', 18.14, center)
    worksheet220.set_column('D:D', 25, left)
    worksheet220.set_column('E:E', 13.14, left)
    worksheet220.set_column('F:F', 8.57, center)
    worksheet220.set_column('G:R', 5, center)
    worksheet220.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PADALARANG', title)
    worksheet220.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet220.write('A5', 'LOKASI', header)
    worksheet220.write('B5', 'TOTAL', header)
    worksheet220.merge_range('A4:B4', 'RANK', header)
    worksheet220.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet220.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet220.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet220.merge_range('F4:F5', 'KELAS', header)
    worksheet220.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet220.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet220.write('G5', 'MAT', body)
    worksheet220.write('H5', 'SEJ', body)
    worksheet220.write('I5', 'GEO', body)
    worksheet220.write('J5', 'EKO', body)
    worksheet220.write('K5', 'SOS', body)
    worksheet220.write('L5', 'JML', body)
    worksheet220.write('M5', 'MAT', body)
    worksheet220.write('N5', 'SEJ', body)
    worksheet220.write('O5', 'GEO', body)
    worksheet220.write('P5', 'EKO', body)
    worksheet220.write('Q5', 'SOS', body)
    worksheet220.write('R5', 'JML', body)

    worksheet220.conditional_format(5, 0, row220_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet220.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PADALARANG', title)
    worksheet220.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet220.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet220.write('A22', 'LOKASI', header)
    worksheet220.write('B22', 'TOTAL', header)
    worksheet220.merge_range('A21:B21', 'RANK', header)
    worksheet220.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet220.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet220.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet220.merge_range('F21:F22', 'KELAS', header)
    worksheet220.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet220.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet220.write('G22', 'MAT', body)
    worksheet220.write('H22', 'SEJ', body)
    worksheet220.write('I22', 'GEO', body)
    worksheet220.write('J22', 'EKO', body)
    worksheet220.write('K22', 'SOS', body)
    worksheet220.write('L22', 'JML', body)
    worksheet220.write('M22', 'MAT', body)
    worksheet220.write('N22', 'SEJ', body)
    worksheet220.write('O22', 'GEO', body)
    worksheet220.write('P22', 'EKO', body)
    worksheet220.write('Q22', 'SOS', body)
    worksheet220.write('R22', 'JML', body)

    worksheet220.conditional_format(22, 0, row220+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 222
    worksheet222.insert_image('A1', r'logo resmi nf.jpg')

    worksheet222.set_column('A:A', 7, center)
    worksheet222.set_column('B:B', 6, center)
    worksheet222.set_column('C:C', 18.14, center)
    worksheet222.set_column('D:D', 25, left)
    worksheet222.set_column('E:E', 13.14, left)
    worksheet222.set_column('F:F', 8.57, center)
    worksheet222.set_column('G:R', 5, center)
    worksheet222.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CILANDAK', title)
    worksheet222.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet222.write('A5', 'LOKASI', header)
    worksheet222.write('B5', 'TOTAL', header)
    worksheet222.merge_range('A4:B4', 'RANK', header)
    worksheet222.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet222.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet222.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet222.merge_range('F4:F5', 'KELAS', header)
    worksheet222.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet222.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet222.write('G5', 'MAT', body)
    worksheet222.write('H5', 'SEJ', body)
    worksheet222.write('I5', 'GEO', body)
    worksheet222.write('J5', 'EKO', body)
    worksheet222.write('K5', 'SOS', body)
    worksheet222.write('L5', 'JML', body)
    worksheet222.write('M5', 'MAT', body)
    worksheet222.write('N5', 'SEJ', body)
    worksheet222.write('O5', 'GEO', body)
    worksheet222.write('P5', 'EKO', body)
    worksheet222.write('Q5', 'SOS', body)
    worksheet222.write('R5', 'JML', body)

    worksheet222.conditional_format(5, 0, row222_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet222.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CILANDAK', title)
    worksheet222.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet222.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet222.write('A22', 'LOKASI', header)
    worksheet222.write('B22', 'TOTAL', header)
    worksheet222.merge_range('A21:B21', 'RANK', header)
    worksheet222.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet222.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet222.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet222.merge_range('F21:F22', 'KELAS', header)
    worksheet222.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet222.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet222.write('G22', 'MAT', body)
    worksheet222.write('H22', 'SEJ', body)
    worksheet222.write('I22', 'GEO', body)
    worksheet222.write('J22', 'EKO', body)
    worksheet222.write('K22', 'SOS', body)
    worksheet222.write('L22', 'JML', body)
    worksheet222.write('M22', 'MAT', body)
    worksheet222.write('N22', 'SEJ', body)
    worksheet222.write('O22', 'GEO', body)
    worksheet222.write('P22', 'EKO', body)
    worksheet222.write('Q22', 'SOS', body)
    worksheet222.write('R22', 'JML', body)

    worksheet222.conditional_format(22, 0, row222+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 226
    worksheet226.insert_image('A1', r'logo resmi nf.jpg')

    worksheet226.set_column('A:A', 7, center)
    worksheet226.set_column('B:B', 6, center)
    worksheet226.set_column('C:C', 18.14, center)
    worksheet226.set_column('D:D', 25, left)
    worksheet226.set_column('E:E', 13.14, left)
    worksheet226.set_column('F:F', 8.57, center)
    worksheet226.set_column('G:R', 5, center)
    worksheet226.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KEBON JERUK', title)
    worksheet226.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet226.write('A5', 'LOKASI', header)
    worksheet226.write('B5', 'TOTAL', header)
    worksheet226.merge_range('A4:B4', 'RANK', header)
    worksheet226.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet226.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet226.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet226.merge_range('F4:F5', 'KELAS', header)
    worksheet226.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet226.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet226.write('G5', 'MAT', body)
    worksheet226.write('H5', 'SEJ', body)
    worksheet226.write('I5', 'GEO', body)
    worksheet226.write('J5', 'EKO', body)
    worksheet226.write('K5', 'SOS', body)
    worksheet226.write('L5', 'JML', body)
    worksheet226.write('M5', 'MAT', body)
    worksheet226.write('N5', 'SEJ', body)
    worksheet226.write('O5', 'GEO', body)
    worksheet226.write('P5', 'EKO', body)
    worksheet226.write('Q5', 'SOS', body)
    worksheet226.write('R5', 'JML', body)

    worksheet226.conditional_format(5, 0, row226_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet226.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF KEBON JERUK', title)
    worksheet226.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet226.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet226.write('A22', 'LOKASI', header)
    worksheet226.write('B22', 'TOTAL', header)
    worksheet226.merge_range('A21:B21', 'RANK', header)
    worksheet226.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet226.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet226.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet226.merge_range('F21:F22', 'KELAS', header)
    worksheet226.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet226.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet226.write('G22', 'MAT', body)
    worksheet226.write('H22', 'SEJ', body)
    worksheet226.write('I22', 'GEO', body)
    worksheet226.write('J22', 'EKO', body)
    worksheet226.write('K22', 'SOS', body)
    worksheet226.write('L22', 'JML', body)
    worksheet226.write('M22', 'MAT', body)
    worksheet226.write('N22', 'SEJ', body)
    worksheet226.write('O22', 'GEO', body)
    worksheet226.write('P22', 'EKO', body)
    worksheet226.write('Q22', 'SOS', body)
    worksheet226.write('R22', 'JML', body)

    worksheet226.conditional_format(22, 0, row226+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 227
    worksheet227.insert_image('A1', r'logo resmi nf.jpg')

    worksheet227.set_column('A:A', 7, center)
    worksheet227.set_column('B:B', 6, center)
    worksheet227.set_column('C:C', 18.14, center)
    worksheet227.set_column('D:D', 25, left)
    worksheet227.set_column('E:E', 13.14, left)
    worksheet227.set_column('F:F', 8.57, center)
    worksheet227.set_column('G:R', 5, center)
    worksheet227.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MERUYA SELATAN', title)
    worksheet227.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet227.write('A5', 'LOKASI', header)
    worksheet227.write('B5', 'TOTAL', header)
    worksheet227.merge_range('A4:B4', 'RANK', header)
    worksheet227.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet227.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet227.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet227.merge_range('F4:F5', 'KELAS', header)
    worksheet227.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet227.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet227.write('G5', 'MAT', body)
    worksheet227.write('H5', 'SEJ', body)
    worksheet227.write('I5', 'GEO', body)
    worksheet227.write('J5', 'EKO', body)
    worksheet227.write('K5', 'SOS', body)
    worksheet227.write('L5', 'JML', body)
    worksheet227.write('M5', 'MAT', body)
    worksheet227.write('N5', 'SEJ', body)
    worksheet227.write('O5', 'GEO', body)
    worksheet227.write('P5', 'EKO', body)
    worksheet227.write('Q5', 'SOS', body)
    worksheet227.write('R5', 'JML', body)

    worksheet227.conditional_format(5, 0, row227_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet227.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF MERUYA SELATAN', title)
    worksheet227.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet227.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet227.write('A22', 'LOKASI', header)
    worksheet227.write('B22', 'TOTAL', header)
    worksheet227.merge_range('A21:B21', 'RANK', header)
    worksheet227.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet227.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet227.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet227.merge_range('F21:F22', 'KELAS', header)
    worksheet227.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet227.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet227.write('G22', 'MAT', body)
    worksheet227.write('H22', 'SEJ', body)
    worksheet227.write('I22', 'GEO', body)
    worksheet227.write('J22', 'EKO', body)
    worksheet227.write('K22', 'SOS', body)
    worksheet227.write('L22', 'JML', body)
    worksheet227.write('M22', 'MAT', body)
    worksheet227.write('N22', 'SEJ', body)
    worksheet227.write('O22', 'GEO', body)
    worksheet227.write('P22', 'EKO', body)
    worksheet227.write('Q22', 'SOS', body)
    worksheet227.write('R22', 'JML', body)

    worksheet227.conditional_format(22, 0, row227+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 228
    worksheet228.insert_image('A1', r'logo resmi nf.jpg')

    worksheet228.set_column('A:A', 7, center)
    worksheet228.set_column('B:B', 6, center)
    worksheet228.set_column('C:C', 18.14, center)
    worksheet228.set_column('D:D', 25, left)
    worksheet228.set_column('E:E', 13.14, left)
    worksheet228.set_column('F:F', 8.57, center)
    worksheet228.set_column('G:R', 5, center)
    worksheet228.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF TANJUNG DUREN', title)
    worksheet228.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet228.write('A5', 'LOKASI', header)
    worksheet228.write('B5', 'TOTAL', header)
    worksheet228.merge_range('A4:B4', 'RANK', header)
    worksheet228.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet228.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet228.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet228.merge_range('F4:F5', 'KELAS', header)
    worksheet228.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet228.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet228.write('G5', 'MAT', body)
    worksheet228.write('H5', 'SEJ', body)
    worksheet228.write('I5', 'GEO', body)
    worksheet228.write('J5', 'EKO', body)
    worksheet228.write('K5', 'SOS', body)
    worksheet228.write('L5', 'JML', body)
    worksheet228.write('M5', 'MAT', body)
    worksheet228.write('N5', 'SEJ', body)
    worksheet228.write('O5', 'GEO', body)
    worksheet228.write('P5', 'EKO', body)
    worksheet228.write('Q5', 'SOS', body)
    worksheet228.write('R5', 'JML', body)

    worksheet228.conditional_format(5, 0, row228_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet228.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF TANJUNG DUREN', title)
    worksheet228.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet228.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet228.write('A22', 'LOKASI', header)
    worksheet228.write('B22', 'TOTAL', header)
    worksheet228.merge_range('A21:B21', 'RANK', header)
    worksheet228.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet228.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet228.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet228.merge_range('F21:F22', 'KELAS', header)
    worksheet228.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet228.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet228.write('G22', 'MAT', body)
    worksheet228.write('H22', 'SEJ', body)
    worksheet228.write('I22', 'GEO', body)
    worksheet228.write('J22', 'EKO', body)
    worksheet228.write('K22', 'SOS', body)
    worksheet228.write('L22', 'JML', body)
    worksheet228.write('M22', 'MAT', body)
    worksheet228.write('N22', 'SEJ', body)
    worksheet228.write('O22', 'GEO', body)
    worksheet228.write('P22', 'EKO', body)
    worksheet228.write('Q22', 'SOS', body)
    worksheet228.write('R22', 'JML', body)

    worksheet228.conditional_format(22, 0, row228+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 229
    worksheet229.insert_image('A1', r'logo resmi nf.jpg')

    worksheet229.set_column('A:A', 7, center)
    worksheet229.set_column('B:B', 6, center)
    worksheet229.set_column('C:C', 18.14, center)
    worksheet229.set_column('D:D', 25, left)
    worksheet229.set_column('E:E', 13.14, left)
    worksheet229.set_column('F:F', 8.57, center)
    worksheet229.set_column('G:R', 5, center)
    worksheet229.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF TOMANG', title)
    worksheet229.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet229.write('A5', 'LOKASI', header)
    worksheet229.write('B5', 'TOTAL', header)
    worksheet229.merge_range('A4:B4', 'RANK', header)
    worksheet229.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet229.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet229.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet229.merge_range('F4:F5', 'KELAS', header)
    worksheet229.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet229.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet229.write('G5', 'MAT', body)
    worksheet229.write('H5', 'SEJ', body)
    worksheet229.write('I5', 'GEO', body)
    worksheet229.write('J5', 'EKO', body)
    worksheet229.write('K5', 'SOS', body)
    worksheet229.write('L5', 'JML', body)
    worksheet229.write('M5', 'MAT', body)
    worksheet229.write('N5', 'SEJ', body)
    worksheet229.write('O5', 'GEO', body)
    worksheet229.write('P5', 'EKO', body)
    worksheet229.write('Q5', 'SOS', body)
    worksheet229.write('R5', 'JML', body)

    worksheet229.conditional_format(5, 0, row229_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet229.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF TOMANG', title)
    worksheet229.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet229.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet229.write('A22', 'LOKASI', header)
    worksheet229.write('B22', 'TOTAL', header)
    worksheet229.merge_range('A21:B21', 'RANK', header)
    worksheet229.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet229.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet229.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet229.merge_range('F21:F22', 'KELAS', header)
    worksheet229.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet229.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet229.write('G22', 'MAT', body)
    worksheet229.write('H22', 'SEJ', body)
    worksheet229.write('I22', 'GEO', body)
    worksheet229.write('J22', 'EKO', body)
    worksheet229.write('K22', 'SOS', body)
    worksheet229.write('L22', 'JML', body)
    worksheet229.write('M22', 'MAT', body)
    worksheet229.write('N22', 'SEJ', body)
    worksheet229.write('O22', 'GEO', body)
    worksheet229.write('P22', 'EKO', body)
    worksheet229.write('Q22', 'SOS', body)
    worksheet229.write('R22', 'JML', body)

    worksheet229.conditional_format(22, 0, row229+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 230
    worksheet230.insert_image('A1', r'logo resmi nf.jpg')

    worksheet230.set_column('A:A', 7, center)
    worksheet230.set_column('B:B', 6, center)
    worksheet230.set_column('C:C', 18.14, center)
    worksheet230.set_column('D:D', 25, left)
    worksheet230.set_column('E:E', 13.14, left)
    worksheet230.set_column('F:F', 8.57, center)
    worksheet230.set_column('G:R', 5, center)
    worksheet230.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KERADENAN', title)
    worksheet230.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet230.write('A5', 'LOKASI', header)
    worksheet230.write('B5', 'TOTAL', header)
    worksheet230.merge_range('A4:B4', 'RANK', header)
    worksheet230.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet230.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet230.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet230.merge_range('F4:F5', 'KELAS', header)
    worksheet230.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet230.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet230.write('G5', 'MAT', body)
    worksheet230.write('H5', 'SEJ', body)
    worksheet230.write('I5', 'GEO', body)
    worksheet230.write('J5', 'EKO', body)
    worksheet230.write('K5', 'SOS', body)
    worksheet230.write('L5', 'JML', body)
    worksheet230.write('M5', 'MAT', body)
    worksheet230.write('N5', 'SEJ', body)
    worksheet230.write('O5', 'GEO', body)
    worksheet230.write('P5', 'EKO', body)
    worksheet230.write('Q5', 'SOS', body)
    worksheet230.write('R5', 'JML', body)

    worksheet230.conditional_format(5, 0, row230_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet230.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF KERADENAN', title)
    worksheet230.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet230.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet230.write('A22', 'LOKASI', header)
    worksheet230.write('B22', 'TOTAL', header)
    worksheet230.merge_range('A21:B21', 'RANK', header)
    worksheet230.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet230.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet230.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet230.merge_range('F21:F22', 'KELAS', header)
    worksheet230.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet230.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet230.write('G22', 'MAT', body)
    worksheet230.write('H22', 'SEJ', body)
    worksheet230.write('I22', 'GEO', body)
    worksheet230.write('J22', 'EKO', body)
    worksheet230.write('K22', 'SOS', body)
    worksheet230.write('L22', 'JML', body)
    worksheet230.write('M22', 'MAT', body)
    worksheet230.write('N22', 'SEJ', body)
    worksheet230.write('O22', 'GEO', body)
    worksheet230.write('P22', 'EKO', body)
    worksheet230.write('Q22', 'SOS', body)
    worksheet230.write('R22', 'JML', body)

    worksheet230.conditional_format(22, 0, row230+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 231
    worksheet231.insert_image('A1', r'logo resmi nf.jpg')

    worksheet231.set_column('A:A', 7, center)
    worksheet231.set_column('B:B', 6, center)
    worksheet231.set_column('C:C', 18.14, center)
    worksheet231.set_column('D:D', 25, left)
    worksheet231.set_column('E:E', 13.14, left)
    worksheet231.set_column('F:F', 8.57, center)
    worksheet231.set_column('G:R', 5, center)
    worksheet231.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF RA KOSASIH SUKABUMI', title)
    worksheet231.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet231.write('A5', 'LOKASI', header)
    worksheet231.write('B5', 'TOTAL', header)
    worksheet231.merge_range('A4:B4', 'RANK', header)
    worksheet231.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet231.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet231.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet231.merge_range('F4:F5', 'KELAS', header)
    worksheet231.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet231.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet231.write('G5', 'MAT', body)
    worksheet231.write('H5', 'SEJ', body)
    worksheet231.write('I5', 'GEO', body)
    worksheet231.write('J5', 'EKO', body)
    worksheet231.write('K5', 'SOS', body)
    worksheet231.write('L5', 'JML', body)
    worksheet231.write('M5', 'MAT', body)
    worksheet231.write('N5', 'SEJ', body)
    worksheet231.write('O5', 'GEO', body)
    worksheet231.write('P5', 'EKO', body)
    worksheet231.write('Q5', 'SOS', body)
    worksheet231.write('R5', 'JML', body)

    worksheet231.conditional_format(5, 0, row231_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet231.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF RA KOSASIH SUKABUMI', title)
    worksheet231.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet231.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet231.write('A22', 'LOKASI', header)
    worksheet231.write('B22', 'TOTAL', header)
    worksheet231.merge_range('A21:B21', 'RANK', header)
    worksheet231.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet231.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet231.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet231.merge_range('F21:F22', 'KELAS', header)
    worksheet231.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet231.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet231.write('G22', 'MAT', body)
    worksheet231.write('H22', 'SEJ', body)
    worksheet231.write('I22', 'GEO', body)
    worksheet231.write('J22', 'EKO', body)
    worksheet231.write('K22', 'SOS', body)
    worksheet231.write('L22', 'JML', body)
    worksheet231.write('M22', 'MAT', body)
    worksheet231.write('N22', 'SEJ', body)
    worksheet231.write('O22', 'GEO', body)
    worksheet231.write('P22', 'EKO', body)
    worksheet231.write('Q22', 'SOS', body)
    worksheet231.write('R22', 'JML', body)

    worksheet231.conditional_format(22, 0, row231+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 233
    worksheet233.insert_image('A1', r'logo resmi nf.jpg')

    worksheet233.set_column('A:A', 7, center)
    worksheet233.set_column('B:B', 6, center)
    worksheet233.set_column('C:C', 18.14, center)
    worksheet233.set_column('D:D', 25, left)
    worksheet233.set_column('E:E', 13.14, left)
    worksheet233.set_column('F:F', 8.57, center)
    worksheet233.set_column('G:R', 5, center)
    worksheet233.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BANGBARUNG', title)
    worksheet233.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet233.write('A5', 'LOKASI', header)
    worksheet233.write('B5', 'TOTAL', header)
    worksheet233.merge_range('A4:B4', 'RANK', header)
    worksheet233.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet233.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet233.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet233.merge_range('F4:F5', 'KELAS', header)
    worksheet233.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet233.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet233.write('G5', 'MAT', body)
    worksheet233.write('H5', 'SEJ', body)
    worksheet233.write('I5', 'GEO', body)
    worksheet233.write('J5', 'EKO', body)
    worksheet233.write('K5', 'SOS', body)
    worksheet233.write('L5', 'JML', body)
    worksheet233.write('M5', 'MAT', body)
    worksheet233.write('N5', 'SEJ', body)
    worksheet233.write('O5', 'GEO', body)
    worksheet233.write('P5', 'EKO', body)
    worksheet233.write('Q5', 'SOS', body)
    worksheet233.write('R5', 'JML', body)

    worksheet233.conditional_format(5, 0, row233_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet233.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF BANGBARUNG', title)
    worksheet233.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet233.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet233.write('A22', 'LOKASI', header)
    worksheet233.write('B22', 'TOTAL', header)
    worksheet233.merge_range('A21:B21', 'RANK', header)
    worksheet233.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet233.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet233.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet233.merge_range('F21:F22', 'KELAS', header)
    worksheet233.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet233.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet233.write('G22', 'MAT', body)
    worksheet233.write('H22', 'SEJ', body)
    worksheet233.write('I22', 'GEO', body)
    worksheet233.write('J22', 'EKO', body)
    worksheet233.write('K22', 'SOS', body)
    worksheet233.write('L22', 'JML', body)
    worksheet233.write('M22', 'MAT', body)
    worksheet233.write('N22', 'SEJ', body)
    worksheet233.write('O22', 'GEO', body)
    worksheet233.write('P22', 'EKO', body)
    worksheet233.write('Q22', 'SOS', body)
    worksheet233.write('R22', 'JML', body)

    worksheet233.conditional_format(22, 0, row233+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 234
    worksheet234.insert_image('A1', r'logo resmi nf.jpg')

    worksheet234.set_column('A:A', 7, center)
    worksheet234.set_column('B:B', 6, center)
    worksheet234.set_column('C:C', 18.14, center)
    worksheet234.set_column('D:D', 25, left)
    worksheet234.set_column('E:E', 13.14, left)
    worksheet234.set_column('F:F', 8.57, center)
    worksheet234.set_column('G:R', 5, center)
    worksheet234.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF LIMUS PRATAMA', title)
    worksheet234.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet234.write('A5', 'LOKASI', header)
    worksheet234.write('B5', 'TOTAL', header)
    worksheet234.merge_range('A4:B4', 'RANK', header)
    worksheet234.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet234.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet234.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet234.merge_range('F4:F5', 'KELAS', header)
    worksheet234.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet234.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet234.write('G5', 'MAT', body)
    worksheet234.write('H5', 'SEJ', body)
    worksheet234.write('I5', 'GEO', body)
    worksheet234.write('J5', 'EKO', body)
    worksheet234.write('K5', 'SOS', body)
    worksheet234.write('L5', 'JML', body)
    worksheet234.write('M5', 'MAT', body)
    worksheet234.write('N5', 'SEJ', body)
    worksheet234.write('O5', 'GEO', body)
    worksheet234.write('P5', 'EKO', body)
    worksheet234.write('Q5', 'SOS', body)
    worksheet234.write('R5', 'JML', body)

    worksheet234.conditional_format(5, 0, row234_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet234.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF LIMUS PRATAMA', title)
    worksheet234.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet234.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet234.write('A22', 'LOKASI', header)
    worksheet234.write('B22', 'TOTAL', header)
    worksheet234.merge_range('A21:B21', 'RANK', header)
    worksheet234.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet234.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet234.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet234.merge_range('F21:F22', 'KELAS', header)
    worksheet234.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet234.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet234.write('G22', 'MAT', body)
    worksheet234.write('H22', 'SEJ', body)
    worksheet234.write('I22', 'GEO', body)
    worksheet234.write('J22', 'EKO', body)
    worksheet234.write('K22', 'SOS', body)
    worksheet234.write('L22', 'JML', body)
    worksheet234.write('M22', 'MAT', body)
    worksheet234.write('N22', 'SEJ', body)
    worksheet234.write('O22', 'GEO', body)
    worksheet234.write('P22', 'EKO', body)
    worksheet234.write('Q22', 'SOS', body)
    worksheet234.write('R22', 'JML', body)

    worksheet234.conditional_format(22, 0, row234+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 235
    worksheet235.insert_image('A1', r'logo resmi nf.jpg')

    worksheet235.set_column('A:A', 7, center)
    worksheet235.set_column('B:B', 6, center)
    worksheet235.set_column('C:C', 18.14, center)
    worksheet235.set_column('D:D', 25, left)
    worksheet235.set_column('E:E', 13.14, left)
    worksheet235.set_column('F:F', 8.57, center)
    worksheet235.set_column('G:R', 5, center)
    worksheet235.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIKARET CIBINONG', title)
    worksheet235.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet235.write('A5', 'LOKASI', header)
    worksheet235.write('B5', 'TOTAL', header)
    worksheet235.merge_range('A4:B4', 'RANK', header)
    worksheet235.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet235.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet235.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet235.merge_range('F4:F5', 'KELAS', header)
    worksheet235.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet235.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet235.write('G5', 'MAT', body)
    worksheet235.write('H5', 'SEJ', body)
    worksheet235.write('I5', 'GEO', body)
    worksheet235.write('J5', 'EKO', body)
    worksheet235.write('K5', 'SOS', body)
    worksheet235.write('L5', 'JML', body)
    worksheet235.write('M5', 'MAT', body)
    worksheet235.write('N5', 'SEJ', body)
    worksheet235.write('O5', 'GEO', body)
    worksheet235.write('P5', 'EKO', body)
    worksheet235.write('Q5', 'SOS', body)
    worksheet235.write('R5', 'JML', body)

    worksheet235.conditional_format(5, 0, row235_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet235.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CIKARET CIBINONG', title)
    worksheet235.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet235.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet235.write('A22', 'LOKASI', header)
    worksheet235.write('B22', 'TOTAL', header)
    worksheet235.merge_range('A21:B21', 'RANK', header)
    worksheet235.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet235.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet235.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet235.merge_range('F21:F22', 'KELAS', header)
    worksheet235.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet235.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet235.write('G22', 'MAT', body)
    worksheet235.write('H22', 'SEJ', body)
    worksheet235.write('I22', 'GEO', body)
    worksheet235.write('J22', 'EKO', body)
    worksheet235.write('K22', 'SOS', body)
    worksheet235.write('L22', 'JML', body)
    worksheet235.write('M22', 'MAT', body)
    worksheet235.write('N22', 'SEJ', body)
    worksheet235.write('O22', 'GEO', body)
    worksheet235.write('P22', 'EKO', body)
    worksheet235.write('Q22', 'SOS', body)
    worksheet235.write('R22', 'JML', body)

    worksheet235.conditional_format(22, 0, row235+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 236
    worksheet236.insert_image('A1', r'logo resmi nf.jpg')

    worksheet236.set_column('A:A', 7, center)
    worksheet236.set_column('B:B', 6, center)
    worksheet236.set_column('C:C', 18.14, center)
    worksheet236.set_column('D:D', 25, left)
    worksheet236.set_column('E:E', 13.14, left)
    worksheet236.set_column('F:F', 8.57, center)
    worksheet236.set_column('G:R', 5, center)
    worksheet236.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF GARUT', title)
    worksheet236.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet236.write('A5', 'LOKASI', header)
    worksheet236.write('B5', 'TOTAL', header)
    worksheet236.merge_range('A4:B4', 'RANK', header)
    worksheet236.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet236.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet236.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet236.merge_range('F4:F5', 'KELAS', header)
    worksheet236.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet236.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet236.write('G5', 'MAT', body)
    worksheet236.write('H5', 'SEJ', body)
    worksheet236.write('I5', 'GEO', body)
    worksheet236.write('J5', 'EKO', body)
    worksheet236.write('K5', 'SOS', body)
    worksheet236.write('L5', 'JML', body)
    worksheet236.write('M5', 'MAT', body)
    worksheet236.write('N5', 'SEJ', body)
    worksheet236.write('O5', 'GEO', body)
    worksheet236.write('P5', 'EKO', body)
    worksheet236.write('Q5', 'SOS', body)
    worksheet236.write('R5', 'JML', body)

    worksheet236.conditional_format(5, 0, row236_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet236.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF GARUT', title)
    worksheet236.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet236.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet236.write('A22', 'LOKASI', header)
    worksheet236.write('B22', 'TOTAL', header)
    worksheet236.merge_range('A21:B21', 'RANK', header)
    worksheet236.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet236.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet236.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet236.merge_range('F21:F22', 'KELAS', header)
    worksheet236.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet236.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet236.write('G22', 'MAT', body)
    worksheet236.write('H22', 'SEJ', body)
    worksheet236.write('I22', 'GEO', body)
    worksheet236.write('J22', 'EKO', body)
    worksheet236.write('K22', 'SOS', body)
    worksheet236.write('L22', 'JML', body)
    worksheet236.write('M22', 'MAT', body)
    worksheet236.write('N22', 'SEJ', body)
    worksheet236.write('O22', 'GEO', body)
    worksheet236.write('P22', 'EKO', body)
    worksheet236.write('Q22', 'SOS', body)
    worksheet236.write('R22', 'JML', body)

    worksheet236.conditional_format(22, 0, row236+21, 17,
                                    {'type': 'no_errors', 'format': border})

    workbook.close()
    st.success("File siap diunduh!")

    # Tombol unduh file
    with open(file_path, "rb") as f:
        bytes_data = f.read()
    st.download_button(label="Unduh File", data=bytes_data,
                       file_name=file_name)


uploaded_file = st.file_uploader(
    'Letakkan file excel NILAI STANDAR [LOKASI 237-299]', type='xlsx')

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # 89
    r = df.shape[0]-5
    # 90
    s = df.shape[0]-4
    # 91
    t = df.shape[0]-3
    # 92
    u = df.shape[0]-2

    # JUMLAH PESERTA
    peserta = df.iloc[r, 23]

    # rata-rata jumlah benar
    rata_mat = df.iloc[r, 156]
    rata_sej = df.iloc[r, 157]
    rata_geo = df.iloc[r, 158]
    rata_eko = df.iloc[r, 159]
    rata_sos = df.iloc[r, 160]
    rata_jml = df.iloc[r, 161]

    # rata-rata nilai standar
    rata_Smat = df.iloc[t, 167]
    rata_Ssej = df.iloc[t, 168]
    rata_Sgeo = df.iloc[t, 169]
    rata_Seko = df.iloc[t, 170]
    rata_Ssos = df.iloc[t, 171]
    rata_Sjml = df.iloc[t, 172]

    max_mat = df.iloc[t, 156]
    max_sej = df.iloc[t, 157]
    max_geo = df.iloc[t, 158]
    max_eko = df.iloc[t, 159]
    max_sos = df.iloc[t, 160]
    max_jml = df.iloc[t, 161]

    # max nilai standar
    max_Smat = df.iloc[r, 167]
    max_Ssej = df.iloc[r, 168]
    max_Sgeo = df.iloc[r, 169]
    max_Seko = df.iloc[r, 170]
    max_Ssos = df.iloc[r, 171]
    max_Sjml = df.iloc[r, 172]

    # min jumlah benar
    min_mat = df.iloc[u, 156]
    min_sej = df.iloc[u, 157]
    min_geo = df.iloc[u, 158]
    min_eko = df.iloc[u, 159]
    min_sos = df.iloc[u, 160]
    min_jml = df.iloc[u, 161]

    # min nilai standar
    min_Smat = df.iloc[s, 167]
    min_Ssej = df.iloc[s, 168]
    min_Sgeo = df.iloc[s, 169]
    min_Seko = df.iloc[s, 170]
    min_Ssos = df.iloc[s, 171]
    min_Sjml = df.iloc[s, 172]

    data_jml_benar = {'BIDANG STUDI': ['MATEMATIKA (MAT)', 'SEJARAH (SEJ)', 'GEOGRAFI (GEO)', 'EKONOMI (EKO)', 'SOSIOLOGI (SOS)', 'JUMLAH (JML)'],
                      'TERENDAH': [min_mat, min_sej, min_geo, min_eko, min_sos, min_jml],
                      'RATA-RATA': [rata_mat, rata_sej, rata_geo, rata_eko, rata_sos, rata_jml],
                      'TERTINGGI': [max_mat, max_sej, max_geo, max_eko, max_sos, max_jml]}

    jml_benar = pd.DataFrame(data_jml_benar)

    data_n_standar = {'BIDANG STUDI': ['MATEMATIKA (MAT)', 'SEJARAH (SEJ)', 'GEOGRAFI (GEO)', 'EKONOMI (EKO)', 'SOSIOLOGI (SOS)', 'JUMLAH (JML)'],
                      'TERENDAH': [min_Smat, min_Ssej, min_Sgeo, min_Seko, min_Ssos, min_Sjml],
                      'RATA-RATA': [rata_Smat, rata_Ssej, rata_Sgeo, rata_Seko, rata_Ssos, rata_Sjml],
                      'TERTINGGI': [max_Smat, max_Ssej, max_Sgeo, max_Seko, max_Ssos, max_Sjml]}

    n_standar = pd.DataFrame(data_n_standar)

    data_jml_peserta = {'JUMLAH PESERTA': [peserta]}

    jml_peserta = pd.DataFrame(data_jml_peserta)

    data_jml_soal = {'BIDANG STUDI': ['MAT', 'SEJ', 'GEO', 'EKO', 'SOS'],
                     'JUMLAH': [JML_SOAL_MAT, JML_SOAL_SEJ, JML_SOAL_GEO, JML_SOAL_EKO, JML_SOAL_SOS]}

    jml_soal = pd.DataFrame(data_jml_soal)

    df = df[['LOKASI', 'RANK LOK.', 'RANK NAS.', 'NOMOR NF', 'NAMA SISWA', 'NAMA SEKOLAH', 'KELAS',
             'MAT', 'SEJ', 'GEO', 'EKO', 'SOS', 'JML', 'S_MAT', 'S_SEJ', 'S_GEO', 'S_EKO', 'S_SOS', 'S_JML']]

    # sort setiap lokasi
    sort237 = df[df['LOKASI'] == 237]
    sort238 = df[df['LOKASI'] == 238]
    sort240 = df[df['LOKASI'] == 240]
    sort241 = df[df['LOKASI'] == 241]
    sort243 = df[df['LOKASI'] == 243]
    sort244 = df[df['LOKASI'] == 244]
    sort245 = df[df['LOKASI'] == 245]
    sort246 = df[df['LOKASI'] == 246]
    sort247 = df[df['LOKASI'] == 247]
    sort248 = df[df['LOKASI'] == 248]
    sort249 = df[df['LOKASI'] == 249]
    sort250 = df[df['LOKASI'] == 250]
    sort251 = df[df['LOKASI'] == 251]
    sort252 = df[df['LOKASI'] == 252]
    sort253 = df[df['LOKASI'] == 253]
    sort254 = df[df['LOKASI'] == 254]
    sort255 = df[df['LOKASI'] == 255]
    sort256 = df[df['LOKASI'] == 256]
    sort258 = df[df['LOKASI'] == 258]
    sort259 = df[df['LOKASI'] == 259]
    sort260 = df[df['LOKASI'] == 260]
    sort261 = df[df['LOKASI'] == 261]
    sort262 = df[df['LOKASI'] == 262]
    sort264 = df[df['LOKASI'] == 264]
    sort265 = df[df['LOKASI'] == 265]
    sort266 = df[df['LOKASI'] == 266]
    sort267 = df[df['LOKASI'] == 267]
    sort268 = df[df['LOKASI'] == 268]
    sort269 = df[df['LOKASI'] == 269]
    sort270 = df[df['LOKASI'] == 270]
    sort271 = df[df['LOKASI'] == 271]
    sort272 = df[df['LOKASI'] == 272]
    sort273 = df[df['LOKASI'] == 273]
    sort274 = df[df['LOKASI'] == 274]
    sort275 = df[df['LOKASI'] == 275]
    sort276 = df[df['LOKASI'] == 276]
    sort277 = df[df['LOKASI'] == 277]
    sort278 = df[df['LOKASI'] == 278]
    sort279 = df[df['LOKASI'] == 279]
    sort280 = df[df['LOKASI'] == 280]
    sort281 = df[df['LOKASI'] == 281]
    sort282 = df[df['LOKASI'] == 282]
    sort283 = df[df['LOKASI'] == 283]
    sort284 = df[df['LOKASI'] == 284]
    sort285 = df[df['LOKASI'] == 285]
    sort286 = df[df['LOKASI'] == 286]
    sort287 = df[df['LOKASI'] == 287]
    sort288 = df[df['LOKASI'] == 288]
    sort289 = df[df['LOKASI'] == 289]
    sort290 = df[df['LOKASI'] == 290]
    sort291 = df[df['LOKASI'] == 291]
    sort292 = df[df['LOKASI'] == 292]
    sort293 = df[df['LOKASI'] == 293]
    sort294 = df[df['LOKASI'] == 294]
    sort295 = df[df['LOKASI'] == 295]
    sort296 = df[df['LOKASI'] == 296]
    sort297 = df[df['LOKASI'] == 297]
    sort298 = df[df['LOKASI'] == 298]
    sort299 = df[df['LOKASI'] == 299]

    # 10 besar setiap lokasi
    # 237
    sort237_10 = sort237.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort237_10['LOKASI']
    sort237_10 = sort237_10.drop(
        sort237_10[(sort237_10['RANK LOK.'] > 10)].index)
    # 238
    sort238_10 = sort238.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort238_10['LOKASI']
    sort238_10 = sort238_10.drop(
        sort238_10[(sort238_10['RANK LOK.'] > 10)].index)
    # 240
    sort240_10 = sort240.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort240_10['LOKASI']
    sort240_10 = sort240_10.drop(
        sort240_10[(sort240_10['RANK LOK.'] > 10)].index)
    # 241
    sort241_10 = sort241.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort241_10['LOKASI']
    sort241_10 = sort241_10.drop(
        sort241_10[(sort241_10['RANK LOK.'] > 10)].index)
    # 243
    sort243_10 = sort243.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort243_10['LOKASI']
    sort243_10 = sort243_10.drop(
        sort243_10[(sort243_10['RANK LOK.'] > 10)].index)
    # 244
    sort244_10 = sort244.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort244_10['LOKASI']
    sort244_10 = sort244_10.drop(
        sort244_10[(sort244_10['RANK LOK.'] > 10)].index)
    # 245
    sort245_10 = sort245.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort245_10['LOKASI']
    sort245_10 = sort245_10.drop(
        sort245_10[(sort245_10['RANK LOK.'] > 10)].index)
    # 246
    sort246_10 = sort246.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort246_10['LOKASI']
    sort246_10 = sort246_10.drop(
        sort246_10[(sort246_10['RANK LOK.'] > 10)].index)
    # 247
    sort247_10 = sort247.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort247_10['LOKASI']
    sort247_10 = sort247_10.drop(
        sort247_10[(sort247_10['RANK LOK.'] > 10)].index)
    # 248
    sort248_10 = sort248.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort248_10['LOKASI']
    sort248_10 = sort248_10.drop(
        sort248_10[(sort248_10['RANK LOK.'] > 10)].index)
    # 249
    sort249_10 = sort249.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort249_10['LOKASI']
    sort249_10 = sort249_10.drop(
        sort249_10[(sort249_10['RANK LOK.'] > 10)].index)
    # 250
    sort250_10 = sort250.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort250_10['LOKASI']
    sort250_10 = sort250_10.drop(
        sort250_10[(sort250_10['RANK LOK.'] > 10)].index)
    # 251
    sort251_10 = sort251.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort251_10['LOKASI']
    sort251_10 = sort251_10.drop(
        sort251_10[(sort251_10['RANK LOK.'] > 10)].index)
    # 252
    sort252_10 = sort252.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort252_10['LOKASI']
    sort252_10 = sort252_10.drop(
        sort252_10[(sort252_10['RANK LOK.'] > 10)].index)
    # 253
    sort253_10 = sort253.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort253_10['LOKASI']
    sort253_10 = sort253_10.drop(
        sort253_10[(sort253_10['RANK LOK.'] > 10)].index)
    # 254
    sort254_10 = sort254.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort254_10['LOKASI']
    sort254_10 = sort254_10.drop(
        sort254_10[(sort254_10['RANK LOK.'] > 10)].index)
    # 255
    sort255_10 = sort255.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort255_10['LOKASI']
    sort255_10 = sort255_10.drop(
        sort255_10[(sort255_10['RANK LOK.'] > 10)].index)
    # 256
    sort256_10 = sort256.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort256_10['LOKASI']
    sort256_10 = sort256_10.drop(
        sort256_10[(sort256_10['RANK LOK.'] > 10)].index)
    # 258
    sort258_10 = sort258.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort258_10['LOKASI']
    sort258_10 = sort258_10.drop(
        sort258_10[(sort258_10['RANK LOK.'] > 10)].index)
    # 259
    sort259_10 = sort259.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort259_10['LOKASI']
    sort259_10 = sort259_10.drop(
        sort259_10[(sort259_10['RANK LOK.'] > 10)].index)
    # 260
    sort260_10 = sort260.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort260_10['LOKASI']
    sort260_10 = sort260_10.drop(
        sort260_10[(sort260_10['RANK LOK.'] > 10)].index)
    # 261
    sort261_10 = sort261.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort261_10['LOKASI']
    sort261_10 = sort261_10.drop(
        sort261_10[(sort261_10['RANK LOK.'] > 10)].index)
    # 262
    sort262_10 = sort262.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort262_10['LOKASI']
    sort262_10 = sort262_10.drop(
        sort262_10[(sort262_10['RANK LOK.'] > 10)].index)
    # 264
    sort264_10 = sort264.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort264_10['LOKASI']
    sort264_10 = sort264_10.drop(
        sort264_10[(sort264_10['RANK LOK.'] > 10)].index)
    # 265
    sort265_10 = sort265.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort265_10['LOKASI']
    sort265_10 = sort265_10.drop(
        sort265_10[(sort265_10['RANK LOK.'] > 10)].index)
    # 266
    sort266_10 = sort266.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort266_10['LOKASI']
    sort266_10 = sort266_10.drop(
        sort266_10[(sort266_10['RANK LOK.'] > 10)].index)
    # 267
    sort267_10 = sort267.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort267_10['LOKASI']
    sort267_10 = sort267_10.drop(
        sort267_10[(sort267_10['RANK LOK.'] > 10)].index)
    # 268
    sort268_10 = sort268.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort268_10['LOKASI']
    sort268_10 = sort268_10.drop(
        sort268_10[(sort268_10['RANK LOK.'] > 10)].index)
    # 269
    sort269_10 = sort269.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort269_10['LOKASI']
    sort269_10 = sort269_10.drop(
        sort269_10[(sort269_10['RANK LOK.'] > 10)].index)
    # 270
    sort270_10 = sort270.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort270_10['LOKASI']
    sort270_10 = sort270_10.drop(
        sort270_10[(sort270_10['RANK LOK.'] > 10)].index)
    # 271
    sort271_10 = sort271.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort271_10['LOKASI']
    sort271_10 = sort271_10.drop(
        sort271_10[(sort271_10['RANK LOK.'] > 10)].index)
    # 272
    sort272_10 = sort272.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort272_10['LOKASI']
    sort272_10 = sort272_10.drop(
        sort272_10[(sort272_10['RANK LOK.'] > 10)].index)
    # 273
    sort273_10 = sort273.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort273_10['LOKASI']
    sort273_10 = sort273_10.drop(
        sort273_10[(sort273_10['RANK LOK.'] > 10)].index)
    # 274
    sort274_10 = sort274.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort274_10['LOKASI']
    sort274_10 = sort274_10.drop(
        sort274_10[(sort274_10['RANK LOK.'] > 10)].index)
    # 275
    sort275_10 = sort275.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort275_10['LOKASI']
    sort275_10 = sort275_10.drop(
        sort275_10[(sort275_10['RANK LOK.'] > 10)].index)
    # 276
    sort276_10 = sort276.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort276_10['LOKASI']
    sort276_10 = sort276_10.drop(
        sort276_10[(sort276_10['RANK LOK.'] > 10)].index)
    # 277
    sort277_10 = sort277.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort277_10['LOKASI']
    sort277_10 = sort277_10.drop(
        sort277_10[(sort277_10['RANK LOK.'] > 10)].index)
    # 278
    sort278_10 = sort278.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort278_10['LOKASI']
    sort278_10 = sort278_10.drop(
        sort278_10[(sort278_10['RANK LOK.'] > 10)].index)
    # 279
    sort279_10 = sort279.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort279_10['LOKASI']
    sort279_10 = sort279_10.drop(
        sort279_10[(sort279_10['RANK LOK.'] > 10)].index)
    # 280
    sort280_10 = sort280.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort280_10['LOKASI']
    sort280_10 = sort280_10.drop(
        sort280_10[(sort280_10['RANK LOK.'] > 10)].index)
    # 281
    sort281_10 = sort281.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort281_10['LOKASI']
    sort281_10 = sort281_10.drop(
        sort281_10[(sort281_10['RANK LOK.'] > 10)].index)
    # 282
    sort282_10 = sort282.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort282_10['LOKASI']
    sort282_10 = sort282_10.drop(
        sort282_10[(sort282_10['RANK LOK.'] > 10)].index)
    # 283
    sort283_10 = sort283.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort283_10['LOKASI']
    sort283_10 = sort283_10.drop(
        sort283_10[(sort283_10['RANK LOK.'] > 10)].index)
    # 284
    sort284_10 = sort284.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort284_10['LOKASI']
    sort284_10 = sort284_10.drop(
        sort284_10[(sort284_10['RANK LOK.'] > 10)].index)
    # 285
    sort285_10 = sort285.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort285_10['LOKASI']
    sort285_10 = sort285_10.drop(
        sort285_10[(sort285_10['RANK LOK.'] > 10)].index)
    # 286
    sort286_10 = sort286.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort286_10['LOKASI']
    sort286_10 = sort286_10.drop(
        sort286_10[(sort286_10['RANK LOK.'] > 10)].index)
    # 287
    sort287_10 = sort287.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort287_10['LOKASI']
    sort287_10 = sort287_10.drop(
        sort287_10[(sort287_10['RANK LOK.'] > 10)].index)
    # 288
    sort288_10 = sort288.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort288_10['LOKASI']
    sort288_10 = sort288_10.drop(
        sort288_10[(sort288_10['RANK LOK.'] > 10)].index)
    # 289
    sort289_10 = sort289.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort289_10['LOKASI']
    sort289_10 = sort289_10.drop(
        sort289_10[(sort289_10['RANK LOK.'] > 10)].index)
    # 290
    sort290_10 = sort290.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort290_10['LOKASI']
    sort290_10 = sort290_10.drop(
        sort290_10[(sort290_10['RANK LOK.'] > 10)].index)
    # 291
    sort291_10 = sort291.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort291_10['LOKASI']
    sort291_10 = sort291_10.drop(
        sort291_10[(sort291_10['RANK LOK.'] > 10)].index)
    # 292
    sort292_10 = sort292.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort292_10['LOKASI']
    sort292_10 = sort292_10.drop(
        sort292_10[(sort292_10['RANK LOK.'] > 10)].index)
    # 293
    sort293_10 = sort293.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort293_10['LOKASI']
    sort293_10 = sort293_10.drop(
        sort293_10[(sort293_10['RANK LOK.'] > 10)].index)
    # 294
    sort294_10 = sort294.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort294_10['LOKASI']
    sort294_10 = sort294_10.drop(
        sort294_10[(sort294_10['RANK LOK.'] > 10)].index)
    # 295
    sort295_10 = sort295.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort295_10['LOKASI']
    sort295_10 = sort295_10.drop(
        sort295_10[(sort295_10['RANK LOK.'] > 10)].index)
    # 296
    sort296_10 = sort296.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort296_10['LOKASI']
    sort296_10 = sort296_10.drop(
        sort296_10[(sort296_10['RANK LOK.'] > 10)].index)
    # 297
    sort297_10 = sort297.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort297_10['LOKASI']
    sort297_10 = sort297_10.drop(
        sort297_10[(sort297_10['RANK LOK.'] > 10)].index)
    # 298
    sort298_10 = sort298.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort298_10['LOKASI']
    sort298_10 = sort298_10.drop(
        sort298_10[(sort298_10['RANK LOK.'] > 10)].index)
    # 299
    sort299_10 = sort299.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort299_10['LOKASI']
    sort299_10 = sort299_10.drop(
        sort299_10[(sort299_10['RANK LOK.'] > 10)].index)

    # All 237
    sort237 = sort237.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort237['LOKASI']
    # All 238
    sort238 = sort238.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort238['LOKASI']
    # All 240
    sort240 = sort240.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort240['LOKASI']
    # All 241
    sort241 = sort241.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort241['LOKASI']
    # All 243
    sort243 = sort243.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort243['LOKASI']
    # All 244
    sort244 = sort244.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort244['LOKASI']
    # All 245
    sort245 = sort245.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort245['LOKASI']
    # All 246
    sort246 = sort246.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort246['LOKASI']
    # All 247
    sort247 = sort247.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort247['LOKASI']
    # All 248
    sort248 = sort248.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort248['LOKASI']
    # All 249
    sort249 = sort249.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort249['LOKASI']
    # All 250
    sort250 = sort250.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort250['LOKASI']
    # All 251
    sort251 = sort251.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort251['LOKASI']
    # All 252
    sort252 = sort252.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort252['LOKASI']
    # All 253
    sort253 = sort253.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort253['LOKASI']
    # All 254
    sort254 = sort254.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort254['LOKASI']
    # All 255
    sort255 = sort255.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort255['LOKASI']
    # All 256
    sort256 = sort256.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort256['LOKASI']
    # All 258
    sort258 = sort258.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort258['LOKASI']
    # All 259
    sort259 = sort259.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort259['LOKASI']
    # All 260
    sort260 = sort260.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort260['LOKASI']
    # All 261
    sort261 = sort261.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort261['LOKASI']
    # All 262
    sort262 = sort262.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort262['LOKASI']
    # All 264
    sort264 = sort264.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort264['LOKASI']
    # All 265
    sort265 = sort265.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort265['LOKASI']
    # All 266
    sort266 = sort266.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort266['LOKASI']
    # All 267
    sort267 = sort267.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort267['LOKASI']
    # All 268
    sort268 = sort268.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort268['LOKASI']
    # All 269
    sort269 = sort269.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort269['LOKASI']
    # All 270
    sort270 = sort270.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort270['LOKASI']
    # All 271
    sort271 = sort271.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort271['LOKASI']
    # All 272
    sort272 = sort272.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort272['LOKASI']
    # All 273
    sort273 = sort273.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort273['LOKASI']
    # All 274
    sort274 = sort274.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort274['LOKASI']
    # All 275
    sort275 = sort275.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort275['LOKASI']
    # All 276
    sort276 = sort276.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort276['LOKASI']
    # All 277
    sort277 = sort277.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort277['LOKASI']
    # All 278
    sort278 = sort278.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort278['LOKASI']
    # All 279
    sort279 = sort279.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort279['LOKASI']
    # All 280
    sort280 = sort280.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort280['LOKASI']
    # All 281
    sort281 = sort281.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort281['LOKASI']
    # All 282
    sort282 = sort282.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort282['LOKASI']
    # All 283
    sort283 = sort283.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort283['LOKASI']
    # All 284
    sort284 = sort284.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort284['LOKASI']
    # All 285
    sort285 = sort285.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort285['LOKASI']
    # All 286
    sort286 = sort286.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort286['LOKASI']
    # All 287
    sort287 = sort287.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort287['LOKASI']
    # All 288
    sort288 = sort288.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort288['LOKASI']
    # All 289
    sort289 = sort289.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort289['LOKASI']
    # All 290
    sort290 = sort290.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort290['LOKASI']
    # All 291
    sort291 = sort291.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort291['LOKASI']
    # All 292
    sort292 = sort292.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort292['LOKASI']
    # All 293
    sort293 = sort293.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort293['LOKASI']
    # All 294
    sort294 = sort294.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort294['LOKASI']
    # All 295
    sort295 = sort295.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort295['LOKASI']
    # All 296
    sort296 = sort296.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort296['LOKASI']
    # All 297
    sort297 = sort297.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort297['LOKASI']
    # All 298
    sort298 = sort298.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort298['LOKASI']
    # All 299
    sort299 = sort299.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort299['LOKASI']

    # jumlah row
    # 237
    row237_10 = sort237_10.shape[0]
    row237 = sort237.shape[0]
    # 238
    row238_10 = sort238_10.shape[0]
    row238 = sort238.shape[0]
    # 240
    row240_10 = sort240_10.shape[0]
    row240 = sort240.shape[0]
    # 241
    row241_10 = sort241_10.shape[0]
    row241 = sort241.shape[0]
    # 243
    row243_10 = sort243_10.shape[0]
    row243 = sort243.shape[0]
    # 244
    row244_10 = sort244_10.shape[0]
    row244 = sort244.shape[0]
    # 245
    row245_10 = sort245_10.shape[0]
    row245 = sort245.shape[0]
    # 246
    row246_10 = sort246_10.shape[0]
    row246 = sort246.shape[0]
    # 247
    row247_10 = sort247_10.shape[0]
    row247 = sort247.shape[0]
    # 248
    row248_10 = sort248_10.shape[0]
    row248 = sort248.shape[0]
    # 249
    row249_10 = sort249_10.shape[0]
    row249 = sort249.shape[0]
    # 250
    row250_10 = sort250_10.shape[0]
    row250 = sort250.shape[0]
    # 251
    row251_10 = sort251_10.shape[0]
    row251 = sort251.shape[0]
    # 252
    row252_10 = sort252_10.shape[0]
    row252 = sort252.shape[0]
    # 253
    row253_10 = sort253_10.shape[0]
    row253 = sort253.shape[0]
    # 254
    row254_10 = sort254_10.shape[0]
    row254 = sort254.shape[0]
    # 255
    row255_10 = sort255_10.shape[0]
    row255 = sort255.shape[0]
    # 256
    row256_10 = sort256_10.shape[0]
    row256 = sort256.shape[0]
    # 258
    row258_10 = sort258_10.shape[0]
    row258 = sort258.shape[0]
    # 259
    row259_10 = sort259_10.shape[0]
    row259 = sort259.shape[0]
    # 260
    row260_10 = sort260_10.shape[0]
    row260 = sort260.shape[0]
    # 261
    row261_10 = sort261_10.shape[0]
    row261 = sort261.shape[0]
    # 262
    row262_10 = sort262_10.shape[0]
    row262 = sort262.shape[0]
    # 264
    row264_10 = sort264_10.shape[0]
    row264 = sort264.shape[0]
    # 265
    row265_10 = sort265_10.shape[0]
    row265 = sort265.shape[0]
    # 266
    row266_10 = sort266_10.shape[0]
    row266 = sort266.shape[0]
    # 267
    row267_10 = sort267_10.shape[0]
    row267 = sort267.shape[0]
    # 268
    row268_10 = sort268_10.shape[0]
    row268 = sort268.shape[0]
    # 269
    row269_10 = sort269_10.shape[0]
    row269 = sort269.shape[0]
    # 270
    row270_10 = sort270_10.shape[0]
    row270 = sort270.shape[0]
    # 271
    row271_10 = sort271_10.shape[0]
    row271 = sort271.shape[0]
    # 272
    row272_10 = sort272_10.shape[0]
    row272 = sort272.shape[0]
    # 273
    row273_10 = sort273_10.shape[0]
    row273 = sort273.shape[0]
    # 274
    row274_10 = sort274_10.shape[0]
    row274 = sort274.shape[0]
    # 275
    row275_10 = sort275_10.shape[0]
    row275 = sort275.shape[0]
    # 276
    row276_10 = sort276_10.shape[0]
    row276 = sort276.shape[0]
    # 277
    row277_10 = sort277_10.shape[0]
    row277 = sort277.shape[0]
    # 278
    row278_10 = sort278_10.shape[0]
    row278 = sort278.shape[0]
    # 279
    row279_10 = sort279_10.shape[0]
    row279 = sort279.shape[0]
    # 280
    row280_10 = sort280_10.shape[0]
    row280 = sort280.shape[0]
    # 281
    row281_10 = sort281_10.shape[0]
    row281 = sort281.shape[0]
    # 282
    row282_10 = sort282_10.shape[0]
    row282 = sort282.shape[0]
    # 283
    row283_10 = sort283_10.shape[0]
    row283 = sort283.shape[0]
    # 284
    row284_10 = sort284_10.shape[0]
    row284 = sort284.shape[0]
    # 285
    row285_10 = sort285_10.shape[0]
    row285 = sort285.shape[0]
    # 286
    row286_10 = sort286_10.shape[0]
    row286 = sort286.shape[0]
    # 287
    row287_10 = sort287_10.shape[0]
    row287 = sort287.shape[0]
    # 288
    row288_10 = sort288_10.shape[0]
    row288 = sort288.shape[0]
    # 289
    row289_10 = sort289_10.shape[0]
    row289 = sort289.shape[0]
    # 290
    row290_10 = sort290_10.shape[0]
    row290 = sort290.shape[0]
    # 291
    row291_10 = sort291_10.shape[0]
    row291 = sort291.shape[0]
    # 292
    row292_10 = sort292_10.shape[0]
    row292 = sort292.shape[0]
    # 293
    row293_10 = sort293_10.shape[0]
    row293 = sort293.shape[0]
    # 294
    row294_10 = sort294_10.shape[0]
    row294 = sort294.shape[0]
    # 295
    row295_10 = sort295_10.shape[0]
    row295 = sort295.shape[0]
    # 296
    row296_10 = sort296_10.shape[0]
    row296 = sort296.shape[0]
    # 297
    row297_10 = sort297_10.shape[0]
    row297 = sort297.shape[0]
    # 298
    row298_10 = sort298_10.shape[0]
    row298 = sort298.shape[0]
    # 299
    row299_10 = sort299_10.shape[0]
    row299 = sort299.shape[0]

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    # Path file hasil penyimpanan
    file_name = f"{kelas}_{penilaian}_{semester}_lokasi_237_299.xlsx"
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
                       startrow=21,
                       startcol=0,
                       index=False,
                       header=False)

    # Convert the dataframe to an XlsxWriter Excel object cover.
    jml_peserta.to_excel(writer, sheet_name='cover',
                         startrow=21,
                         startcol=5,
                         index=False,
                         header=False)

    # Convert the dataframe to an XlsxWriter Excel object cover.
    jml_soal.to_excel(writer, sheet_name='cover',
                      startrow=13,
                      startcol=5,
                      index=False,
                      header=False)

    # 237
    # Convert the dataframe to an XlsxWriter Excel object.
    sort237_10.to_excel(writer, sheet_name='237',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort237.to_excel(writer, sheet_name='237',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 238
    # Convert the dataframe to an XlsxWriter Excel object.
    sort238_10.to_excel(writer, sheet_name='238',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort238.to_excel(writer, sheet_name='238',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 240
    # Convert the dataframe to an XlsxWriter Excel object.
    sort240_10.to_excel(writer, sheet_name='240',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort240.to_excel(writer, sheet_name='240',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 241
    # Convert the dataframe to an XlsxWriter Excel object.
    sort241_10.to_excel(writer, sheet_name='241',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort241.to_excel(writer, sheet_name='241',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 243
    # Convert the dataframe to an XlsxWriter Excel object.
    sort243_10.to_excel(writer, sheet_name='243',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort243.to_excel(writer, sheet_name='243',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 244
    # Convert the dataframe to an XlsxWriter Excel object.
    sort244_10.to_excel(writer, sheet_name='244',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort244.to_excel(writer, sheet_name='244',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 245
    # Convert the dataframe to an XlsxWriter Excel object.
    sort245_10.to_excel(writer, sheet_name='245',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort245.to_excel(writer, sheet_name='245',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 246
    # Convert the dataframe to an XlsxWriter Excel object.
    sort246_10.to_excel(writer, sheet_name='246',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort246.to_excel(writer, sheet_name='246',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 247
    # Convert the dataframe to an XlsxWriter Excel object.
    sort247_10.to_excel(writer, sheet_name='247',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort247.to_excel(writer, sheet_name='247',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 248
    # Convert the dataframe to an XlsxWriter Excel object.
    sort248_10.to_excel(writer, sheet_name='248',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort248.to_excel(writer, sheet_name='248',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 249
    # Convert the dataframe to an XlsxWriter Excel object.
    sort249_10.to_excel(writer, sheet_name='249',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort249.to_excel(writer, sheet_name='249',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 250
    # Convert the dataframe to an XlsxWriter Excel object.
    sort250_10.to_excel(writer, sheet_name='250',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort250.to_excel(writer, sheet_name='250',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 251
    # Convert the dataframe to an XlsxWriter Excel object.
    sort251_10.to_excel(writer, sheet_name='251',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort251.to_excel(writer, sheet_name='251',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 252
    # Convert the dataframe to an XlsxWriter Excel object.
    sort252_10.to_excel(writer, sheet_name='252',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort252.to_excel(writer, sheet_name='252',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 253
    # Convert the dataframe to an XlsxWriter Excel object.
    sort253_10.to_excel(writer, sheet_name='253',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort253.to_excel(writer, sheet_name='253',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 254
    # Convert the dataframe to an XlsxWriter Excel object.
    sort254_10.to_excel(writer, sheet_name='254',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort254.to_excel(writer, sheet_name='254',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 255
    # Convert the dataframe to an XlsxWriter Excel object.
    sort255_10.to_excel(writer, sheet_name='255',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort255.to_excel(writer, sheet_name='255',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 256
    # Convert the dataframe to an XlsxWriter Excel object.
    sort256_10.to_excel(writer, sheet_name='256',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort256.to_excel(writer, sheet_name='256',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 258
    # Convert the dataframe to an XlsxWriter Excel object.
    sort258_10.to_excel(writer, sheet_name='258',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort258.to_excel(writer, sheet_name='258',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 259
    # Convert the dataframe to an XlsxWriter Excel object.
    sort259_10.to_excel(writer, sheet_name='259',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort259.to_excel(writer, sheet_name='259',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 260
    # Convert the dataframe to an XlsxWriter Excel object.
    sort260_10.to_excel(writer, sheet_name='260',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort260.to_excel(writer, sheet_name='260',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 261
    # Convert the dataframe to an XlsxWriter Excel object.
    sort261_10.to_excel(writer, sheet_name='261',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort261.to_excel(writer, sheet_name='261',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 262
    # Convert the dataframe to an XlsxWriter Excel object.
    sort262_10.to_excel(writer, sheet_name='262',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort262.to_excel(writer, sheet_name='262',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 264
    # Convert the dataframe to an XlsxWriter Excel object.
    sort264_10.to_excel(writer, sheet_name='264',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort264.to_excel(writer, sheet_name='264',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 265
    # Convert the dataframe to an XlsxWriter Excel object.
    sort265_10.to_excel(writer, sheet_name='265',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort265.to_excel(writer, sheet_name='265',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 266
    # Convert the dataframe to an XlsxWriter Excel object.
    sort266_10.to_excel(writer, sheet_name='266',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort266.to_excel(writer, sheet_name='266',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 267
    # Convert the dataframe to an XlsxWriter Excel object.
    sort267_10.to_excel(writer, sheet_name='267',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort267.to_excel(writer, sheet_name='267',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 268
    # Convert the dataframe to an XlsxWriter Excel object.
    sort268_10.to_excel(writer, sheet_name='268',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort268.to_excel(writer, sheet_name='268',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 269
    # Convert the dataframe to an XlsxWriter Excel object.
    sort269_10.to_excel(writer, sheet_name='269',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort269.to_excel(writer, sheet_name='269',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 270
    # Convert the dataframe to an XlsxWriter Excel object.
    sort270_10.to_excel(writer, sheet_name='270',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort270.to_excel(writer, sheet_name='270',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 271
    # Convert the dataframe to an XlsxWriter Excel object.
    sort271_10.to_excel(writer, sheet_name='271',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort271.to_excel(writer, sheet_name='271',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 272
    # Convert the dataframe to an XlsxWriter Excel object.
    sort272_10.to_excel(writer, sheet_name='272',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort272.to_excel(writer, sheet_name='272',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 273
    # Convert the dataframe to an XlsxWriter Excel object.
    sort273_10.to_excel(writer, sheet_name='273',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort273.to_excel(writer, sheet_name='273',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 274
    # Convert the dataframe to an XlsxWriter Excel object.
    sort274_10.to_excel(writer, sheet_name='274',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort274.to_excel(writer, sheet_name='274',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 275
    # Convert the dataframe to an XlsxWriter Excel object.
    sort275_10.to_excel(writer, sheet_name='275',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort275.to_excel(writer, sheet_name='275',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 276
    # Convert the dataframe to an XlsxWriter Excel object.
    sort276_10.to_excel(writer, sheet_name='276',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort276.to_excel(writer, sheet_name='276',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 277
    # Convert the dataframe to an XlsxWriter Excel object.
    sort277_10.to_excel(writer, sheet_name='277',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort277.to_excel(writer, sheet_name='277',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 278
    # Convert the dataframe to an XlsxWriter Excel object.
    sort278_10.to_excel(writer, sheet_name='278',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort278.to_excel(writer, sheet_name='278',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 279
    # Convert the dataframe to an XlsxWriter Excel object.
    sort279_10.to_excel(writer, sheet_name='279',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort279.to_excel(writer, sheet_name='279',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 280
    # Convert the dataframe to an XlsxWriter Excel object.
    sort280_10.to_excel(writer, sheet_name='280',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort280.to_excel(writer, sheet_name='280',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 281
    # Convert the dataframe to an XlsxWriter Excel object.
    sort281_10.to_excel(writer, sheet_name='281',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort281.to_excel(writer, sheet_name='281',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 282
    # Convert the dataframe to an XlsxWriter Excel object.
    sort282_10.to_excel(writer, sheet_name='282',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort282.to_excel(writer, sheet_name='282',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 283
    # Convert the dataframe to an XlsxWriter Excel object.
    sort283_10.to_excel(writer, sheet_name='283',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort283.to_excel(writer, sheet_name='283',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 284
    # Convert the dataframe to an XlsxWriter Excel object.
    sort284_10.to_excel(writer, sheet_name='284',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort284.to_excel(writer, sheet_name='284',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 285
    # Convert the dataframe to an XlsxWriter Excel object.
    sort285_10.to_excel(writer, sheet_name='285',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort285.to_excel(writer, sheet_name='285',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 286
    # Convert the dataframe to an XlsxWriter Excel object.
    sort286_10.to_excel(writer, sheet_name='286',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort286.to_excel(writer, sheet_name='286',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 287
    # Convert the dataframe to an XlsxWriter Excel object.
    sort287_10.to_excel(writer, sheet_name='287',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort287.to_excel(writer, sheet_name='287',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 288
    # Convert the dataframe to an XlsxWriter Excel object.
    sort288_10.to_excel(writer, sheet_name='288',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort288.to_excel(writer, sheet_name='288',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 289
    # Convert the dataframe to an XlsxWriter Excel object.
    sort289_10.to_excel(writer, sheet_name='289',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort289.to_excel(writer, sheet_name='289',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 290
    # Convert the dataframe to an XlsxWriter Excel object.
    sort290_10.to_excel(writer, sheet_name='290',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort290.to_excel(writer, sheet_name='290',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 291
    # Convert the dataframe to an XlsxWriter Excel object.
    sort291_10.to_excel(writer, sheet_name='291',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort291.to_excel(writer, sheet_name='291',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 292
    # Convert the dataframe to an XlsxWriter Excel object.
    sort292_10.to_excel(writer, sheet_name='292',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort292.to_excel(writer, sheet_name='292',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 293
    # Convert the dataframe to an XlsxWriter Excel object.
    sort293_10.to_excel(writer, sheet_name='293',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort293.to_excel(writer, sheet_name='293',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 294
    # Convert the dataframe to an XlsxWriter Excel object.
    sort294_10.to_excel(writer, sheet_name='294',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort294.to_excel(writer, sheet_name='294',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 295
    # Convert the dataframe to an XlsxWriter Excel object.
    sort295_10.to_excel(writer, sheet_name='295',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort295.to_excel(writer, sheet_name='295',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 296
    # Convert the dataframe to an XlsxWriter Excel object.
    sort296_10.to_excel(writer, sheet_name='296',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort296.to_excel(writer, sheet_name='296',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 297
    # Convert the dataframe to an XlsxWriter Excel object.
    sort297_10.to_excel(writer, sheet_name='297',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort297.to_excel(writer, sheet_name='297',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 298
    # Convert the dataframe to an XlsxWriter Excel object.
    sort298_10.to_excel(writer, sheet_name='298',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort298.to_excel(writer, sheet_name='298',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 299
    # Convert the dataframe to an XlsxWriter Excel object.
    sort299_10.to_excel(writer, sheet_name='299',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort299.to_excel(writer, sheet_name='299',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)

    # Get the xlsxwriter objects from the dataframe writer object.
    workbook = writer.book

    # membuat worksheet baru
    worksheetcover = writer.sheets['cover']
    worksheet237 = writer.sheets['237']
    worksheet238 = writer.sheets['238']
    worksheet240 = writer.sheets['240']
    worksheet241 = writer.sheets['241']
    worksheet243 = writer.sheets['243']
    worksheet244 = writer.sheets['244']
    worksheet245 = writer.sheets['245']
    worksheet246 = writer.sheets['246']
    worksheet247 = writer.sheets['247']
    worksheet248 = writer.sheets['248']
    worksheet249 = writer.sheets['249']
    worksheet250 = writer.sheets['250']
    worksheet251 = writer.sheets['251']
    worksheet252 = writer.sheets['252']
    worksheet253 = writer.sheets['253']
    worksheet254 = writer.sheets['254']
    worksheet255 = writer.sheets['255']
    worksheet256 = writer.sheets['256']
    worksheet258 = writer.sheets['258']
    worksheet259 = writer.sheets['259']
    worksheet260 = writer.sheets['260']
    worksheet261 = writer.sheets['261']
    worksheet262 = writer.sheets['262']
    worksheet264 = writer.sheets['264']
    worksheet265 = writer.sheets['265']
    worksheet266 = writer.sheets['266']
    worksheet267 = writer.sheets['267']
    worksheet268 = writer.sheets['268']
    worksheet269 = writer.sheets['269']
    worksheet270 = writer.sheets['270']
    worksheet271 = writer.sheets['271']
    worksheet272 = writer.sheets['272']
    worksheet273 = writer.sheets['273']
    worksheet274 = writer.sheets['274']
    worksheet275 = writer.sheets['275']
    worksheet276 = writer.sheets['276']
    worksheet277 = writer.sheets['277']
    worksheet278 = writer.sheets['278']
    worksheet279 = writer.sheets['279']
    worksheet280 = writer.sheets['280']
    worksheet281 = writer.sheets['281']
    worksheet282 = writer.sheets['282']
    worksheet283 = writer.sheets['283']
    worksheet284 = writer.sheets['284']
    worksheet285 = writer.sheets['285']
    worksheet286 = writer.sheets['286']
    worksheet287 = writer.sheets['287']
    worksheet288 = writer.sheets['288']
    worksheet289 = writer.sheets['289']
    worksheet290 = writer.sheets['290']
    worksheet291 = writer.sheets['291']
    worksheet292 = writer.sheets['292']
    worksheet293 = writer.sheets['293']
    worksheet294 = writer.sheets['294']
    worksheet295 = writer.sheets['295']
    worksheet296 = writer.sheets['296']
    worksheet297 = writer.sheets['297']
    worksheet298 = writer.sheets['298']
    worksheet299 = writer.sheets['299']

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
    worksheetcover.conditional_format(16, 0, 11, 3,
                                      {'type': 'no_errors', 'format': borderCover})

    worksheetcover.insert_image('F1', r'logo nf.jpg')

    worksheetcover.merge_range('A10:A11', 'BIDANG STUDI', bodyCover)
    worksheetcover.merge_range('B10:B11', 'TERENDAH', bodyCover)
    worksheetcover.merge_range('C10:C11', 'RATA-RATA', bodyCover)
    worksheetcover.merge_range('D10:D11', 'TERTINGGI', bodyCover)
    worksheetcover.merge_range('A20:A21', 'BIDANG STUDI', bodyCover)
    worksheetcover.merge_range('B20:B21', 'TERENDAH', bodyCover)
    worksheetcover.merge_range('C20:C21', 'RATA-RATA', bodyCover)
    worksheetcover.merge_range('D20:D21', 'TERTINGGI', bodyCover)
    worksheetcover.write('F13', 'BIDANG STUDI', bodyCover)
    worksheetcover.merge_range('F20:F21', 'JUMLAH', sub_header1Cover)
    worksheetcover.merge_range('F23:F24', 'PESERTA', sub_header1Cover)
    worksheetcover.write('G13', 'JUMLAH', bodyCover)
    worksheetcover.set_column('A:A', 25.71, centerCover)
    worksheetcover.set_column('B:B', 15, centerCover)
    worksheetcover.set_column('C:C', 15, centerCover)
    worksheetcover.set_column('D:D', 15, centerCover)
    worksheetcover.set_column('F:F', 25.71, centerCover)
    worksheetcover.set_column('G:G', 13, centerCover)
    worksheetcover.merge_range('A1:F3', 'DAFTAR NILAI', titleCover)
    worksheetcover.merge_range(
        'A4:F5', fr'{penilaian}', sub_titleCover)
    worksheetcover.merge_range(
        'A6:F7', fr'{semester} TAHUN {tahun}', headerCover)
    worksheetcover.write('A9', 'JUMLAH BENAR', sub_headerCover)
    worksheetcover.write('A19', 'NILAI STANDAR', sub_headerCover)
    worksheetcover.merge_range('F8:G9', fr'{kelas}', kelasCover)
    worksheetcover.merge_range('F11:G12', 'JUMLAH SOAL', sub_header1Cover)

    worksheetcover.conditional_format(26, 0, 21, 3,
                                      {'type': 'no_errors', 'format': borderCover})

    worksheetcover.conditional_format(17, 6, 13, 5,
                                      {'type': 'no_errors', 'format': borderCover})

    worksheetcover.conditional_format(21, 5, 21, 5,
                                      {'type': 'no_errors', 'format': borderCover})

    # worksheet 237
    worksheet237.insert_image('A1', r'logo resmi nf.jpg')

    worksheet237.set_column('A:A', 7, center)
    worksheet237.set_column('B:B', 6, center)
    worksheet237.set_column('C:C', 18.14, center)
    worksheet237.set_column('D:D', 25, left)
    worksheet237.set_column('E:E', 13.14, left)
    worksheet237.set_column('F:F', 8.57, center)
    worksheet237.set_column('G:R', 5, center)
    worksheet237.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF TASIKMALAYA', title)
    worksheet237.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet237.write('A5', 'LOKASI', header)
    worksheet237.write('B5', 'TOTAL', header)
    worksheet237.merge_range('A4:B4', 'RANK', header)
    worksheet237.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet237.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet237.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet237.merge_range('F4:F5', 'KELAS', header)
    worksheet237.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet237.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet237.write('G5', 'MAT', body)
    worksheet237.write('H5', 'SEJ', body)
    worksheet237.write('I5', 'GEO', body)
    worksheet237.write('J5', 'EKO', body)
    worksheet237.write('K5', 'SOS', body)
    worksheet237.write('L5', 'JML', body)
    worksheet237.write('M5', 'MAT', body)
    worksheet237.write('N5', 'SEJ', body)
    worksheet237.write('O5', 'GEO', body)
    worksheet237.write('P5', 'EKO', body)
    worksheet237.write('Q5', 'SOS', body)
    worksheet237.write('R5', 'JML', body)

    worksheet237.conditional_format(5, 0, row237_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet237.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF TASIKMALAYA', title)
    worksheet237.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet237.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet237.write('A22', 'LOKASI', header)
    worksheet237.write('B22', 'TOTAL', header)
    worksheet237.merge_range('A21:B21', 'RANK', header)
    worksheet237.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet237.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet237.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet237.merge_range('F21:F22', 'KELAS', header)
    worksheet237.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet237.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet237.write('G22', 'MAT', body)
    worksheet237.write('H22', 'SEJ', body)
    worksheet237.write('I22', 'GEO', body)
    worksheet237.write('J22', 'EKO', body)
    worksheet237.write('K22', 'SOS', body)
    worksheet237.write('L22', 'JML', body)
    worksheet237.write('M22', 'MAT', body)
    worksheet237.write('N22', 'SEJ', body)
    worksheet237.write('O22', 'GEO', body)
    worksheet237.write('P22', 'EKO', body)
    worksheet237.write('Q22', 'SOS', body)
    worksheet237.write('R22', 'JML', body)

    worksheet237.conditional_format(22, 0, row237+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 238
    worksheet238.insert_image('A1', r'logo resmi nf.jpg')

    worksheet238.set_column('A:A', 7, center)
    worksheet238.set_column('B:B', 6, center)
    worksheet238.set_column('C:C', 18.14, center)
    worksheet238.set_column('D:D', 25, left)
    worksheet238.set_column('E:E', 13.14, left)
    worksheet238.set_column('F:F', 8.57, center)
    worksheet238.set_column('G:R', 5, center)
    worksheet238.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SUBANG', title)
    worksheet238.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet238.write('A5', 'LOKASI', header)
    worksheet238.write('B5', 'TOTAL', header)
    worksheet238.merge_range('A4:B4', 'RANK', header)
    worksheet238.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet238.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet238.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet238.merge_range('F4:F5', 'KELAS', header)
    worksheet238.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet238.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet238.write('G5', 'MAT', body)
    worksheet238.write('H5', 'SEJ', body)
    worksheet238.write('I5', 'GEO', body)
    worksheet238.write('J5', 'EKO', body)
    worksheet238.write('K5', 'SOS', body)
    worksheet238.write('L5', 'JML', body)
    worksheet238.write('M5', 'MAT', body)
    worksheet238.write('N5', 'SEJ', body)
    worksheet238.write('O5', 'GEO', body)
    worksheet238.write('P5', 'EKO', body)
    worksheet238.write('Q5', 'SOS', body)
    worksheet238.write('R5', 'JML', body)

    worksheet238.conditional_format(5, 0, row238_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet238.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF SUBANG', title)
    worksheet238.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet238.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet238.write('A22', 'LOKASI', header)
    worksheet238.write('B22', 'TOTAL', header)
    worksheet238.merge_range('A21:B21', 'RANK', header)
    worksheet238.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet238.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet238.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet238.merge_range('F21:F22', 'KELAS', header)
    worksheet238.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet238.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet238.write('G22', 'MAT', body)
    worksheet238.write('H22', 'SEJ', body)
    worksheet238.write('I22', 'GEO', body)
    worksheet238.write('J22', 'EKO', body)
    worksheet238.write('K22', 'SOS', body)
    worksheet238.write('L22', 'JML', body)
    worksheet238.write('M22', 'MAT', body)
    worksheet238.write('N22', 'SEJ', body)
    worksheet238.write('O22', 'GEO', body)
    worksheet238.write('P22', 'EKO', body)
    worksheet238.write('Q22', 'SOS', body)
    worksheet238.write('R22', 'JML', body)

    worksheet238.conditional_format(22, 0, row238+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 240
    worksheet240.insert_image('A1', r'logo resmi nf.jpg')

    worksheet240.set_column('A:A', 7, center)
    worksheet240.set_column('B:B', 6, center)
    worksheet240.set_column('C:C', 18.14, center)
    worksheet240.set_column('D:D', 25, left)
    worksheet240.set_column('E:E', 13.14, left)
    worksheet240.set_column('F:F', 8.57, center)
    worksheet240.set_column('G:R', 5, center)
    worksheet240.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SUMEDANG', title)
    worksheet240.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet240.write('A5', 'LOKASI', header)
    worksheet240.write('B5', 'TOTAL', header)
    worksheet240.merge_range('A4:B4', 'RANK', header)
    worksheet240.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet240.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet240.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet240.merge_range('F4:F5', 'KELAS', header)
    worksheet240.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet240.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet240.write('G5', 'MAT', body)
    worksheet240.write('H5', 'SEJ', body)
    worksheet240.write('I5', 'GEO', body)
    worksheet240.write('J5', 'EKO', body)
    worksheet240.write('K5', 'SOS', body)
    worksheet240.write('L5', 'JML', body)
    worksheet240.write('M5', 'MAT', body)
    worksheet240.write('N5', 'SEJ', body)
    worksheet240.write('O5', 'GEO', body)
    worksheet240.write('P5', 'EKO', body)
    worksheet240.write('Q5', 'SOS', body)
    worksheet240.write('R5', 'JML', body)

    worksheet240.conditional_format(5, 0, row240_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet240.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF SUMEDANG', title)
    worksheet240.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet240.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet240.write('A22', 'LOKASI', header)
    worksheet240.write('B22', 'TOTAL', header)
    worksheet240.merge_range('A21:B21', 'RANK', header)
    worksheet240.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet240.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet240.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet240.merge_range('F21:F22', 'KELAS', header)
    worksheet240.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet240.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet240.write('G22', 'MAT', body)
    worksheet240.write('H22', 'SEJ', body)
    worksheet240.write('I22', 'GEO', body)
    worksheet240.write('J22', 'EKO', body)
    worksheet240.write('K22', 'SOS', body)
    worksheet240.write('L22', 'JML', body)
    worksheet240.write('M22', 'MAT', body)
    worksheet240.write('N22', 'SEJ', body)
    worksheet240.write('O22', 'GEO', body)
    worksheet240.write('P22', 'EKO', body)
    worksheet240.write('Q22', 'SOS', body)
    worksheet240.write('R22', 'JML', body)

    worksheet240.conditional_format(22, 0, row240+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 241
    worksheet241.insert_image('A1', r'logo resmi nf.jpg')

    worksheet241.set_column('A:A', 7, center)
    worksheet241.set_column('B:B', 6, center)
    worksheet241.set_column('C:C', 18.14, center)
    worksheet241.set_column('D:D', 25, left)
    worksheet241.set_column('E:E', 13.14, left)
    worksheet241.set_column('F:F', 8.57, center)
    worksheet241.set_column('G:R', 5, center)
    worksheet241.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MAJALENGKA', title)
    worksheet241.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet241.write('A5', 'LOKASI', header)
    worksheet241.write('B5', 'TOTAL', header)
    worksheet241.merge_range('A4:B4', 'RANK', header)
    worksheet241.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet241.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet241.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet241.merge_range('F4:F5', 'KELAS', header)
    worksheet241.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet241.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet241.write('G5', 'MAT', body)
    worksheet241.write('H5', 'SEJ', body)
    worksheet241.write('I5', 'GEO', body)
    worksheet241.write('J5', 'EKO', body)
    worksheet241.write('K5', 'SOS', body)
    worksheet241.write('L5', 'JML', body)
    worksheet241.write('M5', 'MAT', body)
    worksheet241.write('N5', 'SEJ', body)
    worksheet241.write('O5', 'GEO', body)
    worksheet241.write('P5', 'EKO', body)
    worksheet241.write('Q5', 'SOS', body)
    worksheet241.write('R5', 'JML', body)

    worksheet241.conditional_format(5, 0, row241_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet241.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF MAJALENGKA', title)
    worksheet241.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet241.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet241.write('A22', 'LOKASI', header)
    worksheet241.write('B22', 'TOTAL', header)
    worksheet241.merge_range('A21:B21', 'RANK', header)
    worksheet241.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet241.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet241.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet241.merge_range('F21:F22', 'KELAS', header)
    worksheet241.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet241.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet241.write('G22', 'MAT', body)
    worksheet241.write('H22', 'SEJ', body)
    worksheet241.write('I22', 'GEO', body)
    worksheet241.write('J22', 'EKO', body)
    worksheet241.write('K22', 'SOS', body)
    worksheet241.write('L22', 'JML', body)
    worksheet241.write('M22', 'MAT', body)
    worksheet241.write('N22', 'SEJ', body)
    worksheet241.write('O22', 'GEO', body)
    worksheet241.write('P22', 'EKO', body)
    worksheet241.write('Q22', 'SOS', body)
    worksheet241.write('R22', 'JML', body)

    worksheet241.conditional_format(22, 0, row241+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 243
    worksheet243.insert_image('A1', r'logo resmi nf.jpg')

    worksheet243.set_column('A:A', 7, center)
    worksheet243.set_column('B:B', 6, center)
    worksheet243.set_column('C:C', 18.14, center)
    worksheet243.set_column('D:D', 25, left)
    worksheet243.set_column('E:E', 13.14, left)
    worksheet243.set_column('F:F', 8.57, center)
    worksheet243.set_column('G:R', 5, center)
    worksheet243.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PURBALINGGA', title)
    worksheet243.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet243.write('A5', 'LOKASI', header)
    worksheet243.write('B5', 'TOTAL', header)
    worksheet243.merge_range('A4:B4', 'RANK', header)
    worksheet243.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet243.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet243.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet243.merge_range('F4:F5', 'KELAS', header)
    worksheet243.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet243.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet243.write('G5', 'MAT', body)
    worksheet243.write('H5', 'SEJ', body)
    worksheet243.write('I5', 'GEO', body)
    worksheet243.write('J5', 'EKO', body)
    worksheet243.write('K5', 'SOS', body)
    worksheet243.write('L5', 'JML', body)
    worksheet243.write('M5', 'MAT', body)
    worksheet243.write('N5', 'SEJ', body)
    worksheet243.write('O5', 'GEO', body)
    worksheet243.write('P5', 'EKO', body)
    worksheet243.write('Q5', 'SOS', body)
    worksheet243.write('R5', 'JML', body)

    worksheet243.conditional_format(5, 0, row243_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet243.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PURBALINGGA', title)
    worksheet243.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet243.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet243.write('A22', 'LOKASI', header)
    worksheet243.write('B22', 'TOTAL', header)
    worksheet243.merge_range('A21:B21', 'RANK', header)
    worksheet243.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet243.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet243.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet243.merge_range('F21:F22', 'KELAS', header)
    worksheet243.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet243.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet243.write('G22', 'MAT', body)
    worksheet243.write('H22', 'SEJ', body)
    worksheet243.write('I22', 'GEO', body)
    worksheet243.write('J22', 'EKO', body)
    worksheet243.write('K22', 'SOS', body)
    worksheet243.write('L22', 'JML', body)
    worksheet243.write('M22', 'MAT', body)
    worksheet243.write('N22', 'SEJ', body)
    worksheet243.write('O22', 'GEO', body)
    worksheet243.write('P22', 'EKO', body)
    worksheet243.write('Q22', 'SOS', body)
    worksheet243.write('R22', 'JML', body)

    worksheet243.conditional_format(22, 0, row243+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 244
    worksheet244.insert_image('A1', r'logo resmi nf.jpg')

    worksheet244.set_column('A:A', 7, center)
    worksheet244.set_column('B:B', 6, center)
    worksheet244.set_column('C:C', 18.14, center)
    worksheet244.set_column('D:D', 25, left)
    worksheet244.set_column('E:E', 13.14, left)
    worksheet244.set_column('F:F', 8.57, center)
    worksheet244.set_column('G:R', 5, center)
    worksheet244.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SURAKARTA', title)
    worksheet244.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet244.write('A5', 'LOKASI', header)
    worksheet244.write('B5', 'TOTAL', header)
    worksheet244.merge_range('A4:B4', 'RANK', header)
    worksheet244.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet244.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet244.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet244.merge_range('F4:F5', 'KELAS', header)
    worksheet244.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet244.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet244.write('G5', 'MAT', body)
    worksheet244.write('H5', 'SEJ', body)
    worksheet244.write('I5', 'GEO', body)
    worksheet244.write('J5', 'EKO', body)
    worksheet244.write('K5', 'SOS', body)
    worksheet244.write('L5', 'JML', body)
    worksheet244.write('M5', 'MAT', body)
    worksheet244.write('N5', 'SEJ', body)
    worksheet244.write('O5', 'GEO', body)
    worksheet244.write('P5', 'EKO', body)
    worksheet244.write('Q5', 'SOS', body)
    worksheet244.write('R5', 'JML', body)

    worksheet244.conditional_format(5, 0, row244_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet244.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF SURAKARTA', title)
    worksheet244.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet244.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet244.write('A22', 'LOKASI', header)
    worksheet244.write('B22', 'TOTAL', header)
    worksheet244.merge_range('A21:B21', 'RANK', header)
    worksheet244.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet244.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet244.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet244.merge_range('F21:F22', 'KELAS', header)
    worksheet244.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet244.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet244.write('G22', 'MAT', body)
    worksheet244.write('H22', 'SEJ', body)
    worksheet244.write('I22', 'GEO', body)
    worksheet244.write('J22', 'EKO', body)
    worksheet244.write('K22', 'SOS', body)
    worksheet244.write('L22', 'JML', body)
    worksheet244.write('M22', 'MAT', body)
    worksheet244.write('N22', 'SEJ', body)
    worksheet244.write('O22', 'GEO', body)
    worksheet244.write('P22', 'EKO', body)
    worksheet244.write('Q22', 'SOS', body)
    worksheet244.write('R22', 'JML', body)

    worksheet244.conditional_format(22, 0, row244+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 245
    worksheet245.insert_image('A1', r'logo resmi nf.jpg')

    worksheet245.set_column('A:A', 7, center)
    worksheet245.set_column('B:B', 6, center)
    worksheet245.set_column('C:C', 18.14, center)
    worksheet245.set_column('D:D', 25, left)
    worksheet245.set_column('E:E', 13.14, left)
    worksheet245.set_column('F:F', 8.57, center)
    worksheet245.set_column('G:R', 5, center)
    worksheet245.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SEMARANG', title)
    worksheet245.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet245.write('A5', 'LOKASI', header)
    worksheet245.write('B5', 'TOTAL', header)
    worksheet245.merge_range('A4:B4', 'RANK', header)
    worksheet245.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet245.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet245.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet245.merge_range('F4:F5', 'KELAS', header)
    worksheet245.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet245.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet245.write('G5', 'MAT', body)
    worksheet245.write('H5', 'SEJ', body)
    worksheet245.write('I5', 'GEO', body)
    worksheet245.write('J5', 'EKO', body)
    worksheet245.write('K5', 'SOS', body)
    worksheet245.write('L5', 'JML', body)
    worksheet245.write('M5', 'MAT', body)
    worksheet245.write('N5', 'SEJ', body)
    worksheet245.write('O5', 'GEO', body)
    worksheet245.write('P5', 'EKO', body)
    worksheet245.write('Q5', 'SOS', body)
    worksheet245.write('R5', 'JML', body)

    worksheet245.conditional_format(5, 0, row245_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet245.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF SEMARANG', title)
    worksheet245.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet245.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet245.write('A22', 'LOKASI', header)
    worksheet245.write('B22', 'TOTAL', header)
    worksheet245.merge_range('A21:B21', 'RANK', header)
    worksheet245.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet245.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet245.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet245.merge_range('F21:F22', 'KELAS', header)
    worksheet245.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet245.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet245.write('G22', 'MAT', body)
    worksheet245.write('H22', 'SEJ', body)
    worksheet245.write('I22', 'GEO', body)
    worksheet245.write('J22', 'EKO', body)
    worksheet245.write('K22', 'SOS', body)
    worksheet245.write('L22', 'JML', body)
    worksheet245.write('M22', 'MAT', body)
    worksheet245.write('N22', 'SEJ', body)
    worksheet245.write('O22', 'GEO', body)
    worksheet245.write('P22', 'EKO', body)
    worksheet245.write('Q22', 'SOS', body)
    worksheet245.write('R22', 'JML', body)

    worksheet245.conditional_format(22, 0, row245+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 246
    worksheet246.insert_image('A1', r'logo resmi nf.jpg')

    worksheet246.set_column('A:A', 7, center)
    worksheet246.set_column('B:B', 6, center)
    worksheet246.set_column('C:C', 18.14, center)
    worksheet246.set_column('D:D', 25, left)
    worksheet246.set_column('E:E', 13.14, left)
    worksheet246.set_column('F:F', 8.57, center)
    worksheet246.set_column('G:R', 5, center)
    worksheet246.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KARTASURA', title)
    worksheet246.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet246.write('A5', 'LOKASI', header)
    worksheet246.write('B5', 'TOTAL', header)
    worksheet246.merge_range('A4:B4', 'RANK', header)
    worksheet246.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet246.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet246.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet246.merge_range('F4:F5', 'KELAS', header)
    worksheet246.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet246.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet246.write('G5', 'MAT', body)
    worksheet246.write('H5', 'SEJ', body)
    worksheet246.write('I5', 'GEO', body)
    worksheet246.write('J5', 'EKO', body)
    worksheet246.write('K5', 'SOS', body)
    worksheet246.write('L5', 'JML', body)
    worksheet246.write('M5', 'MAT', body)
    worksheet246.write('N5', 'SEJ', body)
    worksheet246.write('O5', 'GEO', body)
    worksheet246.write('P5', 'EKO', body)
    worksheet246.write('Q5', 'SOS', body)
    worksheet246.write('R5', 'JML', body)

    worksheet246.conditional_format(5, 0, row246_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet246.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF KARTASURA', title)
    worksheet246.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet246.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet246.write('A22', 'LOKASI', header)
    worksheet246.write('B22', 'TOTAL', header)
    worksheet246.merge_range('A21:B21', 'RANK', header)
    worksheet246.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet246.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet246.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet246.merge_range('F21:F22', 'KELAS', header)
    worksheet246.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet246.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet246.write('G22', 'MAT', body)
    worksheet246.write('H22', 'SEJ', body)
    worksheet246.write('I22', 'GEO', body)
    worksheet246.write('J22', 'EKO', body)
    worksheet246.write('K22', 'SOS', body)
    worksheet246.write('L22', 'JML', body)
    worksheet246.write('M22', 'MAT', body)
    worksheet246.write('N22', 'SEJ', body)
    worksheet246.write('O22', 'GEO', body)
    worksheet246.write('P22', 'EKO', body)
    worksheet246.write('Q22', 'SOS', body)
    worksheet246.write('R22', 'JML', body)

    worksheet246.conditional_format(22, 0, row246+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 247
    worksheet247.insert_image('A1', r'logo resmi nf.jpg')

    worksheet247.set_column('A:A', 7, center)
    worksheet247.set_column('B:B', 6, center)
    worksheet247.set_column('C:C', 18.14, center)
    worksheet247.set_column('D:D', 25, left)
    worksheet247.set_column('E:E', 13.14, left)
    worksheet247.set_column('F:F', 8.57, center)
    worksheet247.set_column('G:R', 5, center)
    worksheet247.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF JURANG MANGU', title)
    worksheet247.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet247.write('A5', 'LOKASI', header)
    worksheet247.write('B5', 'TOTAL', header)
    worksheet247.merge_range('A4:B4', 'RANK', header)
    worksheet247.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet247.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet247.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet247.merge_range('F4:F5', 'KELAS', header)
    worksheet247.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet247.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet247.write('G5', 'MAT', body)
    worksheet247.write('H5', 'SEJ', body)
    worksheet247.write('I5', 'GEO', body)
    worksheet247.write('J5', 'EKO', body)
    worksheet247.write('K5', 'SOS', body)
    worksheet247.write('L5', 'JML', body)
    worksheet247.write('M5', 'MAT', body)
    worksheet247.write('N5', 'SEJ', body)
    worksheet247.write('O5', 'GEO', body)
    worksheet247.write('P5', 'EKO', body)
    worksheet247.write('Q5', 'SOS', body)
    worksheet247.write('R5', 'JML', body)

    worksheet247.conditional_format(5, 0, row247_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet247.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF JURANG MANGU', title)
    worksheet247.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet247.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet247.write('A22', 'LOKASI', header)
    worksheet247.write('B22', 'TOTAL', header)
    worksheet247.merge_range('A21:B21', 'RANK', header)
    worksheet247.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet247.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet247.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet247.merge_range('F21:F22', 'KELAS', header)
    worksheet247.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet247.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet247.write('G22', 'MAT', body)
    worksheet247.write('H22', 'SEJ', body)
    worksheet247.write('I22', 'GEO', body)
    worksheet247.write('J22', 'EKO', body)
    worksheet247.write('K22', 'SOS', body)
    worksheet247.write('L22', 'JML', body)
    worksheet247.write('M22', 'MAT', body)
    worksheet247.write('N22', 'SEJ', body)
    worksheet247.write('O22', 'GEO', body)
    worksheet247.write('P22', 'EKO', body)
    worksheet247.write('Q22', 'SOS', body)
    worksheet247.write('R22', 'JML', body)

    worksheet247.conditional_format(22, 0, row247+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 248
    worksheet248.insert_image('A1', r'logo resmi nf.jpg')

    worksheet248.set_column('A:A', 7, center)
    worksheet248.set_column('B:B', 6, center)
    worksheet248.set_column('C:C', 18.14, center)
    worksheet248.set_column('D:D', 25, left)
    worksheet248.set_column('E:E', 13.14, left)
    worksheet248.set_column('F:F', 8.57, center)
    worksheet248.set_column('G:R', 5, center)
    worksheet248.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BSD BOULEVARD', title)
    worksheet248.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet248.write('A5', 'LOKASI', header)
    worksheet248.write('B5', 'TOTAL', header)
    worksheet248.merge_range('A4:B4', 'RANK', header)
    worksheet248.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet248.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet248.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet248.merge_range('F4:F5', 'KELAS', header)
    worksheet248.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet248.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet248.write('G5', 'MAT', body)
    worksheet248.write('H5', 'SEJ', body)
    worksheet248.write('I5', 'GEO', body)
    worksheet248.write('J5', 'EKO', body)
    worksheet248.write('K5', 'SOS', body)
    worksheet248.write('L5', 'JML', body)
    worksheet248.write('M5', 'MAT', body)
    worksheet248.write('N5', 'SEJ', body)
    worksheet248.write('O5', 'GEO', body)
    worksheet248.write('P5', 'EKO', body)
    worksheet248.write('Q5', 'SOS', body)
    worksheet248.write('R5', 'JML', body)

    worksheet248.conditional_format(5, 0, row248_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet248.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF BSD BOULEVARD', title)
    worksheet248.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet248.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet248.write('A22', 'LOKASI', header)
    worksheet248.write('B22', 'TOTAL', header)
    worksheet248.merge_range('A21:B21', 'RANK', header)
    worksheet248.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet248.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet248.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet248.merge_range('F21:F22', 'KELAS', header)
    worksheet248.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet248.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet248.write('G22', 'MAT', body)
    worksheet248.write('H22', 'SEJ', body)
    worksheet248.write('I22', 'GEO', body)
    worksheet248.write('J22', 'EKO', body)
    worksheet248.write('K22', 'SOS', body)
    worksheet248.write('L22', 'JML', body)
    worksheet248.write('M22', 'MAT', body)
    worksheet248.write('N22', 'SEJ', body)
    worksheet248.write('O22', 'GEO', body)
    worksheet248.write('P22', 'EKO', body)
    worksheet248.write('Q22', 'SOS', body)
    worksheet248.write('R22', 'JML', body)

    worksheet248.conditional_format(22, 0, row248+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 249
    worksheet249.insert_image('A1', r'logo resmi nf.jpg')

    worksheet249.set_column('A:A', 7, center)
    worksheet249.set_column('B:B', 6, center)
    worksheet249.set_column('C:C', 18.14, center)
    worksheet249.set_column('D:D', 25, left)
    worksheet249.set_column('E:E', 13.14, left)
    worksheet249.set_column('F:F', 8.57, center)
    worksheet249.set_column('G:R', 5, center)
    worksheet249.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SANGIANG', title)
    worksheet249.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet249.write('A5', 'LOKASI', header)
    worksheet249.write('B5', 'TOTAL', header)
    worksheet249.merge_range('A4:B4', 'RANK', header)
    worksheet249.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet249.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet249.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet249.merge_range('F4:F5', 'KELAS', header)
    worksheet249.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet249.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet249.write('G5', 'MAT', body)
    worksheet249.write('H5', 'SEJ', body)
    worksheet249.write('I5', 'GEO', body)
    worksheet249.write('J5', 'EKO', body)
    worksheet249.write('K5', 'SOS', body)
    worksheet249.write('L5', 'JML', body)
    worksheet249.write('M5', 'MAT', body)
    worksheet249.write('N5', 'SEJ', body)
    worksheet249.write('O5', 'GEO', body)
    worksheet249.write('P5', 'EKO', body)
    worksheet249.write('Q5', 'SOS', body)
    worksheet249.write('R5', 'JML', body)

    worksheet249.conditional_format(5, 0, row249_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet249.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF SANGIANG', title)
    worksheet249.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet249.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet249.write('A22', 'LOKASI', header)
    worksheet249.write('B22', 'TOTAL', header)
    worksheet249.merge_range('A21:B21', 'RANK', header)
    worksheet249.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet249.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet249.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet249.merge_range('F21:F22', 'KELAS', header)
    worksheet249.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet249.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet249.write('G22', 'MAT', body)
    worksheet249.write('H22', 'SEJ', body)
    worksheet249.write('I22', 'GEO', body)
    worksheet249.write('J22', 'EKO', body)
    worksheet249.write('K22', 'SOS', body)
    worksheet249.write('L22', 'JML', body)
    worksheet249.write('M22', 'MAT', body)
    worksheet249.write('N22', 'SEJ', body)
    worksheet249.write('O22', 'GEO', body)
    worksheet249.write('P22', 'EKO', body)
    worksheet249.write('Q22', 'SOS', body)
    worksheet249.write('R22', 'JML', body)

    worksheet249.conditional_format(22, 0, row249+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 250
    worksheet250.insert_image('A1', r'logo resmi nf.jpg')

    worksheet250.set_column('A:A', 7, center)
    worksheet250.set_column('B:B', 6, center)
    worksheet250.set_column('C:C', 18.14, center)
    worksheet250.set_column('D:D', 25, left)
    worksheet250.set_column('E:E', 13.14, left)
    worksheet250.set_column('F:F', 8.57, center)
    worksheet250.set_column('G:R', 5, center)
    worksheet250.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BANJAR WIJAYA', title)
    worksheet250.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet250.write('A5', 'LOKASI', header)
    worksheet250.write('B5', 'TOTAL', header)
    worksheet250.merge_range('A4:B4', 'RANK', header)
    worksheet250.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet250.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet250.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet250.merge_range('F4:F5', 'KELAS', header)
    worksheet250.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet250.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet250.write('G5', 'MAT', body)
    worksheet250.write('H5', 'SEJ', body)
    worksheet250.write('I5', 'GEO', body)
    worksheet250.write('J5', 'EKO', body)
    worksheet250.write('K5', 'SOS', body)
    worksheet250.write('L5', 'JML', body)
    worksheet250.write('M5', 'MAT', body)
    worksheet250.write('N5', 'SEJ', body)
    worksheet250.write('O5', 'GEO', body)
    worksheet250.write('P5', 'EKO', body)
    worksheet250.write('Q5', 'SOS', body)
    worksheet250.write('R5', 'JML', body)

    worksheet250.conditional_format(5, 0, row250_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet250.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF BANJAR WIJAYA', title)
    worksheet250.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet250.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet250.write('A22', 'LOKASI', header)
    worksheet250.write('B22', 'TOTAL', header)
    worksheet250.merge_range('A21:B21', 'RANK', header)
    worksheet250.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet250.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet250.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet250.merge_range('F21:F22', 'KELAS', header)
    worksheet250.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet250.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet250.write('G22', 'MAT', body)
    worksheet250.write('H22', 'SEJ', body)
    worksheet250.write('I22', 'GEO', body)
    worksheet250.write('J22', 'EKO', body)
    worksheet250.write('K22', 'SOS', body)
    worksheet250.write('L22', 'JML', body)
    worksheet250.write('M22', 'MAT', body)
    worksheet250.write('N22', 'SEJ', body)
    worksheet250.write('O22', 'GEO', body)
    worksheet250.write('P22', 'EKO', body)
    worksheet250.write('Q22', 'SOS', body)
    worksheet250.write('R22', 'JML', body)

    worksheet250.conditional_format(22, 0, row250+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 251
    worksheet251.insert_image('A1', r'logo resmi nf.jpg')

    worksheet251.set_column('A:A', 7, center)
    worksheet251.set_column('B:B', 6, center)
    worksheet251.set_column('C:C', 18.14, center)
    worksheet251.set_column('D:D', 25, left)
    worksheet251.set_column('E:E', 13.14, left)
    worksheet251.set_column('F:F', 8.57, center)
    worksheet251.set_column('G:R', 5, center)
    worksheet251.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MELATI MAS', title)
    worksheet251.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet251.write('A5', 'LOKASI', header)
    worksheet251.write('B5', 'TOTAL', header)
    worksheet251.merge_range('A4:B4', 'RANK', header)
    worksheet251.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet251.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet251.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet251.merge_range('F4:F5', 'KELAS', header)
    worksheet251.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet251.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet251.write('G5', 'MAT', body)
    worksheet251.write('H5', 'SEJ', body)
    worksheet251.write('I5', 'GEO', body)
    worksheet251.write('J5', 'EKO', body)
    worksheet251.write('K5', 'SOS', body)
    worksheet251.write('L5', 'JML', body)
    worksheet251.write('M5', 'MAT', body)
    worksheet251.write('N5', 'SEJ', body)
    worksheet251.write('O5', 'GEO', body)
    worksheet251.write('P5', 'EKO', body)
    worksheet251.write('Q5', 'SOS', body)
    worksheet251.write('R5', 'JML', body)

    worksheet251.conditional_format(5, 0, row251_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet251.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF MELATI MAS', title)
    worksheet251.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet251.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet251.write('A22', 'LOKASI', header)
    worksheet251.write('B22', 'TOTAL', header)
    worksheet251.merge_range('A21:B21', 'RANK', header)
    worksheet251.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet251.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet251.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet251.merge_range('F21:F22', 'KELAS', header)
    worksheet251.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet251.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet251.write('G22', 'MAT', body)
    worksheet251.write('H22', 'SEJ', body)
    worksheet251.write('I22', 'GEO', body)
    worksheet251.write('J22', 'EKO', body)
    worksheet251.write('K22', 'SOS', body)
    worksheet251.write('L22', 'JML', body)
    worksheet251.write('M22', 'MAT', body)
    worksheet251.write('N22', 'SEJ', body)
    worksheet251.write('O22', 'GEO', body)
    worksheet251.write('P22', 'EKO', body)
    worksheet251.write('Q22', 'SOS', body)
    worksheet251.write('R22', 'JML', body)

    worksheet251.conditional_format(22, 0, row251+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 252
    worksheet252.insert_image('A1', r'logo resmi nf.jpg')

    worksheet252.set_column('A:A', 7, center)
    worksheet252.set_column('B:B', 6, center)
    worksheet252.set_column('C:C', 18.14, center)
    worksheet252.set_column('D:D', 25, left)
    worksheet252.set_column('E:E', 13.14, left)
    worksheet252.set_column('F:F', 8.57, center)
    worksheet252.set_column('G:R', 5, center)
    worksheet252.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIRENDEU', title)
    worksheet252.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet252.write('A5', 'LOKASI', header)
    worksheet252.write('B5', 'TOTAL', header)
    worksheet252.merge_range('A4:B4', 'RANK', header)
    worksheet252.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet252.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet252.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet252.merge_range('F4:F5', 'KELAS', header)
    worksheet252.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet252.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet252.write('G5', 'MAT', body)
    worksheet252.write('H5', 'SEJ', body)
    worksheet252.write('I5', 'GEO', body)
    worksheet252.write('J5', 'EKO', body)
    worksheet252.write('K5', 'SOS', body)
    worksheet252.write('L5', 'JML', body)
    worksheet252.write('M5', 'MAT', body)
    worksheet252.write('N5', 'SEJ', body)
    worksheet252.write('O5', 'GEO', body)
    worksheet252.write('P5', 'EKO', body)
    worksheet252.write('Q5', 'SOS', body)
    worksheet252.write('R5', 'JML', body)

    worksheet252.conditional_format(5, 0, row252_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet252.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CIRENDEU', title)
    worksheet252.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet252.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet252.write('A22', 'LOKASI', header)
    worksheet252.write('B22', 'TOTAL', header)
    worksheet252.merge_range('A21:B21', 'RANK', header)
    worksheet252.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet252.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet252.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet252.merge_range('F21:F22', 'KELAS', header)
    worksheet252.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet252.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet252.write('G22', 'MAT', body)
    worksheet252.write('H22', 'SEJ', body)
    worksheet252.write('I22', 'GEO', body)
    worksheet252.write('J22', 'EKO', body)
    worksheet252.write('K22', 'SOS', body)
    worksheet252.write('L22', 'JML', body)
    worksheet252.write('M22', 'MAT', body)
    worksheet252.write('N22', 'SEJ', body)
    worksheet252.write('O22', 'GEO', body)
    worksheet252.write('P22', 'EKO', body)
    worksheet252.write('Q22', 'SOS', body)
    worksheet252.write('R22', 'JML', body)

    worksheet252.conditional_format(22, 0, row252+21, 17,
                                    {'type': 'no_errors', 'format': border})
    # worksheet 253
    worksheet253.insert_image('A1', r'logo resmi nf.jpg')

    worksheet253.set_column('A:A', 7, center)
    worksheet253.set_column('B:B', 6, center)
    worksheet253.set_column('C:C', 18.14, center)
    worksheet253.set_column('D:D', 25, left)
    worksheet253.set_column('E:E', 13.14, left)
    worksheet253.set_column('F:F', 8.57, center)
    worksheet253.set_column('G:R', 5, center)
    worksheet253.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KAMPUNG UTAN', title)
    worksheet253.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet253.write('A5', 'LOKASI', header)
    worksheet253.write('B5', 'TOTAL', header)
    worksheet253.merge_range('A4:B4', 'RANK', header)
    worksheet253.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet253.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet253.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet253.merge_range('F4:F5', 'KELAS', header)
    worksheet253.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet253.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet253.write('G5', 'MAT', body)
    worksheet253.write('H5', 'SEJ', body)
    worksheet253.write('I5', 'GEO', body)
    worksheet253.write('J5', 'EKO', body)
    worksheet253.write('K5', 'SOS', body)
    worksheet253.write('L5', 'JML', body)
    worksheet253.write('M5', 'MAT', body)
    worksheet253.write('N5', 'SEJ', body)
    worksheet253.write('O5', 'GEO', body)
    worksheet253.write('P5', 'EKO', body)
    worksheet253.write('Q5', 'SOS', body)
    worksheet253.write('R5', 'JML', body)

    worksheet253.conditional_format(5, 0, row253_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet253.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF KAMPUNG UTAN', title)
    worksheet253.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet253.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet253.write('A22', 'LOKASI', header)
    worksheet253.write('B22', 'TOTAL', header)
    worksheet253.merge_range('A21:B21', 'RANK', header)
    worksheet253.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet253.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet253.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet253.merge_range('F21:F22', 'KELAS', header)
    worksheet253.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet253.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet253.write('G22', 'MAT', body)
    worksheet253.write('H22', 'SEJ', body)
    worksheet253.write('I22', 'GEO', body)
    worksheet253.write('J22', 'EKO', body)
    worksheet253.write('K22', 'SOS', body)
    worksheet253.write('L22', 'JML', body)
    worksheet253.write('M22', 'MAT', body)
    worksheet253.write('N22', 'SEJ', body)
    worksheet253.write('O22', 'GEO', body)
    worksheet253.write('P22', 'EKO', body)
    worksheet253.write('Q22', 'SOS', body)
    worksheet253.write('R22', 'JML', body)

    worksheet253.conditional_format(22, 0, row253+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 254
    worksheet254.insert_image('A1', r'logo resmi nf.jpg')

    worksheet254.set_column('A:A', 7, center)
    worksheet254.set_column('B:B', 6, center)
    worksheet254.set_column('C:C', 18.14, center)
    worksheet254.set_column('D:D', 25, left)
    worksheet254.set_column('E:E', 13.14, left)
    worksheet254.set_column('F:F', 8.57, center)
    worksheet254.set_column('G:R', 5, center)
    worksheet254.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF GRAHA RAYA', title)
    worksheet254.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet254.write('A5', 'LOKASI', header)
    worksheet254.write('B5', 'TOTAL', header)
    worksheet254.merge_range('A4:B4', 'RANK', header)
    worksheet254.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet254.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet254.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet254.merge_range('F4:F5', 'KELAS', header)
    worksheet254.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet254.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet254.write('G5', 'MAT', body)
    worksheet254.write('H5', 'SEJ', body)
    worksheet254.write('I5', 'GEO', body)
    worksheet254.write('J5', 'EKO', body)
    worksheet254.write('K5', 'SOS', body)
    worksheet254.write('L5', 'JML', body)
    worksheet254.write('M5', 'MAT', body)
    worksheet254.write('N5', 'SEJ', body)
    worksheet254.write('O5', 'GEO', body)
    worksheet254.write('P5', 'EKO', body)
    worksheet254.write('Q5', 'SOS', body)
    worksheet254.write('R5', 'JML', body)

    worksheet254.conditional_format(5, 0, row254_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet254.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF GRAHA RAYA', title)
    worksheet254.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet254.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet254.write('A22', 'LOKASI', header)
    worksheet254.write('B22', 'TOTAL', header)
    worksheet254.merge_range('A21:B21', 'RANK', header)
    worksheet254.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet254.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet254.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet254.merge_range('F21:F22', 'KELAS', header)
    worksheet254.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet254.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet254.write('G22', 'MAT', body)
    worksheet254.write('H22', 'SEJ', body)
    worksheet254.write('I22', 'GEO', body)
    worksheet254.write('J22', 'EKO', body)
    worksheet254.write('K22', 'SOS', body)
    worksheet254.write('L22', 'JML', body)
    worksheet254.write('M22', 'MAT', body)
    worksheet254.write('N22', 'SEJ', body)
    worksheet254.write('O22', 'GEO', body)
    worksheet254.write('P22', 'EKO', body)
    worksheet254.write('Q22', 'SOS', body)
    worksheet254.write('R22', 'JML', body)

    worksheet254.conditional_format(22, 0, row254+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 255
    worksheet255.insert_image('A1', r'logo resmi nf.jpg')

    worksheet255.set_column('A:A', 7, center)
    worksheet255.set_column('B:B', 6, center)
    worksheet255.set_column('C:C', 18.14, center)
    worksheet255.set_column('D:D', 25, left)
    worksheet255.set_column('E:E', 13.14, left)
    worksheet255.set_column('F:F', 8.57, center)
    worksheet255.set_column('G:R', 5, center)
    worksheet255.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MERPATI', title)
    worksheet255.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet255.write('A5', 'LOKASI', header)
    worksheet255.write('B5', 'TOTAL', header)
    worksheet255.merge_range('A4:B4', 'RANK', header)
    worksheet255.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet255.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet255.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet255.merge_range('F4:F5', 'KELAS', header)
    worksheet255.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet255.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet255.write('G5', 'MAT', body)
    worksheet255.write('H5', 'SEJ', body)
    worksheet255.write('I5', 'GEO', body)
    worksheet255.write('J5', 'EKO', body)
    worksheet255.write('K5', 'SOS', body)
    worksheet255.write('L5', 'JML', body)
    worksheet255.write('M5', 'MAT', body)
    worksheet255.write('N5', 'SEJ', body)
    worksheet255.write('O5', 'GEO', body)
    worksheet255.write('P5', 'EKO', body)
    worksheet255.write('Q5', 'SOS', body)
    worksheet255.write('R5', 'JML', body)

    worksheet255.conditional_format(5, 0, row255_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet255.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF MERPATI', title)
    worksheet255.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet255.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet255.write('A22', 'LOKASI', header)
    worksheet255.write('B22', 'TOTAL', header)
    worksheet255.merge_range('A21:B21', 'RANK', header)
    worksheet255.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet255.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet255.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet255.merge_range('F21:F22', 'KELAS', header)
    worksheet255.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet255.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet255.write('G22', 'MAT', body)
    worksheet255.write('H22', 'SEJ', body)
    worksheet255.write('I22', 'GEO', body)
    worksheet255.write('J22', 'EKO', body)
    worksheet255.write('K22', 'SOS', body)
    worksheet255.write('L22', 'JML', body)
    worksheet255.write('M22', 'MAT', body)
    worksheet255.write('N22', 'SEJ', body)
    worksheet255.write('O22', 'GEO', body)
    worksheet255.write('P22', 'EKO', body)
    worksheet255.write('Q22', 'SOS', body)
    worksheet255.write('R22', 'JML', body)

    worksheet255.conditional_format(22, 0, row255+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 256
    worksheet256.insert_image('A1', r'logo resmi nf.jpg')

    worksheet256.set_column('A:A', 7, center)
    worksheet256.set_column('B:B', 6, center)
    worksheet256.set_column('C:C', 18.14, center)
    worksheet256.set_column('D:D', 25, left)
    worksheet256.set_column('E:E', 13.14, left)
    worksheet256.set_column('F:F', 8.57, center)
    worksheet256.set_column('G:R', 5, center)
    worksheet256.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIRUAS', title)
    worksheet256.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet256.write('A5', 'LOKASI', header)
    worksheet256.write('B5', 'TOTAL', header)
    worksheet256.merge_range('A4:B4', 'RANK', header)
    worksheet256.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet256.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet256.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet256.merge_range('F4:F5', 'KELAS', header)
    worksheet256.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet256.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet256.write('G5', 'MAT', body)
    worksheet256.write('H5', 'SEJ', body)
    worksheet256.write('I5', 'GEO', body)
    worksheet256.write('J5', 'EKO', body)
    worksheet256.write('K5', 'SOS', body)
    worksheet256.write('L5', 'JML', body)
    worksheet256.write('M5', 'MAT', body)
    worksheet256.write('N5', 'SEJ', body)
    worksheet256.write('O5', 'GEO', body)
    worksheet256.write('P5', 'EKO', body)
    worksheet256.write('Q5', 'SOS', body)
    worksheet256.write('R5', 'JML', body)

    worksheet256.conditional_format(5, 0, row256_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet256.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CIRUAS', title)
    worksheet256.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet256.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet256.write('A22', 'LOKASI', header)
    worksheet256.write('B22', 'TOTAL', header)
    worksheet256.merge_range('A21:B21', 'RANK', header)
    worksheet256.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet256.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet256.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet256.merge_range('F21:F22', 'KELAS', header)
    worksheet256.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet256.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet256.write('G22', 'MAT', body)
    worksheet256.write('H22', 'SEJ', body)
    worksheet256.write('I22', 'GEO', body)
    worksheet256.write('J22', 'EKO', body)
    worksheet256.write('K22', 'SOS', body)
    worksheet256.write('L22', 'JML', body)
    worksheet256.write('M22', 'MAT', body)
    worksheet256.write('N22', 'SEJ', body)
    worksheet256.write('O22', 'GEO', body)
    worksheet256.write('P22', 'EKO', body)
    worksheet256.write('Q22', 'SOS', body)
    worksheet256.write('R22', 'JML', body)

    worksheet256.conditional_format(22, 0, row256+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 258
    worksheet258.insert_image('A1', r'logo resmi nf.jpg')

    worksheet258.set_column('A:A', 7, center)
    worksheet258.set_column('B:B', 6, center)
    worksheet258.set_column('C:C', 18.14, center)
    worksheet258.set_column('D:D', 25, left)
    worksheet258.set_column('E:E', 13.14, left)
    worksheet258.set_column('F:F', 8.57, center)
    worksheet258.set_column('G:R', 5, center)
    worksheet258.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF LHOKSEUMAWE', title)
    worksheet258.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet258.write('A5', 'LOKASI', header)
    worksheet258.write('B5', 'TOTAL', header)
    worksheet258.merge_range('A4:B4', 'RANK', header)
    worksheet258.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet258.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet258.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet258.merge_range('F4:F5', 'KELAS', header)
    worksheet258.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet258.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet258.write('G5', 'MAT', body)
    worksheet258.write('H5', 'SEJ', body)
    worksheet258.write('I5', 'GEO', body)
    worksheet258.write('J5', 'EKO', body)
    worksheet258.write('K5', 'SOS', body)
    worksheet258.write('L5', 'JML', body)
    worksheet258.write('M5', 'MAT', body)
    worksheet258.write('N5', 'SEJ', body)
    worksheet258.write('O5', 'GEO', body)
    worksheet258.write('P5', 'EKO', body)
    worksheet258.write('Q5', 'SOS', body)
    worksheet258.write('R5', 'JML', body)

    worksheet258.conditional_format(5, 0, row258_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet258.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF LHOKSEUMAWE', title)
    worksheet258.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet258.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet258.write('A22', 'LOKASI', header)
    worksheet258.write('B22', 'TOTAL', header)
    worksheet258.merge_range('A21:B21', 'RANK', header)
    worksheet258.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet258.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet258.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet258.merge_range('F21:F22', 'KELAS', header)
    worksheet258.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet258.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet258.write('G22', 'MAT', body)
    worksheet258.write('H22', 'SEJ', body)
    worksheet258.write('I22', 'GEO', body)
    worksheet258.write('J22', 'EKO', body)
    worksheet258.write('K22', 'SOS', body)
    worksheet258.write('L22', 'JML', body)
    worksheet258.write('M22', 'MAT', body)
    worksheet258.write('N22', 'SEJ', body)
    worksheet258.write('O22', 'GEO', body)
    worksheet258.write('P22', 'EKO', body)
    worksheet258.write('Q22', 'SOS', body)
    worksheet258.write('R22', 'JML', body)

    worksheet258.conditional_format(22, 0, row258+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 259
    worksheet259.insert_image('A1', r'logo resmi nf.jpg')

    worksheet259.set_column('A:A', 7, center)
    worksheet259.set_column('B:B', 6, center)
    worksheet259.set_column('C:C', 18.14, center)
    worksheet259.set_column('D:D', 25, left)
    worksheet259.set_column('E:E', 13.14, left)
    worksheet259.set_column('F:F', 8.57, center)
    worksheet259.set_column('G:R', 5, center)
    worksheet259.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PANAM, PKU', title)
    worksheet259.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet259.write('A5', 'LOKASI', header)
    worksheet259.write('B5', 'TOTAL', header)
    worksheet259.merge_range('A4:B4', 'RANK', header)
    worksheet259.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet259.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet259.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet259.merge_range('F4:F5', 'KELAS', header)
    worksheet259.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet259.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet259.write('G5', 'MAT', body)
    worksheet259.write('H5', 'SEJ', body)
    worksheet259.write('I5', 'GEO', body)
    worksheet259.write('J5', 'EKO', body)
    worksheet259.write('K5', 'SOS', body)
    worksheet259.write('L5', 'JML', body)
    worksheet259.write('M5', 'MAT', body)
    worksheet259.write('N5', 'SEJ', body)
    worksheet259.write('O5', 'GEO', body)
    worksheet259.write('P5', 'EKO', body)
    worksheet259.write('Q5', 'SOS', body)
    worksheet259.write('R5', 'JML', body)

    worksheet259.conditional_format(5, 0, row259_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet259.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PANAM, PKU', title)
    worksheet259.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet259.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet259.write('A22', 'LOKASI', header)
    worksheet259.write('B22', 'TOTAL', header)
    worksheet259.merge_range('A21:B21', 'RANK', header)
    worksheet259.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet259.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet259.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet259.merge_range('F21:F22', 'KELAS', header)
    worksheet259.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet259.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet259.write('G22', 'MAT', body)
    worksheet259.write('H22', 'SEJ', body)
    worksheet259.write('I22', 'GEO', body)
    worksheet259.write('J22', 'EKO', body)
    worksheet259.write('K22', 'SOS', body)
    worksheet259.write('L22', 'JML', body)
    worksheet259.write('M22', 'MAT', body)
    worksheet259.write('N22', 'SEJ', body)
    worksheet259.write('O22', 'GEO', body)
    worksheet259.write('P22', 'EKO', body)
    worksheet259.write('Q22', 'SOS', body)
    worksheet259.write('R22', 'JML', body)

    worksheet259.conditional_format(22, 0, row259+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 260
    worksheet260.insert_image('A1', r'logo resmi nf.jpg')

    worksheet260.set_column('A:A', 7, center)
    worksheet260.set_column('B:B', 6, center)
    worksheet260.set_column('C:C', 18.14, center)
    worksheet260.set_column('D:D', 25, left)
    worksheet260.set_column('E:E', 13.14, left)
    worksheet260.set_column('F:F', 8.57, center)
    worksheet260.set_column('G:R', 5, center)
    worksheet260.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF AM. SANGAJI', title)
    worksheet260.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet260.write('A5', 'LOKASI', header)
    worksheet260.write('B5', 'TOTAL', header)
    worksheet260.merge_range('A4:B4', 'RANK', header)
    worksheet260.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet260.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet260.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet260.merge_range('F4:F5', 'KELAS', header)
    worksheet260.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet260.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet260.write('G5', 'MAT', body)
    worksheet260.write('H5', 'SEJ', body)
    worksheet260.write('I5', 'GEO', body)
    worksheet260.write('J5', 'EKO', body)
    worksheet260.write('K5', 'SOS', body)
    worksheet260.write('L5', 'JML', body)
    worksheet260.write('M5', 'MAT', body)
    worksheet260.write('N5', 'SEJ', body)
    worksheet260.write('O5', 'GEO', body)
    worksheet260.write('P5', 'EKO', body)
    worksheet260.write('Q5', 'SOS', body)
    worksheet260.write('R5', 'JML', body)

    worksheet260.conditional_format(5, 0, row260_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet260.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF AM. SANGAJI', title)
    worksheet260.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet260.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet260.write('A22', 'LOKASI', header)
    worksheet260.write('B22', 'TOTAL', header)
    worksheet260.merge_range('A21:B21', 'RANK', header)
    worksheet260.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet260.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet260.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet260.merge_range('F21:F22', 'KELAS', header)
    worksheet260.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet260.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet260.write('G22', 'MAT', body)
    worksheet260.write('H22', 'SEJ', body)
    worksheet260.write('I22', 'GEO', body)
    worksheet260.write('J22', 'EKO', body)
    worksheet260.write('K22', 'SOS', body)
    worksheet260.write('L22', 'JML', body)
    worksheet260.write('M22', 'MAT', body)
    worksheet260.write('N22', 'SEJ', body)
    worksheet260.write('O22', 'GEO', body)
    worksheet260.write('P22', 'EKO', body)
    worksheet260.write('Q22', 'SOS', body)
    worksheet260.write('R22', 'JML', body)

    worksheet260.conditional_format(22, 0, row260+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 261
    worksheet261.insert_image('A1', r'logo resmi nf.jpg')

    worksheet261.set_column('A:A', 7, center)
    worksheet261.set_column('B:B', 6, center)
    worksheet261.set_column('C:C', 18.14, center)
    worksheet261.set_column('D:D', 25, left)
    worksheet261.set_column('E:E', 13.14, left)
    worksheet261.set_column('F:F', 8.57, center)
    worksheet261.set_column('G:R', 5, center)
    worksheet261.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF DURI KOSAMBI', title)
    worksheet261.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet261.write('A5', 'LOKASI', header)
    worksheet261.write('B5', 'TOTAL', header)
    worksheet261.merge_range('A4:B4', 'RANK', header)
    worksheet261.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet261.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet261.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet261.merge_range('F4:F5', 'KELAS', header)
    worksheet261.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet261.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet261.write('G5', 'MAT', body)
    worksheet261.write('H5', 'SEJ', body)
    worksheet261.write('I5', 'GEO', body)
    worksheet261.write('J5', 'EKO', body)
    worksheet261.write('K5', 'SOS', body)
    worksheet261.write('L5', 'JML', body)
    worksheet261.write('M5', 'MAT', body)
    worksheet261.write('N5', 'SEJ', body)
    worksheet261.write('O5', 'GEO', body)
    worksheet261.write('P5', 'EKO', body)
    worksheet261.write('Q5', 'SOS', body)
    worksheet261.write('R5', 'JML', body)

    worksheet261.conditional_format(5, 0, row261_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet261.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF DURI KOSAMBI', title)
    worksheet261.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet261.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet261.write('A22', 'LOKASI', header)
    worksheet261.write('B22', 'TOTAL', header)
    worksheet261.merge_range('A21:B21', 'RANK', header)
    worksheet261.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet261.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet261.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet261.merge_range('F21:F22', 'KELAS', header)
    worksheet261.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet261.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet261.write('G22', 'MAT', body)
    worksheet261.write('H22', 'SEJ', body)
    worksheet261.write('I22', 'GEO', body)
    worksheet261.write('J22', 'EKO', body)
    worksheet261.write('K22', 'SOS', body)
    worksheet261.write('L22', 'JML', body)
    worksheet261.write('M22', 'MAT', body)
    worksheet261.write('N22', 'SEJ', body)
    worksheet261.write('O22', 'GEO', body)
    worksheet261.write('P22', 'EKO', body)
    worksheet261.write('Q22', 'SOS', body)
    worksheet261.write('R22', 'JML', body)

    worksheet261.conditional_format(22, 0, row261+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 262
    worksheet262.insert_image('A1', r'logo resmi nf.jpg')

    worksheet262.set_column('A:A', 7, center)
    worksheet262.set_column('B:B', 6, center)
    worksheet262.set_column('C:C', 18.14, center)
    worksheet262.set_column('D:D', 25, left)
    worksheet262.set_column('E:E', 13.14, left)
    worksheet262.set_column('F:F', 8.57, center)
    worksheet262.set_column('G:R', 5, center)
    worksheet262.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CITRA RAYA CIKUPA', title)
    worksheet262.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet262.write('A5', 'LOKASI', header)
    worksheet262.write('B5', 'TOTAL', header)
    worksheet262.merge_range('A4:B4', 'RANK', header)
    worksheet262.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet262.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet262.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet262.merge_range('F4:F5', 'KELAS', header)
    worksheet262.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet262.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet262.write('G5', 'MAT', body)
    worksheet262.write('H5', 'SEJ', body)
    worksheet262.write('I5', 'GEO', body)
    worksheet262.write('J5', 'EKO', body)
    worksheet262.write('K5', 'SOS', body)
    worksheet262.write('L5', 'JML', body)
    worksheet262.write('M5', 'MAT', body)
    worksheet262.write('N5', 'SEJ', body)
    worksheet262.write('O5', 'GEO', body)
    worksheet262.write('P5', 'EKO', body)
    worksheet262.write('Q5', 'SOS', body)
    worksheet262.write('R5', 'JML', body)

    worksheet262.conditional_format(5, 0, row262_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet262.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CITRA RAYA CIKUPA', title)
    worksheet262.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet262.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet262.write('A22', 'LOKASI', header)
    worksheet262.write('B22', 'TOTAL', header)
    worksheet262.merge_range('A21:B21', 'RANK', header)
    worksheet262.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet262.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet262.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet262.merge_range('F21:F22', 'KELAS', header)
    worksheet262.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet262.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet262.write('G22', 'MAT', body)
    worksheet262.write('H22', 'SEJ', body)
    worksheet262.write('I22', 'GEO', body)
    worksheet262.write('J22', 'EKO', body)
    worksheet262.write('K22', 'SOS', body)
    worksheet262.write('L22', 'JML', body)
    worksheet262.write('M22', 'MAT', body)
    worksheet262.write('N22', 'SEJ', body)
    worksheet262.write('O22', 'GEO', body)
    worksheet262.write('P22', 'EKO', body)
    worksheet262.write('Q22', 'SOS', body)
    worksheet262.write('R22', 'JML', body)

    worksheet262.conditional_format(22, 0, row262+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 264
    worksheet264.insert_image('A1', r'logo resmi nf.jpg')

    worksheet264.set_column('A:A', 7, center)
    worksheet264.set_column('B:B', 6, center)
    worksheet264.set_column('C:C', 18.14, center)
    worksheet264.set_column('D:D', 25, left)
    worksheet264.set_column('E:E', 13.14, left)
    worksheet264.set_column('F:F', 8.57, center)
    worksheet264.set_column('G:R', 5, center)
    worksheet264.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KARAWANG', title)
    worksheet264.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet264.write('A5', 'LOKASI', header)
    worksheet264.write('B5', 'TOTAL', header)
    worksheet264.merge_range('A4:B4', 'RANK', header)
    worksheet264.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet264.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet264.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet264.merge_range('F4:F5', 'KELAS', header)
    worksheet264.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet264.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet264.write('G5', 'MAT', body)
    worksheet264.write('H5', 'SEJ', body)
    worksheet264.write('I5', 'GEO', body)
    worksheet264.write('J5', 'EKO', body)
    worksheet264.write('K5', 'SOS', body)
    worksheet264.write('L5', 'JML', body)
    worksheet264.write('M5', 'MAT', body)
    worksheet264.write('N5', 'SEJ', body)
    worksheet264.write('O5', 'GEO', body)
    worksheet264.write('P5', 'EKO', body)
    worksheet264.write('Q5', 'SOS', body)
    worksheet264.write('R5', 'JML', body)

    worksheet264.conditional_format(5, 0, row264_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet264.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF KARAWANG', title)
    worksheet264.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet264.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet264.write('A22', 'LOKASI', header)
    worksheet264.write('B22', 'TOTAL', header)
    worksheet264.merge_range('A21:B21', 'RANK', header)
    worksheet264.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet264.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet264.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet264.merge_range('F21:F22', 'KELAS', header)
    worksheet264.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet264.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet264.write('G22', 'MAT', body)
    worksheet264.write('H22', 'SEJ', body)
    worksheet264.write('I22', 'GEO', body)
    worksheet264.write('J22', 'EKO', body)
    worksheet264.write('K22', 'SOS', body)
    worksheet264.write('L22', 'JML', body)
    worksheet264.write('M22', 'MAT', body)
    worksheet264.write('N22', 'SEJ', body)
    worksheet264.write('O22', 'GEO', body)
    worksheet264.write('P22', 'EKO', body)
    worksheet264.write('Q22', 'SOS', body)
    worksheet264.write('R22', 'JML', body)

    worksheet264.conditional_format(22, 0, row264+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 265
    worksheet265.insert_image('A1', r'logo resmi nf.jpg')

    worksheet265.set_column('A:A', 7, center)
    worksheet265.set_column('B:B', 6, center)
    worksheet265.set_column('C:C', 18.14, center)
    worksheet265.set_column('D:D', 25, left)
    worksheet265.set_column('E:E', 13.14, left)
    worksheet265.set_column('F:F', 8.57, center)
    worksheet265.set_column('G:R', 5, center)
    worksheet265.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF TAMAN WISMA ASRI', title)
    worksheet265.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet265.write('A5', 'LOKASI', header)
    worksheet265.write('B5', 'TOTAL', header)
    worksheet265.merge_range('A4:B4', 'RANK', header)
    worksheet265.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet265.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet265.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet265.merge_range('F4:F5', 'KELAS', header)
    worksheet265.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet265.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet265.write('G5', 'MAT', body)
    worksheet265.write('H5', 'SEJ', body)
    worksheet265.write('I5', 'GEO', body)
    worksheet265.write('J5', 'EKO', body)
    worksheet265.write('K5', 'SOS', body)
    worksheet265.write('L5', 'JML', body)
    worksheet265.write('M5', 'MAT', body)
    worksheet265.write('N5', 'SEJ', body)
    worksheet265.write('O5', 'GEO', body)
    worksheet265.write('P5', 'EKO', body)
    worksheet265.write('Q5', 'SOS', body)
    worksheet265.write('R5', 'JML', body)

    worksheet265.conditional_format(5, 0, row265_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet265.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF TAMAN WISMA ASRI', title)
    worksheet265.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet265.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet265.write('A22', 'LOKASI', header)
    worksheet265.write('B22', 'TOTAL', header)
    worksheet265.merge_range('A21:B21', 'RANK', header)
    worksheet265.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet265.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet265.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet265.merge_range('F21:F22', 'KELAS', header)
    worksheet265.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet265.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet265.write('G22', 'MAT', body)
    worksheet265.write('H22', 'SEJ', body)
    worksheet265.write('I22', 'GEO', body)
    worksheet265.write('J22', 'EKO', body)
    worksheet265.write('K22', 'SOS', body)
    worksheet265.write('L22', 'JML', body)
    worksheet265.write('M22', 'MAT', body)
    worksheet265.write('N22', 'SEJ', body)
    worksheet265.write('O22', 'GEO', body)
    worksheet265.write('P22', 'EKO', body)
    worksheet265.write('Q22', 'SOS', body)
    worksheet265.write('R22', 'JML', body)

    worksheet265.conditional_format(22, 0, row265+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 266
    worksheet266.insert_image('A1', r'logo resmi nf.jpg')

    worksheet266.set_column('A:A', 7, center)
    worksheet266.set_column('B:B', 6, center)
    worksheet266.set_column('C:C', 18.14, center)
    worksheet266.set_column('D:D', 25, left)
    worksheet266.set_column('E:E', 13.14, left)
    worksheet266.set_column('F:F', 8.57, center)
    worksheet266.set_column('G:R', 5, center)
    worksheet266.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MANGUN JAYA', title)
    worksheet266.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet266.write('A5', 'LOKASI', header)
    worksheet266.write('B5', 'TOTAL', header)
    worksheet266.merge_range('A4:B4', 'RANK', header)
    worksheet266.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet266.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet266.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet266.merge_range('F4:F5', 'KELAS', header)
    worksheet266.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet266.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet266.write('G5', 'MAT', body)
    worksheet266.write('H5', 'SEJ', body)
    worksheet266.write('I5', 'GEO', body)
    worksheet266.write('J5', 'EKO', body)
    worksheet266.write('K5', 'SOS', body)
    worksheet266.write('L5', 'JML', body)
    worksheet266.write('M5', 'MAT', body)
    worksheet266.write('N5', 'SEJ', body)
    worksheet266.write('O5', 'GEO', body)
    worksheet266.write('P5', 'EKO', body)
    worksheet266.write('Q5', 'SOS', body)
    worksheet266.write('R5', 'JML', body)

    worksheet266.conditional_format(5, 0, row266_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet266.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF MANGUN JAYA', title)
    worksheet266.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet266.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet266.write('A22', 'LOKASI', header)
    worksheet266.write('B22', 'TOTAL', header)
    worksheet266.merge_range('A21:B21', 'RANK', header)
    worksheet266.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet266.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet266.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet266.merge_range('F21:F22', 'KELAS', header)
    worksheet266.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet266.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet266.write('G22', 'MAT', body)
    worksheet266.write('H22', 'SEJ', body)
    worksheet266.write('I22', 'GEO', body)
    worksheet266.write('J22', 'EKO', body)
    worksheet266.write('K22', 'SOS', body)
    worksheet266.write('L22', 'JML', body)
    worksheet266.write('M22', 'MAT', body)
    worksheet266.write('N22', 'SEJ', body)
    worksheet266.write('O22', 'GEO', body)
    worksheet266.write('P22', 'EKO', body)
    worksheet266.write('Q22', 'SOS', body)
    worksheet266.write('R22', 'JML', body)

    worksheet266.conditional_format(22, 0, row266+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 267
    worksheet267.insert_image('A1', r'logo resmi nf.jpg')

    worksheet267.set_column('A:A', 7, center)
    worksheet267.set_column('B:B', 6, center)
    worksheet267.set_column('C:C', 18.14, center)
    worksheet267.set_column('D:D', 25, left)
    worksheet267.set_column('E:E', 13.14, left)
    worksheet267.set_column('F:F', 8.57, center)
    worksheet267.set_column('G:R', 5, center)
    worksheet267.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MARAKASH / SEKTOR 5', title)
    worksheet267.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet267.write('A5', 'LOKASI', header)
    worksheet267.write('B5', 'TOTAL', header)
    worksheet267.merge_range('A4:B4', 'RANK', header)
    worksheet267.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet267.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet267.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet267.merge_range('F4:F5', 'KELAS', header)
    worksheet267.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet267.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet267.write('G5', 'MAT', body)
    worksheet267.write('H5', 'SEJ', body)
    worksheet267.write('I5', 'GEO', body)
    worksheet267.write('J5', 'EKO', body)
    worksheet267.write('K5', 'SOS', body)
    worksheet267.write('L5', 'JML', body)
    worksheet267.write('M5', 'MAT', body)
    worksheet267.write('N5', 'SEJ', body)
    worksheet267.write('O5', 'GEO', body)
    worksheet267.write('P5', 'EKO', body)
    worksheet267.write('Q5', 'SOS', body)
    worksheet267.write('R5', 'JML', body)

    worksheet267.conditional_format(5, 0, row267_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet267.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF MARAKASH / SEKTOR 5', title)
    worksheet267.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet267.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet267.write('A22', 'LOKASI', header)
    worksheet267.write('B22', 'TOTAL', header)
    worksheet267.merge_range('A21:B21', 'RANK', header)
    worksheet267.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet267.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet267.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet267.merge_range('F21:F22', 'KELAS', header)
    worksheet267.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet267.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet267.write('G22', 'MAT', body)
    worksheet267.write('H22', 'SEJ', body)
    worksheet267.write('I22', 'GEO', body)
    worksheet267.write('J22', 'EKO', body)
    worksheet267.write('K22', 'SOS', body)
    worksheet267.write('L22', 'JML', body)
    worksheet267.write('M22', 'MAT', body)
    worksheet267.write('N22', 'SEJ', body)
    worksheet267.write('O22', 'GEO', body)
    worksheet267.write('P22', 'EKO', body)
    worksheet267.write('Q22', 'SOS', body)
    worksheet267.write('R22', 'JML', body)

    worksheet267.conditional_format(22, 0, row267+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 268
    worksheet268.insert_image('A1', r'logo resmi nf.jpg')

    worksheet268.set_column('A:A', 7, center)
    worksheet268.set_column('B:B', 6, center)
    worksheet268.set_column('C:C', 18.14, center)
    worksheet268.set_column('D:D', 25, left)
    worksheet268.set_column('E:E', 13.14, left)
    worksheet268.set_column('F:F', 8.57, center)
    worksheet268.set_column('G:R', 5, center)
    worksheet268.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KEBALEN', title)
    worksheet268.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet268.write('A5', 'LOKASI', header)
    worksheet268.write('B5', 'TOTAL', header)
    worksheet268.merge_range('A4:B4', 'RANK', header)
    worksheet268.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet268.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet268.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet268.merge_range('F4:F5', 'KELAS', header)
    worksheet268.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet268.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet268.write('G5', 'MAT', body)
    worksheet268.write('H5', 'SEJ', body)
    worksheet268.write('I5', 'GEO', body)
    worksheet268.write('J5', 'EKO', body)
    worksheet268.write('K5', 'SOS', body)
    worksheet268.write('L5', 'JML', body)
    worksheet268.write('M5', 'MAT', body)
    worksheet268.write('N5', 'SEJ', body)
    worksheet268.write('O5', 'GEO', body)
    worksheet268.write('P5', 'EKO', body)
    worksheet268.write('Q5', 'SOS', body)
    worksheet268.write('R5', 'JML', body)

    worksheet268.conditional_format(5, 0, row268_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet268.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF KEBALEN', title)
    worksheet268.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet268.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet268.write('A22', 'LOKASI', header)
    worksheet268.write('B22', 'TOTAL', header)
    worksheet268.merge_range('A21:B21', 'RANK', header)
    worksheet268.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet268.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet268.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet268.merge_range('F21:F22', 'KELAS', header)
    worksheet268.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet268.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet268.write('G22', 'MAT', body)
    worksheet268.write('H22', 'SEJ', body)
    worksheet268.write('I22', 'GEO', body)
    worksheet268.write('J22', 'EKO', body)
    worksheet268.write('K22', 'SOS', body)
    worksheet268.write('L22', 'JML', body)
    worksheet268.write('M22', 'MAT', body)
    worksheet268.write('N22', 'SEJ', body)
    worksheet268.write('O22', 'GEO', body)
    worksheet268.write('P22', 'EKO', body)
    worksheet268.write('Q22', 'SOS', body)
    worksheet268.write('R22', 'JML', body)

    worksheet268.conditional_format(22, 0, row268+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 269
    worksheet269.insert_image('A1', r'logo resmi nf.jpg')

    worksheet269.set_column('A:A', 7, center)
    worksheet269.set_column('B:B', 6, center)
    worksheet269.set_column('C:C', 18.14, center)
    worksheet269.set_column('D:D', 25, left)
    worksheet269.set_column('E:E', 13.14, left)
    worksheet269.set_column('F:F', 8.57, center)
    worksheet269.set_column('G:R', 5, center)
    worksheet269.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF JATI RANGON', title)
    worksheet269.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet269.write('A5', 'LOKASI', header)
    worksheet269.write('B5', 'TOTAL', header)
    worksheet269.merge_range('A4:B4', 'RANK', header)
    worksheet269.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet269.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet269.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet269.merge_range('F4:F5', 'KELAS', header)
    worksheet269.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet269.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet269.write('G5', 'MAT', body)
    worksheet269.write('H5', 'SEJ', body)
    worksheet269.write('I5', 'GEO', body)
    worksheet269.write('J5', 'EKO', body)
    worksheet269.write('K5', 'SOS', body)
    worksheet269.write('L5', 'JML', body)
    worksheet269.write('M5', 'MAT', body)
    worksheet269.write('N5', 'SEJ', body)
    worksheet269.write('O5', 'GEO', body)
    worksheet269.write('P5', 'EKO', body)
    worksheet269.write('Q5', 'SOS', body)
    worksheet269.write('R5', 'JML', body)

    worksheet269.conditional_format(5, 0, row269_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet269.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF JATI RANGON', title)
    worksheet269.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet269.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet269.write('A22', 'LOKASI', header)
    worksheet269.write('B22', 'TOTAL', header)
    worksheet269.merge_range('A21:B21', 'RANK', header)
    worksheet269.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet269.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet269.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet269.merge_range('F21:F22', 'KELAS', header)
    worksheet269.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet269.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet269.write('G22', 'MAT', body)
    worksheet269.write('H22', 'SEJ', body)
    worksheet269.write('I22', 'GEO', body)
    worksheet269.write('J22', 'EKO', body)
    worksheet269.write('K22', 'SOS', body)
    worksheet269.write('L22', 'JML', body)
    worksheet269.write('M22', 'MAT', body)
    worksheet269.write('N22', 'SEJ', body)
    worksheet269.write('O22', 'GEO', body)
    worksheet269.write('P22', 'EKO', body)
    worksheet269.write('Q22', 'SOS', body)
    worksheet269.write('R22', 'JML', body)

    worksheet269.conditional_format(22, 0, row269+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 270
    worksheet270.insert_image('A1', r'logo resmi nf.jpg')

    worksheet270.set_column('A:A', 7, center)
    worksheet270.set_column('B:B', 6, center)
    worksheet270.set_column('C:C', 18.14, center)
    worksheet270.set_column('D:D', 25, left)
    worksheet270.set_column('E:E', 13.14, left)
    worksheet270.set_column('F:F', 8.57, center)
    worksheet270.set_column('G:R', 5, center)
    worksheet270.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF JATIBENING', title)
    worksheet270.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet270.write('A5', 'LOKASI', header)
    worksheet270.write('B5', 'TOTAL', header)
    worksheet270.merge_range('A4:B4', 'RANK', header)
    worksheet270.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet270.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet270.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet270.merge_range('F4:F5', 'KELAS', header)
    worksheet270.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet270.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet270.write('G5', 'MAT', body)
    worksheet270.write('H5', 'SEJ', body)
    worksheet270.write('I5', 'GEO', body)
    worksheet270.write('J5', 'EKO', body)
    worksheet270.write('K5', 'SOS', body)
    worksheet270.write('L5', 'JML', body)
    worksheet270.write('M5', 'MAT', body)
    worksheet270.write('N5', 'SEJ', body)
    worksheet270.write('O5', 'GEO', body)
    worksheet270.write('P5', 'EKO', body)
    worksheet270.write('Q5', 'SOS', body)
    worksheet270.write('R5', 'JML', body)

    worksheet270.conditional_format(5, 0, row270_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet270.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF JATIBENING', title)
    worksheet270.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet270.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet270.write('A22', 'LOKASI', header)
    worksheet270.write('B22', 'TOTAL', header)
    worksheet270.merge_range('A21:B21', 'RANK', header)
    worksheet270.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet270.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet270.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet270.merge_range('F21:F22', 'KELAS', header)
    worksheet270.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet270.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet270.write('G22', 'MAT', body)
    worksheet270.write('H22', 'SEJ', body)
    worksheet270.write('I22', 'GEO', body)
    worksheet270.write('J22', 'EKO', body)
    worksheet270.write('K22', 'SOS', body)
    worksheet270.write('L22', 'JML', body)
    worksheet270.write('M22', 'MAT', body)
    worksheet270.write('N22', 'SEJ', body)
    worksheet270.write('O22', 'GEO', body)
    worksheet270.write('P22', 'EKO', body)
    worksheet270.write('Q22', 'SOS', body)
    worksheet270.write('R22', 'JML', body)

    worksheet270.conditional_format(22, 0, row270+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 271
    worksheet271.insert_image('A1', r'logo resmi nf.jpg')

    worksheet271.set_column('A:A', 7, center)
    worksheet271.set_column('B:B', 6, center)
    worksheet271.set_column('C:C', 18.14, center)
    worksheet271.set_column('D:D', 25, left)
    worksheet271.set_column('E:E', 13.14, left)
    worksheet271.set_column('F:F', 8.57, center)
    worksheet271.set_column('G:R', 5, center)
    worksheet271.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF JATIMULYA', title)
    worksheet271.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet271.write('A5', 'LOKASI', header)
    worksheet271.write('B5', 'TOTAL', header)
    worksheet271.merge_range('A4:B4', 'RANK', header)
    worksheet271.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet271.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet271.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet271.merge_range('F4:F5', 'KELAS', header)
    worksheet271.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet271.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet271.write('G5', 'MAT', body)
    worksheet271.write('H5', 'SEJ', body)
    worksheet271.write('I5', 'GEO', body)
    worksheet271.write('J5', 'EKO', body)
    worksheet271.write('K5', 'SOS', body)
    worksheet271.write('L5', 'JML', body)
    worksheet271.write('M5', 'MAT', body)
    worksheet271.write('N5', 'SEJ', body)
    worksheet271.write('O5', 'GEO', body)
    worksheet271.write('P5', 'EKO', body)
    worksheet271.write('Q5', 'SOS', body)
    worksheet271.write('R5', 'JML', body)

    worksheet271.conditional_format(5, 0, row271_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet271.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF JATIMULYA', title)
    worksheet271.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet271.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet271.write('A22', 'LOKASI', header)
    worksheet271.write('B22', 'TOTAL', header)
    worksheet271.merge_range('A21:B21', 'RANK', header)
    worksheet271.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet271.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet271.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet271.merge_range('F21:F22', 'KELAS', header)
    worksheet271.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet271.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet271.write('G22', 'MAT', body)
    worksheet271.write('H22', 'SEJ', body)
    worksheet271.write('I22', 'GEO', body)
    worksheet271.write('J22', 'EKO', body)
    worksheet271.write('K22', 'SOS', body)
    worksheet271.write('L22', 'JML', body)
    worksheet271.write('M22', 'MAT', body)
    worksheet271.write('N22', 'SEJ', body)
    worksheet271.write('O22', 'GEO', body)
    worksheet271.write('P22', 'EKO', body)
    worksheet271.write('Q22', 'SOS', body)
    worksheet271.write('R22', 'JML', body)

    worksheet271.conditional_format(22, 0, row271+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 272
    worksheet272.insert_image('A1', r'logo resmi nf.jpg')

    worksheet272.set_column('A:A', 7, center)
    worksheet272.set_column('B:B', 6, center)
    worksheet272.set_column('C:C', 18.14, center)
    worksheet272.set_column('D:D', 25, left)
    worksheet272.set_column('E:E', 13.14, left)
    worksheet272.set_column('F:F', 8.57, center)
    worksheet272.set_column('G:R', 5, center)
    worksheet272.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PERUMNAS 3', title)
    worksheet272.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet272.write('A5', 'LOKASI', header)
    worksheet272.write('B5', 'TOTAL', header)
    worksheet272.merge_range('A4:B4', 'RANK', header)
    worksheet272.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet272.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet272.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet272.merge_range('F4:F5', 'KELAS', header)
    worksheet272.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet272.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet272.write('G5', 'MAT', body)
    worksheet272.write('H5', 'SEJ', body)
    worksheet272.write('I5', 'GEO', body)
    worksheet272.write('J5', 'EKO', body)
    worksheet272.write('K5', 'SOS', body)
    worksheet272.write('L5', 'JML', body)
    worksheet272.write('M5', 'MAT', body)
    worksheet272.write('N5', 'SEJ', body)
    worksheet272.write('O5', 'GEO', body)
    worksheet272.write('P5', 'EKO', body)
    worksheet272.write('Q5', 'SOS', body)
    worksheet272.write('R5', 'JML', body)

    worksheet272.conditional_format(5, 0, row272_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet272.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PERUMNAS 3', title)
    worksheet272.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet272.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet272.write('A22', 'LOKASI', header)
    worksheet272.write('B22', 'TOTAL', header)
    worksheet272.merge_range('A21:B21', 'RANK', header)
    worksheet272.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet272.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet272.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet272.merge_range('F21:F22', 'KELAS', header)
    worksheet272.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet272.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet272.write('G22', 'MAT', body)
    worksheet272.write('H22', 'SEJ', body)
    worksheet272.write('I22', 'GEO', body)
    worksheet272.write('J22', 'EKO', body)
    worksheet272.write('K22', 'SOS', body)
    worksheet272.write('L22', 'JML', body)
    worksheet272.write('M22', 'MAT', body)
    worksheet272.write('N22', 'SEJ', body)
    worksheet272.write('O22', 'GEO', body)
    worksheet272.write('P22', 'EKO', body)
    worksheet272.write('Q22', 'SOS', body)
    worksheet272.write('R22', 'JML', body)

    worksheet272.conditional_format(22, 0, row272+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 273
    worksheet273.insert_image('A1', r'logo resmi nf.jpg')

    worksheet273.set_column('A:A', 7, center)
    worksheet273.set_column('B:B', 6, center)
    worksheet273.set_column('C:C', 18.14, center)
    worksheet273.set_column('D:D', 25, left)
    worksheet273.set_column('E:E', 13.14, left)
    worksheet273.set_column('F:F', 8.57, center)
    worksheet273.set_column('G:R', 5, center)
    worksheet273.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF NAROGONG', title)
    worksheet273.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet273.write('A5', 'LOKASI', header)
    worksheet273.write('B5', 'TOTAL', header)
    worksheet273.merge_range('A4:B4', 'RANK', header)
    worksheet273.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet273.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet273.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet273.merge_range('F4:F5', 'KELAS', header)
    worksheet273.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet273.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet273.write('G5', 'MAT', body)
    worksheet273.write('H5', 'SEJ', body)
    worksheet273.write('I5', 'GEO', body)
    worksheet273.write('J5', 'EKO', body)
    worksheet273.write('K5', 'SOS', body)
    worksheet273.write('L5', 'JML', body)
    worksheet273.write('M5', 'MAT', body)
    worksheet273.write('N5', 'SEJ', body)
    worksheet273.write('O5', 'GEO', body)
    worksheet273.write('P5', 'EKO', body)
    worksheet273.write('Q5', 'SOS', body)
    worksheet273.write('R5', 'JML', body)

    worksheet273.conditional_format(5, 0, row273_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet273.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF NAROGONG', title)
    worksheet273.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet273.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet273.write('A22', 'LOKASI', header)
    worksheet273.write('B22', 'TOTAL', header)
    worksheet273.merge_range('A21:B21', 'RANK', header)
    worksheet273.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet273.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet273.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet273.merge_range('F21:F22', 'KELAS', header)
    worksheet273.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet273.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet273.write('G22', 'MAT', body)
    worksheet273.write('H22', 'SEJ', body)
    worksheet273.write('I22', 'GEO', body)
    worksheet273.write('J22', 'EKO', body)
    worksheet273.write('K22', 'SOS', body)
    worksheet273.write('L22', 'JML', body)
    worksheet273.write('M22', 'MAT', body)
    worksheet273.write('N22', 'SEJ', body)
    worksheet273.write('O22', 'GEO', body)
    worksheet273.write('P22', 'EKO', body)
    worksheet273.write('Q22', 'SOS', body)
    worksheet273.write('R22', 'JML', body)

    worksheet273.conditional_format(22, 0, row273+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 274
    worksheet274.insert_image('A1', r'logo resmi nf.jpg')

    worksheet274.set_column('A:A', 7, center)
    worksheet274.set_column('B:B', 6, center)
    worksheet274.set_column('C:C', 18.14, center)
    worksheet274.set_column('D:D', 25, left)
    worksheet274.set_column('E:E', 13.14, left)
    worksheet274.set_column('F:F', 8.57, center)
    worksheet274.set_column('G:R', 5, center)
    worksheet274.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BEKASI TIMUR REGENCY', title)
    worksheet274.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet274.write('A5', 'LOKASI', header)
    worksheet274.write('B5', 'TOTAL', header)
    worksheet274.merge_range('A4:B4', 'RANK', header)
    worksheet274.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet274.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet274.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet274.merge_range('F4:F5', 'KELAS', header)
    worksheet274.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet274.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet274.write('G5', 'MAT', body)
    worksheet274.write('H5', 'SEJ', body)
    worksheet274.write('I5', 'GEO', body)
    worksheet274.write('J5', 'EKO', body)
    worksheet274.write('K5', 'SOS', body)
    worksheet274.write('L5', 'JML', body)
    worksheet274.write('M5', 'MAT', body)
    worksheet274.write('N5', 'SEJ', body)
    worksheet274.write('O5', 'GEO', body)
    worksheet274.write('P5', 'EKO', body)
    worksheet274.write('Q5', 'SOS', body)
    worksheet274.write('R5', 'JML', body)

    worksheet274.conditional_format(5, 0, row274_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet274.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF BEKASI TIMUR REGENCY', title)
    worksheet274.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet274.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet274.write('A22', 'LOKASI', header)
    worksheet274.write('B22', 'TOTAL', header)
    worksheet274.merge_range('A21:B21', 'RANK', header)
    worksheet274.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet274.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet274.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet274.merge_range('F21:F22', 'KELAS', header)
    worksheet274.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet274.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet274.write('G22', 'MAT', body)
    worksheet274.write('H22', 'SEJ', body)
    worksheet274.write('I22', 'GEO', body)
    worksheet274.write('J22', 'EKO', body)
    worksheet274.write('K22', 'SOS', body)
    worksheet274.write('L22', 'JML', body)
    worksheet274.write('M22', 'MAT', body)
    worksheet274.write('N22', 'SEJ', body)
    worksheet274.write('O22', 'GEO', body)
    worksheet274.write('P22', 'EKO', body)
    worksheet274.write('Q22', 'SOS', body)
    worksheet274.write('R22', 'JML', body)

    worksheet274.conditional_format(22, 0, row274+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 275
    worksheet275.insert_image('A1', r'logo resmi nf.jpg')

    worksheet275.set_column('A:A', 7, center)
    worksheet275.set_column('B:B', 6, center)
    worksheet275.set_column('C:C', 18.14, center)
    worksheet275.set_column('D:D', 25, left)
    worksheet275.set_column('E:E', 13.14, left)
    worksheet275.set_column('F:F', 8.57, center)
    worksheet275.set_column('G:R', 5, center)
    worksheet275.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIKARANG PILAR', title)
    worksheet275.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet275.write('A5', 'LOKASI', header)
    worksheet275.write('B5', 'TOTAL', header)
    worksheet275.merge_range('A4:B4', 'RANK', header)
    worksheet275.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet275.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet275.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet275.merge_range('F4:F5', 'KELAS', header)
    worksheet275.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet275.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet275.write('G5', 'MAT', body)
    worksheet275.write('H5', 'SEJ', body)
    worksheet275.write('I5', 'GEO', body)
    worksheet275.write('J5', 'EKO', body)
    worksheet275.write('K5', 'SOS', body)
    worksheet275.write('L5', 'JML', body)
    worksheet275.write('M5', 'MAT', body)
    worksheet275.write('N5', 'SEJ', body)
    worksheet275.write('O5', 'GEO', body)
    worksheet275.write('P5', 'EKO', body)
    worksheet275.write('Q5', 'SOS', body)
    worksheet275.write('R5', 'JML', body)

    worksheet275.conditional_format(5, 0, row275_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet275.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CIKARANG PILAR', title)
    worksheet275.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet275.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet275.write('A22', 'LOKASI', header)
    worksheet275.write('B22', 'TOTAL', header)
    worksheet275.merge_range('A21:B21', 'RANK', header)
    worksheet275.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet275.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet275.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet275.merge_range('F21:F22', 'KELAS', header)
    worksheet275.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet275.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet275.write('G22', 'MAT', body)
    worksheet275.write('H22', 'SEJ', body)
    worksheet275.write('I22', 'GEO', body)
    worksheet275.write('J22', 'EKO', body)
    worksheet275.write('K22', 'SOS', body)
    worksheet275.write('L22', 'JML', body)
    worksheet275.write('M22', 'MAT', body)
    worksheet275.write('N22', 'SEJ', body)
    worksheet275.write('O22', 'GEO', body)
    worksheet275.write('P22', 'EKO', body)
    worksheet275.write('Q22', 'SOS', body)
    worksheet275.write('R22', 'JML', body)

    worksheet275.conditional_format(22, 0, row275+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 276
    worksheet276.insert_image('A1', r'logo resmi nf.jpg')

    worksheet276.set_column('A:A', 7, center)
    worksheet276.set_column('B:B', 6, center)
    worksheet276.set_column('C:C', 18.14, center)
    worksheet276.set_column('D:D', 25, left)
    worksheet276.set_column('E:E', 13.14, left)
    worksheet276.set_column('F:F', 8.57, center)
    worksheet276.set_column('G:R', 5, center)
    worksheet276.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIKARANG JABABEKA', title)
    worksheet276.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet276.write('A5', 'LOKASI', header)
    worksheet276.write('B5', 'TOTAL', header)
    worksheet276.merge_range('A4:B4', 'RANK', header)
    worksheet276.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet276.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet276.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet276.merge_range('F4:F5', 'KELAS', header)
    worksheet276.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet276.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet276.write('G5', 'MAT', body)
    worksheet276.write('H5', 'SEJ', body)
    worksheet276.write('I5', 'GEO', body)
    worksheet276.write('J5', 'EKO', body)
    worksheet276.write('K5', 'SOS', body)
    worksheet276.write('L5', 'JML', body)
    worksheet276.write('M5', 'MAT', body)
    worksheet276.write('N5', 'SEJ', body)
    worksheet276.write('O5', 'GEO', body)
    worksheet276.write('P5', 'EKO', body)
    worksheet276.write('Q5', 'SOS', body)
    worksheet276.write('R5', 'JML', body)

    worksheet276.conditional_format(5, 0, row276_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet276.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CIKARANG JABABEKA', title)
    worksheet276.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet276.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet276.write('A22', 'LOKASI', header)
    worksheet276.write('B22', 'TOTAL', header)
    worksheet276.merge_range('A21:B21', 'RANK', header)
    worksheet276.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet276.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet276.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet276.merge_range('F21:F22', 'KELAS', header)
    worksheet276.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet276.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet276.write('G22', 'MAT', body)
    worksheet276.write('H22', 'SEJ', body)
    worksheet276.write('I22', 'GEO', body)
    worksheet276.write('J22', 'EKO', body)
    worksheet276.write('K22', 'SOS', body)
    worksheet276.write('L22', 'JML', body)
    worksheet276.write('M22', 'MAT', body)
    worksheet276.write('N22', 'SEJ', body)
    worksheet276.write('O22', 'GEO', body)
    worksheet276.write('P22', 'EKO', body)
    worksheet276.write('Q22', 'SOS', body)
    worksheet276.write('R22', 'JML', body)

    worksheet276.conditional_format(22, 0, row276+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 277
    worksheet277.insert_image('A1', r'logo resmi nf.jpg')

    worksheet277.set_column('A:A', 7, center)
    worksheet277.set_column('B:B', 6, center)
    worksheet277.set_column('C:C', 18.14, center)
    worksheet277.set_column('D:D', 25, left)
    worksheet277.set_column('E:E', 13.14, left)
    worksheet277.set_column('F:F', 8.57, center)
    worksheet277.set_column('G:R', 5, center)
    worksheet277.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PAYAKUMBUH', title)
    worksheet277.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet277.write('A5', 'LOKASI', header)
    worksheet277.write('B5', 'TOTAL', header)
    worksheet277.merge_range('A4:B4', 'RANK', header)
    worksheet277.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet277.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet277.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet277.merge_range('F4:F5', 'KELAS', header)
    worksheet277.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet277.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet277.write('G5', 'MAT', body)
    worksheet277.write('H5', 'SEJ', body)
    worksheet277.write('I5', 'GEO', body)
    worksheet277.write('J5', 'EKO', body)
    worksheet277.write('K5', 'SOS', body)
    worksheet277.write('L5', 'JML', body)
    worksheet277.write('M5', 'MAT', body)
    worksheet277.write('N5', 'SEJ', body)
    worksheet277.write('O5', 'GEO', body)
    worksheet277.write('P5', 'EKO', body)
    worksheet277.write('Q5', 'SOS', body)
    worksheet277.write('R5', 'JML', body)

    worksheet277.conditional_format(5, 0, row277_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet277.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PAYAKUMBUH', title)
    worksheet277.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet277.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet277.write('A22', 'LOKASI', header)
    worksheet277.write('B22', 'TOTAL', header)
    worksheet277.merge_range('A21:B21', 'RANK', header)
    worksheet277.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet277.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet277.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet277.merge_range('F21:F22', 'KELAS', header)
    worksheet277.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet277.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet277.write('G22', 'MAT', body)
    worksheet277.write('H22', 'SEJ', body)
    worksheet277.write('I22', 'GEO', body)
    worksheet277.write('J22', 'EKO', body)
    worksheet277.write('K22', 'SOS', body)
    worksheet277.write('L22', 'JML', body)
    worksheet277.write('M22', 'MAT', body)
    worksheet277.write('N22', 'SEJ', body)
    worksheet277.write('O22', 'GEO', body)
    worksheet277.write('P22', 'EKO', body)
    worksheet277.write('Q22', 'SOS', body)
    worksheet277.write('R22', 'JML', body)

    worksheet277.conditional_format(22, 0, row277+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 278
    worksheet278.insert_image('A1', r'logo resmi nf.jpg')

    worksheet278.set_column('A:A', 7, center)
    worksheet278.set_column('B:B', 6, center)
    worksheet278.set_column('C:C', 18.14, center)
    worksheet278.set_column('D:D', 25, left)
    worksheet278.set_column('E:E', 13.14, left)
    worksheet278.set_column('F:F', 8.57, center)
    worksheet278.set_column('G:R', 5, center)
    worksheet278.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MERDUATI', title)
    worksheet278.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet278.write('A5', 'LOKASI', header)
    worksheet278.write('B5', 'TOTAL', header)
    worksheet278.merge_range('A4:B4', 'RANK', header)
    worksheet278.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet278.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet278.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet278.merge_range('F4:F5', 'KELAS', header)
    worksheet278.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet278.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet278.write('G5', 'MAT', body)
    worksheet278.write('H5', 'SEJ', body)
    worksheet278.write('I5', 'GEO', body)
    worksheet278.write('J5', 'EKO', body)
    worksheet278.write('K5', 'SOS', body)
    worksheet278.write('L5', 'JML', body)
    worksheet278.write('M5', 'MAT', body)
    worksheet278.write('N5', 'SEJ', body)
    worksheet278.write('O5', 'GEO', body)
    worksheet278.write('P5', 'EKO', body)
    worksheet278.write('Q5', 'SOS', body)
    worksheet278.write('R5', 'JML', body)

    worksheet278.conditional_format(5, 0, row278_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet278.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF MERDUATI', title)
    worksheet278.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet278.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet278.write('A22', 'LOKASI', header)
    worksheet278.write('B22', 'TOTAL', header)
    worksheet278.merge_range('A21:B21', 'RANK', header)
    worksheet278.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet278.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet278.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet278.merge_range('F21:F22', 'KELAS', header)
    worksheet278.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet278.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet278.write('G22', 'MAT', body)
    worksheet278.write('H22', 'SEJ', body)
    worksheet278.write('I22', 'GEO', body)
    worksheet278.write('J22', 'EKO', body)
    worksheet278.write('K22', 'SOS', body)
    worksheet278.write('L22', 'JML', body)
    worksheet278.write('M22', 'MAT', body)
    worksheet278.write('N22', 'SEJ', body)
    worksheet278.write('O22', 'GEO', body)
    worksheet278.write('P22', 'EKO', body)
    worksheet278.write('Q22', 'SOS', body)
    worksheet278.write('R22', 'JML', body)

    worksheet278.conditional_format(22, 0, row278+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 279
    worksheet279.insert_image('A1', r'logo resmi nf.jpg')

    worksheet279.set_column('A:A', 7, center)
    worksheet279.set_column('B:B', 6, center)
    worksheet279.set_column('C:C', 18.14, center)
    worksheet279.set_column('D:D', 25, left)
    worksheet279.set_column('E:E', 13.14, left)
    worksheet279.set_column('F:F', 8.57, center)
    worksheet279.set_column('G:R', 5, center)
    worksheet279.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF ANTAPANI', title)
    worksheet279.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet279.write('A5', 'LOKASI', header)
    worksheet279.write('B5', 'TOTAL', header)
    worksheet279.merge_range('A4:B4', 'RANK', header)
    worksheet279.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet279.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet279.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet279.merge_range('F4:F5', 'KELAS', header)
    worksheet279.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet279.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet279.write('G5', 'MAT', body)
    worksheet279.write('H5', 'SEJ', body)
    worksheet279.write('I5', 'GEO', body)
    worksheet279.write('J5', 'EKO', body)
    worksheet279.write('K5', 'SOS', body)
    worksheet279.write('L5', 'JML', body)
    worksheet279.write('M5', 'MAT', body)
    worksheet279.write('N5', 'SEJ', body)
    worksheet279.write('O5', 'GEO', body)
    worksheet279.write('P5', 'EKO', body)
    worksheet279.write('Q5', 'SOS', body)
    worksheet279.write('R5', 'JML', body)

    worksheet279.conditional_format(5, 0, row279_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet279.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF ANTAPANI', title)
    worksheet279.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet279.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet279.write('A22', 'LOKASI', header)
    worksheet279.write('B22', 'TOTAL', header)
    worksheet279.merge_range('A21:B21', 'RANK', header)
    worksheet279.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet279.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet279.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet279.merge_range('F21:F22', 'KELAS', header)
    worksheet279.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet279.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet279.write('G22', 'MAT', body)
    worksheet279.write('H22', 'SEJ', body)
    worksheet279.write('I22', 'GEO', body)
    worksheet279.write('J22', 'EKO', body)
    worksheet279.write('K22', 'SOS', body)
    worksheet279.write('L22', 'JML', body)
    worksheet279.write('M22', 'MAT', body)
    worksheet279.write('N22', 'SEJ', body)
    worksheet279.write('O22', 'GEO', body)
    worksheet279.write('P22', 'EKO', body)
    worksheet279.write('Q22', 'SOS', body)
    worksheet279.write('R22', 'JML', body)

    worksheet279.conditional_format(22, 0, row279+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 280
    worksheet280.insert_image('A1', r'logo resmi nf.jpg')

    worksheet280.set_column('A:A', 7, center)
    worksheet280.set_column('B:B', 6, center)
    worksheet280.set_column('C:C', 18.14, center)
    worksheet280.set_column('D:D', 25, left)
    worksheet280.set_column('E:E', 13.14, left)
    worksheet280.set_column('F:F', 8.57, center)
    worksheet280.set_column('G:R', 5, center)
    worksheet280.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MARGAHAYU', title)
    worksheet280.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet280.write('A5', 'LOKASI', header)
    worksheet280.write('B5', 'TOTAL', header)
    worksheet280.merge_range('A4:B4', 'RANK', header)
    worksheet280.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet280.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet280.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet280.merge_range('F4:F5', 'KELAS', header)
    worksheet280.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet280.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet280.write('G5', 'MAT', body)
    worksheet280.write('H5', 'SEJ', body)
    worksheet280.write('I5', 'GEO', body)
    worksheet280.write('J5', 'EKO', body)
    worksheet280.write('K5', 'SOS', body)
    worksheet280.write('L5', 'JML', body)
    worksheet280.write('M5', 'MAT', body)
    worksheet280.write('N5', 'SEJ', body)
    worksheet280.write('O5', 'GEO', body)
    worksheet280.write('P5', 'EKO', body)
    worksheet280.write('Q5', 'SOS', body)
    worksheet280.write('R5', 'JML', body)

    worksheet280.conditional_format(5, 0, row280_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet280.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF MARGAHAYU', title)
    worksheet280.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet280.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet280.write('A22', 'LOKASI', header)
    worksheet280.write('B22', 'TOTAL', header)
    worksheet280.merge_range('A21:B21', 'RANK', header)
    worksheet280.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet280.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet280.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet280.merge_range('F21:F22', 'KELAS', header)
    worksheet280.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet280.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet280.write('G22', 'MAT', body)
    worksheet280.write('H22', 'SEJ', body)
    worksheet280.write('I22', 'GEO', body)
    worksheet280.write('J22', 'EKO', body)
    worksheet280.write('K22', 'SOS', body)
    worksheet280.write('L22', 'JML', body)
    worksheet280.write('M22', 'MAT', body)
    worksheet280.write('N22', 'SEJ', body)
    worksheet280.write('O22', 'GEO', body)
    worksheet280.write('P22', 'EKO', body)
    worksheet280.write('Q22', 'SOS', body)
    worksheet280.write('R22', 'JML', body)

    worksheet280.conditional_format(22, 0, row280+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 281
    worksheet281.insert_image('A1', r'logo resmi nf.jpg')

    worksheet281.set_column('A:A', 7, center)
    worksheet281.set_column('B:B', 6, center)
    worksheet281.set_column('C:C', 18.14, center)
    worksheet281.set_column('D:D', 25, left)
    worksheet281.set_column('E:E', 13.14, left)
    worksheet281.set_column('F:F', 8.57, center)
    worksheet281.set_column('G:R', 5, center)
    worksheet281.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF RAJAWALI', title)
    worksheet281.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet281.write('A5', 'LOKASI', header)
    worksheet281.write('B5', 'TOTAL', header)
    worksheet281.merge_range('A4:B4', 'RANK', header)
    worksheet281.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet281.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet281.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet281.merge_range('F4:F5', 'KELAS', header)
    worksheet281.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet281.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet281.write('G5', 'MAT', body)
    worksheet281.write('H5', 'SEJ', body)
    worksheet281.write('I5', 'GEO', body)
    worksheet281.write('J5', 'EKO', body)
    worksheet281.write('K5', 'SOS', body)
    worksheet281.write('L5', 'JML', body)
    worksheet281.write('M5', 'MAT', body)
    worksheet281.write('N5', 'SEJ', body)
    worksheet281.write('O5', 'GEO', body)
    worksheet281.write('P5', 'EKO', body)
    worksheet281.write('Q5', 'SOS', body)
    worksheet281.write('R5', 'JML', body)

    worksheet281.conditional_format(5, 0, row281_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet281.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF RAJAWALI', title)
    worksheet281.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet281.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet281.write('A22', 'LOKASI', header)
    worksheet281.write('B22', 'TOTAL', header)
    worksheet281.merge_range('A21:B21', 'RANK', header)
    worksheet281.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet281.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet281.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet281.merge_range('F21:F22', 'KELAS', header)
    worksheet281.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet281.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet281.write('G22', 'MAT', body)
    worksheet281.write('H22', 'SEJ', body)
    worksheet281.write('I22', 'GEO', body)
    worksheet281.write('J22', 'EKO', body)
    worksheet281.write('K22', 'SOS', body)
    worksheet281.write('L22', 'JML', body)
    worksheet281.write('M22', 'MAT', body)
    worksheet281.write('N22', 'SEJ', body)
    worksheet281.write('O22', 'GEO', body)
    worksheet281.write('P22', 'EKO', body)
    worksheet281.write('Q22', 'SOS', body)
    worksheet281.write('R22', 'JML', body)

    worksheet281.conditional_format(22, 0, row281+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 282
    worksheet282.insert_image('A1', r'logo resmi nf.jpg')

    worksheet282.set_column('A:A', 7, center)
    worksheet282.set_column('B:B', 6, center)
    worksheet282.set_column('C:C', 18.14, center)
    worksheet282.set_column('D:D', 25, left)
    worksheet282.set_column('E:E', 13.14, left)
    worksheet282.set_column('F:F', 8.57, center)
    worksheet282.set_column('G:R', 5, center)
    worksheet282.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PAHLAWAN', title)
    worksheet282.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet282.write('A5', 'LOKASI', header)
    worksheet282.write('B5', 'TOTAL', header)
    worksheet282.merge_range('A4:B4', 'RANK', header)
    worksheet282.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet282.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet282.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet282.merge_range('F4:F5', 'KELAS', header)
    worksheet282.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet282.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet282.write('G5', 'MAT', body)
    worksheet282.write('H5', 'SEJ', body)
    worksheet282.write('I5', 'GEO', body)
    worksheet282.write('J5', 'EKO', body)
    worksheet282.write('K5', 'SOS', body)
    worksheet282.write('L5', 'JML', body)
    worksheet282.write('M5', 'MAT', body)
    worksheet282.write('N5', 'SEJ', body)
    worksheet282.write('O5', 'GEO', body)
    worksheet282.write('P5', 'EKO', body)
    worksheet282.write('Q5', 'SOS', body)
    worksheet282.write('R5', 'JML', body)

    worksheet282.conditional_format(5, 0, row282_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet282.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PAHLAWAN', title)
    worksheet282.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet282.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet282.write('A22', 'LOKASI', header)
    worksheet282.write('B22', 'TOTAL', header)
    worksheet282.merge_range('A21:B21', 'RANK', header)
    worksheet282.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet282.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet282.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet282.merge_range('F21:F22', 'KELAS', header)
    worksheet282.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet282.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet282.write('G22', 'MAT', body)
    worksheet282.write('H22', 'SEJ', body)
    worksheet282.write('I22', 'GEO', body)
    worksheet282.write('J22', 'EKO', body)
    worksheet282.write('K22', 'SOS', body)
    worksheet282.write('L22', 'JML', body)
    worksheet282.write('M22', 'MAT', body)
    worksheet282.write('N22', 'SEJ', body)
    worksheet282.write('O22', 'GEO', body)
    worksheet282.write('P22', 'EKO', body)
    worksheet282.write('Q22', 'SOS', body)
    worksheet282.write('R22', 'JML', body)

    worksheet282.conditional_format(22, 0, row282+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 283
    worksheet283.insert_image('A1', r'logo resmi nf.jpg')

    worksheet283.set_column('A:A', 7, center)
    worksheet283.set_column('B:B', 6, center)
    worksheet283.set_column('C:C', 18.14, center)
    worksheet283.set_column('D:D', 25, left)
    worksheet283.set_column('E:E', 13.14, left)
    worksheet283.set_column('F:F', 8.57, center)
    worksheet283.set_column('G:R', 5, center)
    worksheet283.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIJERAH', title)
    worksheet283.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet283.write('A5', 'LOKASI', header)
    worksheet283.write('B5', 'TOTAL', header)
    worksheet283.merge_range('A4:B4', 'RANK', header)
    worksheet283.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet283.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet283.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet283.merge_range('F4:F5', 'KELAS', header)
    worksheet283.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet283.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet283.write('G5', 'MAT', body)
    worksheet283.write('H5', 'SEJ', body)
    worksheet283.write('I5', 'GEO', body)
    worksheet283.write('J5', 'EKO', body)
    worksheet283.write('K5', 'SOS', body)
    worksheet283.write('L5', 'JML', body)
    worksheet283.write('M5', 'MAT', body)
    worksheet283.write('N5', 'SEJ', body)
    worksheet283.write('O5', 'GEO', body)
    worksheet283.write('P5', 'EKO', body)
    worksheet283.write('Q5', 'SOS', body)
    worksheet283.write('R5', 'JML', body)

    worksheet283.conditional_format(5, 0, row283_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet283.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CIJERAH', title)
    worksheet283.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet283.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet283.write('A22', 'LOKASI', header)
    worksheet283.write('B22', 'TOTAL', header)
    worksheet283.merge_range('A21:B21', 'RANK', header)
    worksheet283.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet283.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet283.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet283.merge_range('F21:F22', 'KELAS', header)
    worksheet283.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet283.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet283.write('G22', 'MAT', body)
    worksheet283.write('H22', 'SEJ', body)
    worksheet283.write('I22', 'GEO', body)
    worksheet283.write('J22', 'EKO', body)
    worksheet283.write('K22', 'SOS', body)
    worksheet283.write('L22', 'JML', body)
    worksheet283.write('M22', 'MAT', body)
    worksheet283.write('N22', 'SEJ', body)
    worksheet283.write('O22', 'GEO', body)
    worksheet283.write('P22', 'EKO', body)
    worksheet283.write('Q22', 'SOS', body)
    worksheet283.write('R22', 'JML', body)

    worksheet283.conditional_format(22, 0, row283+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 284
    worksheet284.insert_image('A1', r'logo resmi nf.jpg')

    worksheet284.set_column('A:A', 7, center)
    worksheet284.set_column('B:B', 6, center)
    worksheet284.set_column('C:C', 18.14, center)
    worksheet284.set_column('D:D', 25, left)
    worksheet284.set_column('E:E', 13.14, left)
    worksheet284.set_column('F:F', 8.57, center)
    worksheet284.set_column('G:R', 5, center)
    worksheet284.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF TEGAL', title)
    worksheet284.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet284.write('A5', 'LOKASI', header)
    worksheet284.write('B5', 'TOTAL', header)
    worksheet284.merge_range('A4:B4', 'RANK', header)
    worksheet284.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet284.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet284.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet284.merge_range('F4:F5', 'KELAS', header)
    worksheet284.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet284.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet284.write('G5', 'MAT', body)
    worksheet284.write('H5', 'SEJ', body)
    worksheet284.write('I5', 'GEO', body)
    worksheet284.write('J5', 'EKO', body)
    worksheet284.write('K5', 'SOS', body)
    worksheet284.write('L5', 'JML', body)
    worksheet284.write('M5', 'MAT', body)
    worksheet284.write('N5', 'SEJ', body)
    worksheet284.write('O5', 'GEO', body)
    worksheet284.write('P5', 'EKO', body)
    worksheet284.write('Q5', 'SOS', body)
    worksheet284.write('R5', 'JML', body)

    worksheet284.conditional_format(5, 0, row284_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet284.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF TEGAL', title)
    worksheet284.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet284.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet284.write('A22', 'LOKASI', header)
    worksheet284.write('B22', 'TOTAL', header)
    worksheet284.merge_range('A21:B21', 'RANK', header)
    worksheet284.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet284.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet284.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet284.merge_range('F21:F22', 'KELAS', header)
    worksheet284.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet284.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet284.write('G22', 'MAT', body)
    worksheet284.write('H22', 'SEJ', body)
    worksheet284.write('I22', 'GEO', body)
    worksheet284.write('J22', 'EKO', body)
    worksheet284.write('K22', 'SOS', body)
    worksheet284.write('L22', 'JML', body)
    worksheet284.write('M22', 'MAT', body)
    worksheet284.write('N22', 'SEJ', body)
    worksheet284.write('O22', 'GEO', body)
    worksheet284.write('P22', 'EKO', body)
    worksheet284.write('Q22', 'SOS', body)
    worksheet284.write('R22', 'JML', body)

    worksheet284.conditional_format(22, 0, row284+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 285
    worksheet285.insert_image('A1', r'logo resmi nf.jpg')

    worksheet285.set_column('A:A', 7, center)
    worksheet285.set_column('B:B', 6, center)
    worksheet285.set_column('C:C', 18.14, center)
    worksheet285.set_column('D:D', 25, left)
    worksheet285.set_column('E:E', 13.14, left)
    worksheet285.set_column('F:F', 8.57, center)
    worksheet285.set_column('G:R', 5, center)
    worksheet285.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MEDAN AREA', title)
    worksheet285.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet285.write('A5', 'LOKASI', header)
    worksheet285.write('B5', 'TOTAL', header)
    worksheet285.merge_range('A4:B4', 'RANK', header)
    worksheet285.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet285.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet285.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet285.merge_range('F4:F5', 'KELAS', header)
    worksheet285.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet285.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet285.write('G5', 'MAT', body)
    worksheet285.write('H5', 'SEJ', body)
    worksheet285.write('I5', 'GEO', body)
    worksheet285.write('J5', 'EKO', body)
    worksheet285.write('K5', 'SOS', body)
    worksheet285.write('L5', 'JML', body)
    worksheet285.write('M5', 'MAT', body)
    worksheet285.write('N5', 'SEJ', body)
    worksheet285.write('O5', 'GEO', body)
    worksheet285.write('P5', 'EKO', body)
    worksheet285.write('Q5', 'SOS', body)
    worksheet285.write('R5', 'JML', body)

    worksheet285.conditional_format(5, 0, row285_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet285.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF MEDAN AREA', title)
    worksheet285.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet285.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet285.write('A22', 'LOKASI', header)
    worksheet285.write('B22', 'TOTAL', header)
    worksheet285.merge_range('A21:B21', 'RANK', header)
    worksheet285.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet285.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet285.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet285.merge_range('F21:F22', 'KELAS', header)
    worksheet285.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet285.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet285.write('G22', 'MAT', body)
    worksheet285.write('H22', 'SEJ', body)
    worksheet285.write('I22', 'GEO', body)
    worksheet285.write('J22', 'EKO', body)
    worksheet285.write('K22', 'SOS', body)
    worksheet285.write('L22', 'JML', body)
    worksheet285.write('M22', 'MAT', body)
    worksheet285.write('N22', 'SEJ', body)
    worksheet285.write('O22', 'GEO', body)
    worksheet285.write('P22', 'EKO', body)
    worksheet285.write('Q22', 'SOS', body)
    worksheet285.write('R22', 'JML', body)

    worksheet285.conditional_format(22, 0, row285+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 286
    worksheet286.insert_image('A1', r'logo resmi nf.jpg')

    worksheet286.set_column('A:A', 7, center)
    worksheet286.set_column('B:B', 6, center)
    worksheet286.set_column('C:C', 18.14, center)
    worksheet286.set_column('D:D', 25, left)
    worksheet286.set_column('E:E', 13.14, left)
    worksheet286.set_column('F:F', 8.57, center)
    worksheet286.set_column('G:R', 5, center)
    worksheet286.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MEDAN JOHOR', title)
    worksheet286.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet286.write('A5', 'LOKASI', header)
    worksheet286.write('B5', 'TOTAL', header)
    worksheet286.merge_range('A4:B4', 'RANK', header)
    worksheet286.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet286.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet286.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet286.merge_range('F4:F5', 'KELAS', header)
    worksheet286.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet286.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet286.write('G5', 'MAT', body)
    worksheet286.write('H5', 'SEJ', body)
    worksheet286.write('I5', 'GEO', body)
    worksheet286.write('J5', 'EKO', body)
    worksheet286.write('K5', 'SOS', body)
    worksheet286.write('L5', 'JML', body)
    worksheet286.write('M5', 'MAT', body)
    worksheet286.write('N5', 'SEJ', body)
    worksheet286.write('O5', 'GEO', body)
    worksheet286.write('P5', 'EKO', body)
    worksheet286.write('Q5', 'SOS', body)
    worksheet286.write('R5', 'JML', body)

    worksheet286.conditional_format(5, 0, row286_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet286.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF MEDAN JOHOR', title)
    worksheet286.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet286.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet286.write('A22', 'LOKASI', header)
    worksheet286.write('B22', 'TOTAL', header)
    worksheet286.merge_range('A21:B21', 'RANK', header)
    worksheet286.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet286.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet286.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet286.merge_range('F21:F22', 'KELAS', header)
    worksheet286.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet286.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet286.write('G22', 'MAT', body)
    worksheet286.write('H22', 'SEJ', body)
    worksheet286.write('I22', 'GEO', body)
    worksheet286.write('J22', 'EKO', body)
    worksheet286.write('K22', 'SOS', body)
    worksheet286.write('L22', 'JML', body)
    worksheet286.write('M22', 'MAT', body)
    worksheet286.write('N22', 'SEJ', body)
    worksheet286.write('O22', 'GEO', body)
    worksheet286.write('P22', 'EKO', body)
    worksheet286.write('Q22', 'SOS', body)
    worksheet286.write('R22', 'JML', body)

    worksheet286.conditional_format(22, 0, row286+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 287
    worksheet287.insert_image('A1', r'logo resmi nf.jpg')

    worksheet287.set_column('A:A', 7, center)
    worksheet287.set_column('B:B', 6, center)
    worksheet287.set_column('C:C', 18.14, center)
    worksheet287.set_column('D:D', 25, left)
    worksheet287.set_column('E:E', 13.14, left)
    worksheet287.set_column('F:F', 8.57, center)
    worksheet287.set_column('G:R', 5, center)
    worksheet287.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF JAMBO TAPE', title)
    worksheet287.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet287.write('A5', 'LOKASI', header)
    worksheet287.write('B5', 'TOTAL', header)
    worksheet287.merge_range('A4:B4', 'RANK', header)
    worksheet287.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet287.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet287.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet287.merge_range('F4:F5', 'KELAS', header)
    worksheet287.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet287.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet287.write('G5', 'MAT', body)
    worksheet287.write('H5', 'SEJ', body)
    worksheet287.write('I5', 'GEO', body)
    worksheet287.write('J5', 'EKO', body)
    worksheet287.write('K5', 'SOS', body)
    worksheet287.write('L5', 'JML', body)
    worksheet287.write('M5', 'MAT', body)
    worksheet287.write('N5', 'SEJ', body)
    worksheet287.write('O5', 'GEO', body)
    worksheet287.write('P5', 'EKO', body)
    worksheet287.write('Q5', 'SOS', body)
    worksheet287.write('R5', 'JML', body)

    worksheet287.conditional_format(5, 0, row287_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet287.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF JAMBO TAPE', title)
    worksheet287.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet287.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet287.write('A22', 'LOKASI', header)
    worksheet287.write('B22', 'TOTAL', header)
    worksheet287.merge_range('A21:B21', 'RANK', header)
    worksheet287.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet287.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet287.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet287.merge_range('F21:F22', 'KELAS', header)
    worksheet287.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet287.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet287.write('G22', 'MAT', body)
    worksheet287.write('H22', 'SEJ', body)
    worksheet287.write('I22', 'GEO', body)
    worksheet287.write('J22', 'EKO', body)
    worksheet287.write('K22', 'SOS', body)
    worksheet287.write('L22', 'JML', body)
    worksheet287.write('M22', 'MAT', body)
    worksheet287.write('N22', 'SEJ', body)
    worksheet287.write('O22', 'GEO', body)
    worksheet287.write('P22', 'EKO', body)
    worksheet287.write('Q22', 'SOS', body)
    worksheet287.write('R22', 'JML', body)

    worksheet287.conditional_format(22, 0, row287+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 288
    worksheet288.insert_image('A1', r'logo resmi nf.jpg')

    worksheet288.set_column('A:A', 7, center)
    worksheet288.set_column('B:B', 6, center)
    worksheet288.set_column('C:C', 18.14, center)
    worksheet288.set_column('D:D', 25, left)
    worksheet288.set_column('E:E', 13.14, left)
    worksheet288.set_column('F:F', 8.57, center)
    worksheet288.set_column('G:R', 5, center)
    worksheet288.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF THE HOK', title)
    worksheet288.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet288.write('A5', 'LOKASI', header)
    worksheet288.write('B5', 'TOTAL', header)
    worksheet288.merge_range('A4:B4', 'RANK', header)
    worksheet288.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet288.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet288.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet288.merge_range('F4:F5', 'KELAS', header)
    worksheet288.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet288.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet288.write('G5', 'MAT', body)
    worksheet288.write('H5', 'SEJ', body)
    worksheet288.write('I5', 'GEO', body)
    worksheet288.write('J5', 'EKO', body)
    worksheet288.write('K5', 'SOS', body)
    worksheet288.write('L5', 'JML', body)
    worksheet288.write('M5', 'MAT', body)
    worksheet288.write('N5', 'SEJ', body)
    worksheet288.write('O5', 'GEO', body)
    worksheet288.write('P5', 'EKO', body)
    worksheet288.write('Q5', 'SOS', body)
    worksheet288.write('R5', 'JML', body)

    worksheet288.conditional_format(5, 0, row288_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet288.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF THE HOK', title)
    worksheet288.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet288.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet288.write('A22', 'LOKASI', header)
    worksheet288.write('B22', 'TOTAL', header)
    worksheet288.merge_range('A21:B21', 'RANK', header)
    worksheet288.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet288.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet288.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet288.merge_range('F21:F22', 'KELAS', header)
    worksheet288.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet288.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet288.write('G22', 'MAT', body)
    worksheet288.write('H22', 'SEJ', body)
    worksheet288.write('I22', 'GEO', body)
    worksheet288.write('J22', 'EKO', body)
    worksheet288.write('K22', 'SOS', body)
    worksheet288.write('L22', 'JML', body)
    worksheet288.write('M22', 'MAT', body)
    worksheet288.write('N22', 'SEJ', body)
    worksheet288.write('O22', 'GEO', body)
    worksheet288.write('P22', 'EKO', body)
    worksheet288.write('Q22', 'SOS', body)
    worksheet288.write('R22', 'JML', body)

    worksheet288.conditional_format(22, 0, row288+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 289
    worksheet289.insert_image('A1', r'logo resmi nf.jpg')

    worksheet289.set_column('A:A', 7, center)
    worksheet289.set_column('B:B', 6, center)
    worksheet289.set_column('C:C', 18.14, center)
    worksheet289.set_column('D:D', 25, left)
    worksheet289.set_column('E:E', 13.14, left)
    worksheet289.set_column('F:F', 8.57, center)
    worksheet289.set_column('G:R', 5, center)
    worksheet289.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SAIL', title)
    worksheet289.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet289.write('A5', 'LOKASI', header)
    worksheet289.write('B5', 'TOTAL', header)
    worksheet289.merge_range('A4:B4', 'RANK', header)
    worksheet289.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet289.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet289.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet289.merge_range('F4:F5', 'KELAS', header)
    worksheet289.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet289.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet289.write('G5', 'MAT', body)
    worksheet289.write('H5', 'SEJ', body)
    worksheet289.write('I5', 'GEO', body)
    worksheet289.write('J5', 'EKO', body)
    worksheet289.write('K5', 'SOS', body)
    worksheet289.write('L5', 'JML', body)
    worksheet289.write('M5', 'MAT', body)
    worksheet289.write('N5', 'SEJ', body)
    worksheet289.write('O5', 'GEO', body)
    worksheet289.write('P5', 'EKO', body)
    worksheet289.write('Q5', 'SOS', body)
    worksheet289.write('R5', 'JML', body)

    worksheet289.conditional_format(5, 0, row289_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet289.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF SAIL', title)
    worksheet289.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet289.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet289.write('A22', 'LOKASI', header)
    worksheet289.write('B22', 'TOTAL', header)
    worksheet289.merge_range('A21:B21', 'RANK', header)
    worksheet289.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet289.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet289.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet289.merge_range('F21:F22', 'KELAS', header)
    worksheet289.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet289.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet289.write('G22', 'MAT', body)
    worksheet289.write('H22', 'SEJ', body)
    worksheet289.write('I22', 'GEO', body)
    worksheet289.write('J22', 'EKO', body)
    worksheet289.write('K22', 'SOS', body)
    worksheet289.write('L22', 'JML', body)
    worksheet289.write('M22', 'MAT', body)
    worksheet289.write('N22', 'SEJ', body)
    worksheet289.write('O22', 'GEO', body)
    worksheet289.write('P22', 'EKO', body)
    worksheet289.write('Q22', 'SOS', body)
    worksheet289.write('R22', 'JML', body)

    worksheet289.conditional_format(22, 0, row289+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 290
    worksheet290.insert_image('A1', r'logo resmi nf.jpg')

    worksheet290.set_column('A:A', 7, center)
    worksheet290.set_column('B:B', 6, center)
    worksheet290.set_column('C:C', 18.14, center)
    worksheet290.set_column('D:D', 25, left)
    worksheet290.set_column('E:E', 13.14, left)
    worksheet290.set_column('F:F', 8.57, center)
    worksheet290.set_column('G:R', 5, center)
    worksheet290.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF TELANAI JAMBI', title)
    worksheet290.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet290.write('A5', 'LOKASI', header)
    worksheet290.write('B5', 'TOTAL', header)
    worksheet290.merge_range('A4:B4', 'RANK', header)
    worksheet290.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet290.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet290.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet290.merge_range('F4:F5', 'KELAS', header)
    worksheet290.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet290.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet290.write('G5', 'MAT', body)
    worksheet290.write('H5', 'SEJ', body)
    worksheet290.write('I5', 'GEO', body)
    worksheet290.write('J5', 'EKO', body)
    worksheet290.write('K5', 'SOS', body)
    worksheet290.write('L5', 'JML', body)
    worksheet290.write('M5', 'MAT', body)
    worksheet290.write('N5', 'SEJ', body)
    worksheet290.write('O5', 'GEO', body)
    worksheet290.write('P5', 'EKO', body)
    worksheet290.write('Q5', 'SOS', body)
    worksheet290.write('R5', 'JML', body)

    worksheet290.conditional_format(5, 0, row290_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet290.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF TELANAI JAMBI', title)
    worksheet290.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet290.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet290.write('A22', 'LOKASI', header)
    worksheet290.write('B22', 'TOTAL', header)
    worksheet290.merge_range('A21:B21', 'RANK', header)
    worksheet290.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet290.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet290.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet290.merge_range('F21:F22', 'KELAS', header)
    worksheet290.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet290.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet290.write('G22', 'MAT', body)
    worksheet290.write('H22', 'SEJ', body)
    worksheet290.write('I22', 'GEO', body)
    worksheet290.write('J22', 'EKO', body)
    worksheet290.write('K22', 'SOS', body)
    worksheet290.write('L22', 'JML', body)
    worksheet290.write('M22', 'MAT', body)
    worksheet290.write('N22', 'SEJ', body)
    worksheet290.write('O22', 'GEO', body)
    worksheet290.write('P22', 'EKO', body)
    worksheet290.write('Q22', 'SOS', body)
    worksheet290.write('R22', 'JML', body)

    worksheet290.conditional_format(22, 0, row290+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 291
    worksheet291.insert_image('A1', r'logo resmi nf.jpg')

    worksheet291.set_column('A:A', 7, center)
    worksheet291.set_column('B:B', 6, center)
    worksheet291.set_column('C:C', 18.14, center)
    worksheet291.set_column('D:D', 25, left)
    worksheet291.set_column('E:E', 13.14, left)
    worksheet291.set_column('F:F', 8.57, center)
    worksheet291.set_column('G:R', 5, center)
    worksheet291.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SIDOARJO', title)
    worksheet291.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet291.write('A5', 'LOKASI', header)
    worksheet291.write('B5', 'TOTAL', header)
    worksheet291.merge_range('A4:B4', 'RANK', header)
    worksheet291.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet291.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet291.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet291.merge_range('F4:F5', 'KELAS', header)
    worksheet291.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet291.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet291.write('G5', 'MAT', body)
    worksheet291.write('H5', 'SEJ', body)
    worksheet291.write('I5', 'GEO', body)
    worksheet291.write('J5', 'EKO', body)
    worksheet291.write('K5', 'SOS', body)
    worksheet291.write('L5', 'JML', body)
    worksheet291.write('M5', 'MAT', body)
    worksheet291.write('N5', 'SEJ', body)
    worksheet291.write('O5', 'GEO', body)
    worksheet291.write('P5', 'EKO', body)
    worksheet291.write('Q5', 'SOS', body)
    worksheet291.write('R5', 'JML', body)

    worksheet291.conditional_format(5, 0, row291_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet291.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF SIDOARJO', title)
    worksheet291.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet291.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet291.write('A22', 'LOKASI', header)
    worksheet291.write('B22', 'TOTAL', header)
    worksheet291.merge_range('A21:B21', 'RANK', header)
    worksheet291.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet291.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet291.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet291.merge_range('F21:F22', 'KELAS', header)
    worksheet291.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet291.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet291.write('G22', 'MAT', body)
    worksheet291.write('H22', 'SEJ', body)
    worksheet291.write('I22', 'GEO', body)
    worksheet291.write('J22', 'EKO', body)
    worksheet291.write('K22', 'SOS', body)
    worksheet291.write('L22', 'JML', body)
    worksheet291.write('M22', 'MAT', body)
    worksheet291.write('N22', 'SEJ', body)
    worksheet291.write('O22', 'GEO', body)
    worksheet291.write('P22', 'EKO', body)
    worksheet291.write('Q22', 'SOS', body)
    worksheet291.write('R22', 'JML', body)

    worksheet291.conditional_format(22, 0, row291+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 292
    worksheet292.insert_image('A1', r'logo resmi nf.jpg')

    worksheet292.set_column('A:A', 7, center)
    worksheet292.set_column('B:B', 6, center)
    worksheet292.set_column('C:C', 18.14, center)
    worksheet292.set_column('D:D', 25, left)
    worksheet292.set_column('E:E', 13.14, left)
    worksheet292.set_column('F:F', 8.57, center)
    worksheet292.set_column('G:R', 5, center)
    worksheet292.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PURWOKERTO LOR', title)
    worksheet292.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet292.write('A5', 'LOKASI', header)
    worksheet292.write('B5', 'TOTAL', header)
    worksheet292.merge_range('A4:B4', 'RANK', header)
    worksheet292.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet292.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet292.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet292.merge_range('F4:F5', 'KELAS', header)
    worksheet292.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet292.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet292.write('G5', 'MAT', body)
    worksheet292.write('H5', 'SEJ', body)
    worksheet292.write('I5', 'GEO', body)
    worksheet292.write('J5', 'EKO', body)
    worksheet292.write('K5', 'SOS', body)
    worksheet292.write('L5', 'JML', body)
    worksheet292.write('M5', 'MAT', body)
    worksheet292.write('N5', 'SEJ', body)
    worksheet292.write('O5', 'GEO', body)
    worksheet292.write('P5', 'EKO', body)
    worksheet292.write('Q5', 'SOS', body)
    worksheet292.write('R5', 'JML', body)

    worksheet292.conditional_format(5, 0, row292_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet292.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PURWOKERTO LOR', title)
    worksheet292.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet292.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet292.write('A22', 'LOKASI', header)
    worksheet292.write('B22', 'TOTAL', header)
    worksheet292.merge_range('A21:B21', 'RANK', header)
    worksheet292.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet292.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet292.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet292.merge_range('F21:F22', 'KELAS', header)
    worksheet292.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet292.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet292.write('G22', 'MAT', body)
    worksheet292.write('H22', 'SEJ', body)
    worksheet292.write('I22', 'GEO', body)
    worksheet292.write('J22', 'EKO', body)
    worksheet292.write('K22', 'SOS', body)
    worksheet292.write('L22', 'JML', body)
    worksheet292.write('M22', 'MAT', body)
    worksheet292.write('N22', 'SEJ', body)
    worksheet292.write('O22', 'GEO', body)
    worksheet292.write('P22', 'EKO', body)
    worksheet292.write('Q22', 'SOS', body)
    worksheet292.write('R22', 'JML', body)

    worksheet292.conditional_format(22, 0, row292+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 293
    worksheet293.insert_image('A1', r'logo resmi nf.jpg')

    worksheet293.set_column('A:A', 7, center)
    worksheet293.set_column('B:B', 6, center)
    worksheet293.set_column('C:C', 18.14, center)
    worksheet293.set_column('D:D', 25, left)
    worksheet293.set_column('E:E', 13.14, left)
    worksheet293.set_column('F:F', 8.57, center)
    worksheet293.set_column('G:R', 5, center)
    worksheet293.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF WAY HALIM', title)
    worksheet293.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet293.write('A5', 'LOKASI', header)
    worksheet293.write('B5', 'TOTAL', header)
    worksheet293.merge_range('A4:B4', 'RANK', header)
    worksheet293.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet293.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet293.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet293.merge_range('F4:F5', 'KELAS', header)
    worksheet293.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet293.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet293.write('G5', 'MAT', body)
    worksheet293.write('H5', 'SEJ', body)
    worksheet293.write('I5', 'GEO', body)
    worksheet293.write('J5', 'EKO', body)
    worksheet293.write('K5', 'SOS', body)
    worksheet293.write('L5', 'JML', body)
    worksheet293.write('M5', 'MAT', body)
    worksheet293.write('N5', 'SEJ', body)
    worksheet293.write('O5', 'GEO', body)
    worksheet293.write('P5', 'EKO', body)
    worksheet293.write('Q5', 'SOS', body)
    worksheet293.write('R5', 'JML', body)

    worksheet293.conditional_format(5, 0, row293_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet293.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF WAY HALIM', title)
    worksheet293.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet293.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet293.write('A22', 'LOKASI', header)
    worksheet293.write('B22', 'TOTAL', header)
    worksheet293.merge_range('A21:B21', 'RANK', header)
    worksheet293.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet293.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet293.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet293.merge_range('F21:F22', 'KELAS', header)
    worksheet293.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet293.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet293.write('G22', 'MAT', body)
    worksheet293.write('H22', 'SEJ', body)
    worksheet293.write('I22', 'GEO', body)
    worksheet293.write('J22', 'EKO', body)
    worksheet293.write('K22', 'SOS', body)
    worksheet293.write('L22', 'JML', body)
    worksheet293.write('M22', 'MAT', body)
    worksheet293.write('N22', 'SEJ', body)
    worksheet293.write('O22', 'GEO', body)
    worksheet293.write('P22', 'EKO', body)
    worksheet293.write('Q22', 'SOS', body)
    worksheet293.write('R22', 'JML', body)

    worksheet293.conditional_format(22, 0, row293+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 294
    worksheet294.insert_image('A1', r'logo resmi nf.jpg')

    worksheet294.set_column('A:A', 7, center)
    worksheet294.set_column('B:B', 6, center)
    worksheet294.set_column('C:C', 18.14, center)
    worksheet294.set_column('D:D', 25, left)
    worksheet294.set_column('E:E', 13.14, left)
    worksheet294.set_column('F:F', 8.57, center)
    worksheet294.set_column('G:R', 5, center)
    worksheet294.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF METRO', title)
    worksheet294.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet294.write('A5', 'LOKASI', header)
    worksheet294.write('B5', 'TOTAL', header)
    worksheet294.merge_range('A4:B4', 'RANK', header)
    worksheet294.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet294.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet294.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet294.merge_range('F4:F5', 'KELAS', header)
    worksheet294.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet294.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet294.write('G5', 'MAT', body)
    worksheet294.write('H5', 'SEJ', body)
    worksheet294.write('I5', 'GEO', body)
    worksheet294.write('J5', 'EKO', body)
    worksheet294.write('K5', 'SOS', body)
    worksheet294.write('L5', 'JML', body)
    worksheet294.write('M5', 'MAT', body)
    worksheet294.write('N5', 'SEJ', body)
    worksheet294.write('O5', 'GEO', body)
    worksheet294.write('P5', 'EKO', body)
    worksheet294.write('Q5', 'SOS', body)
    worksheet294.write('R5', 'JML', body)

    worksheet294.conditional_format(5, 0, row294_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet294.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF METRO', title)
    worksheet294.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet294.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet294.write('A22', 'LOKASI', header)
    worksheet294.write('B22', 'TOTAL', header)
    worksheet294.merge_range('A21:B21', 'RANK', header)
    worksheet294.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet294.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet294.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet294.merge_range('F21:F22', 'KELAS', header)
    worksheet294.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet294.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet294.write('G22', 'MAT', body)
    worksheet294.write('H22', 'SEJ', body)
    worksheet294.write('I22', 'GEO', body)
    worksheet294.write('J22', 'EKO', body)
    worksheet294.write('K22', 'SOS', body)
    worksheet294.write('L22', 'JML', body)
    worksheet294.write('M22', 'MAT', body)
    worksheet294.write('N22', 'SEJ', body)
    worksheet294.write('O22', 'GEO', body)
    worksheet294.write('P22', 'EKO', body)
    worksheet294.write('Q22', 'SOS', body)
    worksheet294.write('R22', 'JML', body)

    worksheet294.conditional_format(22, 0, row294+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 295
    worksheet295.insert_image('A1', r'logo resmi nf.jpg')

    worksheet295.set_column('A:A', 7, center)
    worksheet295.set_column('B:B', 6, center)
    worksheet295.set_column('C:C', 18.14, center)
    worksheet295.set_column('D:D', 25, left)
    worksheet295.set_column('E:E', 13.14, left)
    worksheet295.set_column('F:F', 8.57, center)
    worksheet295.set_column('G:R', 5, center)
    worksheet295.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF RAJABASA', title)
    worksheet295.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet295.write('A5', 'LOKASI', header)
    worksheet295.write('B5', 'TOTAL', header)
    worksheet295.merge_range('A4:B4', 'RANK', header)
    worksheet295.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet295.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet295.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet295.merge_range('F4:F5', 'KELAS', header)
    worksheet295.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet295.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet295.write('G5', 'MAT', body)
    worksheet295.write('H5', 'SEJ', body)
    worksheet295.write('I5', 'GEO', body)
    worksheet295.write('J5', 'EKO', body)
    worksheet295.write('K5', 'SOS', body)
    worksheet295.write('L5', 'JML', body)
    worksheet295.write('M5', 'MAT', body)
    worksheet295.write('N5', 'SEJ', body)
    worksheet295.write('O5', 'GEO', body)
    worksheet295.write('P5', 'EKO', body)
    worksheet295.write('Q5', 'SOS', body)
    worksheet295.write('R5', 'JML', body)

    worksheet295.conditional_format(5, 0, row295_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet295.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF RAJABASA', title)
    worksheet295.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet295.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet295.write('A22', 'LOKASI', header)
    worksheet295.write('B22', 'TOTAL', header)
    worksheet295.merge_range('A21:B21', 'RANK', header)
    worksheet295.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet295.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet295.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet295.merge_range('F21:F22', 'KELAS', header)
    worksheet295.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet295.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet295.write('G22', 'MAT', body)
    worksheet295.write('H22', 'SEJ', body)
    worksheet295.write('I22', 'GEO', body)
    worksheet295.write('J22', 'EKO', body)
    worksheet295.write('K22', 'SOS', body)
    worksheet295.write('L22', 'JML', body)
    worksheet295.write('M22', 'MAT', body)
    worksheet295.write('N22', 'SEJ', body)
    worksheet295.write('O22', 'GEO', body)
    worksheet295.write('P22', 'EKO', body)
    worksheet295.write('Q22', 'SOS', body)
    worksheet295.write('R22', 'JML', body)

    worksheet295.conditional_format(22, 0, row295+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 296
    worksheet296.insert_image('A1', r'logo resmi nf.jpg')

    worksheet296.set_column('A:A', 7, center)
    worksheet296.set_column('B:B', 6, center)
    worksheet296.set_column('C:C', 18.14, center)
    worksheet296.set_column('D:D', 25, left)
    worksheet296.set_column('E:E', 13.14, left)
    worksheet296.set_column('F:F', 8.57, center)
    worksheet296.set_column('G:R', 5, center)
    worksheet296.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KEDATON', title)
    worksheet296.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet296.write('A5', 'LOKASI', header)
    worksheet296.write('B5', 'TOTAL', header)
    worksheet296.merge_range('A4:B4', 'RANK', header)
    worksheet296.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet296.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet296.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet296.merge_range('F4:F5', 'KELAS', header)
    worksheet296.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet296.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet296.write('G5', 'MAT', body)
    worksheet296.write('H5', 'SEJ', body)
    worksheet296.write('I5', 'GEO', body)
    worksheet296.write('J5', 'EKO', body)
    worksheet296.write('K5', 'SOS', body)
    worksheet296.write('L5', 'JML', body)
    worksheet296.write('M5', 'MAT', body)
    worksheet296.write('N5', 'SEJ', body)
    worksheet296.write('O5', 'GEO', body)
    worksheet296.write('P5', 'EKO', body)
    worksheet296.write('Q5', 'SOS', body)
    worksheet296.write('R5', 'JML', body)

    worksheet296.conditional_format(5, 0, row296_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet296.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF KEDATON', title)
    worksheet296.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet296.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet296.write('A22', 'LOKASI', header)
    worksheet296.write('B22', 'TOTAL', header)
    worksheet296.merge_range('A21:B21', 'RANK', header)
    worksheet296.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet296.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet296.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet296.merge_range('F21:F22', 'KELAS', header)
    worksheet296.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet296.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet296.write('G22', 'MAT', body)
    worksheet296.write('H22', 'SEJ', body)
    worksheet296.write('I22', 'GEO', body)
    worksheet296.write('J22', 'EKO', body)
    worksheet296.write('K22', 'SOS', body)
    worksheet296.write('L22', 'JML', body)
    worksheet296.write('M22', 'MAT', body)
    worksheet296.write('N22', 'SEJ', body)
    worksheet296.write('O22', 'GEO', body)
    worksheet296.write('P22', 'EKO', body)
    worksheet296.write('Q22', 'SOS', body)
    worksheet296.write('R22', 'JML', body)

    worksheet296.conditional_format(22, 0, row296+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 297
    worksheet297.insert_image('A1', r'logo resmi nf.jpg')

    worksheet297.set_column('A:A', 7, center)
    worksheet297.set_column('B:B', 6, center)
    worksheet297.set_column('C:C', 18.14, center)
    worksheet297.set_column('D:D', 25, left)
    worksheet297.set_column('E:E', 13.14, left)
    worksheet297.set_column('F:F', 8.57, center)
    worksheet297.set_column('G:R', 5, center)
    worksheet297.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PALAPA CUT NYAK DIEN', title)
    worksheet297.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet297.write('A5', 'LOKASI', header)
    worksheet297.write('B5', 'TOTAL', header)
    worksheet297.merge_range('A4:B4', 'RANK', header)
    worksheet297.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet297.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet297.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet297.merge_range('F4:F5', 'KELAS', header)
    worksheet297.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet297.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet297.write('G5', 'MAT', body)
    worksheet297.write('H5', 'SEJ', body)
    worksheet297.write('I5', 'GEO', body)
    worksheet297.write('J5', 'EKO', body)
    worksheet297.write('K5', 'SOS', body)
    worksheet297.write('L5', 'JML', body)
    worksheet297.write('M5', 'MAT', body)
    worksheet297.write('N5', 'SEJ', body)
    worksheet297.write('O5', 'GEO', body)
    worksheet297.write('P5', 'EKO', body)
    worksheet297.write('Q5', 'SOS', body)
    worksheet297.write('R5', 'JML', body)

    worksheet297.conditional_format(5, 0, row297_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet297.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PALAPA CUT NYAK DIEN', title)
    worksheet297.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet297.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet297.write('A22', 'LOKASI', header)
    worksheet297.write('B22', 'TOTAL', header)
    worksheet297.merge_range('A21:B21', 'RANK', header)
    worksheet297.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet297.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet297.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet297.merge_range('F21:F22', 'KELAS', header)
    worksheet297.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet297.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet297.write('G22', 'MAT', body)
    worksheet297.write('H22', 'SEJ', body)
    worksheet297.write('I22', 'GEO', body)
    worksheet297.write('J22', 'EKO', body)
    worksheet297.write('K22', 'SOS', body)
    worksheet297.write('L22', 'JML', body)
    worksheet297.write('M22', 'MAT', body)
    worksheet297.write('N22', 'SEJ', body)
    worksheet297.write('O22', 'GEO', body)
    worksheet297.write('P22', 'EKO', body)
    worksheet297.write('Q22', 'SOS', body)
    worksheet297.write('R22', 'JML', body)

    worksheet297.conditional_format(22, 0, row297+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 298
    worksheet298.insert_image('A1', r'logo resmi nf.jpg')

    worksheet298.set_column('A:A', 7, center)
    worksheet298.set_column('B:B', 6, center)
    worksheet298.set_column('C:C', 18.14, center)
    worksheet298.set_column('D:D', 25, left)
    worksheet298.set_column('E:E', 13.14, left)
    worksheet298.set_column('F:F', 8.57, center)
    worksheet298.set_column('G:R', 5, center)
    worksheet298.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PAHOMAN', title)
    worksheet298.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet298.write('A5', 'LOKASI', header)
    worksheet298.write('B5', 'TOTAL', header)
    worksheet298.merge_range('A4:B4', 'RANK', header)
    worksheet298.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet298.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet298.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet298.merge_range('F4:F5', 'KELAS', header)
    worksheet298.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet298.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet298.write('G5', 'MAT', body)
    worksheet298.write('H5', 'SEJ', body)
    worksheet298.write('I5', 'GEO', body)
    worksheet298.write('J5', 'EKO', body)
    worksheet298.write('K5', 'SOS', body)
    worksheet298.write('L5', 'JML', body)
    worksheet298.write('M5', 'MAT', body)
    worksheet298.write('N5', 'SEJ', body)
    worksheet298.write('O5', 'GEO', body)
    worksheet298.write('P5', 'EKO', body)
    worksheet298.write('Q5', 'SOS', body)
    worksheet298.write('R5', 'JML', body)

    worksheet298.conditional_format(5, 0, row298_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet298.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PAHOMAN', title)
    worksheet298.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet298.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet298.write('A22', 'LOKASI', header)
    worksheet298.write('B22', 'TOTAL', header)
    worksheet298.merge_range('A21:B21', 'RANK', header)
    worksheet298.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet298.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet298.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet298.merge_range('F21:F22', 'KELAS', header)
    worksheet298.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet298.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet298.write('G22', 'MAT', body)
    worksheet298.write('H22', 'SEJ', body)
    worksheet298.write('I22', 'GEO', body)
    worksheet298.write('J22', 'EKO', body)
    worksheet298.write('K22', 'SOS', body)
    worksheet298.write('L22', 'JML', body)
    worksheet298.write('M22', 'MAT', body)
    worksheet298.write('N22', 'SEJ', body)
    worksheet298.write('O22', 'GEO', body)
    worksheet298.write('P22', 'EKO', body)
    worksheet298.write('Q22', 'SOS', body)
    worksheet298.write('R22', 'JML', body)

    worksheet298.conditional_format(22, 0, row298+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 299
    worksheet299.insert_image('A1', r'logo resmi nf.jpg')

    worksheet299.set_column('A:A', 7, center)
    worksheet299.set_column('B:B', 6, center)
    worksheet299.set_column('C:C', 18.14, center)
    worksheet299.set_column('D:D', 25, left)
    worksheet299.set_column('E:E', 13.14, left)
    worksheet299.set_column('F:F', 8.57, center)
    worksheet299.set_column('G:R', 5, center)
    worksheet299.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KEMILING', title)
    worksheet299.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet299.write('A5', 'LOKASI', header)
    worksheet299.write('B5', 'TOTAL', header)
    worksheet299.merge_range('A4:B4', 'RANK', header)
    worksheet299.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet299.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet299.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet299.merge_range('F4:F5', 'KELAS', header)
    worksheet299.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet299.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet299.write('G5', 'MAT', body)
    worksheet299.write('H5', 'SEJ', body)
    worksheet299.write('I5', 'GEO', body)
    worksheet299.write('J5', 'EKO', body)
    worksheet299.write('K5', 'SOS', body)
    worksheet299.write('L5', 'JML', body)
    worksheet299.write('M5', 'MAT', body)
    worksheet299.write('N5', 'SEJ', body)
    worksheet299.write('O5', 'GEO', body)
    worksheet299.write('P5', 'EKO', body)
    worksheet299.write('Q5', 'SOS', body)
    worksheet299.write('R5', 'JML', body)

    worksheet299.conditional_format(5, 0, row299_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet299.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF KEMILING', title)
    worksheet299.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet299.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet299.write('A22', 'LOKASI', header)
    worksheet299.write('B22', 'TOTAL', header)
    worksheet299.merge_range('A21:B21', 'RANK', header)
    worksheet299.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet299.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet299.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet299.merge_range('F21:F22', 'KELAS', header)
    worksheet299.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet299.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet299.write('G22', 'MAT', body)
    worksheet299.write('H22', 'SEJ', body)
    worksheet299.write('I22', 'GEO', body)
    worksheet299.write('J22', 'EKO', body)
    worksheet299.write('K22', 'SOS', body)
    worksheet299.write('L22', 'JML', body)
    worksheet299.write('M22', 'MAT', body)
    worksheet299.write('N22', 'SEJ', body)
    worksheet299.write('O22', 'GEO', body)
    worksheet299.write('P22', 'EKO', body)
    worksheet299.write('Q22', 'SOS', body)
    worksheet299.write('R22', 'JML', body)

    worksheet299.conditional_format(22, 0, row299+21, 17,
                                    {'type': 'no_errors', 'format': border})

    workbook.close()
    st.success("File siap diunduh!")

    # Tombol unduh file
    with open(file_path, "rb") as f:
        bytes_data = f.read()
    st.download_button(label="Unduh File", data=bytes_data,
                       file_name=file_name)


uploaded_file = st.file_uploader(
    'Letakkan file excel NILAI STANDAR [LOKASI DEPOK-PADANG]', type='xlsx')

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # 89
    r = df.shape[0]-5
    # 90
    s = df.shape[0]-4
    # 91
    t = df.shape[0]-3
    # 92
    u = df.shape[0]-2

    # JUMLAH PESERTA
    peserta = df.iloc[r, 23]

    # rata-rata jumlah benar
    rata_mat = df.iloc[r, 156]
    rata_sej = df.iloc[r, 157]
    rata_geo = df.iloc[r, 158]
    rata_eko = df.iloc[r, 159]
    rata_sos = df.iloc[r, 160]
    rata_jml = df.iloc[r, 161]

    # rata-rata nilai standar
    rata_Smat = df.iloc[t, 167]
    rata_Ssej = df.iloc[t, 168]
    rata_Sgeo = df.iloc[t, 169]
    rata_Seko = df.iloc[t, 170]
    rata_Ssos = df.iloc[t, 171]
    rata_Sjml = df.iloc[t, 172]

    max_mat = df.iloc[t, 156]
    max_sej = df.iloc[t, 157]
    max_geo = df.iloc[t, 158]
    max_eko = df.iloc[t, 159]
    max_sos = df.iloc[t, 160]
    max_jml = df.iloc[t, 161]

    # max nilai standar
    max_Smat = df.iloc[r, 167]
    max_Ssej = df.iloc[r, 168]
    max_Sgeo = df.iloc[r, 169]
    max_Seko = df.iloc[r, 170]
    max_Ssos = df.iloc[r, 171]
    max_Sjml = df.iloc[r, 172]

    # min jumlah benar
    min_mat = df.iloc[u, 156]
    min_sej = df.iloc[u, 157]
    min_geo = df.iloc[u, 158]
    min_eko = df.iloc[u, 159]
    min_sos = df.iloc[u, 160]
    min_jml = df.iloc[u, 161]

    # min nilai standar
    min_Smat = df.iloc[s, 167]
    min_Ssej = df.iloc[s, 168]
    min_Sgeo = df.iloc[s, 169]
    min_Seko = df.iloc[s, 170]
    min_Ssos = df.iloc[s, 171]
    min_Sjml = df.iloc[s, 172]

    data_jml_benar = {'BIDANG STUDI': ['MATEMATIKA (MAT)', 'SEJARAH (SEJ)', 'GEOGRAFI (GEO)', 'EKONOMI (EKO)', 'SOSIOLOGI (SOS)', 'JUMLAH (JML)'],
                      'TERENDAH': [min_mat, min_sej, min_geo, min_eko, min_sos, min_jml],
                      'RATA-RATA': [rata_mat, rata_sej, rata_geo, rata_eko, rata_sos, rata_jml],
                      'TERTINGGI': [max_mat, max_sej, max_geo, max_eko, max_sos, max_jml]}

    jml_benar = pd.DataFrame(data_jml_benar)

    data_n_standar = {'BIDANG STUDI': ['MATEMATIKA (MAT)', 'SEJARAH (SEJ)', 'GEOGRAFI (GEO)', 'EKONOMI (EKO)', 'SOSIOLOGI (SOS)', 'JUMLAH (JML)'],
                      'TERENDAH': [min_Smat, min_Ssej, min_Sgeo, min_Seko, min_Ssos, min_Sjml],
                      'RATA-RATA': [rata_Smat, rata_Ssej, rata_Sgeo, rata_Seko, rata_Ssos, rata_Sjml],
                      'TERTINGGI': [max_Smat, max_Ssej, max_Sgeo, max_Seko, max_Ssos, max_Sjml]}

    n_standar = pd.DataFrame(data_n_standar)

    data_jml_peserta = {'JUMLAH PESERTA': [peserta]}

    jml_peserta = pd.DataFrame(data_jml_peserta)

    data_jml_soal = {'BIDANG STUDI': ['MAT', 'SEJ', 'GEO', 'EKO', 'SOS'],
                     'JUMLAH': [JML_SOAL_MAT, JML_SOAL_SEJ, JML_SOAL_GEO, JML_SOAL_EKO, JML_SOAL_SOS]}

    jml_soal = pd.DataFrame(data_jml_soal)

    df = df[['LOKASI', 'RANK LOK.', 'RANK NAS.', 'NOMOR NF', 'NAMA SISWA', 'NAMA SEKOLAH', 'KELAS',
             'MAT', 'SEJ', 'GEO', 'EKO', 'SOS', 'JML', 'S_MAT', 'S_SEJ', 'S_GEO', 'S_EKO', 'S_SOS', 'S_JML']]

    # sort setiap lokasi
    sort530 = df[df['LOKASI'] == 530]
    sort531 = df[df['LOKASI'] == 531]
    sort532 = df[df['LOKASI'] == 532]
    sort533 = df[df['LOKASI'] == 533]
    sort534 = df[df['LOKASI'] == 534]
    sort535 = df[df['LOKASI'] == 535]
    sort546 = df[df['LOKASI'] == 546]
    sort547 = df[df['LOKASI'] == 547]
    sort548 = df[df['LOKASI'] == 548]
    sort549 = df[df['LOKASI'] == 549]
    sort556 = df[df['LOKASI'] == 556]
    sort557 = df[df['LOKASI'] == 557]
    sort558 = df[df['LOKASI'] == 558]
    sort575 = df[df['LOKASI'] == 575]
    sort576 = df[df['LOKASI'] == 576]
    sort577 = df[df['LOKASI'] == 577]
    sort578 = df[df['LOKASI'] == 578]
    sort588 = df[df['LOKASI'] == 588]
    sort589 = df[df['LOKASI'] == 589]
    sort594 = df[df['LOKASI'] == 594]
    sort661 = df[df['LOKASI'] == 661]
    sort662 = df[df['LOKASI'] == 662]
    sort663 = df[df['LOKASI'] == 663]
    sort664 = df[df['LOKASI'] == 664]

    # 10 besar setiap lokasi
    # 530
    sort530_10 = sort530.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort530_10['LOKASI']
    sort530_10 = sort530_10.drop(
        sort530_10[(sort530_10['RANK LOK.'] > 10)].index)
    # 531
    sort531_10 = sort531.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort531_10['LOKASI']
    sort531_10 = sort531_10.drop(
        sort531_10[(sort531_10['RANK LOK.'] > 10)].index)
    # 532
    sort532_10 = sort532.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort532_10['LOKASI']
    sort532_10 = sort532_10.drop(
        sort532_10[(sort532_10['RANK LOK.'] > 10)].index)
    # 533
    sort533_10 = sort533.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort533_10['LOKASI']
    sort533_10 = sort533_10.drop(
        sort533_10[(sort533_10['RANK LOK.'] > 10)].index)
    # 534
    sort534_10 = sort534.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort534_10['LOKASI']
    sort534_10 = sort534_10.drop(
        sort534_10[(sort534_10['RANK LOK.'] > 10)].index)
    # 535
    sort535_10 = sort535.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort535_10['LOKASI']
    sort535_10 = sort535_10.drop(
        sort535_10[(sort535_10['RANK LOK.'] > 10)].index)
    # 546
    sort546_10 = sort546.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort546_10['LOKASI']
    sort546_10 = sort546_10.drop(
        sort546_10[(sort546_10['RANK LOK.'] > 10)].index)
    # 547
    sort547_10 = sort547.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort547_10['LOKASI']
    sort547_10 = sort547_10.drop(
        sort547_10[(sort547_10['RANK LOK.'] > 10)].index)
    # 548
    sort548_10 = sort548.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort548_10['LOKASI']
    sort548_10 = sort548_10.drop(
        sort548_10[(sort548_10['RANK LOK.'] > 10)].index)
    # 549
    sort549_10 = sort549.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort549_10['LOKASI']
    sort549_10 = sort549_10.drop(
        sort549_10[(sort549_10['RANK LOK.'] > 10)].index)
    # 556
    sort556_10 = sort556.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort556_10['LOKASI']
    sort556_10 = sort556_10.drop(
        sort556_10[(sort556_10['RANK LOK.'] > 10)].index)
    # 557
    sort557_10 = sort557.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort557_10['LOKASI']
    sort557_10 = sort557_10.drop(
        sort557_10[(sort557_10['RANK LOK.'] > 10)].index)
    # 558
    sort558_10 = sort558.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort558_10['LOKASI']
    sort558_10 = sort558_10.drop(
        sort558_10[(sort558_10['RANK LOK.'] > 10)].index)
    # 575
    sort575_10 = sort575.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort575_10['LOKASI']
    sort575_10 = sort575_10.drop(
        sort575_10[(sort575_10['RANK LOK.'] > 10)].index)
    # 576
    sort576_10 = sort576.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort576_10['LOKASI']
    sort576_10 = sort576_10.drop(
        sort576_10[(sort576_10['RANK LOK.'] > 10)].index)
    # 577
    sort577_10 = sort577.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort577_10['LOKASI']
    sort577_10 = sort577_10.drop(
        sort577_10[(sort577_10['RANK LOK.'] > 10)].index)
    # 578
    sort578_10 = sort578.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort578_10['LOKASI']
    sort578_10 = sort578_10.drop(
        sort578_10[(sort578_10['RANK LOK.'] > 10)].index)
    # 588
    sort588_10 = sort588.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort588_10['LOKASI']
    sort588_10 = sort588_10.drop(
        sort588_10[(sort588_10['RANK LOK.'] > 10)].index)
    # 589
    sort589_10 = sort589.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort589_10['LOKASI']
    sort589_10 = sort589_10.drop(
        sort589_10[(sort589_10['RANK LOK.'] > 10)].index)
    # 594
    sort594_10 = sort594.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort594_10['LOKASI']
    sort594_10 = sort594_10.drop(
        sort594_10[(sort594_10['RANK LOK.'] > 10)].index)
    # 661
    sort661_10 = sort661.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort661_10['LOKASI']
    sort661_10 = sort661_10.drop(
        sort661_10[(sort661_10['RANK LOK.'] > 10)].index)
    # 662
    sort662_10 = sort662.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort662_10['LOKASI']
    sort662_10 = sort662_10.drop(
        sort662_10[(sort662_10['RANK LOK.'] > 10)].index)
    # 663
    sort663_10 = sort663.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort663_10['LOKASI']
    sort663_10 = sort663_10.drop(
        sort663_10[(sort663_10['RANK LOK.'] > 10)].index)
    # 664
    sort664_10 = sort664.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort664_10['LOKASI']
    sort664_10 = sort664_10.drop(
        sort664_10[(sort664_10['RANK LOK.'] > 10)].index)

    # All 530
    sort530 = sort530.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort530['LOKASI']
    # All 531
    sort531 = sort531.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort531['LOKASI']
    # All 532
    sort532 = sort532.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort532['LOKASI']
    # All 533
    sort533 = sort533.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort533['LOKASI']
    # All 534
    sort534 = sort534.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort534['LOKASI']
    # All 535
    sort535 = sort535.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort535['LOKASI']
    # All 546
    sort546 = sort546.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort546['LOKASI']
    # All 547
    sort547 = sort547.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort547['LOKASI']
    # All 548
    sort548 = sort548.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort548['LOKASI']
    # All 549
    sort549 = sort549.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort549['LOKASI']
    # All 556
    sort556 = sort556.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort556['LOKASI']
    # All 557
    sort557 = sort557.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort557['LOKASI']
    # All 558
    sort558 = sort558.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort558['LOKASI']
    # All 575
    sort575 = sort575.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort575['LOKASI']
    # All 576
    sort576 = sort576.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort576['LOKASI']
    # All 577
    sort577 = sort577.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort577['LOKASI']
    # All 578
    sort578 = sort578.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort578['LOKASI']
    # All 588
    sort588 = sort588.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort588['LOKASI']
    # All 589
    sort589 = sort589.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort589['LOKASI']
    # All 594
    sort594 = sort594.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort594['LOKASI']
    # All 661
    sort661 = sort661.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort661['LOKASI']
    # All 662
    sort662 = sort662.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort662['LOKASI']
    # All 663
    sort663 = sort663.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort663['LOKASI']
    # All 664
    sort664 = sort664.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort664['LOKASI']

    # jumlah row
    # 530
    row530_10 = sort530_10.shape[0]
    row530 = sort530.shape[0]
    # 531
    row531_10 = sort531_10.shape[0]
    row531 = sort531.shape[0]
    # 532
    row532_10 = sort532_10.shape[0]
    row532 = sort532.shape[0]
    # 533
    row533_10 = sort533_10.shape[0]
    row533 = sort533.shape[0]
    # 534
    row534_10 = sort534_10.shape[0]
    row534 = sort534.shape[0]
    # 535
    row535_10 = sort535_10.shape[0]
    row535 = sort535.shape[0]
    # 546
    row546_10 = sort546_10.shape[0]
    row546 = sort546.shape[0]
    # 547
    row547_10 = sort547_10.shape[0]
    row547 = sort547.shape[0]
    # 548
    row548_10 = sort548_10.shape[0]
    row548 = sort548.shape[0]
    # 549
    row549_10 = sort549_10.shape[0]
    row549 = sort549.shape[0]
    # 556
    row556_10 = sort556_10.shape[0]
    row556 = sort556.shape[0]
    # 557
    row557_10 = sort557_10.shape[0]
    row557 = sort557.shape[0]
    # 558
    row558_10 = sort558_10.shape[0]
    row558 = sort558.shape[0]
    # 575
    row575_10 = sort575_10.shape[0]
    row575 = sort575.shape[0]
    # 576
    row576_10 = sort576_10.shape[0]
    row576 = sort576.shape[0]
    # 577
    row577_10 = sort577_10.shape[0]
    row577 = sort577.shape[0]
    # 578
    row578_10 = sort578_10.shape[0]
    row578 = sort578.shape[0]
    # 588
    row588_10 = sort588_10.shape[0]
    row588 = sort588.shape[0]
    # 589
    row589_10 = sort589_10.shape[0]
    row589 = sort589.shape[0]
    # 594
    row594_10 = sort594_10.shape[0]
    row594 = sort594.shape[0]
    # 661
    row661_10 = sort661_10.shape[0]
    row661 = sort661.shape[0]
    # 662
    row662_10 = sort662_10.shape[0]
    row662 = sort662.shape[0]
    # 663
    row663_10 = sort663_10.shape[0]
    row663 = sort663.shape[0]
    # 664
    row664_10 = sort664_10.shape[0]
    row664 = sort664.shape[0]

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    # Path file hasil penyimpanan
    file_name = f"{kelas}_{penilaian}_{semester}_lokasi_depok_padang.xlsx"
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
                       startrow=21,
                       startcol=0,
                       index=False,
                       header=False)

    # Convert the dataframe to an XlsxWriter Excel object cover.
    jml_peserta.to_excel(writer, sheet_name='cover',
                         startrow=21,
                         startcol=5,
                         index=False,
                         header=False)

    # Convert the dataframe to an XlsxWriter Excel object cover.
    jml_soal.to_excel(writer, sheet_name='cover',
                      startrow=13,
                      startcol=5,
                      index=False,
                      header=False)

    # 530
    # Convert the dataframe to an XlsxWriter Excel object.
    sort530_10.to_excel(writer, sheet_name='530',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort530.to_excel(writer, sheet_name='530',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 531
    # Convert the dataframe to an XlsxWriter Excel object.
    sort531_10.to_excel(writer, sheet_name='531',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort531.to_excel(writer, sheet_name='531',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 532
    # Convert the dataframe to an XlsxWriter Excel object.
    sort532_10.to_excel(writer, sheet_name='532',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort532.to_excel(writer, sheet_name='532',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 533
    # Convert the dataframe to an XlsxWriter Excel object.
    sort533_10.to_excel(writer, sheet_name='533',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort533.to_excel(writer, sheet_name='533',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 534
    # Convert the dataframe to an XlsxWriter Excel object.
    sort534_10.to_excel(writer, sheet_name='534',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort534.to_excel(writer, sheet_name='534',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 535
    # Convert the dataframe to an XlsxWriter Excel object.
    sort535_10.to_excel(writer, sheet_name='535',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort535.to_excel(writer, sheet_name='535',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 546
    # Convert the dataframe to an XlsxWriter Excel object.
    sort546_10.to_excel(writer, sheet_name='546',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort546.to_excel(writer, sheet_name='546',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 547
    # Convert the dataframe to an XlsxWriter Excel object.
    sort547_10.to_excel(writer, sheet_name='547',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort547.to_excel(writer, sheet_name='547',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 548
    # Convert the dataframe to an XlsxWriter Excel object.
    sort548_10.to_excel(writer, sheet_name='548',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort548.to_excel(writer, sheet_name='548',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 549
    # Convert the dataframe to an XlsxWriter Excel object.
    sort549_10.to_excel(writer, sheet_name='549',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort549.to_excel(writer, sheet_name='549',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 556
    # Convert the dataframe to an XlsxWriter Excel object.
    sort556_10.to_excel(writer, sheet_name='556',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort556.to_excel(writer, sheet_name='556',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 557
    # Convert the dataframe to an XlsxWriter Excel object.
    sort557_10.to_excel(writer, sheet_name='557',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort557.to_excel(writer, sheet_name='557',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 558
    # Convert the dataframe to an XlsxWriter Excel object.
    sort558_10.to_excel(writer, sheet_name='558',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort558.to_excel(writer, sheet_name='558',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 575
    # Convert the dataframe to an XlsxWriter Excel object.
    sort575_10.to_excel(writer, sheet_name='575',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort575.to_excel(writer, sheet_name='575',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 576
    # Convert the dataframe to an XlsxWriter Excel object.
    sort576_10.to_excel(writer, sheet_name='576',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort576.to_excel(writer, sheet_name='576',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 577
    # Convert the dataframe to an XlsxWriter Excel object.
    sort577_10.to_excel(writer, sheet_name='577',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort577.to_excel(writer, sheet_name='577',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 578
    # Convert the dataframe to an XlsxWriter Excel object.
    sort578_10.to_excel(writer, sheet_name='578',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort578.to_excel(writer, sheet_name='578',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 588
    # Convert the dataframe to an XlsxWriter Excel object.
    sort588_10.to_excel(writer, sheet_name='588',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort588.to_excel(writer, sheet_name='588',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 589
    # Convert the dataframe to an XlsxWriter Excel object.
    sort589_10.to_excel(writer, sheet_name='589',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort589.to_excel(writer, sheet_name='589',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 594
    # Convert the dataframe to an XlsxWriter Excel object.
    sort594_10.to_excel(writer, sheet_name='594',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort594.to_excel(writer, sheet_name='594',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 661
    # Convert the dataframe to an XlsxWriter Excel object.
    sort661_10.to_excel(writer, sheet_name='661',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort661.to_excel(writer, sheet_name='661',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 662
    # Convert the dataframe to an XlsxWriter Excel object.
    sort662_10.to_excel(writer, sheet_name='662',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort662.to_excel(writer, sheet_name='662',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 663
    # Convert the dataframe to an XlsxWriter Excel object.
    sort663_10.to_excel(writer, sheet_name='663',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort663.to_excel(writer, sheet_name='663',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)
    # 664
    # Convert the dataframe to an XlsxWriter Excel object.
    sort664_10.to_excel(writer, sheet_name='664',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort664.to_excel(writer, sheet_name='664',
                     startrow=22,
                     startcol=0,
                     index=False,
                     header=False)

    # Get the xlsxwriter objects from the dataframe writer object.
    workbook = writer.book

    # membuat worksheet baru
    worksheetcover = writer.sheets['cover']
    worksheet530 = writer.sheets['530']
    worksheet531 = writer.sheets['531']
    worksheet532 = writer.sheets['532']
    worksheet533 = writer.sheets['533']
    worksheet534 = writer.sheets['534']
    worksheet535 = writer.sheets['535']
    worksheet546 = writer.sheets['546']
    worksheet547 = writer.sheets['547']
    worksheet548 = writer.sheets['548']
    worksheet549 = writer.sheets['549']
    worksheet556 = writer.sheets['556']
    worksheet557 = writer.sheets['557']
    worksheet558 = writer.sheets['558']
    worksheet575 = writer.sheets['575']
    worksheet576 = writer.sheets['576']
    worksheet577 = writer.sheets['577']
    worksheet578 = writer.sheets['578']
    worksheet588 = writer.sheets['588']
    worksheet589 = writer.sheets['589']
    worksheet594 = writer.sheets['594']
    worksheet661 = writer.sheets['661']
    worksheet662 = writer.sheets['662']
    worksheet663 = writer.sheets['663']
    worksheet664 = writer.sheets['664']

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
    worksheetcover.conditional_format(16, 0, 11, 3,
                                      {'type': 'no_errors', 'format': borderCover})

    worksheetcover.insert_image('F1', r'logo nf.jpg')

    worksheetcover.merge_range('A10:A11', 'BIDANG STUDI', bodyCover)
    worksheetcover.merge_range('B10:B11', 'TERENDAH', bodyCover)
    worksheetcover.merge_range('C10:C11', 'RATA-RATA', bodyCover)
    worksheetcover.merge_range('D10:D11', 'TERTINGGI', bodyCover)
    worksheetcover.merge_range('A20:A21', 'BIDANG STUDI', bodyCover)
    worksheetcover.merge_range('B20:B21', 'TERENDAH', bodyCover)
    worksheetcover.merge_range('C20:C21', 'RATA-RATA', bodyCover)
    worksheetcover.merge_range('D20:D21', 'TERTINGGI', bodyCover)
    worksheetcover.write('F13', 'BIDANG STUDI', bodyCover)
    worksheetcover.merge_range('F20:F21', 'JUMLAH', sub_header1Cover)
    worksheetcover.merge_range('F23:F24', 'PESERTA', sub_header1Cover)
    worksheetcover.write('G13', 'JUMLAH', bodyCover)
    worksheetcover.set_column('A:A', 25.71, centerCover)
    worksheetcover.set_column('B:B', 15, centerCover)
    worksheetcover.set_column('C:C', 15, centerCover)
    worksheetcover.set_column('D:D', 15, centerCover)
    worksheetcover.set_column('F:F', 25.71, centerCover)
    worksheetcover.set_column('G:G', 13, centerCover)
    worksheetcover.merge_range('A1:F3', 'DAFTAR NILAI', titleCover)
    worksheetcover.merge_range(
        'A4:F5', fr'{penilaian}', sub_titleCover)
    worksheetcover.merge_range(
        'A6:F7', fr'{semester} TAHUN {tahun}', headerCover)
    worksheetcover.write('A9', 'JUMLAH BENAR', sub_headerCover)
    worksheetcover.write('A19', 'NILAI STANDAR', sub_headerCover)
    worksheetcover.merge_range('F8:G9', fr'{kelas}', kelasCover)
    worksheetcover.merge_range('F11:G12', 'JUMLAH SOAL', sub_header1Cover)

    worksheetcover.conditional_format(26, 0, 21, 3,
                                      {'type': 'no_errors', 'format': borderCover})

    worksheetcover.conditional_format(17, 6, 13, 5,
                                      {'type': 'no_errors', 'format': borderCover})

    worksheetcover.conditional_format(21, 5, 21, 5,
                                      {'type': 'no_errors', 'format': borderCover})

    # worksheet 530
    worksheet530.insert_image('A1', r'logo resmi nf.jpg')

    worksheet530.set_column('A:A', 7, center)
    worksheet530.set_column('B:B', 6, center)
    worksheet530.set_column('C:C', 18.14, center)
    worksheet530.set_column('D:D', 25, left)
    worksheet530.set_column('E:E', 13.14, left)
    worksheet530.set_column('F:F', 8.57, center)
    worksheet530.set_column('G:R', 5, center)
    worksheet530.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF POLSEK DEPOK', title)
    worksheet530.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet530.write('A5', 'LOKASI', header)
    worksheet530.write('B5', 'TOTAL', header)
    worksheet530.merge_range('A4:B4', 'RANK', header)
    worksheet530.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet530.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet530.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet530.merge_range('F4:F5', 'KELAS', header)
    worksheet530.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet530.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet530.write('G5', 'MAT', body)
    worksheet530.write('H5', 'SEJ', body)
    worksheet530.write('I5', 'GEO', body)
    worksheet530.write('J5', 'EKO', body)
    worksheet530.write('K5', 'SOS', body)
    worksheet530.write('L5', 'JML', body)
    worksheet530.write('M5', 'MAT', body)
    worksheet530.write('N5', 'SEJ', body)
    worksheet530.write('O5', 'GEO', body)
    worksheet530.write('P5', 'EKO', body)
    worksheet530.write('Q5', 'SOS', body)
    worksheet530.write('R5', 'JML', body)

    worksheet530.conditional_format(5, 0, row530_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet530.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF POLSEK DEPOK', title)
    worksheet530.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet530.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet530.write('A22', 'LOKASI', header)
    worksheet530.write('B22', 'TOTAL', header)
    worksheet530.merge_range('A21:B21', 'RANK', header)
    worksheet530.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet530.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet530.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet530.merge_range('F21:F22', 'KELAS', header)
    worksheet530.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet530.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet530.write('G22', 'MAT', body)
    worksheet530.write('H22', 'SEJ', body)
    worksheet530.write('I22', 'GEO', body)
    worksheet530.write('J22', 'EKO', body)
    worksheet530.write('K22', 'SOS', body)
    worksheet530.write('L22', 'JML', body)
    worksheet530.write('M22', 'MAT', body)
    worksheet530.write('N22', 'SEJ', body)
    worksheet530.write('O22', 'GEO', body)
    worksheet530.write('P22', 'EKO', body)
    worksheet530.write('Q22', 'SOS', body)
    worksheet530.write('R22', 'JML', body)

    worksheet530.conditional_format(22, 0, row530+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 531
    worksheet531.insert_image('A1', r'logo resmi nf.jpg')

    worksheet531.set_column('A:A', 7, center)
    worksheet531.set_column('B:B', 6, center)
    worksheet531.set_column('C:C', 18.14, center)
    worksheet531.set_column('D:D', 25, left)
    worksheet531.set_column('E:E', 13.14, left)
    worksheet531.set_column('F:F', 8.57, center)
    worksheet531.set_column('G:R', 5, center)
    worksheet531.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF DEPOK 1', title)
    worksheet531.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet531.write('A5', 'LOKASI', header)
    worksheet531.write('B5', 'TOTAL', header)
    worksheet531.merge_range('A4:B4', 'RANK', header)
    worksheet531.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet531.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet531.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet531.merge_range('F4:F5', 'KELAS', header)
    worksheet531.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet531.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet531.write('G5', 'MAT', body)
    worksheet531.write('H5', 'SEJ', body)
    worksheet531.write('I5', 'GEO', body)
    worksheet531.write('J5', 'EKO', body)
    worksheet531.write('K5', 'SOS', body)
    worksheet531.write('L5', 'JML', body)
    worksheet531.write('M5', 'MAT', body)
    worksheet531.write('N5', 'SEJ', body)
    worksheet531.write('O5', 'GEO', body)
    worksheet531.write('P5', 'EKO', body)
    worksheet531.write('Q5', 'SOS', body)
    worksheet531.write('R5', 'JML', body)

    worksheet531.conditional_format(5, 0, row531_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet531.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF DEPOK 1', title)
    worksheet531.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet531.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet531.write('A22', 'LOKASI', header)
    worksheet531.write('B22', 'TOTAL', header)
    worksheet531.merge_range('A21:B21', 'RANK', header)
    worksheet531.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet531.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet531.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet531.merge_range('F21:F22', 'KELAS', header)
    worksheet531.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet531.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet531.write('G22', 'MAT', body)
    worksheet531.write('H22', 'SEJ', body)
    worksheet531.write('I22', 'GEO', body)
    worksheet531.write('J22', 'EKO', body)
    worksheet531.write('K22', 'SOS', body)
    worksheet531.write('L22', 'JML', body)
    worksheet531.write('M22', 'MAT', body)
    worksheet531.write('N22', 'SEJ', body)
    worksheet531.write('O22', 'GEO', body)
    worksheet531.write('P22', 'EKO', body)
    worksheet531.write('Q22', 'SOS', body)
    worksheet531.write('R22', 'JML', body)

    worksheet531.conditional_format(22, 0, row531+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 532
    worksheet532.insert_image('A1', r'logo resmi nf.jpg')

    worksheet532.set_column('A:A', 7, center)
    worksheet532.set_column('B:B', 6, center)
    worksheet532.set_column('C:C', 18.14, center)
    worksheet532.set_column('D:D', 25, left)
    worksheet532.set_column('E:E', 13.14, left)
    worksheet532.set_column('F:F', 8.57, center)
    worksheet532.set_column('G:R', 5, center)
    worksheet532.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PROKLAMASI DEPOK 2', title)
    worksheet532.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet532.write('A5', 'LOKASI', header)
    worksheet532.write('B5', 'TOTAL', header)
    worksheet532.merge_range('A4:B4', 'RANK', header)
    worksheet532.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet532.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet532.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet532.merge_range('F4:F5', 'KELAS', header)
    worksheet532.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet532.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet532.write('G5', 'MAT', body)
    worksheet532.write('H5', 'SEJ', body)
    worksheet532.write('I5', 'GEO', body)
    worksheet532.write('J5', 'EKO', body)
    worksheet532.write('K5', 'SOS', body)
    worksheet532.write('L5', 'JML', body)
    worksheet532.write('M5', 'MAT', body)
    worksheet532.write('N5', 'SEJ', body)
    worksheet532.write('O5', 'GEO', body)
    worksheet532.write('P5', 'EKO', body)
    worksheet532.write('Q5', 'SOS', body)
    worksheet532.write('R5', 'JML', body)

    worksheet532.conditional_format(5, 0, row532_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet532.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PROKLAMASI DEPOK 2', title)
    worksheet532.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet532.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet532.write('A22', 'LOKASI', header)
    worksheet532.write('B22', 'TOTAL', header)
    worksheet532.merge_range('A21:B21', 'RANK', header)
    worksheet532.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet532.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet532.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet532.merge_range('F21:F22', 'KELAS', header)
    worksheet532.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet532.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet532.write('G22', 'MAT', body)
    worksheet532.write('H22', 'SEJ', body)
    worksheet532.write('I22', 'GEO', body)
    worksheet532.write('J22', 'EKO', body)
    worksheet532.write('K22', 'SOS', body)
    worksheet532.write('L22', 'JML', body)
    worksheet532.write('M22', 'MAT', body)
    worksheet532.write('N22', 'SEJ', body)
    worksheet532.write('O22', 'GEO', body)
    worksheet532.write('P22', 'EKO', body)
    worksheet532.write('Q22', 'SOS', body)
    worksheet532.write('R22', 'JML', body)

    worksheet532.conditional_format(22, 0, row532+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 533
    worksheet533.insert_image('A1', r'logo resmi nf.jpg')

    worksheet533.set_column('A:A', 7, center)
    worksheet533.set_column('B:B', 6, center)
    worksheet533.set_column('C:C', 18.14, center)
    worksheet533.set_column('D:D', 25, left)
    worksheet533.set_column('E:E', 13.14, left)
    worksheet533.set_column('F:F', 8.57, center)
    worksheet533.set_column('G:R', 5, center)
    worksheet533.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIMANGGIS', title)
    worksheet533.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet533.write('A5', 'LOKASI', header)
    worksheet533.write('B5', 'TOTAL', header)
    worksheet533.merge_range('A4:B4', 'RANK', header)
    worksheet533.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet533.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet533.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet533.merge_range('F4:F5', 'KELAS', header)
    worksheet533.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet533.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet533.write('G5', 'MAT', body)
    worksheet533.write('H5', 'SEJ', body)
    worksheet533.write('I5', 'GEO', body)
    worksheet533.write('J5', 'EKO', body)
    worksheet533.write('K5', 'SOS', body)
    worksheet533.write('L5', 'JML', body)
    worksheet533.write('M5', 'MAT', body)
    worksheet533.write('N5', 'SEJ', body)
    worksheet533.write('O5', 'GEO', body)
    worksheet533.write('P5', 'EKO', body)
    worksheet533.write('Q5', 'SOS', body)
    worksheet533.write('R5', 'JML', body)

    worksheet533.conditional_format(5, 0, row533_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet533.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CIMANGGIS', title)
    worksheet533.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet533.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet533.write('A22', 'LOKASI', header)
    worksheet533.write('B22', 'TOTAL', header)
    worksheet533.merge_range('A21:B21', 'RANK', header)
    worksheet533.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet533.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet533.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet533.merge_range('F21:F22', 'KELAS', header)
    worksheet533.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet533.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet533.write('G22', 'MAT', body)
    worksheet533.write('H22', 'SEJ', body)
    worksheet533.write('I22', 'GEO', body)
    worksheet533.write('J22', 'EKO', body)
    worksheet533.write('K22', 'SOS', body)
    worksheet533.write('L22', 'JML', body)
    worksheet533.write('M22', 'MAT', body)
    worksheet533.write('N22', 'SEJ', body)
    worksheet533.write('O22', 'GEO', body)
    worksheet533.write('P22', 'EKO', body)
    worksheet533.write('Q22', 'SOS', body)
    worksheet533.write('R22', 'JML', body)

    worksheet533.conditional_format(22, 0, row533+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 534
    worksheet534.insert_image('A1', r'logo resmi nf.jpg')

    worksheet534.set_column('A:A', 7, center)
    worksheet534.set_column('B:B', 6, center)
    worksheet534.set_column('C:C', 18.14, center)
    worksheet534.set_column('D:D', 25, left)
    worksheet534.set_column('E:E', 13.14, left)
    worksheet534.set_column('F:F', 8.57, center)
    worksheet534.set_column('G:R', 5, center)
    worksheet534.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SAWANGAN', title)
    worksheet534.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet534.write('A5', 'LOKASI', header)
    worksheet534.write('B5', 'TOTAL', header)
    worksheet534.merge_range('A4:B4', 'RANK', header)
    worksheet534.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet534.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet534.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet534.merge_range('F4:F5', 'KELAS', header)
    worksheet534.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet534.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet534.write('G5', 'MAT', body)
    worksheet534.write('H5', 'SEJ', body)
    worksheet534.write('I5', 'GEO', body)
    worksheet534.write('J5', 'EKO', body)
    worksheet534.write('K5', 'SOS', body)
    worksheet534.write('L5', 'JML', body)
    worksheet534.write('M5', 'MAT', body)
    worksheet534.write('N5', 'SEJ', body)
    worksheet534.write('O5', 'GEO', body)
    worksheet534.write('P5', 'EKO', body)
    worksheet534.write('Q5', 'SOS', body)
    worksheet534.write('R5', 'JML', body)

    worksheet534.conditional_format(5, 0, row534_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet534.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF SAWANGAN', title)
    worksheet534.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet534.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet534.write('A22', 'LOKASI', header)
    worksheet534.write('B22', 'TOTAL', header)
    worksheet534.merge_range('A21:B21', 'RANK', header)
    worksheet534.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet534.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet534.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet534.merge_range('F21:F22', 'KELAS', header)
    worksheet534.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet534.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet534.write('G22', 'MAT', body)
    worksheet534.write('H22', 'SEJ', body)
    worksheet534.write('I22', 'GEO', body)
    worksheet534.write('J22', 'EKO', body)
    worksheet534.write('K22', 'SOS', body)
    worksheet534.write('L22', 'JML', body)
    worksheet534.write('M22', 'MAT', body)
    worksheet534.write('N22', 'SEJ', body)
    worksheet534.write('O22', 'GEO', body)
    worksheet534.write('P22', 'EKO', body)
    worksheet534.write('Q22', 'SOS', body)
    worksheet534.write('R22', 'JML', body)

    worksheet534.conditional_format(22, 0, row534+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 535
    worksheet535.insert_image('A1', r'logo resmi nf.jpg')

    worksheet535.set_column('A:A', 7, center)
    worksheet535.set_column('B:B', 6, center)
    worksheet535.set_column('C:C', 18.14, center)
    worksheet535.set_column('D:D', 25, left)
    worksheet535.set_column('E:E', 13.14, left)
    worksheet535.set_column('F:F', 8.57, center)
    worksheet535.set_column('G:R', 5, center)
    worksheet535.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BRIMOB', title)
    worksheet535.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet535.write('A5', 'LOKASI', header)
    worksheet535.write('B5', 'TOTAL', header)
    worksheet535.merge_range('A4:B4', 'RANK', header)
    worksheet535.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet535.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet535.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet535.merge_range('F4:F5', 'KELAS', header)
    worksheet535.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet535.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet535.write('G5', 'MAT', body)
    worksheet535.write('H5', 'SEJ', body)
    worksheet535.write('I5', 'GEO', body)
    worksheet535.write('J5', 'EKO', body)
    worksheet535.write('K5', 'SOS', body)
    worksheet535.write('L5', 'JML', body)
    worksheet535.write('M5', 'MAT', body)
    worksheet535.write('N5', 'SEJ', body)
    worksheet535.write('O5', 'GEO', body)
    worksheet535.write('P5', 'EKO', body)
    worksheet535.write('Q5', 'SOS', body)
    worksheet535.write('R5', 'JML', body)

    worksheet535.conditional_format(5, 0, row535_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet535.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF BRIMOB', title)
    worksheet535.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet535.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet535.write('A22', 'LOKASI', header)
    worksheet535.write('B22', 'TOTAL', header)
    worksheet535.merge_range('A21:B21', 'RANK', header)
    worksheet535.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet535.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet535.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet535.merge_range('F21:F22', 'KELAS', header)
    worksheet535.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet535.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet535.write('G22', 'MAT', body)
    worksheet535.write('H22', 'SEJ', body)
    worksheet535.write('I22', 'GEO', body)
    worksheet535.write('J22', 'EKO', body)
    worksheet535.write('K22', 'SOS', body)
    worksheet535.write('L22', 'JML', body)
    worksheet535.write('M22', 'MAT', body)
    worksheet535.write('N22', 'SEJ', body)
    worksheet535.write('O22', 'GEO', body)
    worksheet535.write('P22', 'EKO', body)
    worksheet535.write('Q22', 'SOS', body)
    worksheet535.write('R22', 'JML', body)

    worksheet535.conditional_format(22, 0, row535+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 546
    worksheet546.insert_image('A1', r'logo resmi nf.jpg')

    worksheet546.set_column('A:A', 7, center)
    worksheet546.set_column('B:B', 6, center)
    worksheet546.set_column('C:C', 18.14, center)
    worksheet546.set_column('D:D', 25, left)
    worksheet546.set_column('E:E', 13.14, left)
    worksheet546.set_column('F:F', 8.57, center)
    worksheet546.set_column('G:R', 5, center)
    worksheet546.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PAJAJARAN (PPIB)', title)
    worksheet546.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet546.write('A5', 'LOKASI', header)
    worksheet546.write('B5', 'TOTAL', header)
    worksheet546.merge_range('A4:B4', 'RANK', header)
    worksheet546.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet546.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet546.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet546.merge_range('F4:F5', 'KELAS', header)
    worksheet546.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet546.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet546.write('G5', 'MAT', body)
    worksheet546.write('H5', 'SEJ', body)
    worksheet546.write('I5', 'GEO', body)
    worksheet546.write('J5', 'EKO', body)
    worksheet546.write('K5', 'SOS', body)
    worksheet546.write('L5', 'JML', body)
    worksheet546.write('M5', 'MAT', body)
    worksheet546.write('N5', 'SEJ', body)
    worksheet546.write('O5', 'GEO', body)
    worksheet546.write('P5', 'EKO', body)
    worksheet546.write('Q5', 'SOS', body)
    worksheet546.write('R5', 'JML', body)

    worksheet546.conditional_format(5, 0, row546_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet546.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PAJAJARAN (PPIB)', title)
    worksheet546.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet546.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet546.write('A22', 'LOKASI', header)
    worksheet546.write('B22', 'TOTAL', header)
    worksheet546.merge_range('A21:B21', 'RANK', header)
    worksheet546.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet546.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet546.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet546.merge_range('F21:F22', 'KELAS', header)
    worksheet546.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet546.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet546.write('G22', 'MAT', body)
    worksheet546.write('H22', 'SEJ', body)
    worksheet546.write('I22', 'GEO', body)
    worksheet546.write('J22', 'EKO', body)
    worksheet546.write('K22', 'SOS', body)
    worksheet546.write('L22', 'JML', body)
    worksheet546.write('M22', 'MAT', body)
    worksheet546.write('N22', 'SEJ', body)
    worksheet546.write('O22', 'GEO', body)
    worksheet546.write('P22', 'EKO', body)
    worksheet546.write('Q22', 'SOS', body)
    worksheet546.write('R22', 'JML', body)

    worksheet546.conditional_format(22, 0, row546+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 547
    worksheet547.insert_image('A1', r'logo resmi nf.jpg')

    worksheet547.set_column('A:A', 7, center)
    worksheet547.set_column('B:B', 6, center)
    worksheet547.set_column('C:C', 18.14, center)
    worksheet547.set_column('D:D', 25, left)
    worksheet547.set_column('E:E', 13.14, left)
    worksheet547.set_column('F:F', 8.57, center)
    worksheet547.set_column('G:R', 5, center)
    worksheet547.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SUKASARI', title)
    worksheet547.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet547.write('A5', 'LOKASI', header)
    worksheet547.write('B5', 'TOTAL', header)
    worksheet547.merge_range('A4:B4', 'RANK', header)
    worksheet547.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet547.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet547.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet547.merge_range('F4:F5', 'KELAS', header)
    worksheet547.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet547.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet547.write('G5', 'MAT', body)
    worksheet547.write('H5', 'SEJ', body)
    worksheet547.write('I5', 'GEO', body)
    worksheet547.write('J5', 'EKO', body)
    worksheet547.write('K5', 'SOS', body)
    worksheet547.write('L5', 'JML', body)
    worksheet547.write('M5', 'MAT', body)
    worksheet547.write('N5', 'SEJ', body)
    worksheet547.write('O5', 'GEO', body)
    worksheet547.write('P5', 'EKO', body)
    worksheet547.write('Q5', 'SOS', body)
    worksheet547.write('R5', 'JML', body)

    worksheet547.conditional_format(5, 0, row547_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet547.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF SUKASARI', title)
    worksheet547.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet547.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet547.write('A22', 'LOKASI', header)
    worksheet547.write('B22', 'TOTAL', header)
    worksheet547.merge_range('A21:B21', 'RANK', header)
    worksheet547.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet547.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet547.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet547.merge_range('F21:F22', 'KELAS', header)
    worksheet547.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet547.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet547.write('G22', 'MAT', body)
    worksheet547.write('H22', 'SEJ', body)
    worksheet547.write('I22', 'GEO', body)
    worksheet547.write('J22', 'EKO', body)
    worksheet547.write('K22', 'SOS', body)
    worksheet547.write('L22', 'JML', body)
    worksheet547.write('M22', 'MAT', body)
    worksheet547.write('N22', 'SEJ', body)
    worksheet547.write('O22', 'GEO', body)
    worksheet547.write('P22', 'EKO', body)
    worksheet547.write('Q22', 'SOS', body)
    worksheet547.write('R22', 'JML', body)

    worksheet547.conditional_format(22, 0, row547+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 548
    worksheet548.insert_image('A1', r'logo resmi nf.jpg')

    worksheet548.set_column('A:A', 7, center)
    worksheet548.set_column('B:B', 6, center)
    worksheet548.set_column('C:C', 18.14, center)
    worksheet548.set_column('D:D', 25, left)
    worksheet548.set_column('E:E', 13.14, left)
    worksheet548.set_column('F:F', 8.57, center)
    worksheet548.set_column('G:R', 5, center)
    worksheet548.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BUBULAK', title)
    worksheet548.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet548.write('A5', 'LOKASI', header)
    worksheet548.write('B5', 'TOTAL', header)
    worksheet548.merge_range('A4:B4', 'RANK', header)
    worksheet548.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet548.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet548.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet548.merge_range('F4:F5', 'KELAS', header)
    worksheet548.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet548.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet548.write('G5', 'MAT', body)
    worksheet548.write('H5', 'SEJ', body)
    worksheet548.write('I5', 'GEO', body)
    worksheet548.write('J5', 'EKO', body)
    worksheet548.write('K5', 'SOS', body)
    worksheet548.write('L5', 'JML', body)
    worksheet548.write('M5', 'MAT', body)
    worksheet548.write('N5', 'SEJ', body)
    worksheet548.write('O5', 'GEO', body)
    worksheet548.write('P5', 'EKO', body)
    worksheet548.write('Q5', 'SOS', body)
    worksheet548.write('R5', 'JML', body)

    worksheet548.conditional_format(5, 0, row548_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet548.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF BUBULAK', title)
    worksheet548.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet548.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet548.write('A22', 'LOKASI', header)
    worksheet548.write('B22', 'TOTAL', header)
    worksheet548.merge_range('A21:B21', 'RANK', header)
    worksheet548.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet548.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet548.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet548.merge_range('F21:F22', 'KELAS', header)
    worksheet548.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet548.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet548.write('G22', 'MAT', body)
    worksheet548.write('H22', 'SEJ', body)
    worksheet548.write('I22', 'GEO', body)
    worksheet548.write('J22', 'EKO', body)
    worksheet548.write('K22', 'SOS', body)
    worksheet548.write('L22', 'JML', body)
    worksheet548.write('M22', 'MAT', body)
    worksheet548.write('N22', 'SEJ', body)
    worksheet548.write('O22', 'GEO', body)
    worksheet548.write('P22', 'EKO', body)
    worksheet548.write('Q22', 'SOS', body)
    worksheet548.write('R22', 'JML', body)

    worksheet548.conditional_format(22, 0, row548+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 549
    worksheet549.insert_image('A1', r'logo resmi nf.jpg')

    worksheet549.set_column('A:A', 7, center)
    worksheet549.set_column('B:B', 6, center)
    worksheet549.set_column('C:C', 18.14, center)
    worksheet549.set_column('D:D', 25, left)
    worksheet549.set_column('E:E', 13.14, left)
    worksheet549.set_column('F:F', 8.57, center)
    worksheet549.set_column('G:R', 5, center)
    worksheet549.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CITEUREUP', title)
    worksheet549.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet549.write('A5', 'LOKASI', header)
    worksheet549.write('B5', 'TOTAL', header)
    worksheet549.merge_range('A4:B4', 'RANK', header)
    worksheet549.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet549.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet549.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet549.merge_range('F4:F5', 'KELAS', header)
    worksheet549.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet549.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet549.write('G5', 'MAT', body)
    worksheet549.write('H5', 'SEJ', body)
    worksheet549.write('I5', 'GEO', body)
    worksheet549.write('J5', 'EKO', body)
    worksheet549.write('K5', 'SOS', body)
    worksheet549.write('L5', 'JML', body)
    worksheet549.write('M5', 'MAT', body)
    worksheet549.write('N5', 'SEJ', body)
    worksheet549.write('O5', 'GEO', body)
    worksheet549.write('P5', 'EKO', body)
    worksheet549.write('Q5', 'SOS', body)
    worksheet549.write('R5', 'JML', body)

    worksheet549.conditional_format(5, 0, row549_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet549.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CITEUREUP', title)
    worksheet549.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet549.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet549.write('A22', 'LOKASI', header)
    worksheet549.write('B22', 'TOTAL', header)
    worksheet549.merge_range('A21:B21', 'RANK', header)
    worksheet549.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet549.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet549.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet549.merge_range('F21:F22', 'KELAS', header)
    worksheet549.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet549.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet549.write('G22', 'MAT', body)
    worksheet549.write('H22', 'SEJ', body)
    worksheet549.write('I22', 'GEO', body)
    worksheet549.write('J22', 'EKO', body)
    worksheet549.write('K22', 'SOS', body)
    worksheet549.write('L22', 'JML', body)
    worksheet549.write('M22', 'MAT', body)
    worksheet549.write('N22', 'SEJ', body)
    worksheet549.write('O22', 'GEO', body)
    worksheet549.write('P22', 'EKO', body)
    worksheet549.write('Q22', 'SOS', body)
    worksheet549.write('R22', 'JML', body)

    worksheet549.conditional_format(22, 0, row549+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 556
    worksheet556.insert_image('A1', r'logo resmi nf.jpg')

    worksheet556.set_column('A:A', 7, center)
    worksheet556.set_column('B:B', 6, center)
    worksheet556.set_column('C:C', 18.14, center)
    worksheet556.set_column('D:D', 25, left)
    worksheet556.set_column('E:E', 13.14, left)
    worksheet556.set_column('F:F', 8.57, center)
    worksheet556.set_column('G:R', 5, center)
    worksheet556.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF YASMIN', title)
    worksheet556.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet556.write('A5', 'LOKASI', header)
    worksheet556.write('B5', 'TOTAL', header)
    worksheet556.merge_range('A4:B4', 'RANK', header)
    worksheet556.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet556.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet556.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet556.merge_range('F4:F5', 'KELAS', header)
    worksheet556.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet556.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet556.write('G5', 'MAT', body)
    worksheet556.write('H5', 'SEJ', body)
    worksheet556.write('I5', 'GEO', body)
    worksheet556.write('J5', 'EKO', body)
    worksheet556.write('K5', 'SOS', body)
    worksheet556.write('L5', 'JML', body)
    worksheet556.write('M5', 'MAT', body)
    worksheet556.write('N5', 'SEJ', body)
    worksheet556.write('O5', 'GEO', body)
    worksheet556.write('P5', 'EKO', body)
    worksheet556.write('Q5', 'SOS', body)
    worksheet556.write('R5', 'JML', body)

    worksheet556.conditional_format(5, 0, row556_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet556.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF YASMIN', title)
    worksheet556.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet556.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet556.write('A22', 'LOKASI', header)
    worksheet556.write('B22', 'TOTAL', header)
    worksheet556.merge_range('A21:B21', 'RANK', header)
    worksheet556.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet556.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet556.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet556.merge_range('F21:F22', 'KELAS', header)
    worksheet556.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet556.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet556.write('G22', 'MAT', body)
    worksheet556.write('H22', 'SEJ', body)
    worksheet556.write('I22', 'GEO', body)
    worksheet556.write('J22', 'EKO', body)
    worksheet556.write('K22', 'SOS', body)
    worksheet556.write('L22', 'JML', body)
    worksheet556.write('M22', 'MAT', body)
    worksheet556.write('N22', 'SEJ', body)
    worksheet556.write('O22', 'GEO', body)
    worksheet556.write('P22', 'EKO', body)
    worksheet556.write('Q22', 'SOS', body)
    worksheet556.write('R22', 'JML', body)

    worksheet556.conditional_format(22, 0, row556+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 557
    worksheet557.insert_image('A1', r'logo resmi nf.jpg')

    worksheet557.set_column('A:A', 7, center)
    worksheet557.set_column('B:B', 6, center)
    worksheet557.set_column('C:C', 18.14, center)
    worksheet557.set_column('D:D', 25, left)
    worksheet557.set_column('E:E', 13.14, left)
    worksheet557.set_column('F:F', 8.57, center)
    worksheet557.set_column('G:R', 5, center)
    worksheet557.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF YOGYA PLAZA', title)
    worksheet557.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet557.write('A5', 'LOKASI', header)
    worksheet557.write('B5', 'TOTAL', header)
    worksheet557.merge_range('A4:B4', 'RANK', header)
    worksheet557.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet557.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet557.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet557.merge_range('F4:F5', 'KELAS', header)
    worksheet557.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet557.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet557.write('G5', 'MAT', body)
    worksheet557.write('H5', 'SEJ', body)
    worksheet557.write('I5', 'GEO', body)
    worksheet557.write('J5', 'EKO', body)
    worksheet557.write('K5', 'SOS', body)
    worksheet557.write('L5', 'JML', body)
    worksheet557.write('M5', 'MAT', body)
    worksheet557.write('N5', 'SEJ', body)
    worksheet557.write('O5', 'GEO', body)
    worksheet557.write('P5', 'EKO', body)
    worksheet557.write('Q5', 'SOS', body)
    worksheet557.write('R5', 'JML', body)

    worksheet557.conditional_format(5, 0, row557_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet557.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF YOGYA PLAZA', title)
    worksheet557.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet557.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet557.write('A22', 'LOKASI', header)
    worksheet557.write('B22', 'TOTAL', header)
    worksheet557.merge_range('A21:B21', 'RANK', header)
    worksheet557.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet557.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet557.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet557.merge_range('F21:F22', 'KELAS', header)
    worksheet557.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet557.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet557.write('G22', 'MAT', body)
    worksheet557.write('H22', 'SEJ', body)
    worksheet557.write('I22', 'GEO', body)
    worksheet557.write('J22', 'EKO', body)
    worksheet557.write('K22', 'SOS', body)
    worksheet557.write('L22', 'JML', body)
    worksheet557.write('M22', 'MAT', body)
    worksheet557.write('N22', 'SEJ', body)
    worksheet557.write('O22', 'GEO', body)
    worksheet557.write('P22', 'EKO', body)
    worksheet557.write('Q22', 'SOS', body)
    worksheet557.write('R22', 'JML', body)

    worksheet557.conditional_format(22, 0, row557+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 558
    worksheet558.insert_image('A1', r'logo resmi nf.jpg')

    worksheet558.set_column('A:A', 7, center)
    worksheet558.set_column('B:B', 6, center)
    worksheet558.set_column('C:C', 18.14, center)
    worksheet558.set_column('D:D', 25, left)
    worksheet558.set_column('E:E', 13.14, left)
    worksheet558.set_column('F:F', 8.57, center)
    worksheet558.set_column('G:R', 5, center)
    worksheet558.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PARUNG', title)
    worksheet558.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet558.write('A5', 'LOKASI', header)
    worksheet558.write('B5', 'TOTAL', header)
    worksheet558.merge_range('A4:B4', 'RANK', header)
    worksheet558.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet558.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet558.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet558.merge_range('F4:F5', 'KELAS', header)
    worksheet558.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet558.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet558.write('G5', 'MAT', body)
    worksheet558.write('H5', 'SEJ', body)
    worksheet558.write('I5', 'GEO', body)
    worksheet558.write('J5', 'EKO', body)
    worksheet558.write('K5', 'SOS', body)
    worksheet558.write('L5', 'JML', body)
    worksheet558.write('M5', 'MAT', body)
    worksheet558.write('N5', 'SEJ', body)
    worksheet558.write('O5', 'GEO', body)
    worksheet558.write('P5', 'EKO', body)
    worksheet558.write('Q5', 'SOS', body)
    worksheet558.write('R5', 'JML', body)

    worksheet558.conditional_format(5, 0, row558_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet558.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PARUNG', title)
    worksheet558.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet558.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet558.write('A22', 'LOKASI', header)
    worksheet558.write('B22', 'TOTAL', header)
    worksheet558.merge_range('A21:B21', 'RANK', header)
    worksheet558.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet558.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet558.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet558.merge_range('F21:F22', 'KELAS', header)
    worksheet558.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet558.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet558.write('G22', 'MAT', body)
    worksheet558.write('H22', 'SEJ', body)
    worksheet558.write('I22', 'GEO', body)
    worksheet558.write('J22', 'EKO', body)
    worksheet558.write('K22', 'SOS', body)
    worksheet558.write('L22', 'JML', body)
    worksheet558.write('M22', 'MAT', body)
    worksheet558.write('N22', 'SEJ', body)
    worksheet558.write('O22', 'GEO', body)
    worksheet558.write('P22', 'EKO', body)
    worksheet558.write('Q22', 'SOS', body)
    worksheet558.write('R22', 'JML', body)

    worksheet558.conditional_format(22, 0, row558+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 575
    worksheet575.insert_image('A1', r'logo resmi nf.jpg')

    worksheet575.set_column('A:A', 7, center)
    worksheet575.set_column('B:B', 6, center)
    worksheet575.set_column('C:C', 18.14, center)
    worksheet575.set_column('D:D', 25, left)
    worksheet575.set_column('E:E', 13.14, left)
    worksheet575.set_column('F:F', 8.57, center)
    worksheet575.set_column('G:R', 5, center)
    worksheet575.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CINERE', title)
    worksheet575.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet575.write('A5', 'LOKASI', header)
    worksheet575.write('B5', 'TOTAL', header)
    worksheet575.merge_range('A4:B4', 'RANK', header)
    worksheet575.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet575.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet575.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet575.merge_range('F4:F5', 'KELAS', header)
    worksheet575.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet575.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet575.write('G5', 'MAT', body)
    worksheet575.write('H5', 'SEJ', body)
    worksheet575.write('I5', 'GEO', body)
    worksheet575.write('J5', 'EKO', body)
    worksheet575.write('K5', 'SOS', body)
    worksheet575.write('L5', 'JML', body)
    worksheet575.write('M5', 'MAT', body)
    worksheet575.write('N5', 'SEJ', body)
    worksheet575.write('O5', 'GEO', body)
    worksheet575.write('P5', 'EKO', body)
    worksheet575.write('Q5', 'SOS', body)
    worksheet575.write('R5', 'JML', body)

    worksheet575.conditional_format(5, 0, row575_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet575.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CINERE', title)
    worksheet575.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet575.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet575.write('A22', 'LOKASI', header)
    worksheet575.write('B22', 'TOTAL', header)
    worksheet575.merge_range('A21:B21', 'RANK', header)
    worksheet575.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet575.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet575.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet575.merge_range('F21:F22', 'KELAS', header)
    worksheet575.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet575.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet575.write('G22', 'MAT', body)
    worksheet575.write('H22', 'SEJ', body)
    worksheet575.write('I22', 'GEO', body)
    worksheet575.write('J22', 'EKO', body)
    worksheet575.write('K22', 'SOS', body)
    worksheet575.write('L22', 'JML', body)
    worksheet575.write('M22', 'MAT', body)
    worksheet575.write('N22', 'SEJ', body)
    worksheet575.write('O22', 'GEO', body)
    worksheet575.write('P22', 'EKO', body)
    worksheet575.write('Q22', 'SOS', body)
    worksheet575.write('R22', 'JML', body)

    worksheet575.conditional_format(22, 0, row575+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 576
    worksheet576.insert_image('A1', r'logo resmi nf.jpg')

    worksheet576.set_column('A:A', 7, center)
    worksheet576.set_column('B:B', 6, center)
    worksheet576.set_column('C:C', 18.14, center)
    worksheet576.set_column('D:D', 25, left)
    worksheet576.set_column('E:E', 13.14, left)
    worksheet576.set_column('F:F', 8.57, center)
    worksheet576.set_column('G:R', 5, center)
    worksheet576.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIBINONG', title)
    worksheet576.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet576.write('A5', 'LOKASI', header)
    worksheet576.write('B5', 'TOTAL', header)
    worksheet576.merge_range('A4:B4', 'RANK', header)
    worksheet576.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet576.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet576.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet576.merge_range('F4:F5', 'KELAS', header)
    worksheet576.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet576.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet576.write('G5', 'MAT', body)
    worksheet576.write('H5', 'SEJ', body)
    worksheet576.write('I5', 'GEO', body)
    worksheet576.write('J5', 'EKO', body)
    worksheet576.write('K5', 'SOS', body)
    worksheet576.write('L5', 'JML', body)
    worksheet576.write('M5', 'MAT', body)
    worksheet576.write('N5', 'SEJ', body)
    worksheet576.write('O5', 'GEO', body)
    worksheet576.write('P5', 'EKO', body)
    worksheet576.write('Q5', 'SOS', body)
    worksheet576.write('R5', 'JML', body)

    worksheet576.conditional_format(5, 0, row576_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet576.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF CIBINONG', title)
    worksheet576.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet576.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet576.write('A22', 'LOKASI', header)
    worksheet576.write('B22', 'TOTAL', header)
    worksheet576.merge_range('A21:B21', 'RANK', header)
    worksheet576.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet576.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet576.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet576.merge_range('F21:F22', 'KELAS', header)
    worksheet576.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet576.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet576.write('G22', 'MAT', body)
    worksheet576.write('H22', 'SEJ', body)
    worksheet576.write('I22', 'GEO', body)
    worksheet576.write('J22', 'EKO', body)
    worksheet576.write('K22', 'SOS', body)
    worksheet576.write('L22', 'JML', body)
    worksheet576.write('M22', 'MAT', body)
    worksheet576.write('N22', 'SEJ', body)
    worksheet576.write('O22', 'GEO', body)
    worksheet576.write('P22', 'EKO', body)
    worksheet576.write('Q22', 'SOS', body)
    worksheet576.write('R22', 'JML', body)

    worksheet576.conditional_format(22, 0, row576+21, 17,
                                    {'type': 'no_errors', 'format': border})
    # worksheet 577
    worksheet577.insert_image('A1', r'logo resmi nf.jpg')

    worksheet577.set_column('A:A', 7, center)
    worksheet577.set_column('B:B', 6, center)
    worksheet577.set_column('C:C', 18.14, center)
    worksheet577.set_column('D:D', 25, left)
    worksheet577.set_column('E:E', 13.14, left)
    worksheet577.set_column('F:F', 8.57, center)
    worksheet577.set_column('G:R', 5, center)
    worksheet577.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF VETERAN (RUMAH SAKIT)', title)
    worksheet577.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet577.write('A5', 'LOKASI', header)
    worksheet577.write('B5', 'TOTAL', header)
    worksheet577.merge_range('A4:B4', 'RANK', header)
    worksheet577.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet577.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet577.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet577.merge_range('F4:F5', 'KELAS', header)
    worksheet577.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet577.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet577.write('G5', 'MAT', body)
    worksheet577.write('H5', 'SEJ', body)
    worksheet577.write('I5', 'GEO', body)
    worksheet577.write('J5', 'EKO', body)
    worksheet577.write('K5', 'SOS', body)
    worksheet577.write('L5', 'JML', body)
    worksheet577.write('M5', 'MAT', body)
    worksheet577.write('N5', 'SEJ', body)
    worksheet577.write('O5', 'GEO', body)
    worksheet577.write('P5', 'EKO', body)
    worksheet577.write('Q5', 'SOS', body)
    worksheet577.write('R5', 'JML', body)

    worksheet577.conditional_format(5, 0, row577_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet577.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF VETERAN (RUMAH SAKIT)', title)
    worksheet577.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet577.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet577.write('A22', 'LOKASI', header)
    worksheet577.write('B22', 'TOTAL', header)
    worksheet577.merge_range('A21:B21', 'RANK', header)
    worksheet577.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet577.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet577.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet577.merge_range('F21:F22', 'KELAS', header)
    worksheet577.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet577.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet577.write('G22', 'MAT', body)
    worksheet577.write('H22', 'SEJ', body)
    worksheet577.write('I22', 'GEO', body)
    worksheet577.write('J22', 'EKO', body)
    worksheet577.write('K22', 'SOS', body)
    worksheet577.write('L22', 'JML', body)
    worksheet577.write('M22', 'MAT', body)
    worksheet577.write('N22', 'SEJ', body)
    worksheet577.write('O22', 'GEO', body)
    worksheet577.write('P22', 'EKO', body)
    worksheet577.write('Q22', 'SOS', body)
    worksheet577.write('R22', 'JML', body)

    worksheet577.conditional_format(22, 0, row577+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 578
    worksheet578.insert_image('A1', r'logo resmi nf.jpg')

    worksheet578.set_column('A:A', 7, center)
    worksheet578.set_column('B:B', 6, center)
    worksheet578.set_column('C:C', 18.14, center)
    worksheet578.set_column('D:D', 25, left)
    worksheet578.set_column('E:E', 13.14, left)
    worksheet578.set_column('F:F', 8.57, center)
    worksheet578.set_column('G:R', 5, center)
    worksheet578.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MARTADINATA', title)
    worksheet578.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet578.write('A5', 'LOKASI', header)
    worksheet578.write('B5', 'TOTAL', header)
    worksheet578.merge_range('A4:B4', 'RANK', header)
    worksheet578.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet578.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet578.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet578.merge_range('F4:F5', 'KELAS', header)
    worksheet578.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet578.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet578.write('G5', 'MAT', body)
    worksheet578.write('H5', 'SEJ', body)
    worksheet578.write('I5', 'GEO', body)
    worksheet578.write('J5', 'EKO', body)
    worksheet578.write('K5', 'SOS', body)
    worksheet578.write('L5', 'JML', body)
    worksheet578.write('M5', 'MAT', body)
    worksheet578.write('N5', 'SEJ', body)
    worksheet578.write('O5', 'GEO', body)
    worksheet578.write('P5', 'EKO', body)
    worksheet578.write('Q5', 'SOS', body)
    worksheet578.write('R5', 'JML', body)

    worksheet578.conditional_format(5, 0, row578_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet578.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF MARTADINATA', title)
    worksheet578.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet578.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet578.write('A22', 'LOKASI', header)
    worksheet578.write('B22', 'TOTAL', header)
    worksheet578.merge_range('A21:B21', 'RANK', header)
    worksheet578.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet578.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet578.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet578.merge_range('F21:F22', 'KELAS', header)
    worksheet578.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet578.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet578.write('G22', 'MAT', body)
    worksheet578.write('H22', 'SEJ', body)
    worksheet578.write('I22', 'GEO', body)
    worksheet578.write('J22', 'EKO', body)
    worksheet578.write('K22', 'SOS', body)
    worksheet578.write('L22', 'JML', body)
    worksheet578.write('M22', 'MAT', body)
    worksheet578.write('N22', 'SEJ', body)
    worksheet578.write('O22', 'GEO', body)
    worksheet578.write('P22', 'EKO', body)
    worksheet578.write('Q22', 'SOS', body)
    worksheet578.write('R22', 'JML', body)

    worksheet578.conditional_format(22, 0, row578+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 588
    worksheet588.insert_image('A1', r'logo resmi nf.jpg')

    worksheet588.set_column('A:A', 7, center)
    worksheet588.set_column('B:B', 6, center)
    worksheet588.set_column('C:C', 18.14, center)
    worksheet588.set_column('D:D', 25, left)
    worksheet588.set_column('E:E', 13.14, left)
    worksheet588.set_column('F:F', 8.57, center)
    worksheet588.set_column('G:R', 5, center)
    worksheet588.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MAHARAJA', title)
    worksheet588.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet588.write('A5', 'LOKASI', header)
    worksheet588.write('B5', 'TOTAL', header)
    worksheet588.merge_range('A4:B4', 'RANK', header)
    worksheet588.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet588.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet588.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet588.merge_range('F4:F5', 'KELAS', header)
    worksheet588.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet588.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet588.write('G5', 'MAT', body)
    worksheet588.write('H5', 'SEJ', body)
    worksheet588.write('I5', 'GEO', body)
    worksheet588.write('J5', 'EKO', body)
    worksheet588.write('K5', 'SOS', body)
    worksheet588.write('L5', 'JML', body)
    worksheet588.write('M5', 'MAT', body)
    worksheet588.write('N5', 'SEJ', body)
    worksheet588.write('O5', 'GEO', body)
    worksheet588.write('P5', 'EKO', body)
    worksheet588.write('Q5', 'SOS', body)
    worksheet588.write('R5', 'JML', body)

    worksheet588.conditional_format(5, 0, row588_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet588.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF MAHARAJA', title)
    worksheet588.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet588.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet588.write('A22', 'LOKASI', header)
    worksheet588.write('B22', 'TOTAL', header)
    worksheet588.merge_range('A21:B21', 'RANK', header)
    worksheet588.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet588.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet588.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet588.merge_range('F21:F22', 'KELAS', header)
    worksheet588.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet588.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet588.write('G22', 'MAT', body)
    worksheet588.write('H22', 'SEJ', body)
    worksheet588.write('I22', 'GEO', body)
    worksheet588.write('J22', 'EKO', body)
    worksheet588.write('K22', 'SOS', body)
    worksheet588.write('L22', 'JML', body)
    worksheet588.write('M22', 'MAT', body)
    worksheet588.write('N22', 'SEJ', body)
    worksheet588.write('O22', 'GEO', body)
    worksheet588.write('P22', 'EKO', body)
    worksheet588.write('Q22', 'SOS', body)
    worksheet588.write('R22', 'JML', body)

    worksheet588.conditional_format(22, 0, row588+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 589
    worksheet589.insert_image('A1', r'logo resmi nf.jpg')

    worksheet589.set_column('A:A', 7, center)
    worksheet589.set_column('B:B', 6, center)
    worksheet589.set_column('C:C', 18.14, center)
    worksheet589.set_column('D:D', 25, left)
    worksheet589.set_column('E:E', 13.14, left)
    worksheet589.set_column('F:F', 8.57, center)
    worksheet589.set_column('G:R', 5, center)
    worksheet589.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF WARUNG JAMBU', title)
    worksheet589.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet589.write('A5', 'LOKASI', header)
    worksheet589.write('B5', 'TOTAL', header)
    worksheet589.merge_range('A4:B4', 'RANK', header)
    worksheet589.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet589.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet589.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet589.merge_range('F4:F5', 'KELAS', header)
    worksheet589.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet589.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet589.write('G5', 'MAT', body)
    worksheet589.write('H5', 'SEJ', body)
    worksheet589.write('I5', 'GEO', body)
    worksheet589.write('J5', 'EKO', body)
    worksheet589.write('K5', 'SOS', body)
    worksheet589.write('L5', 'JML', body)
    worksheet589.write('M5', 'MAT', body)
    worksheet589.write('N5', 'SEJ', body)
    worksheet589.write('O5', 'GEO', body)
    worksheet589.write('P5', 'EKO', body)
    worksheet589.write('Q5', 'SOS', body)
    worksheet589.write('R5', 'JML', body)

    worksheet589.conditional_format(5, 0, row589_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet589.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF WARUNG JAMBU', title)
    worksheet589.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet589.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet589.write('A22', 'LOKASI', header)
    worksheet589.write('B22', 'TOTAL', header)
    worksheet589.merge_range('A21:B21', 'RANK', header)
    worksheet589.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet589.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet589.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet589.merge_range('F21:F22', 'KELAS', header)
    worksheet589.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet589.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet589.write('G22', 'MAT', body)
    worksheet589.write('H22', 'SEJ', body)
    worksheet589.write('I22', 'GEO', body)
    worksheet589.write('J22', 'EKO', body)
    worksheet589.write('K22', 'SOS', body)
    worksheet589.write('L22', 'JML', body)
    worksheet589.write('M22', 'MAT', body)
    worksheet589.write('N22', 'SEJ', body)
    worksheet589.write('O22', 'GEO', body)
    worksheet589.write('P22', 'EKO', body)
    worksheet589.write('Q22', 'SOS', body)
    worksheet589.write('R22', 'JML', body)

    worksheet589.conditional_format(22, 0, row589+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 594
    worksheet594.insert_image('A1', r'logo resmi nf.jpg')

    worksheet594.set_column('A:A', 7, center)
    worksheet594.set_column('B:B', 6, center)
    worksheet594.set_column('C:C', 18.14, center)
    worksheet594.set_column('D:D', 25, left)
    worksheet594.set_column('E:E', 13.14, left)
    worksheet594.set_column('F:F', 8.57, center)
    worksheet594.set_column('G:R', 5, center)
    worksheet594.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PEMDA CIBINONG', title)
    worksheet594.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet594.write('A5', 'LOKASI', header)
    worksheet594.write('B5', 'TOTAL', header)
    worksheet594.merge_range('A4:B4', 'RANK', header)
    worksheet594.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet594.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet594.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet594.merge_range('F4:F5', 'KELAS', header)
    worksheet594.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet594.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet594.write('G5', 'MAT', body)
    worksheet594.write('H5', 'SEJ', body)
    worksheet594.write('I5', 'GEO', body)
    worksheet594.write('J5', 'EKO', body)
    worksheet594.write('K5', 'SOS', body)
    worksheet594.write('L5', 'JML', body)
    worksheet594.write('M5', 'MAT', body)
    worksheet594.write('N5', 'SEJ', body)
    worksheet594.write('O5', 'GEO', body)
    worksheet594.write('P5', 'EKO', body)
    worksheet594.write('Q5', 'SOS', body)
    worksheet594.write('R5', 'JML', body)

    worksheet594.conditional_format(5, 0, row594_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet594.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF PEMDA CIBINONG', title)
    worksheet594.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet594.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet594.write('A22', 'LOKASI', header)
    worksheet594.write('B22', 'TOTAL', header)
    worksheet594.merge_range('A21:B21', 'RANK', header)
    worksheet594.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet594.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet594.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet594.merge_range('F21:F22', 'KELAS', header)
    worksheet594.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet594.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet594.write('G22', 'MAT', body)
    worksheet594.write('H22', 'SEJ', body)
    worksheet594.write('I22', 'GEO', body)
    worksheet594.write('J22', 'EKO', body)
    worksheet594.write('K22', 'SOS', body)
    worksheet594.write('L22', 'JML', body)
    worksheet594.write('M22', 'MAT', body)
    worksheet594.write('N22', 'SEJ', body)
    worksheet594.write('O22', 'GEO', body)
    worksheet594.write('P22', 'EKO', body)
    worksheet594.write('Q22', 'SOS', body)
    worksheet594.write('R22', 'JML', body)

    worksheet594.conditional_format(22, 0, row594+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 661
    worksheet661.insert_image('A1', r'logo resmi nf.jpg')

    worksheet661.set_column('A:A', 7, center)
    worksheet661.set_column('B:B', 6, center)
    worksheet661.set_column('C:C', 18.14, center)
    worksheet661.set_column('D:D', 25, left)
    worksheet661.set_column('E:E', 13.14, left)
    worksheet661.set_column('F:F', 8.57, center)
    worksheet661.set_column('G:R', 5, center)
    worksheet661.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF GAJAH MADA', title)
    worksheet661.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet661.write('A5', 'LOKASI', header)
    worksheet661.write('B5', 'TOTAL', header)
    worksheet661.merge_range('A4:B4', 'RANK', header)
    worksheet661.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet661.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet661.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet661.merge_range('F4:F5', 'KELAS', header)
    worksheet661.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet661.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet661.write('G5', 'MAT', body)
    worksheet661.write('H5', 'SEJ', body)
    worksheet661.write('I5', 'GEO', body)
    worksheet661.write('J5', 'EKO', body)
    worksheet661.write('K5', 'SOS', body)
    worksheet661.write('L5', 'JML', body)
    worksheet661.write('M5', 'MAT', body)
    worksheet661.write('N5', 'SEJ', body)
    worksheet661.write('O5', 'GEO', body)
    worksheet661.write('P5', 'EKO', body)
    worksheet661.write('Q5', 'SOS', body)
    worksheet661.write('R5', 'JML', body)

    worksheet661.conditional_format(5, 0, row661_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet661.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF GAJAH MADA', title)
    worksheet661.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet661.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet661.write('A22', 'LOKASI', header)
    worksheet661.write('B22', 'TOTAL', header)
    worksheet661.merge_range('A21:B21', 'RANK', header)
    worksheet661.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet661.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet661.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet661.merge_range('F21:F22', 'KELAS', header)
    worksheet661.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet661.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet661.write('G22', 'MAT', body)
    worksheet661.write('H22', 'SEJ', body)
    worksheet661.write('I22', 'GEO', body)
    worksheet661.write('J22', 'EKO', body)
    worksheet661.write('K22', 'SOS', body)
    worksheet661.write('L22', 'JML', body)
    worksheet661.write('M22', 'MAT', body)
    worksheet661.write('N22', 'SEJ', body)
    worksheet661.write('O22', 'GEO', body)
    worksheet661.write('P22', 'EKO', body)
    worksheet661.write('Q22', 'SOS', body)
    worksheet661.write('R22', 'JML', body)

    worksheet661.conditional_format(22, 0, row661+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 662
    worksheet662.insert_image('A1', r'logo resmi nf.jpg')

    worksheet662.set_column('A:A', 7, center)
    worksheet662.set_column('B:B', 6, center)
    worksheet662.set_column('C:C', 18.14, center)
    worksheet662.set_column('D:D', 25, left)
    worksheet662.set_column('E:E', 13.14, left)
    worksheet662.set_column('F:F', 8.57, center)
    worksheet662.set_column('G:R', 5, center)
    worksheet662.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF LOLONG BELANTI', title)
    worksheet662.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet662.write('A5', 'LOKASI', header)
    worksheet662.write('B5', 'TOTAL', header)
    worksheet662.merge_range('A4:B4', 'RANK', header)
    worksheet662.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet662.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet662.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet662.merge_range('F4:F5', 'KELAS', header)
    worksheet662.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet662.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet662.write('G5', 'MAT', body)
    worksheet662.write('H5', 'SEJ', body)
    worksheet662.write('I5', 'GEO', body)
    worksheet662.write('J5', 'EKO', body)
    worksheet662.write('K5', 'SOS', body)
    worksheet662.write('L5', 'JML', body)
    worksheet662.write('M5', 'MAT', body)
    worksheet662.write('N5', 'SEJ', body)
    worksheet662.write('O5', 'GEO', body)
    worksheet662.write('P5', 'EKO', body)
    worksheet662.write('Q5', 'SOS', body)
    worksheet662.write('R5', 'JML', body)

    worksheet662.conditional_format(5, 0, row662_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet662.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF LOLONG BELANTI', title)
    worksheet662.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet662.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet662.write('A22', 'LOKASI', header)
    worksheet662.write('B22', 'TOTAL', header)
    worksheet662.merge_range('A21:B21', 'RANK', header)
    worksheet662.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet662.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet662.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet662.merge_range('F21:F22', 'KELAS', header)
    worksheet662.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet662.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet662.write('G22', 'MAT', body)
    worksheet662.write('H22', 'SEJ', body)
    worksheet662.write('I22', 'GEO', body)
    worksheet662.write('J22', 'EKO', body)
    worksheet662.write('K22', 'SOS', body)
    worksheet662.write('L22', 'JML', body)
    worksheet662.write('M22', 'MAT', body)
    worksheet662.write('N22', 'SEJ', body)
    worksheet662.write('O22', 'GEO', body)
    worksheet662.write('P22', 'EKO', body)
    worksheet662.write('Q22', 'SOS', body)
    worksheet662.write('R22', 'JML', body)

    worksheet662.conditional_format(22, 0, row662+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 663
    worksheet663.insert_image('A1', r'logo resmi nf.jpg')

    worksheet663.set_column('A:A', 7, center)
    worksheet663.set_column('B:B', 6, center)
    worksheet663.set_column('C:C', 18.14, center)
    worksheet663.set_column('D:D', 25, left)
    worksheet663.set_column('E:E', 13.14, left)
    worksheet663.set_column('F:F', 8.57, center)
    worksheet663.set_column('G:R', 5, center)
    worksheet663.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SOETOMO', title)
    worksheet663.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet663.write('A5', 'LOKASI', header)
    worksheet663.write('B5', 'TOTAL', header)
    worksheet663.merge_range('A4:B4', 'RANK', header)
    worksheet663.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet663.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet663.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet663.merge_range('F4:F5', 'KELAS', header)
    worksheet663.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet663.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet663.write('G5', 'MAT', body)
    worksheet663.write('H5', 'SEJ', body)
    worksheet663.write('I5', 'GEO', body)
    worksheet663.write('J5', 'EKO', body)
    worksheet663.write('K5', 'SOS', body)
    worksheet663.write('L5', 'JML', body)
    worksheet663.write('M5', 'MAT', body)
    worksheet663.write('N5', 'SEJ', body)
    worksheet663.write('O5', 'GEO', body)
    worksheet663.write('P5', 'EKO', body)
    worksheet663.write('Q5', 'SOS', body)
    worksheet663.write('R5', 'JML', body)

    worksheet663.conditional_format(5, 0, row663_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet663.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF SOETOMO', title)
    worksheet663.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet663.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet663.write('A22', 'LOKASI', header)
    worksheet663.write('B22', 'TOTAL', header)
    worksheet663.merge_range('A21:B21', 'RANK', header)
    worksheet663.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet663.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet663.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet663.merge_range('F21:F22', 'KELAS', header)
    worksheet663.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet663.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet663.write('G22', 'MAT', body)
    worksheet663.write('H22', 'SEJ', body)
    worksheet663.write('I22', 'GEO', body)
    worksheet663.write('J22', 'EKO', body)
    worksheet663.write('K22', 'SOS', body)
    worksheet663.write('L22', 'JML', body)
    worksheet663.write('M22', 'MAT', body)
    worksheet663.write('N22', 'SEJ', body)
    worksheet663.write('O22', 'GEO', body)
    worksheet663.write('P22', 'EKO', body)
    worksheet663.write('Q22', 'SOS', body)
    worksheet663.write('R22', 'JML', body)

    worksheet663.conditional_format(22, 0, row663+21, 17,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 664
    worksheet664.insert_image('A1', r'logo resmi nf.jpg')

    worksheet664.set_column('A:A', 7, center)
    worksheet664.set_column('B:B', 6, center)
    worksheet664.set_column('C:C', 18.14, center)
    worksheet664.set_column('D:D', 25, left)
    worksheet664.set_column('E:E', 13.14, left)
    worksheet664.set_column('F:F', 8.57, center)
    worksheet664.set_column('G:R', 5, center)
    worksheet664.merge_range(
        'A1:R1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF TAN MALAKA', title)
    worksheet664.merge_range(
        'A2:R2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet664.write('A5', 'LOKASI', header)
    worksheet664.write('B5', 'TOTAL', header)
    worksheet664.merge_range('A4:B4', 'RANK', header)
    worksheet664.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet664.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet664.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet664.merge_range('F4:F5', 'KELAS', header)
    worksheet664.merge_range('G4:L4', 'JUMLAH BENAR', header)
    worksheet664.merge_range('M4:R4', 'NILAI STANDAR', header)
    worksheet664.write('G5', 'MAT', body)
    worksheet664.write('H5', 'SEJ', body)
    worksheet664.write('I5', 'GEO', body)
    worksheet664.write('J5', 'EKO', body)
    worksheet664.write('K5', 'SOS', body)
    worksheet664.write('L5', 'JML', body)
    worksheet664.write('M5', 'MAT', body)
    worksheet664.write('N5', 'SEJ', body)
    worksheet664.write('O5', 'GEO', body)
    worksheet664.write('P5', 'EKO', body)
    worksheet664.write('Q5', 'SOS', body)
    worksheet664.write('R5', 'JML', body)

    worksheet664.conditional_format(5, 0, row664_10+4, 17,
                                    {'type': 'no_errors', 'format': border})

    worksheet664.merge_range(
        'A17:R17', fr'KELAS {kelas} - LOKASI NF TAN MALAKA', title)
    worksheet664.merge_range('A18:R18', fr'{penilaian}', subTitle)
    worksheet664.merge_range(
        'A19:R19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet664.write('A22', 'LOKASI', header)
    worksheet664.write('B22', 'TOTAL', header)
    worksheet664.merge_range('A21:B21', 'RANK', header)
    worksheet664.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet664.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet664.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet664.merge_range('F21:F22', 'KELAS', header)
    worksheet664.merge_range('G21:L21', 'JUMLAH BENAR', header)
    worksheet664.merge_range('M21:R21', 'NILAI STANDAR', header)
    worksheet664.write('G22', 'MAT', body)
    worksheet664.write('H22', 'SEJ', body)
    worksheet664.write('I22', 'GEO', body)
    worksheet664.write('J22', 'EKO', body)
    worksheet664.write('K22', 'SOS', body)
    worksheet664.write('L22', 'JML', body)
    worksheet664.write('M22', 'MAT', body)
    worksheet664.write('N22', 'SEJ', body)
    worksheet664.write('O22', 'GEO', body)
    worksheet664.write('P22', 'EKO', body)
    worksheet664.write('Q22', 'SOS', body)
    worksheet664.write('R22', 'JML', body)

    worksheet664.conditional_format(22, 0, row664+21, 17,
                                    {'type': 'no_errors', 'format': border})

    workbook.close()
    st.success("File siap diunduh!")

    # Tombol unduh file
    with open(file_path, "rb") as f:
        bytes_data = f.read()
    st.download_button(label="Unduh File", data=bytes_data,
                       file_name=file_name)
