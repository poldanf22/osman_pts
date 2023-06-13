uploaded_file = st.file_uploader(
    'Letakkan file excel NILAI STANDAR [LOKASI 162-236]', type='xlsx')

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
    sort162 = df[df['LOKASI'] == 162]
    sort163 = df[df['LOKASI'] == 163]
    sort164 = df[df['LOKASI'] == 164]
    sort165 = df[df['LOKASI'] == 165]
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
    # sort139 = df[df['LOKASI']==139]
    sort203 = df[df['LOKASI'] == 203]
    sort210 = df[df['LOKASI'] == 210]
    sort211 = df[df['LOKASI'] == 211]
    sort212 = df[df['LOKASI'] == 212]
    sort216 = df[df['LOKASI'] == 216]
    sort217 = df[df['LOKASI'] == 217]
    sort218 = df[df['LOKASI'] == 218]
    sort219 = df[df['LOKASI'] == 219]
    sort220 = df[df['LOKASI'] == 220]
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
    # sort160 = df[df['LOKASI'] == 160]

    # 10 besar setiap lokasi
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
    # # 104
    # sort104_10=sort104.sort_values(by=['RANK LOK.'], ascending=[True])
    # del sort104_10['LOKASI']
    # sort104_10=sort104_10.drop(sort104_10[(sort104_10['RANK LOK.']>10)].index)
    # 165
    sort165_10 = sort165.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort165_10['LOKASI']
    sort165_10 = sort165_10.drop(
        sort165_10[(sort165_10['RANK LOK.'] > 10)].index)
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
    # # 114
    # sort114_10=sort114.sort_values(by=['RANK LOK.'], ascending=[True])
    # del sort114_10['LOKASI']
    # sort114_10=sort114_10.drop(sort114_10[(sort114_10['RANK LOK.']>10)].index)
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
    # # 139
    # sort139_10=sort139.sort_values(by=['RANK LOK.'], ascending=[True])
    # del sort139_10['LOKASI']
    # sort139_10=sort139_10.drop(sort139_10[(sort139_10['RANK LOK.']>10)].index)
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
    # 160
    # sort160_10 = sort160.sort_values(by=['RANK LOK.'], ascending=[True])
    # del sort160_10['LOKASI']
    # sort160_10 = sort160_10.drop(
    #     sort160_10[(sort160_10['RANK LOK.'] > 10)].index)

    # All 162
    sort162 = sort162.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort162['LOKASI']
    # All 163
    sort163 = sort163.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort163['LOKASI']
    # All 164
    sort164 = sort164.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort164['LOKASI']
    # # All 104
    # sort104=sort104.sort_values(by=['NAMA SISWA'], ascending=[True])
    # del sort104['LOKASI']
    # All 165
    sort165 = sort165.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort165['LOKASI']
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
    # # All 114
    # sort114=sort114.sort_values(by=['NAMA SISWA'], ascending=[True])
    # del sort114['LOKASI']
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
    # # All 139
    # sort139=sort139.sort_values(by=['NAMA SISWA'], ascending=[True])
    # del sort139['LOKASI']
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
    # All 160
    # sort160 = sort160.sort_values(by=['NAMA SISWA'], ascending=[True])
    # del sort160['LOKASI']

    # jumlah row
    # 162
    row162_10 = sort162_10.shape[0]
    row162 = sort162.shape[0]
    # 163
    row163_10 = sort163_10.shape[0]
    row163 = sort163.shape[0]
    # 164
    row164_10 = sort164_10.shape[0]
    row164 = sort164.shape[0]
    # # 104
    # row104_10=sort104_10.shape[0]
    # row104=sort104.shape[0]
    # 165
    row165_10 = sort165_10.shape[0]
    row165 = sort165.shape[0]
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
    # # 114
    # row114_10=sort114_10.shape[0]
    # row114=sort114.shape[0]
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
    # # 139
    # row139_10=sort139_10.shape[0]
    # row139=sort139.shape[0]
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
    # 160
    # row160_10 = sort160_10.shape[0]
    # row160 = sort160.shape[0]

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    file_name = f"{kelas}_{penilaian}_{semester}_lokasi_162_160.xlsx"
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
    # 160
    # Convert the dataframe to an XlsxWriter Excel object.
    # sort160_10.to_excel(writer, sheet_name='160',
    #                     startrow=5,
    #                     startcol=0,
    #                     index=False,
    #                     header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    # sort160.to_excel(writer, sheet_name='160',
    #                  startrow=22,
    #                  startcol=0,
    #                  index=False,
    #                  header=False)

    # Get the xlsxwriter objects from the dataframe writer object.
    workbook = writer.book

    # membuat worksheet baru
    worksheetcover = writer.sheets['cover']
    worksheet162 = writer.sheets['162']
    worksheet163 = writer.sheets['163']
    worksheet164 = writer.sheets['164']
    # worksheet104 = writer.sheets['104']
    worksheet165 = writer.sheets['165']
    worksheet167 = writer.sheets['167']
    worksheet168 = writer.sheets['168']
    worksheet169 = writer.sheets['169']
    worksheet171 = writer.sheets['171']
    worksheet173 = writer.sheets['173']
    worksheet174 = writer.sheets['174']
    worksheet175 = writer.sheets['175']
    worksheet176 = writer.sheets['176']
    # worksheet114 = writer.sheets['114']
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
    # worksheet139 = writer.sheets['139']
    worksheet203 = writer.sheets['203']
    worksheet210 = writer.sheets['210']
    worksheet211 = writer.sheets['211']
    worksheet212 = writer.sheets['212']
    worksheet216 = writer.sheets['216']
    worksheet217 = writer.sheets['217']
    worksheet218 = writer.sheets['218']
    worksheet219 = writer.sheets['219']
    worksheet220 = writer.sheets['220']
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
    # worksheet160 = writer.sheets['160']

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

    # worksheet 162
    worksheet162.insert_image('A1', r'logo resmi nf.jpg')

    worksheet162.set_column('A:A', 7, center)
    worksheet162.set_column('B:B', 6, center)
    worksheet162.set_column('C:C', 18.14, center)
    worksheet162.set_column('D:D', 25, left)
    worksheet162.set_column('E:E', 13.14, left)
    worksheet162.set_column('F:F', 8.57, center)
    worksheet162.set_column('G:V', 5, center)
    worksheet162.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF TAMAN MARGASATWA', title)
    worksheet162.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet162.write('A5', 'LOKASI', header)
    worksheet162.write('B5', 'TOTAL', header)
    worksheet162.merge_range('A4:B4', 'RANK', header)
    worksheet162.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet162.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet162.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet162.merge_range('F4:F5', 'KELAS', header)
    worksheet162.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet162.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet162.write('G5', 'MAW', body)
    worksheet162.write('H5', 'MAP', body)
    worksheet162.write('I5', 'IND', body)
    worksheet162.write('J5', 'ENG', body)
    worksheet162.write('K5', 'SEJ', body)
    worksheet162.write('L5', 'GEO', body)
    worksheet162.write('M5', 'EKO', body)
    worksheet162.write('N5', 'SOS', body)
    worksheet162.write('O5', 'FIS', body)
    worksheet162.write('P5', 'KIM', body)
    worksheet162.write('Q5', 'BIO', body)
    worksheet162.write('R5', 'JML', body)
    worksheet162.write('S5', 'MAW', body)
    worksheet162.write('T5', 'MAP', body)
    worksheet162.write('U5', 'IND', body)
    worksheet162.write('V5', 'ENG', body)
    worksheet162.write('W5', 'SEJ', body)
    worksheet162.write('X5', 'GEO', body)
    worksheet162.write('Y5', 'EKO', body)
    worksheet162.write('Z5', 'SOS', body)
    worksheet162.write('AA5', 'FIS', body)
    worksheet162.write('AB5', 'KIM', body)
    worksheet162.write('AC5', 'BIO', body)
    worksheet162.write('AD5', 'JML', body)

    worksheet162.conditional_format(5, 0, row162_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet162.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF TAMAN MARGASATWA', title)
    worksheet162.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet162.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet162.write('A22', 'LOKASI', header)
    worksheet162.write('B22', 'TOTAL', header)
    worksheet162.merge_range('A21:B21', 'RANK', header)
    worksheet162.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet162.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet162.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet162.merge_range('F21:F22', 'KELAS', header)
    worksheet162.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet162.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet162.write('G22', 'MAW', body)
    worksheet162.write('H22', 'MAP', body)
    worksheet162.write('I22', 'IND', body)
    worksheet162.write('J22', 'ENG', body)
    worksheet162.write('J22', 'SEJ', body)
    worksheet162.write('K22', 'GEO', body)
    worksheet162.write('M22', 'EKO', body)
    worksheet162.write('L22', 'SOS', body)
    worksheet162.write('L22', 'FIS', body)
    worksheet162.write('L22', 'KIM', body)
    worksheet162.write('L22', 'BIO', body)
    worksheet162.write('N22', 'JML', body)
    worksheet162.write('O22', 'MAW', body)
    worksheet162.write('O22', 'MAP', body)
    worksheet162.write('P22', 'IND', body)
    worksheet162.write('Q22', 'ENG', body)
    worksheet162.write('R22', 'SEJ', body)
    worksheet162.write('S22', 'GEO', body)
    worksheet162.write('U22', 'EKO', body)
    worksheet162.write('T22', 'SOS', body)
    worksheet162.write('T22', 'FIS', body)
    worksheet162.write('T22', 'KIM', body)
    worksheet162.write('T22', 'BIO', body)
    worksheet162.write('V22', 'JML', body)

    worksheet162.conditional_format(22, 0, row162+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 163
    worksheet163.insert_image('A1', r'logo resmi nf.jpg')

    worksheet163.set_column('A:A', 7, center)
    worksheet163.set_column('B:B', 6, center)
    worksheet163.set_column('C:C', 18.14, center)
    worksheet163.set_column('D:D', 25, left)
    worksheet163.set_column('E:E', 13.14, left)
    worksheet163.set_column('F:F', 8.57, center)
    worksheet163.set_column('G:V', 5, center)
    worksheet163.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CEMPAKA', title)
    worksheet163.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet163.write('A5', 'LOKASI', header)
    worksheet163.write('B5', 'TOTAL', header)
    worksheet163.merge_range('A4:B4', 'RANK', header)
    worksheet163.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet163.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet163.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet163.merge_range('F4:F5', 'KELAS', header)
    worksheet163.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet163.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet163.write('G5', 'MAW', body)
    worksheet163.write('H5', 'MAP', body)
    worksheet163.write('I5', 'IND', body)
    worksheet163.write('J5', 'ENG', body)
    worksheet163.write('K5', 'SEJ', body)
    worksheet163.write('L5', 'GEO', body)
    worksheet163.write('M5', 'EKO', body)
    worksheet163.write('N5', 'SOS', body)
    worksheet163.write('O5', 'FIS', body)
    worksheet163.write('P5', 'KIM', body)
    worksheet163.write('Q5', 'BIO', body)
    worksheet163.write('R5', 'JML', body)
    worksheet163.write('S5', 'MAW', body)
    worksheet163.write('T5', 'MAP', body)
    worksheet163.write('U5', 'IND', body)
    worksheet163.write('V5', 'ENG', body)
    worksheet163.write('W5', 'SEJ', body)
    worksheet163.write('X5', 'GEO', body)
    worksheet163.write('Y5', 'EKO', body)
    worksheet163.write('Z5', 'SOS', body)
    worksheet163.write('AA5', 'FIS', body)
    worksheet163.write('AB5', 'KIM', body)
    worksheet163.write('AC5', 'BIO', body)
    worksheet163.write('AD5', 'JML', body)

    worksheet163.conditional_format(5, 0, row163_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet163.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF CEMPAKA', title)
    worksheet163.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet163.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet163.write('A22', 'LOKASI', header)
    worksheet163.write('B22', 'TOTAL', header)
    worksheet163.merge_range('A21:B21', 'RANK', header)
    worksheet163.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet163.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet163.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet163.merge_range('F21:F22', 'KELAS', header)
    worksheet163.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet163.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet163.write('G22', 'MAW', body)
    worksheet163.write('H22', 'MAP', body)
    worksheet163.write('I22', 'IND', body)
    worksheet163.write('J22', 'ENG', body)
    worksheet163.write('J22', 'SEJ', body)
    worksheet163.write('K22', 'GEO', body)
    worksheet163.write('M22', 'EKO', body)
    worksheet163.write('L22', 'SOS', body)
    worksheet163.write('L22', 'FIS', body)
    worksheet163.write('L22', 'KIM', body)
    worksheet163.write('L22', 'BIO', body)
    worksheet163.write('N22', 'JML', body)
    worksheet163.write('O22', 'MAW', body)
    worksheet163.write('O22', 'MAP', body)
    worksheet163.write('P22', 'IND', body)
    worksheet163.write('Q22', 'ENG', body)
    worksheet163.write('R22', 'SEJ', body)
    worksheet163.write('S22', 'GEO', body)
    worksheet163.write('U22', 'EKO', body)
    worksheet163.write('T22', 'SOS', body)
    worksheet163.write('T22', 'FIS', body)
    worksheet163.write('T22', 'KIM', body)
    worksheet163.write('T22', 'BIO', body)
    worksheet163.write('V22', 'JML', body)

    worksheet163.conditional_format(22, 0, row163+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 164
    worksheet164.insert_image('A1', r'logo resmi nf.jpg')

    worksheet164.set_column('A:A', 7, center)
    worksheet164.set_column('B:B', 6, center)
    worksheet164.set_column('C:C', 18.14, center)
    worksheet164.set_column('D:D', 25, left)
    worksheet164.set_column('E:E', 13.14, left)
    worksheet164.set_column('F:F', 8.57, center)
    worksheet164.set_column('G:V', 5, center)
    worksheet164.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PANGKALAN JATI', title)
    worksheet164.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet164.write('A5', 'LOKASI', header)
    worksheet164.write('B5', 'TOTAL', header)
    worksheet164.merge_range('A4:B4', 'RANK', header)
    worksheet164.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet164.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet164.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet164.merge_range('F4:F5', 'KELAS', header)
    worksheet164.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet164.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet164.write('G5', 'MAW', body)
    worksheet164.write('H5', 'MAP', body)
    worksheet164.write('I5', 'IND', body)
    worksheet164.write('J5', 'ENG', body)
    worksheet164.write('K5', 'SEJ', body)
    worksheet164.write('L5', 'GEO', body)
    worksheet164.write('M5', 'EKO', body)
    worksheet164.write('N5', 'SOS', body)
    worksheet164.write('O5', 'FIS', body)
    worksheet164.write('P5', 'KIM', body)
    worksheet164.write('Q5', 'BIO', body)
    worksheet164.write('R5', 'JML', body)
    worksheet164.write('S5', 'MAW', body)
    worksheet164.write('T5', 'MAP', body)
    worksheet164.write('U5', 'IND', body)
    worksheet164.write('V5', 'ENG', body)
    worksheet164.write('W5', 'SEJ', body)
    worksheet164.write('X5', 'GEO', body)
    worksheet164.write('Y5', 'EKO', body)
    worksheet164.write('Z5', 'SOS', body)
    worksheet164.write('AA5', 'FIS', body)
    worksheet164.write('AB5', 'KIM', body)
    worksheet164.write('AC5', 'BIO', body)
    worksheet164.write('AD5', 'JML', body)

    worksheet164.conditional_format(5, 0, row164_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet164.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF PANGKALAN JATI', title)
    worksheet164.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet164.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet164.write('A22', 'LOKASI', header)
    worksheet164.write('B22', 'TOTAL', header)
    worksheet164.merge_range('A21:B21', 'RANK', header)
    worksheet164.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet164.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet164.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet164.merge_range('F21:F22', 'KELAS', header)
    worksheet164.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet164.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet164.write('G22', 'MAW', body)
    worksheet164.write('H22', 'MAP', body)
    worksheet164.write('I22', 'IND', body)
    worksheet164.write('J22', 'ENG', body)
    worksheet164.write('J22', 'SEJ', body)
    worksheet164.write('K22', 'GEO', body)
    worksheet164.write('M22', 'EKO', body)
    worksheet164.write('L22', 'SOS', body)
    worksheet164.write('L22', 'FIS', body)
    worksheet164.write('L22', 'KIM', body)
    worksheet164.write('L22', 'BIO', body)
    worksheet164.write('N22', 'JML', body)
    worksheet164.write('O22', 'MAW', body)
    worksheet164.write('O22', 'MAP', body)
    worksheet164.write('P22', 'IND', body)
    worksheet164.write('Q22', 'ENG', body)
    worksheet164.write('R22', 'SEJ', body)
    worksheet164.write('S22', 'GEO', body)
    worksheet164.write('U22', 'EKO', body)
    worksheet164.write('T22', 'SOS', body)
    worksheet164.write('T22', 'FIS', body)
    worksheet164.write('T22', 'KIM', body)
    worksheet164.write('T22', 'BIO', body)
    worksheet164.write('V22', 'JML', body)

    worksheet164.conditional_format(22, 0, row164+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # # worksheet 104
    # worksheet104.insert_image('A1',r'logo resmi nf.jpg')

    # worksheet104.set_column('A:A', 7, center)
    # worksheet104.set_column('B:B', 6, center)
    # worksheet104.set_column('C:C', 18.14, center)
    # worksheet104.set_column('D:D', 25, left)
    # worksheet104.set_column('E:E', 13.14, left)
    # worksheet104.set_column('F:F', 8.57, center)
    # worksheet104.set_column('G:V', 5, center)
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

    # worksheet 165
    worksheet165.insert_image('A1', r'logo resmi nf.jpg')

    worksheet165.set_column('A:A', 7, center)
    worksheet165.set_column('B:B', 6, center)
    worksheet165.set_column('C:C', 18.14, center)
    worksheet165.set_column('D:D', 25, left)
    worksheet165.set_column('E:E', 13.14, left)
    worksheet165.set_column('F:F', 8.57, center)
    worksheet165.set_column('G:V', 5, center)
    worksheet165.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BUARAN', title)
    worksheet165.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet165.write('A5', 'LOKASI', header)
    worksheet165.write('B5', 'TOTAL', header)
    worksheet165.merge_range('A4:B4', 'RANK', header)
    worksheet165.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet165.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet165.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet165.merge_range('F4:F5', 'KELAS', header)
    worksheet165.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet165.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet165.write('G5', 'MAW', body)
    worksheet165.write('H5', 'MAP', body)
    worksheet165.write('I5', 'IND', body)
    worksheet165.write('J5', 'ENG', body)
    worksheet165.write('K5', 'SEJ', body)
    worksheet165.write('L5', 'GEO', body)
    worksheet165.write('M5', 'EKO', body)
    worksheet165.write('N5', 'SOS', body)
    worksheet165.write('O5', 'FIS', body)
    worksheet165.write('P5', 'KIM', body)
    worksheet165.write('Q5', 'BIO', body)
    worksheet165.write('R5', 'JML', body)
    worksheet165.write('S5', 'MAW', body)
    worksheet165.write('T5', 'MAP', body)
    worksheet165.write('U5', 'IND', body)
    worksheet165.write('V5', 'ENG', body)
    worksheet165.write('W5', 'SEJ', body)
    worksheet165.write('X5', 'GEO', body)
    worksheet165.write('Y5', 'EKO', body)
    worksheet165.write('Z5', 'SOS', body)
    worksheet165.write('AA5', 'FIS', body)
    worksheet165.write('AB5', 'KIM', body)
    worksheet165.write('AC5', 'BIO', body)
    worksheet165.write('AD5', 'JML', body)

    worksheet165.conditional_format(5, 0, row165_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet165.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF BUARAN', title)
    worksheet165.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet165.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet165.write('A22', 'LOKASI', header)
    worksheet165.write('B22', 'TOTAL', header)
    worksheet165.merge_range('A21:B21', 'RANK', header)
    worksheet165.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet165.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet165.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet165.merge_range('F21:F22', 'KELAS', header)
    worksheet165.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet165.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet165.write('G22', 'MAW', body)
    worksheet165.write('H22', 'MAP', body)
    worksheet165.write('I22', 'IND', body)
    worksheet165.write('J22', 'ENG', body)
    worksheet165.write('J22', 'SEJ', body)
    worksheet165.write('K22', 'GEO', body)
    worksheet165.write('M22', 'EKO', body)
    worksheet165.write('L22', 'SOS', body)
    worksheet165.write('L22', 'FIS', body)
    worksheet165.write('L22', 'KIM', body)
    worksheet165.write('L22', 'BIO', body)
    worksheet165.write('N22', 'JML', body)
    worksheet165.write('O22', 'MAW', body)
    worksheet165.write('O22', 'MAP', body)
    worksheet165.write('P22', 'IND', body)
    worksheet165.write('Q22', 'ENG', body)
    worksheet165.write('R22', 'SEJ', body)
    worksheet165.write('S22', 'GEO', body)
    worksheet165.write('U22', 'EKO', body)
    worksheet165.write('T22', 'SOS', body)
    worksheet165.write('T22', 'FIS', body)
    worksheet165.write('T22', 'KIM', body)
    worksheet165.write('T22', 'BIO', body)
    worksheet165.write('V22', 'JML', body)

    worksheet165.conditional_format(22, 0, row165+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 167
    worksheet167.insert_image('A1', r'logo resmi nf.jpg')

    worksheet167.set_column('A:A', 7, center)
    worksheet167.set_column('B:B', 6, center)
    worksheet167.set_column('C:C', 18.14, center)
    worksheet167.set_column('D:D', 25, left)
    worksheet167.set_column('E:E', 13.14, left)
    worksheet167.set_column('F:F', 8.57, center)
    worksheet167.set_column('G:V', 5, center)
    worksheet167.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF HEK-KRAMAT JATI', title)
    worksheet167.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet167.write('A5', 'LOKASI', header)
    worksheet167.write('B5', 'TOTAL', header)
    worksheet167.merge_range('A4:B4', 'RANK', header)
    worksheet167.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet167.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet167.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet167.merge_range('F4:F5', 'KELAS', header)
    worksheet167.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet167.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet167.write('G5', 'MAW', body)
    worksheet167.write('H5', 'MAP', body)
    worksheet167.write('I5', 'IND', body)
    worksheet167.write('J5', 'ENG', body)
    worksheet167.write('K5', 'SEJ', body)
    worksheet167.write('L5', 'GEO', body)
    worksheet167.write('M5', 'EKO', body)
    worksheet167.write('N5', 'SOS', body)
    worksheet167.write('O5', 'FIS', body)
    worksheet167.write('P5', 'KIM', body)
    worksheet167.write('Q5', 'BIO', body)
    worksheet167.write('R5', 'JML', body)
    worksheet167.write('S5', 'MAW', body)
    worksheet167.write('T5', 'MAP', body)
    worksheet167.write('U5', 'IND', body)
    worksheet167.write('V5', 'ENG', body)
    worksheet167.write('W5', 'SEJ', body)
    worksheet167.write('X5', 'GEO', body)
    worksheet167.write('Y5', 'EKO', body)
    worksheet167.write('Z5', 'SOS', body)
    worksheet167.write('AA5', 'FIS', body)
    worksheet167.write('AB5', 'KIM', body)
    worksheet167.write('AC5', 'BIO', body)
    worksheet167.write('AD5', 'JML', body)

    worksheet167.conditional_format(5, 0, row167_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet167.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF HEK-KRAMAT JATI', title)
    worksheet167.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet167.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet167.write('A22', 'LOKASI', header)
    worksheet167.write('B22', 'TOTAL', header)
    worksheet167.merge_range('A21:B21', 'RANK', header)
    worksheet167.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet167.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet167.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet167.merge_range('F21:F22', 'KELAS', header)
    worksheet167.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet167.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet167.write('G22', 'MAW', body)
    worksheet167.write('H22', 'MAP', body)
    worksheet167.write('I22', 'IND', body)
    worksheet167.write('J22', 'ENG', body)
    worksheet167.write('J22', 'SEJ', body)
    worksheet167.write('K22', 'GEO', body)
    worksheet167.write('M22', 'EKO', body)
    worksheet167.write('L22', 'SOS', body)
    worksheet167.write('L22', 'FIS', body)
    worksheet167.write('L22', 'KIM', body)
    worksheet167.write('L22', 'BIO', body)
    worksheet167.write('N22', 'JML', body)
    worksheet167.write('O22', 'MAW', body)
    worksheet167.write('O22', 'MAP', body)
    worksheet167.write('P22', 'IND', body)
    worksheet167.write('Q22', 'ENG', body)
    worksheet167.write('R22', 'SEJ', body)
    worksheet167.write('S22', 'GEO', body)
    worksheet167.write('U22', 'EKO', body)
    worksheet167.write('T22', 'SOS', body)
    worksheet167.write('T22', 'FIS', body)
    worksheet167.write('T22', 'KIM', body)
    worksheet167.write('T22', 'BIO', body)
    worksheet167.write('V22', 'JML', body)

    worksheet167.conditional_format(22, 0, row167+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 168
    worksheet168.insert_image('A1', r'logo resmi nf.jpg')

    worksheet168.set_column('A:A', 7, center)
    worksheet168.set_column('B:B', 6, center)
    worksheet168.set_column('C:C', 18.14, center)
    worksheet168.set_column('D:D', 25, left)
    worksheet168.set_column('E:E', 13.14, left)
    worksheet168.set_column('F:F', 8.57, center)
    worksheet168.set_column('G:V', 5, center)
    worksheet168.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MAMPANG', title)
    worksheet168.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet168.write('A5', 'LOKASI', header)
    worksheet168.write('B5', 'TOTAL', header)
    worksheet168.merge_range('A4:B4', 'RANK', header)
    worksheet168.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet168.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet168.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet168.merge_range('F4:F5', 'KELAS', header)
    worksheet168.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet168.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet168.write('G5', 'MAW', body)
    worksheet168.write('H5', 'MAP', body)
    worksheet168.write('I5', 'IND', body)
    worksheet168.write('J5', 'ENG', body)
    worksheet168.write('K5', 'SEJ', body)
    worksheet168.write('L5', 'GEO', body)
    worksheet168.write('M5', 'EKO', body)
    worksheet168.write('N5', 'SOS', body)
    worksheet168.write('O5', 'FIS', body)
    worksheet168.write('P5', 'KIM', body)
    worksheet168.write('Q5', 'BIO', body)
    worksheet168.write('R5', 'JML', body)
    worksheet168.write('S5', 'MAW', body)
    worksheet168.write('T5', 'MAP', body)
    worksheet168.write('U5', 'IND', body)
    worksheet168.write('V5', 'ENG', body)
    worksheet168.write('W5', 'SEJ', body)
    worksheet168.write('X5', 'GEO', body)
    worksheet168.write('Y5', 'EKO', body)
    worksheet168.write('Z5', 'SOS', body)
    worksheet168.write('AA5', 'FIS', body)
    worksheet168.write('AB5', 'KIM', body)
    worksheet168.write('AC5', 'BIO', body)
    worksheet168.write('AD5', 'JML', body)

    worksheet168.conditional_format(5, 0, row168_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet168.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF MAMPANG', title)
    worksheet168.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet168.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet168.write('A22', 'LOKASI', header)
    worksheet168.write('B22', 'TOTAL', header)
    worksheet168.merge_range('A21:B21', 'RANK', header)
    worksheet168.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet168.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet168.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet168.merge_range('F21:F22', 'KELAS', header)
    worksheet168.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet168.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet168.write('G22', 'MAW', body)
    worksheet168.write('H22', 'MAP', body)
    worksheet168.write('I22', 'IND', body)
    worksheet168.write('J22', 'ENG', body)
    worksheet168.write('J22', 'SEJ', body)
    worksheet168.write('K22', 'GEO', body)
    worksheet168.write('M22', 'EKO', body)
    worksheet168.write('L22', 'SOS', body)
    worksheet168.write('L22', 'FIS', body)
    worksheet168.write('L22', 'KIM', body)
    worksheet168.write('L22', 'BIO', body)
    worksheet168.write('N22', 'JML', body)
    worksheet168.write('O22', 'MAW', body)
    worksheet168.write('O22', 'MAP', body)
    worksheet168.write('P22', 'IND', body)
    worksheet168.write('Q22', 'ENG', body)
    worksheet168.write('R22', 'SEJ', body)
    worksheet168.write('S22', 'GEO', body)
    worksheet168.write('U22', 'EKO', body)
    worksheet168.write('T22', 'SOS', body)
    worksheet168.write('T22', 'FIS', body)
    worksheet168.write('T22', 'KIM', body)
    worksheet168.write('T22', 'BIO', body)
    worksheet168.write('V22', 'JML', body)

    worksheet168.conditional_format(22, 0, row168+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 169
    worksheet169.insert_image('A1', r'logo resmi nf.jpg')

    worksheet169.set_column('A:A', 7, center)
    worksheet169.set_column('B:B', 6, center)
    worksheet169.set_column('C:C', 18.14, center)
    worksheet169.set_column('D:D', 25, left)
    worksheet169.set_column('E:E', 13.14, left)
    worksheet169.set_column('F:F', 8.57, center)
    worksheet169.set_column('G:V', 5, center)
    worksheet169.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PALMERAH', title)
    worksheet169.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet169.write('A5', 'LOKASI', header)
    worksheet169.write('B5', 'TOTAL', header)
    worksheet169.merge_range('A4:B4', 'RANK', header)
    worksheet169.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet169.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet169.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet169.merge_range('F4:F5', 'KELAS', header)
    worksheet169.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet169.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet169.write('G5', 'MAW', body)
    worksheet169.write('H5', 'MAP', body)
    worksheet169.write('I5', 'IND', body)
    worksheet169.write('J5', 'ENG', body)
    worksheet169.write('K5', 'SEJ', body)
    worksheet169.write('L5', 'GEO', body)
    worksheet169.write('M5', 'EKO', body)
    worksheet169.write('N5', 'SOS', body)
    worksheet169.write('O5', 'FIS', body)
    worksheet169.write('P5', 'KIM', body)
    worksheet169.write('Q5', 'BIO', body)
    worksheet169.write('R5', 'JML', body)
    worksheet169.write('S5', 'MAW', body)
    worksheet169.write('T5', 'MAP', body)
    worksheet169.write('U5', 'IND', body)
    worksheet169.write('V5', 'ENG', body)
    worksheet169.write('W5', 'SEJ', body)
    worksheet169.write('X5', 'GEO', body)
    worksheet169.write('Y5', 'EKO', body)
    worksheet169.write('Z5', 'SOS', body)
    worksheet169.write('AA5', 'FIS', body)
    worksheet169.write('AB5', 'KIM', body)
    worksheet169.write('AC5', 'BIO', body)
    worksheet169.write('AD5', 'JML', body)

    worksheet169.conditional_format(5, 0, row169_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet169.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF PALMERAH', title)
    worksheet169.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet169.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet169.write('A22', 'LOKASI', header)
    worksheet169.write('B22', 'TOTAL', header)
    worksheet169.merge_range('A21:B21', 'RANK', header)
    worksheet169.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet169.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet169.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet169.merge_range('F21:F22', 'KELAS', header)
    worksheet169.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet169.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet169.write('G22', 'MAW', body)
    worksheet169.write('H22', 'MAP', body)
    worksheet169.write('I22', 'IND', body)
    worksheet169.write('J22', 'ENG', body)
    worksheet169.write('J22', 'SEJ', body)
    worksheet169.write('K22', 'GEO', body)
    worksheet169.write('M22', 'EKO', body)
    worksheet169.write('L22', 'SOS', body)
    worksheet169.write('L22', 'FIS', body)
    worksheet169.write('L22', 'KIM', body)
    worksheet169.write('L22', 'BIO', body)
    worksheet169.write('N22', 'JML', body)
    worksheet169.write('O22', 'MAW', body)
    worksheet169.write('O22', 'MAP', body)
    worksheet169.write('P22', 'IND', body)
    worksheet169.write('Q22', 'ENG', body)
    worksheet169.write('R22', 'SEJ', body)
    worksheet169.write('S22', 'GEO', body)
    worksheet169.write('U22', 'EKO', body)
    worksheet169.write('T22', 'SOS', body)
    worksheet169.write('T22', 'FIS', body)
    worksheet169.write('T22', 'KIM', body)
    worksheet169.write('T22', 'BIO', body)
    worksheet169.write('V22', 'JML', body)

    worksheet169.conditional_format(22, 0, row169+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 171
    worksheet171.insert_image('A1', r'logo resmi nf.jpg')

    worksheet171.set_column('A:A', 7, center)
    worksheet171.set_column('B:B', 6, center)
    worksheet171.set_column('C:C', 18.14, center)
    worksheet171.set_column('D:D', 25, left)
    worksheet171.set_column('E:E', 13.14, left)
    worksheet171.set_column('F:F', 8.57, center)
    worksheet171.set_column('G:V', 5, center)
    worksheet171.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PASAR MINGGU', title)
    worksheet171.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet171.write('A5', 'LOKASI', header)
    worksheet171.write('B5', 'TOTAL', header)
    worksheet171.merge_range('A4:B4', 'RANK', header)
    worksheet171.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet171.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet171.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet171.merge_range('F4:F5', 'KELAS', header)
    worksheet171.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet171.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet171.write('G5', 'MAW', body)
    worksheet171.write('H5', 'MAP', body)
    worksheet171.write('I5', 'IND', body)
    worksheet171.write('J5', 'ENG', body)
    worksheet171.write('K5', 'SEJ', body)
    worksheet171.write('L5', 'GEO', body)
    worksheet171.write('M5', 'EKO', body)
    worksheet171.write('N5', 'SOS', body)
    worksheet171.write('O5', 'FIS', body)
    worksheet171.write('P5', 'KIM', body)
    worksheet171.write('Q5', 'BIO', body)
    worksheet171.write('R5', 'JML', body)
    worksheet171.write('S5', 'MAW', body)
    worksheet171.write('T5', 'MAP', body)
    worksheet171.write('U5', 'IND', body)
    worksheet171.write('V5', 'ENG', body)
    worksheet171.write('W5', 'SEJ', body)
    worksheet171.write('X5', 'GEO', body)
    worksheet171.write('Y5', 'EKO', body)
    worksheet171.write('Z5', 'SOS', body)
    worksheet171.write('AA5', 'FIS', body)
    worksheet171.write('AB5', 'KIM', body)
    worksheet171.write('AC5', 'BIO', body)
    worksheet171.write('AD5', 'JML', body)

    worksheet171.conditional_format(5, 0, row171_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet171.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF PASAR MINGGU', title)
    worksheet171.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet171.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet171.write('A22', 'LOKASI', header)
    worksheet171.write('B22', 'TOTAL', header)
    worksheet171.merge_range('A21:B21', 'RANK', header)
    worksheet171.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet171.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet171.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet171.merge_range('F21:F22', 'KELAS', header)
    worksheet171.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet171.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet171.write('G22', 'MAW', body)
    worksheet171.write('H22', 'MAP', body)
    worksheet171.write('I22', 'IND', body)
    worksheet171.write('J22', 'ENG', body)
    worksheet171.write('J22', 'SEJ', body)
    worksheet171.write('K22', 'GEO', body)
    worksheet171.write('M22', 'EKO', body)
    worksheet171.write('L22', 'SOS', body)
    worksheet171.write('L22', 'FIS', body)
    worksheet171.write('L22', 'KIM', body)
    worksheet171.write('L22', 'BIO', body)
    worksheet171.write('N22', 'JML', body)
    worksheet171.write('O22', 'MAW', body)
    worksheet171.write('O22', 'MAP', body)
    worksheet171.write('P22', 'IND', body)
    worksheet171.write('Q22', 'ENG', body)
    worksheet171.write('R22', 'SEJ', body)
    worksheet171.write('S22', 'GEO', body)
    worksheet171.write('U22', 'EKO', body)
    worksheet171.write('T22', 'SOS', body)
    worksheet171.write('T22', 'FIS', body)
    worksheet171.write('T22', 'KIM', body)
    worksheet171.write('T22', 'BIO', body)
    worksheet171.write('V22', 'JML', body)

    worksheet171.conditional_format(22, 0, row171+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 173
    worksheet173.insert_image('A1', r'logo resmi nf.jpg')

    worksheet173.set_column('A:A', 7, center)
    worksheet173.set_column('B:B', 6, center)
    worksheet173.set_column('C:C', 18.14, center)
    worksheet173.set_column('D:D', 25, left)
    worksheet173.set_column('E:E', 13.14, left)
    worksheet173.set_column('F:F', 8.57, center)
    worksheet173.set_column('G:V', 5, center)
    worksheet173.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BINTARO', title)
    worksheet173.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet173.write('A5', 'LOKASI', header)
    worksheet173.write('B5', 'TOTAL', header)
    worksheet173.merge_range('A4:B4', 'RANK', header)
    worksheet173.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet173.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet173.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet173.merge_range('F4:F5', 'KELAS', header)
    worksheet173.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet173.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet173.write('G5', 'MAW', body)
    worksheet173.write('H5', 'MAP', body)
    worksheet173.write('I5', 'IND', body)
    worksheet173.write('J5', 'ENG', body)
    worksheet173.write('K5', 'SEJ', body)
    worksheet173.write('L5', 'GEO', body)
    worksheet173.write('M5', 'EKO', body)
    worksheet173.write('N5', 'SOS', body)
    worksheet173.write('O5', 'FIS', body)
    worksheet173.write('P5', 'KIM', body)
    worksheet173.write('Q5', 'BIO', body)
    worksheet173.write('R5', 'JML', body)
    worksheet173.write('S5', 'MAW', body)
    worksheet173.write('T5', 'MAP', body)
    worksheet173.write('U5', 'IND', body)
    worksheet173.write('V5', 'ENG', body)
    worksheet173.write('W5', 'SEJ', body)
    worksheet173.write('X5', 'GEO', body)
    worksheet173.write('Y5', 'EKO', body)
    worksheet173.write('Z5', 'SOS', body)
    worksheet173.write('AA5', 'FIS', body)
    worksheet173.write('AB5', 'KIM', body)
    worksheet173.write('AC5', 'BIO', body)
    worksheet173.write('AD5', 'JML', body)

    worksheet173.conditional_format(5, 0, row173_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet173.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF BINTARO', title)
    worksheet173.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet173.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet173.write('A22', 'LOKASI', header)
    worksheet173.write('B22', 'TOTAL', header)
    worksheet173.merge_range('A21:B21', 'RANK', header)
    worksheet173.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet173.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet173.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet173.merge_range('F21:F22', 'KELAS', header)
    worksheet173.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet173.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet173.write('G22', 'MAW', body)
    worksheet173.write('H22', 'MAP', body)
    worksheet173.write('I22', 'IND', body)
    worksheet173.write('J22', 'ENG', body)
    worksheet173.write('J22', 'SEJ', body)
    worksheet173.write('K22', 'GEO', body)
    worksheet173.write('M22', 'EKO', body)
    worksheet173.write('L22', 'SOS', body)
    worksheet173.write('L22', 'FIS', body)
    worksheet173.write('L22', 'KIM', body)
    worksheet173.write('L22', 'BIO', body)
    worksheet173.write('N22', 'JML', body)
    worksheet173.write('O22', 'MAW', body)
    worksheet173.write('O22', 'MAP', body)
    worksheet173.write('P22', 'IND', body)
    worksheet173.write('Q22', 'ENG', body)
    worksheet173.write('R22', 'SEJ', body)
    worksheet173.write('S22', 'GEO', body)
    worksheet173.write('U22', 'EKO', body)
    worksheet173.write('T22', 'SOS', body)
    worksheet173.write('T22', 'FIS', body)
    worksheet173.write('T22', 'KIM', body)
    worksheet173.write('T22', 'BIO', body)
    worksheet173.write('V22', 'JML', body)

    worksheet173.conditional_format(22, 0, row173+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 174
    worksheet174.insert_image('A1', r'logo resmi nf.jpg')

    worksheet174.set_column('A:A', 7, center)
    worksheet174.set_column('B:B', 6, center)
    worksheet174.set_column('C:C', 18.14, center)
    worksheet174.set_column('D:D', 25, left)
    worksheet174.set_column('E:E', 13.14, left)
    worksheet174.set_column('F:F', 8.57, center)
    worksheet174.set_column('G:V', 5, center)
    worksheet174.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF LAMPIRI', title)
    worksheet174.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet174.write('A5', 'LOKASI', header)
    worksheet174.write('B5', 'TOTAL', header)
    worksheet174.merge_range('A4:B4', 'RANK', header)
    worksheet174.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet174.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet174.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet174.merge_range('F4:F5', 'KELAS', header)
    worksheet174.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet174.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet174.write('G5', 'MAW', body)
    worksheet174.write('H5', 'MAP', body)
    worksheet174.write('I5', 'IND', body)
    worksheet174.write('J5', 'ENG', body)
    worksheet174.write('K5', 'SEJ', body)
    worksheet174.write('L5', 'GEO', body)
    worksheet174.write('M5', 'EKO', body)
    worksheet174.write('N5', 'SOS', body)
    worksheet174.write('O5', 'FIS', body)
    worksheet174.write('P5', 'KIM', body)
    worksheet174.write('Q5', 'BIO', body)
    worksheet174.write('R5', 'JML', body)
    worksheet174.write('S5', 'MAW', body)
    worksheet174.write('T5', 'MAP', body)
    worksheet174.write('U5', 'IND', body)
    worksheet174.write('V5', 'ENG', body)
    worksheet174.write('W5', 'SEJ', body)
    worksheet174.write('X5', 'GEO', body)
    worksheet174.write('Y5', 'EKO', body)
    worksheet174.write('Z5', 'SOS', body)
    worksheet174.write('AA5', 'FIS', body)
    worksheet174.write('AB5', 'KIM', body)
    worksheet174.write('AC5', 'BIO', body)
    worksheet174.write('AD5', 'JML', body)

    worksheet174.conditional_format(5, 0, row174_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet174.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF LAMPIRI', title)
    worksheet174.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet174.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet174.write('A22', 'LOKASI', header)
    worksheet174.write('B22', 'TOTAL', header)
    worksheet174.merge_range('A21:B21', 'RANK', header)
    worksheet174.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet174.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet174.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet174.merge_range('F21:F22', 'KELAS', header)
    worksheet174.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet174.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet174.write('G22', 'MAW', body)
    worksheet174.write('H22', 'MAP', body)
    worksheet174.write('I22', 'IND', body)
    worksheet174.write('J22', 'ENG', body)
    worksheet174.write('J22', 'SEJ', body)
    worksheet174.write('K22', 'GEO', body)
    worksheet174.write('M22', 'EKO', body)
    worksheet174.write('L22', 'SOS', body)
    worksheet174.write('L22', 'FIS', body)
    worksheet174.write('L22', 'KIM', body)
    worksheet174.write('L22', 'BIO', body)
    worksheet174.write('N22', 'JML', body)
    worksheet174.write('O22', 'MAW', body)
    worksheet174.write('O22', 'MAP', body)
    worksheet174.write('P22', 'IND', body)
    worksheet174.write('Q22', 'ENG', body)
    worksheet174.write('R22', 'SEJ', body)
    worksheet174.write('S22', 'GEO', body)
    worksheet174.write('U22', 'EKO', body)
    worksheet174.write('T22', 'SOS', body)
    worksheet174.write('T22', 'FIS', body)
    worksheet174.write('T22', 'KIM', body)
    worksheet174.write('T22', 'BIO', body)
    worksheet174.write('V22', 'JML', body)

    worksheet174.conditional_format(22, 0, row174+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 175
    worksheet175.insert_image('A1', r'logo resmi nf.jpg')

    worksheet175.set_column('A:A', 7, center)
    worksheet175.set_column('B:B', 6, center)
    worksheet175.set_column('C:C', 18.14, center)
    worksheet175.set_column('D:D', 25, left)
    worksheet175.set_column('E:E', 13.14, left)
    worksheet175.set_column('F:F', 8.57, center)
    worksheet175.set_column('G:V', 5, center)
    worksheet175.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PONDOK BAMBU', title)
    worksheet175.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet175.write('A5', 'LOKASI', header)
    worksheet175.write('B5', 'TOTAL', header)
    worksheet175.merge_range('A4:B4', 'RANK', header)
    worksheet175.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet175.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet175.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet175.merge_range('F4:F5', 'KELAS', header)
    worksheet175.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet175.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet175.write('G5', 'MAW', body)
    worksheet175.write('H5', 'MAP', body)
    worksheet175.write('I5', 'IND', body)
    worksheet175.write('J5', 'ENG', body)
    worksheet175.write('K5', 'SEJ', body)
    worksheet175.write('L5', 'GEO', body)
    worksheet175.write('M5', 'EKO', body)
    worksheet175.write('N5', 'SOS', body)
    worksheet175.write('O5', 'FIS', body)
    worksheet175.write('P5', 'KIM', body)
    worksheet175.write('Q5', 'BIO', body)
    worksheet175.write('R5', 'JML', body)
    worksheet175.write('S5', 'MAW', body)
    worksheet175.write('T5', 'MAP', body)
    worksheet175.write('U5', 'IND', body)
    worksheet175.write('V5', 'ENG', body)
    worksheet175.write('W5', 'SEJ', body)
    worksheet175.write('X5', 'GEO', body)
    worksheet175.write('Y5', 'EKO', body)
    worksheet175.write('Z5', 'SOS', body)
    worksheet175.write('AA5', 'FIS', body)
    worksheet175.write('AB5', 'KIM', body)
    worksheet175.write('AC5', 'BIO', body)
    worksheet175.write('AD5', 'JML', body)

    worksheet175.conditional_format(5, 0, row175_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet175.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF PONDOK BAMBU', title)
    worksheet175.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet175.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet175.write('A22', 'LOKASI', header)
    worksheet175.write('B22', 'TOTAL', header)
    worksheet175.merge_range('A21:B21', 'RANK', header)
    worksheet175.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet175.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet175.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet175.merge_range('F21:F22', 'KELAS', header)
    worksheet175.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet175.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet175.write('G22', 'MAW', body)
    worksheet175.write('H22', 'MAP', body)
    worksheet175.write('I22', 'IND', body)
    worksheet175.write('J22', 'ENG', body)
    worksheet175.write('J22', 'SEJ', body)
    worksheet175.write('K22', 'GEO', body)
    worksheet175.write('M22', 'EKO', body)
    worksheet175.write('L22', 'SOS', body)
    worksheet175.write('L22', 'FIS', body)
    worksheet175.write('L22', 'KIM', body)
    worksheet175.write('L22', 'BIO', body)
    worksheet175.write('N22', 'JML', body)
    worksheet175.write('O22', 'MAW', body)
    worksheet175.write('O22', 'MAP', body)
    worksheet175.write('P22', 'IND', body)
    worksheet175.write('Q22', 'ENG', body)
    worksheet175.write('R22', 'SEJ', body)
    worksheet175.write('S22', 'GEO', body)
    worksheet175.write('U22', 'EKO', body)
    worksheet175.write('T22', 'SOS', body)
    worksheet175.write('T22', 'FIS', body)
    worksheet175.write('T22', 'KIM', body)
    worksheet175.write('T22', 'BIO', body)
    worksheet175.write('V22', 'JML', body)

    worksheet175.conditional_format(22, 0, row175+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 176
    worksheet176.insert_image('A1', r'logo resmi nf.jpg')

    worksheet176.set_column('A:A', 7, center)
    worksheet176.set_column('B:B', 6, center)
    worksheet176.set_column('C:C', 18.14, center)
    worksheet176.set_column('D:D', 25, left)
    worksheet176.set_column('E:E', 13.14, left)
    worksheet176.set_column('F:F', 8.57, center)
    worksheet176.set_column('G:V', 5, center)
    worksheet176.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF RAWA BADAK', title)
    worksheet176.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet176.write('A5', 'LOKASI', header)
    worksheet176.write('B5', 'TOTAL', header)
    worksheet176.merge_range('A4:B4', 'RANK', header)
    worksheet176.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet176.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet176.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet176.merge_range('F4:F5', 'KELAS', header)
    worksheet176.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet176.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet176.write('G5', 'MAW', body)
    worksheet176.write('H5', 'MAP', body)
    worksheet176.write('I5', 'IND', body)
    worksheet176.write('J5', 'ENG', body)
    worksheet176.write('K5', 'SEJ', body)
    worksheet176.write('L5', 'GEO', body)
    worksheet176.write('M5', 'EKO', body)
    worksheet176.write('N5', 'SOS', body)
    worksheet176.write('O5', 'FIS', body)
    worksheet176.write('P5', 'KIM', body)
    worksheet176.write('Q5', 'BIO', body)
    worksheet176.write('R5', 'JML', body)
    worksheet176.write('S5', 'MAW', body)
    worksheet176.write('T5', 'MAP', body)
    worksheet176.write('U5', 'IND', body)
    worksheet176.write('V5', 'ENG', body)
    worksheet176.write('W5', 'SEJ', body)
    worksheet176.write('X5', 'GEO', body)
    worksheet176.write('Y5', 'EKO', body)
    worksheet176.write('Z5', 'SOS', body)
    worksheet176.write('AA5', 'FIS', body)
    worksheet176.write('AB5', 'KIM', body)
    worksheet176.write('AC5', 'BIO', body)
    worksheet176.write('AD5', 'JML', body)

    worksheet176.conditional_format(5, 0, row176_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet176.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF RAWA BADAK', title)
    worksheet176.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet176.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet176.write('A22', 'LOKASI', header)
    worksheet176.write('B22', 'TOTAL', header)
    worksheet176.merge_range('A21:B21', 'RANK', header)
    worksheet176.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet176.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet176.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet176.merge_range('F21:F22', 'KELAS', header)
    worksheet176.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet176.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet176.write('G22', 'MAW', body)
    worksheet176.write('H22', 'MAP', body)
    worksheet176.write('I22', 'IND', body)
    worksheet176.write('J22', 'ENG', body)
    worksheet176.write('J22', 'SEJ', body)
    worksheet176.write('K22', 'GEO', body)
    worksheet176.write('M22', 'EKO', body)
    worksheet176.write('L22', 'SOS', body)
    worksheet176.write('L22', 'FIS', body)
    worksheet176.write('L22', 'KIM', body)
    worksheet176.write('L22', 'BIO', body)
    worksheet176.write('N22', 'JML', body)
    worksheet176.write('O22', 'MAW', body)
    worksheet176.write('O22', 'MAP', body)
    worksheet176.write('P22', 'IND', body)
    worksheet176.write('Q22', 'ENG', body)
    worksheet176.write('R22', 'SEJ', body)
    worksheet176.write('S22', 'GEO', body)
    worksheet176.write('U22', 'EKO', body)
    worksheet176.write('T22', 'SOS', body)
    worksheet176.write('T22', 'FIS', body)
    worksheet176.write('T22', 'KIM', body)
    worksheet176.write('T22', 'BIO', body)
    worksheet176.write('V22', 'JML', body)

    worksheet176.conditional_format(22, 0, row176+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # # worksheet 114
    # worksheet114.insert_image('A1',r'logo resmi nf.jpg')

    # worksheet114.set_column('A:A', 7, center)
    # worksheet114.set_column('B:B', 6, center)
    # worksheet114.set_column('C:C', 18.14, center)
    # worksheet114.set_column('D:D', 25, left)
    # worksheet114.set_column('E:E', 13.14, left)
    # worksheet114.set_column('F:F', 8.57, center)
    # worksheet114.set_column('G:V', 5, center)
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
    # worksheet 177
    worksheet177.insert_image('A1', r'logo resmi nf.jpg')

    worksheet177.set_column('A:A', 7, center)
    worksheet177.set_column('B:B', 6, center)
    worksheet177.set_column('C:C', 18.14, center)
    worksheet177.set_column('D:D', 25, left)
    worksheet177.set_column('E:E', 13.14, left)
    worksheet177.set_column('F:F', 8.57, center)
    worksheet177.set_column('G:V', 5, center)
    worksheet177.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF RAWAMANGUN', title)
    worksheet177.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet177.write('A5', 'LOKASI', header)
    worksheet177.write('B5', 'TOTAL', header)
    worksheet177.merge_range('A4:B4', 'RANK', header)
    worksheet177.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet177.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet177.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet177.merge_range('F4:F5', 'KELAS', header)
    worksheet177.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet177.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet177.write('G5', 'MAW', body)
    worksheet177.write('H5', 'MAP', body)
    worksheet177.write('I5', 'IND', body)
    worksheet177.write('J5', 'ENG', body)
    worksheet177.write('K5', 'SEJ', body)
    worksheet177.write('L5', 'GEO', body)
    worksheet177.write('M5', 'EKO', body)
    worksheet177.write('N5', 'SOS', body)
    worksheet177.write('O5', 'FIS', body)
    worksheet177.write('P5', 'KIM', body)
    worksheet177.write('Q5', 'BIO', body)
    worksheet177.write('R5', 'JML', body)
    worksheet177.write('S5', 'MAW', body)
    worksheet177.write('T5', 'MAP', body)
    worksheet177.write('U5', 'IND', body)
    worksheet177.write('V5', 'ENG', body)
    worksheet177.write('W5', 'SEJ', body)
    worksheet177.write('X5', 'GEO', body)
    worksheet177.write('Y5', 'EKO', body)
    worksheet177.write('Z5', 'SOS', body)
    worksheet177.write('AA5', 'FIS', body)
    worksheet177.write('AB5', 'KIM', body)
    worksheet177.write('AC5', 'BIO', body)
    worksheet177.write('AD5', 'JML', body)

    worksheet177.conditional_format(5, 0, row177_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet177.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF RAWAMANGUN', title)
    worksheet177.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet177.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet177.write('A22', 'LOKASI', header)
    worksheet177.write('B22', 'TOTAL', header)
    worksheet177.merge_range('A21:B21', 'RANK', header)
    worksheet177.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet177.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet177.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet177.merge_range('F21:F22', 'KELAS', header)
    worksheet177.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet177.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet177.write('G22', 'MAW', body)
    worksheet177.write('H22', 'MAP', body)
    worksheet177.write('I22', 'IND', body)
    worksheet177.write('J22', 'ENG', body)
    worksheet177.write('J22', 'SEJ', body)
    worksheet177.write('K22', 'GEO', body)
    worksheet177.write('M22', 'EKO', body)
    worksheet177.write('L22', 'SOS', body)
    worksheet177.write('L22', 'FIS', body)
    worksheet177.write('L22', 'KIM', body)
    worksheet177.write('L22', 'BIO', body)
    worksheet177.write('N22', 'JML', body)
    worksheet177.write('O22', 'MAW', body)
    worksheet177.write('O22', 'MAP', body)
    worksheet177.write('P22', 'IND', body)
    worksheet177.write('Q22', 'ENG', body)
    worksheet177.write('R22', 'SEJ', body)
    worksheet177.write('S22', 'GEO', body)
    worksheet177.write('U22', 'EKO', body)
    worksheet177.write('T22', 'SOS', body)
    worksheet177.write('T22', 'FIS', body)
    worksheet177.write('T22', 'KIM', body)
    worksheet177.write('T22', 'BIO', body)
    worksheet177.write('V22', 'JML', body)

    worksheet177.conditional_format(22, 0, row177+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 178
    worksheet178.insert_image('A1', r'logo resmi nf.jpg')

    worksheet178.set_column('A:A', 7, center)
    worksheet178.set_column('B:B', 6, center)
    worksheet178.set_column('C:C', 18.14, center)
    worksheet178.set_column('D:D', 25, left)
    worksheet178.set_column('E:E', 13.14, left)
    worksheet178.set_column('F:F', 8.57, center)
    worksheet178.set_column('G:V', 5, center)
    worksheet178.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIRACAS', title)
    worksheet178.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet178.write('A5', 'LOKASI', header)
    worksheet178.write('B5', 'TOTAL', header)
    worksheet178.merge_range('A4:B4', 'RANK', header)
    worksheet178.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet178.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet178.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet178.merge_range('F4:F5', 'KELAS', header)
    worksheet178.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet178.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet178.write('G5', 'MAW', body)
    worksheet178.write('H5', 'MAP', body)
    worksheet178.write('I5', 'IND', body)
    worksheet178.write('J5', 'ENG', body)
    worksheet178.write('K5', 'SEJ', body)
    worksheet178.write('L5', 'GEO', body)
    worksheet178.write('M5', 'EKO', body)
    worksheet178.write('N5', 'SOS', body)
    worksheet178.write('O5', 'FIS', body)
    worksheet178.write('P5', 'KIM', body)
    worksheet178.write('Q5', 'BIO', body)
    worksheet178.write('R5', 'JML', body)
    worksheet178.write('S5', 'MAW', body)
    worksheet178.write('T5', 'MAP', body)
    worksheet178.write('U5', 'IND', body)
    worksheet178.write('V5', 'ENG', body)
    worksheet178.write('W5', 'SEJ', body)
    worksheet178.write('X5', 'GEO', body)
    worksheet178.write('Y5', 'EKO', body)
    worksheet178.write('Z5', 'SOS', body)
    worksheet178.write('AA5', 'FIS', body)
    worksheet178.write('AB5', 'KIM', body)
    worksheet178.write('AC5', 'BIO', body)
    worksheet178.write('AD5', 'JML', body)

    worksheet178.conditional_format(5, 0, row178_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet178.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF CIRACAS', title)
    worksheet178.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet178.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet178.write('A22', 'LOKASI', header)
    worksheet178.write('B22', 'TOTAL', header)
    worksheet178.merge_range('A21:B21', 'RANK', header)
    worksheet178.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet178.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet178.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet178.merge_range('F21:F22', 'KELAS', header)
    worksheet178.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet178.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet178.write('G22', 'MAW', body)
    worksheet178.write('H22', 'MAP', body)
    worksheet178.write('I22', 'IND', body)
    worksheet178.write('J22', 'ENG', body)
    worksheet178.write('J22', 'SEJ', body)
    worksheet178.write('K22', 'GEO', body)
    worksheet178.write('M22', 'EKO', body)
    worksheet178.write('L22', 'SOS', body)
    worksheet178.write('L22', 'FIS', body)
    worksheet178.write('L22', 'KIM', body)
    worksheet178.write('L22', 'BIO', body)
    worksheet178.write('N22', 'JML', body)
    worksheet178.write('O22', 'MAW', body)
    worksheet178.write('O22', 'MAP', body)
    worksheet178.write('P22', 'IND', body)
    worksheet178.write('Q22', 'ENG', body)
    worksheet178.write('R22', 'SEJ', body)
    worksheet178.write('S22', 'GEO', body)
    worksheet178.write('U22', 'EKO', body)
    worksheet178.write('T22', 'SOS', body)
    worksheet178.write('T22', 'FIS', body)
    worksheet178.write('T22', 'KIM', body)
    worksheet178.write('T22', 'BIO', body)
    worksheet178.write('V22', 'JML', body)

    worksheet178.conditional_format(22, 0, row178+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 179
    worksheet179.insert_image('A1', r'logo resmi nf.jpg')

    worksheet179.set_column('A:A', 7, center)
    worksheet179.set_column('B:B', 6, center)
    worksheet179.set_column('C:C', 18.14, center)
    worksheet179.set_column('D:D', 25, left)
    worksheet179.set_column('E:E', 13.14, left)
    worksheet179.set_column('F:F', 8.57, center)
    worksheet179.set_column('G:V', 5, center)
    worksheet179.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KAMPUNG MELAYU', title)
    worksheet179.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet179.write('A5', 'LOKASI', header)
    worksheet179.write('B5', 'TOTAL', header)
    worksheet179.merge_range('A4:B4', 'RANK', header)
    worksheet179.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet179.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet179.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet179.merge_range('F4:F5', 'KELAS', header)
    worksheet179.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet179.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet179.write('G5', 'MAW', body)
    worksheet179.write('H5', 'MAP', body)
    worksheet179.write('I5', 'IND', body)
    worksheet179.write('J5', 'ENG', body)
    worksheet179.write('K5', 'SEJ', body)
    worksheet179.write('L5', 'GEO', body)
    worksheet179.write('M5', 'EKO', body)
    worksheet179.write('N5', 'SOS', body)
    worksheet179.write('O5', 'FIS', body)
    worksheet179.write('P5', 'KIM', body)
    worksheet179.write('Q5', 'BIO', body)
    worksheet179.write('R5', 'JML', body)
    worksheet179.write('S5', 'MAW', body)
    worksheet179.write('T5', 'MAP', body)
    worksheet179.write('U5', 'IND', body)
    worksheet179.write('V5', 'ENG', body)
    worksheet179.write('W5', 'SEJ', body)
    worksheet179.write('X5', 'GEO', body)
    worksheet179.write('Y5', 'EKO', body)
    worksheet179.write('Z5', 'SOS', body)
    worksheet179.write('AA5', 'FIS', body)
    worksheet179.write('AB5', 'KIM', body)
    worksheet179.write('AC5', 'BIO', body)
    worksheet179.write('AD5', 'JML', body)

    worksheet179.conditional_format(5, 0, row179_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet179.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF KAMPUNG MELAYU', title)
    worksheet179.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet179.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet179.write('A22', 'LOKASI', header)
    worksheet179.write('B22', 'TOTAL', header)
    worksheet179.merge_range('A21:B21', 'RANK', header)
    worksheet179.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet179.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet179.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet179.merge_range('F21:F22', 'KELAS', header)
    worksheet179.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet179.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet179.write('G22', 'MAW', body)
    worksheet179.write('H22', 'MAP', body)
    worksheet179.write('I22', 'IND', body)
    worksheet179.write('J22', 'ENG', body)
    worksheet179.write('J22', 'SEJ', body)
    worksheet179.write('K22', 'GEO', body)
    worksheet179.write('M22', 'EKO', body)
    worksheet179.write('L22', 'SOS', body)
    worksheet179.write('L22', 'FIS', body)
    worksheet179.write('L22', 'KIM', body)
    worksheet179.write('L22', 'BIO', body)
    worksheet179.write('N22', 'JML', body)
    worksheet179.write('O22', 'MAW', body)
    worksheet179.write('O22', 'MAP', body)
    worksheet179.write('P22', 'IND', body)
    worksheet179.write('Q22', 'ENG', body)
    worksheet179.write('R22', 'SEJ', body)
    worksheet179.write('S22', 'GEO', body)
    worksheet179.write('U22', 'EKO', body)
    worksheet179.write('T22', 'SOS', body)
    worksheet179.write('T22', 'FIS', body)
    worksheet179.write('T22', 'KIM', body)
    worksheet179.write('T22', 'BIO', body)
    worksheet179.write('V22', 'JML', body)

    worksheet179.conditional_format(22, 0, row179+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 180
    worksheet180.insert_image('A1', r'logo resmi nf.jpg')

    worksheet180.set_column('A:A', 7, center)
    worksheet180.set_column('B:B', 6, center)
    worksheet180.set_column('C:C', 18.14, center)
    worksheet180.set_column('D:D', 25, left)
    worksheet180.set_column('E:E', 13.14, left)
    worksheet180.set_column('F:F', 8.57, center)
    worksheet180.set_column('G:V', 5, center)
    worksheet180.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF AKSES UI', title)
    worksheet180.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet180.write('A5', 'LOKASI', header)
    worksheet180.write('B5', 'TOTAL', header)
    worksheet180.merge_range('A4:B4', 'RANK', header)
    worksheet180.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet180.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet180.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet180.merge_range('F4:F5', 'KELAS', header)
    worksheet180.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet180.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet180.write('G5', 'MAW', body)
    worksheet180.write('H5', 'MAP', body)
    worksheet180.write('I5', 'IND', body)
    worksheet180.write('J5', 'ENG', body)
    worksheet180.write('K5', 'SEJ', body)
    worksheet180.write('L5', 'GEO', body)
    worksheet180.write('M5', 'EKO', body)
    worksheet180.write('N5', 'SOS', body)
    worksheet180.write('O5', 'FIS', body)
    worksheet180.write('P5', 'KIM', body)
    worksheet180.write('Q5', 'BIO', body)
    worksheet180.write('R5', 'JML', body)
    worksheet180.write('S5', 'MAW', body)
    worksheet180.write('T5', 'MAP', body)
    worksheet180.write('U5', 'IND', body)
    worksheet180.write('V5', 'ENG', body)
    worksheet180.write('W5', 'SEJ', body)
    worksheet180.write('X5', 'GEO', body)
    worksheet180.write('Y5', 'EKO', body)
    worksheet180.write('Z5', 'SOS', body)
    worksheet180.write('AA5', 'FIS', body)
    worksheet180.write('AB5', 'KIM', body)
    worksheet180.write('AC5', 'BIO', body)
    worksheet180.write('AD5', 'JML', body)

    worksheet180.conditional_format(5, 0, row180_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet180.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF AKSES UI', title)
    worksheet180.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet180.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet180.write('A22', 'LOKASI', header)
    worksheet180.write('B22', 'TOTAL', header)
    worksheet180.merge_range('A21:B21', 'RANK', header)
    worksheet180.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet180.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet180.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet180.merge_range('F21:F22', 'KELAS', header)
    worksheet180.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet180.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet180.write('G22', 'MAW', body)
    worksheet180.write('H22', 'MAP', body)
    worksheet180.write('I22', 'IND', body)
    worksheet180.write('J22', 'ENG', body)
    worksheet180.write('J22', 'SEJ', body)
    worksheet180.write('K22', 'GEO', body)
    worksheet180.write('M22', 'EKO', body)
    worksheet180.write('L22', 'SOS', body)
    worksheet180.write('L22', 'FIS', body)
    worksheet180.write('L22', 'KIM', body)
    worksheet180.write('L22', 'BIO', body)
    worksheet180.write('N22', 'JML', body)
    worksheet180.write('O22', 'MAW', body)
    worksheet180.write('O22', 'MAP', body)
    worksheet180.write('P22', 'IND', body)
    worksheet180.write('Q22', 'ENG', body)
    worksheet180.write('R22', 'SEJ', body)
    worksheet180.write('S22', 'GEO', body)
    worksheet180.write('U22', 'EKO', body)
    worksheet180.write('T22', 'SOS', body)
    worksheet180.write('T22', 'FIS', body)
    worksheet180.write('T22', 'KIM', body)
    worksheet180.write('T22', 'BIO', body)
    worksheet180.write('V22', 'JML', body)

    worksheet180.conditional_format(22, 0, row180+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 181
    worksheet181.insert_image('A1', r'logo resmi nf.jpg')

    worksheet181.set_column('A:A', 7, center)
    worksheet181.set_column('B:B', 6, center)
    worksheet181.set_column('C:C', 18.14, center)
    worksheet181.set_column('D:D', 25, left)
    worksheet181.set_column('E:E', 13.14, left)
    worksheet181.set_column('F:F', 8.57, center)
    worksheet181.set_column('G:V', 5, center)
    worksheet181.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF JATIMEKAR', title)
    worksheet181.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet181.write('A5', 'LOKASI', header)
    worksheet181.write('B5', 'TOTAL', header)
    worksheet181.merge_range('A4:B4', 'RANK', header)
    worksheet181.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet181.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet181.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet181.merge_range('F4:F5', 'KELAS', header)
    worksheet181.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet181.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet181.write('G5', 'MAW', body)
    worksheet181.write('H5', 'MAP', body)
    worksheet181.write('I5', 'IND', body)
    worksheet181.write('J5', 'ENG', body)
    worksheet181.write('K5', 'SEJ', body)
    worksheet181.write('L5', 'GEO', body)
    worksheet181.write('M5', 'EKO', body)
    worksheet181.write('N5', 'SOS', body)
    worksheet181.write('O5', 'FIS', body)
    worksheet181.write('P5', 'KIM', body)
    worksheet181.write('Q5', 'BIO', body)
    worksheet181.write('R5', 'JML', body)
    worksheet181.write('S5', 'MAW', body)
    worksheet181.write('T5', 'MAP', body)
    worksheet181.write('U5', 'IND', body)
    worksheet181.write('V5', 'ENG', body)
    worksheet181.write('W5', 'SEJ', body)
    worksheet181.write('X5', 'GEO', body)
    worksheet181.write('Y5', 'EKO', body)
    worksheet181.write('Z5', 'SOS', body)
    worksheet181.write('AA5', 'FIS', body)
    worksheet181.write('AB5', 'KIM', body)
    worksheet181.write('AC5', 'BIO', body)
    worksheet181.write('AD5', 'JML', body)

    worksheet181.conditional_format(5, 0, row181_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet181.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF JATIMEKAR', title)
    worksheet181.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet181.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet181.write('A22', 'LOKASI', header)
    worksheet181.write('B22', 'TOTAL', header)
    worksheet181.merge_range('A21:B21', 'RANK', header)
    worksheet181.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet181.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet181.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet181.merge_range('F21:F22', 'KELAS', header)
    worksheet181.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet181.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet181.write('G22', 'MAW', body)
    worksheet181.write('H22', 'MAP', body)
    worksheet181.write('I22', 'IND', body)
    worksheet181.write('J22', 'ENG', body)
    worksheet181.write('J22', 'SEJ', body)
    worksheet181.write('K22', 'GEO', body)
    worksheet181.write('M22', 'EKO', body)
    worksheet181.write('L22', 'SOS', body)
    worksheet181.write('L22', 'FIS', body)
    worksheet181.write('L22', 'KIM', body)
    worksheet181.write('L22', 'BIO', body)
    worksheet181.write('N22', 'JML', body)
    worksheet181.write('O22', 'MAW', body)
    worksheet181.write('O22', 'MAP', body)
    worksheet181.write('P22', 'IND', body)
    worksheet181.write('Q22', 'ENG', body)
    worksheet181.write('R22', 'SEJ', body)
    worksheet181.write('S22', 'GEO', body)
    worksheet181.write('U22', 'EKO', body)
    worksheet181.write('T22', 'SOS', body)
    worksheet181.write('T22', 'FIS', body)
    worksheet181.write('T22', 'KIM', body)
    worksheet181.write('T22', 'BIO', body)
    worksheet181.write('V22', 'JML', body)

    worksheet181.conditional_format(22, 0, row181+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 182
    worksheet182.insert_image('A1', r'logo resmi nf.jpg')

    worksheet182.set_column('A:A', 7, center)
    worksheet182.set_column('B:B', 6, center)
    worksheet182.set_column('C:C', 18.14, center)
    worksheet182.set_column('D:D', 25, left)
    worksheet182.set_column('E:E', 13.14, left)
    worksheet182.set_column('F:F', 8.57, center)
    worksheet182.set_column('G:V', 5, center)
    worksheet182.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF RAWALUMBU', title)
    worksheet182.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet182.write('A5', 'LOKASI', header)
    worksheet182.write('B5', 'TOTAL', header)
    worksheet182.merge_range('A4:B4', 'RANK', header)
    worksheet182.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet182.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet182.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet182.merge_range('F4:F5', 'KELAS', header)
    worksheet182.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet182.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet182.write('G5', 'MAW', body)
    worksheet182.write('H5', 'MAP', body)
    worksheet182.write('I5', 'IND', body)
    worksheet182.write('J5', 'ENG', body)
    worksheet182.write('K5', 'SEJ', body)
    worksheet182.write('L5', 'GEO', body)
    worksheet182.write('M5', 'EKO', body)
    worksheet182.write('N5', 'SOS', body)
    worksheet182.write('O5', 'FIS', body)
    worksheet182.write('P5', 'KIM', body)
    worksheet182.write('Q5', 'BIO', body)
    worksheet182.write('R5', 'JML', body)
    worksheet182.write('S5', 'MAW', body)
    worksheet182.write('T5', 'MAP', body)
    worksheet182.write('U5', 'IND', body)
    worksheet182.write('V5', 'ENG', body)
    worksheet182.write('W5', 'SEJ', body)
    worksheet182.write('X5', 'GEO', body)
    worksheet182.write('Y5', 'EKO', body)
    worksheet182.write('Z5', 'SOS', body)
    worksheet182.write('AA5', 'FIS', body)
    worksheet182.write('AB5', 'KIM', body)
    worksheet182.write('AC5', 'BIO', body)
    worksheet182.write('AD5', 'JML', body)

    worksheet182.conditional_format(5, 0, row182_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet182.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF RAWALUMBU', title)
    worksheet182.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet182.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet182.write('A22', 'LOKASI', header)
    worksheet182.write('B22', 'TOTAL', header)
    worksheet182.merge_range('A21:B21', 'RANK', header)
    worksheet182.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet182.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet182.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet182.merge_range('F21:F22', 'KELAS', header)
    worksheet182.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet182.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet182.write('G22', 'MAW', body)
    worksheet182.write('H22', 'MAP', body)
    worksheet182.write('I22', 'IND', body)
    worksheet182.write('J22', 'ENG', body)
    worksheet182.write('J22', 'SEJ', body)
    worksheet182.write('K22', 'GEO', body)
    worksheet182.write('M22', 'EKO', body)
    worksheet182.write('L22', 'SOS', body)
    worksheet182.write('L22', 'FIS', body)
    worksheet182.write('L22', 'KIM', body)
    worksheet182.write('L22', 'BIO', body)
    worksheet182.write('N22', 'JML', body)
    worksheet182.write('O22', 'MAW', body)
    worksheet182.write('O22', 'MAP', body)
    worksheet182.write('P22', 'IND', body)
    worksheet182.write('Q22', 'ENG', body)
    worksheet182.write('R22', 'SEJ', body)
    worksheet182.write('S22', 'GEO', body)
    worksheet182.write('U22', 'EKO', body)
    worksheet182.write('T22', 'SOS', body)
    worksheet182.write('T22', 'FIS', body)
    worksheet182.write('T22', 'KIM', body)
    worksheet182.write('T22', 'BIO', body)
    worksheet182.write('V22', 'JML', body)

    worksheet182.conditional_format(22, 0, row182+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 183
    worksheet183.insert_image('A1', r'logo resmi nf.jpg')

    worksheet183.set_column('A:A', 7, center)
    worksheet183.set_column('B:B', 6, center)
    worksheet183.set_column('C:C', 18.14, center)
    worksheet183.set_column('D:D', 25, left)
    worksheet183.set_column('E:E', 13.14, left)
    worksheet183.set_column('F:F', 8.57, center)
    worksheet183.set_column('G:V', 5, center)
    worksheet183.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF TAMAN HARAPAN BARU', title)
    worksheet183.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet183.write('A5', 'LOKASI', header)
    worksheet183.write('B5', 'TOTAL', header)
    worksheet183.merge_range('A4:B4', 'RANK', header)
    worksheet183.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet183.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet183.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet183.merge_range('F4:F5', 'KELAS', header)
    worksheet183.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet183.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet183.write('G5', 'MAW', body)
    worksheet183.write('H5', 'MAP', body)
    worksheet183.write('I5', 'IND', body)
    worksheet183.write('J5', 'ENG', body)
    worksheet183.write('K5', 'SEJ', body)
    worksheet183.write('L5', 'GEO', body)
    worksheet183.write('M5', 'EKO', body)
    worksheet183.write('N5', 'SOS', body)
    worksheet183.write('O5', 'FIS', body)
    worksheet183.write('P5', 'KIM', body)
    worksheet183.write('Q5', 'BIO', body)
    worksheet183.write('R5', 'JML', body)
    worksheet183.write('S5', 'MAW', body)
    worksheet183.write('T5', 'MAP', body)
    worksheet183.write('U5', 'IND', body)
    worksheet183.write('V5', 'ENG', body)
    worksheet183.write('W5', 'SEJ', body)
    worksheet183.write('X5', 'GEO', body)
    worksheet183.write('Y5', 'EKO', body)
    worksheet183.write('Z5', 'SOS', body)
    worksheet183.write('AA5', 'FIS', body)
    worksheet183.write('AB5', 'KIM', body)
    worksheet183.write('AC5', 'BIO', body)
    worksheet183.write('AD5', 'JML', body)

    worksheet183.conditional_format(5, 0, row183_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet183.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF TAMAN HARAPAN BARU', title)
    worksheet183.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet183.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet183.write('A22', 'LOKASI', header)
    worksheet183.write('B22', 'TOTAL', header)
    worksheet183.merge_range('A21:B21', 'RANK', header)
    worksheet183.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet183.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet183.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet183.merge_range('F21:F22', 'KELAS', header)
    worksheet183.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet183.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet183.write('G22', 'MAW', body)
    worksheet183.write('H22', 'MAP', body)
    worksheet183.write('I22', 'IND', body)
    worksheet183.write('J22', 'ENG', body)
    worksheet183.write('J22', 'SEJ', body)
    worksheet183.write('K22', 'GEO', body)
    worksheet183.write('M22', 'EKO', body)
    worksheet183.write('L22', 'SOS', body)
    worksheet183.write('L22', 'FIS', body)
    worksheet183.write('L22', 'KIM', body)
    worksheet183.write('L22', 'BIO', body)
    worksheet183.write('N22', 'JML', body)
    worksheet183.write('O22', 'MAW', body)
    worksheet183.write('O22', 'MAP', body)
    worksheet183.write('P22', 'IND', body)
    worksheet183.write('Q22', 'ENG', body)
    worksheet183.write('R22', 'SEJ', body)
    worksheet183.write('S22', 'GEO', body)
    worksheet183.write('U22', 'EKO', body)
    worksheet183.write('T22', 'SOS', body)
    worksheet183.write('T22', 'FIS', body)
    worksheet183.write('T22', 'KIM', body)
    worksheet183.write('T22', 'BIO', body)
    worksheet183.write('V22', 'JML', body)

    worksheet183.conditional_format(22, 0, row183+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 184
    worksheet184.insert_image('A1', r'logo resmi nf.jpg')

    worksheet184.set_column('A:A', 7, center)
    worksheet184.set_column('B:B', 6, center)
    worksheet184.set_column('C:C', 18.14, center)
    worksheet184.set_column('D:D', 25, left)
    worksheet184.set_column('E:E', 13.14, left)
    worksheet184.set_column('F:F', 8.57, center)
    worksheet184.set_column('G:V', 5, center)
    worksheet184.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF VILA NUSA INDAH', title)
    worksheet184.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet184.write('A5', 'LOKASI', header)
    worksheet184.write('B5', 'TOTAL', header)
    worksheet184.merge_range('A4:B4', 'RANK', header)
    worksheet184.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet184.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet184.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet184.merge_range('F4:F5', 'KELAS', header)
    worksheet184.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet184.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet184.write('G5', 'MAW', body)
    worksheet184.write('H5', 'MAP', body)
    worksheet184.write('I5', 'IND', body)
    worksheet184.write('J5', 'ENG', body)
    worksheet184.write('K5', 'SEJ', body)
    worksheet184.write('L5', 'GEO', body)
    worksheet184.write('M5', 'EKO', body)
    worksheet184.write('N5', 'SOS', body)
    worksheet184.write('O5', 'FIS', body)
    worksheet184.write('P5', 'KIM', body)
    worksheet184.write('Q5', 'BIO', body)
    worksheet184.write('R5', 'JML', body)
    worksheet184.write('S5', 'MAW', body)
    worksheet184.write('T5', 'MAP', body)
    worksheet184.write('U5', 'IND', body)
    worksheet184.write('V5', 'ENG', body)
    worksheet184.write('W5', 'SEJ', body)
    worksheet184.write('X5', 'GEO', body)
    worksheet184.write('Y5', 'EKO', body)
    worksheet184.write('Z5', 'SOS', body)
    worksheet184.write('AA5', 'FIS', body)
    worksheet184.write('AB5', 'KIM', body)
    worksheet184.write('AC5', 'BIO', body)
    worksheet184.write('AD5', 'JML', body)

    worksheet184.conditional_format(5, 0, row184_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet184.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF VILA NUSA INDAH', title)
    worksheet184.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet184.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet184.write('A22', 'LOKASI', header)
    worksheet184.write('B22', 'TOTAL', header)
    worksheet184.merge_range('A21:B21', 'RANK', header)
    worksheet184.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet184.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet184.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet184.merge_range('F21:F22', 'KELAS', header)
    worksheet184.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet184.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet184.write('G22', 'MAW', body)
    worksheet184.write('H22', 'MAP', body)
    worksheet184.write('I22', 'IND', body)
    worksheet184.write('J22', 'ENG', body)
    worksheet184.write('J22', 'SEJ', body)
    worksheet184.write('K22', 'GEO', body)
    worksheet184.write('M22', 'EKO', body)
    worksheet184.write('L22', 'SOS', body)
    worksheet184.write('L22', 'FIS', body)
    worksheet184.write('L22', 'KIM', body)
    worksheet184.write('L22', 'BIO', body)
    worksheet184.write('N22', 'JML', body)
    worksheet184.write('O22', 'MAW', body)
    worksheet184.write('O22', 'MAP', body)
    worksheet184.write('P22', 'IND', body)
    worksheet184.write('Q22', 'ENG', body)
    worksheet184.write('R22', 'SEJ', body)
    worksheet184.write('S22', 'GEO', body)
    worksheet184.write('U22', 'EKO', body)
    worksheet184.write('T22', 'SOS', body)
    worksheet184.write('T22', 'FIS', body)
    worksheet184.write('T22', 'KIM', body)
    worksheet184.write('T22', 'BIO', body)
    worksheet184.write('V22', 'JML', body)

    worksheet184.conditional_format(22, 0, row184+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 185
    worksheet185.insert_image('A1', r'logo resmi nf.jpg')

    worksheet185.set_column('A:A', 7, center)
    worksheet185.set_column('B:B', 6, center)
    worksheet185.set_column('C:C', 18.14, center)
    worksheet185.set_column('D:D', 25, left)
    worksheet185.set_column('E:E', 13.14, left)
    worksheet185.set_column('F:F', 8.57, center)
    worksheet185.set_column('G:V', 5, center)
    worksheet185.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF JATIWARNA', title)
    worksheet185.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet185.write('A5', 'LOKASI', header)
    worksheet185.write('B5', 'TOTAL', header)
    worksheet185.merge_range('A4:B4', 'RANK', header)
    worksheet185.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet185.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet185.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet185.merge_range('F4:F5', 'KELAS', header)
    worksheet185.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet185.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet185.write('G5', 'MAW', body)
    worksheet185.write('H5', 'MAP', body)
    worksheet185.write('I5', 'IND', body)
    worksheet185.write('J5', 'ENG', body)
    worksheet185.write('K5', 'SEJ', body)
    worksheet185.write('L5', 'GEO', body)
    worksheet185.write('M5', 'EKO', body)
    worksheet185.write('N5', 'SOS', body)
    worksheet185.write('O5', 'FIS', body)
    worksheet185.write('P5', 'KIM', body)
    worksheet185.write('Q5', 'BIO', body)
    worksheet185.write('R5', 'JML', body)
    worksheet185.write('S5', 'MAW', body)
    worksheet185.write('T5', 'MAP', body)
    worksheet185.write('U5', 'IND', body)
    worksheet185.write('V5', 'ENG', body)
    worksheet185.write('W5', 'SEJ', body)
    worksheet185.write('X5', 'GEO', body)
    worksheet185.write('Y5', 'EKO', body)
    worksheet185.write('Z5', 'SOS', body)
    worksheet185.write('AA5', 'FIS', body)
    worksheet185.write('AB5', 'KIM', body)
    worksheet185.write('AC5', 'BIO', body)
    worksheet185.write('AD5', 'JML', body)

    worksheet185.conditional_format(5, 0, row185_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet185.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF JATIWARNA', title)
    worksheet185.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet185.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet185.write('A22', 'LOKASI', header)
    worksheet185.write('B22', 'TOTAL', header)
    worksheet185.merge_range('A21:B21', 'RANK', header)
    worksheet185.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet185.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet185.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet185.merge_range('F21:F22', 'KELAS', header)
    worksheet185.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet185.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet185.write('G22', 'MAW', body)
    worksheet185.write('H22', 'MAP', body)
    worksheet185.write('I22', 'IND', body)
    worksheet185.write('J22', 'ENG', body)
    worksheet185.write('J22', 'SEJ', body)
    worksheet185.write('K22', 'GEO', body)
    worksheet185.write('M22', 'EKO', body)
    worksheet185.write('L22', 'SOS', body)
    worksheet185.write('L22', 'FIS', body)
    worksheet185.write('L22', 'KIM', body)
    worksheet185.write('L22', 'BIO', body)
    worksheet185.write('N22', 'JML', body)
    worksheet185.write('O22', 'MAW', body)
    worksheet185.write('O22', 'MAP', body)
    worksheet185.write('P22', 'IND', body)
    worksheet185.write('Q22', 'ENG', body)
    worksheet185.write('R22', 'SEJ', body)
    worksheet185.write('S22', 'GEO', body)
    worksheet185.write('U22', 'EKO', body)
    worksheet185.write('T22', 'SOS', body)
    worksheet185.write('T22', 'FIS', body)
    worksheet185.write('T22', 'KIM', body)
    worksheet185.write('T22', 'BIO', body)
    worksheet185.write('V22', 'JML', body)

    worksheet185.conditional_format(22, 0, row185+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 186
    worksheet186.insert_image('A1', r'logo resmi nf.jpg')

    worksheet186.set_column('A:A', 7, center)
    worksheet186.set_column('B:B', 6, center)
    worksheet186.set_column('C:C', 18.14, center)
    worksheet186.set_column('D:D', 25, left)
    worksheet186.set_column('E:E', 13.14, left)
    worksheet186.set_column('F:F', 8.57, center)
    worksheet186.set_column('G:V', 5, center)
    worksheet186.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF TAMBUN', title)
    worksheet186.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet186.write('A5', 'LOKASI', header)
    worksheet186.write('B5', 'TOTAL', header)
    worksheet186.merge_range('A4:B4', 'RANK', header)
    worksheet186.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet186.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet186.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet186.merge_range('F4:F5', 'KELAS', header)
    worksheet186.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet186.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet186.write('G5', 'MAW', body)
    worksheet186.write('H5', 'MAP', body)
    worksheet186.write('I5', 'IND', body)
    worksheet186.write('J5', 'ENG', body)
    worksheet186.write('K5', 'SEJ', body)
    worksheet186.write('L5', 'GEO', body)
    worksheet186.write('M5', 'EKO', body)
    worksheet186.write('N5', 'SOS', body)
    worksheet186.write('O5', 'FIS', body)
    worksheet186.write('P5', 'KIM', body)
    worksheet186.write('Q5', 'BIO', body)
    worksheet186.write('R5', 'JML', body)
    worksheet186.write('S5', 'MAW', body)
    worksheet186.write('T5', 'MAP', body)
    worksheet186.write('U5', 'IND', body)
    worksheet186.write('V5', 'ENG', body)
    worksheet186.write('W5', 'SEJ', body)
    worksheet186.write('X5', 'GEO', body)
    worksheet186.write('Y5', 'EKO', body)
    worksheet186.write('Z5', 'SOS', body)
    worksheet186.write('AA5', 'FIS', body)
    worksheet186.write('AB5', 'KIM', body)
    worksheet186.write('AC5', 'BIO', body)
    worksheet186.write('AD5', 'JML', body)

    worksheet186.conditional_format(5, 0, row186_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet186.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF TAMBUN', title)
    worksheet186.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet186.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet186.write('A22', 'LOKASI', header)
    worksheet186.write('B22', 'TOTAL', header)
    worksheet186.merge_range('A21:B21', 'RANK', header)
    worksheet186.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet186.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet186.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet186.merge_range('F21:F22', 'KELAS', header)
    worksheet186.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet186.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet186.write('G22', 'MAW', body)
    worksheet186.write('H22', 'MAP', body)
    worksheet186.write('I22', 'IND', body)
    worksheet186.write('J22', 'ENG', body)
    worksheet186.write('J22', 'SEJ', body)
    worksheet186.write('K22', 'GEO', body)
    worksheet186.write('M22', 'EKO', body)
    worksheet186.write('L22', 'SOS', body)
    worksheet186.write('L22', 'FIS', body)
    worksheet186.write('L22', 'KIM', body)
    worksheet186.write('L22', 'BIO', body)
    worksheet186.write('N22', 'JML', body)
    worksheet186.write('O22', 'MAW', body)
    worksheet186.write('O22', 'MAP', body)
    worksheet186.write('P22', 'IND', body)
    worksheet186.write('Q22', 'ENG', body)
    worksheet186.write('R22', 'SEJ', body)
    worksheet186.write('S22', 'GEO', body)
    worksheet186.write('U22', 'EKO', body)
    worksheet186.write('T22', 'SOS', body)
    worksheet186.write('T22', 'FIS', body)
    worksheet186.write('T22', 'KIM', body)
    worksheet186.write('T22', 'BIO', body)
    worksheet186.write('V22', 'JML', body)

    worksheet186.conditional_format(22, 0, row186+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 187
    worksheet187.insert_image('A1', r'logo resmi nf.jpg')

    worksheet187.set_column('A:A', 7, center)
    worksheet187.set_column('B:B', 6, center)
    worksheet187.set_column('C:C', 18.14, center)
    worksheet187.set_column('D:D', 25, left)
    worksheet187.set_column('E:E', 13.14, left)
    worksheet187.set_column('F:F', 8.57, center)
    worksheet187.set_column('G:V', 5, center)
    worksheet187.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF DAAN MOGOT', title)
    worksheet187.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet187.write('A5', 'LOKASI', header)
    worksheet187.write('B5', 'TOTAL', header)
    worksheet187.merge_range('A4:B4', 'RANK', header)
    worksheet187.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet187.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet187.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet187.merge_range('F4:F5', 'KELAS', header)
    worksheet187.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet187.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet187.write('G5', 'MAW', body)
    worksheet187.write('H5', 'MAP', body)
    worksheet187.write('I5', 'IND', body)
    worksheet187.write('J5', 'ENG', body)
    worksheet187.write('K5', 'SEJ', body)
    worksheet187.write('L5', 'GEO', body)
    worksheet187.write('M5', 'EKO', body)
    worksheet187.write('N5', 'SOS', body)
    worksheet187.write('O5', 'FIS', body)
    worksheet187.write('P5', 'KIM', body)
    worksheet187.write('Q5', 'BIO', body)
    worksheet187.write('R5', 'JML', body)
    worksheet187.write('S5', 'MAW', body)
    worksheet187.write('T5', 'MAP', body)
    worksheet187.write('U5', 'IND', body)
    worksheet187.write('V5', 'ENG', body)
    worksheet187.write('W5', 'SEJ', body)
    worksheet187.write('X5', 'GEO', body)
    worksheet187.write('Y5', 'EKO', body)
    worksheet187.write('Z5', 'SOS', body)
    worksheet187.write('AA5', 'FIS', body)
    worksheet187.write('AB5', 'KIM', body)
    worksheet187.write('AC5', 'BIO', body)
    worksheet187.write('AD5', 'JML', body)

    worksheet187.conditional_format(5, 0, row187_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet187.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF DAAN MOGOT', title)
    worksheet187.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet187.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet187.write('A22', 'LOKASI', header)
    worksheet187.write('B22', 'TOTAL', header)
    worksheet187.merge_range('A21:B21', 'RANK', header)
    worksheet187.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet187.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet187.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet187.merge_range('F21:F22', 'KELAS', header)
    worksheet187.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet187.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet187.write('G22', 'MAW', body)
    worksheet187.write('H22', 'MAP', body)
    worksheet187.write('I22', 'IND', body)
    worksheet187.write('J22', 'ENG', body)
    worksheet187.write('J22', 'SEJ', body)
    worksheet187.write('K22', 'GEO', body)
    worksheet187.write('M22', 'EKO', body)
    worksheet187.write('L22', 'SOS', body)
    worksheet187.write('L22', 'FIS', body)
    worksheet187.write('L22', 'KIM', body)
    worksheet187.write('L22', 'BIO', body)
    worksheet187.write('N22', 'JML', body)
    worksheet187.write('O22', 'MAW', body)
    worksheet187.write('O22', 'MAP', body)
    worksheet187.write('P22', 'IND', body)
    worksheet187.write('Q22', 'ENG', body)
    worksheet187.write('R22', 'SEJ', body)
    worksheet187.write('S22', 'GEO', body)
    worksheet187.write('U22', 'EKO', body)
    worksheet187.write('T22', 'SOS', body)
    worksheet187.write('T22', 'FIS', body)
    worksheet187.write('T22', 'KIM', body)
    worksheet187.write('T22', 'BIO', body)
    worksheet187.write('V22', 'JML', body)

    worksheet187.conditional_format(22, 0, row187+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 189
    worksheet189.insert_image('A1', r'logo resmi nf.jpg')

    worksheet189.set_column('A:A', 7, center)
    worksheet189.set_column('B:B', 6, center)
    worksheet189.set_column('C:C', 18.14, center)
    worksheet189.set_column('D:D', 25, left)
    worksheet189.set_column('E:E', 13.14, left)
    worksheet189.set_column('F:F', 8.57, center)
    worksheet189.set_column('G:V', 5, center)
    worksheet189.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIBUBUR', title)
    worksheet189.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet189.write('A5', 'LOKASI', header)
    worksheet189.write('B5', 'TOTAL', header)
    worksheet189.merge_range('A4:B4', 'RANK', header)
    worksheet189.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet189.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet189.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet189.merge_range('F4:F5', 'KELAS', header)
    worksheet189.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet189.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet189.write('G5', 'MAW', body)
    worksheet189.write('H5', 'MAP', body)
    worksheet189.write('I5', 'IND', body)
    worksheet189.write('J5', 'ENG', body)
    worksheet189.write('K5', 'SEJ', body)
    worksheet189.write('L5', 'GEO', body)
    worksheet189.write('M5', 'EKO', body)
    worksheet189.write('N5', 'SOS', body)
    worksheet189.write('O5', 'FIS', body)
    worksheet189.write('P5', 'KIM', body)
    worksheet189.write('Q5', 'BIO', body)
    worksheet189.write('R5', 'JML', body)
    worksheet189.write('S5', 'MAW', body)
    worksheet189.write('T5', 'MAP', body)
    worksheet189.write('U5', 'IND', body)
    worksheet189.write('V5', 'ENG', body)
    worksheet189.write('W5', 'SEJ', body)
    worksheet189.write('X5', 'GEO', body)
    worksheet189.write('Y5', 'EKO', body)
    worksheet189.write('Z5', 'SOS', body)
    worksheet189.write('AA5', 'FIS', body)
    worksheet189.write('AB5', 'KIM', body)
    worksheet189.write('AC5', 'BIO', body)
    worksheet189.write('AD5', 'JML', body)

    worksheet189.conditional_format(5, 0, row189_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet189.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF CIBUBUR', title)
    worksheet189.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet189.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet189.write('A22', 'LOKASI', header)
    worksheet189.write('B22', 'TOTAL', header)
    worksheet189.merge_range('A21:B21', 'RANK', header)
    worksheet189.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet189.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet189.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet189.merge_range('F21:F22', 'KELAS', header)
    worksheet189.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet189.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet189.write('G22', 'MAW', body)
    worksheet189.write('H22', 'MAP', body)
    worksheet189.write('I22', 'IND', body)
    worksheet189.write('J22', 'ENG', body)
    worksheet189.write('J22', 'SEJ', body)
    worksheet189.write('K22', 'GEO', body)
    worksheet189.write('M22', 'EKO', body)
    worksheet189.write('L22', 'SOS', body)
    worksheet189.write('L22', 'FIS', body)
    worksheet189.write('L22', 'KIM', body)
    worksheet189.write('L22', 'BIO', body)
    worksheet189.write('N22', 'JML', body)
    worksheet189.write('O22', 'MAW', body)
    worksheet189.write('O22', 'MAP', body)
    worksheet189.write('P22', 'IND', body)
    worksheet189.write('Q22', 'ENG', body)
    worksheet189.write('R22', 'SEJ', body)
    worksheet189.write('S22', 'GEO', body)
    worksheet189.write('U22', 'EKO', body)
    worksheet189.write('T22', 'SOS', body)
    worksheet189.write('T22', 'FIS', body)
    worksheet189.write('T22', 'KIM', body)
    worksheet189.write('T22', 'BIO', body)
    worksheet189.write('V22', 'JML', body)

    worksheet189.conditional_format(22, 0, row189+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 190
    worksheet190.insert_image('A1', r'logo resmi nf.jpg')

    worksheet190.set_column('A:A', 7, center)
    worksheet190.set_column('B:B', 6, center)
    worksheet190.set_column('C:C', 18.14, center)
    worksheet190.set_column('D:D', 25, left)
    worksheet190.set_column('E:E', 13.14, left)
    worksheet190.set_column('F:F', 8.57, center)
    worksheet190.set_column('G:V', 5, center)
    worksheet190.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CENGKARENG', title)
    worksheet190.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet190.write('A5', 'LOKASI', header)
    worksheet190.write('B5', 'TOTAL', header)
    worksheet190.merge_range('A4:B4', 'RANK', header)
    worksheet190.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet190.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet190.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet190.merge_range('F4:F5', 'KELAS', header)
    worksheet190.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet190.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet190.write('G5', 'MAW', body)
    worksheet190.write('H5', 'MAP', body)
    worksheet190.write('I5', 'IND', body)
    worksheet190.write('J5', 'ENG', body)
    worksheet190.write('K5', 'SEJ', body)
    worksheet190.write('L5', 'GEO', body)
    worksheet190.write('M5', 'EKO', body)
    worksheet190.write('N5', 'SOS', body)
    worksheet190.write('O5', 'FIS', body)
    worksheet190.write('P5', 'KIM', body)
    worksheet190.write('Q5', 'BIO', body)
    worksheet190.write('R5', 'JML', body)
    worksheet190.write('S5', 'MAW', body)
    worksheet190.write('T5', 'MAP', body)
    worksheet190.write('U5', 'IND', body)
    worksheet190.write('V5', 'ENG', body)
    worksheet190.write('W5', 'SEJ', body)
    worksheet190.write('X5', 'GEO', body)
    worksheet190.write('Y5', 'EKO', body)
    worksheet190.write('Z5', 'SOS', body)
    worksheet190.write('AA5', 'FIS', body)
    worksheet190.write('AB5', 'KIM', body)
    worksheet190.write('AC5', 'BIO', body)
    worksheet190.write('AD5', 'JML', body)

    worksheet190.conditional_format(5, 0, row190_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet190.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF CENGKARENG', title)
    worksheet190.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet190.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet190.write('A22', 'LOKASI', header)
    worksheet190.write('B22', 'TOTAL', header)
    worksheet190.merge_range('A21:B21', 'RANK', header)
    worksheet190.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet190.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet190.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet190.merge_range('F21:F22', 'KELAS', header)
    worksheet190.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet190.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet190.write('G22', 'MAW', body)
    worksheet190.write('H22', 'MAP', body)
    worksheet190.write('I22', 'IND', body)
    worksheet190.write('J22', 'ENG', body)
    worksheet190.write('J22', 'SEJ', body)
    worksheet190.write('K22', 'GEO', body)
    worksheet190.write('M22', 'EKO', body)
    worksheet190.write('L22', 'SOS', body)
    worksheet190.write('L22', 'FIS', body)
    worksheet190.write('L22', 'KIM', body)
    worksheet190.write('L22', 'BIO', body)
    worksheet190.write('N22', 'JML', body)
    worksheet190.write('O22', 'MAW', body)
    worksheet190.write('O22', 'MAP', body)
    worksheet190.write('P22', 'IND', body)
    worksheet190.write('Q22', 'ENG', body)
    worksheet190.write('R22', 'SEJ', body)
    worksheet190.write('S22', 'GEO', body)
    worksheet190.write('U22', 'EKO', body)
    worksheet190.write('T22', 'SOS', body)
    worksheet190.write('T22', 'FIS', body)
    worksheet190.write('T22', 'KIM', body)
    worksheet190.write('T22', 'BIO', body)
    worksheet190.write('V22', 'JML', body)

    worksheet190.conditional_format(22, 0, row190+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 191
    worksheet191.insert_image('A1', r'logo resmi nf.jpg')

    worksheet191.set_column('A:A', 7, center)
    worksheet191.set_column('B:B', 6, center)
    worksheet191.set_column('C:C', 18.14, center)
    worksheet191.set_column('D:D', 25, left)
    worksheet191.set_column('E:E', 13.14, left)
    worksheet191.set_column('F:F', 8.57, center)
    worksheet191.set_column('G:V', 5, center)
    worksheet191.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PETUKANGAN', title)
    worksheet191.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet191.write('A5', 'LOKASI', header)
    worksheet191.write('B5', 'TOTAL', header)
    worksheet191.merge_range('A4:B4', 'RANK', header)
    worksheet191.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet191.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet191.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet191.merge_range('F4:F5', 'KELAS', header)
    worksheet191.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet191.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet191.write('G5', 'MAW', body)
    worksheet191.write('H5', 'MAP', body)
    worksheet191.write('I5', 'IND', body)
    worksheet191.write('J5', 'ENG', body)
    worksheet191.write('K5', 'SEJ', body)
    worksheet191.write('L5', 'GEO', body)
    worksheet191.write('M5', 'EKO', body)
    worksheet191.write('N5', 'SOS', body)
    worksheet191.write('O5', 'FIS', body)
    worksheet191.write('P5', 'KIM', body)
    worksheet191.write('Q5', 'BIO', body)
    worksheet191.write('R5', 'JML', body)
    worksheet191.write('S5', 'MAW', body)
    worksheet191.write('T5', 'MAP', body)
    worksheet191.write('U5', 'IND', body)
    worksheet191.write('V5', 'ENG', body)
    worksheet191.write('W5', 'SEJ', body)
    worksheet191.write('X5', 'GEO', body)
    worksheet191.write('Y5', 'EKO', body)
    worksheet191.write('Z5', 'SOS', body)
    worksheet191.write('AA5', 'FIS', body)
    worksheet191.write('AB5', 'KIM', body)
    worksheet191.write('AC5', 'BIO', body)
    worksheet191.write('AD5', 'JML', body)

    worksheet191.conditional_format(5, 0, row191_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet191.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF PETUKANGAN', title)
    worksheet191.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet191.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet191.write('A22', 'LOKASI', header)
    worksheet191.write('B22', 'TOTAL', header)
    worksheet191.merge_range('A21:B21', 'RANK', header)
    worksheet191.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet191.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet191.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet191.merge_range('F21:F22', 'KELAS', header)
    worksheet191.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet191.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet191.write('G22', 'MAW', body)
    worksheet191.write('H22', 'MAP', body)
    worksheet191.write('I22', 'IND', body)
    worksheet191.write('J22', 'ENG', body)
    worksheet191.write('J22', 'SEJ', body)
    worksheet191.write('K22', 'GEO', body)
    worksheet191.write('M22', 'EKO', body)
    worksheet191.write('L22', 'SOS', body)
    worksheet191.write('L22', 'FIS', body)
    worksheet191.write('L22', 'KIM', body)
    worksheet191.write('L22', 'BIO', body)
    worksheet191.write('N22', 'JML', body)
    worksheet191.write('O22', 'MAW', body)
    worksheet191.write('O22', 'MAP', body)
    worksheet191.write('P22', 'IND', body)
    worksheet191.write('Q22', 'ENG', body)
    worksheet191.write('R22', 'SEJ', body)
    worksheet191.write('S22', 'GEO', body)
    worksheet191.write('U22', 'EKO', body)
    worksheet191.write('T22', 'SOS', body)
    worksheet191.write('T22', 'FIS', body)
    worksheet191.write('T22', 'KIM', body)
    worksheet191.write('T22', 'BIO', body)
    worksheet191.write('V22', 'JML', body)

    worksheet191.conditional_format(22, 0, row191+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 192
    worksheet192.insert_image('A1', r'logo resmi nf.jpg')

    worksheet192.set_column('A:A', 7, center)
    worksheet192.set_column('B:B', 6, center)
    worksheet192.set_column('C:C', 18.14, center)
    worksheet192.set_column('D:D', 25, left)
    worksheet192.set_column('E:E', 13.14, left)
    worksheet192.set_column('F:F', 8.57, center)
    worksheet192.set_column('G:V', 5, center)
    worksheet192.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MERUYA UTARA', title)
    worksheet192.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet192.write('A5', 'LOKASI', header)
    worksheet192.write('B5', 'TOTAL', header)
    worksheet192.merge_range('A4:B4', 'RANK', header)
    worksheet192.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet192.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet192.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet192.merge_range('F4:F5', 'KELAS', header)
    worksheet192.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet192.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet192.write('G5', 'MAW', body)
    worksheet192.write('H5', 'MAP', body)
    worksheet192.write('I5', 'IND', body)
    worksheet192.write('J5', 'ENG', body)
    worksheet192.write('K5', 'SEJ', body)
    worksheet192.write('L5', 'GEO', body)
    worksheet192.write('M5', 'EKO', body)
    worksheet192.write('N5', 'SOS', body)
    worksheet192.write('O5', 'FIS', body)
    worksheet192.write('P5', 'KIM', body)
    worksheet192.write('Q5', 'BIO', body)
    worksheet192.write('R5', 'JML', body)
    worksheet192.write('S5', 'MAW', body)
    worksheet192.write('T5', 'MAP', body)
    worksheet192.write('U5', 'IND', body)
    worksheet192.write('V5', 'ENG', body)
    worksheet192.write('W5', 'SEJ', body)
    worksheet192.write('X5', 'GEO', body)
    worksheet192.write('Y5', 'EKO', body)
    worksheet192.write('Z5', 'SOS', body)
    worksheet192.write('AA5', 'FIS', body)
    worksheet192.write('AB5', 'KIM', body)
    worksheet192.write('AC5', 'BIO', body)
    worksheet192.write('AD5', 'JML', body)

    worksheet192.conditional_format(5, 0, row192_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet192.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF MERUYA UTARA', title)
    worksheet192.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet192.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet192.write('A22', 'LOKASI', header)
    worksheet192.write('B22', 'TOTAL', header)
    worksheet192.merge_range('A21:B21', 'RANK', header)
    worksheet192.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet192.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet192.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet192.merge_range('F21:F22', 'KELAS', header)
    worksheet192.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet192.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet192.write('G22', 'MAW', body)
    worksheet192.write('H22', 'MAP', body)
    worksheet192.write('I22', 'IND', body)
    worksheet192.write('J22', 'ENG', body)
    worksheet192.write('J22', 'SEJ', body)
    worksheet192.write('K22', 'GEO', body)
    worksheet192.write('M22', 'EKO', body)
    worksheet192.write('L22', 'SOS', body)
    worksheet192.write('L22', 'FIS', body)
    worksheet192.write('L22', 'KIM', body)
    worksheet192.write('L22', 'BIO', body)
    worksheet192.write('N22', 'JML', body)
    worksheet192.write('O22', 'MAW', body)
    worksheet192.write('O22', 'MAP', body)
    worksheet192.write('P22', 'IND', body)
    worksheet192.write('Q22', 'ENG', body)
    worksheet192.write('R22', 'SEJ', body)
    worksheet192.write('S22', 'GEO', body)
    worksheet192.write('U22', 'EKO', body)
    worksheet192.write('T22', 'SOS', body)
    worksheet192.write('T22', 'FIS', body)
    worksheet192.write('T22', 'KIM', body)
    worksheet192.write('T22', 'BIO', body)
    worksheet192.write('V22', 'JML', body)

    worksheet192.conditional_format(22, 0, row192+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 193
    worksheet193.insert_image('A1', r'logo resmi nf.jpg')

    worksheet193.set_column('A:A', 7, center)
    worksheet193.set_column('B:B', 6, center)
    worksheet193.set_column('C:C', 18.14, center)
    worksheet193.set_column('D:D', 25, left)
    worksheet193.set_column('E:E', 13.14, left)
    worksheet193.set_column('F:F', 8.57, center)
    worksheet193.set_column('G:V', 5, center)
    worksheet193.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BINTARA', title)
    worksheet193.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet193.write('A5', 'LOKASI', header)
    worksheet193.write('B5', 'TOTAL', header)
    worksheet193.merge_range('A4:B4', 'RANK', header)
    worksheet193.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet193.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet193.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet193.merge_range('F4:F5', 'KELAS', header)
    worksheet193.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet193.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet193.write('G5', 'MAW', body)
    worksheet193.write('H5', 'MAP', body)
    worksheet193.write('I5', 'IND', body)
    worksheet193.write('J5', 'ENG', body)
    worksheet193.write('K5', 'SEJ', body)
    worksheet193.write('L5', 'GEO', body)
    worksheet193.write('M5', 'EKO', body)
    worksheet193.write('N5', 'SOS', body)
    worksheet193.write('O5', 'FIS', body)
    worksheet193.write('P5', 'KIM', body)
    worksheet193.write('Q5', 'BIO', body)
    worksheet193.write('R5', 'JML', body)
    worksheet193.write('S5', 'MAW', body)
    worksheet193.write('T5', 'MAP', body)
    worksheet193.write('U5', 'IND', body)
    worksheet193.write('V5', 'ENG', body)
    worksheet193.write('W5', 'SEJ', body)
    worksheet193.write('X5', 'GEO', body)
    worksheet193.write('Y5', 'EKO', body)
    worksheet193.write('Z5', 'SOS', body)
    worksheet193.write('AA5', 'FIS', body)
    worksheet193.write('AB5', 'KIM', body)
    worksheet193.write('AC5', 'BIO', body)
    worksheet193.write('AD5', 'JML', body)

    worksheet193.conditional_format(5, 0, row193_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet193.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF BINTARA', title)
    worksheet193.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet193.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet193.write('A22', 'LOKASI', header)
    worksheet193.write('B22', 'TOTAL', header)
    worksheet193.merge_range('A21:B21', 'RANK', header)
    worksheet193.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet193.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet193.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet193.merge_range('F21:F22', 'KELAS', header)
    worksheet193.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet193.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet193.write('G22', 'MAW', body)
    worksheet193.write('H22', 'MAP', body)
    worksheet193.write('I22', 'IND', body)
    worksheet193.write('J22', 'ENG', body)
    worksheet193.write('J22', 'SEJ', body)
    worksheet193.write('K22', 'GEO', body)
    worksheet193.write('M22', 'EKO', body)
    worksheet193.write('L22', 'SOS', body)
    worksheet193.write('L22', 'FIS', body)
    worksheet193.write('L22', 'KIM', body)
    worksheet193.write('L22', 'BIO', body)
    worksheet193.write('N22', 'JML', body)
    worksheet193.write('O22', 'MAW', body)
    worksheet193.write('O22', 'MAP', body)
    worksheet193.write('P22', 'IND', body)
    worksheet193.write('Q22', 'ENG', body)
    worksheet193.write('R22', 'SEJ', body)
    worksheet193.write('S22', 'GEO', body)
    worksheet193.write('U22', 'EKO', body)
    worksheet193.write('T22', 'SOS', body)
    worksheet193.write('T22', 'FIS', body)
    worksheet193.write('T22', 'KIM', body)
    worksheet193.write('T22', 'BIO', body)
    worksheet193.write('V22', 'JML', body)

    worksheet193.conditional_format(22, 0, row193+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 194
    worksheet194.insert_image('A1', r'logo resmi nf.jpg')

    worksheet194.set_column('A:A', 7, center)
    worksheet194.set_column('B:B', 6, center)
    worksheet194.set_column('C:C', 18.14, center)
    worksheet194.set_column('D:D', 25, left)
    worksheet194.set_column('E:E', 13.14, left)
    worksheet194.set_column('F:F', 8.57, center)
    worksheet194.set_column('G:V', 5, center)
    worksheet194.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MALANG', title)
    worksheet194.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet194.write('A5', 'LOKASI', header)
    worksheet194.write('B5', 'TOTAL', header)
    worksheet194.merge_range('A4:B4', 'RANK', header)
    worksheet194.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet194.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet194.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet194.merge_range('F4:F5', 'KELAS', header)
    worksheet194.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet194.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet194.write('G5', 'MAW', body)
    worksheet194.write('H5', 'MAP', body)
    worksheet194.write('I5', 'IND', body)
    worksheet194.write('J5', 'ENG', body)
    worksheet194.write('K5', 'SEJ', body)
    worksheet194.write('L5', 'GEO', body)
    worksheet194.write('M5', 'EKO', body)
    worksheet194.write('N5', 'SOS', body)
    worksheet194.write('O5', 'FIS', body)
    worksheet194.write('P5', 'KIM', body)
    worksheet194.write('Q5', 'BIO', body)
    worksheet194.write('R5', 'JML', body)
    worksheet194.write('S5', 'MAW', body)
    worksheet194.write('T5', 'MAP', body)
    worksheet194.write('U5', 'IND', body)
    worksheet194.write('V5', 'ENG', body)
    worksheet194.write('W5', 'SEJ', body)
    worksheet194.write('X5', 'GEO', body)
    worksheet194.write('Y5', 'EKO', body)
    worksheet194.write('Z5', 'SOS', body)
    worksheet194.write('AA5', 'FIS', body)
    worksheet194.write('AB5', 'KIM', body)
    worksheet194.write('AC5', 'BIO', body)
    worksheet194.write('AD5', 'JML', body)

    worksheet194.conditional_format(5, 0, row194_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet194.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF MALANG', title)
    worksheet194.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet194.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet194.write('A22', 'LOKASI', header)
    worksheet194.write('B22', 'TOTAL', header)
    worksheet194.merge_range('A21:B21', 'RANK', header)
    worksheet194.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet194.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet194.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet194.merge_range('F21:F22', 'KELAS', header)
    worksheet194.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet194.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet194.write('G22', 'MAW', body)
    worksheet194.write('H22', 'MAP', body)
    worksheet194.write('I22', 'IND', body)
    worksheet194.write('J22', 'ENG', body)
    worksheet194.write('J22', 'SEJ', body)
    worksheet194.write('K22', 'GEO', body)
    worksheet194.write('M22', 'EKO', body)
    worksheet194.write('L22', 'SOS', body)
    worksheet194.write('L22', 'FIS', body)
    worksheet194.write('L22', 'KIM', body)
    worksheet194.write('L22', 'BIO', body)
    worksheet194.write('N22', 'JML', body)
    worksheet194.write('O22', 'MAW', body)
    worksheet194.write('O22', 'MAP', body)
    worksheet194.write('P22', 'IND', body)
    worksheet194.write('Q22', 'ENG', body)
    worksheet194.write('R22', 'SEJ', body)
    worksheet194.write('S22', 'GEO', body)
    worksheet194.write('U22', 'EKO', body)
    worksheet194.write('T22', 'SOS', body)
    worksheet194.write('T22', 'FIS', body)
    worksheet194.write('T22', 'KIM', body)
    worksheet194.write('T22', 'BIO', body)
    worksheet194.write('V22', 'JML', body)

    worksheet194.conditional_format(22, 0, row194+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 195
    worksheet195.insert_image('A1', r'logo resmi nf.jpg')

    worksheet195.set_column('A:A', 7, center)
    worksheet195.set_column('B:B', 6, center)
    worksheet195.set_column('C:C', 18.14, center)
    worksheet195.set_column('D:D', 25, left)
    worksheet195.set_column('E:E', 13.14, left)
    worksheet195.set_column('F:F', 8.57, center)
    worksheet195.set_column('G:V', 5, center)
    worksheet195.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MEDAN BARU', title)
    worksheet195.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet195.write('A5', 'LOKASI', header)
    worksheet195.write('B5', 'TOTAL', header)
    worksheet195.merge_range('A4:B4', 'RANK', header)
    worksheet195.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet195.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet195.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet195.merge_range('F4:F5', 'KELAS', header)
    worksheet195.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet195.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet195.write('G5', 'MAW', body)
    worksheet195.write('H5', 'MAP', body)
    worksheet195.write('I5', 'IND', body)
    worksheet195.write('J5', 'ENG', body)
    worksheet195.write('K5', 'SEJ', body)
    worksheet195.write('L5', 'GEO', body)
    worksheet195.write('M5', 'EKO', body)
    worksheet195.write('N5', 'SOS', body)
    worksheet195.write('O5', 'FIS', body)
    worksheet195.write('P5', 'KIM', body)
    worksheet195.write('Q5', 'BIO', body)
    worksheet195.write('R5', 'JML', body)
    worksheet195.write('S5', 'MAW', body)
    worksheet195.write('T5', 'MAP', body)
    worksheet195.write('U5', 'IND', body)
    worksheet195.write('V5', 'ENG', body)
    worksheet195.write('W5', 'SEJ', body)
    worksheet195.write('X5', 'GEO', body)
    worksheet195.write('Y5', 'EKO', body)
    worksheet195.write('Z5', 'SOS', body)
    worksheet195.write('AA5', 'FIS', body)
    worksheet195.write('AB5', 'KIM', body)
    worksheet195.write('AC5', 'BIO', body)
    worksheet195.write('AD5', 'JML', body)

    worksheet195.conditional_format(5, 0, row195_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet195.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF MEDAN BARU', title)
    worksheet195.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet195.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet195.write('A22', 'LOKASI', header)
    worksheet195.write('B22', 'TOTAL', header)
    worksheet195.merge_range('A21:B21', 'RANK', header)
    worksheet195.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet195.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet195.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet195.merge_range('F21:F22', 'KELAS', header)
    worksheet195.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet195.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet195.write('G22', 'MAW', body)
    worksheet195.write('H22', 'MAP', body)
    worksheet195.write('I22', 'IND', body)
    worksheet195.write('J22', 'ENG', body)
    worksheet195.write('J22', 'SEJ', body)
    worksheet195.write('K22', 'GEO', body)
    worksheet195.write('M22', 'EKO', body)
    worksheet195.write('L22', 'SOS', body)
    worksheet195.write('L22', 'FIS', body)
    worksheet195.write('L22', 'KIM', body)
    worksheet195.write('L22', 'BIO', body)
    worksheet195.write('N22', 'JML', body)
    worksheet195.write('O22', 'MAW', body)
    worksheet195.write('O22', 'MAP', body)
    worksheet195.write('P22', 'IND', body)
    worksheet195.write('Q22', 'ENG', body)
    worksheet195.write('R22', 'SEJ', body)
    worksheet195.write('S22', 'GEO', body)
    worksheet195.write('U22', 'EKO', body)
    worksheet195.write('T22', 'SOS', body)
    worksheet195.write('T22', 'FIS', body)
    worksheet195.write('T22', 'KIM', body)
    worksheet195.write('T22', 'BIO', body)
    worksheet195.write('V22', 'JML', body)

    worksheet195.conditional_format(22, 0, row195+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 196
    worksheet196.insert_image('A1', r'logo resmi nf.jpg')

    worksheet196.set_column('A:A', 7, center)
    worksheet196.set_column('B:B', 6, center)
    worksheet196.set_column('C:C', 18.14, center)
    worksheet196.set_column('D:D', 25, left)
    worksheet196.set_column('E:E', 13.14, left)
    worksheet196.set_column('F:F', 8.57, center)
    worksheet196.set_column('G:V', 5, center)
    worksheet196.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MEDAN HELVETIA', title)
    worksheet196.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet196.write('A5', 'LOKASI', header)
    worksheet196.write('B5', 'TOTAL', header)
    worksheet196.merge_range('A4:B4', 'RANK', header)
    worksheet196.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet196.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet196.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet196.merge_range('F4:F5', 'KELAS', header)
    worksheet196.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet196.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet196.write('G5', 'MAW', body)
    worksheet196.write('H5', 'MAP', body)
    worksheet196.write('I5', 'IND', body)
    worksheet196.write('J5', 'ENG', body)
    worksheet196.write('K5', 'SEJ', body)
    worksheet196.write('L5', 'GEO', body)
    worksheet196.write('M5', 'EKO', body)
    worksheet196.write('N5', 'SOS', body)
    worksheet196.write('O5', 'FIS', body)
    worksheet196.write('P5', 'KIM', body)
    worksheet196.write('Q5', 'BIO', body)
    worksheet196.write('R5', 'JML', body)
    worksheet196.write('S5', 'MAW', body)
    worksheet196.write('T5', 'MAP', body)
    worksheet196.write('U5', 'IND', body)
    worksheet196.write('V5', 'ENG', body)
    worksheet196.write('W5', 'SEJ', body)
    worksheet196.write('X5', 'GEO', body)
    worksheet196.write('Y5', 'EKO', body)
    worksheet196.write('Z5', 'SOS', body)
    worksheet196.write('AA5', 'FIS', body)
    worksheet196.write('AB5', 'KIM', body)
    worksheet196.write('AC5', 'BIO', body)
    worksheet196.write('AD5', 'JML', body)

    worksheet196.conditional_format(5, 0, row196_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet196.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF MEDAN HELVETIA', title)
    worksheet196.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet196.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet196.write('A22', 'LOKASI', header)
    worksheet196.write('B22', 'TOTAL', header)
    worksheet196.merge_range('A21:B21', 'RANK', header)
    worksheet196.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet196.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet196.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet196.merge_range('F21:F22', 'KELAS', header)
    worksheet196.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet196.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet196.write('G22', 'MAW', body)
    worksheet196.write('H22', 'MAP', body)
    worksheet196.write('I22', 'IND', body)
    worksheet196.write('J22', 'ENG', body)
    worksheet196.write('J22', 'SEJ', body)
    worksheet196.write('K22', 'GEO', body)
    worksheet196.write('M22', 'EKO', body)
    worksheet196.write('L22', 'SOS', body)
    worksheet196.write('L22', 'FIS', body)
    worksheet196.write('L22', 'KIM', body)
    worksheet196.write('L22', 'BIO', body)
    worksheet196.write('N22', 'JML', body)
    worksheet196.write('O22', 'MAW', body)
    worksheet196.write('O22', 'MAP', body)
    worksheet196.write('P22', 'IND', body)
    worksheet196.write('Q22', 'ENG', body)
    worksheet196.write('R22', 'SEJ', body)
    worksheet196.write('S22', 'GEO', body)
    worksheet196.write('U22', 'EKO', body)
    worksheet196.write('T22', 'SOS', body)
    worksheet196.write('T22', 'FIS', body)
    worksheet196.write('T22', 'KIM', body)
    worksheet196.write('T22', 'BIO', body)
    worksheet196.write('V22', 'JML', body)

    worksheet196.conditional_format(22, 0, row196+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 197
    worksheet197.insert_image('A1', r'logo resmi nf.jpg')

    worksheet197.set_column('A:A', 7, center)
    worksheet197.set_column('B:B', 6, center)
    worksheet197.set_column('C:C', 18.14, center)
    worksheet197.set_column('D:D', 25, left)
    worksheet197.set_column('E:E', 13.14, left)
    worksheet197.set_column('F:F', 8.57, center)
    worksheet197.set_column('G:V', 5, center)
    worksheet197.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIHANJUANG', title)
    worksheet197.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet197.write('A5', 'LOKASI', header)
    worksheet197.write('B5', 'TOTAL', header)
    worksheet197.merge_range('A4:B4', 'RANK', header)
    worksheet197.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet197.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet197.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet197.merge_range('F4:F5', 'KELAS', header)
    worksheet197.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet197.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet197.write('G5', 'MAW', body)
    worksheet197.write('H5', 'MAP', body)
    worksheet197.write('I5', 'IND', body)
    worksheet197.write('J5', 'ENG', body)
    worksheet197.write('K5', 'SEJ', body)
    worksheet197.write('L5', 'GEO', body)
    worksheet197.write('M5', 'EKO', body)
    worksheet197.write('N5', 'SOS', body)
    worksheet197.write('O5', 'FIS', body)
    worksheet197.write('P5', 'KIM', body)
    worksheet197.write('Q5', 'BIO', body)
    worksheet197.write('R5', 'JML', body)
    worksheet197.write('S5', 'MAW', body)
    worksheet197.write('T5', 'MAP', body)
    worksheet197.write('U5', 'IND', body)
    worksheet197.write('V5', 'ENG', body)
    worksheet197.write('W5', 'SEJ', body)
    worksheet197.write('X5', 'GEO', body)
    worksheet197.write('Y5', 'EKO', body)
    worksheet197.write('Z5', 'SOS', body)
    worksheet197.write('AA5', 'FIS', body)
    worksheet197.write('AB5', 'KIM', body)
    worksheet197.write('AC5', 'BIO', body)
    worksheet197.write('AD5', 'JML', body)

    worksheet197.conditional_format(5, 0, row197_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet197.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF CIHANJUANG', title)
    worksheet197.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet197.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet197.write('A22', 'LOKASI', header)
    worksheet197.write('B22', 'TOTAL', header)
    worksheet197.merge_range('A21:B21', 'RANK', header)
    worksheet197.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet197.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet197.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet197.merge_range('F21:F22', 'KELAS', header)
    worksheet197.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet197.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet197.write('G22', 'MAW', body)
    worksheet197.write('H22', 'MAP', body)
    worksheet197.write('I22', 'IND', body)
    worksheet197.write('J22', 'ENG', body)
    worksheet197.write('J22', 'SEJ', body)
    worksheet197.write('K22', 'GEO', body)
    worksheet197.write('M22', 'EKO', body)
    worksheet197.write('L22', 'SOS', body)
    worksheet197.write('L22', 'FIS', body)
    worksheet197.write('L22', 'KIM', body)
    worksheet197.write('L22', 'BIO', body)
    worksheet197.write('N22', 'JML', body)
    worksheet197.write('O22', 'MAW', body)
    worksheet197.write('O22', 'MAP', body)
    worksheet197.write('P22', 'IND', body)
    worksheet197.write('Q22', 'ENG', body)
    worksheet197.write('R22', 'SEJ', body)
    worksheet197.write('S22', 'GEO', body)
    worksheet197.write('U22', 'EKO', body)
    worksheet197.write('T22', 'SOS', body)
    worksheet197.write('T22', 'FIS', body)
    worksheet197.write('T22', 'KIM', body)
    worksheet197.write('T22', 'BIO', body)
    worksheet197.write('V22', 'JML', body)

    worksheet197.conditional_format(22, 0, row197+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 198
    worksheet198.insert_image('A1', r'logo resmi nf.jpg')

    worksheet198.set_column('A:A', 7, center)
    worksheet198.set_column('B:B', 6, center)
    worksheet198.set_column('C:C', 18.14, center)
    worksheet198.set_column('D:D', 25, left)
    worksheet198.set_column('E:E', 13.14, left)
    worksheet198.set_column('F:F', 8.57, center)
    worksheet198.set_column('G:V', 5, center)
    worksheet198.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF BUAH BATU', title)
    worksheet198.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet198.write('A5', 'LOKASI', header)
    worksheet198.write('B5', 'TOTAL', header)
    worksheet198.merge_range('A4:B4', 'RANK', header)
    worksheet198.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet198.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet198.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet198.merge_range('F4:F5', 'KELAS', header)
    worksheet198.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet198.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet198.write('G5', 'MAW', body)
    worksheet198.write('H5', 'MAP', body)
    worksheet198.write('I5', 'IND', body)
    worksheet198.write('J5', 'ENG', body)
    worksheet198.write('K5', 'SEJ', body)
    worksheet198.write('L5', 'GEO', body)
    worksheet198.write('M5', 'EKO', body)
    worksheet198.write('N5', 'SOS', body)
    worksheet198.write('O5', 'FIS', body)
    worksheet198.write('P5', 'KIM', body)
    worksheet198.write('Q5', 'BIO', body)
    worksheet198.write('R5', 'JML', body)
    worksheet198.write('S5', 'MAW', body)
    worksheet198.write('T5', 'MAP', body)
    worksheet198.write('U5', 'IND', body)
    worksheet198.write('V5', 'ENG', body)
    worksheet198.write('W5', 'SEJ', body)
    worksheet198.write('X5', 'GEO', body)
    worksheet198.write('Y5', 'EKO', body)
    worksheet198.write('Z5', 'SOS', body)
    worksheet198.write('AA5', 'FIS', body)
    worksheet198.write('AB5', 'KIM', body)
    worksheet198.write('AC5', 'BIO', body)
    worksheet198.write('AD5', 'JML', body)

    worksheet198.conditional_format(5, 0, row198_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet198.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF BUAH BATU', title)
    worksheet198.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet198.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet198.write('A22', 'LOKASI', header)
    worksheet198.write('B22', 'TOTAL', header)
    worksheet198.merge_range('A21:B21', 'RANK', header)
    worksheet198.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet198.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet198.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet198.merge_range('F21:F22', 'KELAS', header)
    worksheet198.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet198.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet198.write('G22', 'MAW', body)
    worksheet198.write('H22', 'MAP', body)
    worksheet198.write('I22', 'IND', body)
    worksheet198.write('J22', 'ENG', body)
    worksheet198.write('J22', 'SEJ', body)
    worksheet198.write('K22', 'GEO', body)
    worksheet198.write('M22', 'EKO', body)
    worksheet198.write('L22', 'SOS', body)
    worksheet198.write('L22', 'FIS', body)
    worksheet198.write('L22', 'KIM', body)
    worksheet198.write('L22', 'BIO', body)
    worksheet198.write('N22', 'JML', body)
    worksheet198.write('O22', 'MAW', body)
    worksheet198.write('O22', 'MAP', body)
    worksheet198.write('P22', 'IND', body)
    worksheet198.write('Q22', 'ENG', body)
    worksheet198.write('R22', 'SEJ', body)
    worksheet198.write('S22', 'GEO', body)
    worksheet198.write('U22', 'EKO', body)
    worksheet198.write('T22', 'SOS', body)
    worksheet198.write('T22', 'FIS', body)
    worksheet198.write('T22', 'KIM', body)
    worksheet198.write('T22', 'BIO', body)
    worksheet198.write('V22', 'JML', body)

    worksheet198.conditional_format(22, 0, row198+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 199
    worksheet199.insert_image('A1', r'logo resmi nf.jpg')

    worksheet199.set_column('A:A', 7, center)
    worksheet199.set_column('B:B', 6, center)
    worksheet199.set_column('C:C', 18.14, center)
    worksheet199.set_column('D:D', 25, left)
    worksheet199.set_column('E:E', 13.14, left)
    worksheet199.set_column('F:F', 8.57, center)
    worksheet199.set_column('G:V', 5, center)
    worksheet199.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SUMBAWA', title)
    worksheet199.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet199.write('A5', 'LOKASI', header)
    worksheet199.write('B5', 'TOTAL', header)
    worksheet199.merge_range('A4:B4', 'RANK', header)
    worksheet199.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet199.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet199.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet199.merge_range('F4:F5', 'KELAS', header)
    worksheet199.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet199.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet199.write('G5', 'MAW', body)
    worksheet199.write('H5', 'MAP', body)
    worksheet199.write('I5', 'IND', body)
    worksheet199.write('J5', 'ENG', body)
    worksheet199.write('K5', 'SEJ', body)
    worksheet199.write('L5', 'GEO', body)
    worksheet199.write('M5', 'EKO', body)
    worksheet199.write('N5', 'SOS', body)
    worksheet199.write('O5', 'FIS', body)
    worksheet199.write('P5', 'KIM', body)
    worksheet199.write('Q5', 'BIO', body)
    worksheet199.write('R5', 'JML', body)
    worksheet199.write('S5', 'MAW', body)
    worksheet199.write('T5', 'MAP', body)
    worksheet199.write('U5', 'IND', body)
    worksheet199.write('V5', 'ENG', body)
    worksheet199.write('W5', 'SEJ', body)
    worksheet199.write('X5', 'GEO', body)
    worksheet199.write('Y5', 'EKO', body)
    worksheet199.write('Z5', 'SOS', body)
    worksheet199.write('AA5', 'FIS', body)
    worksheet199.write('AB5', 'KIM', body)
    worksheet199.write('AC5', 'BIO', body)
    worksheet199.write('AD5', 'JML', body)

    worksheet199.conditional_format(5, 0, row199_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet199.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF SUMBAWA', title)
    worksheet199.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet199.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet199.write('A22', 'LOKASI', header)
    worksheet199.write('B22', 'TOTAL', header)
    worksheet199.merge_range('A21:B21', 'RANK', header)
    worksheet199.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet199.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet199.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet199.merge_range('F21:F22', 'KELAS', header)
    worksheet199.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet199.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet199.write('G22', 'MAW', body)
    worksheet199.write('H22', 'MAP', body)
    worksheet199.write('I22', 'IND', body)
    worksheet199.write('J22', 'ENG', body)
    worksheet199.write('J22', 'SEJ', body)
    worksheet199.write('K22', 'GEO', body)
    worksheet199.write('M22', 'EKO', body)
    worksheet199.write('L22', 'SOS', body)
    worksheet199.write('L22', 'FIS', body)
    worksheet199.write('L22', 'KIM', body)
    worksheet199.write('L22', 'BIO', body)
    worksheet199.write('N22', 'JML', body)
    worksheet199.write('O22', 'MAW', body)
    worksheet199.write('O22', 'MAP', body)
    worksheet199.write('P22', 'IND', body)
    worksheet199.write('Q22', 'ENG', body)
    worksheet199.write('R22', 'SEJ', body)
    worksheet199.write('S22', 'GEO', body)
    worksheet199.write('U22', 'EKO', body)
    worksheet199.write('T22', 'SOS', body)
    worksheet199.write('T22', 'FIS', body)
    worksheet199.write('T22', 'KIM', body)
    worksheet199.write('T22', 'BIO', body)
    worksheet199.write('V22', 'JML', body)

    worksheet199.conditional_format(22, 0, row199+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 201
    worksheet201.insert_image('A1', r'logo resmi nf.jpg')

    worksheet201.set_column('A:A', 7, center)
    worksheet201.set_column('B:B', 6, center)
    worksheet201.set_column('C:C', 18.14, center)
    worksheet201.set_column('D:D', 25, left)
    worksheet201.set_column('E:E', 13.14, left)
    worksheet201.set_column('F:F', 8.57, center)
    worksheet201.set_column('G:V', 5, center)
    worksheet201.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF UJUNG BERUNG', title)
    worksheet201.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet201.write('A5', 'LOKASI', header)
    worksheet201.write('B5', 'TOTAL', header)
    worksheet201.merge_range('A4:B4', 'RANK', header)
    worksheet201.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet201.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet201.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet201.merge_range('F4:F5', 'KELAS', header)
    worksheet201.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet201.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet201.write('G5', 'MAW', body)
    worksheet201.write('H5', 'MAP', body)
    worksheet201.write('I5', 'IND', body)
    worksheet201.write('J5', 'ENG', body)
    worksheet201.write('K5', 'SEJ', body)
    worksheet201.write('L5', 'GEO', body)
    worksheet201.write('M5', 'EKO', body)
    worksheet201.write('N5', 'SOS', body)
    worksheet201.write('O5', 'FIS', body)
    worksheet201.write('P5', 'KIM', body)
    worksheet201.write('Q5', 'BIO', body)
    worksheet201.write('R5', 'JML', body)
    worksheet201.write('S5', 'MAW', body)
    worksheet201.write('T5', 'MAP', body)
    worksheet201.write('U5', 'IND', body)
    worksheet201.write('V5', 'ENG', body)
    worksheet201.write('W5', 'SEJ', body)
    worksheet201.write('X5', 'GEO', body)
    worksheet201.write('Y5', 'EKO', body)
    worksheet201.write('Z5', 'SOS', body)
    worksheet201.write('AA5', 'FIS', body)
    worksheet201.write('AB5', 'KIM', body)
    worksheet201.write('AC5', 'BIO', body)
    worksheet201.write('AD5', 'JML', body)

    worksheet201.conditional_format(5, 0, row201_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet201.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF UJUNG BERUNG', title)
    worksheet201.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet201.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet201.write('A22', 'LOKASI', header)
    worksheet201.write('B22', 'TOTAL', header)
    worksheet201.merge_range('A21:B21', 'RANK', header)
    worksheet201.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet201.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet201.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet201.merge_range('F21:F22', 'KELAS', header)
    worksheet201.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet201.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet201.write('G22', 'MAW', body)
    worksheet201.write('H22', 'MAP', body)
    worksheet201.write('I22', 'IND', body)
    worksheet201.write('J22', 'ENG', body)
    worksheet201.write('J22', 'SEJ', body)
    worksheet201.write('K22', 'GEO', body)
    worksheet201.write('M22', 'EKO', body)
    worksheet201.write('L22', 'SOS', body)
    worksheet201.write('L22', 'FIS', body)
    worksheet201.write('L22', 'KIM', body)
    worksheet201.write('L22', 'BIO', body)
    worksheet201.write('N22', 'JML', body)
    worksheet201.write('O22', 'MAW', body)
    worksheet201.write('O22', 'MAP', body)
    worksheet201.write('P22', 'IND', body)
    worksheet201.write('Q22', 'ENG', body)
    worksheet201.write('R22', 'SEJ', body)
    worksheet201.write('S22', 'GEO', body)
    worksheet201.write('U22', 'EKO', body)
    worksheet201.write('T22', 'SOS', body)
    worksheet201.write('T22', 'FIS', body)
    worksheet201.write('T22', 'KIM', body)
    worksheet201.write('T22', 'BIO', body)
    worksheet201.write('V22', 'JML', body)

    worksheet201.conditional_format(22, 0, row201+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 202
    worksheet202.insert_image('A1', r'logo resmi nf.jpg')

    worksheet202.set_column('A:A', 7, center)
    worksheet202.set_column('B:B', 6, center)
    worksheet202.set_column('C:C', 18.14, center)
    worksheet202.set_column('D:D', 25, left)
    worksheet202.set_column('E:E', 13.14, left)
    worksheet202.set_column('F:F', 8.57, center)
    worksheet202.set_column('G:V', 5, center)
    worksheet202.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SANGKURIANG', title)
    worksheet202.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet202.write('A5', 'LOKASI', header)
    worksheet202.write('B5', 'TOTAL', header)
    worksheet202.merge_range('A4:B4', 'RANK', header)
    worksheet202.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet202.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet202.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet202.merge_range('F4:F5', 'KELAS', header)
    worksheet202.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet202.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet202.write('G5', 'MAW', body)
    worksheet202.write('H5', 'MAP', body)
    worksheet202.write('I5', 'IND', body)
    worksheet202.write('J5', 'ENG', body)
    worksheet202.write('K5', 'SEJ', body)
    worksheet202.write('L5', 'GEO', body)
    worksheet202.write('M5', 'EKO', body)
    worksheet202.write('N5', 'SOS', body)
    worksheet202.write('O5', 'FIS', body)
    worksheet202.write('P5', 'KIM', body)
    worksheet202.write('Q5', 'BIO', body)
    worksheet202.write('R5', 'JML', body)
    worksheet202.write('S5', 'MAW', body)
    worksheet202.write('T5', 'MAP', body)
    worksheet202.write('U5', 'IND', body)
    worksheet202.write('V5', 'ENG', body)
    worksheet202.write('W5', 'SEJ', body)
    worksheet202.write('X5', 'GEO', body)
    worksheet202.write('Y5', 'EKO', body)
    worksheet202.write('Z5', 'SOS', body)
    worksheet202.write('AA5', 'FIS', body)
    worksheet202.write('AB5', 'KIM', body)
    worksheet202.write('AC5', 'BIO', body)
    worksheet202.write('AD5', 'JML', body)

    worksheet202.conditional_format(5, 0, row202_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet202.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF SANGKURIANG', title)
    worksheet202.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet202.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet202.write('A22', 'LOKASI', header)
    worksheet202.write('B22', 'TOTAL', header)
    worksheet202.merge_range('A21:B21', 'RANK', header)
    worksheet202.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet202.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet202.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet202.merge_range('F21:F22', 'KELAS', header)
    worksheet202.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet202.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet202.write('G22', 'MAW', body)
    worksheet202.write('H22', 'MAP', body)
    worksheet202.write('I22', 'IND', body)
    worksheet202.write('J22', 'ENG', body)
    worksheet202.write('J22', 'SEJ', body)
    worksheet202.write('K22', 'GEO', body)
    worksheet202.write('M22', 'EKO', body)
    worksheet202.write('L22', 'SOS', body)
    worksheet202.write('L22', 'FIS', body)
    worksheet202.write('L22', 'KIM', body)
    worksheet202.write('L22', 'BIO', body)
    worksheet202.write('N22', 'JML', body)
    worksheet202.write('O22', 'MAW', body)
    worksheet202.write('O22', 'MAP', body)
    worksheet202.write('P22', 'IND', body)
    worksheet202.write('Q22', 'ENG', body)
    worksheet202.write('R22', 'SEJ', body)
    worksheet202.write('S22', 'GEO', body)
    worksheet202.write('U22', 'EKO', body)
    worksheet202.write('T22', 'SOS', body)
    worksheet202.write('T22', 'FIS', body)
    worksheet202.write('T22', 'KIM', body)
    worksheet202.write('T22', 'BIO', body)
    worksheet202.write('V22', 'JML', body)

    worksheet202.conditional_format(22, 0, row202+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # # worksheet 139
    # worksheet139.insert_image('A1',r'logo resmi nf.jpg')

    # worksheet139.set_column('A:A', 7, center)
    # worksheet139.set_column('B:B', 6, center)
    # worksheet139.set_column('C:C', 18.14, center)
    # worksheet139.set_column('D:D', 25, left)
    # worksheet139.set_column('E:E', 13.14, left)
    # worksheet139.set_column('F:F', 8.57, center)
    # worksheet139.set_column('G:V', 5, center)
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

    # worksheet 203
    worksheet203.insert_image('A1', r'logo resmi nf.jpg')

    worksheet203.set_column('A:A', 7, center)
    worksheet203.set_column('B:B', 6, center)
    worksheet203.set_column('C:C', 18.14, center)
    worksheet203.set_column('D:D', 25, left)
    worksheet203.set_column('E:E', 13.14, left)
    worksheet203.set_column('F:F', 8.57, center)
    worksheet203.set_column('G:V', 5, center)
    worksheet203.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KARAWACI', title)
    worksheet203.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet203.write('A5', 'LOKASI', header)
    worksheet203.write('B5', 'TOTAL', header)
    worksheet203.merge_range('A4:B4', 'RANK', header)
    worksheet203.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet203.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet203.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet203.merge_range('F4:F5', 'KELAS', header)
    worksheet203.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet203.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet203.write('G5', 'MAW', body)
    worksheet203.write('H5', 'MAP', body)
    worksheet203.write('I5', 'IND', body)
    worksheet203.write('J5', 'ENG', body)
    worksheet203.write('K5', 'SEJ', body)
    worksheet203.write('L5', 'GEO', body)
    worksheet203.write('M5', 'EKO', body)
    worksheet203.write('N5', 'SOS', body)
    worksheet203.write('O5', 'FIS', body)
    worksheet203.write('P5', 'KIM', body)
    worksheet203.write('Q5', 'BIO', body)
    worksheet203.write('R5', 'JML', body)
    worksheet203.write('S5', 'MAW', body)
    worksheet203.write('T5', 'MAP', body)
    worksheet203.write('U5', 'IND', body)
    worksheet203.write('V5', 'ENG', body)
    worksheet203.write('W5', 'SEJ', body)
    worksheet203.write('X5', 'GEO', body)
    worksheet203.write('Y5', 'EKO', body)
    worksheet203.write('Z5', 'SOS', body)
    worksheet203.write('AA5', 'FIS', body)
    worksheet203.write('AB5', 'KIM', body)
    worksheet203.write('AC5', 'BIO', body)
    worksheet203.write('AD5', 'JML', body)

    worksheet203.conditional_format(5, 0, row203_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet203.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF KARAWACI', title)
    worksheet203.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet203.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet203.write('A22', 'LOKASI', header)
    worksheet203.write('B22', 'TOTAL', header)
    worksheet203.merge_range('A21:B21', 'RANK', header)
    worksheet203.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet203.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet203.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet203.merge_range('F21:F22', 'KELAS', header)
    worksheet203.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet203.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet203.write('G22', 'MAW', body)
    worksheet203.write('H22', 'MAP', body)
    worksheet203.write('I22', 'IND', body)
    worksheet203.write('J22', 'ENG', body)
    worksheet203.write('J22', 'SEJ', body)
    worksheet203.write('K22', 'GEO', body)
    worksheet203.write('M22', 'EKO', body)
    worksheet203.write('L22', 'SOS', body)
    worksheet203.write('L22', 'FIS', body)
    worksheet203.write('L22', 'KIM', body)
    worksheet203.write('L22', 'BIO', body)
    worksheet203.write('N22', 'JML', body)
    worksheet203.write('O22', 'MAW', body)
    worksheet203.write('O22', 'MAP', body)
    worksheet203.write('P22', 'IND', body)
    worksheet203.write('Q22', 'ENG', body)
    worksheet203.write('R22', 'SEJ', body)
    worksheet203.write('S22', 'GEO', body)
    worksheet203.write('U22', 'EKO', body)
    worksheet203.write('T22', 'SOS', body)
    worksheet203.write('T22', 'FIS', body)
    worksheet203.write('T22', 'KIM', body)
    worksheet203.write('T22', 'BIO', body)
    worksheet203.write('V22', 'JML', body)

    worksheet203.conditional_format(22, 0, row203+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 210
    worksheet210.insert_image('A1', r'logo resmi nf.jpg')

    worksheet210.set_column('A:A', 7, center)
    worksheet210.set_column('B:B', 6, center)
    worksheet210.set_column('C:C', 18.14, center)
    worksheet210.set_column('D:D', 25, left)
    worksheet210.set_column('E:E', 13.14, left)
    worksheet210.set_column('F:F', 8.57, center)
    worksheet210.set_column('G:V', 5, center)
    worksheet210.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF VETERAN TANGERANG', title)
    worksheet210.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet210.write('A5', 'LOKASI', header)
    worksheet210.write('B5', 'TOTAL', header)
    worksheet210.merge_range('A4:B4', 'RANK', header)
    worksheet210.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet210.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet210.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet210.merge_range('F4:F5', 'KELAS', header)
    worksheet210.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet210.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet210.write('G5', 'MAW', body)
    worksheet210.write('H5', 'MAP', body)
    worksheet210.write('I5', 'IND', body)
    worksheet210.write('J5', 'ENG', body)
    worksheet210.write('K5', 'SEJ', body)
    worksheet210.write('L5', 'GEO', body)
    worksheet210.write('M5', 'EKO', body)
    worksheet210.write('N5', 'SOS', body)
    worksheet210.write('O5', 'FIS', body)
    worksheet210.write('P5', 'KIM', body)
    worksheet210.write('Q5', 'BIO', body)
    worksheet210.write('R5', 'JML', body)
    worksheet210.write('S5', 'MAW', body)
    worksheet210.write('T5', 'MAP', body)
    worksheet210.write('U5', 'IND', body)
    worksheet210.write('V5', 'ENG', body)
    worksheet210.write('W5', 'SEJ', body)
    worksheet210.write('X5', 'GEO', body)
    worksheet210.write('Y5', 'EKO', body)
    worksheet210.write('Z5', 'SOS', body)
    worksheet210.write('AA5', 'FIS', body)
    worksheet210.write('AB5', 'KIM', body)
    worksheet210.write('AC5', 'BIO', body)
    worksheet210.write('AD5', 'JML', body)

    worksheet210.conditional_format(5, 0, row210_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet210.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF VETERAN TANGERANG', title)
    worksheet210.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet210.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet210.write('A22', 'LOKASI', header)
    worksheet210.write('B22', 'TOTAL', header)
    worksheet210.merge_range('A21:B21', 'RANK', header)
    worksheet210.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet210.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet210.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet210.merge_range('F21:F22', 'KELAS', header)
    worksheet210.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet210.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet210.write('G22', 'MAW', body)
    worksheet210.write('H22', 'MAP', body)
    worksheet210.write('I22', 'IND', body)
    worksheet210.write('J22', 'ENG', body)
    worksheet210.write('J22', 'SEJ', body)
    worksheet210.write('K22', 'GEO', body)
    worksheet210.write('M22', 'EKO', body)
    worksheet210.write('L22', 'SOS', body)
    worksheet210.write('L22', 'FIS', body)
    worksheet210.write('L22', 'KIM', body)
    worksheet210.write('L22', 'BIO', body)
    worksheet210.write('N22', 'JML', body)
    worksheet210.write('O22', 'MAW', body)
    worksheet210.write('O22', 'MAP', body)
    worksheet210.write('P22', 'IND', body)
    worksheet210.write('Q22', 'ENG', body)
    worksheet210.write('R22', 'SEJ', body)
    worksheet210.write('S22', 'GEO', body)
    worksheet210.write('U22', 'EKO', body)
    worksheet210.write('T22', 'SOS', body)
    worksheet210.write('T22', 'FIS', body)
    worksheet210.write('T22', 'KIM', body)
    worksheet210.write('T22', 'BIO', body)
    worksheet210.write('V22', 'JML', body)

    worksheet210.conditional_format(22, 0, row210+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 211
    worksheet211.insert_image('A1', r'logo resmi nf.jpg')

    worksheet211.set_column('A:A', 7, center)
    worksheet211.set_column('B:B', 6, center)
    worksheet211.set_column('C:C', 18.14, center)
    worksheet211.set_column('D:D', 25, left)
    worksheet211.set_column('E:E', 13.14, left)
    worksheet211.set_column('F:F', 8.57, center)
    worksheet211.set_column('G:V', 5, center)
    worksheet211.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PERUMNAS 2 TANGERANG', title)
    worksheet211.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet211.write('A5', 'LOKASI', header)
    worksheet211.write('B5', 'TOTAL', header)
    worksheet211.merge_range('A4:B4', 'RANK', header)
    worksheet211.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet211.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet211.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet211.merge_range('F4:F5', 'KELAS', header)
    worksheet211.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet211.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet211.write('G5', 'MAW', body)
    worksheet211.write('H5', 'MAP', body)
    worksheet211.write('I5', 'IND', body)
    worksheet211.write('J5', 'ENG', body)
    worksheet211.write('K5', 'SEJ', body)
    worksheet211.write('L5', 'GEO', body)
    worksheet211.write('M5', 'EKO', body)
    worksheet211.write('N5', 'SOS', body)
    worksheet211.write('O5', 'FIS', body)
    worksheet211.write('P5', 'KIM', body)
    worksheet211.write('Q5', 'BIO', body)
    worksheet211.write('R5', 'JML', body)
    worksheet211.write('S5', 'MAW', body)
    worksheet211.write('T5', 'MAP', body)
    worksheet211.write('U5', 'IND', body)
    worksheet211.write('V5', 'ENG', body)
    worksheet211.write('W5', 'SEJ', body)
    worksheet211.write('X5', 'GEO', body)
    worksheet211.write('Y5', 'EKO', body)
    worksheet211.write('Z5', 'SOS', body)
    worksheet211.write('AA5', 'FIS', body)
    worksheet211.write('AB5', 'KIM', body)
    worksheet211.write('AC5', 'BIO', body)
    worksheet211.write('AD5', 'JML', body)

    worksheet211.conditional_format(5, 0, row211_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet211.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF PERUMNAS 2 TANGERANG', title)
    worksheet211.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet211.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet211.write('A22', 'LOKASI', header)
    worksheet211.write('B22', 'TOTAL', header)
    worksheet211.merge_range('A21:B21', 'RANK', header)
    worksheet211.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet211.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet211.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet211.merge_range('F21:F22', 'KELAS', header)
    worksheet211.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet211.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet211.write('G22', 'MAW', body)
    worksheet211.write('H22', 'MAP', body)
    worksheet211.write('I22', 'IND', body)
    worksheet211.write('J22', 'ENG', body)
    worksheet211.write('J22', 'SEJ', body)
    worksheet211.write('K22', 'GEO', body)
    worksheet211.write('M22', 'EKO', body)
    worksheet211.write('L22', 'SOS', body)
    worksheet211.write('L22', 'FIS', body)
    worksheet211.write('L22', 'KIM', body)
    worksheet211.write('L22', 'BIO', body)
    worksheet211.write('N22', 'JML', body)
    worksheet211.write('O22', 'MAW', body)
    worksheet211.write('O22', 'MAP', body)
    worksheet211.write('P22', 'IND', body)
    worksheet211.write('Q22', 'ENG', body)
    worksheet211.write('R22', 'SEJ', body)
    worksheet211.write('S22', 'GEO', body)
    worksheet211.write('U22', 'EKO', body)
    worksheet211.write('T22', 'SOS', body)
    worksheet211.write('T22', 'FIS', body)
    worksheet211.write('T22', 'KIM', body)
    worksheet211.write('T22', 'BIO', body)
    worksheet211.write('V22', 'JML', body)

    worksheet211.conditional_format(22, 0, row211+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 212
    worksheet212.insert_image('A1', r'logo resmi nf.jpg')

    worksheet212.set_column('A:A', 7, center)
    worksheet212.set_column('B:B', 6, center)
    worksheet212.set_column('C:C', 18.14, center)
    worksheet212.set_column('D:D', 25, left)
    worksheet212.set_column('E:E', 13.14, left)
    worksheet212.set_column('F:F', 8.57, center)
    worksheet212.set_column('G:V', 5, center)
    worksheet212.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KAYURINGIN', title)
    worksheet212.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet212.write('A5', 'LOKASI', header)
    worksheet212.write('B5', 'TOTAL', header)
    worksheet212.merge_range('A4:B4', 'RANK', header)
    worksheet212.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet212.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet212.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet212.merge_range('F4:F5', 'KELAS', header)
    worksheet212.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet212.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet212.write('G5', 'MAW', body)
    worksheet212.write('H5', 'MAP', body)
    worksheet212.write('I5', 'IND', body)
    worksheet212.write('J5', 'ENG', body)
    worksheet212.write('K5', 'SEJ', body)
    worksheet212.write('L5', 'GEO', body)
    worksheet212.write('M5', 'EKO', body)
    worksheet212.write('N5', 'SOS', body)
    worksheet212.write('O5', 'FIS', body)
    worksheet212.write('P5', 'KIM', body)
    worksheet212.write('Q5', 'BIO', body)
    worksheet212.write('R5', 'JML', body)
    worksheet212.write('S5', 'MAW', body)
    worksheet212.write('T5', 'MAP', body)
    worksheet212.write('U5', 'IND', body)
    worksheet212.write('V5', 'ENG', body)
    worksheet212.write('W5', 'SEJ', body)
    worksheet212.write('X5', 'GEO', body)
    worksheet212.write('Y5', 'EKO', body)
    worksheet212.write('Z5', 'SOS', body)
    worksheet212.write('AA5', 'FIS', body)
    worksheet212.write('AB5', 'KIM', body)
    worksheet212.write('AC5', 'BIO', body)
    worksheet212.write('AD5', 'JML', body)

    worksheet212.conditional_format(5, 0, row212_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet212.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF KAYURINGIN', title)
    worksheet212.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet212.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet212.write('A22', 'LOKASI', header)
    worksheet212.write('B22', 'TOTAL', header)
    worksheet212.merge_range('A21:B21', 'RANK', header)
    worksheet212.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet212.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet212.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet212.merge_range('F21:F22', 'KELAS', header)
    worksheet212.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet212.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet212.write('G22', 'MAW', body)
    worksheet212.write('H22', 'MAP', body)
    worksheet212.write('I22', 'IND', body)
    worksheet212.write('J22', 'ENG', body)
    worksheet212.write('J22', 'SEJ', body)
    worksheet212.write('K22', 'GEO', body)
    worksheet212.write('M22', 'EKO', body)
    worksheet212.write('L22', 'SOS', body)
    worksheet212.write('L22', 'FIS', body)
    worksheet212.write('L22', 'KIM', body)
    worksheet212.write('L22', 'BIO', body)
    worksheet212.write('N22', 'JML', body)
    worksheet212.write('O22', 'MAW', body)
    worksheet212.write('O22', 'MAP', body)
    worksheet212.write('P22', 'IND', body)
    worksheet212.write('Q22', 'ENG', body)
    worksheet212.write('R22', 'SEJ', body)
    worksheet212.write('S22', 'GEO', body)
    worksheet212.write('U22', 'EKO', body)
    worksheet212.write('T22', 'SOS', body)
    worksheet212.write('T22', 'FIS', body)
    worksheet212.write('T22', 'KIM', body)
    worksheet212.write('T22', 'BIO', body)
    worksheet212.write('V22', 'JML', body)

    worksheet212.conditional_format(22, 0, row212+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 216
    worksheet216.insert_image('A1', r'logo resmi nf.jpg')

    worksheet216.set_column('A:A', 7, center)
    worksheet216.set_column('B:B', 6, center)
    worksheet216.set_column('C:C', 18.14, center)
    worksheet216.set_column('D:D', 25, left)
    worksheet216.set_column('E:E', 13.14, left)
    worksheet216.set_column('F:F', 8.57, center)
    worksheet216.set_column('G:V', 5, center)
    worksheet216.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF AGUS SALIM', title)
    worksheet216.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet216.write('A5', 'LOKASI', header)
    worksheet216.write('B5', 'TOTAL', header)
    worksheet216.merge_range('A4:B4', 'RANK', header)
    worksheet216.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet216.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet216.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet216.merge_range('F4:F5', 'KELAS', header)
    worksheet216.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet216.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet216.write('G5', 'MAW', body)
    worksheet216.write('H5', 'MAP', body)
    worksheet216.write('I5', 'IND', body)
    worksheet216.write('J5', 'ENG', body)
    worksheet216.write('K5', 'SEJ', body)
    worksheet216.write('L5', 'GEO', body)
    worksheet216.write('M5', 'EKO', body)
    worksheet216.write('N5', 'SOS', body)
    worksheet216.write('O5', 'FIS', body)
    worksheet216.write('P5', 'KIM', body)
    worksheet216.write('Q5', 'BIO', body)
    worksheet216.write('R5', 'JML', body)
    worksheet216.write('S5', 'MAW', body)
    worksheet216.write('T5', 'MAP', body)
    worksheet216.write('U5', 'IND', body)
    worksheet216.write('V5', 'ENG', body)
    worksheet216.write('W5', 'SEJ', body)
    worksheet216.write('X5', 'GEO', body)
    worksheet216.write('Y5', 'EKO', body)
    worksheet216.write('Z5', 'SOS', body)
    worksheet216.write('AA5', 'FIS', body)
    worksheet216.write('AB5', 'KIM', body)
    worksheet216.write('AC5', 'BIO', body)
    worksheet216.write('AD5', 'JML', body)

    worksheet216.conditional_format(5, 0, row216_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet216.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF AGUS SALIM', title)
    worksheet216.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet216.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet216.write('A22', 'LOKASI', header)
    worksheet216.write('B22', 'TOTAL', header)
    worksheet216.merge_range('A21:B21', 'RANK', header)
    worksheet216.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet216.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet216.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet216.merge_range('F21:F22', 'KELAS', header)
    worksheet216.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet216.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet216.write('G22', 'MAW', body)
    worksheet216.write('H22', 'MAP', body)
    worksheet216.write('I22', 'IND', body)
    worksheet216.write('J22', 'ENG', body)
    worksheet216.write('J22', 'SEJ', body)
    worksheet216.write('K22', 'GEO', body)
    worksheet216.write('M22', 'EKO', body)
    worksheet216.write('L22', 'SOS', body)
    worksheet216.write('L22', 'FIS', body)
    worksheet216.write('L22', 'KIM', body)
    worksheet216.write('L22', 'BIO', body)
    worksheet216.write('N22', 'JML', body)
    worksheet216.write('O22', 'MAW', body)
    worksheet216.write('O22', 'MAP', body)
    worksheet216.write('P22', 'IND', body)
    worksheet216.write('Q22', 'ENG', body)
    worksheet216.write('R22', 'SEJ', body)
    worksheet216.write('S22', 'GEO', body)
    worksheet216.write('U22', 'EKO', body)
    worksheet216.write('T22', 'SOS', body)
    worksheet216.write('T22', 'FIS', body)
    worksheet216.write('T22', 'KIM', body)
    worksheet216.write('T22', 'BIO', body)
    worksheet216.write('V22', 'JML', body)

    worksheet216.conditional_format(22, 0, row216+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 217
    worksheet217.insert_image('A1', r'logo resmi nf.jpg')

    worksheet217.set_column('A:A', 7, center)
    worksheet217.set_column('B:B', 6, center)
    worksheet217.set_column('C:C', 18.14, center)
    worksheet217.set_column('D:D', 25, left)
    worksheet217.set_column('E:E', 13.14, left)
    worksheet217.set_column('F:F', 8.57, center)
    worksheet217.set_column('G:V', 5, center)
    worksheet217.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF SUMERU', title)
    worksheet217.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet217.write('A5', 'LOKASI', header)
    worksheet217.write('B5', 'TOTAL', header)
    worksheet217.merge_range('A4:B4', 'RANK', header)
    worksheet217.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet217.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet217.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet217.merge_range('F4:F5', 'KELAS', header)
    worksheet217.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet217.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet217.write('G5', 'MAW', body)
    worksheet217.write('H5', 'MAP', body)
    worksheet217.write('I5', 'IND', body)
    worksheet217.write('J5', 'ENG', body)
    worksheet217.write('K5', 'SEJ', body)
    worksheet217.write('L5', 'GEO', body)
    worksheet217.write('M5', 'EKO', body)
    worksheet217.write('N5', 'SOS', body)
    worksheet217.write('O5', 'FIS', body)
    worksheet217.write('P5', 'KIM', body)
    worksheet217.write('Q5', 'BIO', body)
    worksheet217.write('R5', 'JML', body)
    worksheet217.write('S5', 'MAW', body)
    worksheet217.write('T5', 'MAP', body)
    worksheet217.write('U5', 'IND', body)
    worksheet217.write('V5', 'ENG', body)
    worksheet217.write('W5', 'SEJ', body)
    worksheet217.write('X5', 'GEO', body)
    worksheet217.write('Y5', 'EKO', body)
    worksheet217.write('Z5', 'SOS', body)
    worksheet217.write('AA5', 'FIS', body)
    worksheet217.write('AB5', 'KIM', body)
    worksheet217.write('AC5', 'BIO', body)
    worksheet217.write('AD5', 'JML', body)

    worksheet217.conditional_format(5, 0, row217_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet217.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF SUMERU', title)
    worksheet217.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet217.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet217.write('A22', 'LOKASI', header)
    worksheet217.write('B22', 'TOTAL', header)
    worksheet217.merge_range('A21:B21', 'RANK', header)
    worksheet217.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet217.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet217.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet217.merge_range('F21:F22', 'KELAS', header)
    worksheet217.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet217.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet217.write('G22', 'MAW', body)
    worksheet217.write('H22', 'MAP', body)
    worksheet217.write('I22', 'IND', body)
    worksheet217.write('J22', 'ENG', body)
    worksheet217.write('J22', 'SEJ', body)
    worksheet217.write('K22', 'GEO', body)
    worksheet217.write('M22', 'EKO', body)
    worksheet217.write('L22', 'SOS', body)
    worksheet217.write('L22', 'FIS', body)
    worksheet217.write('L22', 'KIM', body)
    worksheet217.write('L22', 'BIO', body)
    worksheet217.write('N22', 'JML', body)
    worksheet217.write('O22', 'MAW', body)
    worksheet217.write('O22', 'MAP', body)
    worksheet217.write('P22', 'IND', body)
    worksheet217.write('Q22', 'ENG', body)
    worksheet217.write('R22', 'SEJ', body)
    worksheet217.write('S22', 'GEO', body)
    worksheet217.write('U22', 'EKO', body)
    worksheet217.write('T22', 'SOS', body)
    worksheet217.write('T22', 'FIS', body)
    worksheet217.write('T22', 'KIM', body)
    worksheet217.write('T22', 'BIO', body)
    worksheet217.write('V22', 'JML', body)

    worksheet217.conditional_format(22, 0, row217+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 218
    worksheet218.insert_image('A1', r'logo resmi nf.jpg')

    worksheet218.set_column('A:A', 7, center)
    worksheet218.set_column('B:B', 6, center)
    worksheet218.set_column('C:C', 18.14, center)
    worksheet218.set_column('D:D', 25, left)
    worksheet218.set_column('E:E', 13.14, left)
    worksheet218.set_column('F:F', 8.57, center)
    worksheet218.set_column('G:V', 5, center)
    worksheet218.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIKEAS', title)
    worksheet218.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet218.write('A5', 'LOKASI', header)
    worksheet218.write('B5', 'TOTAL', header)
    worksheet218.merge_range('A4:B4', 'RANK', header)
    worksheet218.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet218.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet218.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet218.merge_range('F4:F5', 'KELAS', header)
    worksheet218.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet218.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet218.write('G5', 'MAW', body)
    worksheet218.write('H5', 'MAP', body)
    worksheet218.write('I5', 'IND', body)
    worksheet218.write('J5', 'ENG', body)
    worksheet218.write('K5', 'SEJ', body)
    worksheet218.write('L5', 'GEO', body)
    worksheet218.write('M5', 'EKO', body)
    worksheet218.write('N5', 'SOS', body)
    worksheet218.write('O5', 'FIS', body)
    worksheet218.write('P5', 'KIM', body)
    worksheet218.write('Q5', 'BIO', body)
    worksheet218.write('R5', 'JML', body)
    worksheet218.write('S5', 'MAW', body)
    worksheet218.write('T5', 'MAP', body)
    worksheet218.write('U5', 'IND', body)
    worksheet218.write('V5', 'ENG', body)
    worksheet218.write('W5', 'SEJ', body)
    worksheet218.write('X5', 'GEO', body)
    worksheet218.write('Y5', 'EKO', body)
    worksheet218.write('Z5', 'SOS', body)
    worksheet218.write('AA5', 'FIS', body)
    worksheet218.write('AB5', 'KIM', body)
    worksheet218.write('AC5', 'BIO', body)
    worksheet218.write('AD5', 'JML', body)

    worksheet218.conditional_format(5, 0, row218_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet218.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF CIKEAS', title)
    worksheet218.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet218.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet218.write('A22', 'LOKASI', header)
    worksheet218.write('B22', 'TOTAL', header)
    worksheet218.merge_range('A21:B21', 'RANK', header)
    worksheet218.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet218.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet218.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet218.merge_range('F21:F22', 'KELAS', header)
    worksheet218.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet218.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet218.write('G22', 'MAW', body)
    worksheet218.write('H22', 'MAP', body)
    worksheet218.write('I22', 'IND', body)
    worksheet218.write('J22', 'ENG', body)
    worksheet218.write('J22', 'SEJ', body)
    worksheet218.write('K22', 'GEO', body)
    worksheet218.write('M22', 'EKO', body)
    worksheet218.write('L22', 'SOS', body)
    worksheet218.write('L22', 'FIS', body)
    worksheet218.write('L22', 'KIM', body)
    worksheet218.write('L22', 'BIO', body)
    worksheet218.write('N22', 'JML', body)
    worksheet218.write('O22', 'MAW', body)
    worksheet218.write('O22', 'MAP', body)
    worksheet218.write('P22', 'IND', body)
    worksheet218.write('Q22', 'ENG', body)
    worksheet218.write('R22', 'SEJ', body)
    worksheet218.write('S22', 'GEO', body)
    worksheet218.write('U22', 'EKO', body)
    worksheet218.write('T22', 'SOS', body)
    worksheet218.write('T22', 'FIS', body)
    worksheet218.write('T22', 'KIM', body)
    worksheet218.write('T22', 'BIO', body)
    worksheet218.write('V22', 'JML', body)

    worksheet218.conditional_format(22, 0, row218+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 219
    worksheet219.insert_image('A1', r'logo resmi nf.jpg')

    worksheet219.set_column('A:A', 7, center)
    worksheet219.set_column('B:B', 6, center)
    worksheet219.set_column('C:C', 18.14, center)
    worksheet219.set_column('D:D', 25, left)
    worksheet219.set_column('E:E', 13.14, left)
    worksheet219.set_column('F:F', 8.57, center)
    worksheet219.set_column('G:V', 5, center)
    worksheet219.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIJAWA MASJID', title)
    worksheet219.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet219.write('A5', 'LOKASI', header)
    worksheet219.write('B5', 'TOTAL', header)
    worksheet219.merge_range('A4:B4', 'RANK', header)
    worksheet219.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet219.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet219.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet219.merge_range('F4:F5', 'KELAS', header)
    worksheet219.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet219.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet219.write('G5', 'MAW', body)
    worksheet219.write('H5', 'MAP', body)
    worksheet219.write('I5', 'IND', body)
    worksheet219.write('J5', 'ENG', body)
    worksheet219.write('K5', 'SEJ', body)
    worksheet219.write('L5', 'GEO', body)
    worksheet219.write('M5', 'EKO', body)
    worksheet219.write('N5', 'SOS', body)
    worksheet219.write('O5', 'FIS', body)
    worksheet219.write('P5', 'KIM', body)
    worksheet219.write('Q5', 'BIO', body)
    worksheet219.write('R5', 'JML', body)
    worksheet219.write('S5', 'MAW', body)
    worksheet219.write('T5', 'MAP', body)
    worksheet219.write('U5', 'IND', body)
    worksheet219.write('V5', 'ENG', body)
    worksheet219.write('W5', 'SEJ', body)
    worksheet219.write('X5', 'GEO', body)
    worksheet219.write('Y5', 'EKO', body)
    worksheet219.write('Z5', 'SOS', body)
    worksheet219.write('AA5', 'FIS', body)
    worksheet219.write('AB5', 'KIM', body)
    worksheet219.write('AC5', 'BIO', body)
    worksheet219.write('AD5', 'JML', body)

    worksheet219.conditional_format(5, 0, row219_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet219.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF CIJAWA MASJID', title)
    worksheet219.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet219.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet219.write('A22', 'LOKASI', header)
    worksheet219.write('B22', 'TOTAL', header)
    worksheet219.merge_range('A21:B21', 'RANK', header)
    worksheet219.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet219.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet219.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet219.merge_range('F21:F22', 'KELAS', header)
    worksheet219.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet219.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet219.write('G22', 'MAW', body)
    worksheet219.write('H22', 'MAP', body)
    worksheet219.write('I22', 'IND', body)
    worksheet219.write('J22', 'ENG', body)
    worksheet219.write('J22', 'SEJ', body)
    worksheet219.write('K22', 'GEO', body)
    worksheet219.write('M22', 'EKO', body)
    worksheet219.write('L22', 'SOS', body)
    worksheet219.write('L22', 'FIS', body)
    worksheet219.write('L22', 'KIM', body)
    worksheet219.write('L22', 'BIO', body)
    worksheet219.write('N22', 'JML', body)
    worksheet219.write('O22', 'MAW', body)
    worksheet219.write('O22', 'MAP', body)
    worksheet219.write('P22', 'IND', body)
    worksheet219.write('Q22', 'ENG', body)
    worksheet219.write('R22', 'SEJ', body)
    worksheet219.write('S22', 'GEO', body)
    worksheet219.write('U22', 'EKO', body)
    worksheet219.write('T22', 'SOS', body)
    worksheet219.write('T22', 'FIS', body)
    worksheet219.write('T22', 'KIM', body)
    worksheet219.write('T22', 'BIO', body)
    worksheet219.write('V22', 'JML', body)

    worksheet219.conditional_format(22, 0, row219+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 220
    worksheet220.insert_image('A1', r'logo resmi nf.jpg')

    worksheet220.set_column('A:A', 7, center)
    worksheet220.set_column('B:B', 6, center)
    worksheet220.set_column('C:C', 18.14, center)
    worksheet220.set_column('D:D', 25, left)
    worksheet220.set_column('E:E', 13.14, left)
    worksheet220.set_column('F:F', 8.57, center)
    worksheet220.set_column('G:V', 5, center)
    worksheet220.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PALEDANG', title)
    worksheet220.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet220.write('A5', 'LOKASI', header)
    worksheet220.write('B5', 'TOTAL', header)
    worksheet220.merge_range('A4:B4', 'RANK', header)
    worksheet220.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet220.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet220.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet220.merge_range('F4:F5', 'KELAS', header)
    worksheet220.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet220.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet220.write('G5', 'MAW', body)
    worksheet220.write('H5', 'MAP', body)
    worksheet220.write('I5', 'IND', body)
    worksheet220.write('J5', 'ENG', body)
    worksheet220.write('K5', 'SEJ', body)
    worksheet220.write('L5', 'GEO', body)
    worksheet220.write('M5', 'EKO', body)
    worksheet220.write('N5', 'SOS', body)
    worksheet220.write('O5', 'FIS', body)
    worksheet220.write('P5', 'KIM', body)
    worksheet220.write('Q5', 'BIO', body)
    worksheet220.write('R5', 'JML', body)
    worksheet220.write('S5', 'MAW', body)
    worksheet220.write('T5', 'MAP', body)
    worksheet220.write('U5', 'IND', body)
    worksheet220.write('V5', 'ENG', body)
    worksheet220.write('W5', 'SEJ', body)
    worksheet220.write('X5', 'GEO', body)
    worksheet220.write('Y5', 'EKO', body)
    worksheet220.write('Z5', 'SOS', body)
    worksheet220.write('AA5', 'FIS', body)
    worksheet220.write('AB5', 'KIM', body)
    worksheet220.write('AC5', 'BIO', body)
    worksheet220.write('AD5', 'JML', body)

    worksheet220.conditional_format(5, 0, row220_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet220.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF PALEDANG', title)
    worksheet220.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet220.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet220.write('A22', 'LOKASI', header)
    worksheet220.write('B22', 'TOTAL', header)
    worksheet220.merge_range('A21:B21', 'RANK', header)
    worksheet220.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet220.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet220.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet220.merge_range('F21:F22', 'KELAS', header)
    worksheet220.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet220.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet220.write('G22', 'MAW', body)
    worksheet220.write('H22', 'MAP', body)
    worksheet220.write('I22', 'IND', body)
    worksheet220.write('J22', 'ENG', body)
    worksheet220.write('J22', 'SEJ', body)
    worksheet220.write('K22', 'GEO', body)
    worksheet220.write('M22', 'EKO', body)
    worksheet220.write('L22', 'SOS', body)
    worksheet220.write('L22', 'FIS', body)
    worksheet220.write('L22', 'KIM', body)
    worksheet220.write('L22', 'BIO', body)
    worksheet220.write('N22', 'JML', body)
    worksheet220.write('O22', 'MAW', body)
    worksheet220.write('O22', 'MAP', body)
    worksheet220.write('P22', 'IND', body)
    worksheet220.write('Q22', 'ENG', body)
    worksheet220.write('R22', 'SEJ', body)
    worksheet220.write('S22', 'GEO', body)
    worksheet220.write('U22', 'EKO', body)
    worksheet220.write('T22', 'SOS', body)
    worksheet220.write('T22', 'FIS', body)
    worksheet220.write('T22', 'KIM', body)
    worksheet220.write('T22', 'BIO', body)
    worksheet220.write('V22', 'JML', body)

    worksheet220.conditional_format(22, 0, row220+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 226
    worksheet226.insert_image('A1', r'logo resmi nf.jpg')

    worksheet226.set_column('A:A', 7, center)
    worksheet226.set_column('B:B', 6, center)
    worksheet226.set_column('C:C', 18.14, center)
    worksheet226.set_column('D:D', 25, left)
    worksheet226.set_column('E:E', 13.14, left)
    worksheet226.set_column('F:F', 8.57, center)
    worksheet226.set_column('G:V', 5, center)
    worksheet226.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF GEDONG KUNING', title)
    worksheet226.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet226.write('A5', 'LOKASI', header)
    worksheet226.write('B5', 'TOTAL', header)
    worksheet226.merge_range('A4:B4', 'RANK', header)
    worksheet226.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet226.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet226.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet226.merge_range('F4:F5', 'KELAS', header)
    worksheet226.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet226.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet226.write('G5', 'MAW', body)
    worksheet226.write('H5', 'MAP', body)
    worksheet226.write('I5', 'IND', body)
    worksheet226.write('J5', 'ENG', body)
    worksheet226.write('K5', 'SEJ', body)
    worksheet226.write('L5', 'GEO', body)
    worksheet226.write('M5', 'EKO', body)
    worksheet226.write('N5', 'SOS', body)
    worksheet226.write('O5', 'FIS', body)
    worksheet226.write('P5', 'KIM', body)
    worksheet226.write('Q5', 'BIO', body)
    worksheet226.write('R5', 'JML', body)
    worksheet226.write('S5', 'MAW', body)
    worksheet226.write('T5', 'MAP', body)
    worksheet226.write('U5', 'IND', body)
    worksheet226.write('V5', 'ENG', body)
    worksheet226.write('W5', 'SEJ', body)
    worksheet226.write('X5', 'GEO', body)
    worksheet226.write('Y5', 'EKO', body)
    worksheet226.write('Z5', 'SOS', body)
    worksheet226.write('AA5', 'FIS', body)
    worksheet226.write('AB5', 'KIM', body)
    worksheet226.write('AC5', 'BIO', body)
    worksheet226.write('AD5', 'JML', body)

    worksheet226.conditional_format(5, 0, row226_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet226.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF GEDONG KUNING', title)
    worksheet226.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet226.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet226.write('A22', 'LOKASI', header)
    worksheet226.write('B22', 'TOTAL', header)
    worksheet226.merge_range('A21:B21', 'RANK', header)
    worksheet226.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet226.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet226.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet226.merge_range('F21:F22', 'KELAS', header)
    worksheet226.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet226.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet226.write('G22', 'MAW', body)
    worksheet226.write('H22', 'MAP', body)
    worksheet226.write('I22', 'IND', body)
    worksheet226.write('J22', 'ENG', body)
    worksheet226.write('J22', 'SEJ', body)
    worksheet226.write('K22', 'GEO', body)
    worksheet226.write('M22', 'EKO', body)
    worksheet226.write('L22', 'SOS', body)
    worksheet226.write('L22', 'FIS', body)
    worksheet226.write('L22', 'KIM', body)
    worksheet226.write('L22', 'BIO', body)
    worksheet226.write('N22', 'JML', body)
    worksheet226.write('O22', 'MAW', body)
    worksheet226.write('O22', 'MAP', body)
    worksheet226.write('P22', 'IND', body)
    worksheet226.write('Q22', 'ENG', body)
    worksheet226.write('R22', 'SEJ', body)
    worksheet226.write('S22', 'GEO', body)
    worksheet226.write('U22', 'EKO', body)
    worksheet226.write('T22', 'SOS', body)
    worksheet226.write('T22', 'FIS', body)
    worksheet226.write('T22', 'KIM', body)
    worksheet226.write('T22', 'BIO', body)
    worksheet226.write('V22', 'JML', body)

    worksheet226.conditional_format(22, 0, row226+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 227
    worksheet227.insert_image('A1', r'logo resmi nf.jpg')

    worksheet227.set_column('A:A', 7, center)
    worksheet227.set_column('B:B', 6, center)
    worksheet227.set_column('C:C', 18.14, center)
    worksheet227.set_column('D:D', 25, left)
    worksheet227.set_column('E:E', 13.14, left)
    worksheet227.set_column('F:F', 8.57, center)
    worksheet227.set_column('G:V', 5, center)
    worksheet227.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF JATIWARINGIN', title)
    worksheet227.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet227.write('A5', 'LOKASI', header)
    worksheet227.write('B5', 'TOTAL', header)
    worksheet227.merge_range('A4:B4', 'RANK', header)
    worksheet227.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet227.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet227.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet227.merge_range('F4:F5', 'KELAS', header)
    worksheet227.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet227.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet227.write('G5', 'MAW', body)
    worksheet227.write('H5', 'MAP', body)
    worksheet227.write('I5', 'IND', body)
    worksheet227.write('J5', 'ENG', body)
    worksheet227.write('K5', 'SEJ', body)
    worksheet227.write('L5', 'GEO', body)
    worksheet227.write('M5', 'EKO', body)
    worksheet227.write('N5', 'SOS', body)
    worksheet227.write('O5', 'FIS', body)
    worksheet227.write('P5', 'KIM', body)
    worksheet227.write('Q5', 'BIO', body)
    worksheet227.write('R5', 'JML', body)
    worksheet227.write('S5', 'MAW', body)
    worksheet227.write('T5', 'MAP', body)
    worksheet227.write('U5', 'IND', body)
    worksheet227.write('V5', 'ENG', body)
    worksheet227.write('W5', 'SEJ', body)
    worksheet227.write('X5', 'GEO', body)
    worksheet227.write('Y5', 'EKO', body)
    worksheet227.write('Z5', 'SOS', body)
    worksheet227.write('AA5', 'FIS', body)
    worksheet227.write('AB5', 'KIM', body)
    worksheet227.write('AC5', 'BIO', body)
    worksheet227.write('AD5', 'JML', body)

    worksheet227.conditional_format(5, 0, row227_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet227.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF JATIWARINGIN', title)
    worksheet227.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet227.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet227.write('A22', 'LOKASI', header)
    worksheet227.write('B22', 'TOTAL', header)
    worksheet227.merge_range('A21:B21', 'RANK', header)
    worksheet227.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet227.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet227.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet227.merge_range('F21:F22', 'KELAS', header)
    worksheet227.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet227.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet227.write('G22', 'MAW', body)
    worksheet227.write('H22', 'MAP', body)
    worksheet227.write('I22', 'IND', body)
    worksheet227.write('J22', 'ENG', body)
    worksheet227.write('J22', 'SEJ', body)
    worksheet227.write('K22', 'GEO', body)
    worksheet227.write('M22', 'EKO', body)
    worksheet227.write('L22', 'SOS', body)
    worksheet227.write('L22', 'FIS', body)
    worksheet227.write('L22', 'KIM', body)
    worksheet227.write('L22', 'BIO', body)
    worksheet227.write('N22', 'JML', body)
    worksheet227.write('O22', 'MAW', body)
    worksheet227.write('O22', 'MAP', body)
    worksheet227.write('P22', 'IND', body)
    worksheet227.write('Q22', 'ENG', body)
    worksheet227.write('R22', 'SEJ', body)
    worksheet227.write('S22', 'GEO', body)
    worksheet227.write('U22', 'EKO', body)
    worksheet227.write('T22', 'SOS', body)
    worksheet227.write('T22', 'FIS', body)
    worksheet227.write('T22', 'KIM', body)
    worksheet227.write('T22', 'BIO', body)
    worksheet227.write('V22', 'JML', body)

    worksheet227.conditional_format(22, 0, row227+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 228
    worksheet228.insert_image('A1', r'logo resmi nf.jpg')

    worksheet228.set_column('A:A', 7, center)
    worksheet228.set_column('B:B', 6, center)
    worksheet228.set_column('C:C', 18.14, center)
    worksheet228.set_column('D:D', 25, left)
    worksheet228.set_column('E:E', 13.14, left)
    worksheet228.set_column('F:F', 8.57, center)
    worksheet228.set_column('G:V', 5, center)
    worksheet228.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CILEDUG', title)
    worksheet228.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet228.write('A5', 'LOKASI', header)
    worksheet228.write('B5', 'TOTAL', header)
    worksheet228.merge_range('A4:B4', 'RANK', header)
    worksheet228.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet228.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet228.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet228.merge_range('F4:F5', 'KELAS', header)
    worksheet228.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet228.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet228.write('G5', 'MAW', body)
    worksheet228.write('H5', 'MAP', body)
    worksheet228.write('I5', 'IND', body)
    worksheet228.write('J5', 'ENG', body)
    worksheet228.write('K5', 'SEJ', body)
    worksheet228.write('L5', 'GEO', body)
    worksheet228.write('M5', 'EKO', body)
    worksheet228.write('N5', 'SOS', body)
    worksheet228.write('O5', 'FIS', body)
    worksheet228.write('P5', 'KIM', body)
    worksheet228.write('Q5', 'BIO', body)
    worksheet228.write('R5', 'JML', body)
    worksheet228.write('S5', 'MAW', body)
    worksheet228.write('T5', 'MAP', body)
    worksheet228.write('U5', 'IND', body)
    worksheet228.write('V5', 'ENG', body)
    worksheet228.write('W5', 'SEJ', body)
    worksheet228.write('X5', 'GEO', body)
    worksheet228.write('Y5', 'EKO', body)
    worksheet228.write('Z5', 'SOS', body)
    worksheet228.write('AA5', 'FIS', body)
    worksheet228.write('AB5', 'KIM', body)
    worksheet228.write('AC5', 'BIO', body)
    worksheet228.write('AD5', 'JML', body)

    worksheet228.conditional_format(5, 0, row228_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet228.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF CILEDUG', title)
    worksheet228.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet228.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet228.write('A22', 'LOKASI', header)
    worksheet228.write('B22', 'TOTAL', header)
    worksheet228.merge_range('A21:B21', 'RANK', header)
    worksheet228.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet228.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet228.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet228.merge_range('F21:F22', 'KELAS', header)
    worksheet228.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet228.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet228.write('G22', 'MAW', body)
    worksheet228.write('H22', 'MAP', body)
    worksheet228.write('I22', 'IND', body)
    worksheet228.write('J22', 'ENG', body)
    worksheet228.write('J22', 'SEJ', body)
    worksheet228.write('K22', 'GEO', body)
    worksheet228.write('M22', 'EKO', body)
    worksheet228.write('L22', 'SOS', body)
    worksheet228.write('L22', 'FIS', body)
    worksheet228.write('L22', 'KIM', body)
    worksheet228.write('L22', 'BIO', body)
    worksheet228.write('N22', 'JML', body)
    worksheet228.write('O22', 'MAW', body)
    worksheet228.write('O22', 'MAP', body)
    worksheet228.write('P22', 'IND', body)
    worksheet228.write('Q22', 'ENG', body)
    worksheet228.write('R22', 'SEJ', body)
    worksheet228.write('S22', 'GEO', body)
    worksheet228.write('U22', 'EKO', body)
    worksheet228.write('T22', 'SOS', body)
    worksheet228.write('T22', 'FIS', body)
    worksheet228.write('T22', 'KIM', body)
    worksheet228.write('T22', 'BIO', body)
    worksheet228.write('V22', 'JML', body)

    worksheet228.conditional_format(22, 0, row228+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 229
    worksheet229.insert_image('A1', r'logo resmi nf.jpg')

    worksheet229.set_column('A:A', 7, center)
    worksheet229.set_column('B:B', 6, center)
    worksheet229.set_column('C:C', 18.14, center)
    worksheet229.set_column('D:D', 25, left)
    worksheet229.set_column('E:E', 13.14, left)
    worksheet229.set_column('F:F', 8.57, center)
    worksheet229.set_column('G:V', 5, center)
    worksheet229.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KRANGGAN', title)
    worksheet229.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet229.write('A5', 'LOKASI', header)
    worksheet229.write('B5', 'TOTAL', header)
    worksheet229.merge_range('A4:B4', 'RANK', header)
    worksheet229.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet229.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet229.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet229.merge_range('F4:F5', 'KELAS', header)
    worksheet229.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet229.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet229.write('G5', 'MAW', body)
    worksheet229.write('H5', 'MAP', body)
    worksheet229.write('I5', 'IND', body)
    worksheet229.write('J5', 'ENG', body)
    worksheet229.write('K5', 'SEJ', body)
    worksheet229.write('L5', 'GEO', body)
    worksheet229.write('M5', 'EKO', body)
    worksheet229.write('N5', 'SOS', body)
    worksheet229.write('O5', 'FIS', body)
    worksheet229.write('P5', 'KIM', body)
    worksheet229.write('Q5', 'BIO', body)
    worksheet229.write('R5', 'JML', body)
    worksheet229.write('S5', 'MAW', body)
    worksheet229.write('T5', 'MAP', body)
    worksheet229.write('U5', 'IND', body)
    worksheet229.write('V5', 'ENG', body)
    worksheet229.write('W5', 'SEJ', body)
    worksheet229.write('X5', 'GEO', body)
    worksheet229.write('Y5', 'EKO', body)
    worksheet229.write('Z5', 'SOS', body)
    worksheet229.write('AA5', 'FIS', body)
    worksheet229.write('AB5', 'KIM', body)
    worksheet229.write('AC5', 'BIO', body)
    worksheet229.write('AD5', 'JML', body)

    worksheet229.conditional_format(5, 0, row229_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet229.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF KRANGGAN', title)
    worksheet229.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet229.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet229.write('A22', 'LOKASI', header)
    worksheet229.write('B22', 'TOTAL', header)
    worksheet229.merge_range('A21:B21', 'RANK', header)
    worksheet229.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet229.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet229.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet229.merge_range('F21:F22', 'KELAS', header)
    worksheet229.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet229.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet229.write('G22', 'MAW', body)
    worksheet229.write('H22', 'MAP', body)
    worksheet229.write('I22', 'IND', body)
    worksheet229.write('J22', 'ENG', body)
    worksheet229.write('J22', 'SEJ', body)
    worksheet229.write('K22', 'GEO', body)
    worksheet229.write('M22', 'EKO', body)
    worksheet229.write('L22', 'SOS', body)
    worksheet229.write('L22', 'FIS', body)
    worksheet229.write('L22', 'KIM', body)
    worksheet229.write('L22', 'BIO', body)
    worksheet229.write('N22', 'JML', body)
    worksheet229.write('O22', 'MAW', body)
    worksheet229.write('O22', 'MAP', body)
    worksheet229.write('P22', 'IND', body)
    worksheet229.write('Q22', 'ENG', body)
    worksheet229.write('R22', 'SEJ', body)
    worksheet229.write('S22', 'GEO', body)
    worksheet229.write('U22', 'EKO', body)
    worksheet229.write('T22', 'SOS', body)
    worksheet229.write('T22', 'FIS', body)
    worksheet229.write('T22', 'KIM', body)
    worksheet229.write('T22', 'BIO', body)
    worksheet229.write('V22', 'JML', body)

    worksheet229.conditional_format(22, 0, row229+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 230
    worksheet230.insert_image('A1', r'logo resmi nf.jpg')

    worksheet230.set_column('A:A', 7, center)
    worksheet230.set_column('B:B', 6, center)
    worksheet230.set_column('C:C', 18.14, center)
    worksheet230.set_column('D:D', 25, left)
    worksheet230.set_column('E:E', 13.14, left)
    worksheet230.set_column('F:F', 8.57, center)
    worksheet230.set_column('G:V', 5, center)
    worksheet230.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF MUSTIKA JAYA', title)
    worksheet230.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet230.write('A5', 'LOKASI', header)
    worksheet230.write('B5', 'TOTAL', header)
    worksheet230.merge_range('A4:B4', 'RANK', header)
    worksheet230.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet230.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet230.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet230.merge_range('F4:F5', 'KELAS', header)
    worksheet230.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet230.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet230.write('G5', 'MAW', body)
    worksheet230.write('H5', 'MAP', body)
    worksheet230.write('I5', 'IND', body)
    worksheet230.write('J5', 'ENG', body)
    worksheet230.write('K5', 'SEJ', body)
    worksheet230.write('L5', 'GEO', body)
    worksheet230.write('M5', 'EKO', body)
    worksheet230.write('N5', 'SOS', body)
    worksheet230.write('O5', 'FIS', body)
    worksheet230.write('P5', 'KIM', body)
    worksheet230.write('Q5', 'BIO', body)
    worksheet230.write('R5', 'JML', body)
    worksheet230.write('S5', 'MAW', body)
    worksheet230.write('T5', 'MAP', body)
    worksheet230.write('U5', 'IND', body)
    worksheet230.write('V5', 'ENG', body)
    worksheet230.write('W5', 'SEJ', body)
    worksheet230.write('X5', 'GEO', body)
    worksheet230.write('Y5', 'EKO', body)
    worksheet230.write('Z5', 'SOS', body)
    worksheet230.write('AA5', 'FIS', body)
    worksheet230.write('AB5', 'KIM', body)
    worksheet230.write('AC5', 'BIO', body)
    worksheet230.write('AD5', 'JML', body)

    worksheet230.conditional_format(5, 0, row230_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet230.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF MUSTIKA JAYA', title)
    worksheet230.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet230.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet230.write('A22', 'LOKASI', header)
    worksheet230.write('B22', 'TOTAL', header)
    worksheet230.merge_range('A21:B21', 'RANK', header)
    worksheet230.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet230.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet230.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet230.merge_range('F21:F22', 'KELAS', header)
    worksheet230.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet230.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet230.write('G22', 'MAW', body)
    worksheet230.write('H22', 'MAP', body)
    worksheet230.write('I22', 'IND', body)
    worksheet230.write('J22', 'ENG', body)
    worksheet230.write('J22', 'SEJ', body)
    worksheet230.write('K22', 'GEO', body)
    worksheet230.write('M22', 'EKO', body)
    worksheet230.write('L22', 'SOS', body)
    worksheet230.write('L22', 'FIS', body)
    worksheet230.write('L22', 'KIM', body)
    worksheet230.write('L22', 'BIO', body)
    worksheet230.write('N22', 'JML', body)
    worksheet230.write('O22', 'MAW', body)
    worksheet230.write('O22', 'MAP', body)
    worksheet230.write('P22', 'IND', body)
    worksheet230.write('Q22', 'ENG', body)
    worksheet230.write('R22', 'SEJ', body)
    worksheet230.write('S22', 'GEO', body)
    worksheet230.write('U22', 'EKO', body)
    worksheet230.write('T22', 'SOS', body)
    worksheet230.write('T22', 'FIS', body)
    worksheet230.write('T22', 'KIM', body)
    worksheet230.write('T22', 'BIO', body)
    worksheet230.write('V22', 'JML', body)

    worksheet230.conditional_format(22, 0, row230+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 231
    worksheet231.insert_image('A1', r'logo resmi nf.jpg')

    worksheet231.set_column('A:A', 7, center)
    worksheet231.set_column('B:B', 6, center)
    worksheet231.set_column('C:C', 18.14, center)
    worksheet231.set_column('D:D', 25, left)
    worksheet231.set_column('E:E', 13.14, left)
    worksheet231.set_column('F:F', 8.57, center)
    worksheet231.set_column('G:V', 5, center)
    worksheet231.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF ALEXINDO', title)
    worksheet231.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet231.write('A5', 'LOKASI', header)
    worksheet231.write('B5', 'TOTAL', header)
    worksheet231.merge_range('A4:B4', 'RANK', header)
    worksheet231.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet231.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet231.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet231.merge_range('F4:F5', 'KELAS', header)
    worksheet231.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet231.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet231.write('G5', 'MAW', body)
    worksheet231.write('H5', 'MAP', body)
    worksheet231.write('I5', 'IND', body)
    worksheet231.write('J5', 'ENG', body)
    worksheet231.write('K5', 'SEJ', body)
    worksheet231.write('L5', 'GEO', body)
    worksheet231.write('M5', 'EKO', body)
    worksheet231.write('N5', 'SOS', body)
    worksheet231.write('O5', 'FIS', body)
    worksheet231.write('P5', 'KIM', body)
    worksheet231.write('Q5', 'BIO', body)
    worksheet231.write('R5', 'JML', body)
    worksheet231.write('S5', 'MAW', body)
    worksheet231.write('T5', 'MAP', body)
    worksheet231.write('U5', 'IND', body)
    worksheet231.write('V5', 'ENG', body)
    worksheet231.write('W5', 'SEJ', body)
    worksheet231.write('X5', 'GEO', body)
    worksheet231.write('Y5', 'EKO', body)
    worksheet231.write('Z5', 'SOS', body)
    worksheet231.write('AA5', 'FIS', body)
    worksheet231.write('AB5', 'KIM', body)
    worksheet231.write('AC5', 'BIO', body)
    worksheet231.write('AD5', 'JML', body)

    worksheet231.conditional_format(5, 0, row231_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet231.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF ALEXINDO', title)
    worksheet231.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet231.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet231.write('A22', 'LOKASI', header)
    worksheet231.write('B22', 'TOTAL', header)
    worksheet231.merge_range('A21:B21', 'RANK', header)
    worksheet231.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet231.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet231.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet231.merge_range('F21:F22', 'KELAS', header)
    worksheet231.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet231.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet231.write('G22', 'MAW', body)
    worksheet231.write('H22', 'MAP', body)
    worksheet231.write('I22', 'IND', body)
    worksheet231.write('J22', 'ENG', body)
    worksheet231.write('J22', 'SEJ', body)
    worksheet231.write('K22', 'GEO', body)
    worksheet231.write('M22', 'EKO', body)
    worksheet231.write('L22', 'SOS', body)
    worksheet231.write('L22', 'FIS', body)
    worksheet231.write('L22', 'KIM', body)
    worksheet231.write('L22', 'BIO', body)
    worksheet231.write('N22', 'JML', body)
    worksheet231.write('O22', 'MAW', body)
    worksheet231.write('O22', 'MAP', body)
    worksheet231.write('P22', 'IND', body)
    worksheet231.write('Q22', 'ENG', body)
    worksheet231.write('R22', 'SEJ', body)
    worksheet231.write('S22', 'GEO', body)
    worksheet231.write('U22', 'EKO', body)
    worksheet231.write('T22', 'SOS', body)
    worksheet231.write('T22', 'FIS', body)
    worksheet231.write('T22', 'KIM', body)
    worksheet231.write('T22', 'BIO', body)
    worksheet231.write('V22', 'JML', body)

    worksheet231.conditional_format(22, 0, row231+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 233
    worksheet233.insert_image('A1', r'logo resmi nf.jpg')

    worksheet233.set_column('A:A', 7, center)
    worksheet233.set_column('B:B', 6, center)
    worksheet233.set_column('C:C', 18.14, center)
    worksheet233.set_column('D:D', 25, left)
    worksheet233.set_column('E:E', 13.14, left)
    worksheet233.set_column('F:F', 8.57, center)
    worksheet233.set_column('G:V', 5, center)
    worksheet233.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIBITUNG', title)
    worksheet233.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet233.write('A5', 'LOKASI', header)
    worksheet233.write('B5', 'TOTAL', header)
    worksheet233.merge_range('A4:B4', 'RANK', header)
    worksheet233.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet233.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet233.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet233.merge_range('F4:F5', 'KELAS', header)
    worksheet233.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet233.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet233.write('G5', 'MAW', body)
    worksheet233.write('H5', 'MAP', body)
    worksheet233.write('I5', 'IND', body)
    worksheet233.write('J5', 'ENG', body)
    worksheet233.write('K5', 'SEJ', body)
    worksheet233.write('L5', 'GEO', body)
    worksheet233.write('M5', 'EKO', body)
    worksheet233.write('N5', 'SOS', body)
    worksheet233.write('O5', 'FIS', body)
    worksheet233.write('P5', 'KIM', body)
    worksheet233.write('Q5', 'BIO', body)
    worksheet233.write('R5', 'JML', body)
    worksheet233.write('S5', 'MAW', body)
    worksheet233.write('T5', 'MAP', body)
    worksheet233.write('U5', 'IND', body)
    worksheet233.write('V5', 'ENG', body)
    worksheet233.write('W5', 'SEJ', body)
    worksheet233.write('X5', 'GEO', body)
    worksheet233.write('Y5', 'EKO', body)
    worksheet233.write('Z5', 'SOS', body)
    worksheet233.write('AA5', 'FIS', body)
    worksheet233.write('AB5', 'KIM', body)
    worksheet233.write('AC5', 'BIO', body)
    worksheet233.write('AD5', 'JML', body)

    worksheet233.conditional_format(5, 0, row233_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet233.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF CIBITUNG', title)
    worksheet233.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet233.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet233.write('A22', 'LOKASI', header)
    worksheet233.write('B22', 'TOTAL', header)
    worksheet233.merge_range('A21:B21', 'RANK', header)
    worksheet233.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet233.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet233.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet233.merge_range('F21:F22', 'KELAS', header)
    worksheet233.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet233.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet233.write('G22', 'MAW', body)
    worksheet233.write('H22', 'MAP', body)
    worksheet233.write('I22', 'IND', body)
    worksheet233.write('J22', 'ENG', body)
    worksheet233.write('J22', 'SEJ', body)
    worksheet233.write('K22', 'GEO', body)
    worksheet233.write('M22', 'EKO', body)
    worksheet233.write('L22', 'SOS', body)
    worksheet233.write('L22', 'FIS', body)
    worksheet233.write('L22', 'KIM', body)
    worksheet233.write('L22', 'BIO', body)
    worksheet233.write('N22', 'JML', body)
    worksheet233.write('O22', 'MAW', body)
    worksheet233.write('O22', 'MAP', body)
    worksheet233.write('P22', 'IND', body)
    worksheet233.write('Q22', 'ENG', body)
    worksheet233.write('R22', 'SEJ', body)
    worksheet233.write('S22', 'GEO', body)
    worksheet233.write('U22', 'EKO', body)
    worksheet233.write('T22', 'SOS', body)
    worksheet233.write('T22', 'FIS', body)
    worksheet233.write('T22', 'KIM', body)
    worksheet233.write('T22', 'BIO', body)
    worksheet233.write('V22', 'JML', body)

    worksheet233.conditional_format(22, 0, row233+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 234
    worksheet234.insert_image('A1', r'logo resmi nf.jpg')

    worksheet234.set_column('A:A', 7, center)
    worksheet234.set_column('B:B', 6, center)
    worksheet234.set_column('C:C', 18.14, center)
    worksheet234.set_column('D:D', 25, left)
    worksheet234.set_column('E:E', 13.14, left)
    worksheet234.set_column('F:F', 8.57, center)
    worksheet234.set_column('G:V', 5, center)
    worksheet234.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF KRAMAT JAYA', title)
    worksheet234.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet234.write('A5', 'LOKASI', header)
    worksheet234.write('B5', 'TOTAL', header)
    worksheet234.merge_range('A4:B4', 'RANK', header)
    worksheet234.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet234.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet234.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet234.merge_range('F4:F5', 'KELAS', header)
    worksheet234.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet234.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet234.write('G5', 'MAW', body)
    worksheet234.write('H5', 'MAP', body)
    worksheet234.write('I5', 'IND', body)
    worksheet234.write('J5', 'ENG', body)
    worksheet234.write('K5', 'SEJ', body)
    worksheet234.write('L5', 'GEO', body)
    worksheet234.write('M5', 'EKO', body)
    worksheet234.write('N5', 'SOS', body)
    worksheet234.write('O5', 'FIS', body)
    worksheet234.write('P5', 'KIM', body)
    worksheet234.write('Q5', 'BIO', body)
    worksheet234.write('R5', 'JML', body)
    worksheet234.write('S5', 'MAW', body)
    worksheet234.write('T5', 'MAP', body)
    worksheet234.write('U5', 'IND', body)
    worksheet234.write('V5', 'ENG', body)
    worksheet234.write('W5', 'SEJ', body)
    worksheet234.write('X5', 'GEO', body)
    worksheet234.write('Y5', 'EKO', body)
    worksheet234.write('Z5', 'SOS', body)
    worksheet234.write('AA5', 'FIS', body)
    worksheet234.write('AB5', 'KIM', body)
    worksheet234.write('AC5', 'BIO', body)
    worksheet234.write('AD5', 'JML', body)

    worksheet234.conditional_format(5, 0, row234_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet234.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF KRAMAT JAYA', title)
    worksheet234.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet234.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet234.write('A22', 'LOKASI', header)
    worksheet234.write('B22', 'TOTAL', header)
    worksheet234.merge_range('A21:B21', 'RANK', header)
    worksheet234.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet234.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet234.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet234.merge_range('F21:F22', 'KELAS', header)
    worksheet234.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet234.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet234.write('G22', 'MAW', body)
    worksheet234.write('H22', 'MAP', body)
    worksheet234.write('I22', 'IND', body)
    worksheet234.write('J22', 'ENG', body)
    worksheet234.write('J22', 'SEJ', body)
    worksheet234.write('K22', 'GEO', body)
    worksheet234.write('M22', 'EKO', body)
    worksheet234.write('L22', 'SOS', body)
    worksheet234.write('L22', 'FIS', body)
    worksheet234.write('L22', 'KIM', body)
    worksheet234.write('L22', 'BIO', body)
    worksheet234.write('N22', 'JML', body)
    worksheet234.write('O22', 'MAW', body)
    worksheet234.write('O22', 'MAP', body)
    worksheet234.write('P22', 'IND', body)
    worksheet234.write('Q22', 'ENG', body)
    worksheet234.write('R22', 'SEJ', body)
    worksheet234.write('S22', 'GEO', body)
    worksheet234.write('U22', 'EKO', body)
    worksheet234.write('T22', 'SOS', body)
    worksheet234.write('T22', 'FIS', body)
    worksheet234.write('T22', 'KIM', body)
    worksheet234.write('T22', 'BIO', body)
    worksheet234.write('V22', 'JML', body)

    worksheet234.conditional_format(22, 0, row234+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 235
    worksheet235.insert_image('A1', r'logo resmi nf.jpg')

    worksheet235.set_column('A:A', 7, center)
    worksheet235.set_column('B:B', 6, center)
    worksheet235.set_column('C:C', 18.14, center)
    worksheet235.set_column('D:D', 25, left)
    worksheet235.set_column('E:E', 13.14, left)
    worksheet235.set_column('F:F', 8.57, center)
    worksheet235.set_column('G:V', 5, center)
    worksheet235.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF PONDOK GEDE', title)
    worksheet235.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet235.write('A5', 'LOKASI', header)
    worksheet235.write('B5', 'TOTAL', header)
    worksheet235.merge_range('A4:B4', 'RANK', header)
    worksheet235.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet235.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet235.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet235.merge_range('F4:F5', 'KELAS', header)
    worksheet235.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet235.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet235.write('G5', 'MAW', body)
    worksheet235.write('H5', 'MAP', body)
    worksheet235.write('I5', 'IND', body)
    worksheet235.write('J5', 'ENG', body)
    worksheet235.write('K5', 'SEJ', body)
    worksheet235.write('L5', 'GEO', body)
    worksheet235.write('M5', 'EKO', body)
    worksheet235.write('N5', 'SOS', body)
    worksheet235.write('O5', 'FIS', body)
    worksheet235.write('P5', 'KIM', body)
    worksheet235.write('Q5', 'BIO', body)
    worksheet235.write('R5', 'JML', body)
    worksheet235.write('S5', 'MAW', body)
    worksheet235.write('T5', 'MAP', body)
    worksheet235.write('U5', 'IND', body)
    worksheet235.write('V5', 'ENG', body)
    worksheet235.write('W5', 'SEJ', body)
    worksheet235.write('X5', 'GEO', body)
    worksheet235.write('Y5', 'EKO', body)
    worksheet235.write('Z5', 'SOS', body)
    worksheet235.write('AA5', 'FIS', body)
    worksheet235.write('AB5', 'KIM', body)
    worksheet235.write('AC5', 'BIO', body)
    worksheet235.write('AD5', 'JML', body)

    worksheet235.conditional_format(5, 0, row235_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet235.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF PONDOK GEDE', title)
    worksheet235.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet235.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet235.write('A22', 'LOKASI', header)
    worksheet235.write('B22', 'TOTAL', header)
    worksheet235.merge_range('A21:B21', 'RANK', header)
    worksheet235.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet235.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet235.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet235.merge_range('F21:F22', 'KELAS', header)
    worksheet235.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet235.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet235.write('G22', 'MAW', body)
    worksheet235.write('H22', 'MAP', body)
    worksheet235.write('I22', 'IND', body)
    worksheet235.write('J22', 'ENG', body)
    worksheet235.write('J22', 'SEJ', body)
    worksheet235.write('K22', 'GEO', body)
    worksheet235.write('M22', 'EKO', body)
    worksheet235.write('L22', 'SOS', body)
    worksheet235.write('L22', 'FIS', body)
    worksheet235.write('L22', 'KIM', body)
    worksheet235.write('L22', 'BIO', body)
    worksheet235.write('N22', 'JML', body)
    worksheet235.write('O22', 'MAW', body)
    worksheet235.write('O22', 'MAP', body)
    worksheet235.write('P22', 'IND', body)
    worksheet235.write('Q22', 'ENG', body)
    worksheet235.write('R22', 'SEJ', body)
    worksheet235.write('S22', 'GEO', body)
    worksheet235.write('U22', 'EKO', body)
    worksheet235.write('T22', 'SOS', body)
    worksheet235.write('T22', 'FIS', body)
    worksheet235.write('T22', 'KIM', body)
    worksheet235.write('T22', 'BIO', body)
    worksheet235.write('V22', 'JML', body)

    worksheet235.conditional_format(22, 0, row235+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 236
    worksheet236.insert_image('A1', r'logo resmi nf.jpg')

    worksheet236.set_column('A:A', 7, center)
    worksheet236.set_column('B:B', 6, center)
    worksheet236.set_column('C:C', 18.14, center)
    worksheet236.set_column('D:D', 25, left)
    worksheet236.set_column('E:E', 13.14, left)
    worksheet236.set_column('F:F', 8.57, center)
    worksheet236.set_column('G:V', 5, center)
    worksheet236.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF GALAXY', title)
    worksheet236.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet236.write('A5', 'LOKASI', header)
    worksheet236.write('B5', 'TOTAL', header)
    worksheet236.merge_range('A4:B4', 'RANK', header)
    worksheet236.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet236.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet236.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet236.merge_range('F4:F5', 'KELAS', header)
    worksheet236.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet236.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet236.write('G5', 'MAW', body)
    worksheet236.write('H5', 'MAP', body)
    worksheet236.write('I5', 'IND', body)
    worksheet236.write('J5', 'ENG', body)
    worksheet236.write('K5', 'SEJ', body)
    worksheet236.write('L5', 'GEO', body)
    worksheet236.write('M5', 'EKO', body)
    worksheet236.write('N5', 'SOS', body)
    worksheet236.write('O5', 'FIS', body)
    worksheet236.write('P5', 'KIM', body)
    worksheet236.write('Q5', 'BIO', body)
    worksheet236.write('R5', 'JML', body)
    worksheet236.write('S5', 'MAW', body)
    worksheet236.write('T5', 'MAP', body)
    worksheet236.write('U5', 'IND', body)
    worksheet236.write('V5', 'ENG', body)
    worksheet236.write('W5', 'SEJ', body)
    worksheet236.write('X5', 'GEO', body)
    worksheet236.write('Y5', 'EKO', body)
    worksheet236.write('Z5', 'SOS', body)
    worksheet236.write('AA5', 'FIS', body)
    worksheet236.write('AB5', 'KIM', body)
    worksheet236.write('AC5', 'BIO', body)
    worksheet236.write('AD5', 'JML', body)

    worksheet236.conditional_format(5, 0, row236_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet236.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF GALAXY', title)
    worksheet236.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet236.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet236.write('A22', 'LOKASI', header)
    worksheet236.write('B22', 'TOTAL', header)
    worksheet236.merge_range('A21:B21', 'RANK', header)
    worksheet236.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet236.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet236.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet236.merge_range('F21:F22', 'KELAS', header)
    worksheet236.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet236.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet236.write('G22', 'MAW', body)
    worksheet236.write('H22', 'MAP', body)
    worksheet236.write('I22', 'IND', body)
    worksheet236.write('J22', 'ENG', body)
    worksheet236.write('J22', 'SEJ', body)
    worksheet236.write('K22', 'GEO', body)
    worksheet236.write('M22', 'EKO', body)
    worksheet236.write('L22', 'SOS', body)
    worksheet236.write('L22', 'FIS', body)
    worksheet236.write('L22', 'KIM', body)
    worksheet236.write('L22', 'BIO', body)
    worksheet236.write('N22', 'JML', body)
    worksheet236.write('O22', 'MAW', body)
    worksheet236.write('O22', 'MAP', body)
    worksheet236.write('P22', 'IND', body)
    worksheet236.write('Q22', 'ENG', body)
    worksheet236.write('R22', 'SEJ', body)
    worksheet236.write('S22', 'GEO', body)
    worksheet236.write('U22', 'EKO', body)
    worksheet236.write('T22', 'SOS', body)
    worksheet236.write('T22', 'FIS', body)
    worksheet236.write('T22', 'KIM', body)
    worksheet236.write('T22', 'BIO', body)
    worksheet236.write('V22', 'JML', body)

    worksheet236.conditional_format(22, 0, row236+21, 21,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 160
    # worksheet160.insert_image('A1', r'logo resmi nf.jpg')

    # worksheet160.set_column('A:A', 7, center)
    # worksheet160.set_column('B:B', 6, center)
    # worksheet160.set_column('C:C', 18.14, center)
    # worksheet160.set_column('D:D', 25, left)
    # worksheet160.set_column('E:E', 13.14, left)
    # worksheet160.set_column('F:F', 8.57, center)
    # worksheet160.set_column('G:V', 5, center)
    # worksheet160.merge_range(
    #     'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF CIGANJUR', title)
    # worksheet160.merge_range(
    #     'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    # worksheet160.write('A5', 'LOKASI', header)
    # worksheet160.write('B5', 'TOTAL', header)
    # worksheet160.merge_range('A4:B4', 'RANK', header)
    # worksheet160.merge_range('C4:C5', 'NOMOR NF', header)
    # worksheet160.merge_range('D4:D5', 'NAMA SISWA', header)
    # worksheet160.merge_range('E4:E5', 'SEKOLAH', header)
    # worksheet160.merge_range('F4:F5', 'KELAS', header)
    # worksheet160.merge_range('G4:R4', 'JUMLAH BENAR', header)
    # worksheet160.merge_range('S4:AD4', 'NILAI STANDAR', header)
    # worksheet160.write('G5', 'MAW', body)
    # worksheet160.write('H5', 'MAP', body)
    # worksheet160.write('I5', 'IND', body)
    # worksheet160.write('J5', 'ENG', body)
    # worksheet160.write('K5', 'SEJ', body)
    # worksheet160.write('L5', 'GEO', body)
    # worksheet160.write('M5', 'EKO', body)
    # worksheet160.write('N5', 'SOS', body)
    # worksheet160.write('O5', 'FIS', body)
    # worksheet160.write('P5', 'KIM', body)
    # worksheet160.write('Q5', 'BIO', body)
    # worksheet160.write('R5', 'JML', body)
    # worksheet160.write('S5', 'MAW', body)
    # worksheet160.write('T5', 'MAP', body)
    # worksheet160.write('U5', 'IND', body)
    # worksheet160.write('V5', 'ENG', body)
    # worksheet160.write('W5', 'SEJ', body)
    # worksheet160.write('X5', 'GEO', body)
    # worksheet160.write('Y5', 'EKO', body)
    # worksheet160.write('Z5', 'SOS', body)
    # worksheet160.write('AA5', 'FIS', body)
    # worksheet160.write('AB5', 'KIM', body)
    # worksheet160.write('AC5', 'BIO', body)
    # worksheet160.write('AD5', 'JML', body)

    # worksheet160.conditional_format(5, 0, row160_10+4, 29,
    #                                 {'type': 'no_errors', 'format': border})

    # worksheet160.merge_range(
    #     'A17:AD17', fr'KELAS {kelas} - LOKASI NF CIGANJUR', title)
    # worksheet160.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    # worksheet160.merge_range(
    #     'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    # worksheet160.write('A22', 'LOKASI', header)
    # worksheet160.write('B22', 'TOTAL', header)
    # worksheet160.merge_range('A21:B21', 'RANK', header)
    # worksheet160.merge_range('C21:C22', 'NOMOR NF', header)
    # worksheet160.merge_range('D21:D22', 'NAMA SISWA', header)
    # worksheet160.merge_range('E21:E22', 'SEKOLAH', header)
    # worksheet160.merge_range('F21:F22', 'KELAS', header)
    # worksheet160.merge_range('G21:R21', 'JUMLAH BENAR', header)
    # worksheet160.merge_range('S21:AD21', 'NILAI STANDAR', header)
    # worksheet160.write('G22', 'MAW', body)
    # worksheet160.write('H22', 'MAP', body)
    # worksheet160.write('I22', 'IND', body)
    # worksheet160.write('J22', 'ENG', body)
    # worksheet160.write('J22', 'SEJ', body)
    # worksheet160.write('K22', 'GEO', body)
    # worksheet160.write('M22', 'EKO', body)
    # worksheet160.write('L22', 'SOS', body)
    # worksheet160.write('L22', 'FIS', body)
    # worksheet160.write('L22', 'KIM', body)
    # worksheet160.write('L22', 'BIO', body)
    # worksheet160.write('N22', 'JML', body)
    # worksheet160.write('O22', 'MAW', body)
    # worksheet160.write('O22', 'MAP', body)
    # worksheet160.write('P22', 'IND', body)
    # worksheet160.write('Q22', 'ENG', body)
    # worksheet160.write('R22', 'SEJ', body)
    # worksheet160.write('S22', 'GEO', body)
    # worksheet160.write('U22', 'EKO', body)
    # worksheet160.write('T22', 'SOS', body)
    # worksheet160.write('T22', 'FIS', body)
    # worksheet160.write('T22', 'KIM', body)
    # worksheet160.write('T22', 'BIO', body)
    # worksheet160.write('V22', 'JML', body)

    # worksheet160.conditional_format(22, 0, row160+21, 21,
    #                                 {'type': 'no_errors', 'format': border})

    workbook.close()
    st.success("File siap diunduh!")

    # Tombol unduh file
    with open(file_path, "rb") as f:
        bytes_data = f.read()
    st.download_button(label="Unduh File", data=bytes_data,
                       file_name=file_name)