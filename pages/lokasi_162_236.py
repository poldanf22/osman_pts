# NAMA LOKASI DEPOK-PADANG
lok_530 = 'POLSEK DEPOK'
lok_531 = 'DEPOK 1'
lok_532 = 'PROKLAMASI DEPOK 2'
lok_533 = 'CIMANGGIS'
lok_243 = 'PURBALINGGA'
lok_244 = 'SURAKARTA'
lok_245 = 'SEMARANG'
lok_246 = 'KARTASURA'
lok_248 = 'BSD BOULEVARD'
lok_249 = 'SANGIANG'
lok_250 = 'BANJAR WIJAYA'
lok_252 = 'CIRENDEU'
lok_254 = 'GRAHA RAYA'
lok_255 = 'MERPATI'
lok_256 = 'CIRUAS'
lok_258 = 'LHOSEUMAWE'
lok_259 = 'PANAM, PKU'
lok_260 = 'AM. SANGAJI'
lok_261 = 'DURI KOSAMBI'
lok_262 = 'CITRA RAYA CIKUPA'
lok_263 = 'GRAHA PRIMA'
lok_264 = 'KARAWANG'
lok_265 = 'TAMAN WISMA ASRI'
lok_266 = 'MANGUN JAYA'
lok_267 = 'MARAKASH / SEKTOR 5'
lok_268 = 'KEBALEN'
lok_269 = 'JATI RANGON'
lok_270 = 'JATIBENING'
lok_271 = 'JATIMULYA'
lok_272 = 'PERUMNAS 3'
lok_273 = 'NAROGONG'
lok_274 = 'BEKASI TIMUR REGENCY'
lok_275 = 'CIKARANG PILAR'
lok_276 = 'JABABEKA'
lok_277 = 'PAYAKUMBUH'
lok_278 = 'MERDUATI'
lok_279 = 'ANTAPANI'
lok_280 = 'MARGAHAYU'
lok_282 = 'PAHLAWAN'
lok_283 = 'CIJERAH'
lok_284 = 'TEGAL'
lok_285 = 'MEDAN AREA'
lok_286 = 'MEDAN JOHOR'
lok_287 = 'JAMBO TAPE'
lok_288 = 'THE HOK'
lok_289 = 'SAIL'
lok_290 = 'TELANAI JAMBI'
lok_291 = 'SIDOARJO'
lok_292 = 'PURWOKERTO LOR'
lok_293 = 'WAY HALIM'
lok_294 = 'METRO'
lok_295 = 'RAJABASA'
lok_298 = 'PAHOMAN'
lok_299 = 'KIMILING'

uploaded_file = st.file_uploader(
    'Letakkan file excel NILAI STANDAR [LOKASI 530-299]', type='xlsx')

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
    rata_Smaw = df.iloc[t, len_col-25]
    rata_Smap = df.iloc[t, len_col-24]
    rata_Sind = df.iloc[t, len_col-23]
    rata_Seng = df.iloc[t, len_col-22]
    rata_Ssej = df.iloc[t, len_col-21]
    rata_Sgeo = df.iloc[t, len_col-20]
    rata_Seko = df.iloc[t, len_col-19]
    rata_Ssos = df.iloc[t, len_col-18]
    rata_Sfis = df.iloc[t, len_col-17]
    rata_Skim = df.iloc[t, len_col-16]
    rata_Sbio = df.iloc[t, len_col-15]
    rata_Sjml = df.iloc[t, len_col-14]

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
    max_Smaw = df.iloc[r, len_col-25]
    max_Smap = df.iloc[r, len_col-24]
    max_Sind = df.iloc[r, len_col-23]
    max_Seng = df.iloc[r, len_col-22]
    max_Ssej = df.iloc[r, len_col-21]
    max_Sgeo = df.iloc[r, len_col-20]
    max_Seko = df.iloc[r, len_col-19]
    max_Ssos = df.iloc[r, len_col-18]
    max_Sfis = df.iloc[r, len_col-17]
    max_Skim = df.iloc[r, len_col-16]
    max_Sbio = df.iloc[r, len_col-15]
    max_Sjml = df.iloc[r, len_col-14]

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
    min_Smaw = df.iloc[s, len_col-25]
    min_Smap = df.iloc[s, len_col-24]
    min_Sind = df.iloc[s, len_col-23]
    min_Seng = df.iloc[s, len_col-22]
    min_Ssej = df.iloc[s, len_col-21]
    min_Sgeo = df.iloc[s, len_col-20]
    min_Seko = df.iloc[s, len_col-19]
    min_Ssos = df.iloc[s, len_col-18]
    min_Sfis = df.iloc[s, len_col-17]
    min_Skim = df.iloc[s, len_col-16]
    min_Sbio = df.iloc[s, len_col-15]
    min_Sjml = df.iloc[s, len_col-14]

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
    sort530 = df[df['LOKASI'] == 530]
    sort531 = df[df['LOKASI'] == 531]
    sort532 = df[df['LOKASI'] == 532]
    sort533 = df[df['LOKASI'] == 533]
    sort243 = df[df['LOKASI'] == 243]
    sort244 = df[df['LOKASI'] == 244]
    sort245 = df[df['LOKASI'] == 245]
    sort246 = df[df['LOKASI'] == 246]
    sort248 = df[df['LOKASI'] == 248]
    sort249 = df[df['LOKASI'] == 249]
    sort250 = df[df['LOKASI'] == 250]
    sort252 = df[df['LOKASI'] == 252]
    sort254 = df[df['LOKASI'] == 254]
    sort255 = df[df['LOKASI'] == 255]
    sort256 = df[df['LOKASI'] == 256]
    sort258 = df[df['LOKASI'] == 258]
    sort259 = df[df['LOKASI'] == 259]
    sort260 = df[df['LOKASI'] == 260]
    sort261 = df[df['LOKASI'] == 261]
    sort262 = df[df['LOKASI'] == 262]
    sort263 = df[df['LOKASI'] == 263]
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
    # sort139 = df[df['LOKASI']==139]
    sort279 = df[df['LOKASI'] == 279]
    sort280 = df[df['LOKASI'] == 280]
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
    sort298 = df[df['LOKASI'] == 298]
    sort299 = df[df['LOKASI'] == 299]
    # sort236 = df[df['LOKASI'] == 236]
    # sort160 = df[df['LOKASI'] == 160]

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
    # # 104
    # sort104_10=sort104.sort_values(by=['RANK LOK.'], ascending=[True])
    # del sort104_10['LOKASI']
    # sort104_10=sort104_10.drop(sort104_10[(sort104_10['RANK LOK.']>10)].index)
    # 533
    sort533_10 = sort533.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort533_10['LOKASI']
    sort533_10 = sort533_10.drop(
        sort533_10[(sort533_10['RANK LOK.'] > 10)].index)
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
    # 252
    sort252_10 = sort252.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort252_10['LOKASI']
    sort252_10 = sort252_10.drop(
        sort252_10[(sort252_10['RANK LOK.'] > 10)].index)
    # # 114
    # sort114_10=sort114.sort_values(by=['RANK LOK.'], ascending=[True])
    # del sort114_10['LOKASI']
    # sort114_10=sort114_10.drop(sort114_10[(sort114_10['RANK LOK.']>10)].index)
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
    # 263
    sort263_10 = sort263.sort_values(by=['RANK LOK.'], ascending=[True])
    del sort263_10['LOKASI']
    sort263_10 = sort263_10.drop(
        sort263_10[(sort263_10['RANK LOK.'] > 10)].index)
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
    # # 139
    # sort139_10=sort139.sort_values(by=['RANK LOK.'], ascending=[True])
    # del sort139_10['LOKASI']
    # sort139_10=sort139_10.drop(sort139_10[(sort139_10['RANK LOK.']>10)].index)
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
    # 236
    # sort236_10 = sort236.sort_values(by=['RANK LOK.'], ascending=[True])
    # del sort236_10['LOKASI']
    # sort236_10 = sort236_10.drop(
    #     sort236_10[(sort236_10['RANK LOK.'] > 10)].index)
    # 160
    # sort160_10 = sort160.sort_values(by=['RANK LOK.'], ascending=[True])
    # del sort160_10['LOKASI']
    # sort160_10 = sort160_10.drop(
    #     sort160_10[(sort160_10['RANK LOK.'] > 10)].index)

    # All 530
    sort530 = sort530.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort530['LOKASI']
    # All 531
    sort531 = sort531.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort531['LOKASI']
    # All 532
    sort532 = sort532.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort532['LOKASI']
    # # All 104
    # sort104=sort104.sort_values(by=['NAMA SISWA'], ascending=[True])
    # del sort104['LOKASI']
    # All 533
    sort533 = sort533.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort533['LOKASI']
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
    # All 248
    sort248 = sort248.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort248['LOKASI']
    # All 249
    sort249 = sort249.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort249['LOKASI']
    # All 250
    sort250 = sort250.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort250['LOKASI']
    # All 252
    sort252 = sort252.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort252['LOKASI']
    # # All 114
    # sort114=sort114.sort_values(by=['NAMA SISWA'], ascending=[True])
    # del sort114['LOKASI']
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
    # All 263
    sort263 = sort263.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort263['LOKASI']
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
    # # All 139
    # sort139=sort139.sort_values(by=['NAMA SISWA'], ascending=[True])
    # del sort139['LOKASI']
    # All 279
    sort279 = sort279.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort279['LOKASI']
    # All 280
    sort280 = sort280.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort280['LOKASI']
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
    # All 298
    sort298 = sort298.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort298['LOKASI']
    # All 299
    sort299 = sort299.sort_values(by=['NAMA SISWA'], ascending=[True])
    del sort299['LOKASI']
    # All 236
    # sort236 = sort236.sort_values(by=['NAMA SISWA'], ascending=[True])
    # del sort236['LOKASI']
    # All 160
    # sort160 = sort160.sort_values(by=['NAMA SISWA'], ascending=[True])
    # del sort160['LOKASI']

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
    # # 104
    # row104_10=sort104_10.shape[0]
    # row104=sort104.shape[0]
    # 533
    row533_10 = sort533_10.shape[0]
    row533 = sort533.shape[0]
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
    # 248
    row248_10 = sort248_10.shape[0]
    row248 = sort248.shape[0]
    # 249
    row249_10 = sort249_10.shape[0]
    row249 = sort249.shape[0]
    # 250
    row250_10 = sort250_10.shape[0]
    row250 = sort250.shape[0]
    # 252
    row252_10 = sort252_10.shape[0]
    row252 = sort252.shape[0]
    # # 114
    # row114_10=sort114_10.shape[0]
    # row114=sort114.shape[0]
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
    # 263
    row263_10 = sort263_10.shape[0]
    row263 = sort263.shape[0]
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
    # # 139
    # row139_10=sort139_10.shape[0]
    # row139=sort139.shape[0]
    # 279
    row279_10 = sort279_10.shape[0]
    row279 = sort279.shape[0]
    # 280
    row280_10 = sort280_10.shape[0]
    row280 = sort280.shape[0]
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
    # 298
    row298_10 = sort298_10.shape[0]
    row298 = sort298.shape[0]
    # 299
    row299_10 = sort299_10.shape[0]
    row299 = sort299.shape[0]
    # 236
    # row236_10 = sort236_10.shape[0]
    # row236 = sort236.shape[0]
    # 160
    # row160_10 = sort160_10.shape[0]
    # row160 = sort160.shape[0]

    # Create a Pandas Excel writer using XlsxWriter as the engine.
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
    # 263
    # Convert the dataframe to an XlsxWriter Excel object.
    sort263_10.to_excel(writer, sheet_name='263',
                        startrow=5,
                        startcol=0,
                        index=False,
                        header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    sort263.to_excel(writer, sheet_name='263',
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
    # 236
    # Convert the dataframe to an XlsxWriter Excel object.
    # sort236_10.to_excel(writer, sheet_name='236',
    #                     startrow=5,
    #                     startcol=0,
    #                     index=False,
    #                     header=False)
    # Convert the dataframe to an XlsxWriter Excel object.
    # sort236.to_excel(writer, sheet_name='236',
    #                  startrow=22,
    #                  startcol=0,
    #                  index=False,
    #                  header=False)
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
    worksheet530 = writer.sheets['530']
    worksheet531 = writer.sheets['531']
    worksheet532 = writer.sheets['532']
    # worksheet104 = writer.sheets['104']
    worksheet533 = writer.sheets['533']
    worksheet243 = writer.sheets['243']
    worksheet244 = writer.sheets['244']
    worksheet245 = writer.sheets['245']
    worksheet246 = writer.sheets['246']
    worksheet248 = writer.sheets['248']
    worksheet249 = writer.sheets['249']
    worksheet250 = writer.sheets['250']
    worksheet252 = writer.sheets['252']
    # worksheet114 = writer.sheets['114']
    worksheet254 = writer.sheets['254']
    worksheet255 = writer.sheets['255']
    worksheet256 = writer.sheets['256']
    worksheet258 = writer.sheets['258']
    worksheet259 = writer.sheets['259']
    worksheet260 = writer.sheets['260']
    worksheet261 = writer.sheets['261']
    worksheet262 = writer.sheets['262']
    worksheet263 = writer.sheets['263']
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
    # worksheet139 = writer.sheets['139']
    worksheet279 = writer.sheets['279']
    worksheet280 = writer.sheets['280']
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
    worksheet298 = writer.sheets['298']
    worksheet299 = writer.sheets['299']
    # worksheet236 = writer.sheets['236']
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

    # worksheet 530
    worksheet530.insert_image('A1', r'logo resmi nf.jpg')

    worksheet530.set_column('A:A', 7, center)
    worksheet530.set_column('B:B', 6, center)
    worksheet530.set_column('C:C', 18.14, center)
    worksheet530.set_column('D:D', 25, left)
    worksheet530.set_column('E:E', 13.14, left)
    worksheet530.set_column('F:F', 8.57, center)
    worksheet530.set_column('G:AD', 5, center)
    worksheet530.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_530}', title)
    worksheet530.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet530.write('A5', 'LOKASI', header)
    worksheet530.write('B5', 'TOTAL', header)
    worksheet530.merge_range('A4:B4', 'RANK', header)
    worksheet530.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet530.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet530.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet530.merge_range('F4:F5', 'KELAS', header)
    worksheet530.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet530.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet530.write('G5', 'MAW', body)
    worksheet530.write('H5', 'MAP', body)
    worksheet530.write('I5', 'IND', body)
    worksheet530.write('J5', 'ENG', body)
    worksheet530.write('K5', 'SEJ', body)
    worksheet530.write('L5', 'GEO', body)
    worksheet530.write('M5', 'EKO', body)
    worksheet530.write('N5', 'SOS', body)
    worksheet530.write('O5', 'FIS', body)
    worksheet530.write('P5', 'KIM', body)
    worksheet530.write('Q5', 'BIO', body)
    worksheet530.write('R5', 'JML', body)
    worksheet530.write('S5', 'MAW', body)
    worksheet530.write('T5', 'MAP', body)
    worksheet530.write('U5', 'IND', body)
    worksheet530.write('V5', 'ENG', body)
    worksheet530.write('W5', 'SEJ', body)
    worksheet530.write('X5', 'GEO', body)
    worksheet530.write('Y5', 'EKO', body)
    worksheet530.write('Z5', 'SOS', body)
    worksheet530.write('AA5', 'FIS', body)
    worksheet530.write('AB5', 'KIM', body)
    worksheet530.write('AC5', 'BIO', body)
    worksheet530.write('AD5', 'JML', body)

    worksheet530.conditional_format(5, 0, row530_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet530.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_530}', title)
    worksheet530.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet530.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet530.write('A22', 'LOKASI', header)
    worksheet530.write('B22', 'TOTAL', header)
    worksheet530.merge_range('A21:B21', 'RANK', header)
    worksheet530.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet530.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet530.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet530.merge_range('F21:F22', 'KELAS', header)
    worksheet530.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet530.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet530.write('G22', 'MAW', body)
    worksheet530.write('H22', 'MAP', body)
    worksheet530.write('I22', 'IND', body)
    worksheet530.write('J22', 'ENG', body)
    worksheet530.write('K22', 'SEJ', body)
    worksheet530.write('L22', 'GEO', body)
    worksheet530.write('M22', 'EKO', body)
    worksheet530.write('N22', 'SOS', body)
    worksheet530.write('O22', 'FIS', body)
    worksheet530.write('P22', 'KIM', body)
    worksheet530.write('Q22', 'BIO', body)
    worksheet530.write('R22', 'JML', body)
    worksheet530.write('S22', 'MAW', body)
    worksheet530.write('T22', 'MAP', body)
    worksheet530.write('U22', 'IND', body)
    worksheet530.write('V22', 'ENG', body)
    worksheet530.write('W22', 'SEJ', body)
    worksheet530.write('X22', 'GEO', body)
    worksheet530.write('Y22', 'EKO', body)
    worksheet530.write('Z22', 'SOS', body)
    worksheet530.write('AA22', 'FIS', body)
    worksheet530.write('AB22', 'KIM', body)
    worksheet530.write('AC22', 'BIO', body)
    worksheet530.write('AD22', 'JML', body)

    worksheet530.conditional_format(22, 0, row530+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 531
    worksheet531.insert_image('A1', r'logo resmi nf.jpg')

    worksheet531.set_column('A:A', 7, center)
    worksheet531.set_column('B:B', 6, center)
    worksheet531.set_column('C:C', 18.14, center)
    worksheet531.set_column('D:D', 25, left)
    worksheet531.set_column('E:E', 13.14, left)
    worksheet531.set_column('F:F', 8.57, center)
    worksheet531.set_column('G:AD', 5, center)
    worksheet531.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_531}', title)
    worksheet531.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet531.write('A5', 'LOKASI', header)
    worksheet531.write('B5', 'TOTAL', header)
    worksheet531.merge_range('A4:B4', 'RANK', header)
    worksheet531.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet531.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet531.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet531.merge_range('F4:F5', 'KELAS', header)
    worksheet531.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet531.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet531.write('G5', 'MAW', body)
    worksheet531.write('H5', 'MAP', body)
    worksheet531.write('I5', 'IND', body)
    worksheet531.write('J5', 'ENG', body)
    worksheet531.write('K5', 'SEJ', body)
    worksheet531.write('L5', 'GEO', body)
    worksheet531.write('M5', 'EKO', body)
    worksheet531.write('N5', 'SOS', body)
    worksheet531.write('O5', 'FIS', body)
    worksheet531.write('P5', 'KIM', body)
    worksheet531.write('Q5', 'BIO', body)
    worksheet531.write('R5', 'JML', body)
    worksheet531.write('S5', 'MAW', body)
    worksheet531.write('T5', 'MAP', body)
    worksheet531.write('U5', 'IND', body)
    worksheet531.write('V5', 'ENG', body)
    worksheet531.write('W5', 'SEJ', body)
    worksheet531.write('X5', 'GEO', body)
    worksheet531.write('Y5', 'EKO', body)
    worksheet531.write('Z5', 'SOS', body)
    worksheet531.write('AA5', 'FIS', body)
    worksheet531.write('AB5', 'KIM', body)
    worksheet531.write('AC5', 'BIO', body)
    worksheet531.write('AD5', 'JML', body)

    worksheet531.conditional_format(5, 0, row531_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet531.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_531}', title)
    worksheet531.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet531.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet531.write('A22', 'LOKASI', header)
    worksheet531.write('B22', 'TOTAL', header)
    worksheet531.merge_range('A21:B21', 'RANK', header)
    worksheet531.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet531.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet531.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet531.merge_range('F21:F22', 'KELAS', header)
    worksheet531.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet531.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet531.write('G22', 'MAW', body)
    worksheet531.write('H22', 'MAP', body)
    worksheet531.write('I22', 'IND', body)
    worksheet531.write('J22', 'ENG', body)
    worksheet531.write('K22', 'SEJ', body)
    worksheet531.write('L22', 'GEO', body)
    worksheet531.write('M22', 'EKO', body)
    worksheet531.write('N22', 'SOS', body)
    worksheet531.write('O22', 'FIS', body)
    worksheet531.write('P22', 'KIM', body)
    worksheet531.write('Q22', 'BIO', body)
    worksheet531.write('R22', 'JML', body)
    worksheet531.write('S22', 'MAW', body)
    worksheet531.write('T22', 'MAP', body)
    worksheet531.write('U22', 'IND', body)
    worksheet531.write('V22', 'ENG', body)
    worksheet531.write('W22', 'SEJ', body)
    worksheet531.write('X22', 'GEO', body)
    worksheet531.write('Y22', 'EKO', body)
    worksheet531.write('Z22', 'SOS', body)
    worksheet531.write('AA22', 'FIS', body)
    worksheet531.write('AB22', 'KIM', body)
    worksheet531.write('AC22', 'BIO', body)
    worksheet531.write('AD22', 'JML', body)

    worksheet531.conditional_format(22, 0, row531+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 532
    worksheet532.insert_image('A1', r'logo resmi nf.jpg')

    worksheet532.set_column('A:A', 7, center)
    worksheet532.set_column('B:B', 6, center)
    worksheet532.set_column('C:C', 18.14, center)
    worksheet532.set_column('D:D', 25, left)
    worksheet532.set_column('E:E', 13.14, left)
    worksheet532.set_column('F:F', 8.57, center)
    worksheet532.set_column('G:AD', 5, center)
    worksheet532.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_532}', title)
    worksheet532.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet532.write('A5', 'LOKASI', header)
    worksheet532.write('B5', 'TOTAL', header)
    worksheet532.merge_range('A4:B4', 'RANK', header)
    worksheet532.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet532.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet532.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet532.merge_range('F4:F5', 'KELAS', header)
    worksheet532.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet532.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet532.write('G5', 'MAW', body)
    worksheet532.write('H5', 'MAP', body)
    worksheet532.write('I5', 'IND', body)
    worksheet532.write('J5', 'ENG', body)
    worksheet532.write('K5', 'SEJ', body)
    worksheet532.write('L5', 'GEO', body)
    worksheet532.write('M5', 'EKO', body)
    worksheet532.write('N5', 'SOS', body)
    worksheet532.write('O5', 'FIS', body)
    worksheet532.write('P5', 'KIM', body)
    worksheet532.write('Q5', 'BIO', body)
    worksheet532.write('R5', 'JML', body)
    worksheet532.write('S5', 'MAW', body)
    worksheet532.write('T5', 'MAP', body)
    worksheet532.write('U5', 'IND', body)
    worksheet532.write('V5', 'ENG', body)
    worksheet532.write('W5', 'SEJ', body)
    worksheet532.write('X5', 'GEO', body)
    worksheet532.write('Y5', 'EKO', body)
    worksheet532.write('Z5', 'SOS', body)
    worksheet532.write('AA5', 'FIS', body)
    worksheet532.write('AB5', 'KIM', body)
    worksheet532.write('AC5', 'BIO', body)
    worksheet532.write('AD5', 'JML', body)

    worksheet532.conditional_format(5, 0, row532_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet532.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_532}', title)
    worksheet532.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet532.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet532.write('A22', 'LOKASI', header)
    worksheet532.write('B22', 'TOTAL', header)
    worksheet532.merge_range('A21:B21', 'RANK', header)
    worksheet532.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet532.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet532.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet532.merge_range('F21:F22', 'KELAS', header)
    worksheet532.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet532.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet532.write('G22', 'MAW', body)
    worksheet532.write('H22', 'MAP', body)
    worksheet532.write('I22', 'IND', body)
    worksheet532.write('J22', 'ENG', body)
    worksheet532.write('K22', 'SEJ', body)
    worksheet532.write('L22', 'GEO', body)
    worksheet532.write('M22', 'EKO', body)
    worksheet532.write('N22', 'SOS', body)
    worksheet532.write('O22', 'FIS', body)
    worksheet532.write('P22', 'KIM', body)
    worksheet532.write('Q22', 'BIO', body)
    worksheet532.write('R22', 'JML', body)
    worksheet532.write('S22', 'MAW', body)
    worksheet532.write('T22', 'MAP', body)
    worksheet532.write('U22', 'IND', body)
    worksheet532.write('V22', 'ENG', body)
    worksheet532.write('W22', 'SEJ', body)
    worksheet532.write('X22', 'GEO', body)
    worksheet532.write('Y22', 'EKO', body)
    worksheet532.write('Z22', 'SOS', body)
    worksheet532.write('AA22', 'FIS', body)
    worksheet532.write('AB22', 'KIM', body)
    worksheet532.write('AC22', 'BIO', body)
    worksheet532.write('AD22', 'JML', body)

    worksheet532.conditional_format(22, 0, row532+21, 29,
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

    # worksheet 533
    worksheet533.insert_image('A1', r'logo resmi nf.jpg')

    worksheet533.set_column('A:A', 7, center)
    worksheet533.set_column('B:B', 6, center)
    worksheet533.set_column('C:C', 18.14, center)
    worksheet533.set_column('D:D', 25, left)
    worksheet533.set_column('E:E', 13.14, left)
    worksheet533.set_column('F:F', 8.57, center)
    worksheet533.set_column('G:AD', 5, center)
    worksheet533.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_533}', title)
    worksheet533.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet533.write('A5', 'LOKASI', header)
    worksheet533.write('B5', 'TOTAL', header)
    worksheet533.merge_range('A4:B4', 'RANK', header)
    worksheet533.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet533.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet533.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet533.merge_range('F4:F5', 'KELAS', header)
    worksheet533.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet533.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet533.write('G5', 'MAW', body)
    worksheet533.write('H5', 'MAP', body)
    worksheet533.write('I5', 'IND', body)
    worksheet533.write('J5', 'ENG', body)
    worksheet533.write('K5', 'SEJ', body)
    worksheet533.write('L5', 'GEO', body)
    worksheet533.write('M5', 'EKO', body)
    worksheet533.write('N5', 'SOS', body)
    worksheet533.write('O5', 'FIS', body)
    worksheet533.write('P5', 'KIM', body)
    worksheet533.write('Q5', 'BIO', body)
    worksheet533.write('R5', 'JML', body)
    worksheet533.write('S5', 'MAW', body)
    worksheet533.write('T5', 'MAP', body)
    worksheet533.write('U5', 'IND', body)
    worksheet533.write('V5', 'ENG', body)
    worksheet533.write('W5', 'SEJ', body)
    worksheet533.write('X5', 'GEO', body)
    worksheet533.write('Y5', 'EKO', body)
    worksheet533.write('Z5', 'SOS', body)
    worksheet533.write('AA5', 'FIS', body)
    worksheet533.write('AB5', 'KIM', body)
    worksheet533.write('AC5', 'BIO', body)
    worksheet533.write('AD5', 'JML', body)

    worksheet533.conditional_format(5, 0, row533_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet533.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_533}', title)
    worksheet533.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet533.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet533.write('A22', 'LOKASI', header)
    worksheet533.write('B22', 'TOTAL', header)
    worksheet533.merge_range('A21:B21', 'RANK', header)
    worksheet533.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet533.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet533.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet533.merge_range('F21:F22', 'KELAS', header)
    worksheet533.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet533.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet533.write('G22', 'MAW', body)
    worksheet533.write('H22', 'MAP', body)
    worksheet533.write('I22', 'IND', body)
    worksheet533.write('J22', 'ENG', body)
    worksheet533.write('K22', 'SEJ', body)
    worksheet533.write('L22', 'GEO', body)
    worksheet533.write('M22', 'EKO', body)
    worksheet533.write('N22', 'SOS', body)
    worksheet533.write('O22', 'FIS', body)
    worksheet533.write('P22', 'KIM', body)
    worksheet533.write('Q22', 'BIO', body)
    worksheet533.write('R22', 'JML', body)
    worksheet533.write('S22', 'MAW', body)
    worksheet533.write('T22', 'MAP', body)
    worksheet533.write('U22', 'IND', body)
    worksheet533.write('V22', 'ENG', body)
    worksheet533.write('W22', 'SEJ', body)
    worksheet533.write('X22', 'GEO', body)
    worksheet533.write('Y22', 'EKO', body)
    worksheet533.write('Z22', 'SOS', body)
    worksheet533.write('AA22', 'FIS', body)
    worksheet533.write('AB22', 'KIM', body)
    worksheet533.write('AC22', 'BIO', body)
    worksheet533.write('AD22', 'JML', body)

    worksheet533.conditional_format(22, 0, row533+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 243
    worksheet243.insert_image('A1', r'logo resmi nf.jpg')

    worksheet243.set_column('A:A', 7, center)
    worksheet243.set_column('B:B', 6, center)
    worksheet243.set_column('C:C', 18.14, center)
    worksheet243.set_column('D:D', 25, left)
    worksheet243.set_column('E:E', 13.14, left)
    worksheet243.set_column('F:F', 8.57, center)
    worksheet243.set_column('G:AD', 5, center)
    worksheet243.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_243}', title)
    worksheet243.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet243.write('A5', 'LOKASI', header)
    worksheet243.write('B5', 'TOTAL', header)
    worksheet243.merge_range('A4:B4', 'RANK', header)
    worksheet243.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet243.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet243.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet243.merge_range('F4:F5', 'KELAS', header)
    worksheet243.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet243.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet243.write('G5', 'MAW', body)
    worksheet243.write('H5', 'MAP', body)
    worksheet243.write('I5', 'IND', body)
    worksheet243.write('J5', 'ENG', body)
    worksheet243.write('K5', 'SEJ', body)
    worksheet243.write('L5', 'GEO', body)
    worksheet243.write('M5', 'EKO', body)
    worksheet243.write('N5', 'SOS', body)
    worksheet243.write('O5', 'FIS', body)
    worksheet243.write('P5', 'KIM', body)
    worksheet243.write('Q5', 'BIO', body)
    worksheet243.write('R5', 'JML', body)
    worksheet243.write('S5', 'MAW', body)
    worksheet243.write('T5', 'MAP', body)
    worksheet243.write('U5', 'IND', body)
    worksheet243.write('V5', 'ENG', body)
    worksheet243.write('W5', 'SEJ', body)
    worksheet243.write('X5', 'GEO', body)
    worksheet243.write('Y5', 'EKO', body)
    worksheet243.write('Z5', 'SOS', body)
    worksheet243.write('AA5', 'FIS', body)
    worksheet243.write('AB5', 'KIM', body)
    worksheet243.write('AC5', 'BIO', body)
    worksheet243.write('AD5', 'JML', body)

    worksheet243.conditional_format(5, 0, row243_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet243.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_243}', title)
    worksheet243.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet243.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet243.write('A22', 'LOKASI', header)
    worksheet243.write('B22', 'TOTAL', header)
    worksheet243.merge_range('A21:B21', 'RANK', header)
    worksheet243.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet243.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet243.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet243.merge_range('F21:F22', 'KELAS', header)
    worksheet243.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet243.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet243.write('G22', 'MAW', body)
    worksheet243.write('H22', 'MAP', body)
    worksheet243.write('I22', 'IND', body)
    worksheet243.write('J22', 'ENG', body)
    worksheet243.write('K22', 'SEJ', body)
    worksheet243.write('L22', 'GEO', body)
    worksheet243.write('M22', 'EKO', body)
    worksheet243.write('N22', 'SOS', body)
    worksheet243.write('O22', 'FIS', body)
    worksheet243.write('P22', 'KIM', body)
    worksheet243.write('Q22', 'BIO', body)
    worksheet243.write('R22', 'JML', body)
    worksheet243.write('S22', 'MAW', body)
    worksheet243.write('T22', 'MAP', body)
    worksheet243.write('U22', 'IND', body)
    worksheet243.write('V22', 'ENG', body)
    worksheet243.write('W22', 'SEJ', body)
    worksheet243.write('X22', 'GEO', body)
    worksheet243.write('Y22', 'EKO', body)
    worksheet243.write('Z22', 'SOS', body)
    worksheet243.write('AA22', 'FIS', body)
    worksheet243.write('AB22', 'KIM', body)
    worksheet243.write('AC22', 'BIO', body)
    worksheet243.write('AD22', 'JML', body)

    worksheet243.conditional_format(22, 0, row243+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 244
    worksheet244.insert_image('A1', r'logo resmi nf.jpg')

    worksheet244.set_column('A:A', 7, center)
    worksheet244.set_column('B:B', 6, center)
    worksheet244.set_column('C:C', 18.14, center)
    worksheet244.set_column('D:D', 25, left)
    worksheet244.set_column('E:E', 13.14, left)
    worksheet244.set_column('F:F', 8.57, center)
    worksheet244.set_column('G:AD', 5, center)
    worksheet244.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_244}', title)
    worksheet244.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet244.write('A5', 'LOKASI', header)
    worksheet244.write('B5', 'TOTAL', header)
    worksheet244.merge_range('A4:B4', 'RANK', header)
    worksheet244.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet244.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet244.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet244.merge_range('F4:F5', 'KELAS', header)
    worksheet244.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet244.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet244.write('G5', 'MAW', body)
    worksheet244.write('H5', 'MAP', body)
    worksheet244.write('I5', 'IND', body)
    worksheet244.write('J5', 'ENG', body)
    worksheet244.write('K5', 'SEJ', body)
    worksheet244.write('L5', 'GEO', body)
    worksheet244.write('M5', 'EKO', body)
    worksheet244.write('N5', 'SOS', body)
    worksheet244.write('O5', 'FIS', body)
    worksheet244.write('P5', 'KIM', body)
    worksheet244.write('Q5', 'BIO', body)
    worksheet244.write('R5', 'JML', body)
    worksheet244.write('S5', 'MAW', body)
    worksheet244.write('T5', 'MAP', body)
    worksheet244.write('U5', 'IND', body)
    worksheet244.write('V5', 'ENG', body)
    worksheet244.write('W5', 'SEJ', body)
    worksheet244.write('X5', 'GEO', body)
    worksheet244.write('Y5', 'EKO', body)
    worksheet244.write('Z5', 'SOS', body)
    worksheet244.write('AA5', 'FIS', body)
    worksheet244.write('AB5', 'KIM', body)
    worksheet244.write('AC5', 'BIO', body)
    worksheet244.write('AD5', 'JML', body)

    worksheet244.conditional_format(5, 0, row244_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet244.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_244}', title)
    worksheet244.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet244.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet244.write('A22', 'LOKASI', header)
    worksheet244.write('B22', 'TOTAL', header)
    worksheet244.merge_range('A21:B21', 'RANK', header)
    worksheet244.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet244.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet244.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet244.merge_range('F21:F22', 'KELAS', header)
    worksheet244.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet244.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet244.write('G22', 'MAW', body)
    worksheet244.write('H22', 'MAP', body)
    worksheet244.write('I22', 'IND', body)
    worksheet244.write('J22', 'ENG', body)
    worksheet244.write('K22', 'SEJ', body)
    worksheet244.write('L22', 'GEO', body)
    worksheet244.write('M22', 'EKO', body)
    worksheet244.write('N22', 'SOS', body)
    worksheet244.write('O22', 'FIS', body)
    worksheet244.write('P22', 'KIM', body)
    worksheet244.write('Q22', 'BIO', body)
    worksheet244.write('R22', 'JML', body)
    worksheet244.write('S22', 'MAW', body)
    worksheet244.write('T22', 'MAP', body)
    worksheet244.write('U22', 'IND', body)
    worksheet244.write('V22', 'ENG', body)
    worksheet244.write('W22', 'SEJ', body)
    worksheet244.write('X22', 'GEO', body)
    worksheet244.write('Y22', 'EKO', body)
    worksheet244.write('Z22', 'SOS', body)
    worksheet244.write('AA22', 'FIS', body)
    worksheet244.write('AB22', 'KIM', body)
    worksheet244.write('AC22', 'BIO', body)
    worksheet244.write('AD22', 'JML', body)

    worksheet244.conditional_format(22, 0, row244+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 245
    worksheet245.insert_image('A1', r'logo resmi nf.jpg')

    worksheet245.set_column('A:A', 7, center)
    worksheet245.set_column('B:B', 6, center)
    worksheet245.set_column('C:C', 18.14, center)
    worksheet245.set_column('D:D', 25, left)
    worksheet245.set_column('E:E', 13.14, left)
    worksheet245.set_column('F:F', 8.57, center)
    worksheet245.set_column('G:AD', 5, center)
    worksheet245.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_245}', title)
    worksheet245.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet245.write('A5', 'LOKASI', header)
    worksheet245.write('B5', 'TOTAL', header)
    worksheet245.merge_range('A4:B4', 'RANK', header)
    worksheet245.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet245.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet245.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet245.merge_range('F4:F5', 'KELAS', header)
    worksheet245.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet245.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet245.write('G5', 'MAW', body)
    worksheet245.write('H5', 'MAP', body)
    worksheet245.write('I5', 'IND', body)
    worksheet245.write('J5', 'ENG', body)
    worksheet245.write('K5', 'SEJ', body)
    worksheet245.write('L5', 'GEO', body)
    worksheet245.write('M5', 'EKO', body)
    worksheet245.write('N5', 'SOS', body)
    worksheet245.write('O5', 'FIS', body)
    worksheet245.write('P5', 'KIM', body)
    worksheet245.write('Q5', 'BIO', body)
    worksheet245.write('R5', 'JML', body)
    worksheet245.write('S5', 'MAW', body)
    worksheet245.write('T5', 'MAP', body)
    worksheet245.write('U5', 'IND', body)
    worksheet245.write('V5', 'ENG', body)
    worksheet245.write('W5', 'SEJ', body)
    worksheet245.write('X5', 'GEO', body)
    worksheet245.write('Y5', 'EKO', body)
    worksheet245.write('Z5', 'SOS', body)
    worksheet245.write('AA5', 'FIS', body)
    worksheet245.write('AB5', 'KIM', body)
    worksheet245.write('AC5', 'BIO', body)
    worksheet245.write('AD5', 'JML', body)

    worksheet245.conditional_format(5, 0, row245_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet245.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_245}', title)
    worksheet245.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet245.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet245.write('A22', 'LOKASI', header)
    worksheet245.write('B22', 'TOTAL', header)
    worksheet245.merge_range('A21:B21', 'RANK', header)
    worksheet245.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet245.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet245.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet245.merge_range('F21:F22', 'KELAS', header)
    worksheet245.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet245.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet245.write('G22', 'MAW', body)
    worksheet245.write('H22', 'MAP', body)
    worksheet245.write('I22', 'IND', body)
    worksheet245.write('J22', 'ENG', body)
    worksheet245.write('K22', 'SEJ', body)
    worksheet245.write('L22', 'GEO', body)
    worksheet245.write('M22', 'EKO', body)
    worksheet245.write('N22', 'SOS', body)
    worksheet245.write('O22', 'FIS', body)
    worksheet245.write('P22', 'KIM', body)
    worksheet245.write('Q22', 'BIO', body)
    worksheet245.write('R22', 'JML', body)
    worksheet245.write('S22', 'MAW', body)
    worksheet245.write('T22', 'MAP', body)
    worksheet245.write('U22', 'IND', body)
    worksheet245.write('V22', 'ENG', body)
    worksheet245.write('W22', 'SEJ', body)
    worksheet245.write('X22', 'GEO', body)
    worksheet245.write('Y22', 'EKO', body)
    worksheet245.write('Z22', 'SOS', body)
    worksheet245.write('AA22', 'FIS', body)
    worksheet245.write('AB22', 'KIM', body)
    worksheet245.write('AC22', 'BIO', body)
    worksheet245.write('AD22', 'JML', body)

    worksheet245.conditional_format(22, 0, row245+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 246
    worksheet246.insert_image('A1', r'logo resmi nf.jpg')

    worksheet246.set_column('A:A', 7, center)
    worksheet246.set_column('B:B', 6, center)
    worksheet246.set_column('C:C', 18.14, center)
    worksheet246.set_column('D:D', 25, left)
    worksheet246.set_column('E:E', 13.14, left)
    worksheet246.set_column('F:F', 8.57, center)
    worksheet246.set_column('G:AD', 5, center)
    worksheet246.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_246}', title)
    worksheet246.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet246.write('A5', 'LOKASI', header)
    worksheet246.write('B5', 'TOTAL', header)
    worksheet246.merge_range('A4:B4', 'RANK', header)
    worksheet246.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet246.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet246.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet246.merge_range('F4:F5', 'KELAS', header)
    worksheet246.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet246.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet246.write('G5', 'MAW', body)
    worksheet246.write('H5', 'MAP', body)
    worksheet246.write('I5', 'IND', body)
    worksheet246.write('J5', 'ENG', body)
    worksheet246.write('K5', 'SEJ', body)
    worksheet246.write('L5', 'GEO', body)
    worksheet246.write('M5', 'EKO', body)
    worksheet246.write('N5', 'SOS', body)
    worksheet246.write('O5', 'FIS', body)
    worksheet246.write('P5', 'KIM', body)
    worksheet246.write('Q5', 'BIO', body)
    worksheet246.write('R5', 'JML', body)
    worksheet246.write('S5', 'MAW', body)
    worksheet246.write('T5', 'MAP', body)
    worksheet246.write('U5', 'IND', body)
    worksheet246.write('V5', 'ENG', body)
    worksheet246.write('W5', 'SEJ', body)
    worksheet246.write('X5', 'GEO', body)
    worksheet246.write('Y5', 'EKO', body)
    worksheet246.write('Z5', 'SOS', body)
    worksheet246.write('AA5', 'FIS', body)
    worksheet246.write('AB5', 'KIM', body)
    worksheet246.write('AC5', 'BIO', body)
    worksheet246.write('AD5', 'JML', body)

    worksheet246.conditional_format(5, 0, row246_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet246.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_246}', title)
    worksheet246.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet246.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet246.write('A22', 'LOKASI', header)
    worksheet246.write('B22', 'TOTAL', header)
    worksheet246.merge_range('A21:B21', 'RANK', header)
    worksheet246.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet246.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet246.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet246.merge_range('F21:F22', 'KELAS', header)
    worksheet246.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet246.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet246.write('G22', 'MAW', body)
    worksheet246.write('H22', 'MAP', body)
    worksheet246.write('I22', 'IND', body)
    worksheet246.write('J22', 'ENG', body)
    worksheet246.write('K22', 'SEJ', body)
    worksheet246.write('L22', 'GEO', body)
    worksheet246.write('M22', 'EKO', body)
    worksheet246.write('N22', 'SOS', body)
    worksheet246.write('O22', 'FIS', body)
    worksheet246.write('P22', 'KIM', body)
    worksheet246.write('Q22', 'BIO', body)
    worksheet246.write('R22', 'JML', body)
    worksheet246.write('S22', 'MAW', body)
    worksheet246.write('T22', 'MAP', body)
    worksheet246.write('U22', 'IND', body)
    worksheet246.write('V22', 'ENG', body)
    worksheet246.write('W22', 'SEJ', body)
    worksheet246.write('X22', 'GEO', body)
    worksheet246.write('Y22', 'EKO', body)
    worksheet246.write('Z22', 'SOS', body)
    worksheet246.write('AA22', 'FIS', body)
    worksheet246.write('AB22', 'KIM', body)
    worksheet246.write('AC22', 'BIO', body)
    worksheet246.write('AD22', 'JML', body)

    worksheet246.conditional_format(22, 0, row246+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 248
    worksheet248.insert_image('A1', r'logo resmi nf.jpg')

    worksheet248.set_column('A:A', 7, center)
    worksheet248.set_column('B:B', 6, center)
    worksheet248.set_column('C:C', 18.14, center)
    worksheet248.set_column('D:D', 25, left)
    worksheet248.set_column('E:E', 13.14, left)
    worksheet248.set_column('F:F', 8.57, center)
    worksheet248.set_column('G:AD', 5, center)
    worksheet248.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_248}', title)
    worksheet248.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet248.write('A5', 'LOKASI', header)
    worksheet248.write('B5', 'TOTAL', header)
    worksheet248.merge_range('A4:B4', 'RANK', header)
    worksheet248.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet248.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet248.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet248.merge_range('F4:F5', 'KELAS', header)
    worksheet248.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet248.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet248.write('G5', 'MAW', body)
    worksheet248.write('H5', 'MAP', body)
    worksheet248.write('I5', 'IND', body)
    worksheet248.write('J5', 'ENG', body)
    worksheet248.write('K5', 'SEJ', body)
    worksheet248.write('L5', 'GEO', body)
    worksheet248.write('M5', 'EKO', body)
    worksheet248.write('N5', 'SOS', body)
    worksheet248.write('O5', 'FIS', body)
    worksheet248.write('P5', 'KIM', body)
    worksheet248.write('Q5', 'BIO', body)
    worksheet248.write('R5', 'JML', body)
    worksheet248.write('S5', 'MAW', body)
    worksheet248.write('T5', 'MAP', body)
    worksheet248.write('U5', 'IND', body)
    worksheet248.write('V5', 'ENG', body)
    worksheet248.write('W5', 'SEJ', body)
    worksheet248.write('X5', 'GEO', body)
    worksheet248.write('Y5', 'EKO', body)
    worksheet248.write('Z5', 'SOS', body)
    worksheet248.write('AA5', 'FIS', body)
    worksheet248.write('AB5', 'KIM', body)
    worksheet248.write('AC5', 'BIO', body)
    worksheet248.write('AD5', 'JML', body)

    worksheet248.conditional_format(5, 0, row248_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet248.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_248}', title)
    worksheet248.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet248.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet248.write('A22', 'LOKASI', header)
    worksheet248.write('B22', 'TOTAL', header)
    worksheet248.merge_range('A21:B21', 'RANK', header)
    worksheet248.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet248.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet248.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet248.merge_range('F21:F22', 'KELAS', header)
    worksheet248.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet248.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet248.write('G22', 'MAW', body)
    worksheet248.write('H22', 'MAP', body)
    worksheet248.write('I22', 'IND', body)
    worksheet248.write('J22', 'ENG', body)
    worksheet248.write('K22', 'SEJ', body)
    worksheet248.write('L22', 'GEO', body)
    worksheet248.write('M22', 'EKO', body)
    worksheet248.write('N22', 'SOS', body)
    worksheet248.write('O22', 'FIS', body)
    worksheet248.write('P22', 'KIM', body)
    worksheet248.write('Q22', 'BIO', body)
    worksheet248.write('R22', 'JML', body)
    worksheet248.write('S22', 'MAW', body)
    worksheet248.write('T22', 'MAP', body)
    worksheet248.write('U22', 'IND', body)
    worksheet248.write('V22', 'ENG', body)
    worksheet248.write('W22', 'SEJ', body)
    worksheet248.write('X22', 'GEO', body)
    worksheet248.write('Y22', 'EKO', body)
    worksheet248.write('Z22', 'SOS', body)
    worksheet248.write('AA22', 'FIS', body)
    worksheet248.write('AB22', 'KIM', body)
    worksheet248.write('AC22', 'BIO', body)
    worksheet248.write('AD22', 'JML', body)

    worksheet248.conditional_format(22, 0, row248+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 249
    worksheet249.insert_image('A1', r'logo resmi nf.jpg')

    worksheet249.set_column('A:A', 7, center)
    worksheet249.set_column('B:B', 6, center)
    worksheet249.set_column('C:C', 18.14, center)
    worksheet249.set_column('D:D', 25, left)
    worksheet249.set_column('E:E', 13.14, left)
    worksheet249.set_column('F:F', 8.57, center)
    worksheet249.set_column('G:AD', 5, center)
    worksheet249.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_249}', title)
    worksheet249.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet249.write('A5', 'LOKASI', header)
    worksheet249.write('B5', 'TOTAL', header)
    worksheet249.merge_range('A4:B4', 'RANK', header)
    worksheet249.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet249.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet249.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet249.merge_range('F4:F5', 'KELAS', header)
    worksheet249.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet249.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet249.write('G5', 'MAW', body)
    worksheet249.write('H5', 'MAP', body)
    worksheet249.write('I5', 'IND', body)
    worksheet249.write('J5', 'ENG', body)
    worksheet249.write('K5', 'SEJ', body)
    worksheet249.write('L5', 'GEO', body)
    worksheet249.write('M5', 'EKO', body)
    worksheet249.write('N5', 'SOS', body)
    worksheet249.write('O5', 'FIS', body)
    worksheet249.write('P5', 'KIM', body)
    worksheet249.write('Q5', 'BIO', body)
    worksheet249.write('R5', 'JML', body)
    worksheet249.write('S5', 'MAW', body)
    worksheet249.write('T5', 'MAP', body)
    worksheet249.write('U5', 'IND', body)
    worksheet249.write('V5', 'ENG', body)
    worksheet249.write('W5', 'SEJ', body)
    worksheet249.write('X5', 'GEO', body)
    worksheet249.write('Y5', 'EKO', body)
    worksheet249.write('Z5', 'SOS', body)
    worksheet249.write('AA5', 'FIS', body)
    worksheet249.write('AB5', 'KIM', body)
    worksheet249.write('AC5', 'BIO', body)
    worksheet249.write('AD5', 'JML', body)

    worksheet249.conditional_format(5, 0, row249_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet249.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_249}', title)
    worksheet249.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet249.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet249.write('A22', 'LOKASI', header)
    worksheet249.write('B22', 'TOTAL', header)
    worksheet249.merge_range('A21:B21', 'RANK', header)
    worksheet249.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet249.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet249.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet249.merge_range('F21:F22', 'KELAS', header)
    worksheet249.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet249.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet249.write('G22', 'MAW', body)
    worksheet249.write('H22', 'MAP', body)
    worksheet249.write('I22', 'IND', body)
    worksheet249.write('J22', 'ENG', body)
    worksheet249.write('K22', 'SEJ', body)
    worksheet249.write('L22', 'GEO', body)
    worksheet249.write('M22', 'EKO', body)
    worksheet249.write('N22', 'SOS', body)
    worksheet249.write('O22', 'FIS', body)
    worksheet249.write('P22', 'KIM', body)
    worksheet249.write('Q22', 'BIO', body)
    worksheet249.write('R22', 'JML', body)
    worksheet249.write('S22', 'MAW', body)
    worksheet249.write('T22', 'MAP', body)
    worksheet249.write('U22', 'IND', body)
    worksheet249.write('V22', 'ENG', body)
    worksheet249.write('W22', 'SEJ', body)
    worksheet249.write('X22', 'GEO', body)
    worksheet249.write('Y22', 'EKO', body)
    worksheet249.write('Z22', 'SOS', body)
    worksheet249.write('AA22', 'FIS', body)
    worksheet249.write('AB22', 'KIM', body)
    worksheet249.write('AC22', 'BIO', body)
    worksheet249.write('AD22', 'JML', body)

    worksheet249.conditional_format(22, 0, row249+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 250
    worksheet250.insert_image('A1', r'logo resmi nf.jpg')

    worksheet250.set_column('A:A', 7, center)
    worksheet250.set_column('B:B', 6, center)
    worksheet250.set_column('C:C', 18.14, center)
    worksheet250.set_column('D:D', 25, left)
    worksheet250.set_column('E:E', 13.14, left)
    worksheet250.set_column('F:F', 8.57, center)
    worksheet250.set_column('G:AD', 5, center)
    worksheet250.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_250}', title)
    worksheet250.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet250.write('A5', 'LOKASI', header)
    worksheet250.write('B5', 'TOTAL', header)
    worksheet250.merge_range('A4:B4', 'RANK', header)
    worksheet250.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet250.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet250.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet250.merge_range('F4:F5', 'KELAS', header)
    worksheet250.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet250.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet250.write('G5', 'MAW', body)
    worksheet250.write('H5', 'MAP', body)
    worksheet250.write('I5', 'IND', body)
    worksheet250.write('J5', 'ENG', body)
    worksheet250.write('K5', 'SEJ', body)
    worksheet250.write('L5', 'GEO', body)
    worksheet250.write('M5', 'EKO', body)
    worksheet250.write('N5', 'SOS', body)
    worksheet250.write('O5', 'FIS', body)
    worksheet250.write('P5', 'KIM', body)
    worksheet250.write('Q5', 'BIO', body)
    worksheet250.write('R5', 'JML', body)
    worksheet250.write('S5', 'MAW', body)
    worksheet250.write('T5', 'MAP', body)
    worksheet250.write('U5', 'IND', body)
    worksheet250.write('V5', 'ENG', body)
    worksheet250.write('W5', 'SEJ', body)
    worksheet250.write('X5', 'GEO', body)
    worksheet250.write('Y5', 'EKO', body)
    worksheet250.write('Z5', 'SOS', body)
    worksheet250.write('AA5', 'FIS', body)
    worksheet250.write('AB5', 'KIM', body)
    worksheet250.write('AC5', 'BIO', body)
    worksheet250.write('AD5', 'JML', body)

    worksheet250.conditional_format(5, 0, row250_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet250.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_250}', title)
    worksheet250.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet250.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet250.write('A22', 'LOKASI', header)
    worksheet250.write('B22', 'TOTAL', header)
    worksheet250.merge_range('A21:B21', 'RANK', header)
    worksheet250.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet250.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet250.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet250.merge_range('F21:F22', 'KELAS', header)
    worksheet250.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet250.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet250.write('G22', 'MAW', body)
    worksheet250.write('H22', 'MAP', body)
    worksheet250.write('I22', 'IND', body)
    worksheet250.write('J22', 'ENG', body)
    worksheet250.write('K22', 'SEJ', body)
    worksheet250.write('L22', 'GEO', body)
    worksheet250.write('M22', 'EKO', body)
    worksheet250.write('N22', 'SOS', body)
    worksheet250.write('O22', 'FIS', body)
    worksheet250.write('P22', 'KIM', body)
    worksheet250.write('Q22', 'BIO', body)
    worksheet250.write('R22', 'JML', body)
    worksheet250.write('S22', 'MAW', body)
    worksheet250.write('T22', 'MAP', body)
    worksheet250.write('U22', 'IND', body)
    worksheet250.write('V22', 'ENG', body)
    worksheet250.write('W22', 'SEJ', body)
    worksheet250.write('X22', 'GEO', body)
    worksheet250.write('Y22', 'EKO', body)
    worksheet250.write('Z22', 'SOS', body)
    worksheet250.write('AA22', 'FIS', body)
    worksheet250.write('AB22', 'KIM', body)
    worksheet250.write('AC22', 'BIO', body)
    worksheet250.write('AD22', 'JML', body)

    worksheet250.conditional_format(22, 0, row250+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 252
    worksheet252.insert_image('A1', r'logo resmi nf.jpg')

    worksheet252.set_column('A:A', 7, center)
    worksheet252.set_column('B:B', 6, center)
    worksheet252.set_column('C:C', 18.14, center)
    worksheet252.set_column('D:D', 25, left)
    worksheet252.set_column('E:E', 13.14, left)
    worksheet252.set_column('F:F', 8.57, center)
    worksheet252.set_column('G:AD', 5, center)
    worksheet252.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_252}', title)
    worksheet252.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet252.write('A5', 'LOKASI', header)
    worksheet252.write('B5', 'TOTAL', header)
    worksheet252.merge_range('A4:B4', 'RANK', header)
    worksheet252.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet252.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet252.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet252.merge_range('F4:F5', 'KELAS', header)
    worksheet252.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet252.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet252.write('G5', 'MAW', body)
    worksheet252.write('H5', 'MAP', body)
    worksheet252.write('I5', 'IND', body)
    worksheet252.write('J5', 'ENG', body)
    worksheet252.write('K5', 'SEJ', body)
    worksheet252.write('L5', 'GEO', body)
    worksheet252.write('M5', 'EKO', body)
    worksheet252.write('N5', 'SOS', body)
    worksheet252.write('O5', 'FIS', body)
    worksheet252.write('P5', 'KIM', body)
    worksheet252.write('Q5', 'BIO', body)
    worksheet252.write('R5', 'JML', body)
    worksheet252.write('S5', 'MAW', body)
    worksheet252.write('T5', 'MAP', body)
    worksheet252.write('U5', 'IND', body)
    worksheet252.write('V5', 'ENG', body)
    worksheet252.write('W5', 'SEJ', body)
    worksheet252.write('X5', 'GEO', body)
    worksheet252.write('Y5', 'EKO', body)
    worksheet252.write('Z5', 'SOS', body)
    worksheet252.write('AA5', 'FIS', body)
    worksheet252.write('AB5', 'KIM', body)
    worksheet252.write('AC5', 'BIO', body)
    worksheet252.write('AD5', 'JML', body)

    worksheet252.conditional_format(5, 0, row252_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet252.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_252}', title)
    worksheet252.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet252.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet252.write('A22', 'LOKASI', header)
    worksheet252.write('B22', 'TOTAL', header)
    worksheet252.merge_range('A21:B21', 'RANK', header)
    worksheet252.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet252.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet252.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet252.merge_range('F21:F22', 'KELAS', header)
    worksheet252.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet252.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet252.write('G22', 'MAW', body)
    worksheet252.write('H22', 'MAP', body)
    worksheet252.write('I22', 'IND', body)
    worksheet252.write('J22', 'ENG', body)
    worksheet252.write('K22', 'SEJ', body)
    worksheet252.write('L22', 'GEO', body)
    worksheet252.write('M22', 'EKO', body)
    worksheet252.write('N22', 'SOS', body)
    worksheet252.write('O22', 'FIS', body)
    worksheet252.write('P22', 'KIM', body)
    worksheet252.write('Q22', 'BIO', body)
    worksheet252.write('R22', 'JML', body)
    worksheet252.write('S22', 'MAW', body)
    worksheet252.write('T22', 'MAP', body)
    worksheet252.write('U22', 'IND', body)
    worksheet252.write('V22', 'ENG', body)
    worksheet252.write('W22', 'SEJ', body)
    worksheet252.write('X22', 'GEO', body)
    worksheet252.write('Y22', 'EKO', body)
    worksheet252.write('Z22', 'SOS', body)
    worksheet252.write('AA22', 'FIS', body)
    worksheet252.write('AB22', 'KIM', body)
    worksheet252.write('AC22', 'BIO', body)
    worksheet252.write('AD22', 'JML', body)

    worksheet252.conditional_format(22, 0, row252+21, 29,
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
    # worksheet 254
    worksheet254.insert_image('A1', r'logo resmi nf.jpg')

    worksheet254.set_column('A:A', 7, center)
    worksheet254.set_column('B:B', 6, center)
    worksheet254.set_column('C:C', 18.14, center)
    worksheet254.set_column('D:D', 25, left)
    worksheet254.set_column('E:E', 13.14, left)
    worksheet254.set_column('F:F', 8.57, center)
    worksheet254.set_column('G:AD', 5, center)
    worksheet254.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_254}', title)
    worksheet254.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet254.write('A5', 'LOKASI', header)
    worksheet254.write('B5', 'TOTAL', header)
    worksheet254.merge_range('A4:B4', 'RANK', header)
    worksheet254.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet254.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet254.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet254.merge_range('F4:F5', 'KELAS', header)
    worksheet254.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet254.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet254.write('G5', 'MAW', body)
    worksheet254.write('H5', 'MAP', body)
    worksheet254.write('I5', 'IND', body)
    worksheet254.write('J5', 'ENG', body)
    worksheet254.write('K5', 'SEJ', body)
    worksheet254.write('L5', 'GEO', body)
    worksheet254.write('M5', 'EKO', body)
    worksheet254.write('N5', 'SOS', body)
    worksheet254.write('O5', 'FIS', body)
    worksheet254.write('P5', 'KIM', body)
    worksheet254.write('Q5', 'BIO', body)
    worksheet254.write('R5', 'JML', body)
    worksheet254.write('S5', 'MAW', body)
    worksheet254.write('T5', 'MAP', body)
    worksheet254.write('U5', 'IND', body)
    worksheet254.write('V5', 'ENG', body)
    worksheet254.write('W5', 'SEJ', body)
    worksheet254.write('X5', 'GEO', body)
    worksheet254.write('Y5', 'EKO', body)
    worksheet254.write('Z5', 'SOS', body)
    worksheet254.write('AA5', 'FIS', body)
    worksheet254.write('AB5', 'KIM', body)
    worksheet254.write('AC5', 'BIO', body)
    worksheet254.write('AD5', 'JML', body)

    worksheet254.conditional_format(5, 0, row254_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet254.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_254}', title)
    worksheet254.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet254.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet254.write('A22', 'LOKASI', header)
    worksheet254.write('B22', 'TOTAL', header)
    worksheet254.merge_range('A21:B21', 'RANK', header)
    worksheet254.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet254.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet254.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet254.merge_range('F21:F22', 'KELAS', header)
    worksheet254.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet254.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet254.write('G22', 'MAW', body)
    worksheet254.write('H22', 'MAP', body)
    worksheet254.write('I22', 'IND', body)
    worksheet254.write('J22', 'ENG', body)
    worksheet254.write('K22', 'SEJ', body)
    worksheet254.write('L22', 'GEO', body)
    worksheet254.write('M22', 'EKO', body)
    worksheet254.write('N22', 'SOS', body)
    worksheet254.write('O22', 'FIS', body)
    worksheet254.write('P22', 'KIM', body)
    worksheet254.write('Q22', 'BIO', body)
    worksheet254.write('R22', 'JML', body)
    worksheet254.write('S22', 'MAW', body)
    worksheet254.write('T22', 'MAP', body)
    worksheet254.write('U22', 'IND', body)
    worksheet254.write('V22', 'ENG', body)
    worksheet254.write('W22', 'SEJ', body)
    worksheet254.write('X22', 'GEO', body)
    worksheet254.write('Y22', 'EKO', body)
    worksheet254.write('Z22', 'SOS', body)
    worksheet254.write('AA22', 'FIS', body)
    worksheet254.write('AB22', 'KIM', body)
    worksheet254.write('AC22', 'BIO', body)
    worksheet254.write('AD22', 'JML', body)

    worksheet254.conditional_format(22, 0, row254+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 255
    worksheet255.insert_image('A1', r'logo resmi nf.jpg')

    worksheet255.set_column('A:A', 7, center)
    worksheet255.set_column('B:B', 6, center)
    worksheet255.set_column('C:C', 18.14, center)
    worksheet255.set_column('D:D', 25, left)
    worksheet255.set_column('E:E', 13.14, left)
    worksheet255.set_column('F:F', 8.57, center)
    worksheet255.set_column('G:AD', 5, center)
    worksheet255.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_255}', title)
    worksheet255.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet255.write('A5', 'LOKASI', header)
    worksheet255.write('B5', 'TOTAL', header)
    worksheet255.merge_range('A4:B4', 'RANK', header)
    worksheet255.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet255.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet255.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet255.merge_range('F4:F5', 'KELAS', header)
    worksheet255.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet255.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet255.write('G5', 'MAW', body)
    worksheet255.write('H5', 'MAP', body)
    worksheet255.write('I5', 'IND', body)
    worksheet255.write('J5', 'ENG', body)
    worksheet255.write('K5', 'SEJ', body)
    worksheet255.write('L5', 'GEO', body)
    worksheet255.write('M5', 'EKO', body)
    worksheet255.write('N5', 'SOS', body)
    worksheet255.write('O5', 'FIS', body)
    worksheet255.write('P5', 'KIM', body)
    worksheet255.write('Q5', 'BIO', body)
    worksheet255.write('R5', 'JML', body)
    worksheet255.write('S5', 'MAW', body)
    worksheet255.write('T5', 'MAP', body)
    worksheet255.write('U5', 'IND', body)
    worksheet255.write('V5', 'ENG', body)
    worksheet255.write('W5', 'SEJ', body)
    worksheet255.write('X5', 'GEO', body)
    worksheet255.write('Y5', 'EKO', body)
    worksheet255.write('Z5', 'SOS', body)
    worksheet255.write('AA5', 'FIS', body)
    worksheet255.write('AB5', 'KIM', body)
    worksheet255.write('AC5', 'BIO', body)
    worksheet255.write('AD5', 'JML', body)

    worksheet255.conditional_format(5, 0, row255_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet255.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_255}', title)
    worksheet255.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet255.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet255.write('A22', 'LOKASI', header)
    worksheet255.write('B22', 'TOTAL', header)
    worksheet255.merge_range('A21:B21', 'RANK', header)
    worksheet255.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet255.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet255.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet255.merge_range('F21:F22', 'KELAS', header)
    worksheet255.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet255.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet255.write('G22', 'MAW', body)
    worksheet255.write('H22', 'MAP', body)
    worksheet255.write('I22', 'IND', body)
    worksheet255.write('J22', 'ENG', body)
    worksheet255.write('K22', 'SEJ', body)
    worksheet255.write('L22', 'GEO', body)
    worksheet255.write('M22', 'EKO', body)
    worksheet255.write('N22', 'SOS', body)
    worksheet255.write('O22', 'FIS', body)
    worksheet255.write('P22', 'KIM', body)
    worksheet255.write('Q22', 'BIO', body)
    worksheet255.write('R22', 'JML', body)
    worksheet255.write('S22', 'MAW', body)
    worksheet255.write('T22', 'MAP', body)
    worksheet255.write('U22', 'IND', body)
    worksheet255.write('V22', 'ENG', body)
    worksheet255.write('W22', 'SEJ', body)
    worksheet255.write('X22', 'GEO', body)
    worksheet255.write('Y22', 'EKO', body)
    worksheet255.write('Z22', 'SOS', body)
    worksheet255.write('AA22', 'FIS', body)
    worksheet255.write('AB22', 'KIM', body)
    worksheet255.write('AC22', 'BIO', body)
    worksheet255.write('AD22', 'JML', body)

    worksheet255.conditional_format(22, 0, row255+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 256
    worksheet256.insert_image('A1', r'logo resmi nf.jpg')

    worksheet256.set_column('A:A', 7, center)
    worksheet256.set_column('B:B', 6, center)
    worksheet256.set_column('C:C', 18.14, center)
    worksheet256.set_column('D:D', 25, left)
    worksheet256.set_column('E:E', 13.14, left)
    worksheet256.set_column('F:F', 8.57, center)
    worksheet256.set_column('G:AD', 5, center)
    worksheet256.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_256}', title)
    worksheet256.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet256.write('A5', 'LOKASI', header)
    worksheet256.write('B5', 'TOTAL', header)
    worksheet256.merge_range('A4:B4', 'RANK', header)
    worksheet256.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet256.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet256.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet256.merge_range('F4:F5', 'KELAS', header)
    worksheet256.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet256.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet256.write('G5', 'MAW', body)
    worksheet256.write('H5', 'MAP', body)
    worksheet256.write('I5', 'IND', body)
    worksheet256.write('J5', 'ENG', body)
    worksheet256.write('K5', 'SEJ', body)
    worksheet256.write('L5', 'GEO', body)
    worksheet256.write('M5', 'EKO', body)
    worksheet256.write('N5', 'SOS', body)
    worksheet256.write('O5', 'FIS', body)
    worksheet256.write('P5', 'KIM', body)
    worksheet256.write('Q5', 'BIO', body)
    worksheet256.write('R5', 'JML', body)
    worksheet256.write('S5', 'MAW', body)
    worksheet256.write('T5', 'MAP', body)
    worksheet256.write('U5', 'IND', body)
    worksheet256.write('V5', 'ENG', body)
    worksheet256.write('W5', 'SEJ', body)
    worksheet256.write('X5', 'GEO', body)
    worksheet256.write('Y5', 'EKO', body)
    worksheet256.write('Z5', 'SOS', body)
    worksheet256.write('AA5', 'FIS', body)
    worksheet256.write('AB5', 'KIM', body)
    worksheet256.write('AC5', 'BIO', body)
    worksheet256.write('AD5', 'JML', body)

    worksheet256.conditional_format(5, 0, row256_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet256.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_256}', title)
    worksheet256.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet256.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet256.write('A22', 'LOKASI', header)
    worksheet256.write('B22', 'TOTAL', header)
    worksheet256.merge_range('A21:B21', 'RANK', header)
    worksheet256.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet256.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet256.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet256.merge_range('F21:F22', 'KELAS', header)
    worksheet256.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet256.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet256.write('G22', 'MAW', body)
    worksheet256.write('H22', 'MAP', body)
    worksheet256.write('I22', 'IND', body)
    worksheet256.write('J22', 'ENG', body)
    worksheet256.write('K22', 'SEJ', body)
    worksheet256.write('L22', 'GEO', body)
    worksheet256.write('M22', 'EKO', body)
    worksheet256.write('N22', 'SOS', body)
    worksheet256.write('O22', 'FIS', body)
    worksheet256.write('P22', 'KIM', body)
    worksheet256.write('Q22', 'BIO', body)
    worksheet256.write('R22', 'JML', body)
    worksheet256.write('S22', 'MAW', body)
    worksheet256.write('T22', 'MAP', body)
    worksheet256.write('U22', 'IND', body)
    worksheet256.write('V22', 'ENG', body)
    worksheet256.write('W22', 'SEJ', body)
    worksheet256.write('X22', 'GEO', body)
    worksheet256.write('Y22', 'EKO', body)
    worksheet256.write('Z22', 'SOS', body)
    worksheet256.write('AA22', 'FIS', body)
    worksheet256.write('AB22', 'KIM', body)
    worksheet256.write('AC22', 'BIO', body)
    worksheet256.write('AD22', 'JML', body)

    worksheet256.conditional_format(22, 0, row256+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 258
    worksheet258.insert_image('A1', r'logo resmi nf.jpg')

    worksheet258.set_column('A:A', 7, center)
    worksheet258.set_column('B:B', 6, center)
    worksheet258.set_column('C:C', 18.14, center)
    worksheet258.set_column('D:D', 25, left)
    worksheet258.set_column('E:E', 13.14, left)
    worksheet258.set_column('F:F', 8.57, center)
    worksheet258.set_column('G:AD', 5, center)
    worksheet258.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_258}', title)
    worksheet258.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet258.write('A5', 'LOKASI', header)
    worksheet258.write('B5', 'TOTAL', header)
    worksheet258.merge_range('A4:B4', 'RANK', header)
    worksheet258.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet258.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet258.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet258.merge_range('F4:F5', 'KELAS', header)
    worksheet258.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet258.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet258.write('G5', 'MAW', body)
    worksheet258.write('H5', 'MAP', body)
    worksheet258.write('I5', 'IND', body)
    worksheet258.write('J5', 'ENG', body)
    worksheet258.write('K5', 'SEJ', body)
    worksheet258.write('L5', 'GEO', body)
    worksheet258.write('M5', 'EKO', body)
    worksheet258.write('N5', 'SOS', body)
    worksheet258.write('O5', 'FIS', body)
    worksheet258.write('P5', 'KIM', body)
    worksheet258.write('Q5', 'BIO', body)
    worksheet258.write('R5', 'JML', body)
    worksheet258.write('S5', 'MAW', body)
    worksheet258.write('T5', 'MAP', body)
    worksheet258.write('U5', 'IND', body)
    worksheet258.write('V5', 'ENG', body)
    worksheet258.write('W5', 'SEJ', body)
    worksheet258.write('X5', 'GEO', body)
    worksheet258.write('Y5', 'EKO', body)
    worksheet258.write('Z5', 'SOS', body)
    worksheet258.write('AA5', 'FIS', body)
    worksheet258.write('AB5', 'KIM', body)
    worksheet258.write('AC5', 'BIO', body)
    worksheet258.write('AD5', 'JML', body)

    worksheet258.conditional_format(5, 0, row258_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet258.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_258}', title)
    worksheet258.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet258.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet258.write('A22', 'LOKASI', header)
    worksheet258.write('B22', 'TOTAL', header)
    worksheet258.merge_range('A21:B21', 'RANK', header)
    worksheet258.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet258.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet258.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet258.merge_range('F21:F22', 'KELAS', header)
    worksheet258.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet258.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet258.write('G22', 'MAW', body)
    worksheet258.write('H22', 'MAP', body)
    worksheet258.write('I22', 'IND', body)
    worksheet258.write('J22', 'ENG', body)
    worksheet258.write('K22', 'SEJ', body)
    worksheet258.write('L22', 'GEO', body)
    worksheet258.write('M22', 'EKO', body)
    worksheet258.write('N22', 'SOS', body)
    worksheet258.write('O22', 'FIS', body)
    worksheet258.write('P22', 'KIM', body)
    worksheet258.write('Q22', 'BIO', body)
    worksheet258.write('R22', 'JML', body)
    worksheet258.write('S22', 'MAW', body)
    worksheet258.write('T22', 'MAP', body)
    worksheet258.write('U22', 'IND', body)
    worksheet258.write('V22', 'ENG', body)
    worksheet258.write('W22', 'SEJ', body)
    worksheet258.write('X22', 'GEO', body)
    worksheet258.write('Y22', 'EKO', body)
    worksheet258.write('Z22', 'SOS', body)
    worksheet258.write('AA22', 'FIS', body)
    worksheet258.write('AB22', 'KIM', body)
    worksheet258.write('AC22', 'BIO', body)
    worksheet258.write('AD22', 'JML', body)

    worksheet258.conditional_format(22, 0, row258+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 259
    worksheet259.insert_image('A1', r'logo resmi nf.jpg')

    worksheet259.set_column('A:A', 7, center)
    worksheet259.set_column('B:B', 6, center)
    worksheet259.set_column('C:C', 18.14, center)
    worksheet259.set_column('D:D', 25, left)
    worksheet259.set_column('E:E', 13.14, left)
    worksheet259.set_column('F:F', 8.57, center)
    worksheet259.set_column('G:AD', 5, center)
    worksheet259.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_259}', title)
    worksheet259.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet259.write('A5', 'LOKASI', header)
    worksheet259.write('B5', 'TOTAL', header)
    worksheet259.merge_range('A4:B4', 'RANK', header)
    worksheet259.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet259.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet259.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet259.merge_range('F4:F5', 'KELAS', header)
    worksheet259.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet259.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet259.write('G5', 'MAW', body)
    worksheet259.write('H5', 'MAP', body)
    worksheet259.write('I5', 'IND', body)
    worksheet259.write('J5', 'ENG', body)
    worksheet259.write('K5', 'SEJ', body)
    worksheet259.write('L5', 'GEO', body)
    worksheet259.write('M5', 'EKO', body)
    worksheet259.write('N5', 'SOS', body)
    worksheet259.write('O5', 'FIS', body)
    worksheet259.write('P5', 'KIM', body)
    worksheet259.write('Q5', 'BIO', body)
    worksheet259.write('R5', 'JML', body)
    worksheet259.write('S5', 'MAW', body)
    worksheet259.write('T5', 'MAP', body)
    worksheet259.write('U5', 'IND', body)
    worksheet259.write('V5', 'ENG', body)
    worksheet259.write('W5', 'SEJ', body)
    worksheet259.write('X5', 'GEO', body)
    worksheet259.write('Y5', 'EKO', body)
    worksheet259.write('Z5', 'SOS', body)
    worksheet259.write('AA5', 'FIS', body)
    worksheet259.write('AB5', 'KIM', body)
    worksheet259.write('AC5', 'BIO', body)
    worksheet259.write('AD5', 'JML', body)

    worksheet259.conditional_format(5, 0, row259_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet259.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_259}', title)
    worksheet259.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet259.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet259.write('A22', 'LOKASI', header)
    worksheet259.write('B22', 'TOTAL', header)
    worksheet259.merge_range('A21:B21', 'RANK', header)
    worksheet259.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet259.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet259.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet259.merge_range('F21:F22', 'KELAS', header)
    worksheet259.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet259.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet259.write('G22', 'MAW', body)
    worksheet259.write('H22', 'MAP', body)
    worksheet259.write('I22', 'IND', body)
    worksheet259.write('J22', 'ENG', body)
    worksheet259.write('K22', 'SEJ', body)
    worksheet259.write('L22', 'GEO', body)
    worksheet259.write('M22', 'EKO', body)
    worksheet259.write('N22', 'SOS', body)
    worksheet259.write('O22', 'FIS', body)
    worksheet259.write('P22', 'KIM', body)
    worksheet259.write('Q22', 'BIO', body)
    worksheet259.write('R22', 'JML', body)
    worksheet259.write('S22', 'MAW', body)
    worksheet259.write('T22', 'MAP', body)
    worksheet259.write('U22', 'IND', body)
    worksheet259.write('V22', 'ENG', body)
    worksheet259.write('W22', 'SEJ', body)
    worksheet259.write('X22', 'GEO', body)
    worksheet259.write('Y22', 'EKO', body)
    worksheet259.write('Z22', 'SOS', body)
    worksheet259.write('AA22', 'FIS', body)
    worksheet259.write('AB22', 'KIM', body)
    worksheet259.write('AC22', 'BIO', body)
    worksheet259.write('AD22', 'JML', body)

    worksheet259.conditional_format(22, 0, row259+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 260
    worksheet260.insert_image('A1', r'logo resmi nf.jpg')

    worksheet260.set_column('A:A', 7, center)
    worksheet260.set_column('B:B', 6, center)
    worksheet260.set_column('C:C', 18.14, center)
    worksheet260.set_column('D:D', 25, left)
    worksheet260.set_column('E:E', 13.14, left)
    worksheet260.set_column('F:F', 8.57, center)
    worksheet260.set_column('G:AD', 5, center)
    worksheet260.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_260}', title)
    worksheet260.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet260.write('A5', 'LOKASI', header)
    worksheet260.write('B5', 'TOTAL', header)
    worksheet260.merge_range('A4:B4', 'RANK', header)
    worksheet260.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet260.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet260.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet260.merge_range('F4:F5', 'KELAS', header)
    worksheet260.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet260.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet260.write('G5', 'MAW', body)
    worksheet260.write('H5', 'MAP', body)
    worksheet260.write('I5', 'IND', body)
    worksheet260.write('J5', 'ENG', body)
    worksheet260.write('K5', 'SEJ', body)
    worksheet260.write('L5', 'GEO', body)
    worksheet260.write('M5', 'EKO', body)
    worksheet260.write('N5', 'SOS', body)
    worksheet260.write('O5', 'FIS', body)
    worksheet260.write('P5', 'KIM', body)
    worksheet260.write('Q5', 'BIO', body)
    worksheet260.write('R5', 'JML', body)
    worksheet260.write('S5', 'MAW', body)
    worksheet260.write('T5', 'MAP', body)
    worksheet260.write('U5', 'IND', body)
    worksheet260.write('V5', 'ENG', body)
    worksheet260.write('W5', 'SEJ', body)
    worksheet260.write('X5', 'GEO', body)
    worksheet260.write('Y5', 'EKO', body)
    worksheet260.write('Z5', 'SOS', body)
    worksheet260.write('AA5', 'FIS', body)
    worksheet260.write('AB5', 'KIM', body)
    worksheet260.write('AC5', 'BIO', body)
    worksheet260.write('AD5', 'JML', body)

    worksheet260.conditional_format(5, 0, row260_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet260.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_260}', title)
    worksheet260.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet260.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet260.write('A22', 'LOKASI', header)
    worksheet260.write('B22', 'TOTAL', header)
    worksheet260.merge_range('A21:B21', 'RANK', header)
    worksheet260.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet260.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet260.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet260.merge_range('F21:F22', 'KELAS', header)
    worksheet260.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet260.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet260.write('G22', 'MAW', body)
    worksheet260.write('H22', 'MAP', body)
    worksheet260.write('I22', 'IND', body)
    worksheet260.write('J22', 'ENG', body)
    worksheet260.write('K22', 'SEJ', body)
    worksheet260.write('L22', 'GEO', body)
    worksheet260.write('M22', 'EKO', body)
    worksheet260.write('N22', 'SOS', body)
    worksheet260.write('O22', 'FIS', body)
    worksheet260.write('P22', 'KIM', body)
    worksheet260.write('Q22', 'BIO', body)
    worksheet260.write('R22', 'JML', body)
    worksheet260.write('S22', 'MAW', body)
    worksheet260.write('T22', 'MAP', body)
    worksheet260.write('U22', 'IND', body)
    worksheet260.write('V22', 'ENG', body)
    worksheet260.write('W22', 'SEJ', body)
    worksheet260.write('X22', 'GEO', body)
    worksheet260.write('Y22', 'EKO', body)
    worksheet260.write('Z22', 'SOS', body)
    worksheet260.write('AA22', 'FIS', body)
    worksheet260.write('AB22', 'KIM', body)
    worksheet260.write('AC22', 'BIO', body)
    worksheet260.write('AD22', 'JML', body)

    worksheet260.conditional_format(22, 0, row260+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 261
    worksheet261.insert_image('A1', r'logo resmi nf.jpg')

    worksheet261.set_column('A:A', 7, center)
    worksheet261.set_column('B:B', 6, center)
    worksheet261.set_column('C:C', 18.14, center)
    worksheet261.set_column('D:D', 25, left)
    worksheet261.set_column('E:E', 13.14, left)
    worksheet261.set_column('F:F', 8.57, center)
    worksheet261.set_column('G:AD', 5, center)
    worksheet261.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_261}', title)
    worksheet261.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet261.write('A5', 'LOKASI', header)
    worksheet261.write('B5', 'TOTAL', header)
    worksheet261.merge_range('A4:B4', 'RANK', header)
    worksheet261.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet261.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet261.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet261.merge_range('F4:F5', 'KELAS', header)
    worksheet261.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet261.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet261.write('G5', 'MAW', body)
    worksheet261.write('H5', 'MAP', body)
    worksheet261.write('I5', 'IND', body)
    worksheet261.write('J5', 'ENG', body)
    worksheet261.write('K5', 'SEJ', body)
    worksheet261.write('L5', 'GEO', body)
    worksheet261.write('M5', 'EKO', body)
    worksheet261.write('N5', 'SOS', body)
    worksheet261.write('O5', 'FIS', body)
    worksheet261.write('P5', 'KIM', body)
    worksheet261.write('Q5', 'BIO', body)
    worksheet261.write('R5', 'JML', body)
    worksheet261.write('S5', 'MAW', body)
    worksheet261.write('T5', 'MAP', body)
    worksheet261.write('U5', 'IND', body)
    worksheet261.write('V5', 'ENG', body)
    worksheet261.write('W5', 'SEJ', body)
    worksheet261.write('X5', 'GEO', body)
    worksheet261.write('Y5', 'EKO', body)
    worksheet261.write('Z5', 'SOS', body)
    worksheet261.write('AA5', 'FIS', body)
    worksheet261.write('AB5', 'KIM', body)
    worksheet261.write('AC5', 'BIO', body)
    worksheet261.write('AD5', 'JML', body)

    worksheet261.conditional_format(5, 0, row261_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet261.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_261}', title)
    worksheet261.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet261.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet261.write('A22', 'LOKASI', header)
    worksheet261.write('B22', 'TOTAL', header)
    worksheet261.merge_range('A21:B21', 'RANK', header)
    worksheet261.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet261.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet261.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet261.merge_range('F21:F22', 'KELAS', header)
    worksheet261.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet261.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet261.write('G22', 'MAW', body)
    worksheet261.write('H22', 'MAP', body)
    worksheet261.write('I22', 'IND', body)
    worksheet261.write('J22', 'ENG', body)
    worksheet261.write('K22', 'SEJ', body)
    worksheet261.write('L22', 'GEO', body)
    worksheet261.write('M22', 'EKO', body)
    worksheet261.write('N22', 'SOS', body)
    worksheet261.write('O22', 'FIS', body)
    worksheet261.write('P22', 'KIM', body)
    worksheet261.write('Q22', 'BIO', body)
    worksheet261.write('R22', 'JML', body)
    worksheet261.write('S22', 'MAW', body)
    worksheet261.write('T22', 'MAP', body)
    worksheet261.write('U22', 'IND', body)
    worksheet261.write('V22', 'ENG', body)
    worksheet261.write('W22', 'SEJ', body)
    worksheet261.write('X22', 'GEO', body)
    worksheet261.write('Y22', 'EKO', body)
    worksheet261.write('Z22', 'SOS', body)
    worksheet261.write('AA22', 'FIS', body)
    worksheet261.write('AB22', 'KIM', body)
    worksheet261.write('AC22', 'BIO', body)
    worksheet261.write('AD22', 'JML', body)

    worksheet261.conditional_format(22, 0, row261+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 262
    worksheet262.insert_image('A1', r'logo resmi nf.jpg')

    worksheet262.set_column('A:A', 7, center)
    worksheet262.set_column('B:B', 6, center)
    worksheet262.set_column('C:C', 18.14, center)
    worksheet262.set_column('D:D', 25, left)
    worksheet262.set_column('E:E', 13.14, left)
    worksheet262.set_column('F:F', 8.57, center)
    worksheet262.set_column('G:AD', 5, center)
    worksheet262.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_262}', title)
    worksheet262.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet262.write('A5', 'LOKASI', header)
    worksheet262.write('B5', 'TOTAL', header)
    worksheet262.merge_range('A4:B4', 'RANK', header)
    worksheet262.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet262.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet262.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet262.merge_range('F4:F5', 'KELAS', header)
    worksheet262.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet262.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet262.write('G5', 'MAW', body)
    worksheet262.write('H5', 'MAP', body)
    worksheet262.write('I5', 'IND', body)
    worksheet262.write('J5', 'ENG', body)
    worksheet262.write('K5', 'SEJ', body)
    worksheet262.write('L5', 'GEO', body)
    worksheet262.write('M5', 'EKO', body)
    worksheet262.write('N5', 'SOS', body)
    worksheet262.write('O5', 'FIS', body)
    worksheet262.write('P5', 'KIM', body)
    worksheet262.write('Q5', 'BIO', body)
    worksheet262.write('R5', 'JML', body)
    worksheet262.write('S5', 'MAW', body)
    worksheet262.write('T5', 'MAP', body)
    worksheet262.write('U5', 'IND', body)
    worksheet262.write('V5', 'ENG', body)
    worksheet262.write('W5', 'SEJ', body)
    worksheet262.write('X5', 'GEO', body)
    worksheet262.write('Y5', 'EKO', body)
    worksheet262.write('Z5', 'SOS', body)
    worksheet262.write('AA5', 'FIS', body)
    worksheet262.write('AB5', 'KIM', body)
    worksheet262.write('AC5', 'BIO', body)
    worksheet262.write('AD5', 'JML', body)

    worksheet262.conditional_format(5, 0, row262_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet262.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_262}', title)
    worksheet262.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet262.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet262.write('A22', 'LOKASI', header)
    worksheet262.write('B22', 'TOTAL', header)
    worksheet262.merge_range('A21:B21', 'RANK', header)
    worksheet262.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet262.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet262.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet262.merge_range('F21:F22', 'KELAS', header)
    worksheet262.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet262.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet262.write('G22', 'MAW', body)
    worksheet262.write('H22', 'MAP', body)
    worksheet262.write('I22', 'IND', body)
    worksheet262.write('J22', 'ENG', body)
    worksheet262.write('K22', 'SEJ', body)
    worksheet262.write('L22', 'GEO', body)
    worksheet262.write('M22', 'EKO', body)
    worksheet262.write('N22', 'SOS', body)
    worksheet262.write('O22', 'FIS', body)
    worksheet262.write('P22', 'KIM', body)
    worksheet262.write('Q22', 'BIO', body)
    worksheet262.write('R22', 'JML', body)
    worksheet262.write('S22', 'MAW', body)
    worksheet262.write('T22', 'MAP', body)
    worksheet262.write('U22', 'IND', body)
    worksheet262.write('V22', 'ENG', body)
    worksheet262.write('W22', 'SEJ', body)
    worksheet262.write('X22', 'GEO', body)
    worksheet262.write('Y22', 'EKO', body)
    worksheet262.write('Z22', 'SOS', body)
    worksheet262.write('AA22', 'FIS', body)
    worksheet262.write('AB22', 'KIM', body)
    worksheet262.write('AC22', 'BIO', body)
    worksheet262.write('AD22', 'JML', body)

    worksheet262.conditional_format(22, 0, row262+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 263
    worksheet263.insert_image('A1', r'logo resmi nf.jpg')

    worksheet263.set_column('A:A', 7, center)
    worksheet263.set_column('B:B', 6, center)
    worksheet263.set_column('C:C', 18.14, center)
    worksheet263.set_column('D:D', 25, left)
    worksheet263.set_column('E:E', 13.14, left)
    worksheet263.set_column('F:F', 8.57, center)
    worksheet263.set_column('G:AD', 5, center)
    worksheet263.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_263}', title)
    worksheet263.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet263.write('A5', 'LOKASI', header)
    worksheet263.write('B5', 'TOTAL', header)
    worksheet263.merge_range('A4:B4', 'RANK', header)
    worksheet263.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet263.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet263.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet263.merge_range('F4:F5', 'KELAS', header)
    worksheet263.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet263.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet263.write('G5', 'MAW', body)
    worksheet263.write('H5', 'MAP', body)
    worksheet263.write('I5', 'IND', body)
    worksheet263.write('J5', 'ENG', body)
    worksheet263.write('K5', 'SEJ', body)
    worksheet263.write('L5', 'GEO', body)
    worksheet263.write('M5', 'EKO', body)
    worksheet263.write('N5', 'SOS', body)
    worksheet263.write('O5', 'FIS', body)
    worksheet263.write('P5', 'KIM', body)
    worksheet263.write('Q5', 'BIO', body)
    worksheet263.write('R5', 'JML', body)
    worksheet263.write('S5', 'MAW', body)
    worksheet263.write('T5', 'MAP', body)
    worksheet263.write('U5', 'IND', body)
    worksheet263.write('V5', 'ENG', body)
    worksheet263.write('W5', 'SEJ', body)
    worksheet263.write('X5', 'GEO', body)
    worksheet263.write('Y5', 'EKO', body)
    worksheet263.write('Z5', 'SOS', body)
    worksheet263.write('AA5', 'FIS', body)
    worksheet263.write('AB5', 'KIM', body)
    worksheet263.write('AC5', 'BIO', body)
    worksheet263.write('AD5', 'JML', body)

    worksheet263.conditional_format(5, 0, row263_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet263.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_263}', title)
    worksheet263.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet263.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet263.write('A22', 'LOKASI', header)
    worksheet263.write('B22', 'TOTAL', header)
    worksheet263.merge_range('A21:B21', 'RANK', header)
    worksheet263.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet263.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet263.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet263.merge_range('F21:F22', 'KELAS', header)
    worksheet263.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet263.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet263.write('G22', 'MAW', body)
    worksheet263.write('H22', 'MAP', body)
    worksheet263.write('I22', 'IND', body)
    worksheet263.write('J22', 'ENG', body)
    worksheet263.write('K22', 'SEJ', body)
    worksheet263.write('L22', 'GEO', body)
    worksheet263.write('M22', 'EKO', body)
    worksheet263.write('N22', 'SOS', body)
    worksheet263.write('O22', 'FIS', body)
    worksheet263.write('P22', 'KIM', body)
    worksheet263.write('Q22', 'BIO', body)
    worksheet263.write('R22', 'JML', body)
    worksheet263.write('S22', 'MAW', body)
    worksheet263.write('T22', 'MAP', body)
    worksheet263.write('U22', 'IND', body)
    worksheet263.write('V22', 'ENG', body)
    worksheet263.write('W22', 'SEJ', body)
    worksheet263.write('X22', 'GEO', body)
    worksheet263.write('Y22', 'EKO', body)
    worksheet263.write('Z22', 'SOS', body)
    worksheet263.write('AA22', 'FIS', body)
    worksheet263.write('AB22', 'KIM', body)
    worksheet263.write('AC22', 'BIO', body)
    worksheet263.write('AD22', 'JML', body)

    worksheet263.conditional_format(22, 0, row263+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 264
    worksheet264.insert_image('A1', r'logo resmi nf.jpg')

    worksheet264.set_column('A:A', 7, center)
    worksheet264.set_column('B:B', 6, center)
    worksheet264.set_column('C:C', 18.14, center)
    worksheet264.set_column('D:D', 25, left)
    worksheet264.set_column('E:E', 13.14, left)
    worksheet264.set_column('F:F', 8.57, center)
    worksheet264.set_column('G:AD', 5, center)
    worksheet264.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_264}', title)
    worksheet264.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet264.write('A5', 'LOKASI', header)
    worksheet264.write('B5', 'TOTAL', header)
    worksheet264.merge_range('A4:B4', 'RANK', header)
    worksheet264.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet264.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet264.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet264.merge_range('F4:F5', 'KELAS', header)
    worksheet264.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet264.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet264.write('G5', 'MAW', body)
    worksheet264.write('H5', 'MAP', body)
    worksheet264.write('I5', 'IND', body)
    worksheet264.write('J5', 'ENG', body)
    worksheet264.write('K5', 'SEJ', body)
    worksheet264.write('L5', 'GEO', body)
    worksheet264.write('M5', 'EKO', body)
    worksheet264.write('N5', 'SOS', body)
    worksheet264.write('O5', 'FIS', body)
    worksheet264.write('P5', 'KIM', body)
    worksheet264.write('Q5', 'BIO', body)
    worksheet264.write('R5', 'JML', body)
    worksheet264.write('S5', 'MAW', body)
    worksheet264.write('T5', 'MAP', body)
    worksheet264.write('U5', 'IND', body)
    worksheet264.write('V5', 'ENG', body)
    worksheet264.write('W5', 'SEJ', body)
    worksheet264.write('X5', 'GEO', body)
    worksheet264.write('Y5', 'EKO', body)
    worksheet264.write('Z5', 'SOS', body)
    worksheet264.write('AA5', 'FIS', body)
    worksheet264.write('AB5', 'KIM', body)
    worksheet264.write('AC5', 'BIO', body)
    worksheet264.write('AD5', 'JML', body)

    worksheet264.conditional_format(5, 0, row264_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet264.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_264}', title)
    worksheet264.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet264.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet264.write('A22', 'LOKASI', header)
    worksheet264.write('B22', 'TOTAL', header)
    worksheet264.merge_range('A21:B21', 'RANK', header)
    worksheet264.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet264.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet264.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet264.merge_range('F21:F22', 'KELAS', header)
    worksheet264.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet264.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet264.write('G22', 'MAW', body)
    worksheet264.write('H22', 'MAP', body)
    worksheet264.write('I22', 'IND', body)
    worksheet264.write('J22', 'ENG', body)
    worksheet264.write('K22', 'SEJ', body)
    worksheet264.write('L22', 'GEO', body)
    worksheet264.write('M22', 'EKO', body)
    worksheet264.write('N22', 'SOS', body)
    worksheet264.write('O22', 'FIS', body)
    worksheet264.write('P22', 'KIM', body)
    worksheet264.write('Q22', 'BIO', body)
    worksheet264.write('R22', 'JML', body)
    worksheet264.write('S22', 'MAW', body)
    worksheet264.write('T22', 'MAP', body)
    worksheet264.write('U22', 'IND', body)
    worksheet264.write('V22', 'ENG', body)
    worksheet264.write('W22', 'SEJ', body)
    worksheet264.write('X22', 'GEO', body)
    worksheet264.write('Y22', 'EKO', body)
    worksheet264.write('Z22', 'SOS', body)
    worksheet264.write('AA22', 'FIS', body)
    worksheet264.write('AB22', 'KIM', body)
    worksheet264.write('AC22', 'BIO', body)
    worksheet264.write('AD22', 'JML', body)

    worksheet264.conditional_format(22, 0, row264+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 265
    worksheet265.insert_image('A1', r'logo resmi nf.jpg')

    worksheet265.set_column('A:A', 7, center)
    worksheet265.set_column('B:B', 6, center)
    worksheet265.set_column('C:C', 18.14, center)
    worksheet265.set_column('D:D', 25, left)
    worksheet265.set_column('E:E', 13.14, left)
    worksheet265.set_column('F:F', 8.57, center)
    worksheet265.set_column('G:AD', 5, center)
    worksheet265.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_265}', title)
    worksheet265.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet265.write('A5', 'LOKASI', header)
    worksheet265.write('B5', 'TOTAL', header)
    worksheet265.merge_range('A4:B4', 'RANK', header)
    worksheet265.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet265.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet265.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet265.merge_range('F4:F5', 'KELAS', header)
    worksheet265.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet265.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet265.write('G5', 'MAW', body)
    worksheet265.write('H5', 'MAP', body)
    worksheet265.write('I5', 'IND', body)
    worksheet265.write('J5', 'ENG', body)
    worksheet265.write('K5', 'SEJ', body)
    worksheet265.write('L5', 'GEO', body)
    worksheet265.write('M5', 'EKO', body)
    worksheet265.write('N5', 'SOS', body)
    worksheet265.write('O5', 'FIS', body)
    worksheet265.write('P5', 'KIM', body)
    worksheet265.write('Q5', 'BIO', body)
    worksheet265.write('R5', 'JML', body)
    worksheet265.write('S5', 'MAW', body)
    worksheet265.write('T5', 'MAP', body)
    worksheet265.write('U5', 'IND', body)
    worksheet265.write('V5', 'ENG', body)
    worksheet265.write('W5', 'SEJ', body)
    worksheet265.write('X5', 'GEO', body)
    worksheet265.write('Y5', 'EKO', body)
    worksheet265.write('Z5', 'SOS', body)
    worksheet265.write('AA5', 'FIS', body)
    worksheet265.write('AB5', 'KIM', body)
    worksheet265.write('AC5', 'BIO', body)
    worksheet265.write('AD5', 'JML', body)

    worksheet265.conditional_format(5, 0, row265_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet265.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_265}', title)
    worksheet265.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet265.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet265.write('A22', 'LOKASI', header)
    worksheet265.write('B22', 'TOTAL', header)
    worksheet265.merge_range('A21:B21', 'RANK', header)
    worksheet265.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet265.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet265.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet265.merge_range('F21:F22', 'KELAS', header)
    worksheet265.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet265.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet265.write('G22', 'MAW', body)
    worksheet265.write('H22', 'MAP', body)
    worksheet265.write('I22', 'IND', body)
    worksheet265.write('J22', 'ENG', body)
    worksheet265.write('K22', 'SEJ', body)
    worksheet265.write('L22', 'GEO', body)
    worksheet265.write('M22', 'EKO', body)
    worksheet265.write('N22', 'SOS', body)
    worksheet265.write('O22', 'FIS', body)
    worksheet265.write('P22', 'KIM', body)
    worksheet265.write('Q22', 'BIO', body)
    worksheet265.write('R22', 'JML', body)
    worksheet265.write('S22', 'MAW', body)
    worksheet265.write('T22', 'MAP', body)
    worksheet265.write('U22', 'IND', body)
    worksheet265.write('V22', 'ENG', body)
    worksheet265.write('W22', 'SEJ', body)
    worksheet265.write('X22', 'GEO', body)
    worksheet265.write('Y22', 'EKO', body)
    worksheet265.write('Z22', 'SOS', body)
    worksheet265.write('AA22', 'FIS', body)
    worksheet265.write('AB22', 'KIM', body)
    worksheet265.write('AC22', 'BIO', body)
    worksheet265.write('AD22', 'JML', body)

    worksheet265.conditional_format(22, 0, row265+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 266
    worksheet266.insert_image('A1', r'logo resmi nf.jpg')

    worksheet266.set_column('A:A', 7, center)
    worksheet266.set_column('B:B', 6, center)
    worksheet266.set_column('C:C', 18.14, center)
    worksheet266.set_column('D:D', 25, left)
    worksheet266.set_column('E:E', 13.14, left)
    worksheet266.set_column('F:F', 8.57, center)
    worksheet266.set_column('G:AD', 5, center)
    worksheet266.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_266}', title)
    worksheet266.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet266.write('A5', 'LOKASI', header)
    worksheet266.write('B5', 'TOTAL', header)
    worksheet266.merge_range('A4:B4', 'RANK', header)
    worksheet266.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet266.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet266.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet266.merge_range('F4:F5', 'KELAS', header)
    worksheet266.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet266.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet266.write('G5', 'MAW', body)
    worksheet266.write('H5', 'MAP', body)
    worksheet266.write('I5', 'IND', body)
    worksheet266.write('J5', 'ENG', body)
    worksheet266.write('K5', 'SEJ', body)
    worksheet266.write('L5', 'GEO', body)
    worksheet266.write('M5', 'EKO', body)
    worksheet266.write('N5', 'SOS', body)
    worksheet266.write('O5', 'FIS', body)
    worksheet266.write('P5', 'KIM', body)
    worksheet266.write('Q5', 'BIO', body)
    worksheet266.write('R5', 'JML', body)
    worksheet266.write('S5', 'MAW', body)
    worksheet266.write('T5', 'MAP', body)
    worksheet266.write('U5', 'IND', body)
    worksheet266.write('V5', 'ENG', body)
    worksheet266.write('W5', 'SEJ', body)
    worksheet266.write('X5', 'GEO', body)
    worksheet266.write('Y5', 'EKO', body)
    worksheet266.write('Z5', 'SOS', body)
    worksheet266.write('AA5', 'FIS', body)
    worksheet266.write('AB5', 'KIM', body)
    worksheet266.write('AC5', 'BIO', body)
    worksheet266.write('AD5', 'JML', body)

    worksheet266.conditional_format(5, 0, row266_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet266.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_266}', title)
    worksheet266.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet266.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet266.write('A22', 'LOKASI', header)
    worksheet266.write('B22', 'TOTAL', header)
    worksheet266.merge_range('A21:B21', 'RANK', header)
    worksheet266.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet266.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet266.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet266.merge_range('F21:F22', 'KELAS', header)
    worksheet266.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet266.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet266.write('G22', 'MAW', body)
    worksheet266.write('H22', 'MAP', body)
    worksheet266.write('I22', 'IND', body)
    worksheet266.write('J22', 'ENG', body)
    worksheet266.write('K22', 'SEJ', body)
    worksheet266.write('L22', 'GEO', body)
    worksheet266.write('M22', 'EKO', body)
    worksheet266.write('N22', 'SOS', body)
    worksheet266.write('O22', 'FIS', body)
    worksheet266.write('P22', 'KIM', body)
    worksheet266.write('Q22', 'BIO', body)
    worksheet266.write('R22', 'JML', body)
    worksheet266.write('S22', 'MAW', body)
    worksheet266.write('T22', 'MAP', body)
    worksheet266.write('U22', 'IND', body)
    worksheet266.write('V22', 'ENG', body)
    worksheet266.write('W22', 'SEJ', body)
    worksheet266.write('X22', 'GEO', body)
    worksheet266.write('Y22', 'EKO', body)
    worksheet266.write('Z22', 'SOS', body)
    worksheet266.write('AA22', 'FIS', body)
    worksheet266.write('AB22', 'KIM', body)
    worksheet266.write('AC22', 'BIO', body)
    worksheet266.write('AD22', 'JML', body)

    worksheet266.conditional_format(22, 0, row266+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 267
    worksheet267.insert_image('A1', r'logo resmi nf.jpg')

    worksheet267.set_column('A:A', 7, center)
    worksheet267.set_column('B:B', 6, center)
    worksheet267.set_column('C:C', 18.14, center)
    worksheet267.set_column('D:D', 25, left)
    worksheet267.set_column('E:E', 13.14, left)
    worksheet267.set_column('F:F', 8.57, center)
    worksheet267.set_column('G:AD', 5, center)
    worksheet267.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_267}', title)
    worksheet267.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet267.write('A5', 'LOKASI', header)
    worksheet267.write('B5', 'TOTAL', header)
    worksheet267.merge_range('A4:B4', 'RANK', header)
    worksheet267.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet267.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet267.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet267.merge_range('F4:F5', 'KELAS', header)
    worksheet267.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet267.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet267.write('G5', 'MAW', body)
    worksheet267.write('H5', 'MAP', body)
    worksheet267.write('I5', 'IND', body)
    worksheet267.write('J5', 'ENG', body)
    worksheet267.write('K5', 'SEJ', body)
    worksheet267.write('L5', 'GEO', body)
    worksheet267.write('M5', 'EKO', body)
    worksheet267.write('N5', 'SOS', body)
    worksheet267.write('O5', 'FIS', body)
    worksheet267.write('P5', 'KIM', body)
    worksheet267.write('Q5', 'BIO', body)
    worksheet267.write('R5', 'JML', body)
    worksheet267.write('S5', 'MAW', body)
    worksheet267.write('T5', 'MAP', body)
    worksheet267.write('U5', 'IND', body)
    worksheet267.write('V5', 'ENG', body)
    worksheet267.write('W5', 'SEJ', body)
    worksheet267.write('X5', 'GEO', body)
    worksheet267.write('Y5', 'EKO', body)
    worksheet267.write('Z5', 'SOS', body)
    worksheet267.write('AA5', 'FIS', body)
    worksheet267.write('AB5', 'KIM', body)
    worksheet267.write('AC5', 'BIO', body)
    worksheet267.write('AD5', 'JML', body)

    worksheet267.conditional_format(5, 0, row267_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet267.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_267}', title)
    worksheet267.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet267.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet267.write('A22', 'LOKASI', header)
    worksheet267.write('B22', 'TOTAL', header)
    worksheet267.merge_range('A21:B21', 'RANK', header)
    worksheet267.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet267.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet267.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet267.merge_range('F21:F22', 'KELAS', header)
    worksheet267.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet267.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet267.write('G22', 'MAW', body)
    worksheet267.write('H22', 'MAP', body)
    worksheet267.write('I22', 'IND', body)
    worksheet267.write('J22', 'ENG', body)
    worksheet267.write('K22', 'SEJ', body)
    worksheet267.write('L22', 'GEO', body)
    worksheet267.write('M22', 'EKO', body)
    worksheet267.write('N22', 'SOS', body)
    worksheet267.write('O22', 'FIS', body)
    worksheet267.write('P22', 'KIM', body)
    worksheet267.write('Q22', 'BIO', body)
    worksheet267.write('R22', 'JML', body)
    worksheet267.write('S22', 'MAW', body)
    worksheet267.write('T22', 'MAP', body)
    worksheet267.write('U22', 'IND', body)
    worksheet267.write('V22', 'ENG', body)
    worksheet267.write('W22', 'SEJ', body)
    worksheet267.write('X22', 'GEO', body)
    worksheet267.write('Y22', 'EKO', body)
    worksheet267.write('Z22', 'SOS', body)
    worksheet267.write('AA22', 'FIS', body)
    worksheet267.write('AB22', 'KIM', body)
    worksheet267.write('AC22', 'BIO', body)
    worksheet267.write('AD22', 'JML', body)

    worksheet267.conditional_format(22, 0, row267+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 268
    worksheet268.insert_image('A1', r'logo resmi nf.jpg')

    worksheet268.set_column('A:A', 7, center)
    worksheet268.set_column('B:B', 6, center)
    worksheet268.set_column('C:C', 18.14, center)
    worksheet268.set_column('D:D', 25, left)
    worksheet268.set_column('E:E', 13.14, left)
    worksheet268.set_column('F:F', 8.57, center)
    worksheet268.set_column('G:AD', 5, center)
    worksheet268.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_268}', title)
    worksheet268.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet268.write('A5', 'LOKASI', header)
    worksheet268.write('B5', 'TOTAL', header)
    worksheet268.merge_range('A4:B4', 'RANK', header)
    worksheet268.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet268.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet268.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet268.merge_range('F4:F5', 'KELAS', header)
    worksheet268.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet268.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet268.write('G5', 'MAW', body)
    worksheet268.write('H5', 'MAP', body)
    worksheet268.write('I5', 'IND', body)
    worksheet268.write('J5', 'ENG', body)
    worksheet268.write('K5', 'SEJ', body)
    worksheet268.write('L5', 'GEO', body)
    worksheet268.write('M5', 'EKO', body)
    worksheet268.write('N5', 'SOS', body)
    worksheet268.write('O5', 'FIS', body)
    worksheet268.write('P5', 'KIM', body)
    worksheet268.write('Q5', 'BIO', body)
    worksheet268.write('R5', 'JML', body)
    worksheet268.write('S5', 'MAW', body)
    worksheet268.write('T5', 'MAP', body)
    worksheet268.write('U5', 'IND', body)
    worksheet268.write('V5', 'ENG', body)
    worksheet268.write('W5', 'SEJ', body)
    worksheet268.write('X5', 'GEO', body)
    worksheet268.write('Y5', 'EKO', body)
    worksheet268.write('Z5', 'SOS', body)
    worksheet268.write('AA5', 'FIS', body)
    worksheet268.write('AB5', 'KIM', body)
    worksheet268.write('AC5', 'BIO', body)
    worksheet268.write('AD5', 'JML', body)

    worksheet268.conditional_format(5, 0, row268_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet268.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_268}', title)
    worksheet268.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet268.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet268.write('A22', 'LOKASI', header)
    worksheet268.write('B22', 'TOTAL', header)
    worksheet268.merge_range('A21:B21', 'RANK', header)
    worksheet268.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet268.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet268.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet268.merge_range('F21:F22', 'KELAS', header)
    worksheet268.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet268.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet268.write('G22', 'MAW', body)
    worksheet268.write('H22', 'MAP', body)
    worksheet268.write('I22', 'IND', body)
    worksheet268.write('J22', 'ENG', body)
    worksheet268.write('K22', 'SEJ', body)
    worksheet268.write('L22', 'GEO', body)
    worksheet268.write('M22', 'EKO', body)
    worksheet268.write('N22', 'SOS', body)
    worksheet268.write('O22', 'FIS', body)
    worksheet268.write('P22', 'KIM', body)
    worksheet268.write('Q22', 'BIO', body)
    worksheet268.write('R22', 'JML', body)
    worksheet268.write('S22', 'MAW', body)
    worksheet268.write('T22', 'MAP', body)
    worksheet268.write('U22', 'IND', body)
    worksheet268.write('V22', 'ENG', body)
    worksheet268.write('W22', 'SEJ', body)
    worksheet268.write('X22', 'GEO', body)
    worksheet268.write('Y22', 'EKO', body)
    worksheet268.write('Z22', 'SOS', body)
    worksheet268.write('AA22', 'FIS', body)
    worksheet268.write('AB22', 'KIM', body)
    worksheet268.write('AC22', 'BIO', body)
    worksheet268.write('AD22', 'JML', body)

    worksheet268.conditional_format(22, 0, row268+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 269
    worksheet269.insert_image('A1', r'logo resmi nf.jpg')

    worksheet269.set_column('A:A', 7, center)
    worksheet269.set_column('B:B', 6, center)
    worksheet269.set_column('C:C', 18.14, center)
    worksheet269.set_column('D:D', 25, left)
    worksheet269.set_column('E:E', 13.14, left)
    worksheet269.set_column('F:F', 8.57, center)
    worksheet269.set_column('G:AD', 5, center)
    worksheet269.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_269}', title)
    worksheet269.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet269.write('A5', 'LOKASI', header)
    worksheet269.write('B5', 'TOTAL', header)
    worksheet269.merge_range('A4:B4', 'RANK', header)
    worksheet269.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet269.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet269.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet269.merge_range('F4:F5', 'KELAS', header)
    worksheet269.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet269.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet269.write('G5', 'MAW', body)
    worksheet269.write('H5', 'MAP', body)
    worksheet269.write('I5', 'IND', body)
    worksheet269.write('J5', 'ENG', body)
    worksheet269.write('K5', 'SEJ', body)
    worksheet269.write('L5', 'GEO', body)
    worksheet269.write('M5', 'EKO', body)
    worksheet269.write('N5', 'SOS', body)
    worksheet269.write('O5', 'FIS', body)
    worksheet269.write('P5', 'KIM', body)
    worksheet269.write('Q5', 'BIO', body)
    worksheet269.write('R5', 'JML', body)
    worksheet269.write('S5', 'MAW', body)
    worksheet269.write('T5', 'MAP', body)
    worksheet269.write('U5', 'IND', body)
    worksheet269.write('V5', 'ENG', body)
    worksheet269.write('W5', 'SEJ', body)
    worksheet269.write('X5', 'GEO', body)
    worksheet269.write('Y5', 'EKO', body)
    worksheet269.write('Z5', 'SOS', body)
    worksheet269.write('AA5', 'FIS', body)
    worksheet269.write('AB5', 'KIM', body)
    worksheet269.write('AC5', 'BIO', body)
    worksheet269.write('AD5', 'JML', body)

    worksheet269.conditional_format(5, 0, row269_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet269.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_269}', title)
    worksheet269.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet269.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet269.write('A22', 'LOKASI', header)
    worksheet269.write('B22', 'TOTAL', header)
    worksheet269.merge_range('A21:B21', 'RANK', header)
    worksheet269.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet269.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet269.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet269.merge_range('F21:F22', 'KELAS', header)
    worksheet269.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet269.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet269.write('G22', 'MAW', body)
    worksheet269.write('H22', 'MAP', body)
    worksheet269.write('I22', 'IND', body)
    worksheet269.write('J22', 'ENG', body)
    worksheet269.write('K22', 'SEJ', body)
    worksheet269.write('L22', 'GEO', body)
    worksheet269.write('M22', 'EKO', body)
    worksheet269.write('N22', 'SOS', body)
    worksheet269.write('O22', 'FIS', body)
    worksheet269.write('P22', 'KIM', body)
    worksheet269.write('Q22', 'BIO', body)
    worksheet269.write('R22', 'JML', body)
    worksheet269.write('S22', 'MAW', body)
    worksheet269.write('T22', 'MAP', body)
    worksheet269.write('U22', 'IND', body)
    worksheet269.write('V22', 'ENG', body)
    worksheet269.write('W22', 'SEJ', body)
    worksheet269.write('X22', 'GEO', body)
    worksheet269.write('Y22', 'EKO', body)
    worksheet269.write('Z22', 'SOS', body)
    worksheet269.write('AA22', 'FIS', body)
    worksheet269.write('AB22', 'KIM', body)
    worksheet269.write('AC22', 'BIO', body)
    worksheet269.write('AD22', 'JML', body)

    worksheet269.conditional_format(22, 0, row269+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 270
    worksheet270.insert_image('A1', r'logo resmi nf.jpg')

    worksheet270.set_column('A:A', 7, center)
    worksheet270.set_column('B:B', 6, center)
    worksheet270.set_column('C:C', 18.14, center)
    worksheet270.set_column('D:D', 25, left)
    worksheet270.set_column('E:E', 13.14, left)
    worksheet270.set_column('F:F', 8.57, center)
    worksheet270.set_column('G:AD', 5, center)
    worksheet270.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_270}', title)
    worksheet270.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet270.write('A5', 'LOKASI', header)
    worksheet270.write('B5', 'TOTAL', header)
    worksheet270.merge_range('A4:B4', 'RANK', header)
    worksheet270.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet270.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet270.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet270.merge_range('F4:F5', 'KELAS', header)
    worksheet270.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet270.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet270.write('G5', 'MAW', body)
    worksheet270.write('H5', 'MAP', body)
    worksheet270.write('I5', 'IND', body)
    worksheet270.write('J5', 'ENG', body)
    worksheet270.write('K5', 'SEJ', body)
    worksheet270.write('L5', 'GEO', body)
    worksheet270.write('M5', 'EKO', body)
    worksheet270.write('N5', 'SOS', body)
    worksheet270.write('O5', 'FIS', body)
    worksheet270.write('P5', 'KIM', body)
    worksheet270.write('Q5', 'BIO', body)
    worksheet270.write('R5', 'JML', body)
    worksheet270.write('S5', 'MAW', body)
    worksheet270.write('T5', 'MAP', body)
    worksheet270.write('U5', 'IND', body)
    worksheet270.write('V5', 'ENG', body)
    worksheet270.write('W5', 'SEJ', body)
    worksheet270.write('X5', 'GEO', body)
    worksheet270.write('Y5', 'EKO', body)
    worksheet270.write('Z5', 'SOS', body)
    worksheet270.write('AA5', 'FIS', body)
    worksheet270.write('AB5', 'KIM', body)
    worksheet270.write('AC5', 'BIO', body)
    worksheet270.write('AD5', 'JML', body)

    worksheet270.conditional_format(5, 0, row270_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet270.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_270}', title)
    worksheet270.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet270.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet270.write('A22', 'LOKASI', header)
    worksheet270.write('B22', 'TOTAL', header)
    worksheet270.merge_range('A21:B21', 'RANK', header)
    worksheet270.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet270.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet270.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet270.merge_range('F21:F22', 'KELAS', header)
    worksheet270.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet270.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet270.write('G22', 'MAW', body)
    worksheet270.write('H22', 'MAP', body)
    worksheet270.write('I22', 'IND', body)
    worksheet270.write('J22', 'ENG', body)
    worksheet270.write('K22', 'SEJ', body)
    worksheet270.write('L22', 'GEO', body)
    worksheet270.write('M22', 'EKO', body)
    worksheet270.write('N22', 'SOS', body)
    worksheet270.write('O22', 'FIS', body)
    worksheet270.write('P22', 'KIM', body)
    worksheet270.write('Q22', 'BIO', body)
    worksheet270.write('R22', 'JML', body)
    worksheet270.write('S22', 'MAW', body)
    worksheet270.write('T22', 'MAP', body)
    worksheet270.write('U22', 'IND', body)
    worksheet270.write('V22', 'ENG', body)
    worksheet270.write('W22', 'SEJ', body)
    worksheet270.write('X22', 'GEO', body)
    worksheet270.write('Y22', 'EKO', body)
    worksheet270.write('Z22', 'SOS', body)
    worksheet270.write('AA22', 'FIS', body)
    worksheet270.write('AB22', 'KIM', body)
    worksheet270.write('AC22', 'BIO', body)
    worksheet270.write('AD22', 'JML', body)

    worksheet270.conditional_format(22, 0, row270+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 271
    worksheet271.insert_image('A1', r'logo resmi nf.jpg')

    worksheet271.set_column('A:A', 7, center)
    worksheet271.set_column('B:B', 6, center)
    worksheet271.set_column('C:C', 18.14, center)
    worksheet271.set_column('D:D', 25, left)
    worksheet271.set_column('E:E', 13.14, left)
    worksheet271.set_column('F:F', 8.57, center)
    worksheet271.set_column('G:AD', 5, center)
    worksheet271.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_271}', title)
    worksheet271.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet271.write('A5', 'LOKASI', header)
    worksheet271.write('B5', 'TOTAL', header)
    worksheet271.merge_range('A4:B4', 'RANK', header)
    worksheet271.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet271.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet271.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet271.merge_range('F4:F5', 'KELAS', header)
    worksheet271.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet271.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet271.write('G5', 'MAW', body)
    worksheet271.write('H5', 'MAP', body)
    worksheet271.write('I5', 'IND', body)
    worksheet271.write('J5', 'ENG', body)
    worksheet271.write('K5', 'SEJ', body)
    worksheet271.write('L5', 'GEO', body)
    worksheet271.write('M5', 'EKO', body)
    worksheet271.write('N5', 'SOS', body)
    worksheet271.write('O5', 'FIS', body)
    worksheet271.write('P5', 'KIM', body)
    worksheet271.write('Q5', 'BIO', body)
    worksheet271.write('R5', 'JML', body)
    worksheet271.write('S5', 'MAW', body)
    worksheet271.write('T5', 'MAP', body)
    worksheet271.write('U5', 'IND', body)
    worksheet271.write('V5', 'ENG', body)
    worksheet271.write('W5', 'SEJ', body)
    worksheet271.write('X5', 'GEO', body)
    worksheet271.write('Y5', 'EKO', body)
    worksheet271.write('Z5', 'SOS', body)
    worksheet271.write('AA5', 'FIS', body)
    worksheet271.write('AB5', 'KIM', body)
    worksheet271.write('AC5', 'BIO', body)
    worksheet271.write('AD5', 'JML', body)

    worksheet271.conditional_format(5, 0, row271_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet271.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_271}', title)
    worksheet271.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet271.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet271.write('A22', 'LOKASI', header)
    worksheet271.write('B22', 'TOTAL', header)
    worksheet271.merge_range('A21:B21', 'RANK', header)
    worksheet271.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet271.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet271.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet271.merge_range('F21:F22', 'KELAS', header)
    worksheet271.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet271.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet271.write('G22', 'MAW', body)
    worksheet271.write('H22', 'MAP', body)
    worksheet271.write('I22', 'IND', body)
    worksheet271.write('J22', 'ENG', body)
    worksheet271.write('K22', 'SEJ', body)
    worksheet271.write('L22', 'GEO', body)
    worksheet271.write('M22', 'EKO', body)
    worksheet271.write('N22', 'SOS', body)
    worksheet271.write('O22', 'FIS', body)
    worksheet271.write('P22', 'KIM', body)
    worksheet271.write('Q22', 'BIO', body)
    worksheet271.write('R22', 'JML', body)
    worksheet271.write('S22', 'MAW', body)
    worksheet271.write('T22', 'MAP', body)
    worksheet271.write('U22', 'IND', body)
    worksheet271.write('V22', 'ENG', body)
    worksheet271.write('W22', 'SEJ', body)
    worksheet271.write('X22', 'GEO', body)
    worksheet271.write('Y22', 'EKO', body)
    worksheet271.write('Z22', 'SOS', body)
    worksheet271.write('AA22', 'FIS', body)
    worksheet271.write('AB22', 'KIM', body)
    worksheet271.write('AC22', 'BIO', body)
    worksheet271.write('AD22', 'JML', body)

    worksheet271.conditional_format(22, 0, row271+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 272
    worksheet272.insert_image('A1', r'logo resmi nf.jpg')

    worksheet272.set_column('A:A', 7, center)
    worksheet272.set_column('B:B', 6, center)
    worksheet272.set_column('C:C', 18.14, center)
    worksheet272.set_column('D:D', 25, left)
    worksheet272.set_column('E:E', 13.14, left)
    worksheet272.set_column('F:F', 8.57, center)
    worksheet272.set_column('G:AD', 5, center)
    worksheet272.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_272}', title)
    worksheet272.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet272.write('A5', 'LOKASI', header)
    worksheet272.write('B5', 'TOTAL', header)
    worksheet272.merge_range('A4:B4', 'RANK', header)
    worksheet272.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet272.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet272.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet272.merge_range('F4:F5', 'KELAS', header)
    worksheet272.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet272.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet272.write('G5', 'MAW', body)
    worksheet272.write('H5', 'MAP', body)
    worksheet272.write('I5', 'IND', body)
    worksheet272.write('J5', 'ENG', body)
    worksheet272.write('K5', 'SEJ', body)
    worksheet272.write('L5', 'GEO', body)
    worksheet272.write('M5', 'EKO', body)
    worksheet272.write('N5', 'SOS', body)
    worksheet272.write('O5', 'FIS', body)
    worksheet272.write('P5', 'KIM', body)
    worksheet272.write('Q5', 'BIO', body)
    worksheet272.write('R5', 'JML', body)
    worksheet272.write('S5', 'MAW', body)
    worksheet272.write('T5', 'MAP', body)
    worksheet272.write('U5', 'IND', body)
    worksheet272.write('V5', 'ENG', body)
    worksheet272.write('W5', 'SEJ', body)
    worksheet272.write('X5', 'GEO', body)
    worksheet272.write('Y5', 'EKO', body)
    worksheet272.write('Z5', 'SOS', body)
    worksheet272.write('AA5', 'FIS', body)
    worksheet272.write('AB5', 'KIM', body)
    worksheet272.write('AC5', 'BIO', body)
    worksheet272.write('AD5', 'JML', body)

    worksheet272.conditional_format(5, 0, row272_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet272.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_272}', title)
    worksheet272.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet272.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet272.write('A22', 'LOKASI', header)
    worksheet272.write('B22', 'TOTAL', header)
    worksheet272.merge_range('A21:B21', 'RANK', header)
    worksheet272.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet272.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet272.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet272.merge_range('F21:F22', 'KELAS', header)
    worksheet272.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet272.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet272.write('G22', 'MAW', body)
    worksheet272.write('H22', 'MAP', body)
    worksheet272.write('I22', 'IND', body)
    worksheet272.write('J22', 'ENG', body)
    worksheet272.write('K22', 'SEJ', body)
    worksheet272.write('L22', 'GEO', body)
    worksheet272.write('M22', 'EKO', body)
    worksheet272.write('N22', 'SOS', body)
    worksheet272.write('O22', 'FIS', body)
    worksheet272.write('P22', 'KIM', body)
    worksheet272.write('Q22', 'BIO', body)
    worksheet272.write('R22', 'JML', body)
    worksheet272.write('S22', 'MAW', body)
    worksheet272.write('T22', 'MAP', body)
    worksheet272.write('U22', 'IND', body)
    worksheet272.write('V22', 'ENG', body)
    worksheet272.write('W22', 'SEJ', body)
    worksheet272.write('X22', 'GEO', body)
    worksheet272.write('Y22', 'EKO', body)
    worksheet272.write('Z22', 'SOS', body)
    worksheet272.write('AA22', 'FIS', body)
    worksheet272.write('AB22', 'KIM', body)
    worksheet272.write('AC22', 'BIO', body)
    worksheet272.write('AD22', 'JML', body)

    worksheet272.conditional_format(22, 0, row272+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 273
    worksheet273.insert_image('A1', r'logo resmi nf.jpg')

    worksheet273.set_column('A:A', 7, center)
    worksheet273.set_column('B:B', 6, center)
    worksheet273.set_column('C:C', 18.14, center)
    worksheet273.set_column('D:D', 25, left)
    worksheet273.set_column('E:E', 13.14, left)
    worksheet273.set_column('F:F', 8.57, center)
    worksheet273.set_column('G:AD', 5, center)
    worksheet273.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_273}', title)
    worksheet273.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet273.write('A5', 'LOKASI', header)
    worksheet273.write('B5', 'TOTAL', header)
    worksheet273.merge_range('A4:B4', 'RANK', header)
    worksheet273.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet273.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet273.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet273.merge_range('F4:F5', 'KELAS', header)
    worksheet273.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet273.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet273.write('G5', 'MAW', body)
    worksheet273.write('H5', 'MAP', body)
    worksheet273.write('I5', 'IND', body)
    worksheet273.write('J5', 'ENG', body)
    worksheet273.write('K5', 'SEJ', body)
    worksheet273.write('L5', 'GEO', body)
    worksheet273.write('M5', 'EKO', body)
    worksheet273.write('N5', 'SOS', body)
    worksheet273.write('O5', 'FIS', body)
    worksheet273.write('P5', 'KIM', body)
    worksheet273.write('Q5', 'BIO', body)
    worksheet273.write('R5', 'JML', body)
    worksheet273.write('S5', 'MAW', body)
    worksheet273.write('T5', 'MAP', body)
    worksheet273.write('U5', 'IND', body)
    worksheet273.write('V5', 'ENG', body)
    worksheet273.write('W5', 'SEJ', body)
    worksheet273.write('X5', 'GEO', body)
    worksheet273.write('Y5', 'EKO', body)
    worksheet273.write('Z5', 'SOS', body)
    worksheet273.write('AA5', 'FIS', body)
    worksheet273.write('AB5', 'KIM', body)
    worksheet273.write('AC5', 'BIO', body)
    worksheet273.write('AD5', 'JML', body)

    worksheet273.conditional_format(5, 0, row273_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet273.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_273}', title)
    worksheet273.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet273.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet273.write('A22', 'LOKASI', header)
    worksheet273.write('B22', 'TOTAL', header)
    worksheet273.merge_range('A21:B21', 'RANK', header)
    worksheet273.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet273.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet273.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet273.merge_range('F21:F22', 'KELAS', header)
    worksheet273.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet273.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet273.write('G22', 'MAW', body)
    worksheet273.write('H22', 'MAP', body)
    worksheet273.write('I22', 'IND', body)
    worksheet273.write('J22', 'ENG', body)
    worksheet273.write('K22', 'SEJ', body)
    worksheet273.write('L22', 'GEO', body)
    worksheet273.write('M22', 'EKO', body)
    worksheet273.write('N22', 'SOS', body)
    worksheet273.write('O22', 'FIS', body)
    worksheet273.write('P22', 'KIM', body)
    worksheet273.write('Q22', 'BIO', body)
    worksheet273.write('R22', 'JML', body)
    worksheet273.write('S22', 'MAW', body)
    worksheet273.write('T22', 'MAP', body)
    worksheet273.write('U22', 'IND', body)
    worksheet273.write('V22', 'ENG', body)
    worksheet273.write('W22', 'SEJ', body)
    worksheet273.write('X22', 'GEO', body)
    worksheet273.write('Y22', 'EKO', body)
    worksheet273.write('Z22', 'SOS', body)
    worksheet273.write('AA22', 'FIS', body)
    worksheet273.write('AB22', 'KIM', body)
    worksheet273.write('AC22', 'BIO', body)
    worksheet273.write('AD22', 'JML', body)

    worksheet273.conditional_format(22, 0, row273+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 274
    worksheet274.insert_image('A1', r'logo resmi nf.jpg')

    worksheet274.set_column('A:A', 7, center)
    worksheet274.set_column('B:B', 6, center)
    worksheet274.set_column('C:C', 18.14, center)
    worksheet274.set_column('D:D', 25, left)
    worksheet274.set_column('E:E', 13.14, left)
    worksheet274.set_column('F:F', 8.57, center)
    worksheet274.set_column('G:AD', 5, center)
    worksheet274.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_274}', title)
    worksheet274.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet274.write('A5', 'LOKASI', header)
    worksheet274.write('B5', 'TOTAL', header)
    worksheet274.merge_range('A4:B4', 'RANK', header)
    worksheet274.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet274.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet274.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet274.merge_range('F4:F5', 'KELAS', header)
    worksheet274.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet274.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet274.write('G5', 'MAW', body)
    worksheet274.write('H5', 'MAP', body)
    worksheet274.write('I5', 'IND', body)
    worksheet274.write('J5', 'ENG', body)
    worksheet274.write('K5', 'SEJ', body)
    worksheet274.write('L5', 'GEO', body)
    worksheet274.write('M5', 'EKO', body)
    worksheet274.write('N5', 'SOS', body)
    worksheet274.write('O5', 'FIS', body)
    worksheet274.write('P5', 'KIM', body)
    worksheet274.write('Q5', 'BIO', body)
    worksheet274.write('R5', 'JML', body)
    worksheet274.write('S5', 'MAW', body)
    worksheet274.write('T5', 'MAP', body)
    worksheet274.write('U5', 'IND', body)
    worksheet274.write('V5', 'ENG', body)
    worksheet274.write('W5', 'SEJ', body)
    worksheet274.write('X5', 'GEO', body)
    worksheet274.write('Y5', 'EKO', body)
    worksheet274.write('Z5', 'SOS', body)
    worksheet274.write('AA5', 'FIS', body)
    worksheet274.write('AB5', 'KIM', body)
    worksheet274.write('AC5', 'BIO', body)
    worksheet274.write('AD5', 'JML', body)

    worksheet274.conditional_format(5, 0, row274_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet274.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_274}', title)
    worksheet274.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet274.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet274.write('A22', 'LOKASI', header)
    worksheet274.write('B22', 'TOTAL', header)
    worksheet274.merge_range('A21:B21', 'RANK', header)
    worksheet274.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet274.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet274.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet274.merge_range('F21:F22', 'KELAS', header)
    worksheet274.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet274.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet274.write('G22', 'MAW', body)
    worksheet274.write('H22', 'MAP', body)
    worksheet274.write('I22', 'IND', body)
    worksheet274.write('J22', 'ENG', body)
    worksheet274.write('K22', 'SEJ', body)
    worksheet274.write('L22', 'GEO', body)
    worksheet274.write('M22', 'EKO', body)
    worksheet274.write('N22', 'SOS', body)
    worksheet274.write('O22', 'FIS', body)
    worksheet274.write('P22', 'KIM', body)
    worksheet274.write('Q22', 'BIO', body)
    worksheet274.write('R22', 'JML', body)
    worksheet274.write('S22', 'MAW', body)
    worksheet274.write('T22', 'MAP', body)
    worksheet274.write('U22', 'IND', body)
    worksheet274.write('V22', 'ENG', body)
    worksheet274.write('W22', 'SEJ', body)
    worksheet274.write('X22', 'GEO', body)
    worksheet274.write('Y22', 'EKO', body)
    worksheet274.write('Z22', 'SOS', body)
    worksheet274.write('AA22', 'FIS', body)
    worksheet274.write('AB22', 'KIM', body)
    worksheet274.write('AC22', 'BIO', body)
    worksheet274.write('AD22', 'JML', body)

    worksheet274.conditional_format(22, 0, row274+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 275
    worksheet275.insert_image('A1', r'logo resmi nf.jpg')

    worksheet275.set_column('A:A', 7, center)
    worksheet275.set_column('B:B', 6, center)
    worksheet275.set_column('C:C', 18.14, center)
    worksheet275.set_column('D:D', 25, left)
    worksheet275.set_column('E:E', 13.14, left)
    worksheet275.set_column('F:F', 8.57, center)
    worksheet275.set_column('G:AD', 5, center)
    worksheet275.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_275}', title)
    worksheet275.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet275.write('A5', 'LOKASI', header)
    worksheet275.write('B5', 'TOTAL', header)
    worksheet275.merge_range('A4:B4', 'RANK', header)
    worksheet275.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet275.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet275.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet275.merge_range('F4:F5', 'KELAS', header)
    worksheet275.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet275.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet275.write('G5', 'MAW', body)
    worksheet275.write('H5', 'MAP', body)
    worksheet275.write('I5', 'IND', body)
    worksheet275.write('J5', 'ENG', body)
    worksheet275.write('K5', 'SEJ', body)
    worksheet275.write('L5', 'GEO', body)
    worksheet275.write('M5', 'EKO', body)
    worksheet275.write('N5', 'SOS', body)
    worksheet275.write('O5', 'FIS', body)
    worksheet275.write('P5', 'KIM', body)
    worksheet275.write('Q5', 'BIO', body)
    worksheet275.write('R5', 'JML', body)
    worksheet275.write('S5', 'MAW', body)
    worksheet275.write('T5', 'MAP', body)
    worksheet275.write('U5', 'IND', body)
    worksheet275.write('V5', 'ENG', body)
    worksheet275.write('W5', 'SEJ', body)
    worksheet275.write('X5', 'GEO', body)
    worksheet275.write('Y5', 'EKO', body)
    worksheet275.write('Z5', 'SOS', body)
    worksheet275.write('AA5', 'FIS', body)
    worksheet275.write('AB5', 'KIM', body)
    worksheet275.write('AC5', 'BIO', body)
    worksheet275.write('AD5', 'JML', body)

    worksheet275.conditional_format(5, 0, row275_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet275.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_275}', title)
    worksheet275.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet275.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet275.write('A22', 'LOKASI', header)
    worksheet275.write('B22', 'TOTAL', header)
    worksheet275.merge_range('A21:B21', 'RANK', header)
    worksheet275.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet275.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet275.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet275.merge_range('F21:F22', 'KELAS', header)
    worksheet275.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet275.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet275.write('G22', 'MAW', body)
    worksheet275.write('H22', 'MAP', body)
    worksheet275.write('I22', 'IND', body)
    worksheet275.write('J22', 'ENG', body)
    worksheet275.write('K22', 'SEJ', body)
    worksheet275.write('L22', 'GEO', body)
    worksheet275.write('M22', 'EKO', body)
    worksheet275.write('N22', 'SOS', body)
    worksheet275.write('O22', 'FIS', body)
    worksheet275.write('P22', 'KIM', body)
    worksheet275.write('Q22', 'BIO', body)
    worksheet275.write('R22', 'JML', body)
    worksheet275.write('S22', 'MAW', body)
    worksheet275.write('T22', 'MAP', body)
    worksheet275.write('U22', 'IND', body)
    worksheet275.write('V22', 'ENG', body)
    worksheet275.write('W22', 'SEJ', body)
    worksheet275.write('X22', 'GEO', body)
    worksheet275.write('Y22', 'EKO', body)
    worksheet275.write('Z22', 'SOS', body)
    worksheet275.write('AA22', 'FIS', body)
    worksheet275.write('AB22', 'KIM', body)
    worksheet275.write('AC22', 'BIO', body)
    worksheet275.write('AD22', 'JML', body)

    worksheet275.conditional_format(22, 0, row275+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 276
    worksheet276.insert_image('A1', r'logo resmi nf.jpg')

    worksheet276.set_column('A:A', 7, center)
    worksheet276.set_column('B:B', 6, center)
    worksheet276.set_column('C:C', 18.14, center)
    worksheet276.set_column('D:D', 25, left)
    worksheet276.set_column('E:E', 13.14, left)
    worksheet276.set_column('F:F', 8.57, center)
    worksheet276.set_column('G:AD', 5, center)
    worksheet276.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_276}', title)
    worksheet276.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet276.write('A5', 'LOKASI', header)
    worksheet276.write('B5', 'TOTAL', header)
    worksheet276.merge_range('A4:B4', 'RANK', header)
    worksheet276.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet276.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet276.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet276.merge_range('F4:F5', 'KELAS', header)
    worksheet276.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet276.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet276.write('G5', 'MAW', body)
    worksheet276.write('H5', 'MAP', body)
    worksheet276.write('I5', 'IND', body)
    worksheet276.write('J5', 'ENG', body)
    worksheet276.write('K5', 'SEJ', body)
    worksheet276.write('L5', 'GEO', body)
    worksheet276.write('M5', 'EKO', body)
    worksheet276.write('N5', 'SOS', body)
    worksheet276.write('O5', 'FIS', body)
    worksheet276.write('P5', 'KIM', body)
    worksheet276.write('Q5', 'BIO', body)
    worksheet276.write('R5', 'JML', body)
    worksheet276.write('S5', 'MAW', body)
    worksheet276.write('T5', 'MAP', body)
    worksheet276.write('U5', 'IND', body)
    worksheet276.write('V5', 'ENG', body)
    worksheet276.write('W5', 'SEJ', body)
    worksheet276.write('X5', 'GEO', body)
    worksheet276.write('Y5', 'EKO', body)
    worksheet276.write('Z5', 'SOS', body)
    worksheet276.write('AA5', 'FIS', body)
    worksheet276.write('AB5', 'KIM', body)
    worksheet276.write('AC5', 'BIO', body)
    worksheet276.write('AD5', 'JML', body)

    worksheet276.conditional_format(5, 0, row276_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet276.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_275}', title)
    worksheet276.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet276.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet276.write('A22', 'LOKASI', header)
    worksheet276.write('B22', 'TOTAL', header)
    worksheet276.merge_range('A21:B21', 'RANK', header)
    worksheet276.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet276.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet276.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet276.merge_range('F21:F22', 'KELAS', header)
    worksheet276.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet276.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet276.write('G22', 'MAW', body)
    worksheet276.write('H22', 'MAP', body)
    worksheet276.write('I22', 'IND', body)
    worksheet276.write('J22', 'ENG', body)
    worksheet276.write('K22', 'SEJ', body)
    worksheet276.write('L22', 'GEO', body)
    worksheet276.write('M22', 'EKO', body)
    worksheet276.write('N22', 'SOS', body)
    worksheet276.write('O22', 'FIS', body)
    worksheet276.write('P22', 'KIM', body)
    worksheet276.write('Q22', 'BIO', body)
    worksheet276.write('R22', 'JML', body)
    worksheet276.write('S22', 'MAW', body)
    worksheet276.write('T22', 'MAP', body)
    worksheet276.write('U22', 'IND', body)
    worksheet276.write('V22', 'ENG', body)
    worksheet276.write('W22', 'SEJ', body)
    worksheet276.write('X22', 'GEO', body)
    worksheet276.write('Y22', 'EKO', body)
    worksheet276.write('Z22', 'SOS', body)
    worksheet276.write('AA22', 'FIS', body)
    worksheet276.write('AB22', 'KIM', body)
    worksheet276.write('AC22', 'BIO', body)
    worksheet276.write('AD22', 'JML', body)

    worksheet276.conditional_format(22, 0, row276+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 277
    worksheet277.insert_image('A1', r'logo resmi nf.jpg')

    worksheet277.set_column('A:A', 7, center)
    worksheet277.set_column('B:B', 6, center)
    worksheet277.set_column('C:C', 18.14, center)
    worksheet277.set_column('D:D', 25, left)
    worksheet277.set_column('E:E', 13.14, left)
    worksheet277.set_column('F:F', 8.57, center)
    worksheet277.set_column('G:AD', 5, center)
    worksheet277.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_277}', title)
    worksheet277.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet277.write('A5', 'LOKASI', header)
    worksheet277.write('B5', 'TOTAL', header)
    worksheet277.merge_range('A4:B4', 'RANK', header)
    worksheet277.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet277.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet277.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet277.merge_range('F4:F5', 'KELAS', header)
    worksheet277.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet277.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet277.write('G5', 'MAW', body)
    worksheet277.write('H5', 'MAP', body)
    worksheet277.write('I5', 'IND', body)
    worksheet277.write('J5', 'ENG', body)
    worksheet277.write('K5', 'SEJ', body)
    worksheet277.write('L5', 'GEO', body)
    worksheet277.write('M5', 'EKO', body)
    worksheet277.write('N5', 'SOS', body)
    worksheet277.write('O5', 'FIS', body)
    worksheet277.write('P5', 'KIM', body)
    worksheet277.write('Q5', 'BIO', body)
    worksheet277.write('R5', 'JML', body)
    worksheet277.write('S5', 'MAW', body)
    worksheet277.write('T5', 'MAP', body)
    worksheet277.write('U5', 'IND', body)
    worksheet277.write('V5', 'ENG', body)
    worksheet277.write('W5', 'SEJ', body)
    worksheet277.write('X5', 'GEO', body)
    worksheet277.write('Y5', 'EKO', body)
    worksheet277.write('Z5', 'SOS', body)
    worksheet277.write('AA5', 'FIS', body)
    worksheet277.write('AB5', 'KIM', body)
    worksheet277.write('AC5', 'BIO', body)
    worksheet277.write('AD5', 'JML', body)

    worksheet277.conditional_format(5, 0, row277_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet277.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_277}', title)
    worksheet277.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet277.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet277.write('A22', 'LOKASI', header)
    worksheet277.write('B22', 'TOTAL', header)
    worksheet277.merge_range('A21:B21', 'RANK', header)
    worksheet277.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet277.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet277.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet277.merge_range('F21:F22', 'KELAS', header)
    worksheet277.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet277.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet277.write('G22', 'MAW', body)
    worksheet277.write('H22', 'MAP', body)
    worksheet277.write('I22', 'IND', body)
    worksheet277.write('J22', 'ENG', body)
    worksheet277.write('K22', 'SEJ', body)
    worksheet277.write('L22', 'GEO', body)
    worksheet277.write('M22', 'EKO', body)
    worksheet277.write('N22', 'SOS', body)
    worksheet277.write('O22', 'FIS', body)
    worksheet277.write('P22', 'KIM', body)
    worksheet277.write('Q22', 'BIO', body)
    worksheet277.write('R22', 'JML', body)
    worksheet277.write('S22', 'MAW', body)
    worksheet277.write('T22', 'MAP', body)
    worksheet277.write('U22', 'IND', body)
    worksheet277.write('V22', 'ENG', body)
    worksheet277.write('W22', 'SEJ', body)
    worksheet277.write('X22', 'GEO', body)
    worksheet277.write('Y22', 'EKO', body)
    worksheet277.write('Z22', 'SOS', body)
    worksheet277.write('AA22', 'FIS', body)
    worksheet277.write('AB22', 'KIM', body)
    worksheet277.write('AC22', 'BIO', body)
    worksheet277.write('AD22', 'JML', body)

    worksheet277.conditional_format(22, 0, row277+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 278
    worksheet278.insert_image('A1', r'logo resmi nf.jpg')

    worksheet278.set_column('A:A', 7, center)
    worksheet278.set_column('B:B', 6, center)
    worksheet278.set_column('C:C', 18.14, center)
    worksheet278.set_column('D:D', 25, left)
    worksheet278.set_column('E:E', 13.14, left)
    worksheet278.set_column('F:F', 8.57, center)
    worksheet278.set_column('G:AD', 5, center)
    worksheet278.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_278}', title)
    worksheet278.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet278.write('A5', 'LOKASI', header)
    worksheet278.write('B5', 'TOTAL', header)
    worksheet278.merge_range('A4:B4', 'RANK', header)
    worksheet278.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet278.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet278.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet278.merge_range('F4:F5', 'KELAS', header)
    worksheet278.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet278.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet278.write('G5', 'MAW', body)
    worksheet278.write('H5', 'MAP', body)
    worksheet278.write('I5', 'IND', body)
    worksheet278.write('J5', 'ENG', body)
    worksheet278.write('K5', 'SEJ', body)
    worksheet278.write('L5', 'GEO', body)
    worksheet278.write('M5', 'EKO', body)
    worksheet278.write('N5', 'SOS', body)
    worksheet278.write('O5', 'FIS', body)
    worksheet278.write('P5', 'KIM', body)
    worksheet278.write('Q5', 'BIO', body)
    worksheet278.write('R5', 'JML', body)
    worksheet278.write('S5', 'MAW', body)
    worksheet278.write('T5', 'MAP', body)
    worksheet278.write('U5', 'IND', body)
    worksheet278.write('V5', 'ENG', body)
    worksheet278.write('W5', 'SEJ', body)
    worksheet278.write('X5', 'GEO', body)
    worksheet278.write('Y5', 'EKO', body)
    worksheet278.write('Z5', 'SOS', body)
    worksheet278.write('AA5', 'FIS', body)
    worksheet278.write('AB5', 'KIM', body)
    worksheet278.write('AC5', 'BIO', body)
    worksheet278.write('AD5', 'JML', body)

    worksheet278.conditional_format(5, 0, row278_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet278.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_278}', title)
    worksheet278.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet278.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet278.write('A22', 'LOKASI', header)
    worksheet278.write('B22', 'TOTAL', header)
    worksheet278.merge_range('A21:B21', 'RANK', header)
    worksheet278.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet278.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet278.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet278.merge_range('F21:F22', 'KELAS', header)
    worksheet278.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet278.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet278.write('G22', 'MAW', body)
    worksheet278.write('H22', 'MAP', body)
    worksheet278.write('I22', 'IND', body)
    worksheet278.write('J22', 'ENG', body)
    worksheet278.write('K22', 'SEJ', body)
    worksheet278.write('L22', 'GEO', body)
    worksheet278.write('M22', 'EKO', body)
    worksheet278.write('N22', 'SOS', body)
    worksheet278.write('O22', 'FIS', body)
    worksheet278.write('P22', 'KIM', body)
    worksheet278.write('Q22', 'BIO', body)
    worksheet278.write('R22', 'JML', body)
    worksheet278.write('S22', 'MAW', body)
    worksheet278.write('T22', 'MAP', body)
    worksheet278.write('U22', 'IND', body)
    worksheet278.write('V22', 'ENG', body)
    worksheet278.write('W22', 'SEJ', body)
    worksheet278.write('X22', 'GEO', body)
    worksheet278.write('Y22', 'EKO', body)
    worksheet278.write('Z22', 'SOS', body)
    worksheet278.write('AA22', 'FIS', body)
    worksheet278.write('AB22', 'KIM', body)
    worksheet278.write('AC22', 'BIO', body)
    worksheet278.write('AD22', 'JML', body)

    worksheet278.conditional_format(22, 0, row278+21, 29,
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

    # worksheet 279
    worksheet279.insert_image('A1', r'logo resmi nf.jpg')

    worksheet279.set_column('A:A', 7, center)
    worksheet279.set_column('B:B', 6, center)
    worksheet279.set_column('C:C', 18.14, center)
    worksheet279.set_column('D:D', 25, left)
    worksheet279.set_column('E:E', 13.14, left)
    worksheet279.set_column('F:F', 8.57, center)
    worksheet279.set_column('G:AD', 5, center)
    worksheet279.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_279}', title)
    worksheet279.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet279.write('A5', 'LOKASI', header)
    worksheet279.write('B5', 'TOTAL', header)
    worksheet279.merge_range('A4:B4', 'RANK', header)
    worksheet279.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet279.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet279.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet279.merge_range('F4:F5', 'KELAS', header)
    worksheet279.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet279.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet279.write('G5', 'MAW', body)
    worksheet279.write('H5', 'MAP', body)
    worksheet279.write('I5', 'IND', body)
    worksheet279.write('J5', 'ENG', body)
    worksheet279.write('K5', 'SEJ', body)
    worksheet279.write('L5', 'GEO', body)
    worksheet279.write('M5', 'EKO', body)
    worksheet279.write('N5', 'SOS', body)
    worksheet279.write('O5', 'FIS', body)
    worksheet279.write('P5', 'KIM', body)
    worksheet279.write('Q5', 'BIO', body)
    worksheet279.write('R5', 'JML', body)
    worksheet279.write('S5', 'MAW', body)
    worksheet279.write('T5', 'MAP', body)
    worksheet279.write('U5', 'IND', body)
    worksheet279.write('V5', 'ENG', body)
    worksheet279.write('W5', 'SEJ', body)
    worksheet279.write('X5', 'GEO', body)
    worksheet279.write('Y5', 'EKO', body)
    worksheet279.write('Z5', 'SOS', body)
    worksheet279.write('AA5', 'FIS', body)
    worksheet279.write('AB5', 'KIM', body)
    worksheet279.write('AC5', 'BIO', body)
    worksheet279.write('AD5', 'JML', body)

    worksheet279.conditional_format(5, 0, row279_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet279.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_279}', title)
    worksheet279.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet279.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet279.write('A22', 'LOKASI', header)
    worksheet279.write('B22', 'TOTAL', header)
    worksheet279.merge_range('A21:B21', 'RANK', header)
    worksheet279.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet279.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet279.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet279.merge_range('F21:F22', 'KELAS', header)
    worksheet279.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet279.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet279.write('G22', 'MAW', body)
    worksheet279.write('H22', 'MAP', body)
    worksheet279.write('I22', 'IND', body)
    worksheet279.write('J22', 'ENG', body)
    worksheet279.write('K22', 'SEJ', body)
    worksheet279.write('L22', 'GEO', body)
    worksheet279.write('M22', 'EKO', body)
    worksheet279.write('N22', 'SOS', body)
    worksheet279.write('O22', 'FIS', body)
    worksheet279.write('P22', 'KIM', body)
    worksheet279.write('Q22', 'BIO', body)
    worksheet279.write('R22', 'JML', body)
    worksheet279.write('S22', 'MAW', body)
    worksheet279.write('T22', 'MAP', body)
    worksheet279.write('U22', 'IND', body)
    worksheet279.write('V22', 'ENG', body)
    worksheet279.write('W22', 'SEJ', body)
    worksheet279.write('X22', 'GEO', body)
    worksheet279.write('Y22', 'EKO', body)
    worksheet279.write('Z22', 'SOS', body)
    worksheet279.write('AA22', 'FIS', body)
    worksheet279.write('AB22', 'KIM', body)
    worksheet279.write('AC22', 'BIO', body)
    worksheet279.write('AD22', 'JML', body)

    worksheet279.conditional_format(22, 0, row279+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 280
    worksheet280.insert_image('A1', r'logo resmi nf.jpg')

    worksheet280.set_column('A:A', 7, center)
    worksheet280.set_column('B:B', 6, center)
    worksheet280.set_column('C:C', 18.14, center)
    worksheet280.set_column('D:D', 25, left)
    worksheet280.set_column('E:E', 13.14, left)
    worksheet280.set_column('F:F', 8.57, center)
    worksheet280.set_column('G:AD', 5, center)
    worksheet280.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_280}', title)
    worksheet280.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet280.write('A5', 'LOKASI', header)
    worksheet280.write('B5', 'TOTAL', header)
    worksheet280.merge_range('A4:B4', 'RANK', header)
    worksheet280.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet280.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet280.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet280.merge_range('F4:F5', 'KELAS', header)
    worksheet280.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet280.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet280.write('G5', 'MAW', body)
    worksheet280.write('H5', 'MAP', body)
    worksheet280.write('I5', 'IND', body)
    worksheet280.write('J5', 'ENG', body)
    worksheet280.write('K5', 'SEJ', body)
    worksheet280.write('L5', 'GEO', body)
    worksheet280.write('M5', 'EKO', body)
    worksheet280.write('N5', 'SOS', body)
    worksheet280.write('O5', 'FIS', body)
    worksheet280.write('P5', 'KIM', body)
    worksheet280.write('Q5', 'BIO', body)
    worksheet280.write('R5', 'JML', body)
    worksheet280.write('S5', 'MAW', body)
    worksheet280.write('T5', 'MAP', body)
    worksheet280.write('U5', 'IND', body)
    worksheet280.write('V5', 'ENG', body)
    worksheet280.write('W5', 'SEJ', body)
    worksheet280.write('X5', 'GEO', body)
    worksheet280.write('Y5', 'EKO', body)
    worksheet280.write('Z5', 'SOS', body)
    worksheet280.write('AA5', 'FIS', body)
    worksheet280.write('AB5', 'KIM', body)
    worksheet280.write('AC5', 'BIO', body)
    worksheet280.write('AD5', 'JML', body)

    worksheet280.conditional_format(5, 0, row280_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet280.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_280}', title)
    worksheet280.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet280.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet280.write('A22', 'LOKASI', header)
    worksheet280.write('B22', 'TOTAL', header)
    worksheet280.merge_range('A21:B21', 'RANK', header)
    worksheet280.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet280.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet280.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet280.merge_range('F21:F22', 'KELAS', header)
    worksheet280.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet280.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet280.write('G22', 'MAW', body)
    worksheet280.write('H22', 'MAP', body)
    worksheet280.write('I22', 'IND', body)
    worksheet280.write('J22', 'ENG', body)
    worksheet280.write('K22', 'SEJ', body)
    worksheet280.write('L22', 'GEO', body)
    worksheet280.write('M22', 'EKO', body)
    worksheet280.write('N22', 'SOS', body)
    worksheet280.write('O22', 'FIS', body)
    worksheet280.write('P22', 'KIM', body)
    worksheet280.write('Q22', 'BIO', body)
    worksheet280.write('R22', 'JML', body)
    worksheet280.write('S22', 'MAW', body)
    worksheet280.write('T22', 'MAP', body)
    worksheet280.write('U22', 'IND', body)
    worksheet280.write('V22', 'ENG', body)
    worksheet280.write('W22', 'SEJ', body)
    worksheet280.write('X22', 'GEO', body)
    worksheet280.write('Y22', 'EKO', body)
    worksheet280.write('Z22', 'SOS', body)
    worksheet280.write('AA22', 'FIS', body)
    worksheet280.write('AB22', 'KIM', body)
    worksheet280.write('AC22', 'BIO', body)
    worksheet280.write('AD22', 'JML', body)

    worksheet280.conditional_format(22, 0, row280+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 282
    worksheet282.insert_image('A1', r'logo resmi nf.jpg')

    worksheet282.set_column('A:A', 7, center)
    worksheet282.set_column('B:B', 6, center)
    worksheet282.set_column('C:C', 18.14, center)
    worksheet282.set_column('D:D', 25, left)
    worksheet282.set_column('E:E', 13.14, left)
    worksheet282.set_column('F:F', 8.57, center)
    worksheet282.set_column('G:AD', 5, center)
    worksheet282.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_282}', title)
    worksheet282.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet282.write('A5', 'LOKASI', header)
    worksheet282.write('B5', 'TOTAL', header)
    worksheet282.merge_range('A4:B4', 'RANK', header)
    worksheet282.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet282.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet282.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet282.merge_range('F4:F5', 'KELAS', header)
    worksheet282.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet282.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet282.write('G5', 'MAW', body)
    worksheet282.write('H5', 'MAP', body)
    worksheet282.write('I5', 'IND', body)
    worksheet282.write('J5', 'ENG', body)
    worksheet282.write('K5', 'SEJ', body)
    worksheet282.write('L5', 'GEO', body)
    worksheet282.write('M5', 'EKO', body)
    worksheet282.write('N5', 'SOS', body)
    worksheet282.write('O5', 'FIS', body)
    worksheet282.write('P5', 'KIM', body)
    worksheet282.write('Q5', 'BIO', body)
    worksheet282.write('R5', 'JML', body)
    worksheet282.write('S5', 'MAW', body)
    worksheet282.write('T5', 'MAP', body)
    worksheet282.write('U5', 'IND', body)
    worksheet282.write('V5', 'ENG', body)
    worksheet282.write('W5', 'SEJ', body)
    worksheet282.write('X5', 'GEO', body)
    worksheet282.write('Y5', 'EKO', body)
    worksheet282.write('Z5', 'SOS', body)
    worksheet282.write('AA5', 'FIS', body)
    worksheet282.write('AB5', 'KIM', body)
    worksheet282.write('AC5', 'BIO', body)
    worksheet282.write('AD5', 'JML', body)

    worksheet282.conditional_format(5, 0, row282_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet282.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_282}', title)
    worksheet282.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet282.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet282.write('A22', 'LOKASI', header)
    worksheet282.write('B22', 'TOTAL', header)
    worksheet282.merge_range('A21:B21', 'RANK', header)
    worksheet282.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet282.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet282.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet282.merge_range('F21:F22', 'KELAS', header)
    worksheet282.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet282.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet282.write('G22', 'MAW', body)
    worksheet282.write('H22', 'MAP', body)
    worksheet282.write('I22', 'IND', body)
    worksheet282.write('J22', 'ENG', body)
    worksheet282.write('K22', 'SEJ', body)
    worksheet282.write('L22', 'GEO', body)
    worksheet282.write('M22', 'EKO', body)
    worksheet282.write('N22', 'SOS', body)
    worksheet282.write('O22', 'FIS', body)
    worksheet282.write('P22', 'KIM', body)
    worksheet282.write('Q22', 'BIO', body)
    worksheet282.write('R22', 'JML', body)
    worksheet282.write('S22', 'MAW', body)
    worksheet282.write('T22', 'MAP', body)
    worksheet282.write('U22', 'IND', body)
    worksheet282.write('V22', 'ENG', body)
    worksheet282.write('W22', 'SEJ', body)
    worksheet282.write('X22', 'GEO', body)
    worksheet282.write('Y22', 'EKO', body)
    worksheet282.write('Z22', 'SOS', body)
    worksheet282.write('AA22', 'FIS', body)
    worksheet282.write('AB22', 'KIM', body)
    worksheet282.write('AC22', 'BIO', body)
    worksheet282.write('AD22', 'JML', body)

    worksheet282.conditional_format(22, 0, row282+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 283
    worksheet283.insert_image('A1', r'logo resmi nf.jpg')

    worksheet283.set_column('A:A', 7, center)
    worksheet283.set_column('B:B', 6, center)
    worksheet283.set_column('C:C', 18.14, center)
    worksheet283.set_column('D:D', 25, left)
    worksheet283.set_column('E:E', 13.14, left)
    worksheet283.set_column('F:F', 8.57, center)
    worksheet283.set_column('G:AD', 5, center)
    worksheet283.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_283}', title)
    worksheet283.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet283.write('A5', 'LOKASI', header)
    worksheet283.write('B5', 'TOTAL', header)
    worksheet283.merge_range('A4:B4', 'RANK', header)
    worksheet283.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet283.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet283.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet283.merge_range('F4:F5', 'KELAS', header)
    worksheet283.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet283.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet283.write('G5', 'MAW', body)
    worksheet283.write('H5', 'MAP', body)
    worksheet283.write('I5', 'IND', body)
    worksheet283.write('J5', 'ENG', body)
    worksheet283.write('K5', 'SEJ', body)
    worksheet283.write('L5', 'GEO', body)
    worksheet283.write('M5', 'EKO', body)
    worksheet283.write('N5', 'SOS', body)
    worksheet283.write('O5', 'FIS', body)
    worksheet283.write('P5', 'KIM', body)
    worksheet283.write('Q5', 'BIO', body)
    worksheet283.write('R5', 'JML', body)
    worksheet283.write('S5', 'MAW', body)
    worksheet283.write('T5', 'MAP', body)
    worksheet283.write('U5', 'IND', body)
    worksheet283.write('V5', 'ENG', body)
    worksheet283.write('W5', 'SEJ', body)
    worksheet283.write('X5', 'GEO', body)
    worksheet283.write('Y5', 'EKO', body)
    worksheet283.write('Z5', 'SOS', body)
    worksheet283.write('AA5', 'FIS', body)
    worksheet283.write('AB5', 'KIM', body)
    worksheet283.write('AC5', 'BIO', body)
    worksheet283.write('AD5', 'JML', body)

    worksheet283.conditional_format(5, 0, row283_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet283.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_283}', title)
    worksheet283.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet283.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet283.write('A22', 'LOKASI', header)
    worksheet283.write('B22', 'TOTAL', header)
    worksheet283.merge_range('A21:B21', 'RANK', header)
    worksheet283.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet283.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet283.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet283.merge_range('F21:F22', 'KELAS', header)
    worksheet283.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet283.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet283.write('G22', 'MAW', body)
    worksheet283.write('H22', 'MAP', body)
    worksheet283.write('I22', 'IND', body)
    worksheet283.write('J22', 'ENG', body)
    worksheet283.write('K22', 'SEJ', body)
    worksheet283.write('L22', 'GEO', body)
    worksheet283.write('M22', 'EKO', body)
    worksheet283.write('N22', 'SOS', body)
    worksheet283.write('O22', 'FIS', body)
    worksheet283.write('P22', 'KIM', body)
    worksheet283.write('Q22', 'BIO', body)
    worksheet283.write('R22', 'JML', body)
    worksheet283.write('S22', 'MAW', body)
    worksheet283.write('T22', 'MAP', body)
    worksheet283.write('U22', 'IND', body)
    worksheet283.write('V22', 'ENG', body)
    worksheet283.write('W22', 'SEJ', body)
    worksheet283.write('X22', 'GEO', body)
    worksheet283.write('Y22', 'EKO', body)
    worksheet283.write('Z22', 'SOS', body)
    worksheet283.write('AA22', 'FIS', body)
    worksheet283.write('AB22', 'KIM', body)
    worksheet283.write('AC22', 'BIO', body)
    worksheet283.write('AD22', 'JML', body)

    worksheet283.conditional_format(22, 0, row283+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 284
    worksheet284.insert_image('A1', r'logo resmi nf.jpg')

    worksheet284.set_column('A:A', 7, center)
    worksheet284.set_column('B:B', 6, center)
    worksheet284.set_column('C:C', 18.14, center)
    worksheet284.set_column('D:D', 25, left)
    worksheet284.set_column('E:E', 13.14, left)
    worksheet284.set_column('F:F', 8.57, center)
    worksheet284.set_column('G:AD', 5, center)
    worksheet284.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_284}', title)
    worksheet284.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet284.write('A5', 'LOKASI', header)
    worksheet284.write('B5', 'TOTAL', header)
    worksheet284.merge_range('A4:B4', 'RANK', header)
    worksheet284.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet284.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet284.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet284.merge_range('F4:F5', 'KELAS', header)
    worksheet284.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet284.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet284.write('G5', 'MAW', body)
    worksheet284.write('H5', 'MAP', body)
    worksheet284.write('I5', 'IND', body)
    worksheet284.write('J5', 'ENG', body)
    worksheet284.write('K5', 'SEJ', body)
    worksheet284.write('L5', 'GEO', body)
    worksheet284.write('M5', 'EKO', body)
    worksheet284.write('N5', 'SOS', body)
    worksheet284.write('O5', 'FIS', body)
    worksheet284.write('P5', 'KIM', body)
    worksheet284.write('Q5', 'BIO', body)
    worksheet284.write('R5', 'JML', body)
    worksheet284.write('S5', 'MAW', body)
    worksheet284.write('T5', 'MAP', body)
    worksheet284.write('U5', 'IND', body)
    worksheet284.write('V5', 'ENG', body)
    worksheet284.write('W5', 'SEJ', body)
    worksheet284.write('X5', 'GEO', body)
    worksheet284.write('Y5', 'EKO', body)
    worksheet284.write('Z5', 'SOS', body)
    worksheet284.write('AA5', 'FIS', body)
    worksheet284.write('AB5', 'KIM', body)
    worksheet284.write('AC5', 'BIO', body)
    worksheet284.write('AD5', 'JML', body)

    worksheet284.conditional_format(5, 0, row284_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet284.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_284}', title)
    worksheet284.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet284.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet284.write('A22', 'LOKASI', header)
    worksheet284.write('B22', 'TOTAL', header)
    worksheet284.merge_range('A21:B21', 'RANK', header)
    worksheet284.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet284.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet284.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet284.merge_range('F21:F22', 'KELAS', header)
    worksheet284.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet284.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet284.write('G22', 'MAW', body)
    worksheet284.write('H22', 'MAP', body)
    worksheet284.write('I22', 'IND', body)
    worksheet284.write('J22', 'ENG', body)
    worksheet284.write('K22', 'SEJ', body)
    worksheet284.write('L22', 'GEO', body)
    worksheet284.write('M22', 'EKO', body)
    worksheet284.write('N22', 'SOS', body)
    worksheet284.write('O22', 'FIS', body)
    worksheet284.write('P22', 'KIM', body)
    worksheet284.write('Q22', 'BIO', body)
    worksheet284.write('R22', 'JML', body)
    worksheet284.write('S22', 'MAW', body)
    worksheet284.write('T22', 'MAP', body)
    worksheet284.write('U22', 'IND', body)
    worksheet284.write('V22', 'ENG', body)
    worksheet284.write('W22', 'SEJ', body)
    worksheet284.write('X22', 'GEO', body)
    worksheet284.write('Y22', 'EKO', body)
    worksheet284.write('Z22', 'SOS', body)
    worksheet284.write('AA22', 'FIS', body)
    worksheet284.write('AB22', 'KIM', body)
    worksheet284.write('AC22', 'BIO', body)
    worksheet284.write('AD22', 'JML', body)

    worksheet284.conditional_format(22, 0, row284+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 285
    worksheet285.insert_image('A1', r'logo resmi nf.jpg')

    worksheet285.set_column('A:A', 7, center)
    worksheet285.set_column('B:B', 6, center)
    worksheet285.set_column('C:C', 18.14, center)
    worksheet285.set_column('D:D', 25, left)
    worksheet285.set_column('E:E', 13.14, left)
    worksheet285.set_column('F:F', 8.57, center)
    worksheet285.set_column('G:AD', 5, center)
    worksheet285.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_285}', title)
    worksheet285.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet285.write('A5', 'LOKASI', header)
    worksheet285.write('B5', 'TOTAL', header)
    worksheet285.merge_range('A4:B4', 'RANK', header)
    worksheet285.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet285.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet285.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet285.merge_range('F4:F5', 'KELAS', header)
    worksheet285.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet285.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet285.write('G5', 'MAW', body)
    worksheet285.write('H5', 'MAP', body)
    worksheet285.write('I5', 'IND', body)
    worksheet285.write('J5', 'ENG', body)
    worksheet285.write('K5', 'SEJ', body)
    worksheet285.write('L5', 'GEO', body)
    worksheet285.write('M5', 'EKO', body)
    worksheet285.write('N5', 'SOS', body)
    worksheet285.write('O5', 'FIS', body)
    worksheet285.write('P5', 'KIM', body)
    worksheet285.write('Q5', 'BIO', body)
    worksheet285.write('R5', 'JML', body)
    worksheet285.write('S5', 'MAW', body)
    worksheet285.write('T5', 'MAP', body)
    worksheet285.write('U5', 'IND', body)
    worksheet285.write('V5', 'ENG', body)
    worksheet285.write('W5', 'SEJ', body)
    worksheet285.write('X5', 'GEO', body)
    worksheet285.write('Y5', 'EKO', body)
    worksheet285.write('Z5', 'SOS', body)
    worksheet285.write('AA5', 'FIS', body)
    worksheet285.write('AB5', 'KIM', body)
    worksheet285.write('AC5', 'BIO', body)
    worksheet285.write('AD5', 'JML', body)

    worksheet285.conditional_format(5, 0, row285_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet285.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_285}', title)
    worksheet285.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet285.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet285.write('A22', 'LOKASI', header)
    worksheet285.write('B22', 'TOTAL', header)
    worksheet285.merge_range('A21:B21', 'RANK', header)
    worksheet285.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet285.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet285.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet285.merge_range('F21:F22', 'KELAS', header)
    worksheet285.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet285.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet285.write('G22', 'MAW', body)
    worksheet285.write('H22', 'MAP', body)
    worksheet285.write('I22', 'IND', body)
    worksheet285.write('J22', 'ENG', body)
    worksheet285.write('K22', 'SEJ', body)
    worksheet285.write('L22', 'GEO', body)
    worksheet285.write('M22', 'EKO', body)
    worksheet285.write('N22', 'SOS', body)
    worksheet285.write('O22', 'FIS', body)
    worksheet285.write('P22', 'KIM', body)
    worksheet285.write('Q22', 'BIO', body)
    worksheet285.write('R22', 'JML', body)
    worksheet285.write('S22', 'MAW', body)
    worksheet285.write('T22', 'MAP', body)
    worksheet285.write('U22', 'IND', body)
    worksheet285.write('V22', 'ENG', body)
    worksheet285.write('W22', 'SEJ', body)
    worksheet285.write('X22', 'GEO', body)
    worksheet285.write('Y22', 'EKO', body)
    worksheet285.write('Z22', 'SOS', body)
    worksheet285.write('AA22', 'FIS', body)
    worksheet285.write('AB22', 'KIM', body)
    worksheet285.write('AC22', 'BIO', body)
    worksheet285.write('AD22', 'JML', body)

    worksheet285.conditional_format(22, 0, row285+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 286
    worksheet286.insert_image('A1', r'logo resmi nf.jpg')

    worksheet286.set_column('A:A', 7, center)
    worksheet286.set_column('B:B', 6, center)
    worksheet286.set_column('C:C', 18.14, center)
    worksheet286.set_column('D:D', 25, left)
    worksheet286.set_column('E:E', 13.14, left)
    worksheet286.set_column('F:F', 8.57, center)
    worksheet286.set_column('G:AD', 5, center)
    worksheet286.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_286}', title)
    worksheet286.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet286.write('A5', 'LOKASI', header)
    worksheet286.write('B5', 'TOTAL', header)
    worksheet286.merge_range('A4:B4', 'RANK', header)
    worksheet286.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet286.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet286.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet286.merge_range('F4:F5', 'KELAS', header)
    worksheet286.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet286.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet286.write('G5', 'MAW', body)
    worksheet286.write('H5', 'MAP', body)
    worksheet286.write('I5', 'IND', body)
    worksheet286.write('J5', 'ENG', body)
    worksheet286.write('K5', 'SEJ', body)
    worksheet286.write('L5', 'GEO', body)
    worksheet286.write('M5', 'EKO', body)
    worksheet286.write('N5', 'SOS', body)
    worksheet286.write('O5', 'FIS', body)
    worksheet286.write('P5', 'KIM', body)
    worksheet286.write('Q5', 'BIO', body)
    worksheet286.write('R5', 'JML', body)
    worksheet286.write('S5', 'MAW', body)
    worksheet286.write('T5', 'MAP', body)
    worksheet286.write('U5', 'IND', body)
    worksheet286.write('V5', 'ENG', body)
    worksheet286.write('W5', 'SEJ', body)
    worksheet286.write('X5', 'GEO', body)
    worksheet286.write('Y5', 'EKO', body)
    worksheet286.write('Z5', 'SOS', body)
    worksheet286.write('AA5', 'FIS', body)
    worksheet286.write('AB5', 'KIM', body)
    worksheet286.write('AC5', 'BIO', body)
    worksheet286.write('AD5', 'JML', body)

    worksheet286.conditional_format(5, 0, row286_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet286.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_286}', title)
    worksheet286.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet286.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet286.write('A22', 'LOKASI', header)
    worksheet286.write('B22', 'TOTAL', header)
    worksheet286.merge_range('A21:B21', 'RANK', header)
    worksheet286.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet286.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet286.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet286.merge_range('F21:F22', 'KELAS', header)
    worksheet286.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet286.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet286.write('G22', 'MAW', body)
    worksheet286.write('H22', 'MAP', body)
    worksheet286.write('I22', 'IND', body)
    worksheet286.write('J22', 'ENG', body)
    worksheet286.write('K22', 'SEJ', body)
    worksheet286.write('L22', 'GEO', body)
    worksheet286.write('M22', 'EKO', body)
    worksheet286.write('N22', 'SOS', body)
    worksheet286.write('O22', 'FIS', body)
    worksheet286.write('P22', 'KIM', body)
    worksheet286.write('Q22', 'BIO', body)
    worksheet286.write('R22', 'JML', body)
    worksheet286.write('S22', 'MAW', body)
    worksheet286.write('T22', 'MAP', body)
    worksheet286.write('U22', 'IND', body)
    worksheet286.write('V22', 'ENG', body)
    worksheet286.write('W22', 'SEJ', body)
    worksheet286.write('X22', 'GEO', body)
    worksheet286.write('Y22', 'EKO', body)
    worksheet286.write('Z22', 'SOS', body)
    worksheet286.write('AA22', 'FIS', body)
    worksheet286.write('AB22', 'KIM', body)
    worksheet286.write('AC22', 'BIO', body)
    worksheet286.write('AD22', 'JML', body)

    worksheet286.conditional_format(22, 0, row286+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 287
    worksheet287.insert_image('A1', r'logo resmi nf.jpg')

    worksheet287.set_column('A:A', 7, center)
    worksheet287.set_column('B:B', 6, center)
    worksheet287.set_column('C:C', 18.14, center)
    worksheet287.set_column('D:D', 25, left)
    worksheet287.set_column('E:E', 13.14, left)
    worksheet287.set_column('F:F', 8.57, center)
    worksheet287.set_column('G:AD', 5, center)
    worksheet287.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_287}', title)
    worksheet287.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet287.write('A5', 'LOKASI', header)
    worksheet287.write('B5', 'TOTAL', header)
    worksheet287.merge_range('A4:B4', 'RANK', header)
    worksheet287.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet287.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet287.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet287.merge_range('F4:F5', 'KELAS', header)
    worksheet287.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet287.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet287.write('G5', 'MAW', body)
    worksheet287.write('H5', 'MAP', body)
    worksheet287.write('I5', 'IND', body)
    worksheet287.write('J5', 'ENG', body)
    worksheet287.write('K5', 'SEJ', body)
    worksheet287.write('L5', 'GEO', body)
    worksheet287.write('M5', 'EKO', body)
    worksheet287.write('N5', 'SOS', body)
    worksheet287.write('O5', 'FIS', body)
    worksheet287.write('P5', 'KIM', body)
    worksheet287.write('Q5', 'BIO', body)
    worksheet287.write('R5', 'JML', body)
    worksheet287.write('S5', 'MAW', body)
    worksheet287.write('T5', 'MAP', body)
    worksheet287.write('U5', 'IND', body)
    worksheet287.write('V5', 'ENG', body)
    worksheet287.write('W5', 'SEJ', body)
    worksheet287.write('X5', 'GEO', body)
    worksheet287.write('Y5', 'EKO', body)
    worksheet287.write('Z5', 'SOS', body)
    worksheet287.write('AA5', 'FIS', body)
    worksheet287.write('AB5', 'KIM', body)
    worksheet287.write('AC5', 'BIO', body)
    worksheet287.write('AD5', 'JML', body)

    worksheet287.conditional_format(5, 0, row287_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet287.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_287}', title)
    worksheet287.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet287.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet287.write('A22', 'LOKASI', header)
    worksheet287.write('B22', 'TOTAL', header)
    worksheet287.merge_range('A21:B21', 'RANK', header)
    worksheet287.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet287.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet287.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet287.merge_range('F21:F22', 'KELAS', header)
    worksheet287.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet287.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet287.write('G22', 'MAW', body)
    worksheet287.write('H22', 'MAP', body)
    worksheet287.write('I22', 'IND', body)
    worksheet287.write('J22', 'ENG', body)
    worksheet287.write('K22', 'SEJ', body)
    worksheet287.write('L22', 'GEO', body)
    worksheet287.write('M22', 'EKO', body)
    worksheet287.write('N22', 'SOS', body)
    worksheet287.write('O22', 'FIS', body)
    worksheet287.write('P22', 'KIM', body)
    worksheet287.write('Q22', 'BIO', body)
    worksheet287.write('R22', 'JML', body)
    worksheet287.write('S22', 'MAW', body)
    worksheet287.write('T22', 'MAP', body)
    worksheet287.write('U22', 'IND', body)
    worksheet287.write('V22', 'ENG', body)
    worksheet287.write('W22', 'SEJ', body)
    worksheet287.write('X22', 'GEO', body)
    worksheet287.write('Y22', 'EKO', body)
    worksheet287.write('Z22', 'SOS', body)
    worksheet287.write('AA22', 'FIS', body)
    worksheet287.write('AB22', 'KIM', body)
    worksheet287.write('AC22', 'BIO', body)
    worksheet287.write('AD22', 'JML', body)

    worksheet287.conditional_format(22, 0, row287+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 288
    worksheet288.insert_image('A1', r'logo resmi nf.jpg')

    worksheet288.set_column('A:A', 7, center)
    worksheet288.set_column('B:B', 6, center)
    worksheet288.set_column('C:C', 18.14, center)
    worksheet288.set_column('D:D', 25, left)
    worksheet288.set_column('E:E', 13.14, left)
    worksheet288.set_column('F:F', 8.57, center)
    worksheet288.set_column('G:AD', 5, center)
    worksheet288.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_288}', title)
    worksheet288.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet288.write('A5', 'LOKASI', header)
    worksheet288.write('B5', 'TOTAL', header)
    worksheet288.merge_range('A4:B4', 'RANK', header)
    worksheet288.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet288.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet288.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet288.merge_range('F4:F5', 'KELAS', header)
    worksheet288.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet288.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet288.write('G5', 'MAW', body)
    worksheet288.write('H5', 'MAP', body)
    worksheet288.write('I5', 'IND', body)
    worksheet288.write('J5', 'ENG', body)
    worksheet288.write('K5', 'SEJ', body)
    worksheet288.write('L5', 'GEO', body)
    worksheet288.write('M5', 'EKO', body)
    worksheet288.write('N5', 'SOS', body)
    worksheet288.write('O5', 'FIS', body)
    worksheet288.write('P5', 'KIM', body)
    worksheet288.write('Q5', 'BIO', body)
    worksheet288.write('R5', 'JML', body)
    worksheet288.write('S5', 'MAW', body)
    worksheet288.write('T5', 'MAP', body)
    worksheet288.write('U5', 'IND', body)
    worksheet288.write('V5', 'ENG', body)
    worksheet288.write('W5', 'SEJ', body)
    worksheet288.write('X5', 'GEO', body)
    worksheet288.write('Y5', 'EKO', body)
    worksheet288.write('Z5', 'SOS', body)
    worksheet288.write('AA5', 'FIS', body)
    worksheet288.write('AB5', 'KIM', body)
    worksheet288.write('AC5', 'BIO', body)
    worksheet288.write('AD5', 'JML', body)

    worksheet288.conditional_format(5, 0, row288_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet288.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_288}', title)
    worksheet288.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet288.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet288.write('A22', 'LOKASI', header)
    worksheet288.write('B22', 'TOTAL', header)
    worksheet288.merge_range('A21:B21', 'RANK', header)
    worksheet288.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet288.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet288.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet288.merge_range('F21:F22', 'KELAS', header)
    worksheet288.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet288.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet288.write('G22', 'MAW', body)
    worksheet288.write('H22', 'MAP', body)
    worksheet288.write('I22', 'IND', body)
    worksheet288.write('J22', 'ENG', body)
    worksheet288.write('K22', 'SEJ', body)
    worksheet288.write('L22', 'GEO', body)
    worksheet288.write('M22', 'EKO', body)
    worksheet288.write('N22', 'SOS', body)
    worksheet288.write('O22', 'FIS', body)
    worksheet288.write('P22', 'KIM', body)
    worksheet288.write('Q22', 'BIO', body)
    worksheet288.write('R22', 'JML', body)
    worksheet288.write('S22', 'MAW', body)
    worksheet288.write('T22', 'MAP', body)
    worksheet288.write('U22', 'IND', body)
    worksheet288.write('V22', 'ENG', body)
    worksheet288.write('W22', 'SEJ', body)
    worksheet288.write('X22', 'GEO', body)
    worksheet288.write('Y22', 'EKO', body)
    worksheet288.write('Z22', 'SOS', body)
    worksheet288.write('AA22', 'FIS', body)
    worksheet288.write('AB22', 'KIM', body)
    worksheet288.write('AC22', 'BIO', body)
    worksheet288.write('AD22', 'JML', body)

    worksheet288.conditional_format(22, 0, row288+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 289
    worksheet289.insert_image('A1', r'logo resmi nf.jpg')

    worksheet289.set_column('A:A', 7, center)
    worksheet289.set_column('B:B', 6, center)
    worksheet289.set_column('C:C', 18.14, center)
    worksheet289.set_column('D:D', 25, left)
    worksheet289.set_column('E:E', 13.14, left)
    worksheet289.set_column('F:F', 8.57, center)
    worksheet289.set_column('G:AD', 5, center)
    worksheet289.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_289}', title)
    worksheet289.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet289.write('A5', 'LOKASI', header)
    worksheet289.write('B5', 'TOTAL', header)
    worksheet289.merge_range('A4:B4', 'RANK', header)
    worksheet289.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet289.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet289.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet289.merge_range('F4:F5', 'KELAS', header)
    worksheet289.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet289.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet289.write('G5', 'MAW', body)
    worksheet289.write('H5', 'MAP', body)
    worksheet289.write('I5', 'IND', body)
    worksheet289.write('J5', 'ENG', body)
    worksheet289.write('K5', 'SEJ', body)
    worksheet289.write('L5', 'GEO', body)
    worksheet289.write('M5', 'EKO', body)
    worksheet289.write('N5', 'SOS', body)
    worksheet289.write('O5', 'FIS', body)
    worksheet289.write('P5', 'KIM', body)
    worksheet289.write('Q5', 'BIO', body)
    worksheet289.write('R5', 'JML', body)
    worksheet289.write('S5', 'MAW', body)
    worksheet289.write('T5', 'MAP', body)
    worksheet289.write('U5', 'IND', body)
    worksheet289.write('V5', 'ENG', body)
    worksheet289.write('W5', 'SEJ', body)
    worksheet289.write('X5', 'GEO', body)
    worksheet289.write('Y5', 'EKO', body)
    worksheet289.write('Z5', 'SOS', body)
    worksheet289.write('AA5', 'FIS', body)
    worksheet289.write('AB5', 'KIM', body)
    worksheet289.write('AC5', 'BIO', body)
    worksheet289.write('AD5', 'JML', body)

    worksheet289.conditional_format(5, 0, row289_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet289.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_289}', title)
    worksheet289.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet289.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet289.write('A22', 'LOKASI', header)
    worksheet289.write('B22', 'TOTAL', header)
    worksheet289.merge_range('A21:B21', 'RANK', header)
    worksheet289.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet289.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet289.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet289.merge_range('F21:F22', 'KELAS', header)
    worksheet289.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet289.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet289.write('G22', 'MAW', body)
    worksheet289.write('H22', 'MAP', body)
    worksheet289.write('I22', 'IND', body)
    worksheet289.write('J22', 'ENG', body)
    worksheet289.write('K22', 'SEJ', body)
    worksheet289.write('L22', 'GEO', body)
    worksheet289.write('M22', 'EKO', body)
    worksheet289.write('N22', 'SOS', body)
    worksheet289.write('O22', 'FIS', body)
    worksheet289.write('P22', 'KIM', body)
    worksheet289.write('Q22', 'BIO', body)
    worksheet289.write('R22', 'JML', body)
    worksheet289.write('S22', 'MAW', body)
    worksheet289.write('T22', 'MAP', body)
    worksheet289.write('U22', 'IND', body)
    worksheet289.write('V22', 'ENG', body)
    worksheet289.write('W22', 'SEJ', body)
    worksheet289.write('X22', 'GEO', body)
    worksheet289.write('Y22', 'EKO', body)
    worksheet289.write('Z22', 'SOS', body)
    worksheet289.write('AA22', 'FIS', body)
    worksheet289.write('AB22', 'KIM', body)
    worksheet289.write('AC22', 'BIO', body)
    worksheet289.write('AD22', 'JML', body)

    worksheet289.conditional_format(22, 0, row289+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 290
    worksheet290.insert_image('A1', r'logo resmi nf.jpg')

    worksheet290.set_column('A:A', 7, center)
    worksheet290.set_column('B:B', 6, center)
    worksheet290.set_column('C:C', 18.14, center)
    worksheet290.set_column('D:D', 25, left)
    worksheet290.set_column('E:E', 13.14, left)
    worksheet290.set_column('F:F', 8.57, center)
    worksheet290.set_column('G:AD', 5, center)
    worksheet290.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_290}', title)
    worksheet290.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet290.write('A5', 'LOKASI', header)
    worksheet290.write('B5', 'TOTAL', header)
    worksheet290.merge_range('A4:B4', 'RANK', header)
    worksheet290.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet290.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet290.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet290.merge_range('F4:F5', 'KELAS', header)
    worksheet290.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet290.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet290.write('G5', 'MAW', body)
    worksheet290.write('H5', 'MAP', body)
    worksheet290.write('I5', 'IND', body)
    worksheet290.write('J5', 'ENG', body)
    worksheet290.write('K5', 'SEJ', body)
    worksheet290.write('L5', 'GEO', body)
    worksheet290.write('M5', 'EKO', body)
    worksheet290.write('N5', 'SOS', body)
    worksheet290.write('O5', 'FIS', body)
    worksheet290.write('P5', 'KIM', body)
    worksheet290.write('Q5', 'BIO', body)
    worksheet290.write('R5', 'JML', body)
    worksheet290.write('S5', 'MAW', body)
    worksheet290.write('T5', 'MAP', body)
    worksheet290.write('U5', 'IND', body)
    worksheet290.write('V5', 'ENG', body)
    worksheet290.write('W5', 'SEJ', body)
    worksheet290.write('X5', 'GEO', body)
    worksheet290.write('Y5', 'EKO', body)
    worksheet290.write('Z5', 'SOS', body)
    worksheet290.write('AA5', 'FIS', body)
    worksheet290.write('AB5', 'KIM', body)
    worksheet290.write('AC5', 'BIO', body)
    worksheet290.write('AD5', 'JML', body)

    worksheet290.conditional_format(5, 0, row290_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet290.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_290}', title)
    worksheet290.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet290.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet290.write('A22', 'LOKASI', header)
    worksheet290.write('B22', 'TOTAL', header)
    worksheet290.merge_range('A21:B21', 'RANK', header)
    worksheet290.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet290.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet290.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet290.merge_range('F21:F22', 'KELAS', header)
    worksheet290.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet290.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet290.write('G22', 'MAW', body)
    worksheet290.write('H22', 'MAP', body)
    worksheet290.write('I22', 'IND', body)
    worksheet290.write('J22', 'ENG', body)
    worksheet290.write('K22', 'SEJ', body)
    worksheet290.write('L22', 'GEO', body)
    worksheet290.write('M22', 'EKO', body)
    worksheet290.write('N22', 'SOS', body)
    worksheet290.write('O22', 'FIS', body)
    worksheet290.write('P22', 'KIM', body)
    worksheet290.write('Q22', 'BIO', body)
    worksheet290.write('R22', 'JML', body)
    worksheet290.write('S22', 'MAW', body)
    worksheet290.write('T22', 'MAP', body)
    worksheet290.write('U22', 'IND', body)
    worksheet290.write('V22', 'ENG', body)
    worksheet290.write('W22', 'SEJ', body)
    worksheet290.write('X22', 'GEO', body)
    worksheet290.write('Y22', 'EKO', body)
    worksheet290.write('Z22', 'SOS', body)
    worksheet290.write('AA22', 'FIS', body)
    worksheet290.write('AB22', 'KIM', body)
    worksheet290.write('AC22', 'BIO', body)
    worksheet290.write('AD22', 'JML', body)

    worksheet290.conditional_format(22, 0, row290+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 291
    worksheet291.insert_image('A1', r'logo resmi nf.jpg')

    worksheet291.set_column('A:A', 7, center)
    worksheet291.set_column('B:B', 6, center)
    worksheet291.set_column('C:C', 18.14, center)
    worksheet291.set_column('D:D', 25, left)
    worksheet291.set_column('E:E', 13.14, left)
    worksheet291.set_column('F:F', 8.57, center)
    worksheet291.set_column('G:AD', 5, center)
    worksheet291.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_291}', title)
    worksheet291.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet291.write('A5', 'LOKASI', header)
    worksheet291.write('B5', 'TOTAL', header)
    worksheet291.merge_range('A4:B4', 'RANK', header)
    worksheet291.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet291.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet291.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet291.merge_range('F4:F5', 'KELAS', header)
    worksheet291.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet291.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet291.write('G5', 'MAW', body)
    worksheet291.write('H5', 'MAP', body)
    worksheet291.write('I5', 'IND', body)
    worksheet291.write('J5', 'ENG', body)
    worksheet291.write('K5', 'SEJ', body)
    worksheet291.write('L5', 'GEO', body)
    worksheet291.write('M5', 'EKO', body)
    worksheet291.write('N5', 'SOS', body)
    worksheet291.write('O5', 'FIS', body)
    worksheet291.write('P5', 'KIM', body)
    worksheet291.write('Q5', 'BIO', body)
    worksheet291.write('R5', 'JML', body)
    worksheet291.write('S5', 'MAW', body)
    worksheet291.write('T5', 'MAP', body)
    worksheet291.write('U5', 'IND', body)
    worksheet291.write('V5', 'ENG', body)
    worksheet291.write('W5', 'SEJ', body)
    worksheet291.write('X5', 'GEO', body)
    worksheet291.write('Y5', 'EKO', body)
    worksheet291.write('Z5', 'SOS', body)
    worksheet291.write('AA5', 'FIS', body)
    worksheet291.write('AB5', 'KIM', body)
    worksheet291.write('AC5', 'BIO', body)
    worksheet291.write('AD5', 'JML', body)

    worksheet291.conditional_format(5, 0, row291_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet291.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_291}', title)
    worksheet291.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet291.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet291.write('A22', 'LOKASI', header)
    worksheet291.write('B22', 'TOTAL', header)
    worksheet291.merge_range('A21:B21', 'RANK', header)
    worksheet291.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet291.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet291.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet291.merge_range('F21:F22', 'KELAS', header)
    worksheet291.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet291.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet291.write('G22', 'MAW', body)
    worksheet291.write('H22', 'MAP', body)
    worksheet291.write('I22', 'IND', body)
    worksheet291.write('J22', 'ENG', body)
    worksheet291.write('K22', 'SEJ', body)
    worksheet291.write('L22', 'GEO', body)
    worksheet291.write('M22', 'EKO', body)
    worksheet291.write('N22', 'SOS', body)
    worksheet291.write('O22', 'FIS', body)
    worksheet291.write('P22', 'KIM', body)
    worksheet291.write('Q22', 'BIO', body)
    worksheet291.write('R22', 'JML', body)
    worksheet291.write('S22', 'MAW', body)
    worksheet291.write('T22', 'MAP', body)
    worksheet291.write('U22', 'IND', body)
    worksheet291.write('V22', 'ENG', body)
    worksheet291.write('W22', 'SEJ', body)
    worksheet291.write('X22', 'GEO', body)
    worksheet291.write('Y22', 'EKO', body)
    worksheet291.write('Z22', 'SOS', body)
    worksheet291.write('AA22', 'FIS', body)
    worksheet291.write('AB22', 'KIM', body)
    worksheet291.write('AC22', 'BIO', body)
    worksheet291.write('AD22', 'JML', body)

    worksheet291.conditional_format(22, 0, row291+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 292
    worksheet292.insert_image('A1', r'logo resmi nf.jpg')

    worksheet292.set_column('A:A', 7, center)
    worksheet292.set_column('B:B', 6, center)
    worksheet292.set_column('C:C', 18.14, center)
    worksheet292.set_column('D:D', 25, left)
    worksheet292.set_column('E:E', 13.14, left)
    worksheet292.set_column('F:F', 8.57, center)
    worksheet292.set_column('G:AD', 5, center)
    worksheet292.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_292}', title)
    worksheet292.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet292.write('A5', 'LOKASI', header)
    worksheet292.write('B5', 'TOTAL', header)
    worksheet292.merge_range('A4:B4', 'RANK', header)
    worksheet292.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet292.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet292.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet292.merge_range('F4:F5', 'KELAS', header)
    worksheet292.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet292.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet292.write('G5', 'MAW', body)
    worksheet292.write('H5', 'MAP', body)
    worksheet292.write('I5', 'IND', body)
    worksheet292.write('J5', 'ENG', body)
    worksheet292.write('K5', 'SEJ', body)
    worksheet292.write('L5', 'GEO', body)
    worksheet292.write('M5', 'EKO', body)
    worksheet292.write('N5', 'SOS', body)
    worksheet292.write('O5', 'FIS', body)
    worksheet292.write('P5', 'KIM', body)
    worksheet292.write('Q5', 'BIO', body)
    worksheet292.write('R5', 'JML', body)
    worksheet292.write('S5', 'MAW', body)
    worksheet292.write('T5', 'MAP', body)
    worksheet292.write('U5', 'IND', body)
    worksheet292.write('V5', 'ENG', body)
    worksheet292.write('W5', 'SEJ', body)
    worksheet292.write('X5', 'GEO', body)
    worksheet292.write('Y5', 'EKO', body)
    worksheet292.write('Z5', 'SOS', body)
    worksheet292.write('AA5', 'FIS', body)
    worksheet292.write('AB5', 'KIM', body)
    worksheet292.write('AC5', 'BIO', body)
    worksheet292.write('AD5', 'JML', body)

    worksheet292.conditional_format(5, 0, row292_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet292.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_292}', title)
    worksheet292.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet292.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet292.write('A22', 'LOKASI', header)
    worksheet292.write('B22', 'TOTAL', header)
    worksheet292.merge_range('A21:B21', 'RANK', header)
    worksheet292.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet292.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet292.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet292.merge_range('F21:F22', 'KELAS', header)
    worksheet292.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet292.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet292.write('G22', 'MAW', body)
    worksheet292.write('H22', 'MAP', body)
    worksheet292.write('I22', 'IND', body)
    worksheet292.write('J22', 'ENG', body)
    worksheet292.write('K22', 'SEJ', body)
    worksheet292.write('L22', 'GEO', body)
    worksheet292.write('M22', 'EKO', body)
    worksheet292.write('N22', 'SOS', body)
    worksheet292.write('O22', 'FIS', body)
    worksheet292.write('P22', 'KIM', body)
    worksheet292.write('Q22', 'BIO', body)
    worksheet292.write('R22', 'JML', body)
    worksheet292.write('S22', 'MAW', body)
    worksheet292.write('T22', 'MAP', body)
    worksheet292.write('U22', 'IND', body)
    worksheet292.write('V22', 'ENG', body)
    worksheet292.write('W22', 'SEJ', body)
    worksheet292.write('X22', 'GEO', body)
    worksheet292.write('Y22', 'EKO', body)
    worksheet292.write('Z22', 'SOS', body)
    worksheet292.write('AA22', 'FIS', body)
    worksheet292.write('AB22', 'KIM', body)
    worksheet292.write('AC22', 'BIO', body)
    worksheet292.write('AD22', 'JML', body)

    worksheet292.conditional_format(22, 0, row292+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 293
    worksheet293.insert_image('A1', r'logo resmi nf.jpg')

    worksheet293.set_column('A:A', 7, center)
    worksheet293.set_column('B:B', 6, center)
    worksheet293.set_column('C:C', 18.14, center)
    worksheet293.set_column('D:D', 25, left)
    worksheet293.set_column('E:E', 13.14, left)
    worksheet293.set_column('F:F', 8.57, center)
    worksheet293.set_column('G:AD', 5, center)
    worksheet293.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_293}', title)
    worksheet293.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet293.write('A5', 'LOKASI', header)
    worksheet293.write('B5', 'TOTAL', header)
    worksheet293.merge_range('A4:B4', 'RANK', header)
    worksheet293.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet293.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet293.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet293.merge_range('F4:F5', 'KELAS', header)
    worksheet293.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet293.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet293.write('G5', 'MAW', body)
    worksheet293.write('H5', 'MAP', body)
    worksheet293.write('I5', 'IND', body)
    worksheet293.write('J5', 'ENG', body)
    worksheet293.write('K5', 'SEJ', body)
    worksheet293.write('L5', 'GEO', body)
    worksheet293.write('M5', 'EKO', body)
    worksheet293.write('N5', 'SOS', body)
    worksheet293.write('O5', 'FIS', body)
    worksheet293.write('P5', 'KIM', body)
    worksheet293.write('Q5', 'BIO', body)
    worksheet293.write('R5', 'JML', body)
    worksheet293.write('S5', 'MAW', body)
    worksheet293.write('T5', 'MAP', body)
    worksheet293.write('U5', 'IND', body)
    worksheet293.write('V5', 'ENG', body)
    worksheet293.write('W5', 'SEJ', body)
    worksheet293.write('X5', 'GEO', body)
    worksheet293.write('Y5', 'EKO', body)
    worksheet293.write('Z5', 'SOS', body)
    worksheet293.write('AA5', 'FIS', body)
    worksheet293.write('AB5', 'KIM', body)
    worksheet293.write('AC5', 'BIO', body)
    worksheet293.write('AD5', 'JML', body)

    worksheet293.conditional_format(5, 0, row293_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet293.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_293}', title)
    worksheet293.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet293.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet293.write('A22', 'LOKASI', header)
    worksheet293.write('B22', 'TOTAL', header)
    worksheet293.merge_range('A21:B21', 'RANK', header)
    worksheet293.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet293.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet293.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet293.merge_range('F21:F22', 'KELAS', header)
    worksheet293.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet293.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet293.write('G22', 'MAW', body)
    worksheet293.write('H22', 'MAP', body)
    worksheet293.write('I22', 'IND', body)
    worksheet293.write('J22', 'ENG', body)
    worksheet293.write('K22', 'SEJ', body)
    worksheet293.write('L22', 'GEO', body)
    worksheet293.write('M22', 'EKO', body)
    worksheet293.write('N22', 'SOS', body)
    worksheet293.write('O22', 'FIS', body)
    worksheet293.write('P22', 'KIM', body)
    worksheet293.write('Q22', 'BIO', body)
    worksheet293.write('R22', 'JML', body)
    worksheet293.write('S22', 'MAW', body)
    worksheet293.write('T22', 'MAP', body)
    worksheet293.write('U22', 'IND', body)
    worksheet293.write('V22', 'ENG', body)
    worksheet293.write('W22', 'SEJ', body)
    worksheet293.write('X22', 'GEO', body)
    worksheet293.write('Y22', 'EKO', body)
    worksheet293.write('Z22', 'SOS', body)
    worksheet293.write('AA22', 'FIS', body)
    worksheet293.write('AB22', 'KIM', body)
    worksheet293.write('AC22', 'BIO', body)
    worksheet293.write('AD22', 'JML', body)

    worksheet293.conditional_format(22, 0, row293+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 294
    worksheet294.insert_image('A1', r'logo resmi nf.jpg')

    worksheet294.set_column('A:A', 7, center)
    worksheet294.set_column('B:B', 6, center)
    worksheet294.set_column('C:C', 18.14, center)
    worksheet294.set_column('D:D', 25, left)
    worksheet294.set_column('E:E', 13.14, left)
    worksheet294.set_column('F:F', 8.57, center)
    worksheet294.set_column('G:AD', 5, center)
    worksheet294.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_294}', title)
    worksheet294.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet294.write('A5', 'LOKASI', header)
    worksheet294.write('B5', 'TOTAL', header)
    worksheet294.merge_range('A4:B4', 'RANK', header)
    worksheet294.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet294.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet294.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet294.merge_range('F4:F5', 'KELAS', header)
    worksheet294.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet294.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet294.write('G5', 'MAW', body)
    worksheet294.write('H5', 'MAP', body)
    worksheet294.write('I5', 'IND', body)
    worksheet294.write('J5', 'ENG', body)
    worksheet294.write('K5', 'SEJ', body)
    worksheet294.write('L5', 'GEO', body)
    worksheet294.write('M5', 'EKO', body)
    worksheet294.write('N5', 'SOS', body)
    worksheet294.write('O5', 'FIS', body)
    worksheet294.write('P5', 'KIM', body)
    worksheet294.write('Q5', 'BIO', body)
    worksheet294.write('R5', 'JML', body)
    worksheet294.write('S5', 'MAW', body)
    worksheet294.write('T5', 'MAP', body)
    worksheet294.write('U5', 'IND', body)
    worksheet294.write('V5', 'ENG', body)
    worksheet294.write('W5', 'SEJ', body)
    worksheet294.write('X5', 'GEO', body)
    worksheet294.write('Y5', 'EKO', body)
    worksheet294.write('Z5', 'SOS', body)
    worksheet294.write('AA5', 'FIS', body)
    worksheet294.write('AB5', 'KIM', body)
    worksheet294.write('AC5', 'BIO', body)
    worksheet294.write('AD5', 'JML', body)

    worksheet294.conditional_format(5, 0, row294_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet294.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_294}', title)
    worksheet294.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet294.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet294.write('A22', 'LOKASI', header)
    worksheet294.write('B22', 'TOTAL', header)
    worksheet294.merge_range('A21:B21', 'RANK', header)
    worksheet294.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet294.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet294.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet294.merge_range('F21:F22', 'KELAS', header)
    worksheet294.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet294.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet294.write('G22', 'MAW', body)
    worksheet294.write('H22', 'MAP', body)
    worksheet294.write('I22', 'IND', body)
    worksheet294.write('J22', 'ENG', body)
    worksheet294.write('K22', 'SEJ', body)
    worksheet294.write('L22', 'GEO', body)
    worksheet294.write('M22', 'EKO', body)
    worksheet294.write('N22', 'SOS', body)
    worksheet294.write('O22', 'FIS', body)
    worksheet294.write('P22', 'KIM', body)
    worksheet294.write('Q22', 'BIO', body)
    worksheet294.write('R22', 'JML', body)
    worksheet294.write('S22', 'MAW', body)
    worksheet294.write('T22', 'MAP', body)
    worksheet294.write('U22', 'IND', body)
    worksheet294.write('V22', 'ENG', body)
    worksheet294.write('W22', 'SEJ', body)
    worksheet294.write('X22', 'GEO', body)
    worksheet294.write('Y22', 'EKO', body)
    worksheet294.write('Z22', 'SOS', body)
    worksheet294.write('AA22', 'FIS', body)
    worksheet294.write('AB22', 'KIM', body)
    worksheet294.write('AC22', 'BIO', body)
    worksheet294.write('AD22', 'JML', body)

    worksheet294.conditional_format(22, 0, row294+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 295
    worksheet295.insert_image('A1', r'logo resmi nf.jpg')

    worksheet295.set_column('A:A', 7, center)
    worksheet295.set_column('B:B', 6, center)
    worksheet295.set_column('C:C', 18.14, center)
    worksheet295.set_column('D:D', 25, left)
    worksheet295.set_column('E:E', 13.14, left)
    worksheet295.set_column('F:F', 8.57, center)
    worksheet295.set_column('G:AD', 5, center)
    worksheet295.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_295}', title)
    worksheet295.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet295.write('A5', 'LOKASI', header)
    worksheet295.write('B5', 'TOTAL', header)
    worksheet295.merge_range('A4:B4', 'RANK', header)
    worksheet295.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet295.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet295.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet295.merge_range('F4:F5', 'KELAS', header)
    worksheet295.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet295.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet295.write('G5', 'MAW', body)
    worksheet295.write('H5', 'MAP', body)
    worksheet295.write('I5', 'IND', body)
    worksheet295.write('J5', 'ENG', body)
    worksheet295.write('K5', 'SEJ', body)
    worksheet295.write('L5', 'GEO', body)
    worksheet295.write('M5', 'EKO', body)
    worksheet295.write('N5', 'SOS', body)
    worksheet295.write('O5', 'FIS', body)
    worksheet295.write('P5', 'KIM', body)
    worksheet295.write('Q5', 'BIO', body)
    worksheet295.write('R5', 'JML', body)
    worksheet295.write('S5', 'MAW', body)
    worksheet295.write('T5', 'MAP', body)
    worksheet295.write('U5', 'IND', body)
    worksheet295.write('V5', 'ENG', body)
    worksheet295.write('W5', 'SEJ', body)
    worksheet295.write('X5', 'GEO', body)
    worksheet295.write('Y5', 'EKO', body)
    worksheet295.write('Z5', 'SOS', body)
    worksheet295.write('AA5', 'FIS', body)
    worksheet295.write('AB5', 'KIM', body)
    worksheet295.write('AC5', 'BIO', body)
    worksheet295.write('AD5', 'JML', body)

    worksheet295.conditional_format(5, 0, row295_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet295.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_295}', title)
    worksheet295.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet295.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet295.write('A22', 'LOKASI', header)
    worksheet295.write('B22', 'TOTAL', header)
    worksheet295.merge_range('A21:B21', 'RANK', header)
    worksheet295.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet295.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet295.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet295.merge_range('F21:F22', 'KELAS', header)
    worksheet295.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet295.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet295.write('G22', 'MAW', body)
    worksheet295.write('H22', 'MAP', body)
    worksheet295.write('I22', 'IND', body)
    worksheet295.write('J22', 'ENG', body)
    worksheet295.write('K22', 'SEJ', body)
    worksheet295.write('L22', 'GEO', body)
    worksheet295.write('M22', 'EKO', body)
    worksheet295.write('N22', 'SOS', body)
    worksheet295.write('O22', 'FIS', body)
    worksheet295.write('P22', 'KIM', body)
    worksheet295.write('Q22', 'BIO', body)
    worksheet295.write('R22', 'JML', body)
    worksheet295.write('S22', 'MAW', body)
    worksheet295.write('T22', 'MAP', body)
    worksheet295.write('U22', 'IND', body)
    worksheet295.write('V22', 'ENG', body)
    worksheet295.write('W22', 'SEJ', body)
    worksheet295.write('X22', 'GEO', body)
    worksheet295.write('Y22', 'EKO', body)
    worksheet295.write('Z22', 'SOS', body)
    worksheet295.write('AA22', 'FIS', body)
    worksheet295.write('AB22', 'KIM', body)
    worksheet295.write('AC22', 'BIO', body)
    worksheet295.write('AD22', 'JML', body)

    worksheet295.conditional_format(22, 0, row295+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 298
    worksheet298.insert_image('A1', r'logo resmi nf.jpg')

    worksheet298.set_column('A:A', 7, center)
    worksheet298.set_column('B:B', 6, center)
    worksheet298.set_column('C:C', 18.14, center)
    worksheet298.set_column('D:D', 25, left)
    worksheet298.set_column('E:E', 13.14, left)
    worksheet298.set_column('F:F', 8.57, center)
    worksheet298.set_column('G:AD', 5, center)
    worksheet298.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_298}', title)
    worksheet298.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet298.write('A5', 'LOKASI', header)
    worksheet298.write('B5', 'TOTAL', header)
    worksheet298.merge_range('A4:B4', 'RANK', header)
    worksheet298.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet298.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet298.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet298.merge_range('F4:F5', 'KELAS', header)
    worksheet298.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet298.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet298.write('G5', 'MAW', body)
    worksheet298.write('H5', 'MAP', body)
    worksheet298.write('I5', 'IND', body)
    worksheet298.write('J5', 'ENG', body)
    worksheet298.write('K5', 'SEJ', body)
    worksheet298.write('L5', 'GEO', body)
    worksheet298.write('M5', 'EKO', body)
    worksheet298.write('N5', 'SOS', body)
    worksheet298.write('O5', 'FIS', body)
    worksheet298.write('P5', 'KIM', body)
    worksheet298.write('Q5', 'BIO', body)
    worksheet298.write('R5', 'JML', body)
    worksheet298.write('S5', 'MAW', body)
    worksheet298.write('T5', 'MAP', body)
    worksheet298.write('U5', 'IND', body)
    worksheet298.write('V5', 'ENG', body)
    worksheet298.write('W5', 'SEJ', body)
    worksheet298.write('X5', 'GEO', body)
    worksheet298.write('Y5', 'EKO', body)
    worksheet298.write('Z5', 'SOS', body)
    worksheet298.write('AA5', 'FIS', body)
    worksheet298.write('AB5', 'KIM', body)
    worksheet298.write('AC5', 'BIO', body)
    worksheet298.write('AD5', 'JML', body)

    worksheet298.conditional_format(5, 0, row298_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet298.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_298}', title)
    worksheet298.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet298.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet298.write('A22', 'LOKASI', header)
    worksheet298.write('B22', 'TOTAL', header)
    worksheet298.merge_range('A21:B21', 'RANK', header)
    worksheet298.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet298.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet298.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet298.merge_range('F21:F22', 'KELAS', header)
    worksheet298.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet298.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet298.write('G22', 'MAW', body)
    worksheet298.write('H22', 'MAP', body)
    worksheet298.write('I22', 'IND', body)
    worksheet298.write('J22', 'ENG', body)
    worksheet298.write('K22', 'SEJ', body)
    worksheet298.write('L22', 'GEO', body)
    worksheet298.write('M22', 'EKO', body)
    worksheet298.write('N22', 'SOS', body)
    worksheet298.write('O22', 'FIS', body)
    worksheet298.write('P22', 'KIM', body)
    worksheet298.write('Q22', 'BIO', body)
    worksheet298.write('R22', 'JML', body)
    worksheet298.write('S22', 'MAW', body)
    worksheet298.write('T22', 'MAP', body)
    worksheet298.write('U22', 'IND', body)
    worksheet298.write('V22', 'ENG', body)
    worksheet298.write('W22', 'SEJ', body)
    worksheet298.write('X22', 'GEO', body)
    worksheet298.write('Y22', 'EKO', body)
    worksheet298.write('Z22', 'SOS', body)
    worksheet298.write('AA22', 'FIS', body)
    worksheet298.write('AB22', 'KIM', body)
    worksheet298.write('AC22', 'BIO', body)
    worksheet298.write('AD22', 'JML', body)

    worksheet298.conditional_format(22, 0, row298+21, 29,
                                    {'type': 'no_errors', 'format': border})

    # worksheet 299
    worksheet299.insert_image('A1', r'logo resmi nf.jpg')

    worksheet299.set_column('A:A', 7, center)
    worksheet299.set_column('B:B', 6, center)
    worksheet299.set_column('C:C', 18.14, center)
    worksheet299.set_column('D:D', 25, left)
    worksheet299.set_column('E:E', 13.14, left)
    worksheet299.set_column('F:F', 8.57, center)
    worksheet299.set_column('G:AD', 5, center)
    worksheet299.merge_range(
        'A1:AD1', fr'10 SISWA KELAS {kelas} PERINGKAT TERTINGGI NF {lok_299}', title)
    worksheet299.merge_range(
        'A2:AD2', fr'{penilaian} - {semester} TAHUN {tahun}', sub_title)
    worksheet299.write('A5', 'LOKASI', header)
    worksheet299.write('B5', 'TOTAL', header)
    worksheet299.merge_range('A4:B4', 'RANK', header)
    worksheet299.merge_range('C4:C5', 'NOMOR NF', header)
    worksheet299.merge_range('D4:D5', 'NAMA SISWA', header)
    worksheet299.merge_range('E4:E5', 'SEKOLAH', header)
    worksheet299.merge_range('F4:F5', 'KELAS', header)
    worksheet299.merge_range('G4:R4', 'JUMLAH BENAR', header)
    worksheet299.merge_range('S4:AD4', 'NILAI STANDAR', header)
    worksheet299.write('G5', 'MAW', body)
    worksheet299.write('H5', 'MAP', body)
    worksheet299.write('I5', 'IND', body)
    worksheet299.write('J5', 'ENG', body)
    worksheet299.write('K5', 'SEJ', body)
    worksheet299.write('L5', 'GEO', body)
    worksheet299.write('M5', 'EKO', body)
    worksheet299.write('N5', 'SOS', body)
    worksheet299.write('O5', 'FIS', body)
    worksheet299.write('P5', 'KIM', body)
    worksheet299.write('Q5', 'BIO', body)
    worksheet299.write('R5', 'JML', body)
    worksheet299.write('S5', 'MAW', body)
    worksheet299.write('T5', 'MAP', body)
    worksheet299.write('U5', 'IND', body)
    worksheet299.write('V5', 'ENG', body)
    worksheet299.write('W5', 'SEJ', body)
    worksheet299.write('X5', 'GEO', body)
    worksheet299.write('Y5', 'EKO', body)
    worksheet299.write('Z5', 'SOS', body)
    worksheet299.write('AA5', 'FIS', body)
    worksheet299.write('AB5', 'KIM', body)
    worksheet299.write('AC5', 'BIO', body)
    worksheet299.write('AD5', 'JML', body)

    worksheet299.conditional_format(5, 0, row299_10+4, 29,
                                    {'type': 'no_errors', 'format': border})

    worksheet299.merge_range(
        'A17:AD17', fr'KELAS {kelas} - LOKASI NF {lok_299}', title)
    worksheet299.merge_range('A18:AD18', fr'{penilaian}', subTitle)
    worksheet299.merge_range(
        'A19:AD19', fr'{semester} TAHUN {tahun}', sub_title)
    worksheet299.write('A22', 'LOKASI', header)
    worksheet299.write('B22', 'TOTAL', header)
    worksheet299.merge_range('A21:B21', 'RANK', header)
    worksheet299.merge_range('C21:C22', 'NOMOR NF', header)
    worksheet299.merge_range('D21:D22', 'NAMA SISWA', header)
    worksheet299.merge_range('E21:E22', 'SEKOLAH', header)
    worksheet299.merge_range('F21:F22', 'KELAS', header)
    worksheet299.merge_range('G21:R21', 'JUMLAH BENAR', header)
    worksheet299.merge_range('S21:AD21', 'NILAI STANDAR', header)
    worksheet299.write('G22', 'MAW', body)
    worksheet299.write('H22', 'MAP', body)
    worksheet299.write('I22', 'IND', body)
    worksheet299.write('J22', 'ENG', body)
    worksheet299.write('K22', 'SEJ', body)
    worksheet299.write('L22', 'GEO', body)
    worksheet299.write('M22', 'EKO', body)
    worksheet299.write('N22', 'SOS', body)
    worksheet299.write('O22', 'FIS', body)
    worksheet299.write('P22', 'KIM', body)
    worksheet299.write('Q22', 'BIO', body)
    worksheet299.write('R22', 'JML', body)
    worksheet299.write('S22', 'MAW', body)
    worksheet299.write('T22', 'MAP', body)
    worksheet299.write('U22', 'IND', body)
    worksheet299.write('V22', 'ENG', body)
    worksheet299.write('W22', 'SEJ', body)
    worksheet299.write('X22', 'GEO', body)
    worksheet299.write('Y22', 'EKO', body)
    worksheet299.write('Z22', 'SOS', body)
    worksheet299.write('AA22', 'FIS', body)
    worksheet299.write('AB22', 'KIM', body)
    worksheet299.write('AC22', 'BIO', body)
    worksheet299.write('AD22', 'JML', body)

    worksheet299.conditional_format(22, 0, row299+21, 29,
                                    {'type': 'no_errors', 'format': border})

    

    workbook.close()
    st.success("File siap diunduh!")

    # Tombol unduh file
    with open(file_path, "rb") as f:
        bytes_data = f.read()
    st.download_button(label="Unduh File", data=bytes_data,
                       file_name=file_name)
