import pickle
import streamlit as st
import streamlit_authenticator as stauth
from pathlib import Path
# User Authentication
names = ["TI Polda NF", "TI Polda NF"]
usernames = ["admin1", "admin2"]

# load hashed kd_akses
file_path = Path(__file__).parent / "hashed_pw.pkl"
with file_path.open("rb") as file:
    hashed_kd_akses = pickle.load(file)

authenticator = stauth.Authenticate(
    names, usernames, hashed_kd_akses, "lookup", "abcdef")
name, authentication_status, username = authenticator.login("Login", "main")

if authentication_status == False:
    st.error("Username/kode akses salah")

if authentication_status == None:
    st.warning("Silahkan masukan username dan kode akses")

if authentication_status:
    # kurikulum - kelas - mapel
    # 4sd k13
    k13_4sd_mat = 'LHG94EEQ'
    k13_4sd_ind = 'LHG9KCRA'
    k13_4sd_eng = 'LHGA44Y9'
    k13_4sd_ipa = 'LHGALT9N'
    k13_4sd_ips = 'LHH0F32F'
    k13_4sd = [k13_4sd_mat, k13_4sd_ind, k13_4sd_eng, k13_4sd_ipa, k13_4sd_ips]
    column_order_k13_4sd = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_4SD', 'IND_4SD',
                            'ENG_4SD', 'IPA_4SD', 'IPS_4SD']

    # 4sd km
    km_4sd_mat = 'LHH0U12P'
    km_4sd_ind = 'LHH19TQN'
    km_4sd_eng = 'LHH47YLV'
    km_4sd_ipas = 'LHH4U3Q0'
    km_4sd = [km_4sd_mat, km_4sd_ind, km_4sd_eng, km_4sd_ipas]
    column_order_km_4sd = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_4KM', 'IND_4KM',
                           'ENG_4KM', 'IPAS_4KM']

    # 5sd k13
    k13_5sd_mat = 'LHH5V62M'
    k13_5sd_ind = 'LHH6WL2C'
    k13_5sd_eng = 'LHH7NAB5'
    k13_5sd_ipa = 'LHHCO0Q4'
    k13_5sd_ips = 'LHHDAY7I'
    k13_5sd = [k13_5sd_mat, k13_5sd_ind, k13_5sd_eng, k13_5sd_ipa, k13_5sd_ips]
    column_order_k13_5sd = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_5SD', 'IND_5SD',
                            'ENG_5SD', 'IPA_5SD', 'IPS_5SD']

    # 7smp k13
    k13_7smp_mat = 'LHHDRBXZ'
    k13_7smp_ind = 'LHHDUWKS'
    k13_7smp_eng = 'LHHDX6U7'
    k13_7smp_ipa = 'LHHDZC8Y'
    k13_7smp_ips = 'LHHE476J'
    k13_7smp = [k13_7smp_mat, k13_7smp_ind,
                k13_7smp_eng, k13_7smp_ipa, k13_7smp_ips]
    column_order_k13_7smp = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_7SMP', 'IND_7SMP',
                             'ENG_7SMP', 'IPA_7SMP', 'IPS_7SMP']

    # 8smp k13
    k13_8smp_mat = 'LHH6H3F6'
    k13_8smp_ind = 'LHH6TEO5'
    k13_8smp_eng = 'LHHN9AZH'
    k13_8smp_ipa = 'LHHNDOAI'
    k13_8smp_ips = 'LHHNFJ3E'
    k13_8smp = [k13_8smp_mat, k13_8smp_ind,
                k13_8smp_eng, k13_8smp_ipa, k13_8smp_ips]
    column_order_k13_8smp = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_8SMP', 'IND_8SMP',
                             'ENG_8SMP', 'IPA_8SMP', 'IPS_8SMP']

    # 7smp km
    km_7smp_mat = 'LHHE7GC8'
    km_7smp_ind = 'LHHEAQWK'
    km_7smp_eng = 'LHHEEEB5'
    km_7smp_ipa = 'LHHF9Q62'
    km_7smp_ips = 'LHHFBCWT'
    km_7smp = [km_7smp_mat, km_7smp_ind, km_7smp_eng, km_7smp_ipa, km_7smp_ips]
    column_order_km_7smp = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_7KM', 'IND_7KM',
                            'ENG_7KM', 'IPA_7KM', 'IPS_7KM']

    # 10sma ipa k13
    k13_10ipa_mat = 'LHHO4J0W'
    k13_10ipa_bio = 'LHHO78FV'
    k13_10ipa_fis = 'LHHOB3L0'
    k13_10ipa_kim = 'LHHODJIH'
    k13_10ipa = [k13_10ipa_mat, k13_10ipa_bio, k13_10ipa_fis, k13_10ipa_kim]
    column_order_k13_10ipa = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_10IPA', 'FIS_10IPA',
                              'KIM_10IPA', 'BIO_10IPA']

    # 10sma ips k13
    k13_10ips_mat = 'LHHOHQW3'
    k13_10ips_sos = 'LHHOL2GH'
    k13_10ips_eng = 'LHHOOPEJ'
    k13_10ips_eko = 'LHHOR6Q5'
    k13_10ips_ind = 'LHHOUB5D'
    k13_10ips_sej = 'LHHOXG3D'
    k13_10ips_geo = 'LHHP0FDK'
    k13_10ips = [k13_10ips_mat, k13_10ips_sos, k13_10ips_eng,
                 k13_10ips_eko, k13_10ips_ind, k13_10ips_sej, k13_10ips_geo]
    column_order_k13_10ips = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_10IPS', 'IND_10IPS',
                              'ENG_10IPS', 'SEJ_10IPS', 'GEO_10IPS', 'EKO_10IPS', 'SOS_10IPS']

    # 10sma km
    km_10sma_mat = 'LHHP5ZFU'
    km_10sma_ind = 'LHHP8HQU'
    km_10sma_eng = 'LHHPD245'
    km_10sma_ipa = 'LHHPGDX7'
    km_10sma_ips = 'LHHPI776'
    km_10sma = [km_10sma_mat, km_10sma_ind,
                km_10sma_eng, km_10sma_ipa, km_10sma_ips]
    column_order_km_10sma = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_10KM', 'IND_10KM',
                             'ENG_10KM', 'IPA_10KM', 'IPS_10KM']

    # 11sma ipa k13
    k13_11ipa_mat = 'LHHPMD5N'
    k13_11ipa_bio = 'LHHPPGY1'
    k13_11ipa_fis = 'LHHPSAH8'
    k13_11ipa_kim = 'LHHPVWEU'
    k13_11ipa = [k13_11ipa_mat, k13_11ipa_bio,
                 k13_11ipa_fis, k13_11ipa_kim]
    column_order_k13_11ipa = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_11IPA', 'FIS_11IPA',
                              'KIM_11IPA', 'BIO_11IPA']

    # 11sma ips k13
    k13_11ips_mat = 'LHHPY4DQ'
    k13_11ips_sos = 'LHHQ0GEG'
    k13_11ips_eng = 'LHHQ2SOF'
    k13_11ips_eko = 'LHHQ5KA7'
    k13_11ips_ind = 'LHHQ9G2X'
    k13_11ips_sej = 'LHHQBZF1'
    k13_11ips_geo = 'LHHQEL75'
    k13_11ips = [k13_11ips_mat, k13_11ips_sos, k13_11ips_eng,
                 k13_11ips_eko, k13_11ips_ind, k13_11ips_sej, k13_11ips_geo]
    column_order_k13_11ips = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAT_11IPS', 'IND_11IPS',
                              'ENG_11IPS', 'SEJ_11IPS', 'GEO_11IPS', 'EKO_11IPS', 'SOS_11IPS']

    # 11sma km
    km_11sma_maw = 'LHHQHX38'
    km_11sma_map = 'LHPP32X4'
    km_11sma_ind = 'LHHR041C'
    km_11sma_eng = 'LHHQTDZN'
    km_11sma_sej = 'LHHR2936'
    km_11sma_geo = 'LHHR7D5N'
    km_11sma_eko = 'LHHQXWPS'
    km_11sma_sos = 'LHHQNOEV'
    km_11sma_fis = 'LHHQVJ2O'
    km_11sma_kim = 'LHHR50DU'
    km_11sma_bio = 'LHHQQ7WE'
    km_11sma = [km_11sma_maw, km_11sma_map, km_11sma_ind, km_11sma_eng,
                km_11sma_sej, km_11sma_geo, km_11sma_eko, km_11sma_sos,
                km_11sma_fis, km_11sma_kim, km_11sma_bio]
    column_order_km_11sma = ['IDTAHUN', 'NAMA', 'NONF', 'KELAS', 'NAMA_SKLH', 'KD_LOK', 'MAW_11KM', 'MAP_11KM', 'IND_11KM',
                             'ENG_11KM', 'SEJ_11KM', 'GEO_11KM', 'EKO_11KM', 'SOS_11KM', 'FIS_11KM', 'KIM_11KM', 'BIO_11KM']

    image = Image.open('logo resmi nf resize.png')
    st.image(image)

    st.title("PIVOT - PAT")

    col1 = st.container()
    with col1:
        KURIKULUM = st.selectbox(
            "KURIKULUM",
            ("--Pilih Kurikulum--", "K13", "KM"))

    col2 = st.container()
    with col2:
        KELAS = st.selectbox(
            "KELAS",
            ("--Pilih Kelas--", "4 SD", "5 SD", "7 SMP", "8 SMP", "10 IPA", "10 IPS", "10 SMA", "11 IPA", "11 IPS", "11 SMA"))

    TAHUN = st.text_input("Masukkan Tahun Ajaran",
                          placeholder="contoh: 2022-2023")

    uploaded_bobot = st.file_uploader(
        'Letakkan file excel bobot TO', type='xlsx')
    uploaded_jwb = st.file_uploader('Letakkan file excel jwb TO', type='xlsx')

    bobot = None
    jwb = None

    if uploaded_bobot is not None:
        bobot = pd.read_excel(uploaded_bobot)

    if uploaded_jwb is not None:
        jwb = pd.read_excel(uploaded_jwb)

    if bobot is not None and jwb is not None:
        bobot = bobot.drop(['id', 'jns_pkt', 'jns_tes', 'kel_studi', 'nama_tes', 'no_soal', 'bobot', 'kd_studi', 'bab', 'eigen', 'kode_soal', 'st_eigen',
                            'modified_time', 'kode_naskah', 'group_tes', 'kunci', 'sequence', 'label', 'item_id'], axis=1)  # Menghilangkan kolom sebelum dilakukan merge

        result = pd.merge(bobot, jwb[['kode', 'nama', 'nonf', 'kd_lok',
                                      'nama_sklh', 'kelas', 'jml_benar']], on='kode', how='left')
        # Menghapus nilai NaN dari kolom 'nonf'
        result = result.dropna(subset=['nonf'])

        # k13
        if KELAS == "4 SD" and KURIKULUM == "K13":
            kode_kls_kur = k13_4sd
            column_order = column_order_k13_4sd
        elif KELAS == "5 SD" and KURIKULUM == "K13":
            kode_kls_kur = k13_5sd
            column_order = column_order_k13_5sd
        elif KELAS == "7 SMP" and KURIKULUM == "K13":
            kode_kls_kur = k13_7smp
            column_order = column_order_k13_7smp
        elif KELAS == "8 SMP" and KURIKULUM == "K13":
            kode_kls_kur = k13_8smp
            column_order = column_order_k13_8smp
        elif KELAS == "10 IPA" and KURIKULUM == "K13":
            kode_kls_kur = k13_10ipa
            column_order = column_order_k13_10ipa
        elif KELAS == "11 IPA" and KURIKULUM == "K13":
            kode_kls_kur = k13_11ipa
            column_order = column_order_k13_11ipa
        elif KELAS == "10 IPS" and KURIKULUM == "K13":
            kode_kls_kur = k13_10ips
            column_order = column_order_k13_10ips
        elif KELAS == "11 IPS" and KURIKULUM == "K13":
            kode_kls_kur = k13_11ips
            column_order = column_order_k13_11ips
        # km
        elif KELAS == "4 SD" and KURIKULUM == "KM":
            kode_kls_kur = km_4sd
            column_order = column_order_km_4sd
        elif KELAS == "7 SMP" and KURIKULUM == "KM":
            kode_kls_kur = km_7smp
            column_order = column_order_km_7smp
        elif KELAS == "10 SMA" and KURIKULUM == "KM":
            kode_kls_kur = km_10sma
            column_order = column_order_km_10sma
        elif KELAS == "11 SMA" and KURIKULUM == "KM":
            kode_kls_kur = km_11sma
            column_order = column_order_km_11sma

        result_filtered = result[result['kode'].isin(kode_kls_kur)]
        result_filtered.drop_duplicates(
            subset=['nama', 'kode'], keep='first', inplace=True)

        # Menggunakan pivot_table untuk menjadikan konten kolom 'studi' sebagai header dan menghilangkan duplikat
        result_pivot = pd.pivot_table(result_filtered, index=[
            'nama', 'nonf', 'kd_lok', 'nama_sklh', 'kelas', 'idtahun'], columns='kode', values='jml_benar', aggfunc='first')
        result_pivot.reset_index(inplace=True)  # Mengatur ulang indeks

        # Ubah nama kolom
        result_pivot = result_pivot.rename(
            columns={'nama': 'NAMA', 'nonf': 'NONF', 'kd_lok': 'KD_LOK', 'nama_sklh': 'NAMA_SKLH', 'kelas': 'KELAS', 'idtahun': 'IDTAHUN',
                     'LHHQHX38': 'MAW_11KM', 'LHHQNOEV': 'SOS_11KM', 'LHHQQ7WE': 'BIO_11KM', 'LHHQTDZN': 'ENG_11KM', 'LHHQVJ2O': 'FIS_11KM', 'LHHQXWPS': 'EKO_11KM', 'LHHR041C': 'IND_11KM', 'LHHR2936': 'SEJ_11KM', 'LHHR50DU': 'KIM_11KM', 'LHHR7D5N': 'GEO_11KM', 'LHPP32X4': 'MAP_11KM',
                     'LHHPY4DQ': 'MAT_11IPS', 'LHHQ0GEG': 'SOS_11IPS', 'LHHQ2SOF': 'ENG_11IPS', 'LHHQ5KA7': 'EKO_11IPS', 'LHHQ9G2X': 'IND_11IPS', 'LHHQBZF1': 'SEJ_11IPS', 'LHHQEL75': 'GEO_11IPS',
                     'LHHPMD5N': 'MAT_11IPA', 'LHHPPGY1': 'BIO_11IPA', 'LHHPSAH8': 'FIS_11IPA', 'LHHPVWEU': 'KIM_11IPA',
                     'LHHP5ZFU': 'MAT_10KM', 'LHHP8HQU': 'IND_10KM', 'LHHPD245': 'ENG_10KM', 'LHHPGDX7': 'IPA_10KM', 'LHHPI776': 'IPS_10KM',
                     'LHHOHQW3': 'MAT_10IPS', 'LHHOL2GH': 'SOS_10IPS', 'LHHOOPEJ': 'ENG_10IPS', 'LHHOR6Q5': 'EKO_10IPS', 'LHHOUB5D': 'IND_10IPS', 'LHHOXG3D': 'SEJ_10IPS', 'LHHP0FDK': 'GEO_10IPS',
                     'LHHO4J0W': 'MAT_10IPA', 'LHHO78FV': 'BIO_10IPA', 'LHHOB3L0': 'FIS_10IPA', 'LHHODJIH': 'KIM_10IPA',
                     'LHH6H3F6': 'MAT_8SMP', 'LHH6TEO5': 'IND_8SMP', 'LHHN9AZH': 'ENG_8SMP', 'LHHNDOAI': 'IPA_8SMP', 'LHHNFJ3E': 'IPS_8SMP',
                     'LHHE7GC8': 'MAT_7KM', 'LHHEAQWK': 'IND_7KM', 'LHHEEEB5': 'ENG_7KM', 'LHHF9Q62': 'IPA_7KM', 'LHHFBCWT': 'IPS_7KM',
                     'LHHDRBXZ': 'MAT_7SMP', 'LHHDUWKS': 'IND_7SMP', 'LHHDX6U7': 'ENG_7SMP', 'LHHDZC8Y': 'IPA_7SMP', 'LHHE476J': 'IPS_7SMP',
                     'LHH5V62M': 'MAT_5SD', 'LHH6WL2C': 'IND_5SD', 'LHH7NAB5': 'ENG_5SD', 'LHHCO0Q4': 'IPA_5SD', 'LHHDAY7I': 'IPS_5SD',
                     'LHH0U12P': 'MAT_4KM', 'LHH19TQN': 'IND_4KM', 'LHH47YLV': 'ENG_4KM', 'LHH4U3Q0': 'IPAS_4KM',
                     'LHG94EEQ': 'MAT_4SD', 'LHG9KCRA': 'IND_4SD', 'LHGA44Y9': 'ENG_4SD', 'LHGALT9N': 'IPA_4SD', 'LHH0F32F': 'IPS_4SD'})

        result_pivot = result_pivot.reindex(columns=column_order)

        kelas = KELAS.lower().replace(" ", "")
        kurikulum = KURIKULUM.lower()
        tahun = TAHUN.replace("-", "")

        path_file = f"{kelas}_pat_sm2_{kurikulum}_{tahun}_pivot.xlsx"

        # Simpan file ke direktori temporer
        temp_dir = tempfile.gettempdir()
        file_path = temp_dir + '/' + path_file
        # wb.save(file_path)

        # Menyimpan DataFrame ke file Excel
        result_pivot.to_excel(file_path, index=False)
        st.success("File siap diunduh!")

        # Tombol unduh file
        with open(file_path, "rb") as f:
            bytes_data = f.read()
        st.download_button(label="Unduh File", data=bytes_data,
                           file_name=path_file)
    #     st.write(result_pivot)
    # else:
    #     st.write("File tidak ditemukan atau gagal diunggah.")
