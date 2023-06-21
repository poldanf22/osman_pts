import streamlit as st
from passlib.hash import pbkdf2_sha256
from sqlalchemy import create_engine, Column, String
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.declarative import declarative_base
import subprocess

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

# # Membuat objek SQLAlchemy Engine untuk koneksi MySQL
# engine = create_engine(
#     'mysql+pymysql://root@localhost/db_streamlit')

# Membuat objek SQLAlchemy Engine untuk koneksi Server
engine = create_engine(
    'mysql+pymysql://poldanf:polda4lhamdulillaHoke@10.212.37.103/db_streamlit')

# Membuat objek Session
Session = sessionmaker(bind=engine)
session = Session()

# Membuat kelas model User untuk tabel pengguna
Base = declarative_base()


class User(Base):
    __tablename__ = 'users'
    nopeg = Column(String, primary_key=True)
    password = Column(String)


# Membuat fungsi untuk mengenkripsi password
def encrypt_password(password):
    return pbkdf2_sha256.hash(password)


# Fungsi untuk memverifikasi password
def verify_password(password, hashed_password):
    return pbkdf2_sha256.verify(password, hashed_password)


# Fungsi untuk memeriksa kredensial pengguna
def authenticate(nopeg, password):
    user = session.query(User).filter_by(nopeg=nopeg).first()
    if user and verify_password(password, user.password):
        return True
    return False


# Halaman login dan registrasi
def login_register():
    st.title("Login / Registrasi")

    # Form input
    nopeg = st.text_input("Nopeg")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        # Periksa autentikasi
        if authenticate(nopeg, password):
            # Set session state untuk menandai login berhasil
            st.session_state.is_authenticated = True
            st.experimental_rerun()
        else:
            st.error("nopeg atau password salah")

    if st.button("Daftar"):
        # Periksa apakah pengguna sudah terdaftar
        if session.query(User).filter_by(nopeg=nopeg).first():
            st.error("nopeg sudah digunakan")
        else:
            # Enkripsi password
            hashed_password = encrypt_password(password)

            # Buat objek User baru
            user = User(nopeg=nopeg, password=hashed_password)

            # Simpan objek User ke database
            session.add(user)
            session.commit()

            st.success("Registrasi berhasil, silakan login.")


# Halaman konten setelah login berhasil
def after_login():
    # Cek apakah pengguna sudah terotentikasi
    if "is_authenticated" not in st.session_state:
        st.session_state.is_authenticated = False

    # Menampilkan halaman yang sesuai berdasarkan status autentikasi
    if st.session_state.is_authenticated:
        # Tambahkan file lain yang ingin diakses
        selected_file = st.selectbox(
            "Pilih file:", ("pivot.py", "nilai_std_sd_smp_10km.py"))

        if st.button("Buka File"):
            # Ganti folder_path dengan jalur folder yang berisi file-file tersebut
            file_path = f"halaman/{selected_file}"
            subprocess.Popen(["streamlit", "run", file_path])
            st.warning("Mohon ditunggu sampai muncul Tab Baru!")


if __name__ == "__main__":
    login_register()
    after_login()
