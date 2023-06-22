import pickle
import streamlit as st
import streamlit_authenticator as stauth
from pathlib import Path
import subprocess

# User Authentication
names = ["Peter Parker", "Rebecca Miller"]
usernames = ["pparker", "rmiller"]

# Load hashed passwords
file_path = Path(__file__).parent / "hashed_pw.pkl"
with file_path.open("rb") as file:
    hashed_passwords = pickle.load(file)

authenticator = stauth.Authenticate(
    names, usernames, hashed_passwords, "lookup", "abcdef")
name, authentication_status, username = authenticator.login("Login", "main")

# Cek apakah pengguna sudah terotentikasi
if "is_authenticated" not in st.session_state:
    st.session_state.is_authenticated = False

# Menampilkan halaman yang sesuai berdasarkan status autentikasi
if authentication_status:
    st.session_state.is_authenticated = True

if st.session_state.is_authenticated:
    # Tambahkan file lain yang ingin diakses
    selected_file = st.selectbox(
        "Pilih file:", ("pivot.py", "nilai_std_sd_smp_10km.py"))

    if st.button("Buka File"):
        # Ganti folder_path dengan jalur folder yang berisi file-file tersebut
        file_path = f"halaman/{selected_file}"
        subprocess.Popen(["streamlit", "run", file_path])
        st.warning("Mohon ditunggu sampai muncul Tab Baru!")
else:
    if authentication_status is False:
        st.error("Username/Password is incorrect")

    if authentication_status is None:
        st.warning("Please enter your username and password")
