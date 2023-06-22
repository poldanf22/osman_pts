import pickle
import streamlit as st
import streamlit_authenticator as stauth
from pathlib import Path
from PIL import Image
import subprocess

# User Authentication
names = ["TI Polda NF 1", "TI Polda NF 2"]
usernames = ["admin1", "admin2"]

# load hashed kd_akses
file_path = Path(__file__).parent / "hashed_pw.pkl"
with file_path.open("rb") as file:
    hashed_kd_akses = pickle.load(file)

authenticator = stauth.Authenticate(
    names, usernames, hashed_kd_akses, "lookup", "abcdef")
name, authentication_status, username = authenticator.login("Login", "main")


def after_login():
    if authentication_status == False:
        st.error("Username/kode akses salah")

    if authentication_status == None:
        st.warning("Silahkan masukan username dan kode akses")

    if authentication_status:

        authenticator.logout("Logout", "sidebar")
        # st.sidebar.title(f"Welcome {name}")
        selected_file = st.sidebar.selectbox(
            "Pilih file:", ("pivot.py", "nilai_std_sd_smp_10km.py"))

        if st.sidebar.button("Buka File"):
           try:
            with open(selected_file, "r") as file:
                contents = file.read()
                st.code(contents)
                except FileNotFoundError:
                    st.sidebar.error("File tidak ditemukan")


if __name__ == "__main__":
    after_login()
