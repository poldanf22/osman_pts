import pickle
import streamlit as st
import streamlit_authenticator as stauth
from pathlib import Path
from PIL import Image
import subprocess
from streamlit_option_menu import option_menu

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

if authentication_status == False:
    st.error("Username/kode akses salah")

if authentication_status == None:
    st.warning("Silahkan masukan username dan kode akses")

if authentication_status:

    authenticator.logout("Logout", "sidebar")
    with st.sidebar:
        selected_file = option_menu(
            menu_title="Pilih file:",
            options=["Pivot"],
        )
    if selected_file == "Pivot":
        subprocess.Popen(["streamlit", "run", "pivot.py"])

        # st.sidebar.title(f"Assalamu'alaikum")
        # selected_file = st.sidebar.selectbox(
        #     "Pilih file:", ("pivot.py", "nilai_std_sd_smp_10km.py"))

        # if st.sidebar.button("Buka File"):
        #     # Ganti folder_path dengan jalur folder yang berisi file-file tersebut
        #     path_file = f"{selected_file}"
        #     subprocess.Popen(["streamlit", "run", path_file])
        #     st.sidebar.warning("Mohon ditunggu sampai muncul Tab Baru!")
