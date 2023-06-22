import pickle
import streamlit as st
import streamlit_authenticator as stauth
from pathlib import Path
import subprocess

# User Authentication
names = ["TI NF", "Polda NF"]
usernames = ["admin1", "admin2"]

# load hashed kd_akses
file_path = Path(__file__).parent / "hashed_pw.pkl"
with file_path.open("rb") as file:
    hashed_kd_akses = pickle.load(file)

authenticator = stauth.Authenticate(
    names, usernames, hashed_kd_akses, "lookup", "abcdef")
name, authentication_status, username = authenticator.login("Login", "main")

if authentication_status == False:
    st.error("Username/kode akses salah!")

if authentication_status == None:
    st.warning("Silahkan masukan username dan kode akses")
