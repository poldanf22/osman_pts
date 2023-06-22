import pickle
from pathlib import Path

import streamlit_authenticator as stauth

name = ["TI Polda NF", "TI Polda NF"]
username = ["admin1", "admin2"]
kd_akses = ["nfcemerlang", "nfcemerlang"]

hashed_kd_akses = stauth.Hasher(kd_akses).generate()

file_path = Path(__file__).parent / "hashed_pw.pkl"
with file_path.open("wb") as file:
    pickle.dump(hashed_kd_akses, file)
