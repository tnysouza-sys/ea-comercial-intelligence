import streamlit_authenticator as stauth

def gerar_hash(senha):
    return stauth.Hasher.hash(senha)