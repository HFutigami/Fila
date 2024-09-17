import streamlit as st
import pandas as pd
from office365.sharepoint.files.file import File
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from datetime import datetime
import io
import plotly.express as px
from workalendar.america import Brazil
from datetime import timedelta
import webbrowser
import os

### VERIFICA SE FOI ESTABELECIDA UMA CONEXÃO,
### CASO CONTRÁRIO DIRECIONA O USUÁRIO PARA A TELA INICIAL
if 'connection' not in st.session_state:
    with st.form("login"):
        st.write("Login")
        user = st.text_input('Usuário')
        senha = st.text_input('Senha', type='password')
        submitted = st.form_submit_button("Login")
        if submitted:
            if senha == st.secrets['credenciais'].SENHA and user == st.secrets['credenciais'].USER:
                st.session_state['connection'] = 'editor'
            else:
                st.warning('Usuário ou senha inválidos!', icon="⚠️")

