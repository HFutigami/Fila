import streamlit as st
import pandas as pd
from office365.sharepoint.files.file import File
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from datetime import datetime
import io
import plotly.express as px
from workalendar.america import Brazil
from time import sleep

if 'connection' not in st.session_state:
    with st.form('login', clear_on_submit=True):
        usuario = st.text_input(label='Usuário')
        senha = st.text_input(label='Senha', type='password')
        submitted = st.form_submit_button("Login")
        if submitted:
            with st.spinner('Conectando...'):
                if usuario == 'julio':
                    st.session_state['connection'] = 'editor'
                    st.rerun()
            st.session_state['connection'] = 'a'
            st.rerun()
else:
    st.set_page_config('ESTOQUE • FILA', page_icon='https://i.imgur.com/mOEfCM8.png', layout='wide')
        
    st.image('https://seeklogo.com/images/G/gertec-logo-D1C911377C-seeklogo.com.png?v=637843433630000000', width=200)
    st.header('', divider='gray')

    st.sidebar.title('MÓDULOS')
    if st.session_state['connection'] == 'editor':
        st.sidebar.page_link('pages/4_Movimentações.py', label='MOVIMENTAÇÕES')
    st.sidebar.page_link('pages/1_Contratos.py', label='CONTRATO')
    st.sidebar.page_link('pages/2_Varejo.py', label='VAREJO')
    st.sidebar.page_link('pages/3_OS Interna.py', label='OS INTERNA')
