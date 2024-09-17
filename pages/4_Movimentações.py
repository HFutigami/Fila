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
    st.switch_page('Painel.py')

else:

    ### CONFIGURAÇÕES INICIAS DO STREAMLIT
    st.set_page_config('ESTOQUE • FILA', page_icon='https://i.imgur.com/mOEfCM8.png', layout='wide')

    st.image('https://seeklogo.com/images/G/gertec-logo-D1C911377C-seeklogo.com.png?v=637843433630000000', width=200)
    st.header('', divider='gray')

    st.sidebar.title('MÓDULOS')
    if st.session_state['connection'] == 'editor':
        st.sidebar.page_link('pages/4_Movimentações.py', label='MOVIMENTAÇÕES', disabled=True)
    st.sidebar.page_link('pages/1_Contratos.py', label='CONTRATO')
    st.sidebar.page_link('pages/2_Varejo.py', label='VAREJO')
    st.sidebar.page_link('pages/3_OS Interna.py', label='OS INTERNA')

    ### LINKS ONDE SÃO ARMAZENADOS OS DADOS DO FILA
    sharepoint_fila_url = 'https://gertecsao.sharepoint.com/sites/PowerBi-Estoque/'
    sharepoint_os_url = 'https://gertecsao.sharepoint.com/sites/RecebimentoLogstica/'
    folder_in_sharepoint = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Fila'
    sharepoint_user = st.secrets.sharepoint.USER
    sharepoint_password = st.secrets.sharepoint.SENHA

    saldo_fila_url = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Fila/saldo.parquet'
    varejo_liberado_url = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Fila/Varejo%20Liberado/'
    sla_contratos_url = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Fila/SlaContratos.csv'
    abertura_os_url = '/sites/RecebimentoLogstica/Documentos%20Compartilhados/General/Recebimento%20-%20Abertura%20de%20OS.xlsx'
    prioridade_url = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Fila/Prioridades.xlsx'

    calendario = Brazil()


    ### FUNÇÕES
    def df_sharep(file_url, tipo='parquet', sheet='', site=sharepoint_fila_url):
        """Gera um DataFrame a partir de um diretório do SharePoint."""
        auth = AuthenticationContext(site)
        auth.acquire_token_for_user(sharepoint_user, sharepoint_password)
        ctx = ClientContext(saldo_fila_url, auth)
        web = ctx.web
        ctx.execute_query()

        file_response = File.open_binary(ctx, file_url)
        bytes_file_obj = io.BytesIO()
        bytes_file_obj.write(file_response.content)
        bytes_file_obj.seek(0)
        if tipo == 'parquet':
            return pd.read_parquet(bytes_file_obj)
        elif tipo == 'csv':
            return pd.read_csv(bytes_file_obj, sep=";")
        elif tipo == 'excel':
            if sheet != '':
                return pd.read_excel(bytes_file_obj, sheet, dtype='str')
            else:
                return pd.read_excel(bytes_file_obj, dtype='str')


    def create_df_historico_movimentações():
        # SLA Contratos
        sla_contratos = df_sharep(sla_contratos_url, tipo="csv")
        sla_contratos.rename(columns={'SLA': 'PRAZO'}, inplace=True)
        sla_contratos['FLUXO'] = 'CONTRATO'
        sla_contratos.set_index(['CLIENTE', 'FLUXO'], inplace=True)

        # Saldo geral
        historico_fila = df_sharep(saldo_fila_url)

        historico_fila['FLUXO'] = historico_fila['FLUXO'].str.replace('000001', 'CONTRATO')
        historico_fila['FLUXO'] = historico_fila['FLUXO'].str.replace('000002', 'VAREJO')
        historico_fila['FLUXO'] = historico_fila['FLUXO'].str.replace('000003', 'VAREJO')
        historico_fila['FLUXO'] = historico_fila['FLUXO'].str.replace('000004', 'OS INTERNA')

        historico_fila['GARANTIA'] = historico_fila['GARANTIA'].str.upper()
        historico_fila['CLIENTE'] = historico_fila['CLIENTE'].str.upper()

        historico_fila = historico_fila[historico_fila['ENTRADA GERFLOOR'] != 'Nenhum registro encontrado']
        historico_fila['ENTRADA GERFLOOR'] = pd.to_datetime(
            historico_fila.loc[historico_fila['ENTRADA GERFLOOR'] != 'Nenhum registro encontrado', 'ENTRADA GERFLOOR'],
            format='%d/%m/%Y %I:%M:%S %p')

        historico_fila = historico_fila[['ENDEREÇO',
                                         'CAIXA',
                                         'SERIAL',
                                         'CLIENTE',
                                         'EQUIPAMENTO',
                                         'NUM OS',
                                         'FLUXO',
                                         'GARANTIA',
                                         'ENTRADA GERFLOOR',
                                         'ENTRADA FILA',
                                         'SAÍDA FILA']]

        historico_fila['ULTIMA DATA'] = historico_fila['SAÍDA FILA']
        historico_fila.loc[historico_fila['ULTIMA DATA'].isna(), 'ULTIMA DATA'] = datetime.now()
        historico_fila['AGING TOTAL'] = historico_fila.apply(
            lambda row: calendario.get_working_days_delta(row['ENTRADA GERFLOOR'], row['ULTIMA DATA']), axis=1) + 1
        historico_fila['AGING TOTAL'] = historico_fila['AGING TOTAL'].astype('int')
        historico_fila['AGING FILA'] = historico_fila.apply(
            lambda row: calendario.get_working_days_delta(row['ENTRADA FILA'], row['ULTIMA DATA']), axis=1) + 1
        historico_fila['AGING FILA'] = historico_fila['AGING FILA'].astype('int')

        historico_fila = historico_fila.join(sla_contratos, on=['CLIENTE', 'FLUXO'], how='left')
        historico_fila.loc[historico_fila['PRAZO'].isna(), 'PRAZO'] = 30
        historico_fila['% DO SLA'] = None
        historico_fila.loc[~historico_fila['PRAZO'].isna(), '% DO SLA'] = historico_fila['AGING TOTAL'] / \
                                                                          historico_fila['PRAZO']
        historico_fila['STATUS'] = None
        historico_fila.loc[
            (historico_fila['% DO SLA'] > 0.0) & (historico_fila['% DO SLA'] <= 0.1), 'STATUS'] = "RÁPIDO"
        historico_fila.loc[(historico_fila['% DO SLA'] > 0.1) & (historico_fila['% DO SLA'] <= 0.3), 'STATUS'] = "MÉDIO"
        historico_fila.loc[(historico_fila['% DO SLA'] > 0.3) & (historico_fila['% DO SLA'] <= 0.6), 'STATUS'] = "LENTO"
        historico_fila.loc[
            (historico_fila['% DO SLA'] > 0.6) & (historico_fila['% DO SLA'] <= 1.0), 'STATUS'] = "CRÍTICO"
        historico_fila.loc[(historico_fila['% DO SLA'] > 1.0), 'STATUS'] = "SLA ESTOURADO"

        return historico_fila

    
    def create_df_prioridades(df):
        df_prioridades = df[df['CAIXA'].isin(df_sharep(prioridade_url, tipo='excel')['CAIXAS'])].copy()
        df_prioridades['FILA'] = df['ENDEREÇO'].apply(lambda  x: 1 if x not in ['LAB', 'EQUIPE TECNICA', 'GESTAO DE ATIVOS', 'QUALIDADE', 'RETRIAGEM'] else 0)
        df_prioridades['SAÍDA'] = df['ENDEREÇO'].apply(lambda x: 1 if x in ['LAB', 'EQUIPE TECNICA', 'GESTAO DE ATIVOS', 'QUALIDADE', 'RETRIAGEM'] else 0)
        
        return df_prioridades


    def create_df_saldo_contratos_resumido(df):
        
        abertura_os = df_sharep(abertura_os_url, 'excel', 'BASE', sharepoint_os_url)
        abertura_os = abertura_os[abertura_os['ABRIR O.S'] != "0"]
        abertura_os.reset_index(drop=True, inplace=True)
        abertura_os.loc[abertura_os['CLIENTE GERFLOOR'].isna(), 'CLIENTE GERFLOOR'] = abertura_os.loc[
            abertura_os['CLIENTE GERFLOOR'].isna(), 'CLIENTES'].apply(lambda x: x.split(" - ", maxsplit=1)[0])
        abertura_os.loc[abertura_os['EQUIPAMENTO GERFLOOR'].isna(), 'EQUIPAMENTO GERFLOOR'] = abertura_os.loc[
            abertura_os['EQUIPAMENTO GERFLOOR'].isna(), 'CLIENTES'].apply(lambda x: x.split(" - ", maxsplit=1)[1])
        abertura_os = abertura_os.rename(columns={'CLIENTE GERFLOOR': 'CLIENTE',
                                                  'EQUIPAMENTO GERFLOOR': 'EQUIPAMENTO'}).set_index(
            ['CLIENTE', 'EQUIPAMENTO']).drop(['PENDÊNCIA', 'O.S ABERTA', 'CLIENTES'], axis=1)

        df.loc[df['CLIENTE'].str.startswith('COBRA'), 'CLIENTE'] = 'COBRA'
        df.loc[df['CLIENTE'].str.startswith('BB'), 'CLIENTE'] = 'COBRA'

        df_saldo_atual_contratos_resumido = df.groupby(['CLIENTE', 'EQUIPAMENTO'])[['SERIAL']].count().reset_index()

        df_saldo_atual_contratos_resumido = df_saldo_atual_contratos_resumido.join(abertura_os,
                                                                                   on=['CLIENTE', 'EQUIPAMENTO'],
                                                                                   how='outer')
        df_saldo_atual_contratos_resumido.loc[df_saldo_atual_contratos_resumido['SERIAL'].isna(), 'SERIAL'] = 0
        df_saldo_atual_contratos_resumido.SERIAL = df_saldo_atual_contratos_resumido.SERIAL.astype(int)
        df_saldo_atual_contratos_resumido.loc[df_saldo_atual_contratos_resumido['ABRIR O.S'].isna(), 'ABRIR O.S'] = 0
        df_saldo_atual_contratos_resumido['ABRIR O.S'] = df_saldo_atual_contratos_resumido['ABRIR O.S'].astype(int)
        df_saldo_atual_contratos_resumido.rename(columns={'SERIAL': 'QTD FILA',
                                                          'ABRIR O.S': 'QTD OS'}, inplace=True)
        df_saldo_atual_contratos_resumido = df_saldo_atual_contratos_resumido[
            ['CLIENTE', 'EQUIPAMENTO', 'QTD OS', 'QTD FILA']]
        try:
            df_saldo_atual_contratos_resumido.sort_values(['CLIENTE', 'EQUIPAMENTO'], inplace=True)
        except:
            pass

        return df_saldo_atual_contratos_resumido

    
    if 'historico_fila' not in st.session_state:
        st.session_state['historico_fila'] = create_df_historico_movimentações()
        historico_fila = st.session_state['historico_fila']
    else:
        historico_fila = st.session_state['historico_fila']

    if 'prioridades_df' not in st.session_state:
        st.session_state['prioridades_df'] = create_df_prioridades(historico_fila)
        prioridades_df = st.session_state['prioridades_df']
    else:
        prioridades_df = st.session_state['prioridades_df']

    st.dataframe(prioridades_df)

