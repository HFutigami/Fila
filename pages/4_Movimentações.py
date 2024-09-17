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
        prioridade = df_sharep(prioridade_url, tipo='excel')
        df_prioridades = df[(df['CAIXA'].isin(prioridade['CAIXAS'])) | (df[df['NUM OS'].isin(prioridade['CAIXAS'])])].copy()
        df_prioridades['FILA'] = df['ENDEREÇO'].apply(lambda  x: 1 if x not in ['LAB', 'EQUIPE TECNICA', 'GESTAO DE ATIVOS', 'QUALIDADE', 'RETRIAGEM'] else 0)
        df_prioridades['SAÍDA'] = df['ENDEREÇO'].apply(lambda x: 1 if x in ['LAB', 'EQUIPE TECNICA', 'GESTAO DE ATIVOS', 'QUALIDADE', 'RETRIAGEM'] else 0)
        
        return df_prioridades


    def create_df_prioridades_resumido(df):

        df_prioridades_resumido = df.groupby(['CLIENTE', 'EQUIPAMENTO'])[['FILA', 'SAÍDA']].sum().reset_index()
        df_prioridades_resumido.rename(columns={'FILA':'QTD EM FILA', 'SAÍDA':'QTD FORA DO FILA'}, inplace=True)
        try:
            df_prioridades_resumido.sort_values(['CLIENTE', 'EQUIPAMENTO'], inplace=True)
        except:
            pass

        return df_prioridades_resumido


    def create_fig_criticos(df):
        df = df[~df['ENDEREÇO'].isin(['LAB', 'EQUIPE TECNICA', 'GESTAO DE ATIVOS', 'QUALIDADE', 'RETRIAGEM'])]
        df['CAIXA'] = df['CAIXA'].astype('str')
        df['CAIXA'] = "ㅤ" + df['CAIXA']
        df['ENTRADA FILA'] = df['ENTRADA FILA'].astype('str')
        df['RÓTULO'] = df['CLIENTE'] + ' - ' + df['ENDEREÇO'] + ' - ' + \
                       df['ENTRADA FILA'].str.replace('-', '/').str.split(" ").str[0]

        df = df.groupby(['CAIXA', 'RÓTULO', '% DO SLA'])['SERIAL'].count().reset_index().sort_values('% DO SLA',
                                                                                                     ascending=True).tail(
            10)

        fig = px.bar(df,
                     x='% DO SLA',
                     y='CAIXA',
                     color='% DO SLA',
                     orientation='h',
                     text='RÓTULO',
                     color_continuous_scale=[(0, "#008000"),
                                             (0.2, "#32CD32"),
                                             (0.45, "#FFD700"),
                                             (0.8, "#FF8C00"),
                                             (1, "#8B0000")],
                     range_color=[0, 1])
        return fig


    def create_fig_status(df):
        df = df.copy()
        df.loc[~df['ENDEREÇO'].isin(['LAB', 'EQUIPE TECNICA', 'GESTAO DE ATIVOS', 'QUALIDADE', 'RETRIAGEM']), 'ENDEREÇO'] = 'FILA'
        df.loc[df['ENDEREÇO'].isin(['LAB', 'EQUIPE TECNICA', 'GESTAO DE ATIVOS', 'QUALIDADE', 'RETRIAGEM']), 'ENDEREÇO'] = 'FORA DO FILA'
        df = df.groupby(['ENDEREÇO'])[['SERIAL']].count().reset_index()

        fig = px.pie(df,
                                    names='ENDEREÇO',
                                    values='SERIAL',
                                    color='ENDEREÇO',
                                    hole=0.4,
                                    color_discrete_map={'FORA DO FILA':'#008000',
                                                        'FILA':'#8B0000'},)
        fig.update_traces(textinfo='value+percent')

        return fig

    
    if 'historico_fila' not in st.session_state:
        st.session_state['historico_fila'] = create_df_historico_movimentações()
        historico_fila = st.session_state['historico_fila']
    else:
        historico_fila = st.session_state['historico_fila']

    if 'prioridades_df' not in st.session_state or 'prioridade_resumido_df' not in st.session_state:
        st.session_state['prioridades_df'] = create_df_prioridades(historico_fila)
        prioridades_df = st.session_state['prioridades_df']

        st.session_state['prioridade_resumido_df'] = create_df_prioridades_resumido(prioridades_df)
        prioridade_resumido_df = st.session_state['prioridade_resumido_df']
    else:
        prioridades_df = st.session_state['prioridades_df']
        prioridade_resumido_df = st.session_state['prioridade_resumido_df']


    r0c1, r0c2, r0c3, r0c4 = st.columns(4, gap='large')
    st.write('')
    r1c1, r1c2 = st.columns(2, gap='large')
    r2c1, r2c2 = st.columns([0.7, 0.3], gap='large')

    r1c1.write('Saldo resumido de equipamentos prioritários.')
    prioridade_resumido_stdf = r1c1.dataframe(prioridade_resumido_df[['CLIENTE', 'EQUIPAMENTO', 'QTD EM FILA', 'QTD FORA DO FILA']],
                                            hide_index=True,
                                            use_container_width=True,
                                            on_select='rerun')

    if prioridade_resumido_stdf.selection.rows:
        prioridade_resumido_df['CONCATENADO'] = prioridade_resumido_df['CLIENTE'] + prioridade_resumido_df['EQUIPAMENTO']
        prioridades_df['CONCATENADO'] = prioridades_df['CLIENTE'] + prioridades_df['EQUIPAMENTO']
        filtro_saldo = list(prioridade_resumido_df.iloc[prioridade_resumido_stdf.selection.rows]['CONCATENADO'])
        prioridades_selecao = prioridades_df[prioridades_df['CONCATENADO'].isin(filtro_saldo)]
        st.session_state['prioridades_selecao'] = prioridades_selecao
        r0c1.metric('Total de equipamentos (seleção)',
                    '{:,}'.format(sum(prioridade_resumido_df.iloc[prioridade_resumido_stdf.selection.rows]['QTD EM FILA'])).replace(',', '.'))
    else:
        r0c1.metric('Total de equipamentos',
                    '{:,}'.format(sum(prioridade_resumido_df['QTD EM FILA'])).replace(',', '.'))
        

    if 'prioridades_selecao' in st.session_state and prioridade_resumido_stdf.selection.rows:
        if len(st.session_state['prioridades_selecao']) > 0:
            r1c2.write('Classificação dos equipamentos no fila de acordo com % do SLA.')
            r1c2.plotly_chart(create_fig_criticos(st.session_state['prioridades_selecao'][
                                                      ~st.session_state['prioridades_selecao'][
                                                          '% DO SLA'].isna()].copy()))

            r2c1.write('Saldo detalhado de equipamentos prioritários.')
            r2c1.dataframe(prioridades_selecao[[
                'ENDEREÇO', 'CAIXA', 'SERIAL', 'CLIENTE',
                'EQUIPAMENTO', 'NUM OS', 'ENTRADA GERFLOOR',
                'ENTRADA FILA', 'AGING TOTAL', 'AGING FILA',
                'STATUS'
            ]],
                           hide_index=True,
                           use_container_width=True,
                           column_config={
                               'ENTRADA GERFLOOR':st.column_config.DateColumn('ENTRADA GERFLOOR', format="DD/MM/YYYY"),
                               'ENTRADA FILA':st.column_config.DateColumn('ENTRADA FILA', format="DD/MM/YYYY HH:mm:ss")
                           })
            r2c2.write('Relação de equipamentos em fila x fora do fila.')
            r2c2.plotly_chart(create_fig_status(st.session_state['prioridades_selecao']))
        

