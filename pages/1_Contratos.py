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
        st.sidebar.page_link('pages/4_Movimentações.py', label='MOVIMENTAÇÕES')
    st.sidebar.page_link('pages/1_Contratos.py', label='CONTRATO', disabled=True)
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
        sla_contratos.rename(columns={'SLA':'PRAZO'}, inplace=True)
        sla_contratos['FLUXO'] = 'CONTRATO'
        sla_contratos.set_index(['CLIENTE', 'FLUXO'], inplace=True)
        
        # Saldo geral
        historico_fila = df_sharep(saldo_fila_url)

        lista_clientes_contrato = ['REDE',
                                    'Stone',
                                    'Wyntech',
                                    'Getnet',
                                    'MERCADOPAGO.COM',
                                    'CIELO',
                                    'Procomp',
                                    'NCR',
                                    'COBRA_JOAO PESSOA',
                                    'COBRA_GOIANIA',
                                    'COBRA',
                                    'COBRA_PORTO VELHO',
                                    'COBRA_TERESINA',
                                    'FD',
                                    'COBRA_RIO DE JANEIRO_2',
                                    'COBRA_BAURU',
                                    'COBRA_CURITIBA',
                                    'COBRA_RIBEIRAO PRETO',
                                    'BB_CAMPINAS',
                                    'COBRA_FLORIANOPOLIS',
                                    'COBRA_BELEM',
                                    'COBRA_JOINVILLE',
                                    'COBRA_LONDRINA',
                                    'COBRA_BARUERI',
                                    'BB',
                                    'COBRA_JABOATAO DOS GUARARAPES',
                                    'MARTINS',
                                    'Paygo',
                                    'ALEXANDRE',
                                    'MERCADO PAGO',
                                    'SERVICO',
                                    'COBRA_SAO LUIS',
                                    'COBRA_PORTO ALEGRE',
                                    'PAGSEGURO_OSASCO',
                                    'COBRA_BRASILIA',
                                    'MOOZ']
        lista_clientes_varejo = ['MARSISTEM']
        historico_fila.loc[historico_fila['CLIENTE'].isin(lista_clientes_contrato), 'FLUXO'] = '000001'
        historico_fila.loc[historico_fila['CLIENTE'] == 'GERTEC', 'FLUXO'] = '000004'
        historico_fila.loc[historico_fila['CLIENTE'].isin(lista_clientes_varejo), 'FLUXO'] = '000002'
        historico_fila.loc[historico_fila['FLUXO'] == 'NÃO DEFINIDO', 'FLUXO'] = '000002'
        historico_fila['FLUXO'] = historico_fila['FLUXO'].str.replace('000001', 'CONTRATO')
        historico_fila['FLUXO'] = historico_fila['FLUXO'].str.replace('000002', 'VAREJO')
        historico_fila['FLUXO'] = historico_fila['FLUXO'].str.replace('000003', 'VAREJO')
        historico_fila['FLUXO'] = historico_fila['FLUXO'].str.replace('000004', 'OS INTERNA')

        filtro_3m = ((historico_fila['SAÍDA FILA'] >= datetime.today()-timedelta(month=3, hours=datetime.today().hour, minutes=datetime.today().minute)) &
                    (historico_fila['SAÍDA FILA'] <= datetime.today()-timedelta(hours=datetime.today().hour+1, minutes=datetime.today().minute)))

        historico_fila = historico_fila[(filtro_3m) | (historico_fila['SAÍDA FILA'].isna())]
                     
        historico_fila['GARANTIA'] = historico_fila['GARANTIA'].str.upper()
        historico_fila['CLIENTE'] = historico_fila['CLIENTE'].str.upper()
        
        historico_fila = historico_fila[historico_fila['ENTRADA GERFLOOR'] != 'Nenhum registro encontrado']
        historico_fila['ENTRADA GERFLOOR'] = pd.to_datetime(historico_fila.loc[historico_fila['ENTRADA GERFLOOR'] != 'Nenhum registro encontrado', 'ENTRADA GERFLOOR'], format='%d/%m/%Y %I:%M:%S %p')
        
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
        historico_fila['AGING TOTAL'] = historico_fila.apply(lambda row: calendario.get_working_days_delta(row['ENTRADA GERFLOOR'], row['ULTIMA DATA']), axis=1) + 1
        historico_fila['AGING TOTAL'] = historico_fila['AGING TOTAL'].astype('int')
        historico_fila['AGING FILA'] = historico_fila.apply(lambda row: calendario.get_working_days_delta(row['ENTRADA FILA'], row['ULTIMA DATA']), axis=1) + 1
        historico_fila['AGING FILA'] = historico_fila['AGING FILA'].astype('int')
        
        historico_fila = historico_fila.join(sla_contratos, on=['CLIENTE', 'FLUXO'], how='left')
        historico_fila.loc[historico_fila['PRAZO'].isna(), 'PRAZO'] = 30
        historico_fila['% DO SLA'] = None
        historico_fila.loc[~historico_fila['PRAZO'].isna(), '% DO SLA'] = historico_fila['AGING TOTAL']/historico_fila['PRAZO']
        historico_fila['STATUS'] = None
        historico_fila.loc[(historico_fila['% DO SLA'] > 0.0) & (historico_fila['% DO SLA'] <= 0.1), 'STATUS'] = "RÁPIDO"
        historico_fila.loc[(historico_fila['% DO SLA'] > 0.1) & (historico_fila['% DO SLA'] <= 0.3), 'STATUS'] = "MÉDIO"
        historico_fila.loc[(historico_fila['% DO SLA'] > 0.3) & (historico_fila['% DO SLA'] <= 0.6), 'STATUS'] = "LENTO"
        historico_fila.loc[(historico_fila['% DO SLA'] > 0.6) & (historico_fila['% DO SLA'] <= 1.0), 'STATUS'] = "CRÍTICO"
        historico_fila.loc[(historico_fila['% DO SLA'] > 1.0), 'STATUS'] = "SLA ESTOURADO"

        return historico_fila


    def create_df_saldo_contratos(df):
        df_saldo_atual_contratos = df
        df_saldo_atual_contratos = df_saldo_atual_contratos[(df_saldo_atual_contratos['FLUXO'] == 'CONTRATO') & (df_saldo_atual_contratos['ENDEREÇO'] != 'LAB')]

        return df_saldo_atual_contratos
    

    def create_df_saldo_contratos_resumido(df):

        abertura_os = df_sharep(abertura_os_url, 'excel', 'BASE', sharepoint_os_url)
        abertura_os = abertura_os[abertura_os['ABRIR O.S'] != "0"]
        abertura_os.reset_index(drop=True, inplace=True)
        abertura_os.loc[abertura_os['CLIENTE GERFLOOR'].isna(), 'CLIENTE GERFLOOR'] = abertura_os.loc[abertura_os['CLIENTE GERFLOOR'].isna(), 'CLIENTES'].apply(lambda x: x.split(" - ", maxsplit=1)[0])
        abertura_os.loc[abertura_os['EQUIPAMENTO GERFLOOR'].isna(), 'EQUIPAMENTO GERFLOOR'] = abertura_os.loc[abertura_os['EQUIPAMENTO GERFLOOR'].isna(), 'CLIENTES'].apply(lambda x: x.split(" - ", maxsplit=1)[1])
        abertura_os = abertura_os.rename(columns={'CLIENTE GERFLOOR':'CLIENTE',
                                    'EQUIPAMENTO GERFLOOR':'EQUIPAMENTO'}).set_index(['CLIENTE', 'EQUIPAMENTO']).drop(['PENDÊNCIA', 'O.S ABERTA', 'CLIENTES'], axis=1)
 
        df.loc[df['CLIENTE'].str.startswith('COBRA'), 'CLIENTE'] = 'COBRA'
        df.loc[df['CLIENTE'].str.startswith('BB'), 'CLIENTE'] = 'COBRA'

        df_saldo_atual_contratos_resumido = df.groupby(['CLIENTE', 'EQUIPAMENTO'])[['SERIAL']].count().reset_index()
        
        df_saldo_atual_contratos_resumido = df_saldo_atual_contratos_resumido.join(abertura_os, on=['CLIENTE', 'EQUIPAMENTO'], how='outer')
        df_saldo_atual_contratos_resumido.loc[df_saldo_atual_contratos_resumido['SERIAL'].isna(), 'SERIAL'] = 0
        df_saldo_atual_contratos_resumido.SERIAL = df_saldo_atual_contratos_resumido.SERIAL.astype(int)
        df_saldo_atual_contratos_resumido.loc[df_saldo_atual_contratos_resumido['ABRIR O.S'].isna(), 'ABRIR O.S'] = 0
        df_saldo_atual_contratos_resumido['ABRIR O.S'] = df_saldo_atual_contratos_resumido['ABRIR O.S'].astype(int)
        df_saldo_atual_contratos_resumido.rename(columns={'SERIAL':'QTD FILA',
                                                          'ABRIR O.S':'QTD OS'}, inplace=True)
        df_saldo_atual_contratos_resumido = df_saldo_atual_contratos_resumido[['CLIENTE', 'EQUIPAMENTO', 'QTD OS', 'QTD FILA']]
        try:
            df_saldo_atual_contratos_resumido.sort_values(['CLIENTE', 'EQUIPAMENTO'], inplace=True)
        except:
            pass

        return df_saldo_atual_contratos_resumido


    def create_df_saidas_contratos(df):
        df_saldo_atual_contratos = df
        df_saldo_atual_contratos = df_saldo_atual_contratos[(df_saldo_atual_contratos['FLUXO'] == 'CONTRATO') & (df_saldo_atual_contratos['ENDEREÇO'] == 'LAB')]

        return df_saldo_atual_contratos


    def create_df_saidas_contratos_resumido(df):
        df = df
        df.loc[df['CLIENTE'].str.startswith('COBRA'), 'CLIENTE'] = 'COBRA'
        df.loc[df['CLIENTE'].str.startswith('BB'), 'CLIENTE'] = 'COBRA'

        df = df.groupby(['CLIENTE', 'EQUIPAMENTO'])[['SERIAL']].count().reset_index()
        df = df.rename(columns={'SERIAL':'QUANTIDADE'})
        try:
            df = df.sort_values([['CLIENTE', 'EQUIPAMENTO']])
        except:
            pass

        return df


    def create_fig_criticos():
        df = st.session_state['saldo_atual_contratos_selecao'][~st.session_state['saldo_atual_contratos_selecao']['% DO SLA'].isna()].copy()
        df['CAIXA'] = df['CAIXA'].astype('str')
        df['CAIXA'] = "ㅤ" + df['CAIXA']
        df['ENTRADA FILA'] = df['ENTRADA FILA'].astype('str')
        df['RÓTULO'] = df['CLIENTE'] + ' - ' + df['ENDEREÇO'] + ' - ' + df['ENTRADA FILA'].str.replace('-','/').str.split(" ").str[0]
        df = df.groupby(['CAIXA', 'RÓTULO', '% DO SLA'])['SERIAL'].count().reset_index().sort_values('% DO SLA', ascending=True).tail(10)
        
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
                        range_color=[0,1])
        
        return fig


    def create_fig_status(df):
        df = df.groupby(['STATUS'])[['SERIAL']].count().reset_index()

        fig = px.pie(df,
                                    names='STATUS',
                                    values='SERIAL',
                                    color='STATUS',
                                    hole=0.4,
                                    color_discrete_map={'RÁPIDO':'#008000',
                                                    'MÉDIO':'#32CD32',
                                                    'LENTO':'#FFD700',
                                                    'CRÍTICO':'#FF8C00',
                                                    'SLA ESTOURADO':'#8B0000'},
                                category_orders={'STATUS':['RÁPIDO', 'MÉDIO', 'LENTO', 'CRÍTICO', 'SLA ESTOURADO']})
        fig.update_traces(textinfo='value+percent')

        return fig


    def create_fig_status_saidas():
        df = st.session_state['saidas_contratos_selecao'].copy()
        df['SAÍDA FILA'] = df['SAÍDA FILA'].dt.strftime('%Y/%m')
        df = df.groupby(['SAÍDA FILA', 'STATUS'])['EQUIPAMENTO'].count().reset_index()
        df.rename(columns={'EQUIPAMENTO':'QUANTIDADE'}, inplace=True)
        
        fig = px.bar(df,
                     x='SAÍDA FILA',
                     y='QUANTIDADE',
                     color='STATUS',
                     color_discrete_map={
                        'RÁPIDO':'#008000',
                        'MÉDIO':'#32CD32',
                        'LENTO':'#FFD700',
                        'CRÍTICO':'#FF8C00',
                        'SLA ESTOURADO':'#8B0000'
                     },
                     orientation='v',
                     barmode='group',
                     text='QUANTIDADE',
                     category_orders={'STATUS':['RÁPIDO', 'MÉDIO', 'LENTO', 'CRÍTICO', 'SLA ESTOURADO']})
        
        fig.update_traces(textposition='inside',
                          orientation='v')
        
        fig.update_layout(yaxis_title=None,
                          xaxis_title=None,
                          yaxis_visible=False)
        
        return fig


    def create_fig_volume_fila(rows):
        df = df_saldo_atual_contratos_resumido.iloc[rows][['CLIENTE',
                                                           'EQUIPAMENTO',
                                                           'QTD FILA']].groupby(
                                                                ['CLIENTE'])['QTD FILA'].sum(
                                                           ).reset_index().sort_values(['QTD FILA'], ascending=False).head(5)

        fig = px.bar(df,
                     x='CLIENTE',
                     y='QTD FILA',
                     color_discrete_sequence=['#13399A'],
                     orientation='v',
                     text='QTD FILA')
        
        fig.update_traces(textposition='inside',
                          orientation='v')
      
        fig.update_layout(yaxis_title=None,
                          xaxis_title=None,
                          yaxis_visible=False)

        return fig
    

    def create_fig_volume_os(rows):
        df = df_saldo_atual_contratos_resumido.iloc[rows][['CLIENTE',
                                                           'EQUIPAMENTO',
                                                           'QTD OS']].groupby(
                                                                ['CLIENTE'])['QTD OS'].sum(
                                                           ).reset_index().sort_values(['QTD OS'], ascending=False).head(5)

        fig = px.bar(df,
                     x='CLIENTE',
                     y='QTD OS',
                     color_discrete_sequence=['#E8C406'],
                     orientation='v',
                     text='QTD OS')
        
        fig.update_traces(textposition='inside',
                          orientation='v')
      
        fig.update_layout(yaxis_title=None,
                          xaxis_title=None,
                          yaxis_visible=False)
        
        return fig


    def html_saldo_contrato():
        df = st.session_state['df_saldo_atual_contratos_resumido'].copy()
        df.loc[df['CLIENTE'].str.startswith('COBRA'), 'CLIENTE'] = 'COBRA'
        df.loc[df['CLIENTE'].str.startswith('BB'), 'CLIENTE'] = 'COBRA'
        df.loc[df['CLIENTE'].str.startswith('MERCADO'), 'CLIENTE'] = 'MERCADO PAGO'

        df.loc[df['EQUIPAMENTO'].str.contains('PPC930'), 'EQUIPAMENTO'] = 'PPC930'
        
        df = df.groupby(['CLIENTE', 'EQUIPAMENTO'])[['QTD OS', 'QTD FILA']].sum().reset_index()

        html_contratos = df[['CLIENTE', 'EQUIPAMENTO', 'QTD OS', 'QTD FILA']].to_html(index=False, index_names=False, justify='left', na_rep='')
        html_contratos = html_contratos.replace('<table border="1" class="dataframe">',
                                            '<style>\ntable {\n  border-collapse: collapse;\n  width: 100%;\n}\n\nth, td {\n  text-align: center;\n  padding-top: 2px;\n  padding-bottom: 1px;\n  padding-left: 8px;\n  padding-right: 8px;\n}\n\ntr:nth-child(even) {\n  background-color: #DCDCDC;\n}\n\ntable, th, td {\n  border: 2px solid black;\n  border-collapse: collapse;\n}\n</style>\n<table border="1" class="dataframe">')
        
        return html_contratos


    @st.experimental_dialog("Filtros de Saldo", width='large')
    def open_dialog_filtros_saldo():
        df = st.session_state['historico_fila']
        df = df[(df['FLUXO'] == 'CONTRATO') & (df['ENDEREÇO'] != 'LAB')]

        df2 = st.session_state['df_saldo_atual_contratos_resumido']

        fr1c1, fr1c2 = st.columns(2)
        fr2c1, fr2c2 = st.columns(2)
        fr3c1, fr3c2 = st.columns(2)
        fr4c1, fr4c2 = st.columns(2)
        fr5c1, fr5c2 = st.columns(2)

        ft_cliente = fr1c1.multiselect('CLIENTE', df2['CLIENTE'].unique())
        ft_equip = fr1c2.multiselect('EQUIPAMENTO', df2['EQUIPAMENTO'].unique())

        ft_os = fr2c1.multiselect('NUM OS', df['NUM OS'].unique())
        ft_ns = fr2c2.multiselect('SERIAL', df['SERIAL'].unique())

        ft_end = fr3c1.multiselect('ENDEREÇO', df['ENDEREÇO'].unique())
        ft_caixa = fr3c2.multiselect('CAIXA', df['CAIXA'].unique())

        ft_dtger_min = fr4c1.date_input('DATA ENTRADA GERFLOOR', value=min(df['ENTRADA GERFLOOR']), format='DD/MM/YYYY')
        ft_dtger_max = fr4c2.date_input('', value=max(df['ENTRADA GERFLOOR']), format='DD/MM/YYYY')

        ft_dtfila_min = fr5c1.date_input('DATA ENTRADA FILA', value=min(df['ENTRADA FILA']), format='DD/MM/YYYY')
        ft_dtfila_max = fr5c2.date_input(' ', value=max(df['ENTRADA FILA']), format='DD/MM/YYYY')

        if st.button('APLICAR FILTROS', use_container_width=True):
            if ft_cliente:
                df = df[df['CLIENTE'].isin(ft_cliente)]
            if ft_equip:
                df = df[df['EQUIPAMENTO'].isin(ft_equip)]
            if ft_os:
                df = df[df['NUM OS'].isin(ft_os)]
            if ft_ns:
                df = df[df['SERIAL'].isin(ft_ns)]
            if ft_end:
                df = df[df['ENDEREÇO'].isin(ft_end)]
            if ft_caixa:
                df = df[df['CAIXA'].isin(ft_caixa)]

            df = df[(df['ENTRADA GERFLOOR'] >= pd.to_datetime(ft_dtger_min)) & (df['ENTRADA GERFLOOR'] <= pd.to_datetime(ft_dtger_max))]
            df = df[(df['ENTRADA FILA'] >= pd.to_datetime(ft_dtfila_min)) & (df['ENTRADA FILA'] <= pd.to_datetime(ft_dtfila_max))]

            st.session_state['df_saldo_atual_contratos'] = create_df_saldo_contratos(df)
            df_sacr = create_df_saldo_contratos_resumido(st.session_state['df_saldo_atual_contratos'])

            if ft_cliente:
                df_sacr = df_sacr[df_sacr['CLIENTE'].isin(ft_cliente)]
            if ft_equip:
                df_sacr = df_sacr[df_sacr['EQUIPAMENTO'].isin(ft_equip)]

            st.session_state['df_saldo_atual_contratos_resumido'] = df_sacr

            st.rerun()


    @st.experimental_dialog("Filtros de Saída", width='large')
    def open_dialog_filtros_saida():
        df = st.session_state['historico_fila']
        df = df[(df['FLUXO'] == 'CONTRATO') & (df['ENDEREÇO'] == 'LAB')]

        df2 = st.session_state['df_saidas_contratos_resumido']

        fr1c1, fr1c2 = st.columns(2)
        fr2c1, fr2c2 = st.columns(2)
        fr3c1, fr3c2 = st.columns(2)
        fr4c1, fr4c2 = st.columns(2)
        fr5c1, fr5c2 = st.columns(2)

        ft_cliente = fr1c1.multiselect('CLIENTE', df2['CLIENTE'].unique())
        ft_equip = fr1c2.multiselect('EQUIPAMENTO', df2['EQUIPAMENTO'].unique())

        ft_os = fr2c1.multiselect('NUM OS', df['NUM OS'].unique())
        ft_ns = fr2c2.multiselect('SERIAL', df['SERIAL'].unique())

        ft_end = fr3c1.multiselect('ENDEREÇO', df['ENDEREÇO'].unique())
        ft_caixa = fr3c2.multiselect('CAIXA', df['CAIXA'].unique())

        ft_dtger_min = fr4c1.date_input('DATA ENTRADA GERFLOOR', value=min(df['ENTRADA GERFLOOR']), format='DD/MM/YYYY')
        ft_dtger_max = fr4c2.date_input('', value=max(df['ENTRADA GERFLOOR']), format='DD/MM/YYYY')

        ft_dtfila_min = fr5c1.date_input('DATA ENTRADA FILA', value=min(df['ENTRADA FILA']), format='DD/MM/YYYY')
        ft_dtfila_max = fr5c2.date_input(' ', value=max(df['ENTRADA FILA']), format='DD/MM/YYYY')

        ft_dtsfila_min = fr5c1.date_input('DATA SAÍDA FILA', value=min(df['SAÍDA FILA']), format='DD/MM/YYYY')
        ft_dtsfila_max = fr5c2.date_input('  ', value=max(df['SAÍDA FILA']), format='DD/MM/YYYY')

        if st.button('APLICAR FILTROS', use_container_width=True):
            if ft_cliente:
                df = df[df['CLIENTE'].isin(ft_cliente)]
            if ft_equip:
                df = df[df['EQUIPAMENTO'].isin(ft_equip)]
            if ft_os:
                df = df[df['NUM OS'].isin(ft_os)]
            if ft_ns:
                df = df[df['SERIAL'].isin(ft_ns)]
            if ft_end:
                df = df[df['ENDEREÇO'].isin(ft_end)]
            if ft_caixa:
                df = df[df['CAIXA'].isin(ft_caixa)]

            df = df[(df['ENTRADA GERFLOOR'] >= pd.to_datetime(ft_dtger_min)) & (df['ENTRADA GERFLOOR'] <= pd.to_datetime(ft_dtger_max))]
            df = df[(df['ENTRADA FILA'] >= pd.to_datetime(ft_dtfila_min)) & (df['ENTRADA FILA'] <= pd.to_datetime(ft_dtfila_max))]
            df = df[(df['SAÍDA FILA'] >= pd.to_datetime(ft_dtsfila_min)) & (df['SAÍDA FILA'] <= pd.to_datetime(ft_dtsfila_max))]

            st.session_state['df_saidas_contratos'] = create_df_saidas_contratos(df)
            df_scr = create_df_saidas_contratos_resumido(st.session_state['df_saidas_contratos'])

            if ft_cliente:
                df_scr = df_scr[df_scr['CLIENTE'].isin(ft_cliente)]
            if ft_equip:
                df_scr = df_scr[df_scr['EQUIPAMENTO'].isin(ft_equip)]

            st.session_state['df_saida_contratos_resumido'] = df_scr

            st.rerun()


    if 'historico_fila' not in st.session_state:
        st.session_state['historico_fila'] = create_df_historico_movimentações()
        historico_fila = st.session_state['historico_fila']
    else:
        historico_fila = st.session_state['historico_fila']

    st.sidebar.header('')
    st.sidebar.title('AÇÕES')

    tabs_saldo, tabs_saida, tabs_geral = st.tabs(['Saldo', 'Saídas', 'Tabela Geral'])

    tabs_saldo.title('Saldo de Contratos')
    r0c1, r0c2, r0c3, r0c4 = tabs_saldo.columns(4)
    tabs_saldo.write('')
    r1c1, r1c2 = tabs_saldo.columns(2, gap='large')
    r2c1, r2c2 = tabs_saldo.columns([0.7, 0.3], gap='large')
    tabs_saldo.write('')
    r3c1, r3c2 = tabs_saldo.columns(2, gap='large')

    if 'df_saldo_atual_contratos' not in st.session_state or 'df_saldo_atual_contratos_resumido' not in st.session_state:
        st.session_state['df_saldo_atual_contratos'] = create_df_saldo_contratos(historico_fila)
        st.session_state['df_saldo_atual_contratos_resumido'] = create_df_saldo_contratos_resumido(st.session_state['df_saldo_atual_contratos'])
    
        df_saldo_atual_contratos = st.session_state['df_saldo_atual_contratos']
        df_saldo_atual_contratos_resumido = st.session_state['df_saldo_atual_contratos_resumido']
    else:
        df_saldo_atual_contratos = st.session_state['df_saldo_atual_contratos']
        df_saldo_atual_contratos_resumido = st.session_state['df_saldo_atual_contratos_resumido']
        
    st.sidebar.download_button('BAIXAR RESUMO', html_saldo_contrato(), use_container_width=True, file_name='Contratos.html')

    r1c1.write('Resumo de saldo de equipamentos.')
    saldo_atual_contratos = r1c1.dataframe(
        df_saldo_atual_contratos_resumido[['CLIENTE', 'EQUIPAMENTO', 'QTD OS', 'QTD FILA']],
        hide_index=True,
        use_container_width=True,
        on_select='rerun',
        column_config={'SERIAL':st.column_config.NumberColumn('QTD FILA')})
    
    if saldo_atual_contratos.selection.rows:
        df_saldo_atual_contratos_resumido['CONCATENADO'] = df_saldo_atual_contratos_resumido['CLIENTE'] + df_saldo_atual_contratos_resumido['EQUIPAMENTO']
        df_saldo_atual_contratos['CONCATENADO'] = df_saldo_atual_contratos['CLIENTE'] + df_saldo_atual_contratos['EQUIPAMENTO']
        filtro_saldo = list(df_saldo_atual_contratos_resumido.iloc[saldo_atual_contratos.selection.rows]['CONCATENADO'])
        saldo_atual_contratos_selecao = df_saldo_atual_contratos[df_saldo_atual_contratos['CONCATENADO'].isin(filtro_saldo)]
        st.session_state['saldo_atual_contratos_selecao'] = saldo_atual_contratos_selecao
        r0c1.metric('Total de equipamentos (seleção)',
                    '{:,}'.format(sum(df_saldo_atual_contratos_resumido.iloc[saldo_atual_contratos.selection.rows]['QTD FILA']) + 
                    sum(df_saldo_atual_contratos_resumido.iloc[saldo_atual_contratos.selection.rows]['QTD OS'])).replace(',', '.'))
        r0c2.metric('Equipamentos aguardando OS (seleção)',
                    '{:,}'.format(sum(df_saldo_atual_contratos_resumido.iloc[saldo_atual_contratos.selection.rows]['QTD OS'])).replace(',', '.'))
        r0c3.metric('Equipamentos em fila (seleção)',
                    '{:,}'.format(sum(df_saldo_atual_contratos_resumido.iloc[saldo_atual_contratos.selection.rows]['QTD FILA'])).replace(',', '.'))
    else:
        r0c1.metric('Total de equipamentos',
                    '{:,}'.format(sum(df_saldo_atual_contratos_resumido['QTD FILA']) +
                    sum(df_saldo_atual_contratos_resumido['QTD OS'])).replace(',', '.'))
        r0c2.metric('Equipamentos aguardando OS',
                    '{:,}'.format(sum(df_saldo_atual_contratos_resumido['QTD OS'])).replace(',', '.'))
        r0c3.metric('Equipamentos em fila',
                    '{:,}'.format(sum(df_saldo_atual_contratos_resumido['QTD FILA'])).replace(',', '.'))
        
    if r0c4.button('FILTROS DE SALDO', use_container_width=True):
        open_dialog_filtros_saldo()


    if 'saldo_atual_contratos_selecao' in st.session_state and saldo_atual_contratos.selection.rows:
        if len(st.session_state['saldo_atual_contratos_selecao']) > 0:
            r1c2.write('Classificação dos equipamentos no fila de acordo com % do SLA.')
            r1c2.plotly_chart(create_fig_criticos())

            r2c1.write('Saldo detalhado de equipamentos no fila.')
            r2c1.dataframe(saldo_atual_contratos_selecao[[
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
            r2c2.write('Status dos equipamentos em relação a entrega do SLA.')
            r2c2.plotly_chart(create_fig_status(st.session_state['saldo_atual_contratos_selecao']))

            r3c2.write('Maiores volumetrias em fila.')
            r3c2.plotly_chart(create_fig_volume_fila(saldo_atual_contratos.selection.rows))

        if sum(df_saldo_atual_contratos_resumido.iloc[saldo_atual_contratos.selection.rows]['QTD OS']) > 0: 
            r3c1.write('Maiores volumetrias aguardando abertura de OS.')
            r3c1.plotly_chart(create_fig_volume_os(saldo_atual_contratos.selection.rows))

    tabs_saida.title('Saída de Equipamentos')
    t2r0c1, t2r0c2, t2r0c3, t2r0c4 = tabs_saida.columns(4)
    tabs_saida.write('')
    t2r1c1, t2r1c2 = tabs_saida.columns(2, gap='large')
    t2r2c1 = tabs_saida.container()
    tabs_saida.write('')
    t2r3c1 = tabs_saida.container()

    if 'df_saidas_contratos' not in st.session_state or 'df_saidas_contratos_resumido' not in st.session_state:
        st.session_state['df_saidas_contratos'] = create_df_saidas_contratos(historico_fila)
        st.session_state['df_saidas_contratos_resumido'] = create_df_saidas_contratos_resumido(st.session_state['df_saidas_contratos'])
    
        df_saidas_contratos = st.session_state['df_saidas_contratos']
        df_saidas_contratos_resumido = st.session_state['df_saidas_contratos_resumido']
    else:
        df_saidas_contratos = st.session_state['df_saidas_contratos']
        df_saidas_contratos_resumido = st.session_state['df_saidas_contratos_resumido']

    t2r1c1.write('Resumo de equipamentos enviados ao laboratório.')
    saidas_contratos = t2r1c1.dataframe(df_saidas_contratos_resumido[['CLIENTE', 'EQUIPAMENTO', 'QUANTIDADE']],
                      hide_index=True,
                      use_container_width=True,
                      on_select='rerun')
    
    if saidas_contratos.selection.rows:
        df_saidas_contratos_resumido['CONCATENADO'] = df_saidas_contratos_resumido['CLIENTE'] + df_saidas_contratos_resumido['EQUIPAMENTO']
        df_saidas_contratos['CONCATENADO'] = df_saidas_contratos['CLIENTE'] + df_saidas_contratos['EQUIPAMENTO']
        filtro_saldo = list(df_saidas_contratos_resumido.iloc[saidas_contratos.selection.rows]['CONCATENADO'])
        saidas_contratos_selecao = df_saidas_contratos[df_saidas_contratos['CONCATENADO'].isin(filtro_saldo)]
        st.session_state['saidas_contratos_selecao'] = saidas_contratos_selecao

        t2r0c1.metric('Total de saídas (seleção)', '{:,}'.format(len(saidas_contratos_selecao['SERIAL'])).replace(',','.'))
        if len(saidas_contratos_selecao[saidas_contratos_selecao['SAÍDA FILA'] >= datetime.today()-timedelta(hours=datetime.today().hour+1)]) > 0:
            filtro_ontem = ((saidas_contratos_selecao['SAÍDA FILA'] >= datetime.today()-timedelta(days=1, hours=datetime.today().hour, minutes=datetime.today().minute)) &
                            (saidas_contratos_selecao['SAÍDA FILA'] <= datetime.today()-timedelta(hours=datetime.today().hour+1, minutes=datetime.today().minute)))
            try:
                t2r0c2.metric('Saídas do dia (seleção)', '{:,}'.format(len(saidas_contratos_selecao[saidas_contratos_selecao['SAÍDA FILA'] >= datetime.today()-timedelta(hours=datetime.today().hour+1)])).replace(',','.'),
                            delta='{:.2%}'.format(((len(saidas_contratos_selecao[saidas_contratos_selecao['SAÍDA FILA'] >= datetime.today()-timedelta(hours=datetime.today().hour+1)])) - len(saidas_contratos_selecao[filtro_ontem])) / len(saidas_contratos_selecao[filtro_ontem])))
            except:
                t2r0c2.metric('Saídas do dia (seleção)', '{:,}'.format(len(saidas_contratos_selecao[saidas_contratos_selecao['SAÍDA FILA'] >= datetime.today()-timedelta(hours=datetime.today().hour+1)])).replace(',','.'),
                            delta='{:.2%}'.format(0))
        else: t2r0c2.metric('Saídas do dia (seleção)', 0)
    else:
        t2r0c1.metric('Total de saídas', '{:,}'.format(sum(df_saidas_contratos_resumido['QUANTIDADE'])).replace(',','.'))
        if len(df_saidas_contratos[df_saidas_contratos['SAÍDA FILA'] >= datetime.today()-timedelta(hours=datetime.today().hour+1)]) > 0:
            filtro_ontem = ((df_saidas_contratos['SAÍDA FILA'] >= datetime.today()-timedelta(days=1, hours=datetime.today().hour, minutes=datetime.today().minute)) &
                            (df_saidas_contratos['SAÍDA FILA'] <= datetime.today()-timedelta(hours=datetime.today().hour, minutes=datetime.today().minute)))
            try:
                t2r0c2.metric('Saídas do dia', '{:,}'.format(len(df_saidas_contratos[df_saidas_contratos['SAÍDA FILA'] >= datetime.today()-timedelta(hours=datetime.today().hour+1)])).replace(',','.'),
                            delta='{:.2%}'.format(((len(df_saidas_contratos[df_saidas_contratos['SAÍDA FILA'] >= datetime.today()-timedelta(hours=datetime.today().hour+1)])) - len(df_saidas_contratos[filtro_ontem])) / len(df_saidas_contratos[filtro_ontem])))
            except:
                t2r0c2.metric('Saídas do dia', '{:,}'.format(len(df_saidas_contratos[df_saidas_contratos['SAÍDA FILA'] >= datetime.today()-timedelta(hours=datetime.today().hour+1)])).replace(',','.'),
                            delta='{:.2%}'.format(0))
        else: t2r0c2.metric('Saídas do dia', 0)

    if t2r0c4.button('FILTROS DE SAÍDA', use_container_width=True):
        open_dialog_filtros_saida()

    if 'saidas_contratos_selecao' in st.session_state and saidas_contratos.selection.rows:
        t2r1c2.write('Status dos equipamentos entregues em relação ao SLA.')
        t2r1c2.plotly_chart(create_fig_status(st.session_state['saidas_contratos_selecao']))

        t2r2c1.write('Histórico detalhado de equipamentos entregues ao laboratório.')
        t2r2c1.dataframe(st.session_state['saidas_contratos_selecao'][['CAIXA', 'SERIAL', 'CLIENTE', 'EQUIPAMENTO',
                                                                       'NUM OS', 'ENTRADA GERFLOOR', 'ENTRADA FILA',
                                                                       'SAÍDA FILA', 'AGING TOTAL', 'AGING FILA', 'STATUS']].sort_values(['SAÍDA FILA']),
                         hide_index=True,
                         use_container_width=True,
                         column_config={
                             'ENTRADA GERFLOOR': st.column_config.DateColumn('ENTRADA GERFLOOR', format='DD/MM/YYYY'),
                             'ENTRADA FILA': st.column_config.DateColumn('ENTRADA FILA', format='DD/MM/YYYY HH:mm:ss'),
                             'SAÍDA FILA': st.column_config.DateColumn('SAÍDA FILA', format='DD/MM/YYYY HH:mm:ss')
                         })
        
        t2r3c1.write('Distribuição do status dos equipamentos entregues ao longo dos meses.')
        t2r3c1.plotly_chart(create_fig_status_saidas())

    tabs_geral.dataframe(st.session_state['historico_fila'][st.session_state['historico_fila']['FLUXO'] == 'CONTRATO'])
