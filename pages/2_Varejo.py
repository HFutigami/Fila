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
    st.sidebar.page_link('pages/1_Contratos.py', label='CONTRATO')
    st.sidebar.page_link('pages/2_Varejo.py', label='VAREJO', disabled=True)
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

        historico_fila['FLUXO'] = historico_fila['FLUXO'].str.replace('000001', 'CONTRATO')
        historico_fila['FLUXO'] = historico_fila['FLUXO'].str.replace('000002', 'VAREJO')
        historico_fila['FLUXO'] = historico_fila['FLUXO'].str.replace('000003', 'VAREJO')
        historico_fila['FLUXO'] = historico_fila['FLUXO'].str.replace('000004', 'OS INTERNA')
        
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


    def create_df_saldo_varejo(df):
        df_saldo_atual_varejo = df.copy()
        df_saldo_atual_varejo = df_saldo_atual_varejo[(~df_saldo_atual_varejo['FLUXO'].isin(['CONTRATO', 'OS INTERNA'])) & (df_saldo_atual_varejo['ENDEREÇO'] != 'LAB')]

        return df_saldo_atual_varejo
    

    def create_df_saldo_varejo_resumido(df):
        
        df_saldo_atual_varejo_resumido = df
        df_saldo_atual_varejo_resumido = df_saldo_atual_varejo_resumido.sort_values(['ENTRADA FILA']).drop_duplicates(['SERIAL', 'CLIENTE', 'NUM OS'], keep='last')
        df_saldo_atual_varejo_resumido = df_saldo_atual_varejo_resumido.groupby(['CLIENTE', 'EQUIPAMENTO'])[['SERIAL']].count().reset_index()
        df_saldo_atual_varejo_resumido.rename(columns={'SERIAL':'QUANTIDADE'}, inplace=True)
        df_saldo_atual_varejo_resumido = df_saldo_atual_varejo_resumido[['CLIENTE', 'EQUIPAMENTO', 'QUANTIDADE']]
        try:
            df_saldo_atual_varejo_resumido.sort_values(['CLIENTE', 'EQUIPAMENTO'], inplace=True)
        except:
            pass

        return df_saldo_atual_varejo_resumido


    def create_df_saidas_varejo(df):
        df_saldo_atual_varejo = df.copy()
        df_saldo_atual_varejo = df_saldo_atual_varejo[(~df_saldo_atual_varejo['FLUXO'].isin(['CONTRATO', 'OS INTERNA'])) & (df_saldo_atual_varejo['ENDEREÇO'] == 'LAB')]

        return df_saldo_atual_varejo


    def create_df_saidas_varejo_resumido(df):
        df = df.groupby(['CLIENTE', 'EQUIPAMENTO'])[['SERIAL']].count().reset_index().copy()
        df = df.rename(columns={'SERIAL':'QUANTIDADE'})
        try:
            df = df.sort_values([['CLIENTE', 'EQUIPAMENTO']])
        except:
            pass

        return df


    def create_df_varejo_liberado(data_liberacao):
        try:
            auth = AuthenticationContext(sharepoint_fila_url)
            auth.acquire_token_for_user(sharepoint_user, sharepoint_password)
            ctx = ClientContext(saldo_fila_url, auth)
            web = ctx.web
            ctx.execute_query()
    
            file_response = File.open_binary(ctx, varejo_liberado_url + str(data_liberacao) + ".xlsx")
            bytes_file_obj = io.BytesIO()
            bytes_file_obj.write(file_response.content)
            bytes_file_obj.seek(0)
            df = pd.read_excel(bytes_file_obj,
                                sheet_name='LAB. - SEPARAÇÃO',
                                dtype='str')
            processos_varejo = df
            processos_varejo = processos_varejo[['Nr Serie', 'Num OS', 'Produto_1', 'Client Final', 'Dt Aber. OS']]
            processos_varejo.rename(columns={'Nr Serie':'SERIAL', 'Num OS':'NUM OS'}, inplace=True)
            processos_varejo.set_index(['SERIAL', 'NUM OS'], inplace=True)
        
            varejo_liberado = historico_fila.join(processos_varejo,
                                                on=['SERIAL', 'NUM OS'],
                                                how='right')
            varejo_liberado.loc[varejo_liberado['ENDEREÇO'].isna(), 'CLIENTE'] = varejo_liberado.loc[varejo_liberado['ENDEREÇO'].isna(), 'Client Final']
            varejo_liberado.loc[varejo_liberado['ENDEREÇO'].isna(), 'EQUIPAMENTO'] = varejo_liberado.loc[varejo_liberado['ENDEREÇO'].isna(), 'Produto_1']
            varejo_liberado.loc[varejo_liberado['ENDEREÇO'].isna(), 'ENTRADA GERFLOOR'] = varejo_liberado.loc[varejo_liberado['ENDEREÇO'].isna(), 'Dt Aber. OS']
            varejo_liberado.drop(columns=['Produto_1',
                                          'Dt Aber. OS',
                                          'Client Final',
                                          'ULTIMA DATA',
                                          'AGING TOTAL',
                                          'AGING FILA',
                                          'PRAZO',
                                          'FLUXO',
                                          '% DO SLA',
                                          'STATUS'], inplace=True)
            varejo_liberado.sort_values('ENDEREÇO', inplace=True)

            st.session_state['data_liberação'] = data_liberacao
            st.session_state['varejo_liberado'] = varejo_liberado
        except:
            varejo_liberado = pd.DataFrame(columns=['!'])
            st.session_state['data_liberação'] = data_liberacao
            st.session_state['varejo_liberado'] = varejo_liberado
            
        return varejo_liberado


    def create_df_varejo_liberado_resumido(df):
        df = df.groupby(['NUM OS', 'CLIENTE', 'EQUIPAMENTO'])[['SERIAL']].count().reset_index().copy()
        df.rename(columns={'SERIAL':'QUANTIDADE'}, inplace=True)

        return df


    def create_df_terceiros_varejo(df):
        df_saldo_atual_contratos = df.copy()
        df_saldo_atual_contratos = df_saldo_atual_contratos[
            (df_saldo_atual_contratos['FLUXO'] == 'VAREJO') & (df_saldo_atual_contratos['ENDEREÇO'].isin(
                ['EQUIPE TECNICA', 'QUALIDADE', 'RETRIAGEM', 'GESTAO DE ATIVOS']))]
        df_saldo_atual_contratos.rename(columns={'ENDEREÇO': 'TERCEIROS'}, inplace=True)

        return df_saldo_atual_contratos


    def create_df_terceiros_varejo_resumido(df):

        df = df.groupby(['TERCEIROS'])[['SERIAL']].count().reset_index()
        df = df.rename(columns={'SERIAL': 'QUANTIDADE'})
        try:
            df = df.sort_values([['TERCEIROS', 'EQUIPAMENTO']])
        except:
            pass

        return df


    def create_fig_criticos(df):
        df['CAIXA'] = df['CAIXA'].astype('str')
        df['CAIXA'] = "ㅤ" + df['CAIXA']
        df['SERIAL'] = df['SERIAL'].astype('str')
        df['SERIAL'] = "ㅤ" + df['SERIAL']
        df['ENTRADA FILA'] = df['ENTRADA FILA'].astype('str')
        try:
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

        except:
            df['RÓTULO'] = df['CLIENTE'] + ' - ' + df['TERCEIROS'] + ' - ' + \
                           df['ENTRADA FILA'].str.replace('-', '/').str.split(" ").str[0]

            df = df.groupby(['SERIAL', 'RÓTULO', '% DO SLA'])['CAIXA'].count().reset_index().sort_values('% DO SLA',
                                                                                                         ascending=True).tail(
                10)

            fig = px.bar(df,
                         x='% DO SLA',
                         y='SERIAL',
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
        df = st.session_state['saidas_varejo_selecao'].copy()
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


    def create_fig_status_liberacao():
        df = st.session_state['varejo_liberado_selecao'].copy()
        df['STATUS'] = ''
        df.loc[df['ENDEREÇO'].isna(), 'STATUS'] = 'NÃO LOCALIZADO'
        df.loc[df['ENDEREÇO'] == 'LAB', 'STATUS'] = 'EM LABORATÓRIO'
        df.loc[~df['STATUS'].isin(['EM LABORATÓRIO', 'NÃO LOCALIZADO']), 'STATUS'] = 'LOCALIZADO'
        df = df.groupby('STATUS')['SERIAL'].count().reset_index()
        df.rename(columns={'SERIAL':'QUANTIDADE'}, inplace=True)

        fig = px.pie(data_frame=df,
                    values='QUANTIDADE',
                    names='STATUS',
                    color='STATUS',
                    hole=0.4,
                    color_discrete_map={
                        'NÃO LOCALIZADO':'#8B0000',
                        'LOCALIZADO':'#13399A',
                        'EM LABORATÓRIO':'#008000'
                    }).update_traces(textinfo='value+percent')
        
        return fig
        


    def create_fig_volume_fila(rows):
        df = df_saldo_atual_varejo_resumido.iloc[rows][['CLIENTE',
                                                           'EQUIPAMENTO',
                                                           'QUANTIDADE']].groupby(
                                                                ['CLIENTE'])['QUANTIDADE'].sum(
                                                           ).reset_index().sort_values(['QUANTIDADE'], ascending=False).head(10)

        fig = px.bar(df,
                     x='CLIENTE',
                     y='QUANTIDADE',
                     color_discrete_sequence=['#13399A'],
                     orientation='v',
                     text='QUANTIDADE')
        
        fig.update_traces(textposition='inside',
                          orientation='v')
      
        fig.update_layout(yaxis_title=None,
                          xaxis_title=None,
                          yaxis_visible=False)

        return fig

    
    def html_varejo():
        df = st.session_state['varejo_liberado'].copy()
        varejo_compactado = df.groupby(['NUM OS', 'CLIENTE', 'ENDEREÇO'])['SERIAL'].count().reset_index().copy()
        varejo_compactado['SERIAL'] = varejo_compactado['SERIAL'].apply(lambda x: "TOTAL: " + str(x))
            
        df = pd.concat([varejo_compactado, df])
        df['SEPARADO'] = ''
        df = df[['NUM OS', 'SERIAL',
            'CAIXA', 'CLIENTE',
            'EQUIPAMENTO', 'ENDEREÇO',
            'SEPARADO', 'GARANTIA']].sort_values(['ENDEREÇO', 'NUM OS', 'SERIAL'])
        df.loc[df['SERIAL'].str.startswith('TOTAL'), ['NUM OS', 'CAIXA', 'CLIENTE', 'EQUIPAMENTO', 'ENDEREÇO', 'SEPARADO', 'GARANTIA']] = ''
        
        html_content = df.to_html(index=False, index_names=False, justify='left', na_rep='')
        html_content = html_content.replace('<table border="1" class="dataframe">',
                                          '<style>\ntable {\n  border-collapse: collapse;\n  width: 100%;\n}\n\nth, td {\n  text-align: center;\n  padding-top: 2px;\n  padding-bottom: 1px;\n  padding-left: 8px;\n  padding-right: 8px;\n}\n\ntr:nth-child(even) {\n  background-color: #DCDCDC;\n}\n\ntable, th, td {\n  border: 2px solid black;\n  border-collapse: collapse;\n}\n</style>\n<table border="1" class="dataframe">')
        return html_content
    

    @st.experimental_dialog("Filtros de Saldo", width='large')
    def open_dialog_filtros_saldo():
        df = st.session_state['historico_fila']
        df = df[(~df['FLUXO'].isin(['CONTRATO', 'OS INTERNA'])) & (df['ENDEREÇO'] != 'LAB')]

        df2 = st.session_state['df_saldo_atual_varejo_resumido']

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

            st.session_state['df_saldo_atual_varejo'] = create_df_saldo_varejo(df)
            df_sacr = create_df_saldo_varejo_resumido(st.session_state['df_saldo_atual_varejo'])

            if ft_cliente:
                df_sacr = df_sacr[df_sacr['CLIENTE'].isin(ft_cliente)]
            if ft_equip:
                df_sacr = df_sacr[df_sacr['EQUIPAMENTO'].isin(ft_equip)]

            st.session_state['df_saldo_atual_varejo_resumido'] = df_sacr

            st.rerun()


    @st.experimental_dialog("Filtros de Saída", width='large')
    def open_dialog_filtros_saida():
        df = st.session_state['historico_fila']
        df = df[(~df['FLUXO'].isin(['CONTRATO', 'OS INTERNA'])) & (df['ENDEREÇO'] == 'LAB')]

        df2 = st.session_state['df_saidas_varejo_resumido']

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

            st.session_state['df_saidas_varejo'] = create_df_saidas_varejo(df)
            df_scr = create_df_saidas_varejo_resumido(st.session_state['df_saidas_varejo'])

            if ft_cliente:
                df_scr = df_scr[df_scr['CLIENTE'].isin(ft_cliente)]
            if ft_equip:
                df_scr = df_scr[df_scr['EQUIPAMENTO'].isin(ft_equip)]

            st.session_state['df_saldo_atual_varejo_resumido'] = df_scr

            st.rerun()


    if 'historico_fila' not in st.session_state:
        st.session_state['historico_fila'] = create_df_historico_movimentações()
        historico_fila = st.session_state['historico_fila']
    else:
        historico_fila = st.session_state['historico_fila']

    st.sidebar.header('')
    st.sidebar.title('AÇÕES')

    tabs_saldo, tabs_liberado, tabs_saida, tabs_terceiros, tabs_geral = st.tabs(['Saldo', 'Liberado', 'Saídas', 'Terceiros', 'Tabela Geral'])

    tabs_saldo.title('Saldo de Varejo')
    r0c1, r0c2, r0c3, r0c4 = tabs_saldo.columns(4, gap='large')
    tabs_saldo.write('')
    r1c1, r1c2 = tabs_saldo.columns(2, gap='large')
    r2c1, r2c2 = tabs_saldo.columns([0.7, 0.3], gap='large')
    tabs_saldo.write('')
    r3c1 = tabs_saldo.container()

    if 'df_saldo_atual_varejo' not in st.session_state or 'df_saldo_atual_varejo_resumido' not in st.session_state:
        st.session_state['df_saldo_atual_varejo'] = create_df_saldo_varejo(historico_fila)
        st.session_state['df_saldo_atual_varejo_resumido'] = create_df_saldo_varejo_resumido(st.session_state['df_saldo_atual_varejo'])
    
        df_saldo_atual_varejo = st.session_state['df_saldo_atual_varejo']
        df_saldo_atual_varejo_resumido = st.session_state['df_saldo_atual_varejo_resumido']
    else:
        df_saldo_atual_varejo = st.session_state['df_saldo_atual_varejo']
        df_saldo_atual_varejo_resumido = st.session_state['df_saldo_atual_varejo_resumido']

    r1c1.write('Resumo de saldo de equipamentos.')
    saldo_atual_varejo = r1c1.dataframe(
        df_saldo_atual_varejo_resumido[['CLIENTE', 'EQUIPAMENTO', 'QUANTIDADE']],
        hide_index=True,
        use_container_width=True,
        on_select='rerun',
        column_config={'SERIAL':st.column_config.NumberColumn('QUANTIDADE')})
    
    if saldo_atual_varejo.selection.rows:
        df_saldo_atual_varejo_resumido['CONCATENADO'] = df_saldo_atual_varejo_resumido['CLIENTE'] + df_saldo_atual_varejo_resumido['EQUIPAMENTO']
        df_saldo_atual_varejo['CONCATENADO'] = df_saldo_atual_varejo['CLIENTE'] + df_saldo_atual_varejo['EQUIPAMENTO']
        filtro_saldo = list(df_saldo_atual_varejo_resumido.iloc[saldo_atual_varejo.selection.rows]['CONCATENADO'])
        saldo_atual_varejo_selecao = df_saldo_atual_varejo[df_saldo_atual_varejo['CONCATENADO'].isin(filtro_saldo)]
        st.session_state['saldo_atual_varejo_selecao'] = saldo_atual_varejo_selecao
        r0c1.metric('Total de equipamentos (seleção)',
                    '{:,}'.format(sum(df_saldo_atual_varejo_resumido.iloc[saldo_atual_varejo.selection.rows]['QUANTIDADE'])).replace(',', '.'))
    else:
        r0c1.metric('Total de equipamentos',
                    '{:,}'.format(sum(df_saldo_atual_varejo_resumido['QUANTIDADE'])).replace(',', '.'))
       
    if r0c4.button('FILTROS DE SALDO', use_container_width=True):
        open_dialog_filtros_saldo()


    if 'saldo_atual_varejo_selecao' in st.session_state and saldo_atual_varejo.selection.rows:
        if len(st.session_state['saldo_atual_varejo_selecao']) > 0:
            r1c2.write('Classificação dos equipamentos no fila de acordo com % do SLA.')
            r1c2.plotly_chart(create_fig_criticos(st.session_state['saldo_atual_varejo_selecao'][
                                                      ~st.session_state['saldo_atual_varejo_selecao'][
                                                          '% DO SLA'].isna()].copy()))

            r2c1.write('Saldo detalhado de equipamentos no fila.')
            r2c1.dataframe(saldo_atual_varejo_selecao[[
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
            r2c2.plotly_chart(create_fig_status(st.session_state['saldo_atual_varejo_selecao']))

            r3c1.write('Maiores volumetrias em fila.')
            r3c1.plotly_chart(create_fig_volume_fila(saldo_atual_varejo.selection.rows))


    tabs_liberado.title('Varejo Liberado')
    t2r0c1, t2r0c2, t2r0c3, t2r0c4 = tabs_liberado.columns(4)
    t2r1c1, _ = tabs_liberado.columns(2, gap='large')
    tabs_liberado.write('')
    t2r2c1, t2r2c2 = tabs_liberado.columns(2, gap='large')
    t2r3c1 = tabs_liberado.container()

    dt_varejo = t2r1c1.date_input(label='Data de liberação')

    if 'data_liberação' not in st.session_state:
        df_varejo_liberado = create_df_varejo_liberado(dt_varejo)
    elif st.session_state['data_liberação'] != dt_varejo:
        df_varejo_liberado = create_df_varejo_liberado(dt_varejo)
        try:
            st.session_state.pop('varejo_liberado_resumido')
        except:
            pass
    else:
        df_varejo_liberado = st.session_state['varejo_liberado']
    
    if 'varejo_liberado' in st.session_state:
        if len(st.session_state['varejo_liberado']) > 0:

            st.sidebar.download_button(label='BAIXAR LIBERADOS', data=html_varejo(), file_name=f'Varejo {str(dt_varejo)}.html', use_container_width=True)
            
            if 'varejo_liberado_resumido' not in st.session_state:
                st.session_state['varejo_liberado_resumido'] = create_df_varejo_liberado_resumido(df_varejo_liberado)
                try:
                    df_varejo_liberado_resumido = st.session_state['varejo_liberado_resumido']
                except:
                    pass
            else:
                df_varejo_liberado_resumido = st.session_state['varejo_liberado_resumido']
            
            if st.session_state['varejo_liberado_resumido'].shape[0] > 0:
                t2r2c1.write('Resumo de equipamentos liberado.')
                varejo_liberado = t2r2c1.dataframe(df_varejo_liberado_resumido[['NUM OS', 'CLIENTE', 'EQUIPAMENTO', 'QUANTIDADE']],
                                                            hide_index=True,
                                                            use_container_width=True,
                                                            on_select='rerun')

            if varejo_liberado.selection.rows:
                df_varejo_liberado_resumido['CONCATENADO'] = df_varejo_liberado_resumido['NUM OS'] + df_varejo_liberado_resumido['CLIENTE'] + df_varejo_liberado_resumido['EQUIPAMENTO']
                df_varejo_liberado['CONCATENADO'] = df_varejo_liberado['NUM OS'] + df_varejo_liberado['CLIENTE'] + df_varejo_liberado['EQUIPAMENTO']
                filtro_liberado = list(df_varejo_liberado_resumido.iloc[varejo_liberado.selection.rows]['CONCATENADO'])
                varejo_liberado_selecao = df_varejo_liberado[df_varejo_liberado['CONCATENADO'].isin(filtro_liberado)]
                st.session_state['varejo_liberado_selecao'] = varejo_liberado_selecao

            if t2r0c4.button('FILTROS DE LIBERAÇÃO', use_container_width=True):
                open_dialog_filtros_saida()

            if 'varejo_liberado_selecao' in st.session_state and varejo_liberado.selection.rows:
                t2r2c2.write('Relação de equipamentos localizados.')
                t2r2c2.plotly_chart(create_fig_status_liberacao())

                t2r3c1.write('Lista de varejo liberado detalhada.')
                t2r3c1.dataframe(st.session_state['varejo_liberado_selecao'][['ENDEREÇO', 'CAIXA', 'SERIAL', 'CLIENTE', 'EQUIPAMENTO',
                                                                            'NUM OS', 'ENTRADA GERFLOOR', 'ENTRADA FILA',
                                                                            'SAÍDA FILA']].sort_values(['SAÍDA FILA']),
                                hide_index=True,
                                use_container_width=True,
                                column_config={
                                    'ENTRADA GERFLOOR': st.column_config.DateColumn('ENTRADA GERFLOOR', format='DD/MM/YYYY'),
                                    'ENTRADA FILA': st.column_config.DateColumn('ENTRADA FILA', format='DD/MM/YYYY HH:mm:ss'),
                                    'SAÍDA FILA': st.column_config.DateColumn('SAÍDA FILA', format='DD/MM/YYYY HH:mm:ss')
                                })
    
        else:
            t2r2c1.header('Sem liberação de varejo para a data informada!')


    tabs_saida.title('Saída de Equipamentos')
    t3r0c1, t3r0c2, t3r0c3, t3r0c4 = tabs_saida.columns(4)
    tabs_saida.write('')
    t3r1c1, t3r1c2 = tabs_saida.columns(2, gap='large')
    t3r2c1 = tabs_saida.container()
    tabs_saida.write('')
    t3r3c1 = tabs_saida.container()

    if 'df_saidas_varejo' not in st.session_state or 'df_saidas_varejo_resumido' not in st.session_state:
        st.session_state['df_saidas_varejo'] = create_df_saidas_varejo(historico_fila)
        st.session_state['df_saidas_varejo_resumido'] = create_df_saidas_varejo_resumido(st.session_state['df_saidas_varejo'])
    
        df_saidas_varejo = st.session_state['df_saidas_varejo']
        df_saidas_varejo_resumido = st.session_state['df_saidas_varejo_resumido']
    else:
        df_saidas_varejo = st.session_state['df_saidas_varejo']
        df_saidas_varejo_resumido = st.session_state['df_saidas_varejo_resumido']

    t3r1c1.write('Resumo de equipamentos enviados ao laboratório.')
    saidas_varejo = t3r1c1.dataframe(df_saidas_varejo_resumido[['CLIENTE', 'EQUIPAMENTO', 'QUANTIDADE']],
                      hide_index=True,
                      use_container_width=True,
                      on_select='rerun')
    
    if saidas_varejo.selection.rows:
        df_saidas_varejo_resumido['CONCATENADO'] = df_saidas_varejo_resumido['CLIENTE'] + df_saidas_varejo_resumido['EQUIPAMENTO']
        df_saidas_varejo['CONCATENADO'] = df_saidas_varejo['CLIENTE'] + df_saidas_varejo['EQUIPAMENTO']
        filtro_saldo = list(df_saidas_varejo_resumido.iloc[saidas_varejo.selection.rows]['CONCATENADO'])
        saidas_varejo_selecao = df_saidas_varejo[df_saidas_varejo['CONCATENADO'].isin(filtro_saldo)]
        st.session_state['saidas_varejo_selecao'] = saidas_varejo_selecao

        t3r0c1.metric('Total de saídas (seleção)', '{:,}'.format(len(saidas_varejo_selecao['SERIAL'])).replace(',','.'))
        if len(saidas_varejo_selecao[saidas_varejo_selecao['SAÍDA FILA'] >= datetime.today()-timedelta(hours=datetime.today().hour+1)]) > 0:
            filtro_ontem = ((saidas_varejo_selecao['SAÍDA FILA'] >= datetime.today()-timedelta(days=1, hours=datetime.today().hour, minutes=datetime.today().minute)) &
                            (saidas_varejo_selecao['SAÍDA FILA'] <= datetime.today()-timedelta(hours=datetime.today().hour, minutes=datetime.today().minute)))
            try:
                t3r0c2.metric('Saídas do dia (seleção)', '{:,}'.format(len(saidas_varejo_selecao[saidas_varejo_selecao['SAÍDA FILA'] >= datetime.today()-timedelta(hours=datetime.today().hour+1)])).replace(',','.'),
                            delta='{:.2%}'.format(((len(saidas_varejo_selecao[saidas_varejo_selecao['SAÍDA FILA'] >= datetime.today()])) - len(saidas_varejo_selecao[filtro_ontem])) / len(saidas_varejo_selecao[saidas_varejo_selecao['SAÍDA FILA'] >= datetime.today()])))
            except:
                t3r0c2.metric('Saídas do dia (seleção)', '{:,}'.format(len(saidas_varejo_selecao[saidas_varejo_selecao['SAÍDA FILA'] >= datetime.today()-timedelta(hours=datetime.today().hour+1)])).replace(',','.'),
                            delta='{:.2%}'.format(0))
        else: t3r0c2.metric('Saídas do dia (seleção)', 0)
    else:
        t3r0c1.metric('Total de saídas', '{:,}'.format(sum(df_saidas_varejo_resumido['QUANTIDADE'])).replace(',','.'))
        if len(df_saidas_varejo[df_saidas_varejo['SAÍDA FILA'] >= datetime.today()-timedelta(hours=datetime.today().hour+1)]) > 0:
            filtro_ontem = ((df_saidas_varejo['SAÍDA FILA'] >= datetime.today()-timedelta(days=1, hours=datetime.today().hour, minutes=datetime.today().minute)) &
                            (df_saidas_varejo['SAÍDA FILA'] <= datetime.today()-timedelta(hours=datetime.today().hour, minutes=datetime.today().minute)))
            try:
                t3r0c2.metric('Saídas do dia', '{:,}'.format(len(df_saidas_varejo[df_saidas_varejo['SAÍDA FILA'] >= datetime.today()-timedelta(hours=datetime.today().hour+1)])).replace(',','.'),
                            delta='{:.2%}'.format(((len(df_saidas_varejo[df_saidas_varejo['SAÍDA FILA'] >= datetime.today()])) - len(df_saidas_varejo[filtro_ontem])) / len(df_saidas_varejo[df_saidas_varejo['SAÍDA FILA'] >= datetime.today()])))
            except:
                t3r0c2.metric('Saídas do dia', '{:,}'.format(len(df_saidas_varejo[df_saidas_varejo['SAÍDA FILA'] >= datetime.today()-timedelta(hours=datetime.today().hour+1)])).replace(',','.'),
                            delta='{:.2%}'.format(0))
        else: t3r0c2.metric('Saídas do dia', 0)

    if t3r0c4.button('FILTROS DE SAÍDA', use_container_width=True):
        open_dialog_filtros_saida()

    if 'saidas_varejo_selecao' in st.session_state and saidas_varejo.selection.rows:
        t3r1c2.write('Status dos equipamentos entregues em relação ao SLA.')
        t3r1c2.plotly_chart(create_fig_status(st.session_state['saidas_varejo_selecao']))

        t3r2c1.write('Histórico detalhado de equipamentos entregues ao laboratório.')
        t3r2c1.dataframe(st.session_state['saidas_varejo_selecao'][['CAIXA', 'SERIAL', 'CLIENTE', 'EQUIPAMENTO',
                                                                       'NUM OS', 'ENTRADA GERFLOOR', 'ENTRADA FILA',
                                                                       'SAÍDA FILA', 'AGING TOTAL', 'AGING FILA', 'STATUS']].sort_values(['SAÍDA FILA']),
                         hide_index=True,
                         use_container_width=True,
                         column_config={
                             'ENTRADA GERFLOOR': st.column_config.DateColumn('ENTRADA GERFLOOR', format='DD/MM/YYYY'),
                             'ENTRADA FILA': st.column_config.DateColumn('ENTRADA FILA', format='DD/MM/YYYY HH:mm:ss'),
                             'SAÍDA FILA': st.column_config.DateColumn('SAÍDA FILA', format='DD/MM/YYYY HH:mm:ss')
                         })
        
        t3r3c1.write('Distribuição do status dos equipamentos entregues ao longo dos meses.')
        t3r3c1.plotly_chart(create_fig_status_saidas())

    
    tabs_terceiros.title('Equipamentos Retirados por Terceiros')
    t4r0c1, t4r0c2, t4r0c3, t4r0c4 = tabs_terceiros.columns(4)
    tabs_terceiros.write('')
    t4r1c1, t4r1c2 = tabs_terceiros.columns(2, gap='large')
    t4r2c1, t4r2c2 = tabs_terceiros.columns([6, 4], gap='large')

    if 'df_terceiros_varejo' not in st.session_state or 'df_terceiros_varejo_resumido' not in st.session_state:
        st.session_state['df_terceiros_varejo'] = create_df_terceiros_varejo(historico_fila)
        st.session_state['df_terceiros_varejo_resumido'] = create_df_terceiros_varejo_resumido(
            st.session_state['df_terceiros_varejo'])

        df_terceiros_varejo = st.session_state['df_terceiros_varejo']
        df_terceiros_varejo_resumido = st.session_state['df_terceiros_varejo_resumido']
    else:
        df_terceiros_varejo = st.session_state['df_terceiros_varejo']
        df_terceiros_varejo_resumido = st.session_state['df_terceiros_varejo_resumido']

    t4r1c1.write('Resumo de equipamentos entregues a outros setores.')
    terceiros_varejo = t4r1c1.dataframe(df_terceiros_varejo_resumido[['TERCEIROS', 'QUANTIDADE']],
                                           hide_index=True,
                                           use_container_width=True,
                                           on_select='rerun')

    if terceiros_varejo.selection.rows:
        filtro_saldo = list(df_terceiros_varejo_resumido.iloc[terceiros_varejo.selection.rows]['TERCEIROS'])
        terceiros_varejo_selecao = df_terceiros_varejo[df_terceiros_varejo['TERCEIROS'].isin(filtro_saldo)]
        st.session_state['terceiros_varejo_selecao'] = terceiros_varejo_selecao

        t4r0c1.metric('Total em posse de terceiros (seleção)',
                      '{:,}'.format(len(terceiros_varejo_selecao['SERIAL'])).replace(',', '.'))
    else:
        t4r0c1.metric('Total em posse de terceiros',
                      '{:,}'.format(sum(df_terceiros_varejo_resumido['QUANTIDADE'])).replace(',', '.'))

    if 'terceiros_varejo_selecao' in st.session_state and terceiros_varejo.selection.rows:
        t4r1c2.write('Classificação dos equipamentos em posse de terceiros de acordo com % do SLA.')
        t4r1c2.plotly_chart(create_fig_criticos(st.session_state['terceiros_varejo_selecao'][
                                                    ~st.session_state['terceiros_varejo_selecao'][
                                                        '% DO SLA'].isna()].copy()))

        t4r2c2.write('Status dos equipamentos em relação ao SLA.')
        t4r2c2.plotly_chart(create_fig_status(st.session_state['terceiros_varejo_selecao']))

        t4r2c1.write('Histórico detalhado de equipamentos entregues ao laboratório.')
        t4r2c1.dataframe(st.session_state['terceiros_varejo_selecao'][['CAIXA', 'SERIAL', 'CLIENTE', 'EQUIPAMENTO',
                                                                          'NUM OS', 'ENTRADA GERFLOOR', 'ENTRADA FILA',
                                                                          'SAÍDA FILA', 'AGING TOTAL', 'AGING FILA',
                                                                          'STATUS']].sort_values(['SAÍDA FILA']),
                         hide_index=True,
                         use_container_width=True,
                         column_config={
                             'ENTRADA GERFLOOR': st.column_config.DateColumn('ENTRADA GERFLOOR', format='DD/MM/YYYY'),
                             'ENTRADA FILA': st.column_config.DateColumn('ENTRADA FILA', format='DD/MM/YYYY HH:mm:ss'),
                             'SAÍDA FILA': st.column_config.DateColumn('SAÍDA FILA', format='DD/MM/YYYY HH:mm:ss')
                         })
        
    tabs_geral.dataframe(st.session_state['historico_fila'][~st.session_state['historico_fila']['FLUXO'].isin(['CONTRATO', 'OS INTERNA'])])
