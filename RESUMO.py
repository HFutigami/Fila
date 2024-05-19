import streamlit as st
import pandas as pd
from office365.sharepoint.files.file import File
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from datetime import datetime
import io
import plotly.express as px

### Autenticação ao Sharepoint

sharepoint_base_url = 'https://gertecsao.sharepoint.com/sites/PowerBi-Estoque/'
folder_in_sharepoint = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Fila'
sharepoint_user = 'gertec.visualizador@gertec.com.br'
sharepoint_password = 'VY&ks28@AM2!hs1'

saldo_fila_url = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Fila/saldo.parquet'
saldo_file_url = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Fila/'

auth = AuthenticationContext(sharepoint_base_url)
auth.acquire_token_for_user(sharepoint_user, sharepoint_password)
ctx = ClientContext(saldo_fila_url, auth)
web = ctx.web
ctx.execute_query()


### Funções

def df_sharep(file_url):
    """Gera um DataFrame a partir de um diretório do SharePoint."""
    file_response = File.open_binary(ctx, file_url)
    bytes_file_obj = io.BytesIO()
    bytes_file_obj.write(file_response.content)
    bytes_file_obj.seek(0)
    return pd.read_parquet(bytes_file_obj)


### DattaFrames

# SLA Contratos
sla_contratos = pd.DataFrame(columns=['Cliente', 'Prazo'], data=[['STONE',15], ['GERTEC',10], ['REK',5]])
sla_contratos.rename(columns={'Cliente':'CLIENTE', 'Prazo':'PRAZO'}, inplace=True)
sla_contratos.set_index('CLIENTE', inplace=True)


# Endereços
lista_de_enderecos = pd.DataFrame(columns=['Vagas'], data=[['FILA'], ['CX 01'], ['CX 02'], ['B 01 2']]) 
lista_de_enderecos = pd.concat([pd.DataFrame(columns=['Vagas'], data=['LAB']), lista_de_enderecos])


# Saldo geral
historico_fila = df_sharep(saldo_fila_url)
historico_fila['FLUXO'] = historico_fila['FLUXO'].str.replace('000001', 'CONTRATO')
historico_fila['FLUXO'] = historico_fila['FLUXO'].str.replace('000002', 'VAREJO')
historico_fila['FLUXO'] = historico_fila['FLUXO'].str.replace('000003', 'VAREJO')
historico_fila['FLUXO'] = historico_fila['FLUXO'].str.replace('000004', 'OS INTERNA')

historico_fila['GARANTIA'] = historico_fila['GARANTIA'].str.upper()
historico_fila['CLIENTE'] = historico_fila['CLIENTE'].str.upper()
historico_fila['ENTRADA GERFLOOR'] = pd.to_datetime(historico_fila['ENTRADA GERFLOOR'], format='%d/%m/%Y %I:%M:%S %p')

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
historico_fila['AGING TOTAL'] = (historico_fila['ULTIMA DATA'] - historico_fila['ENTRADA GERFLOOR']).dt.days
historico_fila['AGING TOTAL'] = historico_fila['AGING TOTAL'].astype('int')
historico_fila['AGING FILA'] = (historico_fila['ULTIMA DATA'] - historico_fila['ENTRADA FILA']).dt.days
historico_fila['AGING FILA'] = historico_fila['AGING FILA'].astype('int')

historico_fila = historico_fila.join(sla_contratos, on=['CLIENTE'], how='left')
historico_fila['%CRITIC'] = None
historico_fila.loc[~historico_fila['PRAZO'].isna(), '%CRITIC'] = historico_fila['AGING TOTAL']/historico_fila['PRAZO']
historico_fila['STATUS'] = None
historico_fila.loc[(historico_fila['%CRITIC'] > 0.0) & (historico_fila['%CRITIC'] <= 0.1), 'STATUS'] = "RÁPIDO"
historico_fila.loc[(historico_fila['%CRITIC'] > 0.1) & (historico_fila['%CRITIC'] <= 0.3), 'STATUS'] = "MÉDIO"
historico_fila.loc[(historico_fila['%CRITIC'] > 0.3) & (historico_fila['%CRITIC'] <= 0.6), 'STATUS'] = "LENTO"
historico_fila.loc[(historico_fila['%CRITIC'] > 0.6) & (historico_fila['%CRITIC'] <= 1.0), 'STATUS'] = "CRÍTICO"
historico_fila.loc[(historico_fila['%CRITIC'] > 1.0), 'STATUS'] = "SLA ESTOURADO"


# Saldo Geral
saldo_geral = historico_fila.copy()
saldo_geral.drop(columns=['ULTIMA DATA', '%CRITIC', 'SAÍDA FILA', 'PRAZO'], inplace=True)


# Saldo atual
saldo_atual = historico_fila[historico_fila['ENDEREÇO'] != "LAB"].copy()
saldo_atual.drop(columns=['SAÍDA FILA'], inplace=True)


# Saldo atual contratos
saldo_atual_contratos = saldo_atual[saldo_atual['FLUXO'] != "xCONTRATO"].copy()
saldo_atual_contratos = saldo_atual_contratos.groupby(['CLIENTE', 'EQUIPAMENTO'])['SERIAL'].count().reset_index()
saldo_atual_contratos.rename(columns={'SERIAL':'QUANTIDADE'}, inplace=True)


# Varejo liberado
processos_varejo = pd.read_excel('PROCESSOS VAREJO_ BACKLOG_SEPARAÇÃO_ANALISE TECNICA 16.05.24.xlsx',
                                 sheet_name='LAB. - SEPARAÇÃO',
                                 dtype='str')
processos_varejo = processos_varejo[['Nr Serie', 'Num OS', 'Produto_1']]
processos_varejo.rename(columns={'Nr Serie':'SERIAL', 'Num OS':'NUM OS'}, inplace=True)
processos_varejo.set_index(['SERIAL', 'NUM OS'], inplace=True)

varejo_liberado = historico_fila.join(processos_varejo,
                                      on=['SERIAL', 'NUM OS'],
                                      how='right')
varejo_liberado.drop(columns=['Produto_1',
                              'SAÍDA FILA',
                              'ULTIMA DATA',
                              'AGING TOTAL',
                              'AGING FILA',
                              'PRAZO',
                              '%CRITIC',
                              'STATUS'], inplace=True)
varejo_liberado.sort_values('ENDEREÇO', inplace=True)


# Status liberados
porcent_varejo_liberado = varejo_liberado.copy()
porcent_varejo_liberado.drop(columns=['FLUXO', 'ENTRADA GERFLOOR'], inplace=True)
porcent_varejo_liberado['STATUS'] = ''
porcent_varejo_liberado.loc[porcent_varejo_liberado['ENDEREÇO'].isna(), 'STATUS'] = 'NÃO LOCALIZADO'
porcent_varejo_liberado.loc[~porcent_varejo_liberado['STATUS'].isin(['LAB', 'NÃO LOCALIZADO']), 'STATUS'] = 'LOCALIZADO'
porcent_varejo_liberado.loc[~porcent_varejo_liberado['STATUS'].isin(['LOCALIZADO', 'NÃO LOCALIZADO']), 'STATUS'] = 'EM LABORATÓRIO'
porcent_varejo_liberado = porcent_varejo_liberado.groupby('STATUS')['SERIAL'].count().reset_index()
porcent_varejo_liberado.rename(columns={'SERIAL':'QUANTIDADE'}, inplace=True)


### GRÁFICOS

# Saldo Contratos
figdf_saldo_atual = saldo_atual_contratos[saldo_atual_contratos['CLIENTE'].isin(saldo_atual_contratos.groupby('CLIENTE')['QUANTIDADE'].sum().reset_index().sort_values('QUANTIDADE', ascending=False)['CLIENTE'].head(3))]
figdf_saldo_atual.sort_values('QUANTIDADE', ascending=False, inplace=True)

fig_saldo_contratos = px.bar(figdf_saldo_atual,
                             x='CLIENTE',
                             y='QUANTIDADE',
                             color='EQUIPAMENTO',
                             barmode='stack',
                             orientation='v',
                             text='QUANTIDADE')

fig_saldo_contratos.update_traces(textposition='inside',
                                  orientation='v')

fig_saldo_contratos.update_layout(yaxis_title=None,
                                  xaxis_title=None,
                                  yaxis_visible=False)


# Relação Status de Prazo
figdf_status_prazo = saldo_geral.groupby('STATUS')['EQUIPAMENTO'].count().reset_index().copy()
figdf_status_prazo.rename(columns={'EQUIPAMENTO':'QUANTIDADE'}, inplace=True)
figdf_status_prazo.sort_values('QUANTIDADE', ascending=False, inplace=True)

colors_status_prazo = []
for i in (figdf_status_prazo['STATUS']):
    if i == "RÁPIDO":
        colors_status_prazo.append('#008000')
    elif i == "MÉDIO":
        colors_status_prazo.append('#32CD32')
    elif i == "LENTO":
        colors_status_prazo.append('#FFD700')
    elif i == "CRÍTICO":
        colors_status_prazo.append('#FF8C00')
    elif i == "SLA ESTOURADO":
        colors_status_prazo.append('#8B0000')

fig_status_prazo = px.pie(figdf_status_prazo,
                          names='STATUS',
                          values='QUANTIDADE',
                          color='QUANTIDADE',
                          color_discrete_sequence=colors_status_prazo)
fig_status_prazo.update_traces(textinfo='value+percent')


# Relação Status de Saídas
porcent_historico_saidas = historico_fila[historico_fila['ENDEREÇO'] != "LAB"].groupby(['STATUS'])['EQUIPAMENTO'].count().reset_index().copy()
porcent_historico_saidas.rename(columns={'EQUIPAMENTO':'QUANTIDADE'}, inplace=True)

colors_status_saidas = []
for i in (porcent_historico_saidas['STATUS']):
    if i == "RÁPIDO":
        colors_status_saidas.append('#008000')
    elif i == "MÉDIO":
        colors_status_saidas.append('#32CD32')
    elif i == "LENTO":
        colors_status_saidas.append('#FFD700')
    elif i == "CRÍTICO":
        colors_status_saidas.append('#FF8C00')
    elif i == "SLA ESTOURADO":
        colors_status_saidas.append('#8B0000')

fig_status_saidas = px.pie(porcent_historico_saidas,
                           names='STATUS',
                           values='QUANTIDADE',
                           color='QUANTIDADE',
                           color_discrete_sequence=colors_status_saidas)
fig_status_saidas.update_traces(textinfo='value+percent')
fig_status_saidas.update_layout(showlegend=False)


# Barras Status Saídas
figdf_historico_saidas = historico_fila.copy()
figdf_historico_saidas['ENTRADA FILA'] = figdf_historico_saidas['ENTRADA FILA'].dt.strftime('%Y/%m')
figdf_historico_saidas = figdf_historico_saidas.groupby(['ENTRADA FILA', 'STATUS'])['EQUIPAMENTO'].count().reset_index()
figdf_historico_saidas.rename(columns={'EQUIPAMENTO':'QUANTIDADE'}, inplace=True)

colors_historico_saidas = []
for i in (figdf_historico_saidas['STATUS']):
    if i == "RÁPIDO":
        colors_historico_saidas.append('#008000')
    elif i == "MÉDIO":
        colors_historico_saidas.append('#32CD32')
    elif i == "LENTO":
        colors_historico_saidas.append('#FFD700')
    elif i == "CRÍTICO":
        colors_historico_saidas.append('#FF8C00')
    elif i == "SLA ESTOURADO":
        colors_historico_saidas.append('#8B0000')

fig_historico_saidas = px.bar(figdf_historico_saidas,
                              x='ENTRADA FILA',
                              y='QUANTIDADE',
                              color='STATUS',
                              color_discrete_sequence=colors_historico_saidas,
                              orientation='v',
                              barmode='group',
                              text='QUANTIDADE')

fig_historico_saidas.update_traces(textposition='inside',
                                  orientation='v')

fig_historico_saidas.update_layout(yaxis_title=None,
                                   xaxis_title=None,
                                   yaxis_visible=False)



### Design do Streamlit

st.set_page_config(page_title='Dash', page_icon='https://i.imgur.com/mOEfCM8.png', layout='wide')

st.image('https://seeklogo.com/images/G/gertec-logo-D1C911377C-seeklogo.com.png?v=637843433630000000', width=200)
st.header('', divider='gray')

st.sidebar.title('FILTROS')
filter_equipamento = st.sidebar.multiselect('EQUIPAMENTO', [1,2,3], )
filter_serial = st.sidebar.multiselect('SERIAL', [1,2,3])
filter_numos = st.sidebar.multiselect('NÚMERO DE OS', [1,2,3])
filter_caixa = st.sidebar.multiselect('CAIXA', [1,2,3])
filter_cliente = st.sidebar.multiselect('CLIENTE', [1,2,3])
filter_endereco = st.sidebar.multiselect('ENDEREÇO', lista_de_enderecos)
filter_fluxo = st.sidebar.multiselect('FLUXO', [1,2,3])
filter_garantia = st.sidebar.multiselect('GARANTIA', [1,2,3])
filter_entrada_gerfloor = st.sidebar.multiselect('DATA DE ENTRADA GERFLOOR', [1,2,3])
filter_entrada_fila = st.sidebar.multiselect('DATA DE ENTRADA FILA', [1,2,3])
filter_saida_fila = st.sidebar.multiselect('DATA DE SAÍDA FILA', [1,2,3])


st.header('Saldo de Equipamentos do Fila')
st.plotly_chart(fig_status_prazo)
st.dataframe(saldo_geral,
             use_container_width=True,
             hide_index=True,
             column_config={'ENTRADA GERFLOOR': st.column_config.DateColumn(format="DD/MM/YYYY HH:mm:ss"),
                            'ENTRADA FILA': st.column_config.DateColumn(format="DD/MM/YYYY HH:mm:ss")})

st.header('')
st.header('')

st.header('Saldo de Contratos')

sac_col1, sac_col2 = st.columns([8, 12])
sac_col1.dataframe(saldo_atual_contratos, hide_index=True, use_container_width=True)
sac_col2.plotly_chart(fig_saldo_contratos, use_container_width=True)

st.header('')
st.header('')

st.header('Varejo Liberado')

pvlp_col1, _ = st.columns([5,15])
pvlp_col1.date_input(label='Data de liberação')
pvl_col1, pvl_col2 = st.columns([12, 8])
pvl_col1.dataframe(varejo_liberado,
                   hide_index=True,
                   use_container_width=True,
                   column_config={'ENTRADA GERFLOOR': st.column_config.DateColumn(format="DD/MM/YYYY HH:mm:ss"),
                                  'ENTRADA FILA': st.column_config.DateColumn(format="DD/MM/YYYY HH:mm:ss"),
                                  'SAÍDA FILA': st.column_config.DateColumn(format="DD/MM/YYYY HH:mm:ss")})
pvl_col2.plotly_chart(px.pie(data_frame=porcent_varejo_liberado,
                             values='QUANTIDADE',
                             names='STATUS').update_traces(textinfo='value+percent'),
                      use_container_width=True)

st.header('')
st.header('')

st.header('Histórico de Movimentações')

st.dataframe(historico_fila,
             hide_index=True,
             use_container_width=True,
             column_config={'ENTRADA GERFLOOR': st.column_config.DateColumn(format="DD/MM/YYYY HH:mm:ss"),
                            'ENTRADA FILA': st.column_config.DateColumn(format="DD/MM/YYYY HH:mm:ss"),
                            'SAÍDA FILA': st.column_config.DateColumn(format="DD/MM/YYYY HH:mm:ss")})

hsf_col1, hsf_col2 = st.columns([8, 12])
hsf_col1.plotly_chart(fig_status_saidas, use_container_width=True)
hsf_col2.plotly_chart(fig_historico_saidas, use_container_width=True)
