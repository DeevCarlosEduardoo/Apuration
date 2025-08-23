
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from tkinter import Canvas as TkCanvas, Button, PhotoImage
import numpy as np
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from reportlab.pdfgen import canvas as pdf_canvas  # Renomeando para evitar conflitos
from datetime import datetime, timedelta
import locale
import re
from babel.numbers import format_currency
from tkinter.filedialog import askdirectory
import os
from dateutil.relativedelta import relativedelta
import requests
import locale
from tkcalendar import DateEntry 
from boxsdk import Client, OAuth2
import io
import time 
from pandas.tseries.offsets import DateOffset
from datetime import datetime
import requests
import webbrowser
import threading
from flask import Flask, request
import time
import sys
import os
from tkinter import ttk
import calendar
import datetime as dt

# === CONFIGURAÇÕES BOX ===
CLIENT_ID = 'zkacla486aw46nrxpk58oapx4aqm84ze'
CLIENT_SECRET = 'x0iZRVgP41qHjR6QkLcJf1OL3Eh6PMww'
REDIRECT_URI = 'http://localhost:5000/callback'
AUTH_URL = f'https://account.box.com/api/oauth2/authorize?response_type=code&client_id={CLIENT_ID}&redirect_uri={REDIRECT_URI}'
TOKEN_URL = 'https://api.box.com/oauth2/token'
UPLOAD_URL = 'https://upload.box.com/api/2.0/files/content'
FOLDER_ID = '304180333772'  # ✅ Sem "d_" aqui

access_token_global = None

app = Flask(__name__)

@app.route('/callback')
def callback():
    global access_token_global, refresh_token_global

    code = request.args.get('code')
    if not code:
        return 'Erro: código não recebido.'

    data = {
        "grant_type": "authorization_code",
        "code": code,
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "redirect_uri": REDIRECT_URI
    }

    response = requests.post(TOKEN_URL, data=data)
    tokens = response.json()

    if response.status_code == 200:
        access_token_global = tokens.get('access_token')
        refresh_token_global = tokens.get('refresh_token')
        return 'Token recebido com sucesso. Pode fechar esta janela.'
    else:
        return f"Erro ao obter token: {response.text}"
    
def refresh_access_token():
    global access_token_global, refresh_token_global

    if not refresh_token_global:
        raise Exception("❌ Refresh token não disponível.")

    data = {
        "grant_type": "refresh_token",
        "refresh_token": refresh_token_global,
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET
    }

    response = requests.post(TOKEN_URL, data=data)
    tokens = response.json()

    if response.status_code == 200:
        access_token_global = tokens.get('access_token')
        refresh_token_global = tokens.get('refresh_token')  # atualiza também!
        print("🔄 Access token atualizado com sucesso.")
    else:
        raise Exception(f"❌ Erro ao atualizar token: {response.text}")

def iniciar_flask():
    app.run(port=5000)

def autenticar_box():
    # Inicia Flask em outra thread
    thread = threading.Thread(target=iniciar_flask)
    thread.daemon = True
    thread.start()

    # Abre o navegador para login
    time.sleep(1)
    webbrowser.open(AUTH_URL)

    # Aguarda o token chegar
    print("🔐 Aguardando autenticação...")
    while access_token_global is None:
        time.sleep(1)

    return access_token_global


def validar_colunas(df):
    # --- Verifica INICIO DA APURAÇÃO ---
    if "INICIO DA APURAÇÃO" not in df.columns:
        messagebox.showerror("Erro", "Coluna 'INICIO DA APURAÇÃO' não encontrada!")
        sys.exit()

    coluna_inicio_apuracao = df["INICIO DA APURAÇÃO"]

    # Não pode ter datas
    if np.issubdtype(coluna_inicio_apuracao.dtype, np.datetime64):
        messagebox.showerror("Erro", "A coluna INICIO DA APURAÇÃO contém valores no formato data!")
        sys.exit()

    for valor in coluna_inicio_apuracao:
        if isinstance(valor, (dt.datetime, dt.date, np.datetime64)):
            messagebox.showerror("Erro", f"A coluna INICIO DA APURAÇÃO contém um valor de data: {valor}")
            sys.exit()

    # --- Verifica e converte DT. INÍCIO ---
    if "DT. INÍCIO" not in df.columns:
        messagebox.showerror("Erro", "Coluna 'DT. INÍCIO' não encontrada!")
        sys.exit()

    try:
        # Converte qualquer valor possível para datetime
        df["DT. INÍCIO"] = pd.to_datetime(df["DT. INÍCIO"], errors="raise")
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível converter todos os valores da coluna DT. INÍCIO para data.\nDetalhes: {e}")
        sys.exit()

    # Confirma que agora a coluna tem somente datas
    if not np.issubdtype(df["DT. INÍCIO"].dtype, np.datetime64):
        messagebox.showerror("Erro", "A coluna DT. INÍCIO ainda contém valores que não são datas!")
        sys.exit()

    return df


def processar_arquivos():
    BaseCurtoCaminho = caminho_arquivo1.get()
    BaseLongoCaminho = caminho_arquivo2.get()
    ColigadosCaminho = caminho_arquivo3.get()
    BaseHistoricaCaminho = caminho_arquivo4.get()
    TituloRelatorio = TitleInput.get()
    ValorCheckBox = CheckboxValue.get()

    access_token = autenticar_box()

    print(access_token)
    meses_portugues = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}
    
    if BaseCurtoCaminho and BaseHistoricaCaminho and TituloRelatorio:


        df_validacao = pd.read_excel(BaseCurtoCaminho)

        # Validação
        validar_colunas(df_validacao)


        if  ValorCheckBox == True:
             
            data_atual = DateValue
                    
                    
        elif ValorCheckBox == False:
           
            data_atual = datetime.today()  # Pega a data de hoje
                
                

        messagebox.showinfo("Processo Iniciado", "Arquivos selecinados, iniciando o processo")

        Base = pd.read_excel(BaseCurtoCaminho)
        BaseConsumo = pd.read_excel(BaseCurtoCaminho,engine='openpyxl',sheet_name='Bases - Consumo')


        Base['INICIO DA APURAÇÃO'] = pd.to_numeric(Base['INICIO DA APURAÇÃO'], errors='coerce')

        Base['DATA INICIAL'] = pd.to_datetime('1899-12-30') + pd.to_timedelta(Base['INICIO DA APURAÇÃO'], unit='D')

        Base['PRAZO APURACAO'] = pd.to_numeric(Base['PRAZO APURACAO'], errors='coerce').fillna(0).astype(int)

        Base['DATA FINAL'] = Base['DATA INICIAL'] + Base['PRAZO APURACAO'].apply(lambda x: DateOffset(months=int(x)))

        Base = Base[Base['DATA INICIAL'] < data_atual]
        
        Base = Base[Base['ATIVO OU INATIVO'] == 'ATIVO']

        ColigadosFiltros = (
                            BaseConsumo[
                                (BaseConsumo['SAP'] == 'Coligado') &
                                (BaseConsumo['ATIVO OU INATIVO'] == 'ATIVO')
                            ][['CÓDIGO SAP', 'RAZÃO SOCIAL', 'SAP PRINCIPAL']]
                            .drop_duplicates()
                        )

        Base = Base[Base['ATIVO OU INATIVO'] == 'ATIVO']

        BaseCurto = Base
        Coligados = ColigadosFiltros
        BaseLongo = Base
        BaseHistorica = pd.read_excel(BaseHistoricaCaminho,engine='openpyxl',sheet_name='Sheet1')
        BaseHistoricaCompleta = pd.read_excel(BaseHistoricaCaminho,engine='openpyxl',sheet_name='Sheet1')

        df = BaseCurto
        BaseLongo = BaseLongo



        # Filtrar os dados
        df_filtrado = df[(df['ATIVO OU INATIVO'] == 'ATIVO') & (df['LINHA DO CONTRATO'] == 'Principal') & (df['ATIVO OU INATIVO'] != 'CONTRATO ENCERRADO') & (BaseLongo['INICIO DA APURAÇÃO'].notna())&  (BaseLongo['MODALIDADE'] == 'Compra e Venda com consumo')].drop_duplicates(subset='SAP PRINCIPAL')

        

        BaseLongoFiltrado = BaseLongo[(BaseLongo['ATIVO OU INATIVO'] == 'ATIVO') & (BaseLongo['LINHA DO CONTRATO'] == 'Principal') & (BaseLongo['INICIO DA APURAÇÃO'].notna())& (BaseLongo['MODALIDADE'] == 'NOVA LOCAÇÃO')].drop_duplicates(subset='SAP PRINCIPAL')
        
        
    
        LongoPrazoApuração = BaseLongo[(BaseLongo['ATIVO OU INATIVO'] == 'ATIVO') & (BaseLongo['LINHA DO CONTRATO'] == 'Principal') & (BaseLongo['INICIO DA APURAÇÃO'].notna()) & (BaseLongo['PRAZO APURACAO'].notna())].drop_duplicates(subset='SAP PRINCIPAL')

        MANUTENÇÃO = BaseLongo[
            (BaseLongo['ATIVO OU INATIVO'] == 'ATIVO') & 
            (BaseLongo['LINHA DO CONTRATO'] == 'Principal') & 
            (BaseLongo['INICIO DA APURAÇÃO'].notna()) & 
            ((BaseLongo['MODALIDADE'] == 'MANUTENÇÃO') & (BaseLongo['CONSUMO ANO 1'].notna()))
        ].drop_duplicates(subset='SAP PRINCIPAL')

        NovoComodato = BaseLongo[
            (BaseLongo['ATIVO OU INATIVO'] == 'ATIVO') & 
            (BaseLongo['LINHA DO CONTRATO'] == 'Principal') & 
            (BaseLongo['INICIO DA APURAÇÃO'].notna()) & 
            (BaseLongo['MODALIDADE'] == 'NOVO COMODATO') 
        ].drop_duplicates(subset='SAP PRINCIPAL')

        acordodeconsumo = BaseLongo[
            (BaseLongo['ATIVO OU INATIVO'] == 'ATIVO') & 
            (BaseLongo['LINHA DO CONTRATO'] == 'Principal') & 
            (BaseLongo['INICIO DA APURAÇÃO'].notna()) & 
            (BaseLongo['MODALIDADE'] == 'Acordo de Consumo') 
        ].drop_duplicates(subset='SAP PRINCIPAL')

        EquipamentosGerais = BaseLongo[
            (BaseLongo['ATIVO OU INATIVO'] == 'ATIVO')  & 
            (BaseLongo['EQUIPAMENTO'].notnull())
        ]


        BaseCurtoFiltrado = df_filtrado[(df_filtrado['ATIVO OU INATIVO'] == 'ATIVO') & (df_filtrado['LINHA DO CONTRATO'] == 'Principal' ) & (df_filtrado['INICIO DA APURAÇÃO'].notna()) &  (df_filtrado['MODALIDADE'] == 'Compra e Venda com consumo') ].drop_duplicates(subset='SAP PRINCIPAL')

        CurtoPrazoApuração = df[(df['ATIVO OU INATIVO'] == 'ATIVO') & (df['LINHA DO CONTRATO'] == 'Principal') & (df['INICIO DA APURAÇÃO'].notna()) & (df['PRAZO APURACAO'].notna())].drop_duplicates(subset='SAP PRINCIPAL')
        PrazodeApuraçãoConcatenados = pd.concat([LongoPrazoApuração, CurtoPrazoApuração], ignore_index=True)


  

        try:
            locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')
        except:
            # Caso esteja no Windows, tente:
            locale.setlocale(locale.LC_TIME, 'Portuguese')

        # Configuração de estilo das tabelas
        style = TableStyle([
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # Alinhamento horizontal à esquerda
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Alinhamento vertical centralizado
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
            ('TOPPADDING', (0, 0), (-1, -1), 1),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ])


        StyleTituloMudado = TableStyle([
            # Estilo geral
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 5),  # Tamanho da fonte geral
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # Alinhamento à esquerda
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Centralização vertical
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno inferior
            ('TOPPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno superior

            # Estilo para a linha de cabeçalho
            ('SPAN', (0, 0), (-1, 0)),  # Mescla as colunas na primeira linha
            ('BACKGROUND', (0, 0), (-1, 0), (68/255, 83/255, 106/255)),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),  # Texto branco
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Alinhamento centralizado
            ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),  # Linha acima da primeira linha
            ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),  # Linha abaixo da primeira linha
            ('FONTSIZE', (0, 0), (-1, 0), 6),  # Aumenta o tamanho da fonte na primeira linha
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Deixa a primeira linha em negrito


            # Estilo para os títulos das linhas
            ('TEXTCOLOR', (0, 1), (0, -1), colors.black),  # Texto preto                
            ('FONTNAME', (0, 1), (0, -1), 'Helvetica-Bold'),  # Texto em negrito

            # Fundo das células de conteúdo
            ('BACKGROUND', (1, 1), (1, -1), colors.white),  # Fundo branco
            ('INNERGRID', (0, 1), (-1, -1), 0, colors.white),  # Sem grade interna
            ('BOX', (0, 1), (-1, -1), 0, colors.white),  # Sem borda externa
        ])

        StyleColigados = TableStyle([
            # Estilo geral
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),  # Fonte padrão
            ('FONTSIZE', (0, 0), (-1, -1), 5),  # Tamanho da fonte geral
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # Alinhamento à esquerda
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Centralização vertical
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno inferior
            ('TOPPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno superior

            # Estilo para a linha de cabeçalho
            ('SPAN', (0, 0), (-1, 0)),  # Mescla as colunas na primeira linha
            ('BACKGROUND', (0, 0), (-1, 0), (68/255, 83/255, 106/255)),  # Fundo azul-escuro
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),  # Texto branco
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Alinhamento centralizado
            ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),  # Linha acima da primeira linha
            ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),  # Linha abaixo da primeira linha
            ('FONTSIZE', (0, 0), (-1, 0), 6),  # Aumenta o tamanho da fonte na primeira linha
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Negrito na primeira linha

            # Estilo para a segunda linha (também em negrito)
            ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),  # Negrito na segunda linha

            # Estilo para os títulos das linhas restantes
            ('TEXTCOLOR', (0, 2), (0, -1), colors.black),  # Texto preto (da terceira linha em diante)
            ('FONTNAME', (0, 2), (-1, -1), 'Helvetica'),  # Fonte normal da terceira linha em diante

            # Fundo das células de conteúdo
            ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Grade interna
            ('BOX', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Borda externa

        ])

        StyleBaseHistorica = TableStyle([
            # Estilo geral
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),  # Fonte padrão
            ('FONTSIZE', (0, 0), (-1, -1), 5),  # Tamanho da fonte geral
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # Alinhamento à esquerda
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Centralização vertical
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno inferior
            ('TOPPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno superior

            # Estilo para a linha de cabeçalho
            ('SPAN', (0, 0), (-1, 0)),  # Mescla as colunas na primeira linha
            ('BACKGROUND', (0, 0), (-1, 0), (68/255, 83/255, 106/255)),  # Fundo azul-escuro
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),  # Texto branco
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Alinhamento centralizado
            ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),  # Linha acima da primeira linha
            ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),  # Linha abaixo da primeira linha
            ('FONTSIZE', (0, 0), (-1, 0), 6),  # Aumenta o tamanho da fonte na primeira linha
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Negrito na primeira linha

            # Estilo para a segunda linha (também em negrito)
            ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),  # Negrito na segunda linha

            # Estilo para os títulos das linhas restantes
            ('TEXTCOLOR', (0, 2), (0, -1), colors.black),  # Texto preto (da terceira linha em diante)
            ('FONTNAME', (0, 2), (-1, -1), 'Helvetica'),  # Fonte normal da terceira linha em diante

            # Fundo das células de conteúdo
            ('BACKGROUND', (1, 1), (1, -1), colors.white),  # Fundo branco
            ('INNERGRID', (0, 1), (-1, -1), 0, colors.white),  # Sem grade interna
            ('BOX', (0, 1), (-1, -1), 0, colors.white),  # Sem borda externa

            # Divisões da tabela com cinza claro
            ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Grade interna
            ('BOX', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Borda externa
        ])



        def calcular_altura_tabela(num_linhas):
            altura_linha = 15  # Altura de cada linha da tabela (ajustada)
            return num_linhas * altura_linha

        messagebox.showinfo("Salvar Arquivos", "Iniciar o processo de salvar!")

     
        contador = 0
        SAPZADA = 0
        RAZAOZADA = 0

       
  
        df_concatenado = pd.concat([BaseLongoFiltrado, BaseCurtoFiltrado,MANUTENÇÃO,NovoComodato,acordodeconsumo], ignore_index=True)

        df_concatenado = df_concatenado.drop_duplicates(subset=["SAP PRINCIPAL"], keep="first")

        print(SapSelecionado)
        print(ValorSAP)

        
        
        if SapSelecionado == True:
            ValorSapInt = int(ValorSAP)
            df_concatenado = df_concatenado[df_concatenado["SAP PRINCIPAL"] == ValorSapInt]

        df_concatenado = df_concatenado.drop_duplicates(subset=["SAP PRINCIPAL"], keep="first")
        
        df_concatenado.to_excel("excelfiltrado.xlsx")
        
        # Criar PDFs para as linhas filtradas
        for index, row in df_concatenado.iterrows():
            
            sap_principal_filtro = row['SAP PRINCIPAL']
            Ninterno = row['Nº INTERNO']
            Versão = row['VERSÃO']

            FiltrandoLentes = BaseConsumo[(BaseConsumo['CÓDIGO SAP'] == sap_principal_filtro) & (BaseConsumo['Nº INTERNO'] == Ninterno) & (BaseConsumo['ATIVO OU INATIVO'] == "ATIVO")]

            print(sap_principal_filtro)

            contador += 1
        

            lentesFiltroHistorico = FiltrandoLentes['SKU PRODUTO'].dropna().unique().tolist()

            if not lentesFiltroHistorico:
                lentesFiltroHistorico = [
                "ICB00",
                "PCB00",
                "ZCB00",
                "ZCT00",
                "ZFR00",
                "ZKB00",
                "ZLB00",
                "ZMA00",
                "ZMB00",
                "ZMT00",
                "ZXR00",
                "ZXT00",
                "DFW00",
                "DCB00",
                "DIB00",
                "DIU00",
                "DFR00",
                "DET00",
                "DEN00"
            ]

            data_base_excel = datetime(1899, 12, 30)
            data_base_serial = 2

            data_atual_serial = (datetime.today() - datetime(1899, 12, 30)).days + data_base_serial

            def excel_serial_to_date(serial):
                return data_base_excel + pd.to_timedelta(serial, unit="D")  # Ajuste do Excel (base 1900)

            ColigadosFiltrado = Coligados[Coligados['SAP PRINCIPAL'] == sap_principal_filtro]
            ContacatenadoPrazo = PrazodeApuraçãoConcatenados[PrazodeApuraçãoConcatenados['SAP PRINCIPAL'] == sap_principal_filtro]

        

            ClientesManutenção = MANUTENÇÃO[(MANUTENÇÃO['SAP PRINCIPAL'] == sap_principal_filtro) & (MANUTENÇÃO['DATA FINAL'] > data_atual)]

            ClientesNovoComodato = NovoComodato[(NovoComodato['SAP PRINCIPAL'] == sap_principal_filtro) & (NovoComodato['DATA FINAL'] > data_atual)]

            ClientesacordodeConsumo = acordodeconsumo[(acordodeconsumo['SAP PRINCIPAL'] == sap_principal_filtro) & (acordodeconsumo['DATA FINAL'] > data_atual)]

            BaseLongoFiltradoCliente = BaseLongoFiltrado[(BaseLongoFiltrado['SAP PRINCIPAL'] == sap_principal_filtro) & (BaseLongoFiltrado['DATA FINAL'] > data_atual)]

        

            BaseCurtoFiltradoCliente = BaseCurtoFiltrado[(BaseCurtoFiltrado['SAP PRINCIPAL'] == sap_principal_filtro ) & (BaseCurtoFiltrado['DATA FINAL'] > data_atual)]
           




           
           
            


            if not BaseLongoFiltradoCliente.empty:
                print("Passou Longo")
                DataDaApuraçãoLongo = data_base_excel + timedelta(BaseLongoFiltradoCliente.iloc[0].get('INICIO DA APURAÇÃO'))
                
                DataFimApuração = DataDaApuraçãoLongo + relativedelta(months=int(BaseLongoFiltradoCliente.iloc[0].get('PRAZO APURACAO', 0)))
                
                mes_extenso = meses_portugues[DataDaApuraçãoLongo.month]
                anodeApuracaosemfromatar = DataDaApuraçãoLongo.year
                
                
                DataDaApuraçãoFormatada = f"{mes_extenso} de {anodeApuracaosemfromatar}"
                DataDaApuraçãoFiltro = DataDaApuraçãoLongo.strftime('%Y-%m')
                DataDaApuraçãoFiltroCurto = DataDaApuraçãoLongo.strftime('%Y-%m')
                DataFimApuraçãoFormatada = DataFimApuração.strftime('%d/%m/%Y')


                dif_anos = DataFimApuração.year - data_atual.year
                dif_meses = DataFimApuração.month - data_atual.month

                # Total de meses
                total_meses = (dif_anos * 12) + dif_meses - 1
                print(total_meses)
                
                DataInicioApuraçãoLongoFormatada = DataDaApuraçãoLongo.strftime('%d/%m/%Y')
                DataFimApuraçãoLongoFormatada = DataFimApuração.strftime('%d/%m/%Y')


                try:
                    # Tenta converter diretamente usando timedelta
                    if isinstance(BaseLongoFiltradoCliente.iloc[0]['DT. INÍCIO'], (int, float)) and BaseLongoFiltradoCliente.iloc[0]['DT. INÍCIO'] > 60:
                            DataInicioLongo = data_base_excel + timedelta(BaseLongoFiltradoCliente.iloc[0]['DT. INÍCIO'] - pd.Timedelta(days=2))    
                    else:
                            # Tenta converter diretamente para datetime, assumindo que já é uma data válida
                            DataInicioLongo = pd.to_datetime(BaseLongoFiltradoCliente.iloc[0]['DT. INÍCIO'])
                            
                    if isinstance(BaseLongoFiltradoCliente.iloc[0]['DT. FINAL'], (int, float)) and BaseLongoFiltradoCliente.iloc[0]['DT. FINAL'] > 60:
                            DataFimLongo = data_base_excel + timedelta(BaseLongoFiltradoCliente.iloc[0]['DT. FINAL'] - pd.Timedelta(days=2))    
                    else:
                            # Tenta converter diretamente para datetime, assumindo que já é uma data válida
                            DataFimLongo = pd.to_datetime(BaseLongoFiltradoCliente.iloc[0]['DT. FINAL'])     
                    
                except TypeError:
                    # Caso dê erro, trata como número serial do Excel
                    serial_inicio = int(BaseLongoFiltradoCliente.iloc[0]['DT. INÍCIO'])  # Conversão explícita
                    serial_fim = int(BaseLongoFiltradoCliente.iloc[0]['DT. FINAL'])      # Conversão explícita
                    DataInicioLongo = datetime(1900, 1, 1) + timedelta(days=serial_inicio - 2)
                    DataFimLongo = datetime(1900, 1, 1) + timedelta(days=serial_fim - 2)

                # Formata as datas
                DataInicioLongoFormatada = DataInicioLongo.strftime('%d/%m/%Y')
                DataFimLongoFormatada = DataFimLongo.strftime('%d/%m/%Y')
                Vigencia = f"{DataInicioLongoFormatada} - {DataFimLongoFormatada}"

                InicioApuração = DataDaApuraçãoLongo

            if not BaseCurtoFiltradoCliente.empty:
                print("Passou Curtp")
                DataDaApuraçãoCurto = data_base_excel + timedelta(BaseCurtoFiltradoCliente.iloc[0].get('INICIO DA APURAÇÃO'))
               
                DataFimApuraçãoCurto = DataDaApuraçãoCurto + relativedelta(months=int(BaseCurtoFiltradoCliente.iloc[0].get('PRAZO APURACAO', 0)))
                
                mes_extenso = meses_portugues[DataDaApuraçãoCurto.month]
                anodeApuracaosemfromatar = DataDaApuraçãoCurto.year
                
                
                DataDaApuraçãoFormatada = f"{mes_extenso} de {anodeApuracaosemfromatar}"
                DataDaApuraçãoFiltro = DataDaApuraçãoCurto.strftime('%Y-%m')
                DataDaApuraçãoFiltroCurto = DataDaApuraçãoCurto.strftime('%Y-%m')
                DataFimApuraçãoFormatadaCurto = DataFimApuraçãoCurto.strftime('%d/%m/%Y')


                dif_anos = DataFimApuraçãoCurto.year - data_atual.year
                dif_meses = DataFimApuraçãoCurto.month - data_atual.month

                # Total de meses
                total_mesesCurto = (dif_anos * 12) + dif_meses - 1
                print(total_mesesCurto)
                
                InicioApuração = DataDaApuraçãoCurto
                DataInicioApuraçãoCurtoFormatada = DataDaApuraçãoCurto.strftime('%d/%m/%Y')
                    
            if not ClientesManutenção.empty:
                print("Passou Manutenção")
                DataDaApuraçãoManutenção = data_base_excel + timedelta(ClientesManutenção.iloc[0].get('INICIO DA APURAÇÃO'))
                
                DataFimApuraçãoManutenção = DataDaApuraçãoManutenção + relativedelta(months=int(ClientesManutenção.iloc[0].get('PRAZO APURACAO', 0)))
                
                mes_extenso = meses_portugues[DataDaApuraçãoManutenção.month]
                anodeApuracaosemfromatar = DataDaApuraçãoManutenção.year
                
                
                DataDaApuraçãoFormatada = f"{mes_extenso} de {anodeApuracaosemfromatar}"
                DataDaApuraçãoFiltro = DataDaApuraçãoManutenção.strftime('%Y-%m')
                DataFimApuraçãoFormatada = DataFimApuraçãoManutenção.strftime('%d/%m/%Y')


                dif_anos = DataFimApuraçãoManutenção.year - data_atual.year
                dif_meses = DataDaApuraçãoManutenção.month - data_atual.month

                # Total de meses
                total_meses_manutenção = (dif_anos * 12) + dif_meses - 1
                print(total_meses_manutenção)
                
                DataInicioApuraçãoManutençãoFormatada = DataDaApuraçãoManutenção.strftime('%d/%m/%Y')

                DataDaApuraçãoLongo = DataDaApuraçãoManutenção

            if not ClientesNovoComodato.empty:
                print("Passou Novo Comodato")
                DataDaApuraçãoNovoComodato = data_base_excel + timedelta(ClientesNovoComodato.iloc[0].get('INICIO DA APURAÇÃO'))
                
                DataFimApuraçãoNovoComodato = DataDaApuraçãoNovoComodato + relativedelta(months=int(ClientesNovoComodato.iloc[0].get('PRAZO APURACAO', 0)))
                
                mes_extenso = meses_portugues[DataDaApuraçãoNovoComodato.month]
                anodeApuracaosemfromatar = DataDaApuraçãoNovoComodato.year
                
                
                DataDaApuraçãoFormatada = f"{mes_extenso} de {anodeApuracaosemfromatar}"
                DataDaApuraçãoFiltro = DataDaApuraçãoNovoComodato.strftime('%Y-%m')
                DataFimApuraçãoFormatada = DataFimApuraçãoNovoComodato.strftime('%d/%m/%Y')


                dif_anos = DataFimApuraçãoNovoComodato.year - data_atual.year
                dif_meses = DataDaApuraçãoNovoComodato.month - data_atual.month

                # Total de meses
                total_meses_NovoComodato = (dif_anos * 12) + dif_meses  - 1
                print(total_meses_NovoComodato)
                
                DataInicioApuraçãoNovoComodatoFormatada = DataDaApuraçãoNovoComodato.strftime('%d/%m/%Y')

                DataDaApuraçãoLongo = DataDaApuraçãoNovoComodato          

            if not ClientesacordodeConsumo.empty:
                print("Passou Novo Comodato")
                DataDaApuraçãoacordodeconsumo = data_base_excel + timedelta(ClientesacordodeConsumo.iloc[0].get('INICIO DA APURAÇÃO'))
                
                DataFimApuraçãoNacordodeconsumo = DataDaApuraçãoacordodeconsumo + relativedelta(months=int(ClientesacordodeConsumo.iloc[0].get('PRAZO APURACAO', 0)))
                
                mes_extenso = meses_portugues[DataDaApuraçãoacordodeconsumo.month]
                anodeApuracaosemfromatar = DataDaApuraçãoacordodeconsumo.year
                
                
                DataDaApuraçãoFormatada = f"{mes_extenso} de {anodeApuracaosemfromatar}"
                DataDaApuraçãoFiltro = DataDaApuraçãoacordodeconsumo.strftime('%Y-%m')
                DataFimApuraçãoFormatada = DataFimApuraçãoNacordodeconsumo.strftime('%d/%m/%Y')



                dif_anos = DataFimApuraçãoNacordodeconsumo.year - data_atual.year
                dif_meses = DataDaApuraçãoacordodeconsumo.month - data_atual.month

                # Total de meses
                total_meses_acordodeconsumo = (dif_anos * 12) + dif_meses  - 1
                print(total_meses_acordodeconsumo)
                
                DataInicioApuraçãoacordodeconsumoFormatada = DataDaApuraçãoacordodeconsumo.strftime('%d/%m/%Y')

                DataDaApuraçãoLongo = DataDaApuraçãoacordodeconsumo              
                

        
            
            if not BaseCurtoFiltradoCliente.empty:
                DataDaApuraçãoCurto = data_base_excel + timedelta(BaseCurtoFiltradoCliente.iloc[0].get('INICIO DA APURAÇÃO'))
                
                mes_extenso_curto = meses_portugues[DataDaApuraçãoCurto.month]
                anoCurtoApuraçao = DataDaApuraçãoCurto.year
                    
                DataDaApuraçãoFormatadaCurto = f"{mes_extenso_curto} de {anoCurtoApuraçao}"
                print(f"Data curto: {DataDaApuraçãoFormatadaCurto}")

            else:
                DataDaApuraçãoFormatadaCurto = None

            if not BaseLongoFiltradoCliente.empty:
                DataDaApuraçãoLongo = data_base_excel + timedelta(BaseLongoFiltradoCliente.iloc[0].get('INICIO DA APURAÇÃO'))
                
                mes_extenso_Longo = meses_portugues[DataDaApuraçãoLongo.month]
                anoLongoApuraçao = DataDaApuraçãoLongo.year
                    
                DataDaApuraçãoFormatadaLongo = f"{mes_extenso_Longo} de {anoLongoApuraçao}"
                print(f"Data Longo: {DataDaApuraçãoFormatadaLongo}")

            else:
                DataDaApuraçãoFormatadaLongo = None

            if not ClientesManutenção.empty:
                DataDaApuraçãoMANUTENÇÃO= data_base_excel + timedelta(ClientesManutenção.iloc[0].get('INICIO DA APURAÇÃO'))
                
                mes_extensoMANUTENÇÃO = meses_portugues[DataDaApuraçãoMANUTENÇÃO.month]
                anoMANUTENÇÃOApuraçao = DataDaApuraçãoMANUTENÇÃO.year
                
                DataDaApuraçãoFormatadaMANUTENÇÃO= f"{mes_extensoMANUTENÇÃO} de {anoMANUTENÇÃOApuraçao}"

            else:
                DataDaApuraçãoFormatadaMANUTENÇÃO = None
            
            if not ClientesNovoComodato.empty:
                DataDaApuraçãoNovoComodato= data_base_excel + timedelta(ClientesNovoComodato.iloc[0].get('INICIO DA APURAÇÃO'))
                
                mes_extensoNovoComodato = meses_portugues[DataDaApuraçãoNovoComodato.month]
                anoNovoComodatoApuraçao = DataDaApuraçãoNovoComodato.year
                
                DataDaApuraçãoFormatadaNovoComodato= f"{mes_extensoNovoComodato} de {anoNovoComodatoApuraçao}"

            else:
                DataDaApuraçãoFormatadaNovoComodato = None
            
            if not ClientesacordodeConsumo.empty:
                DataDaApuraçãoAcordodeConsumo= data_base_excel + timedelta(ClientesacordodeConsumo.iloc[0].get('INICIO DA APURAÇÃO'))
                
                mes_extensoAcordodeConsumo = meses_portugues[DataDaApuraçãoAcordodeConsumo.month]
                anoAcordoDeConsumo = DataDaApuraçãoAcordodeConsumo.year
                
                DataDaApuraçãoFormatadaAcordoConsumo= f"{mes_extensoAcordodeConsumo} de {anoAcordoDeConsumo}"

            else:
                DataDaApuraçãoFormatadaAcordoConsumo = None

            if  ValorCheckBox == True:
                def calcular_ano_referencia(data_inicio):
                    data_inicio = datetime.strptime(data_inicio, "%d/%m/%Y")  # Converte string para datetime
                    data_atual = DateValue
                    datetime
                    diferenca_meses = (data_atual.year - data_inicio.year) * 12 + (data_atual.month - data_inicio.month)

                    # Ajuste para garantir que o ciclo comece no mesmo mês e dure 12 meses completos
                    if data_atual.day < data_inicio.day:  
                        diferenca_meses -= 1  # Ainda não completou o ciclo

                    # Calcula o ano do ciclo
                    ano = (diferenca_meses // 12) + 1

                    return f"Ano {ano}"
                    
            elif ValorCheckBox == False:
                def calcular_ano_referencia(data_inicio):
                    data_inicio = datetime.strptime(data_inicio, "%d/%m/%Y")  # Converte string para datetime
                    data_atual = datetime.today()  # Pega a data de hoje
                    
                    diferenca_meses = (data_atual.year - data_inicio.year) * 12 + (data_atual.month - data_inicio.month)

                    # Ajuste para garantir que o ciclo comece no mesmo mês e dure 12 meses completos
                    if data_atual.day < data_inicio.day:  
                        diferenca_meses -= 1  # Ainda não completou o ciclo

                    # Calcula o ano do ciclo
                    ano = (diferenca_meses // 12) + 1 
                    
                    return f"Ano {ano}" 


            # Converte a entrada do usuário para um objeto de data
            if not BaseLongoFiltradoCliente.empty:
                AnodaApuração = calcular_ano_referencia(DataInicioApuraçãoLongoFormatada)

            if not BaseCurtoFiltradoCliente.empty:
                AnodaApuração = calcular_ano_referencia(DataInicioApuraçãoCurtoFormatada)
            
            if not ClientesManutenção.empty:
                AnodaApuração = calcular_ano_referencia(DataInicioApuraçãoManutençãoFormatada)
            
            if not ClientesNovoComodato.empty:
                AnodaApuração = calcular_ano_referencia(DataInicioApuraçãoNovoComodatoFormatada)

            if not ClientesacordodeConsumo.empty:
                AnodaApuração = calcular_ano_referencia(DataInicioApuraçãoacordodeconsumoFormatada)

  
            
            if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                data_menor = min(DataInicioApuraçãoLongoFormatada, DataInicioApuraçãoCurtoFormatada)
                AnodaApuração = calcular_ano_referencia(data_menor)
                print(data_menor)

            BaseHistorica['DataApuração'] = pd.to_datetime(BaseHistorica['Ano'].astype(str) + '-' + BaseHistorica['Mês'].astype(str).str.zfill(2))
            BaseHistoricaCompleta['DataApuração'] = pd.to_datetime(BaseHistoricaCompleta['Ano'].astype(str) + '-' + BaseHistoricaCompleta['Mês'].astype(str).str.zfill(2))

            if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                print(DataDaApuraçãoCurto)
                print(DataDaApuraçãoLongo)
                DataApuraçãoMenor = min(DataDaApuraçãoCurto, DataDaApuraçãoLongo)
                DataDaApuraçãoFiltro = DataApuraçãoMenor
                DataDaApuraçãoLongo = DataApuraçãoMenor
                print(DataDaApuraçãoLongo)

                
            
            if AnodaApuração == "Ano 1":   
                BaseHistoricaFiltrada = BaseHistorica[
                (
                    (BaseHistorica['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistorica['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistorica['Item 2'].isin(lentesFiltroHistorico))&
                (BaseHistorica['DataApuração'] >= DataDaApuraçãoFiltro)&
                (BaseHistorica['DataApuração'] <= (DataDaApuraçãoLongo + relativedelta(months=11)).strftime('%Y-%m'))
                ]
                
                BaseHistoricaFiltradaCompleta = BaseHistoricaCompleta[
                (
                    (BaseHistoricaCompleta['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistoricaCompleta['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistoricaCompleta['Item 2'].isin(lentesFiltroHistorico))&
                (BaseHistoricaCompleta['DataApuração'] >= DataDaApuraçãoFiltro)
                
                ]
                if not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:
                    data_inicio = DataDaApuraçãoLongo
                elif  BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    data_inicio = DataDaApuraçãoCurto
                
                if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    data_inicio = DataDaApuraçãoLongo

                if not ClientesacordodeConsumo.empty:
                    data_inicio = DataDaApuraçãoAcordodeConsumo

                if not  ClientesManutenção.empty :
                    data_inicio = DataDaApuraçãoManutenção

                if not  ClientesNovoComodato.empty :
                    data_inicio = DataDaApuraçãoNovoComodato
                

                meses_apurados = (data_atual.year - data_inicio.year) * 12 + (data_atual.month - data_inicio.month)

                # Se o dia atual for maior ou igual ao dia da data de início, inclui o mês atual
                if data_atual.day >= data_inicio.day:
                    meses_apurados += 1

                # Calcula o mês dentro do ciclo de 12 meses
                meses_passados = (meses_apurados % 12) or 12


                
            elif AnodaApuração == "Ano 2":

                BaseHistoricaFiltradaAno1 = BaseHistorica[
                (
                    (BaseHistorica['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistorica['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistorica['Item 2'].isin(lentesFiltroHistorico))&
                (BaseHistorica['DataApuração'] >= DataDaApuraçãoFiltro)&
                (BaseHistorica['DataApuração'] <= (DataDaApuraçãoLongo + relativedelta(months=11)).strftime('%Y-%m'))
                ]

                BaseHistoricaFiltrada = BaseHistorica[
                (
                    (BaseHistorica['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistorica['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistorica['Item 2'].isin(lentesFiltroHistorico))&
                (BaseHistorica['DataApuração'] >=(DataDaApuraçãoLongo + relativedelta(months=12)).strftime('%Y-%m'))&
                (BaseHistorica['DataApuração'] <= (DataDaApuraçãoLongo + relativedelta(months=23)).strftime('%Y-%m'))
                ]
                
                BaseHistoricaFiltradaCompleta = BaseHistoricaCompleta[
                (
                    (BaseHistoricaCompleta['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistoricaCompleta['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistoricaCompleta['Item 2'].isin(lentesFiltroHistorico)) &
                (BaseHistoricaCompleta['DataApuração'] >= DataDaApuraçãoLongo)
                ]

                if not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:
                    data_inicio = DataDaApuraçãoLongo + relativedelta(months=12)
                elif  BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    data_inicio = DataDaApuraçãoCurto + relativedelta(months=12)

                if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    data_inicio = DataDaApuraçãoLongo + relativedelta(months=12)

                if not ClientesacordodeConsumo.empty:
                    data_inicio = DataDaApuraçãoAcordodeConsumo + relativedelta(months=12)
                    
                if  not ClientesManutenção.empty:
                    data_inicio = DataDaApuraçãoManutenção + relativedelta(months=12)

                if not  ClientesNovoComodato.empty :
                    data_inicio = DataDaApuraçãoNovoComodato + relativedelta(months=12)

              
                meses_apurados = (data_atual.year - data_inicio.year) * 12 + (data_atual.month - data_inicio.month)

                # Se o dia atual for maior ou igual ao dia da data de início, inclui o mês atual
                if data_atual.day >= data_inicio.day:
                    meses_apurados += 1

                # Calcula o mês dentro do ciclo de 12 meses
                meses_passados = (meses_apurados % 12) or 12

            
                

                AnoFiltroINicio = (DataDaApuraçãoLongo + relativedelta(months=11)).strftime('%Y-%m')
                AnoFiltroIdim= (DataDaApuraçãoLongo + relativedelta(months=22)).strftime('%Y-%m')    

           

                print(AnoFiltroINicio)    
                print(AnoFiltroIdim)  
                print(AnodaApuração)        

            
            elif AnodaApuração == "Ano 3":
                BaseHistoricaFiltradaAno1 = BaseHistorica[
                (
                    (BaseHistorica['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistorica['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistorica['Item 2'].isin(lentesFiltroHistorico))&
                (BaseHistorica['DataApuração'] >= DataDaApuraçãoFiltro)&
                (BaseHistorica['DataApuração'] <= (DataDaApuraçãoLongo + relativedelta(months=11)).strftime('%Y-%m'))
                ]

                BaseHistoricaFiltradaAno2 = BaseHistorica[
                (
                    (BaseHistorica['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistorica['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistorica['Item 2'].isin(lentesFiltroHistorico))&
                (BaseHistorica['DataApuração'] >=(DataDaApuraçãoLongo + relativedelta(months=12)).strftime('%Y-%m'))&
                (BaseHistorica['DataApuração'] <= (DataDaApuraçãoLongo + relativedelta(months=23)).strftime('%Y-%m'))
                ]


                BaseHistoricaFiltrada = BaseHistorica[
                (
                    (BaseHistorica['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistorica['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistorica['Item 2'].isin(lentesFiltroHistorico))&
                (BaseHistorica['DataApuração'] >=(DataDaApuraçãoLongo + relativedelta(months=24)).strftime('%Y-%m'))&
                (BaseHistorica['DataApuração'] <= (DataDaApuraçãoLongo + relativedelta(months=35)).strftime('%Y-%m'))
                ]
                
                BaseHistoricaFiltradaCompleta = BaseHistoricaCompleta[
                (
                    (BaseHistoricaCompleta['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistoricaCompleta['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistoricaCompleta['Item 2'].isin(lentesFiltroHistorico))&
                (BaseHistoricaCompleta['DataApuração'] >= DataDaApuraçãoFiltro)
                ]

                if not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:
                    data_inicio = DataDaApuraçãoLongo + relativedelta(months=24)
                elif  BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    data_inicio = DataDaApuraçãoCurto + relativedelta(months=24)
                if  not ClientesManutenção.empty :
                    data_inicio = DataDaApuraçãoManutenção + relativedelta(months=24)

                if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    data_inicio = DataDaApuraçãoLongo + relativedelta(months=24)

                if not ClientesacordodeConsumo.empty:
                    data_inicio = DataDaApuraçãoAcordodeConsumo + relativedelta(months=24)
                
                if not  ClientesNovoComodato.empty :
                    data_inicio = DataDaApuraçãoNovoComodato + relativedelta(months=24)

             
                meses_apurados = (data_atual.year - data_inicio.year) * 12 + (data_atual.month - data_inicio.month)

                # Se o dia atual for maior ou igual ao dia da data de início, inclui o mês atual
                if data_atual.day >= data_inicio.day:
                    meses_apurados += 1

                # Calcula o mês dentro do ciclo de 12 meses
                meses_passados = (meses_apurados % 12) or 12


            elif AnodaApuração == "Ano 4":
                BaseHistoricaFiltradaAno1 = BaseHistorica[
                (
                    (BaseHistorica['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistorica['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistorica['Item 2'].isin(lentesFiltroHistorico))&
                (BaseHistorica['DataApuração'] >= DataDaApuraçãoFiltro)&
                (BaseHistorica['DataApuração'] <= (DataDaApuraçãoLongo + relativedelta(months=11)).strftime('%Y-%m'))
                ]

                BaseHistoricaFiltradaAno2 = BaseHistorica[
                (
                    (BaseHistorica['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistorica['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistorica['Item 2'].isin(lentesFiltroHistorico))&
                (BaseHistorica['DataApuração'] >=(DataDaApuraçãoLongo + relativedelta(months=12)).strftime('%Y-%m'))&
                (BaseHistorica['DataApuração'] <= (DataDaApuraçãoLongo + relativedelta(months=23)).strftime('%Y-%m'))
                ]

                BaseHistoricaFiltradaAno3 = BaseHistorica[
                (
                    (BaseHistorica['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistorica['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistorica['Item 2'].isin(lentesFiltroHistorico))&
                (BaseHistorica['DataApuração'] >=(DataDaApuraçãoLongo + relativedelta(months=24)).strftime('%Y-%m'))&
                (BaseHistorica['DataApuração'] <= (DataDaApuraçãoLongo + relativedelta(months=35)).strftime('%Y-%m'))
                ]

                BaseHistoricaFiltrada = BaseHistorica[
                (
                    (BaseHistorica['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistorica['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistorica['Item 2'].isin(lentesFiltroHistorico))&
                (BaseHistorica['DataApuração'] >=(DataDaApuraçãoLongo + relativedelta(months=36)).strftime('%Y-%m'))&
                (BaseHistorica['DataApuração'] <= (DataDaApuraçãoLongo + relativedelta(months=47)).strftime('%Y-%m'))
                ]
                
                BaseHistoricaFiltradaCompleta = BaseHistoricaCompleta[
                (
                    (BaseHistoricaCompleta['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistoricaCompleta['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistoricaCompleta['Item 2'].isin(lentesFiltroHistorico))&
                (BaseHistoricaCompleta['DataApuração'] >= DataDaApuraçãoFiltro)
                ]

                if not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:
                    data_inicio = DataDaApuraçãoLongo + relativedelta(months=36)
                elif  BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    data_inicio = DataDaApuraçãoCurto + relativedelta(months=36)
                if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    data_inicio = DataDaApuraçãoLongo + relativedelta(months=36)

                if not ClientesacordodeConsumo.empty:
                    data_inicio = DataDaApuraçãoAcordodeConsumo + relativedelta(months=36)
                if not ClientesManutenção.empty:
                    data_inicio = DataDaApuraçãoManutenção + relativedelta(months=36)
                if not  ClientesNovoComodato.empty :
                    data_inicio = DataDaApuraçãoNovoComodato + relativedelta(months=36)

                meses_apurados = (data_atual.year - data_inicio.year) * 12 + (data_atual.month - data_inicio.month)

                # Se o dia atual for maior ou igual ao dia da data de início, inclui o mês atual
                if data_atual.day >= data_inicio.day:
                    meses_apurados += 1

                # Calcula o mês dentro do ciclo de 12 meses
                meses_passados = (meses_apurados % 12) or 12

           
            elif AnodaApuração == "Ano 5":
                BaseHistoricaFiltradaAno1 = BaseHistorica[
                (
                    (BaseHistorica['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistorica['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistorica['Item 2'].isin(lentesFiltroHistorico))&
                (BaseHistorica['DataApuração'] >= DataDaApuraçãoFiltro)&
                (BaseHistorica['DataApuração'] <= (DataDaApuraçãoLongo + relativedelta(months=11)).strftime('%Y-%m'))
                ]

                BaseHistoricaFiltradaAno2 = BaseHistorica[
                (
                    (BaseHistorica['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistorica['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistorica['Item 2'].isin(lentesFiltroHistorico))&
                (BaseHistorica['DataApuração'] >=(DataDaApuraçãoLongo + relativedelta(months=12)).strftime('%Y-%m'))&
                (BaseHistorica['DataApuração'] <= (DataDaApuraçãoLongo + relativedelta(months=23)).strftime('%Y-%m'))
                ]

                BaseHistoricaFiltradaAno3 = BaseHistorica[
                (
                    (BaseHistorica['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistorica['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistorica['Item 2'].isin(lentesFiltroHistorico))&
                (BaseHistorica['DataApuração'] >=(DataDaApuraçãoLongo + relativedelta(months=24)).strftime('%Y-%m'))&
                (BaseHistorica['DataApuração'] <= (DataDaApuraçãoLongo + relativedelta(months=35)).strftime('%Y-%m'))
                ]

                BaseHistoricaFiltradaAno4 = BaseHistorica[
                (
                    (BaseHistorica['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistorica['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistorica['Item 2'].isin(lentesFiltroHistorico))&
                (BaseHistorica['DataApuração'] >=(DataDaApuraçãoLongo + relativedelta(months=36)).strftime('%Y-%m'))&
                (BaseHistorica['DataApuração'] <= (DataDaApuraçãoLongo + relativedelta(months=47)).strftime('%Y-%m'))
                ]

                BaseHistoricaFiltrada = BaseHistorica[
                (
                    (BaseHistorica['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistorica['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistorica['Item 2'].isin(lentesFiltroHistorico))&
                (BaseHistorica['DataApuração'] >=(DataDaApuraçãoLongo + relativedelta(months=48)).strftime('%Y-%m'))&
                (BaseHistorica['DataApuração'] <= (DataDaApuraçãoLongo + relativedelta(months=59)).strftime('%Y-%m'))
                ]
                
                BaseHistoricaFiltradaCompleta = BaseHistoricaCompleta[
                (
                    (BaseHistoricaCompleta['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistoricaCompleta['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistoricaCompleta['Item 2'].isin(lentesFiltroHistorico))&
                (BaseHistoricaCompleta['DataApuração'] >= DataDaApuraçãoFiltro)
                ]

                if not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:
                    data_inicio = DataDaApuraçãoLongo + relativedelta(months=48)
                elif  BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    data_inicio = DataDaApuraçãoCurto + relativedelta(months=48)
                if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    data_inicio = DataDaApuraçãoLongo + relativedelta(months=48)

                if not ClientesacordodeConsumo.empty:
                    data_inicio = DataDaApuraçãoAcordodeConsumo + relativedelta(months=48)
                if  not ClientesManutenção.empty:
                    data_inicio = DataDaApuraçãoManutenção + relativedelta(months=48)
                if not  ClientesNovoComodato.empty :
                    data_inicio = DataDaApuraçãoNovoComodato + relativedelta(months=48 )
                
            
                meses_apurados = (data_atual.year - data_inicio.year) * 12 + (data_atual.month - data_inicio.month)

                # Se o dia atual for maior ou igual ao dia da data de início, inclui o mês atual
                if data_atual.day >= data_inicio.day:
                    meses_apurados += 1

                # Calcula o mês dentro do ciclo de 12 meses
                meses_passados = (meses_apurados % 12) or 12

            # Converter as datas
       

            # if not BaseLongoFiltradoCliente.empty:
            #     dataInicioLongo = data_base_excel + timedelta(row['DT. INÍCIO'] - 2)
            if isinstance(row['DT. INÍCIO'], (int, float)) and row['DT. INÍCIO'] > 60:
                dataInicio = data_base_excel + timedelta(row['DT. INÍCIO'] - pd.Timedelta(days=2))    
            else:
                # Tenta converter diretamente para datetime, assumindo que já é uma data válida
                dataInicioLongo = pd.to_datetime(row['DT. INÍCIO'])

            if isinstance(row['DT. INÍCIO'], (int, float)) and row['DT. INÍCIO'] > 60:
                dataInicio = data_base_excel + timedelta(row['DT. INÍCIO'] - pd.Timedelta(days=2))    
            else:
                # Tenta converter diretamente para datetime, assumindo que já é uma data válida
                dataInicio = pd.to_datetime(row['DT. INÍCIO'])
            if isinstance(row['DT. FINAL'], (int, float)) and row['DT. FINAL'] > 60:
                dataFim = data_base_excel + timedelta(row['DT. FINAL'] - pd.Timedelta(days=2))    
            else:
                # Tenta converter diretamente para datetime, assumindo que já é uma data válida
                dataFim = pd.to_datetime(row['DT. INÍCIO'])

        
            dataInicioFormatada = dataInicioLongo.strftime('%d/%m/%Y')
            dataFimFormatada = dataFim.strftime('%d/%m/%Y')
            DataFormatadaComMêseAno = dataInicio.strftime('%B de %Y')
            

            if not BaseCurtoFiltradoCliente.empty:
                try:
                    # Tenta converter diretamente usando timedelta
                    if isinstance(BaseCurtoFiltradoCliente.iloc[0]['DT. INÍCIO'], (int, float)) and BaseCurtoFiltradoCliente.iloc[0]['DT. INÍCIO'] > 60:
                            DataInicioCurto = data_base_excel + timedelta(BaseCurtoFiltradoCliente.iloc[0]['DT. INÍCIO'] - pd.Timedelta(days=2))    
                    else:
                            # Tenta converter diretamente para datetime, assumindo que já é uma data válida
                            DataInicioCurto = pd.to_datetime(BaseCurtoFiltradoCliente.iloc[0]['DT. INÍCIO'])
                            
                    if isinstance(BaseCurtoFiltradoCliente.iloc[0]['DT. FINAL'], (int, float)) and BaseCurtoFiltradoCliente.iloc[0]['DT. FINAL'] > 60:
                            DataFimCurto = data_base_excel + timedelta(BaseCurtoFiltradoCliente.iloc[0]['DT. FINAL'] - pd.Timedelta(days=2))    
                    else:
                            # Tenta converter diretamente para datetime, assumindo que já é uma data válida
                            DataFimCurto = pd.to_datetime(BaseCurtoFiltradoCliente.iloc[0]['DT. FINAL'])     
                    
                except TypeError:
                    # Caso dê erro, trata como número serial do Excel
                    serial_inicio = int(BaseCurtoFiltradoCliente.iloc[0]['DT. INÍCIO'])  # Conversão explícita
                    serial_fim = int(BaseCurtoFiltradoCliente.iloc[0]['DT. FINAL'])      # Conversão explícita
                    DataInicioCurto = datetime(1900, 1, 1) + timedelta(days=serial_inicio - 2)
                    DataFimCurto = datetime(1900, 1, 1) + timedelta(days=serial_fim - 2)

                # Formata as datas
                DataInicioCurtoFormatada = DataInicioCurto.strftime('%d/%m/%Y')
                DataFimCurtoFormatada = DataFimCurto.strftime('%d/%m/%Y')
                VigenciaCurto = f"{DataInicioCurtoFormatada} - {DataFimCurtoFormatada}"
            else:
                VigenciaCurto = " - "

            if not ClientesManutenção.empty:
                try:
                    # Tenta converter diretamente usando timedelta
                    if isinstance(ClientesManutenção.iloc[0]['DT. INÍCIO'], (int, float)) and MANUTENÇÃO.iloc[0]['DT. INÍCIO'] > 60:
                            DataInicioMANUTENÇÃO = data_base_excel + timedelta(ClientesManutenção.iloc[0]['DT. INÍCIO'] - pd.Timedelta(days=2))    
                    else:
                            # Tenta converter diretamente para datetime, assumindo que já é uma data válida
                            DataInicioMANUTENÇÃO = pd.to_datetime(ClientesManutenção.iloc[0]['DT. INÍCIO'])
                            
                    if isinstance(ClientesManutenção.iloc[0]['DT. FINAL'], (int, float)) and MANUTENÇÃO.iloc[0]['DT. FINAL'] > 60:
                            DataFimMANUTENÇÃO = data_base_excel + timedelta(ClientesManutenção.iloc[0]['DT. FINAL'] - pd.Timedelta(days=2))    
                    else:
                            # Tenta converter diretamente para datetime, assumindo que já é uma data válida
                            DataFimMANUTENÇÃO = pd.to_datetime(ClientesManutenção.iloc[0]['DT. FINAL'])     
                    
                except TypeError:
                    # Caso dê erro, trata como número serial do Excel
                    serial_inicio = int(ClientesManutenção.iloc[0]['DT. INÍCIO'])  # Conversão explícita
                    serial_fim = int(ClientesManutenção.iloc[0]['DT. FINAL'])      # Conversão explícita
                    DataInicioCurto = datetime(1900, 1, 1) + timedelta(days=serial_inicio - 2)
                    DataFimCurto = datetime(1900, 1, 1) + timedelta(days=serial_fim - 2)

                # Formata as datas
                DataInicioMANUTENÇÃOFormatada = DataInicioMANUTENÇÃO.strftime('%d/%m/%Y')
                DataFimMANUTENÇÃOFormatada = DataFimMANUTENÇÃO.strftime('%d/%m/%Y')
                VigenciaMANUTENÇÃO = f"{DataInicioMANUTENÇÃOFormatada} - {DataFimMANUTENÇÃOFormatada}"
            else:
                VigenciaMANUTENÇÃO = " - "
            
            if not ClientesNovoComodato.empty:
                try:
                    # Tenta converter diretamente usando timedelta
                    if isinstance(ClientesNovoComodato.iloc[0]['DT. INÍCIO'], (int, float)) and ClientesNovoComodato.iloc[0]['DT. INÍCIO'] > 60:
                            DataInicioNovoComodato = data_base_excel + timedelta(ClientesNovoComodato.iloc[0]['DT. INÍCIO'] - pd.Timedelta(days=2))    
                    else:
                            # Tenta converter diretamente para datetime, assumindo que já é uma data válida
                            DataInicioNovoComodato = pd.to_datetime(ClientesNovoComodato.iloc[0]['DT. INÍCIO'])
                            
                    if isinstance(ClientesNovoComodato.iloc[0]['DT. FINAL'], (int, float)) and ClientesNovoComodato.iloc[0]['DT. FINAL'] > 60:
                            DataFimNovoComodato = data_base_excel + timedelta(ClientesNovoComodato.iloc[0]['DT. FINAL'] - pd.Timedelta(days=2))    
                    else:
                            # Tenta converter diretamente para datetime, assumindo que já é uma data válida
                            DataFimNovoComodato = pd.to_datetime(ClientesNovoComodato.iloc[0]['DT. FINAL'])     
                    
                except TypeError:
                    # Caso dê erro, trata como número serial do Excel
                    serial_inicio = int(ClientesNovoComodato.iloc[0]['DT. INÍCIO'])  # Conversão explícita
                    serial_fim = int(ClientesManutenção.iloc[0]['DT. FINAL'])      # Conversão explícita
                    DataInicioCurto = datetime(1900, 1, 1) + timedelta(days=serial_inicio - 2)
                    DataFimCurto = datetime(1900, 1, 1) + timedelta(days=serial_fim - 2)

                # Formata as datas
                DataInicioNovoComodatoFormatada = DataInicioNovoComodato.strftime('%d/%m/%Y')
                DataFimNovoComodatoFormatada = DataFimNovoComodato.strftime('%d/%m/%Y')
                VigenciaNovoComodato= f"{DataInicioNovoComodatoFormatada} - {DataFimNovoComodatoFormatada}"
            else:
                VigenciaNovoComodato = " - "
            
            if not ClientesacordodeConsumo.empty:
                try:
                    # Tenta converter diretamente usando timedelta
                    if isinstance(ClientesacordodeConsumo.iloc[0]['DT. INÍCIO'], (int, float)) and ClientesacordodeConsumo.iloc[0]['DT. INÍCIO'] > 60:
                            DataInicioAcordodeConsumo = data_base_excel + timedelta(ClientesacordodeConsumo.iloc[0]['DT. INÍCIO'] - pd.Timedelta(days=2))    
                    else:
                            # Tenta converter diretamente para datetime, assumindo que já é uma data válida
                            DataInicioAcordodeConsumo = pd.to_datetime(ClientesacordodeConsumo.iloc[0]['DT. INÍCIO'])
                            
                    if isinstance(ClientesacordodeConsumo.iloc[0]['DT. FINAL'], (int, float)) and ClientesacordodeConsumo.iloc[0]['DT. FINAL'] > 60:
                            DataFimAcordoDeConsumo = data_base_excel + timedelta(ClientesacordodeConsumo.iloc[0]['DT. FINAL'] - pd.Timedelta(days=2))    
                    else:
                            # Tenta converter diretamente para datetime, assumindo que já é uma data válida
                            DataFimAcordoDeConsumo = pd.to_datetime(ClientesacordodeConsumo.iloc[0]['DT. FINAL'])     
                    
                except TypeError:
                    # Caso dê erro, trata como número serial do Excel
                    serial_inicio = int(ClientesacordodeConsumo.iloc[0]['DT. INÍCIO'])  # Conversão explícita
                    serial_fim = int(ClientesacordodeConsumo.iloc[0]['DT. FINAL'])      # Conversão explícita
                    DataInicioCurto = datetime(1900, 1, 1) + timedelta(days=serial_inicio - 2)
                    DataFimCurto = datetime(1900, 1, 1) + timedelta(days=serial_fim - 2)

                # Formata as datas
                DataInicioAcordoDeConsumoFormatada = DataInicioAcordodeConsumo.strftime('%d/%m/%Y')
                DataFimAcordoDeConsumoFormatada = DataFimAcordoDeConsumo.strftime('%d/%m/%Y')
                VigenciaAcordoDeConsumo= f"{DataInicioAcordoDeConsumoFormatada} - {DataFimAcordoDeConsumoFormatada}"
            else:
                VigenciaAcordoDeConsumo = " - "
            
            
            
            df['VALOR TOTAL'] = pd.to_numeric(df['VALOR TOTAL'], errors='coerce')

            # Certifique-se de que não há valores NaN na coluna antes de formatar
            df = df.dropna(subset=['VALOR TOTAL'])

            # Formata os valores
            df['VALOR TOTAL'] = df['VALOR TOTAL'].apply(lambda x: f"R$ {x:.2f}".replace('.', ','))

            # Filtrar equipamentos
            


            EquipamentosGeraisFiltrado = EquipamentosGerais[(EquipamentosGerais['SAP PRINCIPAL'] == sap_principal_filtro)][['EQUIPAMENTO', 'DESCRIÇÃO EQUIPAMENTO','Nº INTERNO','SÉRIE']]


            EquipamentosGeraisFiltrado = EquipamentosGeraisFiltrado.drop_duplicates(subset=['EQUIPAMENTO', 'DESCRIÇÃO EQUIPAMENTO','Nº INTERNO','SÉRIE'])

        # ...

            # Criar a tabela de equipamentos de longo prazo
            equipamentos_longo_com_cabecalho = [['SKU Equipamento', 'Descrição','N INTERNO','Série']] + EquipamentosGeraisFiltrado.iloc[::-1].values.tolist()
            tabela_equipamentos_longo = Table(equipamentos_longo_com_cabecalho, colWidths=[100, 250])
            tabela_equipamentos_longo.setStyle(StyleColigados)

            # ... (restante do código)

            

            # ... (restante do código)

            # Dados para as tabelas
            RazaoSocialCompleta = f"{row['SAP PRINCIPAL']} - {row['RAZÃO SOCIAL']}"
            
            
            if pd.isnull(row['SAM']) or row['SAM'] == '':
                InfClientes = [['Informações do Cliente'],['Sap Principal', RazaoSocialCompleta], ['Consultor', row['CONSULTOR']], ['Distrital', row['DISTRITAL']], ['Sam', '']]
            else:
                InfClientes = [['Informações do Cliente'],['Sap Principal', RazaoSocialCompleta], ['Consultor', row['CONSULTOR']], ['Distrital', row['DISTRITAL']], ['Sam', row['SAM']]]

            # Concatenando os DataFrames e realizando os cálculos como no código anterior
            

            skus = FiltrandoLentes['SKU PRODUTO'].dropna().astype(str).tolist()
            descricoes = FiltrandoLentes['DESCRIÇÃO CONSUMO'].dropna().astype(str).tolist()


            lentes_dados =  [f"{sku} {desc}" for sku, desc in zip(skus, descricoes)]

            if not lentes_dados:
                lentes_dados = [
                    "ICB00 LIO TECNIS Eyhance",
                    "PCB00 LIO TECNIS ITEC PRELOADED",
                    "ZCB00 LIO TECNIS ONE",
                    "ZCT00 LIO TECNIS ONE TÓRICA",
                    "ZFR00 LIO TECNIS Synergy 0,00 R$",
                    "ZKB00 TECNIS ONE MF Low Add",
                    "ZLB00 TECNIS ONE MF Low Add",
                    "ZMA00 LIO TECNIS MF",
                    "ZMB00 TECNIS ONE MF",
                    "ZMT00 TECNIS ONE TÓRICA MF",
                    "ZXR00 TECNIS SYMFONY",
                    "ZXT00 TECNIS SYMFONY TÓRICA",
                    "DFW00 TECNIS SYNERGY TORIC SIMPLICITY",
                    "DCB00 TECNIS SIMPLICITY DCB",
                    "DIB00 TECNIS EYHANCE SIMPLICITY",
                    "DIU00 EYHANCE TORIC II SIMPLICITY",
                    "DFR00 TECNIS SYNERGY SIMPLICITY",
                    "DEN00 TECNIS PURESEE Simplicity - peça única",
                    "DET00 TECNIS PURESEE Simplicity Tórica - peça única"
                ]
            # Separando as colunas
            lentes = [linha.split(maxsplit=1)[0] for linha in lentes_dados]
            descricao = [linha.split(maxsplit=1)[1] for linha in lentes_dados]

            # Criando o DataFrame
            dados_lentes = pd.DataFrame({
                "LENTES": lentes,
                "DESCRIÇÃO CONSUMO": descricao
            })  
            
            dados_lentes2 = pd.DataFrame({
                "LENTES": lentes,
                "DESCRIÇÃO CONSUMO": descricao
            })  

            dados_lentes3 = pd.DataFrame({
                "LENTES": lentes,
                "DESCRIÇÃO CONSUMO": descricao
            }) 

            dados_lentes4 = pd.DataFrame({
                "LENTES": lentes,
                "DESCRIÇÃO CONSUMO": descricao
            }) 

            dados_lentes5 = pd.DataFrame({
                "LENTES": lentes,
                "DESCRIÇÃO CONSUMO": descricao
            }) 

            def formatar_moeda(valor):
                # Formata manualmente os números no formato de moeda brasileira
                    return f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

            if AnodaApuração == 'Ano 1':
                def formatar_moeda(valor):
                # Formata manualmente os números no formato de moeda brasileira
                    return f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

                # Criando uma nova coluna para somar valores da base_historica com base nas repetições
                def obter_valor_somado(lente):
                    valores_correspondentes = BaseHistoricaFiltrada[BaseHistoricaFiltrada['Item 2'] == lente]['Total Gross']
                    return valores_correspondentes.sum() if not valores_correspondentes.empty else 0

                # Aplicando a função para calcular o valor total para cada lente
                dados_lentes['VALOR TOTAL'] = dados_lentes['LENTES'].apply(obter_valor_somado)

                # Removendo duplicados para evitar redundância
                dados_lentes = dados_lentes.drop_duplicates(subset=['LENTES'])

                soma_valor_total = dados_lentes['VALOR TOTAL'].sum()

                # Convertendo o valor somado para o formato de moeda brasileira
                valor_total_formatado = formatar_moeda(soma_valor_total)        
            elif AnodaApuração == 'Ano 2':
                

                # Criando uma nova coluna para somar valores da base_historica com base nas repetições
                def obter_valor_somado(lente):
                    valores_correspondentes = BaseHistoricaFiltrada[BaseHistoricaFiltrada['Item 2'] == lente]['Total Gross']
                    return valores_correspondentes.sum() if not valores_correspondentes.empty else 0

                def obter_valor_somado_Ano1(lente):
                    valores_correspondentes = BaseHistoricaFiltradaAno1[BaseHistoricaFiltradaAno1['Item 2'] == lente]['Total Gross']
                    return valores_correspondentes.sum() if not valores_correspondentes.empty else 0

                # Aplicando a função para calcular o valor total para cada lente
                dados_lentes['VALOR TOTAL'] = dados_lentes['LENTES'].apply(obter_valor_somado)

                dados_lentes2['VALOR TOTAL'] = dados_lentes2['LENTES'].apply(obter_valor_somado_Ano1)

                # Removendo duplicados para evitar redundância
                dados_lentes = dados_lentes.drop_duplicates(subset=['LENTES'])
                dados_lentes2 = dados_lentes2.drop_duplicates(subset=['LENTES'])

                soma_valor_total = dados_lentes['VALOR TOTAL'].sum()
                soma_valor_total_Ano1 = dados_lentes2['VALOR TOTAL'].sum()

                # Convertendo o valor somado para o formato de moeda brasileira
                valor_total_formatado = formatar_moeda(soma_valor_total)
                valor_total_formatadoAno1 = formatar_moeda(soma_valor_total_Ano1)
            elif AnodaApuração == 'Ano 3':
                
                def obter_valor_somado(lente):
                    valores_correspondentes = BaseHistoricaFiltrada[BaseHistoricaFiltrada['Item 2'] == lente]['Total Gross']
                    return valores_correspondentes.sum() if not valores_correspondentes.empty else 0

                def obter_valor_somado_Ano1(lente):
                    valores_correspondentes = BaseHistoricaFiltradaAno1[BaseHistoricaFiltradaAno1['Item 2'] == lente]['Total Gross']
                    return valores_correspondentes.sum() if not valores_correspondentes.empty else 0

                def obter_valor_somado_Ano2(lente):
                    valores_correspondentes = BaseHistoricaFiltradaAno2[BaseHistoricaFiltradaAno2['Item 2'] == lente]['Total Gross']
                    return valores_correspondentes.sum() if not valores_correspondentes.empty else 0
                # Aplicando a função para calcular o valor total para cada lente
                dados_lentes['VALOR TOTAL'] = dados_lentes['LENTES'].apply(obter_valor_somado)

                # Referente Ao Ano 1
                dados_lentes2['VALOR TOTAL'] = dados_lentes2['LENTES'].apply(obter_valor_somado_Ano1)
                # Referente Ao Ano 2
                dados_lentes3['VALOR TOTAL'] = dados_lentes3['LENTES'].apply(obter_valor_somado_Ano2)

                # Removendo duplicados para evitar redundância
                dados_lentes = dados_lentes.drop_duplicates(subset=['LENTES'])

                # Referente Ao Ano 1
                dados_lentes2 = dados_lentes2.drop_duplicates(subset=['LENTES'])

                # Referente Ao Ano 2
                dados_lentes3 = dados_lentes3.drop_duplicates(subset=['LENTES'])

                soma_valor_total = dados_lentes['VALOR TOTAL'].sum()
                soma_valor_total_Ano1 = dados_lentes2['VALOR TOTAL'].sum()
                soma_valor_total_Ano2 = dados_lentes3['VALOR TOTAL'].sum()

                # Convertendo o valor somado para o formato de moeda brasileira
                valor_total_formatado = formatar_moeda(soma_valor_total)
                valor_total_formatadoAno1 = formatar_moeda(soma_valor_total_Ano1)
                valor_total_formatadoAno2 = formatar_moeda(soma_valor_total_Ano2)
            elif AnodaApuração == 'Ano 4':
                
                def obter_valor_somado(lente):
                    valores_correspondentes = BaseHistoricaFiltrada[BaseHistoricaFiltrada['Item 2'] == lente]['Total Gross']
                    return valores_correspondentes.sum() if not valores_correspondentes.empty else 0

                def obter_valor_somado_Ano1(lente):
                    valores_correspondentes = BaseHistoricaFiltradaAno1[BaseHistoricaFiltradaAno1['Item 2'] == lente]['Total Gross']
                    return valores_correspondentes.sum() if not valores_correspondentes.empty else 0

                def obter_valor_somado_Ano2(lente):
                    valores_correspondentes = BaseHistoricaFiltradaAno2[BaseHistoricaFiltradaAno2['Item 2'] == lente]['Total Gross']
                    return valores_correspondentes.sum() if not valores_correspondentes.empty else 0
                def obter_valor_somado_Ano3(lente):
                    valores_correspondentes = BaseHistoricaFiltradaAno3[BaseHistoricaFiltradaAno3['Item 2'] == lente]['Total Gross']
                    return valores_correspondentes.sum() if not valores_correspondentes.empty else 0
                # Aplicando a função para calcular o valor total para cada lente
                dados_lentes['VALOR TOTAL'] = dados_lentes['LENTES'].apply(obter_valor_somado)

                # Referente Ao Ano 1
                dados_lentes2['VALOR TOTAL'] = dados_lentes2['LENTES'].apply(obter_valor_somado_Ano1)
                # Referente Ao Ano 2
                dados_lentes3['VALOR TOTAL'] = dados_lentes3['LENTES'].apply(obter_valor_somado_Ano2)
                # Referente Ao Ano 3
                dados_lentes4['VALOR TOTAL'] = dados_lentes4['LENTES'].apply(obter_valor_somado_Ano3)

                # Removendo duplicados para evitar redundância
                dados_lentes = dados_lentes.drop_duplicates(subset=['LENTES'])

                # Referente Ao Ano 1
                dados_lentes2 = dados_lentes2.drop_duplicates(subset=['LENTES'])

                # Referente Ao Ano 2
                dados_lentes3 = dados_lentes3.drop_duplicates(subset=['LENTES'])

                # Referente Ao Ano 3
                dados_lentes4 = dados_lentes4.drop_duplicates(subset=['LENTES'])

                soma_valor_total = dados_lentes['VALOR TOTAL'].sum()
                soma_valor_total_Ano1 = dados_lentes2['VALOR TOTAL'].sum()
                soma_valor_total_Ano2 = dados_lentes3['VALOR TOTAL'].sum()
                soma_valor_total_Ano3 = dados_lentes4['VALOR TOTAL'].sum()

                # Convertendo o valor somado para o formato de moeda brasileira
                valor_total_formatado = formatar_moeda(soma_valor_total)
                valor_total_formatadoAno1 = formatar_moeda(soma_valor_total_Ano1)
                valor_total_formatadoAno2 = formatar_moeda(soma_valor_total_Ano2)
                valor_total_formatadoAno3 = formatar_moeda(soma_valor_total_Ano3)
            elif AnodaApuração == 'Ano 5':
                
                def obter_valor_somado(lente):
                    valores_correspondentes = BaseHistoricaFiltrada[BaseHistoricaFiltrada['Item 2'] == lente]['Total Gross']
                    return valores_correspondentes.sum() if not valores_correspondentes.empty else 0

                def obter_valor_somado_Ano1(lente):
                    valores_correspondentes = BaseHistoricaFiltradaAno1[BaseHistoricaFiltradaAno1['Item 2'] == lente]['Total Gross']
                    return valores_correspondentes.sum() if not valores_correspondentes.empty else 0

                def obter_valor_somado_Ano2(lente):
                    valores_correspondentes = BaseHistoricaFiltradaAno2[BaseHistoricaFiltradaAno2['Item 2'] == lente]['Total Gross']
                    return valores_correspondentes.sum() if not valores_correspondentes.empty else 0
                def obter_valor_somado_Ano3(lente):
                    valores_correspondentes = BaseHistoricaFiltradaAno3[BaseHistoricaFiltradaAno3['Item 2'] == lente]['Total Gross']
                    return valores_correspondentes.sum() if not valores_correspondentes.empty else 0
                def obter_valor_somado_Ano4(lente):
                    valores_correspondentes = BaseHistoricaFiltradaAno4[BaseHistoricaFiltradaAno4['Item 2'] == lente]['Total Gross']
                    return valores_correspondentes.sum() if not valores_correspondentes.empty else 0
                # Aplicando a função para calcular o valor total para cada lente
                dados_lentes['VALOR TOTAL'] = dados_lentes['LENTES'].apply(obter_valor_somado)

                # Referente Ao Ano 1
                dados_lentes2['VALOR TOTAL'] = dados_lentes2['LENTES'].apply(obter_valor_somado_Ano1)
                # Referente Ao Ano 2
                dados_lentes3['VALOR TOTAL'] = dados_lentes3['LENTES'].apply(obter_valor_somado_Ano2)
                # Referente Ao Ano 3
                dados_lentes4['VALOR TOTAL'] = dados_lentes4['LENTES'].apply(obter_valor_somado_Ano3)
                # Referente Ao Ano 4
                dados_lentes5['VALOR TOTAL'] = dados_lentes5['LENTES'].apply(obter_valor_somado_Ano3)

                # Removendo duplicados para evitar redundância
                dados_lentes = dados_lentes.drop_duplicates(subset=['LENTES'])

                # Referente Ao Ano 1
                dados_lentes2 = dados_lentes2.drop_duplicates(subset=['LENTES'])

                # Referente Ao Ano 2
                dados_lentes3 = dados_lentes3.drop_duplicates(subset=['LENTES'])

                # Referente Ao Ano 3
                dados_lentes4 = dados_lentes4.drop_duplicates(subset=['LENTES'])
                # Referente Ao Ano 4

                dados_lentes5 = dados_lentes5.drop_duplicates(subset=['LENTES'])

                soma_valor_total = dados_lentes['VALOR TOTAL'].sum()
                soma_valor_total_Ano1 = dados_lentes2['VALOR TOTAL'].sum()
                soma_valor_total_Ano2 = dados_lentes3['VALOR TOTAL'].sum()
                soma_valor_total_Ano3 = dados_lentes4['VALOR TOTAL'].sum()
                soma_valor_total_Ano4 = dados_lentes5['VALOR TOTAL'].sum()

                # Convertendo o valor somado para o formato de moeda brasileira
                valor_total_formatado = formatar_moeda(soma_valor_total)
                valor_total_formatadoAno1 = formatar_moeda(soma_valor_total_Ano1)
                valor_total_formatadoAno2 = formatar_moeda(soma_valor_total_Ano2)
                valor_total_formatadoAno3 = formatar_moeda(soma_valor_total_Ano3)
                valor_total_formatadoAno4 = formatar_moeda(soma_valor_total_Ano4)
            # Calculando altura da página e ajustando layout para o relatório
            altura_pagina_Consumo = 400

            # Convertendo para uma lista de listas para o ReportLab
            
            


        # Função para validar e formatar um valor
            def validar_formatar_consumo(valor):
                if pd.notna(valor):  # Verifica se o valor não é NaN
                    MultaSemformatar = valor
                    return formatar_moeda(valor)  # Formata para moeda brasileira
                return ""  # Retorna valor padrão
            
            def tratar_valor_para_soma_dos_valores(valor):
                if pd.notna(valor):  # Verifica se o valor não é NaN
                    return valor
                return 0  # Substitui valores inválidos por 0 para soma]
            
            def validar_formatar_consumo_total(valor):
                if pd.notna(valor) and valor != 0:  # Verifica se o valor não é NaN e não é igual a 0
                    return formatar_moeda(valor)  # Formata para moeda brasileira
                return 0
            def somar_sem_perder_valor(a, b):
                if a and not b:  # Se 'a' tem valor e 'b' é None ou 0
                    return a
                if b and not a:  # Se 'b' tem valor e 'a' é None ou 0
                    return b
                return (a or 0) + (b or 0)  #

            # Inicializa as variáveis como valores padrão
            ValorConsumoAno1Longo = ValorConsumoAno2Longo = ValorConsumoAno3Longo = 0
            ValorConsumoAno4Longo = ValorConsumoAno5Longo = 0

            ValorConsumoAno1Curto = ValorConsumoAno2Curto = ValorConsumoAno3Curto = 0
            ValorConsumoAno4Curto = ValorConsumoAno5Curto = 0

            consumo_ano_1_curto = 0
            consumo_ano_2_curto = 0
            consumo_ano_3_curto = 0
            consumo_ano_4_curto = 0
            consumo_ano_5_curto = 0

            

            # Verifica e processa BaseLongoFiltradoCliente
            if not BaseLongoFiltradoCliente.empty and len(BaseLongoFiltradoCliente) > 0:
                consumo_ano_1 = pd.to_numeric(BaseLongoFiltradoCliente.iloc[0].get('CONSUMO ANO 1'), errors='coerce')
                consumo_ano_2 = pd.to_numeric(BaseLongoFiltradoCliente.iloc[0].get('CONSUMO ANO 2'), errors='coerce')
                consumo_ano_3 = pd.to_numeric(BaseLongoFiltradoCliente.iloc[0].get('CONSUMO ANO 3'), errors='coerce')
                consumo_ano_4 = pd.to_numeric(BaseLongoFiltradoCliente.iloc[0].get('CONSUMO ANO 4'), errors='coerce')
                consumo_ano_5 = pd.to_numeric(BaseLongoFiltradoCliente.iloc[0].get('CONSUMO ANO 5'), errors='coerce')

                ValorConsumoAno1Longo = validar_formatar_consumo(consumo_ano_1)
                ValorConsumoAno2Longo = validar_formatar_consumo(consumo_ano_2)
                ValorConsumoAno3Longo = validar_formatar_consumo(consumo_ano_3)
                ValorConsumoAno4Longo = validar_formatar_consumo(consumo_ano_4)
                ValorConsumoAno5Longo = validar_formatar_consumo(consumo_ano_5)

            # Verifica e processa BaseCurtoFiltradoCliente
            if not BaseCurtoFiltradoCliente.empty and len(BaseCurtoFiltradoCliente) > 0:

                consumo_ano_1_curto = 0
                consumo_ano_2_curto = 0
                consumo_ano_3_curto = 0
                consumo_ano_4_curto = 0
                consumo_ano_5_curto = 0

                consumo_ano_1_curto = pd.to_numeric(BaseCurtoFiltradoCliente.iloc[0].get('CONSUMO ANO 1'), errors='coerce')
                consumo_ano_2_curto = pd.to_numeric(BaseCurtoFiltradoCliente.iloc[0].get('CONSUMO ANO 2'), errors='coerce')
                consumo_ano_3_curto = pd.to_numeric(BaseCurtoFiltradoCliente.iloc[0].get('CONSUMO ANO 3'), errors='coerce')
                consumo_ano_4_curto = pd.to_numeric(BaseCurtoFiltradoCliente.iloc[0].get('CONSUMO ANO 4'), errors='coerce')
                consumo_ano_5_curto = pd.to_numeric(BaseCurtoFiltradoCliente.iloc[0].get('CONSUMO ANO 5'), errors='coerce')

                ValorConsumoAno1Curto = validar_formatar_consumo(consumo_ano_1_curto)
                ValorConsumoAno2Curto = validar_formatar_consumo(consumo_ano_2_curto)
                ValorConsumoAno3Curto = validar_formatar_consumo(consumo_ano_3_curto)
                ValorConsumoAno4Curto = validar_formatar_consumo(consumo_ano_4_curto)
                ValorConsumoAno5Curto = validar_formatar_consumo(consumo_ano_5_curto)

            
            if not ClientesManutenção.empty and len(ClientesManutenção) > 0: 

                consumo_ano_1_Manutenção = 0
                consumo_ano_2_Manutenção = 0
                consumo_ano_3_Manutenção = 0
                consumo_ano_4_Manutenção = 0
                consumo_ano_5_Manutenção = 0

                consumo_ano_1_Manutenção = pd.to_numeric(ClientesManutenção.iloc[0].get('CONSUMO ANO 1'), errors='coerce')
                consumo_ano_2_Manutenção = pd.to_numeric(ClientesManutenção.iloc[0].get('CONSUMO ANO 2'), errors='coerce')
                consumo_ano_3_Manutenção = pd.to_numeric(ClientesManutenção.iloc[0].get('CONSUMO ANO 3'), errors='coerce')
                consumo_ano_4_Manutenção = pd.to_numeric(ClientesManutenção.iloc[0].get('CONSUMO ANO 4'), errors='coerce')
                consumo_ano_5_Manutenção = pd.to_numeric(ClientesManutenção.iloc[0].get('CONSUMO ANO 5'), errors='coerce')

                ValorConsumoAno1Manutenção = validar_formatar_consumo(consumo_ano_1_Manutenção)
                ValorConsumoAno2Manutenção = validar_formatar_consumo(consumo_ano_2_Manutenção)
                ValorConsumoAno3Manutenção = validar_formatar_consumo(consumo_ano_3_Manutenção)
                ValorConsumoAno4Manutenção = validar_formatar_consumo(consumo_ano_4_Manutenção)
                ValorConsumoAno5Manutenção = validar_formatar_consumo(consumo_ano_5_Manutenção)

            if not ClientesNovoComodato.empty and len(ClientesNovoComodato) > 0: 

                consumo_ano_1_NovoComodato = 0
                consumo_ano_2_NovoComodato = 0
                consumo_ano_3_NovoComodato = 0
                consumo_ano_4_NovoComodato = 0
                consumo_ano_5_NovoComodato = 0

                consumo_ano_1_NovoComodato = pd.to_numeric(ClientesNovoComodato.iloc[0].get('CONSUMO ANO 1'), errors='coerce')
                consumo_ano_2_NovoComodato = pd.to_numeric(ClientesNovoComodato.iloc[0].get('CONSUMO ANO 2'), errors='coerce')
                consumo_ano_3_NovoComodato = pd.to_numeric(ClientesNovoComodato.iloc[0].get('CONSUMO ANO 3'), errors='coerce')
                consumo_ano_4_NovoComodato = pd.to_numeric(ClientesNovoComodato.iloc[0].get('CONSUMO ANO 4'), errors='coerce')
                consumo_ano_5_NovoComodato = pd.to_numeric(ClientesNovoComodato.iloc[0].get('CONSUMO ANO 5'), errors='coerce')

                ValorConsumoAno1NovoComodato = validar_formatar_consumo(consumo_ano_1_NovoComodato)
                ValorConsumoAno2NovoComodato = validar_formatar_consumo(consumo_ano_2_NovoComodato)
                ValorConsumoAno3NovoComodato = validar_formatar_consumo(consumo_ano_3_NovoComodato)
                ValorConsumoAno4NovoComodato = validar_formatar_consumo(consumo_ano_4_NovoComodato)
                ValorConsumoAno5NovoComodato = validar_formatar_consumo(consumo_ano_5_NovoComodato)

            if not ClientesacordodeConsumo.empty and len(ClientesacordodeConsumo) > 0: 

                consumo_ano_1_AcordoDeConsumo = 0
                consumo_ano_2_AcordoDeConsumo = 0
                consumo_ano_3_AcordoDeConsumo = 0
                consumo_ano_4_AcordoDeConsumo = 0
                consumo_ano_5_AcordoDeConsumo = 0

                consumo_ano_1_AcordoDeConsumo = pd.to_numeric(ClientesacordodeConsumo.iloc[0].get('CONSUMO ANO 1'), errors='coerce')
                consumo_ano_2_AcordoDeConsumo = pd.to_numeric(ClientesacordodeConsumo.iloc[0].get('CONSUMO ANO 2'), errors='coerce')
                consumo_ano_3_AcordoDeConsumo = pd.to_numeric(ClientesacordodeConsumo.iloc[0].get('CONSUMO ANO 3'), errors='coerce')
                consumo_ano_4_AcordoDeConsumo = pd.to_numeric(ClientesacordodeConsumo.iloc[0].get('CONSUMO ANO 4'), errors='coerce')
                consumo_ano_5_AcordoDeConsumo = pd.to_numeric(ClientesacordodeConsumo.iloc[0].get('CONSUMO ANO 5'), errors='coerce')

                ValorConsumoAno1AcordoDeConsumo = validar_formatar_consumo(consumo_ano_1_AcordoDeConsumo)
                ValorConsumoAno2AcordoDeConsumo = validar_formatar_consumo(consumo_ano_2_AcordoDeConsumo)
                ValorConsumoAno3AcordoDeConsumo = validar_formatar_consumo(consumo_ano_3_AcordoDeConsumo)
                ValorConsumoAno4AcordoDeConsumo = validar_formatar_consumo(consumo_ano_4_AcordoDeConsumo)
                ValorConsumoAno5AcordoDeConsumo = validar_formatar_consumo(consumo_ano_5_AcordoDeConsumo)

            consumo_ano_1_total = 0
            consumo_ano_2_total = 0
            consumo_ano_3_total = 0
            consumo_ano_4_total = 0
            consumo_ano_5_total = 0


            if ('BaseCurtoFiltradoCliente' in locals() and 'BaseLongoFiltradoCliente' in locals() and
                    not BaseCurtoFiltradoCliente.empty and not BaseLongoFiltradoCliente.empty):

                # Validar e somar os valores para cada ano
                consumo_ano_1_total = somar_sem_perder_valor(
                    tratar_valor_para_soma_dos_valores(BaseCurtoFiltradoCliente.iloc[0].get('CONSUMO ANO 1')),
                    tratar_valor_para_soma_dos_valores(BaseLongoFiltradoCliente.iloc[0].get('CONSUMO ANO 1'))
                )

                consumo_ano_2_total = somar_sem_perder_valor(
                    tratar_valor_para_soma_dos_valores(consumo_ano_2),
                    tratar_valor_para_soma_dos_valores(consumo_ano_2_curto)
                )

                consumo_ano_3_total = somar_sem_perder_valor(
                    tratar_valor_para_soma_dos_valores(consumo_ano_3),
                    tratar_valor_para_soma_dos_valores(consumo_ano_3_curto)
                )

                consumo_ano_4_total = somar_sem_perder_valor(
                    tratar_valor_para_soma_dos_valores(consumo_ano_4),
                    tratar_valor_para_soma_dos_valores(consumo_ano_4_curto)
                )

                consumo_ano_5_total = somar_sem_perder_valor(
                    tratar_valor_para_soma_dos_valores(consumo_ano_5),
                    tratar_valor_para_soma_dos_valores(consumo_ano_5_curto)
                )


            ValorConsumoTotalAno1 = validar_formatar_consumo_total(consumo_ano_1_total)
            ValorConsumoTotalAno2 = validar_formatar_consumo_total(consumo_ano_2_total)
            ValorConsumoTotalAno3 = validar_formatar_consumo_total(consumo_ano_3_total)
            ValorConsumoTotalAno4 = validar_formatar_consumo_total(consumo_ano_4_total)
            ValorConsumoTotalAno5 = validar_formatar_consumo_total(consumo_ano_5_total)
            



            def calcular_porcentagem_a_mais(valor_planejado, valor_comprado):
                """Calcula quantos por cento a mais foi comprado."""
                if valor_planejado == 0:
                    return 0  # Ou outro valor padrão que faça sentido para você
                
                variacao = (valor_comprado  / valor_planejado) * 100
                return round(variacao, 2)  # Limita a 2 casas decimais

            # Valores 
            if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                if AnodaApuração == "Ano 1":
                    valor_planejado = consumo_ano_1_total
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                if AnodaApuração == "Ano 2":
                    valor_planejado = consumo_ano_2_total
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_total
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)


                if AnodaApuração == "Ano 3":
                    valor_planejado = consumo_ano_3_total
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_total
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                    valor_planejadoAno2 = consumo_ano_2_total
                    valor_compradoAno2 = soma_valor_total_Ano2

                    # Calculando a porcentagem a mais
                    resultadoAno2 = calcular_porcentagem_a_mais(valor_planejadoAno2, valor_compradoAno2)

                if AnodaApuração == "Ano 4":
                    valor_planejado = consumo_ano_4_total
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_total
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                    valor_planejadoAno2 = consumo_ano_2_total
                    valor_compradoAno2 = soma_valor_total_Ano2

                    # Calculando a porcentagem a mais
                    resultadoAno2 = calcular_porcentagem_a_mais(valor_planejadoAno2, valor_compradoAno2)

                    valor_planejadoAno3 = consumo_ano_3_total
                    valor_compradoAno3 = soma_valor_total_Ano3

                    # Calculando a porcentagem a mais
                    resultadoAno3 = calcular_porcentagem_a_mais(valor_planejadoAno3, valor_compradoAno3)
                
                if AnodaApuração == "Ano 5":
                    valor_planejado = consumo_ano_5_total
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_total
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                    valor_planejadoAno2 = consumo_ano_2_total
                    valor_compradoAno2 = soma_valor_total_Ano2

                    # Calculando a porcentagem a mais
                    resultadoAno2 = calcular_porcentagem_a_mais(valor_planejadoAno2, valor_compradoAno2)

                    valor_planejadoAno3 = consumo_ano_3_total
                    valor_compradoAno3 = soma_valor_total_Ano3

                    # Calculando a porcentagem a mais
                    resultadoAno3 = calcular_porcentagem_a_mais(valor_planejadoAno3, valor_compradoAno3)

                    valor_planejadoAno4 = consumo_ano_4_total
                    valor_compradoAno4 = soma_valor_total_Ano4

                    # Calculando a porcentagem a mais
                    resultadoAno4 = calcular_porcentagem_a_mais(valor_planejadoAno4, valor_compradoAno4)



            elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:
                if AnodaApuração == "Ano 1":
                    valor_planejado = consumo_ano_1
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                if AnodaApuração == "Ano 2":
                    valor_planejado = consumo_ano_2
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                if AnodaApuração == "Ano 3":
                    valor_planejado = consumo_ano_3
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                    valor_planejadoAno2 = consumo_ano_2
                    valor_compradoAno2 = soma_valor_total_Ano2

                    # Calculando a porcentagem a mais
                    resultadoAno2 = calcular_porcentagem_a_mais(valor_planejadoAno2, valor_compradoAno2)

                if AnodaApuração == "Ano 4":
                    valor_planejado = consumo_ano_4
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                    valor_planejadoAno2 = consumo_ano_2
                    valor_compradoAno2 = soma_valor_total_Ano2

                    # Calculando a porcentagem a mais
                    resultadoAno2 = calcular_porcentagem_a_mais(valor_planejadoAno2, valor_compradoAno2)

                    valor_planejadoAno3 = consumo_ano_3
                    valor_compradoAno3 = soma_valor_total_Ano3

                    # Calculando a porcentagem a mais
                    resultadoAno3 = calcular_porcentagem_a_mais(valor_planejadoAno3, valor_compradoAno3)
                
                if AnodaApuração == "Ano 5":
                    valor_planejado = consumo_ano_5
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                    valor_planejadoAno2 = consumo_ano_2
                    valor_compradoAno2 = soma_valor_total_Ano2

                    # Calculando a porcentagem a mais
                    resultadoAno2 = calcular_porcentagem_a_mais(valor_planejadoAno2, valor_compradoAno2)

                    valor_planejadoAno3 = consumo_ano_3
                    valor_compradoAno3 = soma_valor_total_Ano3

                    # Calculando a porcentagem a mais
                    resultadoAno3 = calcular_porcentagem_a_mais(valor_planejadoAno3, valor_compradoAno3)

                    valor_planejadoAno4 = consumo_ano_4
                    valor_compradoAno4 = soma_valor_total_Ano4

                    # Calculando a porcentagem a mais
                    resultadoAno4 = calcular_porcentagem_a_mais(valor_planejadoAno4, valor_compradoAno4)

            elif not BaseCurtoFiltradoCliente.empty and BaseLongoFiltradoCliente.empty:
                if AnodaApuração == "Ano 1":
                    valor_planejado = consumo_ano_1_curto
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                if AnodaApuração == "Ano 2":
                    valor_planejado = consumo_ano_2_curto
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_curto
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                if AnodaApuração == "Ano 3":
                    valor_planejado = consumo_ano_3_curto
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_curto
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                    valor_planejadoAno2 = consumo_ano_2_curto
                    valor_compradoAno2 = soma_valor_total_Ano2

                    # Calculando a porcentagem a mais
                    resultadoAno2 = calcular_porcentagem_a_mais(valor_planejadoAno2, valor_compradoAno2)

                if AnodaApuração == "Ano 4":
                    valor_planejado = consumo_ano_4_curto
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_curto
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                    valor_planejadoAno2 = consumo_ano_2_curto
                    valor_compradoAno2 = soma_valor_total_Ano2

                    # Calculando a porcentagem a mais
                    resultadoAno2 = calcular_porcentagem_a_mais(valor_planejadoAno2, valor_compradoAno2)

                    valor_planejadoAno3 = consumo_ano_3_curto
                    valor_compradoAno3 = soma_valor_total_Ano3

                    # Calculando a porcentagem a mais
                    resultadoAno3 = calcular_porcentagem_a_mais(valor_planejadoAno3, valor_compradoAno3)
                
                if AnodaApuração == "Ano 5":
                    valor_planejado = consumo_ano_5_curto
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_curto
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                    valor_planejadoAno2 = consumo_ano_2_curto
                    valor_compradoAno2 = soma_valor_total_Ano2

                    # Calculando a porcentagem a mais
                    resultadoAno2 = calcular_porcentagem_a_mais(valor_planejadoAno2, valor_compradoAno2)

                    valor_planejadoAno3 = consumo_ano_3_curto
                    valor_compradoAno3 = soma_valor_total_Ano3

                    # Calculando a porcentagem a mais
                    resultadoAno3 = calcular_porcentagem_a_mais(valor_planejadoAno3, valor_compradoAno3)

                    valor_planejadoAno4 = consumo_ano_4_curto
                    valor_compradoAno4 = soma_valor_total_Ano4

                    # Calculando a porcentagem a mais
                    resultadoAno4 = calcular_porcentagem_a_mais(valor_planejadoAno4, valor_compradoAno4)

            if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                if AnodaApuração == "Ano 1":
                    valor_planejado = consumo_ano_1_total
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                if AnodaApuração == "Ano 2":
                    valor_planejado = consumo_ano_2_total
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_total
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)


                if AnodaApuração == "Ano 3":
                    valor_planejado = consumo_ano_3_total
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_total
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                    valor_planejadoAno2 = consumo_ano_2_total
                    valor_compradoAno2 = soma_valor_total_Ano2

                    # Calculando a porcentagem a mais
                    resultadoAno2 = calcular_porcentagem_a_mais(valor_planejadoAno2, valor_compradoAno2)

                if AnodaApuração == "Ano 4":
                    valor_planejado = consumo_ano_4_total
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_total
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                    valor_planejadoAno2 = consumo_ano_2_total
                    valor_compradoAno2 = soma_valor_total_Ano2

                    # Calculando a porcentagem a mais
                    resultadoAno2 = calcular_porcentagem_a_mais(valor_planejadoAno2, valor_compradoAno2)

                    valor_planejadoAno3 = consumo_ano_3_total
                    valor_compradoAno3 = soma_valor_total_Ano3

                    # Calculando a porcentagem a mais
                    resultadoAno3 = calcular_porcentagem_a_mais(valor_planejadoAno3, valor_compradoAno3)
                
                if AnodaApuração == "Ano 5":
                    valor_planejado = consumo_ano_5_total
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_total
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                    valor_planejadoAno2 = consumo_ano_2_total
                    valor_compradoAno2 = soma_valor_total_Ano2

                    # Calculando a porcentagem a mais
                    resultadoAno2 = calcular_porcentagem_a_mais(valor_planejadoAno2, valor_compradoAno2)

                    valor_planejadoAno3 = consumo_ano_3_total
                    valor_compradoAno3 = soma_valor_total_Ano3

                    # Calculando a porcentagem a mais
                    resultadoAno3 = calcular_porcentagem_a_mais(valor_planejadoAno3, valor_compradoAno3)

                    valor_planejadoAno4 = consumo_ano_4_total
                    valor_compradoAno4 = soma_valor_total_Ano4

                    # Calculando a porcentagem a mais
                    resultadoAno4 = calcular_porcentagem_a_mais(valor_planejadoAno4, valor_compradoAno4)



            elif not ClientesManutenção.empty:
                if AnodaApuração == "Ano 1":
                    valor_planejado = consumo_ano_1_Manutenção
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                if AnodaApuração == "Ano 2":
                    valor_planejado = consumo_ano_2_Manutenção
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_Manutenção
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                if AnodaApuração == "Ano 3":
                    valor_planejado = consumo_ano_3_Manutenção
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_Manutenção
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                    valor_planejadoAno2 = consumo_ano_2_Manutenção
                    valor_compradoAno2 = soma_valor_total_Ano2

                    # Calculando a porcentagem a mais
                    resultadoAno2 = calcular_porcentagem_a_mais(valor_planejadoAno2, valor_compradoAno2)

                if AnodaApuração == "Ano 4":
                    valor_planejado = consumo_ano_4_Manutenção
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_Manutenção
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                    valor_planejadoAno2 = consumo_ano_2_Manutenção
                    valor_compradoAno2 = soma_valor_total_Ano2

                    # Calculando a porcentagem a mais
                    resultadoAno2 = calcular_porcentagem_a_mais(valor_planejadoAno2, valor_compradoAno2)

                    valor_planejadoAno3 = consumo_ano_3_Manutenção
                    valor_compradoAno3 = soma_valor_total_Ano3

                    # Calculando a porcentagem a mais
                    resultadoAno3 = calcular_porcentagem_a_mais(valor_planejadoAno3, valor_compradoAno3)
                
                if AnodaApuração == "Ano 5":
                    valor_planejado = consumo_ano_5_Manutenção
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_Manutenção
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                    valor_planejadoAno2 = consumo_ano_2_Manutenção
                    valor_compradoAno2 = soma_valor_total_Ano2

                    # Calculando a porcentagem a mais
                    resultadoAno2 = calcular_porcentagem_a_mais(valor_planejadoAno2, valor_compradoAno2)

                    valor_planejadoAno3 = consumo_ano_3_Manutenção
                    valor_compradoAno3 = soma_valor_total_Ano3

                    # Calculando a porcentagem a mais
                    resultadoAno3 = calcular_porcentagem_a_mais(valor_planejadoAno3, valor_compradoAno3)

                    valor_planejadoAno4 = consumo_ano_4_Manutenção
                    valor_compradoAno4 = soma_valor_total_Ano4

                    # Calculando a porcentagem a mais
                    resultadoAno4 = calcular_porcentagem_a_mais(valor_planejadoAno4, valor_compradoAno4)

            
            elif not ClientesNovoComodato.empty:
                if AnodaApuração == "Ano 1":
                    valor_planejado = consumo_ano_1_NovoComodato
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                if AnodaApuração == "Ano 2":
                    valor_planejado = consumo_ano_2_NovoComodato
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_NovoComodato
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                if AnodaApuração == "Ano 3":
                    valor_planejado = consumo_ano_3_NovoComodato
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_NovoComodato
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                    valor_planejadoAno2 = consumo_ano_2_NovoComodato
                    valor_compradoAno2 = soma_valor_total_Ano2

                    # Calculando a porcentagem a mais
                    resultadoAno2 = calcular_porcentagem_a_mais(valor_planejadoAno2, valor_compradoAno2)

                if AnodaApuração == "Ano 4":
                    valor_planejado = consumo_ano_4_NovoComodato
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_NovoComodato
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                    valor_planejadoAno2 = consumo_ano_2_NovoComodato
                    valor_compradoAno2 = soma_valor_total_Ano2

                    # Calculando a porcentagem a mais
                    resultadoAno2 = calcular_porcentagem_a_mais(valor_planejadoAno2, valor_compradoAno2)

                    valor_planejadoAno3 = consumo_ano_3_NovoComodato
                    valor_compradoAno3 = soma_valor_total_Ano3

                    # Calculando a porcentagem a mais
                    resultadoAno3 = calcular_porcentagem_a_mais(valor_planejadoAno3, valor_compradoAno3)
                
                if AnodaApuração == "Ano 5":
                    valor_planejado = consumo_ano_5_NovoComodato
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_NovoComodato
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                    valor_planejadoAno2 = consumo_ano_2_NovoComodato
                    valor_compradoAno2 = soma_valor_total_Ano2

                    # Calculando a porcentagem a mais
                    resultadoAno2 = calcular_porcentagem_a_mais(valor_planejadoAno2, valor_compradoAno2)

                    valor_planejadoAno3 = consumo_ano_3_NovoComodato
                    valor_compradoAno3 = soma_valor_total_Ano3

                    # Calculando a porcentagem a mais
                    resultadoAno3 = calcular_porcentagem_a_mais(valor_planejadoAno3, valor_compradoAno3)

                    valor_planejadoAno4 = consumo_ano_4_NovoComodato
                    valor_compradoAno4 = soma_valor_total_Ano4

                    # Calculando a porcentagem a mais
                    resultadoAno4 = calcular_porcentagem_a_mais(valor_planejadoAno4, valor_compradoAno4)

            elif not ClientesacordodeConsumo.empty:
                if AnodaApuração == "Ano 1":
                    valor_planejado = consumo_ano_1_AcordoDeConsumo
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                if AnodaApuração == "Ano 2":
                    valor_planejado = consumo_ano_2_AcordoDeConsumo
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_AcordoDeConsumo
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                if AnodaApuração == "Ano 3":
                    valor_planejado = consumo_ano_3_AcordoDeConsumo
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_AcordoDeConsumo
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                    valor_planejadoAno2 = consumo_ano_2_AcordoDeConsumo
                    valor_compradoAno2 = soma_valor_total_Ano2

                    # Calculando a porcentagem a mais
                    resultadoAno2 = calcular_porcentagem_a_mais(valor_planejadoAno2, valor_compradoAno2)

                if AnodaApuração == "Ano 4":
                    valor_planejado = consumo_ano_4_AcordoDeConsumo
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_AcordoDeConsumo
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                    valor_planejadoAno2 = consumo_ano_2_AcordoDeConsumo
                    valor_compradoAno2 = soma_valor_total_Ano2

                    # Calculando a porcentagem a mais
                    resultadoAno2 = calcular_porcentagem_a_mais(valor_planejadoAno2, valor_compradoAno2)

                    valor_planejadoAno3 = consumo_ano_3_AcordoDeConsumo
                    valor_compradoAno3 = soma_valor_total_Ano3

                    # Calculando a porcentagem a mais
                    resultadoAno3 = calcular_porcentagem_a_mais(valor_planejadoAno3, valor_compradoAno3)
                
                if AnodaApuração == "Ano 5":
                    valor_planejado = consumo_ano_5_AcordoDeConsumo
                    valor_comprado = soma_valor_total

                    # Calculando a porcentagem a mais
                    resultado = calcular_porcentagem_a_mais(valor_planejado, valor_comprado)

                    valor_planejadoAno1 = consumo_ano_1_AcordoDeConsumo
                    valor_compradoAno1 = soma_valor_total_Ano1

                    # Calculando a porcentagem a mais
                    resultadoAno1 = calcular_porcentagem_a_mais(valor_planejadoAno1, valor_compradoAno1)

                    valor_planejadoAno2 = consumo_ano_2_AcordoDeConsumo
                    valor_compradoAno2 = soma_valor_total_Ano2

                    # Calculando a porcentagem a mais
                    resultadoAno2 = calcular_porcentagem_a_mais(valor_planejadoAno2, valor_compradoAno2)

                    valor_planejadoAno3 = consumo_ano_3_AcordoDeConsumo
                    valor_compradoAno3 = soma_valor_total_Ano3

                    # Calculando a porcentagem a mais
                    resultadoAno3 = calcular_porcentagem_a_mais(valor_planejadoAno3, valor_compradoAno3)

                    valor_planejadoAno4 = consumo_ano_4_AcordoDeConsumo
                    valor_compradoAno4 = soma_valor_total_Ano4

                    # Calculando a porcentagem a mais
                    resultadoAno4 = calcular_porcentagem_a_mais(valor_planejadoAno4, valor_compradoAno4)

            if AnodaApuração == 'Ano 1':
                if resultado >= 100:
                    StyleTabelaTarget = TableStyle([
                    # Estilo geral
                    ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 0), (-1, -1), 5),  # Tamanho da fonte geral
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # Alinhamento à esquerda
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Centralização vertical
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno inferior
                    ('TOPPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno superior

                    # Estilo para a linha de cabeçalho
                    ('SPAN', (0, 0), (-1, 0)),  # Mescla as colunas na primeira linha
                    ('BACKGROUND', (0, 0), (-1, 0), (192/255, 0/255, 10/255)),  # Fundo vermelho
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),  # Texto branco
                    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Alinhamento centralizado
                    ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),  # Linha acima da primeira linha
                    ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),  # Linha abaixo da primeira linha
                    ('FONTSIZE', (0, 0), (-1, 0), 6),  # Aumenta o tamanho da fonte na primeira linha
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Deixa a primeira linha em negrito

                    # Estilo para a segunda linha
                    ('BACKGROUND', (0, 1), (-1, 1), (192/255, 0/255, 10/255)),  # Fundo vermelho
                    ('TEXTCOLOR', (0, 1), (-1, 1), colors.white),  # Texto branco
                    ('ALIGN', (0, 1), (-1, 1), 'CENTER'),  # Alinhamento centralizado

                    # Estilo para a terceira linha
                    ('BACKGROUND', (0, 2), (-1, 2), (230/255, 230/255, 230/255)),  # Fundo cinza claro
                    ('TEXTCOLOR', (0, 2), (-1, 2), colors.black),  # Texto preto
                    ('ALIGN', (0, 2), (-1, 2), 'LEFT'),  # Alinhamento à esquerda

                    # Estilo para a primeira coluna a partir da terceira linha
                    ('BACKGROUND', (0, 2), (0, -1), (230/255, 230/255, 230/255)),  # Fundo cinza claro
                    ('TEXTCOLOR', (0, 2), (0, -1), colors.black),  # Texto preto

                    # Estilo para os títulos das linhas restantes
                    ('TEXTCOLOR', (0, 3), (0, -1), colors.black),  # Texto preto
                    ('FONTNAME', (0, 3), (0, -1), 'Helvetica-Bold'),  # Texto em negrito

                    # Fundo das células de conteúdo restante
                    ('BACKGROUND', (1, 3), (-1, -1), colors.white),  # Fundo branco

                        # Deixar a **coluna 2 da terceira linha pra baixo** em negrito
                    ('FONTNAME', (1, 2), (1, -1), 'Helvetica-Bold'),

                    # Divisões da tabela com cinza claro
                    ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Grade interna
                    ('BOX', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Borda externa
                ])
                else:
                    StyleTabelaTarget = TableStyle([
                    # Estilo geral
                    ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 0), (-1, -1), 5),  # Tamanho da fonte geral
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # Alinhamento à esquerda
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Centralização vertical
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno inferior
                    ('TOPPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno superior

                    # Estilo para a linha de cabeçalho
                    ('SPAN', (0, 0), (-1, 0)),  # Mescla as colunas na primeira linha
                    ('BACKGROUND', (0, 0), (-1, 0), (192/255, 0/255, 10/255)),  # Fundo vermelho
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),  # Texto branco
                    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Alinhamento centralizado
                    ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),  # Linha acima da primeira linha
                    ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),  # Linha abaixo da primeira linha
                    ('FONTSIZE', (0, 0), (-1, 0), 6),  # Aumenta o tamanho da fonte na primeira linha
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Deixa a primeira linha em negrito

                    # Estilo para a segunda linha
                    ('BACKGROUND', (0, 1), (-1, 1), (192/255, 0/255, 10/255)),  # Fundo vermelho
                    ('TEXTCOLOR', (0, 1), (-1, 1), colors.white),  # Texto branco
                    ('ALIGN', (0, 1), (-1, 1), 'CENTER'),  # Alinhamento centralizado

                    # Estilo para a terceira linha
                    ('BACKGROUND', (0, 2), (-1, 2), (230/255, 230/255, 230/255)),  # Fundo cinza claro
                    ('TEXTCOLOR', (0, 2), (-1, 2), colors.black),  # Texto preto
                    ('ALIGN', (0, 2), (-1, 2), 'LEFT'),  # Alinhamento à esquerda

                    # Estilo para a primeira coluna a partir da terceira linha
                    ('BACKGROUND', (0, 2), (0, -1), (230/255, 230/255, 230/255)),  # Fundo cinza claro
                    ('TEXTCOLOR', (0, 2), (0, -1), colors.black),  # Texto preto

                    # Estilo para os títulos das linhas restantes
                    ('TEXTCOLOR', (0, 3), (0, -1), colors.black),  # Texto preto
                    ('FONTNAME', (0, 3), (0, -1), 'Helvetica-Bold'),  # Texto em negrito

                    # Fundo das células de conteúdo restante
                    ('BACKGROUND', (1, 3), (-1, -1), colors.white),  # Fundo branco

                        # Deixar a **coluna 2 da terceira linha pra baixo** em negrito
                    ('FONTNAME', (1, 2), (1, -1), 'Helvetica-Bold'),

                    # Divisões da tabela com cinza claro
                    ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Grade interna
                    ('BOX', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Borda e

                    ('TEXTCOLOR', (1, -1), (1, -1), (192/255, 0/255, 10/255)),  # Texto vermelho
                    ('FONTNAME', (1, -1), (1, -1), 'Helvetica-Bold'),  # Texto em negrito
                    ])
            elif AnodaApuração == 'Ano 2':
                    if resultado >= 100:
                        StyleTabelaTarget = TableStyle([
                        # Estilo geral
                        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 0), (-1, -1), 5),  # Tamanho da fonte geral
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # Alinhamento à esquerda
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Centralização vertical
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno inferior
                        ('TOPPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno superior

                        # Estilo para a linha de cabeçalho
                        ('SPAN', (0, 0), (-1, 0)),  # Mescla as colunas na primeira linha
                        ('BACKGROUND', (0, 0), (-1, 0), (192/255, 0/255, 10/255)),  # Fundo vermelho
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),  # Texto branco
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Alinhamento centralizado
                        ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),  # Linha acima da primeira linha
                        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),  # Linha abaixo da primeira linha
                        ('FONTSIZE', (0, 0), (-1, 0), 6),  # Aumenta o tamanho da fonte na primeira linha
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Deixa a primeira linha em negrito

                        # Estilo para a segunda linha
                        ('BACKGROUND', (0, 1), (-1, 1), (192/255, 0/255, 10/255)),  # Fundo vermelho
                        ('TEXTCOLOR', (0, 1), (-1, 1), colors.white),  # Texto branco
                        ('ALIGN', (0, 1), (-1, 1), 'CENTER'),  # Alinhamento centralizado

                        # Estilo para a terceira linha
                        ('BACKGROUND', (0, 2), (-1, 2), (230/255, 230/255, 230/255)),  # Fundo cinza claro
                        ('TEXTCOLOR', (0, 2), (-1, 2), colors.black),  # Texto preto
                        ('ALIGN', (0, 2), (-1, 2), 'LEFT'),  # Alinhamento à esquerda

                        # Estilo para a primeira coluna a partir da terceira linha
                        ('BACKGROUND', (0, 2), (0, -1), (230/255, 230/255, 230/255)),  # Fundo cinza claro
                        ('TEXTCOLOR', (0, 2), (0, -1), colors.black),  # Texto preto

                        # Estilo para os títulos das linhas restantes
                        ('TEXTCOLOR', (0, 3), (0, -1), colors.black),  # Texto preto
                        ('FONTNAME', (0, 3), (0, -1), 'Helvetica-Bold'),  # Texto em negrito

                        # Fundo das células de conteúdo restante
                        ('BACKGROUND', (1, 3), (-1, -1), colors.white),  # Fundo branco

                        # Deixar a **coluna 2 da terceira linha pra baixo** em negrito
                        ('FONTNAME', (2, 2), (2, -1), 'Helvetica-Bold'),

                        # Divisões da tabela com cinza claro
                        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Grade interna
                        ('BOX', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Borda externa
                    ])
                    else:
                        StyleTabelaTarget = TableStyle([
                        # Estilo geral
                        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 0), (-1, -1), 5),  # Tamanho da fonte geral
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # Alinhamento à esquerda
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Centralização vertical
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno inferior
                        ('TOPPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno superior

                        # Estilo para a linha de cabeçalho
                        ('SPAN', (0, 0), (-1, 0)),  # Mescla as colunas na primeira linha
                        ('BACKGROUND', (0, 0), (-1, 0), (192/255, 0/255, 10/255)),  # Fundo vermelho
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),  # Texto branco
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Alinhamento centralizado
                        ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),  # Linha acima da primeira linha
                        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),  # Linha abaixo da primeira linha
                        ('FONTSIZE', (0, 0), (-1, 0), 6),  # Aumenta o tamanho da fonte na primeira linha
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Deixa a primeira linha em negrito

                        # Estilo para a segunda linha
                        ('BACKGROUND', (0, 1), (-1, 1), (192/255, 0/255, 10/255)),  # Fundo vermelho
                        ('TEXTCOLOR', (0, 1), (-1, 1), colors.white),  # Texto branco
                        ('ALIGN', (0, 1), (-1, 1), 'CENTER'),  # Alinhamento centralizado

                        # Estilo para a terceira linha
                        ('BACKGROUND', (0, 2), (-1, 2), (230/255, 230/255, 230/255)),  # Fundo cinza claro
                        ('TEXTCOLOR', (0, 2), (-1, 2), colors.black),  # Texto preto
                        ('ALIGN', (0, 2), (-1, 2), 'LEFT'),  # Alinhamento à esquerda

                        # Estilo para a primeira coluna a partir da terceira linha
                        ('BACKGROUND', (0, 2), (0, -1), (230/255, 230/255, 230/255)),  # Fundo cinza claro
                        ('TEXTCOLOR', (0, 2), (0, -1), colors.black),  # Texto preto

                        # Estilo para os títulos das linhas restantes
                        ('TEXTCOLOR', (0, 3), (0, -1), colors.black),  # Texto preto
                        ('FONTNAME', (0, 3), (0, -1), 'Helvetica-Bold'),  # Texto em negrito

                        # Fundo das células de conteúdo restante
                        ('BACKGROUND', (1, 3), (-1, -1), colors.white),  # Fundo branco

                            # Deixar a **coluna 2 da terceira linha pra baixo** em negrito
                        ('FONTNAME', (2, 2), (2, -1), 'Helvetica-Bold'),

                        # Divisões da tabela com cinza claro
                        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Grade interna
                        ('BOX', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Borda e

                        ('TEXTCOLOR', (2, -1), (2, -1), (192/255, 0/255, 10/255)),  # Texto vermelho
                        ('FONTNAME', (2, -1), (2, -1), 'Helvetica-Bold'),  # Texto em negrito
                        ])
            elif AnodaApuração == 'Ano 3':
                    if resultado >= 100:
                        StyleTabelaTarget = TableStyle([
                        # Estilo geral
                        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 0), (-1, -1), 5),  # Tamanho da fonte geral
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # Alinhamento à esquerda
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Centralização vertical
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno inferior
                        ('TOPPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno superior

                        # Estilo para a linha de cabeçalho
                        ('SPAN', (0, 0), (-1, 0)),  # Mescla as colunas na primeira linha
                        ('BACKGROUND', (0, 0), (-1, 0), (192/255, 0/255, 10/255)),  # Fundo vermelho
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),  # Texto branco
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Alinhamento centralizado
                        ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),  # Linha acima da primeira linha
                        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),  # Linha abaixo da primeira linha
                        ('FONTSIZE', (0, 0), (-1, 0), 6),  # Aumenta o tamanho da fonte na primeira linha
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Deixa a primeira linha em negrito

                        # Estilo para a segunda linha
                        ('BACKGROUND', (0, 1), (-1, 1), (192/255, 0/255, 10/255)),  # Fundo vermelho
                        ('TEXTCOLOR', (0, 1), (-1, 1), colors.white),  # Texto branco
                        ('ALIGN', (0, 1), (-1, 1), 'CENTER'),  # Alinhamento centralizado

                        # Estilo para a terceira linha
                        ('BACKGROUND', (0, 2), (-1, 2), (230/255, 230/255, 230/255)),  # Fundo cinza claro
                        ('TEXTCOLOR', (0, 2), (-1, 2), colors.black),  # Texto preto
                        ('ALIGN', (0, 2), (-1, 2), 'LEFT'),  # Alinhamento à esquerda

                        # Estilo para a primeira coluna a partir da terceira linha
                        ('BACKGROUND', (0, 2), (0, -1), (230/255, 230/255, 230/255)),  # Fundo cinza claro
                        ('TEXTCOLOR', (0, 2), (0, -1), colors.black),  # Texto preto

                        # Estilo para os títulos das linhas restantes
                        ('TEXTCOLOR', (0, 3), (0, -1), colors.black),  # Texto preto
                        ('FONTNAME', (0, 3), (0, -1), 'Helvetica-Bold'),  # Texto em negrito

                        # Fundo das células de conteúdo restante
                        ('BACKGROUND', (1, 3), (-1, -1), colors.white),  # Fundo branco

                        # Deixar a **coluna 2 da terceira linha pra baixo** em negrito
                        ('FONTNAME', (3, 2), (3, -1), 'Helvetica-Bold'),

                        # Divisões da tabela com cinza claro
                        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Grade interna
                        ('BOX', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Borda externa
                    ])
                    else:
                        StyleTabelaTarget = TableStyle([
                        # Estilo geral
                        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 0), (-1, -1), 5),  # Tamanho da fonte geral
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # Alinhamento à esquerda
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Centralização vertical
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno inferior
                        ('TOPPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno superior

                        # Estilo para a linha de cabeçalho
                        ('SPAN', (0, 0), (-1, 0)),  # Mescla as colunas na primeira linha
                        ('BACKGROUND', (0, 0), (-1, 0), (192/255, 0/255, 10/255)),  # Fundo vermelho
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),  # Texto branco
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Alinhamento centralizado
                        ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),  # Linha acima da primeira linha
                        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),  # Linha abaixo da primeira linha
                        ('FONTSIZE', (0, 0), (-1, 0), 6),  # Aumenta o tamanho da fonte na primeira linha
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Deixa a primeira linha em negrito

                        # Estilo para a segunda linha
                        ('BACKGROUND', (0, 1), (-1, 1), (192/255, 0/255, 10/255)),  # Fundo vermelho
                        ('TEXTCOLOR', (0, 1), (-1, 1), colors.white),  # Texto branco
                        ('ALIGN', (0, 1), (-1, 1), 'CENTER'),  # Alinhamento centralizado

                        # Estilo para a terceira linha
                        ('BACKGROUND', (0, 2), (-1, 2), (230/255, 230/255, 230/255)),  # Fundo cinza claro
                        ('TEXTCOLOR', (0, 2), (-1, 2), colors.black),  # Texto preto
                        ('ALIGN', (0, 2), (-1, 2), 'LEFT'),  # Alinhamento à esquerda

                        # Estilo para a primeira coluna a partir da terceira linha
                        ('BACKGROUND', (0, 2), (0, -1), (230/255, 230/255, 230/255)),  # Fundo cinza claro
                        ('TEXTCOLOR', (0, 2), (0, -1), colors.black),  # Texto preto

                        # Estilo para os títulos das linhas restantes
                        ('TEXTCOLOR', (0, 3), (0, -1), colors.black),  # Texto preto
                        ('FONTNAME', (0, 3), (0, -1), 'Helvetica-Bold'),  # Texto em negrito

                        # Fundo das células de conteúdo restante
                        ('BACKGROUND', (1, 3), (-1, -1), colors.white),  # Fundo branco

                            # Deixar a **coluna 2 da terceira linha pra baixo** em negrito
                        ('FONTNAME', (3, 2), (3, -1), 'Helvetica-Bold'),

                        # Divisões da tabela com cinza claro
                        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Grade interna
                        ('BOX', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Borda e

                        ('TEXTCOLOR', (3, -1), (3, -1), (192/255, 0/255, 10/255)),  # Texto vermelho
                        ('FONTNAME', (3, -1), (3, -1), 'Helvetica-Bold'),  # Texto em negrito
                        ])
            elif AnodaApuração == 'Ano 4':
                    if resultado >= 100:
                        StyleTabelaTarget = TableStyle([
                        # Estilo geral
                        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 0), (-1, -1), 5),  # Tamanho da fonte geral
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # Alinhamento à esquerda
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Centralização vertical
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno inferior
                        ('TOPPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno superior

                        # Estilo para a linha de cabeçalho
                        ('SPAN', (0, 0), (-1, 0)),  # Mescla as colunas na primeira linha
                        ('BACKGROUND', (0, 0), (-1, 0), (192/255, 0/255, 10/255)),  # Fundo vermelho
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),  # Texto branco
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Alinhamento centralizado
                        ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),  # Linha acima da primeira linha
                        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),  # Linha abaixo da primeira linha
                        ('FONTSIZE', (0, 0), (-1, 0), 6),  # Aumenta o tamanho da fonte na primeira linha
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Deixa a primeira linha em negrito

                        # Estilo para a segunda linha
                        ('BACKGROUND', (0, 1), (-1, 1), (192/255, 0/255, 10/255)),  # Fundo vermelho
                        ('TEXTCOLOR', (0, 1), (-1, 1), colors.white),  # Texto branco
                        ('ALIGN', (0, 1), (-1, 1), 'CENTER'),  # Alinhamento centralizado

                        # Estilo para a terceira linha
                        ('BACKGROUND', (0, 2), (-1, 2), (230/255, 230/255, 230/255)),  # Fundo cinza claro
                        ('TEXTCOLOR', (0, 2), (-1, 2), colors.black),  # Texto preto
                        ('ALIGN', (0, 2), (-1, 2), 'LEFT'),  # Alinhamento à esquerda

                        # Estilo para a primeira coluna a partir da terceira linha
                        ('BACKGROUND', (0, 2), (0, -1), (230/255, 230/255, 230/255)),  # Fundo cinza claro
                        ('TEXTCOLOR', (0, 2), (0, -1), colors.black),  # Texto preto

                        # Estilo para os títulos das linhas restantes
                        ('TEXTCOLOR', (0, 3), (0, -1), colors.black),  # Texto preto
                        ('FONTNAME', (0, 3), (0, -1), 'Helvetica-Bold'),  # Texto em negrito

                        # Fundo das células de conteúdo restante
                        ('BACKGROUND', (1, 3), (-1, -1), colors.white),  # Fundo branco

                        # Deixar a **coluna 2 da terceira linha pra baixo** em negrito
                        ('FONTNAME', (4, 2), (4, -1), 'Helvetica-Bold'),

                        # Divisões da tabela com cinza claro
                        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Grade interna
                        ('BOX', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Borda externa
                    ])
                    else:
                        StyleTabelaTarget = TableStyle([
                        # Estilo geral
                        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 0), (-1, -1), 5),  # Tamanho da fonte geral
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # Alinhamento à esquerda
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Centralização vertical
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno inferior
                        ('TOPPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno superior

                        # Estilo para a linha de cabeçalho
                        ('SPAN', (0, 0), (-1, 0)),  # Mescla as colunas na primeira linha
                        ('BACKGROUND', (0, 0), (-1, 0), (192/255, 0/255, 10/255)),  # Fundo vermelho
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),  # Texto branco
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Alinhamento centralizado
                        ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),  # Linha acima da primeira linha
                        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),  # Linha abaixo da primeira linha
                        ('FONTSIZE', (0, 0), (-1, 0), 6),  # Aumenta o tamanho da fonte na primeira linha
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Deixa a primeira linha em negrito

                        # Estilo para a segunda linha
                        ('BACKGROUND', (0, 1), (-1, 1), (192/255, 0/255, 10/255)),  # Fundo vermelho
                        ('TEXTCOLOR', (0, 1), (-1, 1), colors.white),  # Texto branco
                        ('ALIGN', (0, 1), (-1, 1), 'CENTER'),  # Alinhamento centralizado

                        # Estilo para a terceira linha
                        ('BACKGROUND', (0, 2), (-1, 2), (230/255, 230/255, 230/255)),  # Fundo cinza claro
                        ('TEXTCOLOR', (0, 2), (-1, 2), colors.black),  # Texto preto
                        ('ALIGN', (0, 2), (-1, 2), 'LEFT'),  # Alinhamento à esquerda

                        # Estilo para a primeira coluna a partir da terceira linha
                        ('BACKGROUND', (0, 2), (0, -1), (230/255, 230/255, 230/255)),  # Fundo cinza claro
                        ('TEXTCOLOR', (0, 2), (0, -1), colors.black),  # Texto preto

                        # Estilo para os títulos das linhas restantes
                        ('TEXTCOLOR', (0, 3), (0, -1), colors.black),  # Texto preto
                        ('FONTNAME', (0, 3), (0, -1), 'Helvetica-Bold'),  # Texto em negrito

                        # Fundo das células de conteúdo restante
                        ('BACKGROUND', (1, 3), (-1, -1), colors.white),  # Fundo branco

                            # Deixar a **coluna 2 da terceira linha pra baixo** em negrito
                        ('FONTNAME', (4, 2), (5, -1), 'Helvetica-Bold'),

                        # Divisões da tabela com cinza claro
                        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Grade interna
                        ('BOX', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Borda e

                        ('TEXTCOLOR', (4, -1), (4, -1), (192/255, 0/255, 10/255)),  # Texto vermelho
                        ('FONTNAME', (4, -1), (4, -1), 'Helvetica-Bold'),  # Texto em negrito
                        ])
            elif AnodaApuração == 'Ano 5':
                    if resultado >= 100:
                        StyleTabelaTarget = TableStyle([
                        # Estilo geral
                        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 0), (-1, -1), 5),  # Tamanho da fonte geral
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # Alinhamento à esquerda
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Centralização vertical
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno inferior
                        ('TOPPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno superior

                        # Estilo para a linha de cabeçalho
                        ('SPAN', (0, 0), (-1, 0)),  # Mescla as colunas na primeira linha
                        ('BACKGROUND', (0, 0), (-1, 0), (192/255, 0/255, 10/255)),  # Fundo vermelho
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),  # Texto branco
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Alinhamento centralizado
                        ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),  # Linha acima da primeira linha
                        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),  # Linha abaixo da primeira linha
                        ('FONTSIZE', (0, 0), (-1, 0), 6),  # Aumenta o tamanho da fonte na primeira linha
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Deixa a primeira linha em negrito

                        # Estilo para a segunda linha
                        ('BACKGROUND', (0, 1), (-1, 1), (192/255, 0/255, 10/255)),  # Fundo vermelho
                        ('TEXTCOLOR', (0, 1), (-1, 1), colors.white),  # Texto branco
                        ('ALIGN', (0, 1), (-1, 1), 'CENTER'),  # Alinhamento centralizado

                        # Estilo para a terceira linha
                        ('BACKGROUND', (0, 2), (-1, 2), (230/255, 230/255, 230/255)),  # Fundo cinza claro
                        ('TEXTCOLOR', (0, 2), (-1, 2), colors.black),  # Texto preto
                        ('ALIGN', (0, 2), (-1, 2), 'LEFT'),  # Alinhamento à esquerda

                        # Estilo para a primeira coluna a partir da terceira linha
                        ('BACKGROUND', (0, 2), (0, -1), (230/255, 230/255, 230/255)),  # Fundo cinza claro
                        ('TEXTCOLOR', (0, 2), (0, -1), colors.black),  # Texto preto

                        # Estilo para os títulos das linhas restantes
                        ('TEXTCOLOR', (0, 3), (0, -1), colors.black),  # Texto preto
                        ('FONTNAME', (0, 3), (0, -1), 'Helvetica-Bold'),  # Texto em negrito

                        # Fundo das células de conteúdo restante
                        ('BACKGROUND', (1, 3), (-1, -1), colors.white),  # Fundo branco

                        # Deixar a **coluna 2 da terceira linha pra baixo** em negrito
                        ('FONTNAME', (5, 2), (5, -1), 'Helvetica-Bold'),

                        # Divisões da tabela com cinza claro
                        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Grade interna
                        ('BOX', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Borda externa
                    ])
                    else:
                        StyleTabelaTarget = TableStyle([
                        # Estilo geral
                        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 0), (-1, -1), 5),  # Tamanho da fonte geral
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # Alinhamento à esquerda
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Centralização vertical
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno inferior
                        ('TOPPADDING', (0, 0), (-1, -1), 1),  # Espaçamento interno superior

                        # Estilo para a linha de cabeçalho
                        ('SPAN', (0, 0), (-1, 0)),  # Mescla as colunas na primeira linha
                        ('BACKGROUND', (0, 0), (-1, 0), (192/255, 0/255, 10/255)),  # Fundo vermelho
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),  # Texto branco
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Alinhamento centralizado
                        ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),  # Linha acima da primeira linha
                        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),  # Linha abaixo da primeira linha
                        ('FONTSIZE', (0, 0), (-1, 0), 6),  # Aumenta o tamanho da fonte na primeira linha
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Deixa a primeira linha em negrito

                        # Estilo para a segunda linha
                        ('BACKGROUND', (0, 1), (-1, 1), (192/255, 0/255, 10/255)),  # Fundo vermelho
                        ('TEXTCOLOR', (0, 1), (-1, 1), colors.white),  # Texto branco
                        ('ALIGN', (0, 1), (-1, 1), 'CENTER'),  # Alinhamento centralizado

                        # Estilo para a terceira linha
                        ('BACKGROUND', (0, 2), (-1, 2), (230/255, 230/255, 230/255)),  # Fundo cinza claro
                        ('TEXTCOLOR', (0, 2), (-1, 2), colors.black),  # Texto preto
                        ('ALIGN', (0, 2), (-1, 2), 'LEFT'),  # Alinhamento à esquerda

                        # Estilo para a primeira coluna a partir da terceira linha
                        ('BACKGROUND', (0, 2), (0, -1), (230/255, 230/255, 230/255)),  # Fundo cinza claro
                        ('TEXTCOLOR', (0, 2), (0, -1), colors.black),  # Texto preto

                        # Estilo para os títulos das linhas restantes
                        ('TEXTCOLOR', (0, 3), (0, -1), colors.black),  # Texto preto
                        ('FONTNAME', (0, 3), (0, -1), 'Helvetica-Bold'),  # Texto em negrito

                        # Fundo das células de conteúdo restante
                        ('BACKGROUND', (1, 3), (-1, -1), colors.white),  # Fundo branco

                            # Deixar a **coluna 2 da terceira linha pra baixo** em negrito
                        ('FONTNAME', (5, 2), (5, -1), 'Helvetica-Bold'),

                        # Divisões da tabela com cinza claro
                        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Grade interna
                        ('BOX', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Borda e

                        ('TEXTCOLOR', (5, -1), (5, -1), (192/255, 0/255, 10/255)),  # Texto vermelho
                        ('FONTNAME', (5, -1), (5, -1), 'Helvetica-Bold'),  # Texto em negrito
                        ])
            
            
            
            
            if AnodaApuração == "Ano 1":
                if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_1_total )

                elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_1 )
                elif not BaseCurtoFiltradoCliente.empty and BaseLongoFiltradoCliente.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_1_curto )

                if not ClientesManutenção.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_1_Manutenção )     
                
                if not ClientesNovoComodato.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_1_NovoComodato )   

                if not ClientesacordodeConsumo.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_1_AcordoDeConsumo)   


            elif AnodaApuração == "Ano 2":

                if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_2_total )

                elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_2 )
                elif not BaseCurtoFiltradoCliente.empty and BaseLongoFiltradoCliente.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_2_curto )

                if not ClientesNovoComodato.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_2_NovoComodato )  
                
                if not ClientesManutenção.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_2_Manutenção )     
                if not ClientesacordodeConsumo.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_3_AcordoDeConsumo)   

                

            elif AnodaApuração == "Ano 3":

                if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_3_total )

                elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_3 )
                elif not BaseCurtoFiltradoCliente.empty and BaseLongoFiltradoCliente.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_3_curto )
                
                if not ClientesManutenção.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_3_Manutenção )  

                if not ClientesNovoComodato.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_3_NovoComodato )  
                if not ClientesacordodeConsumo.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_3_AcordoDeConsumo)   


            elif AnodaApuração == "Ano 4":

                if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_4_total )

                elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_4 )
                elif not BaseCurtoFiltradoCliente.empty and BaseLongoFiltradoCliente.empty:
                    
                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_4_curto )

                if not ClientesManutenção.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_4_Manutenção )  

                if not ClientesNovoComodato.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_4_NovoComodato )  
                if not ClientesacordodeConsumo.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_4_AcordoDeConsumo)   

            elif AnodaApuração == "Ano 5":

                if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_5_total )

                elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_5 )
                elif not BaseCurtoFiltradoCliente.empty and BaseLongoFiltradoCliente.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_5_curto )
                
                if not ClientesManutenção.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_5_Manutenção )  

                if not ClientesNovoComodato.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_5_NovoComodato )  
                if not ClientesacordodeConsumo.empty:

                    diferencaCurtoeLongo = validar_formatar_consumo(soma_valor_total - consumo_ano_5_AcordoDeConsumo)   


            def calcular_penalidade(valor, porcentagem):
                if porcentagem < 100:  # Verifica se a porcentagem é válida
                    if 91 < porcentagem <= 99:  # Maior que 91 e menor ou igual a 99
                        penalidade = valor * 0.06
                        valor = penalidade
                    elif 81 < porcentagem <= 90:  # Maior que 81 e menor ou igual a 90
                        penalidade = valor * 0.14
                        valor = penalidade
                    elif 71 < porcentagem <= 80:  # Maior que 71 e menor ou igual a 80
                        penalidade = valor * 0.21
                        valor = penalidade
                    elif 61 < porcentagem <= 70:  # Maior que 61 e menor ou igual a 70
                        penalidade = valor * 0.28
                        valor = penalidade
                    elif 51 < porcentagem <= 60:  # Maior que 51 e menor ou igual a 60
                        penalidade = valor * 0.34
                        valor = penalidade
                    elif 41 < porcentagem <= 50:  # Maior que 41 e menor ou igual a 50
                        penalidade = valor * 0.40
                        valor = penalidade
                    elif 31 < porcentagem <= 40:  # Maior que 31 e menor ou igual a 40
                        penalidade = valor * 0.45
                        valor = penalidade
                    elif 21 < porcentagem <= 30:  # Maior que 21 e menor ou igual a 30
                        penalidade = valor * 0.48
                        valor = penalidade
                    elif 0 <= porcentagem <= 20:  # Entre 0 e 20, incluindo ambos
                        penalidade = valor * 0.50
                        valor = penalidade
                else:
                    return 0 
                    
                return valor

            # Exemplo de uso
            # Exemplo de porcentagem negativa
            if AnodaApuração == 'Ano 1':
                if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    valor_original = consumo_ano_1_total
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)

                elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:
                    valor_original = consumo_ano_1
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)

                elif not BaseCurtoFiltradoCliente.empty and BaseLongoFiltradoCliente.empty:
                    valor_original = consumo_ano_1_curto
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)
                if not ClientesManutenção.empty:
                    valor_original = consumo_ano_1_Manutenção
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)
                
                if not ClientesNovoComodato.empty:
                    valor_original = consumo_ano_1_NovoComodato
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)

                if not ClientesacordodeConsumo.empty:
                    valor_original = consumo_ano_1_AcordoDeConsumo
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)

            if AnodaApuração == 'Ano 2':
                if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    valor_original = consumo_ano_2_total
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)

                elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:
                    valor_original = consumo_ano_2
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)

                elif not BaseCurtoFiltradoCliente.empty and BaseLongoFiltradoCliente.empty:
                    valor_original = consumo_ano_2_curto
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)

                if not ClientesManutenção.empty:
                    valor_original = consumo_ano_2_Manutenção
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)

                if not ClientesNovoComodato.empty:
                    valor_original = consumo_ano_2_NovoComodato
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)
                
                if not ClientesacordodeConsumo.empty:
                    valor_original = consumo_ano_2_AcordoDeConsumo
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)
            if AnodaApuração == 'Ano 3':
                if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    valor_original = consumo_ano_3_total
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)

                elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:
                    valor_original = consumo_ano_3
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)

                elif not BaseCurtoFiltradoCliente.empty and BaseLongoFiltradoCliente.empty:
                    valor_original = consumo_ano_3_curto
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)
                    
                if not ClientesManutenção.empty:
                    valor_original = consumo_ano_3_Manutenção
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)
                
                if not ClientesNovoComodato.empty:
                    valor_original = consumo_ano_3_NovoComodato
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)

                if not ClientesacordodeConsumo.empty:
                    valor_original = consumo_ano_3_AcordoDeConsumo
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)

            if AnodaApuração == 'Ano 4':
                if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    valor_original = consumo_ano_4_total
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)

                elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:
                    valor_original = consumo_ano_4
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)

                elif not BaseCurtoFiltradoCliente.empty and BaseLongoFiltradoCliente.empty:
                    valor_original = consumo_ano_4_curto
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)
                if not ClientesManutenção.empty:
                    valor_original = consumo_ano_4_Manutenção
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)
                if not ClientesNovoComodato.empty:
                    valor_original = consumo_ano_4_NovoComodato
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)
                if not ClientesacordodeConsumo.empty:
                    valor_original = consumo_ano_4_AcordoDeConsumo
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)
                    
            if AnodaApuração == 'Ano 5':
                if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    valor_original = consumo_ano_5_total
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)

                elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:
                    valor_original = consumo_ano_5
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)

                elif not BaseCurtoFiltradoCliente.empty and BaseLongoFiltradoCliente.empty:
                    valor_original = consumo_ano_5_curto
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)
                
                if not ClientesManutenção.empty:
                    valor_original = consumo_ano_5_Manutenção
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)
                if not ClientesNovoComodato.empty:
                    valor_original = consumo_ano_5_NovoComodato
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)
                if not ClientesacordodeConsumo.empty:
                    valor_original = consumo_ano_5_AcordoDeConsumo
                    porcentagem_atual = resultado
                    valor_final = calcular_penalidade(valor_original, porcentagem_atual)

            MultaFormatada = validar_formatar_consumo(valor_final)
            # Adicionar linha de BaseLongoFiltradoCliente, se houver dados


            if AnodaApuração == "Ano 1":

                if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    Target = [['Consumo Unificado'],
                    ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                    ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                    [f'Target - ' + BaseCurtoFiltradoCliente.iloc[0]['Nº INTERNO'] + '(Compra e Venda)', ValorConsumoAno1Curto, ValorConsumoAno2Curto,ValorConsumoAno3Curto,ValorConsumoAno4Curto,ValorConsumoAno5Curto],  # Dados
                    [f'Target - ' + BaseLongoFiltradoCliente.iloc[0]['Nº INTERNO'] + '(Locação)', ValorConsumoAno1Longo,ValorConsumoAno2Longo,ValorConsumoAno3Longo , ValorConsumoAno5Longo,ValorConsumoAno5Longo],
                    ['Target Unificado', ValorConsumoTotalAno1, '','','','' ],
                    ['Valor Consumido - Unificado', valor_total_formatado, '', '', '', ''],
                    ['Percentual de Atingimento', f"{resultado}%", '', '', '', '']   # Dados
                    ]
                    

                elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:  # Verifica se apenas BaseLongoFiltradoCliente não está vazio
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + BaseLongoFiltradoCliente.iloc[0]['Nº INTERNO'] + '(Locação)', ValorConsumoAno1Longo,ValorConsumoAno2Longo,ValorConsumoAno3Longo , ValorConsumoAno5Longo,ValorConsumoAno5Longo],
                        ['Target Unificado', ValorConsumoAno1Longo, '', '','', ''],
                        ['Valor Consumido - Unificado', valor_total_formatado, '', '', '', ''], 
                        ['Percentual de Atingimento', f"{resultado}%", '', '', '', '']   # Dados
                    ]
                    

                elif not BaseCurtoFiltradoCliente.empty and BaseLongoFiltradoCliente.empty: # Caso nenhuma das condições anteriores seja atendida
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + BaseCurtoFiltradoCliente.iloc[0]['Nº INTERNO'] + '(Compra e Venda)', ValorConsumoAno1Curto, ValorConsumoAno2Curto,ValorConsumoAno3Curto,ValorConsumoAno4Curto,ValorConsumoAno5Curto],
                        ['Target Unificado', ValorConsumoAno1Curto, '', '','', ''],
                        ['Valor Consumido - Unificado', valor_total_formatado, '', '', '', ''] ,
                        ['Percentual de Atingimento', f"{resultado}%", '', '', '', '']   # Dados
                        # Dados
                    ]

                elif not ClientesManutenção.empty: # Caso nenhuma das condições anteriores seja atendida
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + ClientesManutenção.iloc[0]['Nº INTERNO'] + '(Manutenção)', ValorConsumoAno1Manutenção, ValorConsumoAno2Manutenção,ValorConsumoAno3Manutenção,ValorConsumoAno4Manutenção,ValorConsumoAno5Manutenção],
                        ['Target Unificado', ValorConsumoAno1Manutenção, '', '','', ''],
                        ['Valor Consumido - Unificado', valor_total_formatado, '', '', '', ''] ,
                        ['Percentual de Atingimento', f"{resultado}%", '', '', '', '']   # Dados
                        # Dados
                    ]
                elif not ClientesNovoComodato.empty: # Caso nenhuma das condições anteriores seja atendida
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + ClientesNovoComodato.iloc[0]['Nº INTERNO'] + '(Novo Comodato)', ValorConsumoAno1NovoComodato, ValorConsumoAno2NovoComodato,ValorConsumoAno3NovoComodato,ValorConsumoAno4NovoComodato,ValorConsumoAno5NovoComodato],
                        ['Target Unificado', ValorConsumoAno1NovoComodato, '', '','', ''],
                        ['Valor Consumido - Unificado', valor_total_formatado, '', '', '', ''] ,
                        ['Percentual de Atingimento', f"{resultado}%", '', '', '', '']   # Dados
                        # Dados
                    ]
                elif not ClientesacordodeConsumo.empty: # Caso nenhuma das condições anteriores seja atendida
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + ClientesacordodeConsumo.iloc[0]['Nº INTERNO'] + '(Novo Comodato)', ValorConsumoAno1AcordoDeConsumo, ValorConsumoAno2AcordoDeConsumo,ValorConsumoAno3AcordoDeConsumo,ValorConsumoAno4AcordoDeConsumo,ValorConsumoAno5AcordoDeConsumo],
                        ['Target Unificado', ValorConsumoAno1AcordoDeConsumo, '', '','', ''],
                        ['Valor Consumido - Unificado', valor_total_formatado, '', '', '', ''] ,
                        ['Percentual de Atingimento', f"{resultado}%", '', '', '', '']   # Dados
                        # Dados
                    ]
                    
            elif  AnodaApuração == "Ano 2":


                if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    Target = [['Consumo Unificado'],
                    ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                    ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                    [f'Target - ' + BaseCurtoFiltradoCliente.iloc[0]['Nº INTERNO'] + '(Compra e Venda)', ValorConsumoAno1Curto, ValorConsumoAno2Curto, ValorConsumoAno3Curto, ValorConsumoAno4Curto,ValorConsumoAno5Curto],  # Dados
                    [f'Target - ' + BaseLongoFiltradoCliente.iloc[0]['Nº INTERNO'] + '(Locação)', ValorConsumoAno1Longo, ValorConsumoAno2Longo, ValorConsumoAno3Longo, ValorConsumoAno4Longo,ValorConsumoAno5Longo],
                    ['Target Unificado', ValorConsumoTotalAno1, ValorConsumoTotalAno2, ValorConsumoTotalAno3, ValorConsumoTotalAno4, ValorConsumoTotalAno5],
                    ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatado, '', '', ''],
                    ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultado}%", '', '', '']   # Dados
                    ]
                    
                elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:  # Verifica se apenas BaseLongoFiltradoCliente não está vazio
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + BaseLongoFiltradoCliente.iloc[0]['Nº INTERNO'] + '(Locação)', ValorConsumoAno1Longo, ValorConsumoAno2Longo, ValorConsumoAno3Longo, ValorConsumoAno4Longo, ValorConsumoAno5Longo],
                        ['Target Unificado', ValorConsumoAno1Longo, ValorConsumoAno2Longo, ValorConsumoAno3Longo,ValorConsumoAno4Longo, ValorConsumoAno5Longo],
                        ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatado, '', '', ''],
                        ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultado}%", '', '', '']   # Dados
                    ]
                    
                elif not BaseCurtoFiltradoCliente.empty and BaseLongoFiltradoCliente.empty: # Caso nenhuma das condições anteriores seja atendida
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + BaseCurtoFiltradoCliente.iloc[0]['Nº INTERNO'] + '(Compra e Venda)', ValorConsumoAno1Curto,ValorConsumoAno2Curto,ValorConsumoAno3Curto,ValorConsumoAno4Curto, ValorConsumoAno5Curto],
                        ['Target Unificado', ValorConsumoAno1Curto, ValorConsumoAno2Curto, ValorConsumoAno3Curto,ValorConsumoAno4Curto, ValorConsumoAno5Curto],
                        ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatado, '', '', ''],
                        ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultado}%", '', '', '']   # Dados
                        # Dados
                    ]
                elif not ClientesManutenção.empty: # Caso nenhuma das condições anteriores seja atendida
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + ClientesManutenção.iloc[0]['Nº INTERNO'] + '(Manutenção)', ValorConsumoAno1Manutenção, ValorConsumoAno2Manutenção,ValorConsumoAno3Manutenção,ValorConsumoAno4Manutenção,ValorConsumoAno5Manutenção],
                        ['Target Unificado', ValorConsumoAno1Manutenção, ValorConsumoAno2Manutenção, '','', ''],
                        ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatado, '', '', ''] ,
                        ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultado}%", '', '', '']   # Dados
                        # Dados
                    ]
                elif not ClientesNovoComodato.empty: # Caso nenhuma das condições anteriores seja atendida
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + ClientesNovoComodato.iloc[0]['Nº INTERNO'] + '(Novo Comodato)', ValorConsumoAno1NovoComodato, ValorConsumoAno2NovoComodato,ValorConsumoAno3NovoComodato,ValorConsumoAno4NovoComodato,ValorConsumoAno5NovoComodato],
                        ['Target Unificado', ValorConsumoAno1NovoComodato, ValorConsumoAno2NovoComodato, '','', ''],
                        ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatado, '', '', ''] ,
                        ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultado}%", '', '', '']   # Dados
                        # Dados
                    ]  
                elif not ClientesacordodeConsumo.empty: # Caso nenhuma das condições anteriores seja atendida
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + ClientesacordodeConsumo.iloc[0]['Nº INTERNO'] + '(Novo Comodato)', ValorConsumoAno1AcordoDeConsumo, ValorConsumoAno2AcordoDeConsumo,ValorConsumoAno3AcordoDeConsumo,ValorConsumoAno4AcordoDeConsumo,ValorConsumoAno5AcordoDeConsumo],
                        ['Target Unificado', ValorConsumoAno1AcordoDeConsumo, ValorConsumoAno2AcordoDeConsumo, '','', ''],
                        ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatado, '', '', ''] ,
                        ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultado}%", '', '', '']   # Dados
                        # Dados
                    ]  

            elif  AnodaApuração == "Ano 3":


                if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    Target = [['Consumo Unificado'],
                    ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                    ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                    [f'Target - ' + BaseCurtoFiltradoCliente.iloc[0]['Nº INTERNO'] + '(Compra e Venda)', ValorConsumoAno1Curto, ValorConsumoAno2Curto, ValorConsumoAno3Curto, ValorConsumoAno4Curto,ValorConsumoAno5Curto],  # Dados
                    [f'Target - ' + BaseLongoFiltradoCliente.iloc[0]['Nº INTERNO'] + '(Locação)', ValorConsumoAno1Longo, ValorConsumoAno2Longo, ValorConsumoAno3Longo, ValorConsumoAno4Longo,ValorConsumoAno5Longo],
                    ['Target Unificado', ValorConsumoTotalAno1, ValorConsumoTotalAno2, ValorConsumoTotalAno3, ValorConsumoTotalAno4, ValorConsumoTotalAno5],
                    ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatadoAno2, valor_total_formatado, '', ''],
                    ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultadoAno2}%", f"{resultado}%", '', '']   # Dados
                    ]
                    
                elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:  # Verifica se apenas BaseLongoFiltradoCliente não está vazio
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + BaseLongoFiltradoCliente.iloc[0]['Nº INTERNO'] + '(Locação)', ValorConsumoAno1Longo, ValorConsumoAno2Longo, ValorConsumoAno3Longo, ValorConsumoAno4Longo, ValorConsumoAno5Longo],
                        ['Target Unificado', ValorConsumoAno1Longo, ValorConsumoAno2Longo, ValorConsumoAno3Longo,ValorConsumoAno4Longo, ValorConsumoAno5Longo],
                        ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatadoAno2, valor_total_formatado, '', ''],
                        ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultadoAno2}%", f"{resultado}%", '', '']   # Dados
                    ]
                    


                elif not BaseCurtoFiltradoCliente.empty and BaseLongoFiltradoCliente.empty: # Caso nenhuma das condições anteriores seja atendida
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + BaseCurtoFiltradoCliente.iloc[0]['Nº INTERNO'] + '(Compra e Venda)', ValorConsumoAno1Curto,ValorConsumoAno2Curto,ValorConsumoAno3Curto,ValorConsumoAno4Curto, ValorConsumoAno5Curto],
                        ['Target Unificado', ValorConsumoAno1Curto, ValorConsumoAno2Curto, ValorConsumoAno3Curto,ValorConsumoAno4Curto, ValorConsumoAno5Curto],
                        ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatadoAno2, valor_total_formatado, '', ''],
                        ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultadoAno2}%", f"{resultado}%", '', '']   # Dados
                        # Dados
                    ]
                elif not ClientesManutenção.empty: # Caso nenhuma das condições anteriores seja atendida
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + ClientesManutenção.iloc[0]['Nº INTERNO'] + '(Manutenção)', ValorConsumoAno1Manutenção, ValorConsumoAno2Manutenção,ValorConsumoAno3Manutenção,ValorConsumoAno4Manutenção,ValorConsumoAno5Manutenção],
                        ['Target Unificado', ValorConsumoAno1Manutenção, ValorConsumoAno2Manutenção, ValorConsumoAno3Manutenção,'', ''],
                        ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatadoAno2, valor_total_formatado, '', ''] ,
                        ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultadoAno2}%", f"{resultado}%", '', '']   # Dados
                        # Dados
                    ]
                elif not ClientesNovoComodato.empty: # Caso nenhuma das condições anteriores seja atendida
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + ClientesNovoComodato.iloc[0]['Nº INTERNO'] + '(Novo Comodato)', ValorConsumoAno1NovoComodato, ValorConsumoAno2NovoComodato,ValorConsumoAno3NovoComodato,ValorConsumoAno4NovoComodato,ValorConsumoAno5NovoComodato],
                        ['Target Unificado', ValorConsumoAno1NovoComodato, ValorConsumoAno2NovoComodato, ValorConsumoAno3NovoComodato,'', ''],
                        ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatadoAno2, valor_total_formatado, '', ''] ,
                        ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultadoAno2}%", f"{resultado}%", '', '']   # Dados
                        # Dados
                    ]
                elif not ClientesacordodeConsumo.empty: # Caso nenhuma das condições anteriores seja atendida
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + ClientesacordodeConsumo.iloc[0]['Nº INTERNO'] + '(Novo Comodato)', ValorConsumoAno1AcordoDeConsumo, ValorConsumoAno2AcordoDeConsumo,ValorConsumoAno3AcordoDeConsumo,ValorConsumoAno4AcordoDeConsumo,ValorConsumoAno5AcordoDeConsumo],
                        ['Target Unificado', ValorConsumoAno1AcordoDeConsumo, ValorConsumoAno2AcordoDeConsumo, ValorConsumoAno3AcordoDeConsumo,'', ''],
                        ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatadoAno2, valor_total_formatado, '', ''] ,
                        ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultadoAno2}%", f"{resultado}%", '', '']   # Dados
                        # Dados
                    ]
                    
            elif  AnodaApuração == "Ano 4":


                if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    Target = [['Consumo Unificado'],
                    ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                    ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                    [f'Target - ' + BaseCurtoFiltradoCliente.iloc[0]['Nº INTERNO'] + '(Compra e Venda)', ValorConsumoAno1Curto, ValorConsumoAno2Curto, ValorConsumoAno3Curto, ValorConsumoAno4Curto,ValorConsumoAno5Curto],  # Dados
                    [f'Target - ' + BaseLongoFiltradoCliente.iloc[0]['Nº INTERNO'] + '(Locação)', ValorConsumoAno1Longo, ValorConsumoAno2Longo, ValorConsumoAno3Longo, ValorConsumoAno4Longo,ValorConsumoAno5Longo],
                    ['Target Unificado', ValorConsumoTotalAno1, ValorConsumoTotalAno2, ValorConsumoTotalAno3, ValorConsumoTotalAno4, ValorConsumoTotalAno5],
                    ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatadoAno2, valor_total_formatadoAno3, valor_total_formatado, ''],
                    ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultadoAno2}%", f"{resultadoAno3}%", f"{resultado}%", '']   # Dados
                    ]
                    
                elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:  # Verifica se apenas BaseLongoFiltradoCliente não está vazio
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + BaseLongoFiltradoCliente.iloc[0]['Nº INTERNO'] + '(Locação)', ValorConsumoAno1Longo, ValorConsumoAno2Longo, ValorConsumoAno3Longo, ValorConsumoAno4Longo, ValorConsumoAno5Longo],
                        ['Target Unificado', ValorConsumoAno1Longo, ValorConsumoAno2Longo, ValorConsumoAno3Longo,ValorConsumoAno4Longo, ValorConsumoAno5Longo],
                        ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatadoAno2, valor_total_formatadoAno3, valor_total_formatado, ''],
                        ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultadoAno2}%", f"{resultadoAno3}%", f"{resultado}%", '']   # Dados
                    ]
                    
                elif not BaseCurtoFiltradoCliente.empty and BaseLongoFiltradoCliente.empty: # Caso nenhuma das condições anteriores seja atendida
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + BaseCurtoFiltradoCliente.iloc[0]['Nº INTERNO'] + '(Compra e Venda)', ValorConsumoAno1Curto,ValorConsumoAno2Curto,ValorConsumoAno3Curto,ValorConsumoAno4Curto, ValorConsumoAno5Curto],
                        ['Target Unificado', ValorConsumoAno1Curto, ValorConsumoAno2Curto, ValorConsumoAno3Curto,ValorConsumoAno4Curto, ValorConsumoAno5Curto],
                        ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatadoAno2, valor_total_formatadoAno3, valor_total_formatado, ''],
                        ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultadoAno2}%", f"{resultadoAno3}%", f"{resultado}%", '']   # Dados
                        # Dados
                    ]
                elif not ClientesManutenção.empty: # Caso nenhuma das condições anteriores seja atendida
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + ClientesManutenção.iloc[0]['Nº INTERNO'] + '(Manutenção)', ValorConsumoAno1Manutenção, ValorConsumoAno2Manutenção,ValorConsumoAno3Manutenção,ValorConsumoAno4Manutenção,ValorConsumoAno5Manutenção],
                        ['Target Unificado', ValorConsumoAno1Manutenção, ValorConsumoAno2Manutenção, ValorConsumoAno3Manutenção,ValorConsumoAno4Manutenção, ValorConsumoAno5Manutenção],
                        ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatadoAno2, valor_total_formatadoAno3, valor_total_formatado, ''] ,
                        ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultadoAno2}%",  f"{resultadoAno3}%", f"{resultado}%", '']   # Dados
                        # Dados
                    ]
                elif not ClientesNovoComodato.empty: # Caso nenhuma das condições anteriores seja atendida
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + ClientesNovoComodato.iloc[0]['Nº INTERNO'] + '(Novo Comodato)', ValorConsumoAno1NovoComodato, ValorConsumoAno2NovoComodato,ValorConsumoAno3NovoComodato,ValorConsumoAno4NovoComodato,ValorConsumoAno5NovoComodato],
                        ['Target Unificado', ValorConsumoAno1NovoComodato, ValorConsumoAno2NovoComodato, ValorConsumoAno3NovoComodato, ValorConsumoAno4NovoComodato, ValorConsumoAno5NovoComodato],
                        ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatadoAno2,valor_total_formatadoAno3, valor_total_formatado, ''] ,
                        ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultadoAno2}%", f"{resultadoAno3}%", f"{resultado}%", '']   # Dados
                        # Dados
                    ] 
                elif not ClientesacordodeConsumo.empty: # Caso nenhuma das condições anteriores seja atendida
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + ClientesacordodeConsumo.iloc[0]['Nº INTERNO'] + '(Novo Comodato)', ValorConsumoAno1AcordoDeConsumo, ValorConsumoAno2AcordoDeConsumo,ValorConsumoAno3AcordoDeConsumo,ValorConsumoAno4AcordoDeConsumo,ValorConsumoAno5AcordoDeConsumo],
                        ['Target Unificado', ValorConsumoAno1AcordoDeConsumo, ValorConsumoAno2AcordoDeConsumo, ValorConsumoAno3AcordoDeConsumo, ValorConsumoAno4AcordoDeConsumo,ValorConsumoAno5AcordoDeConsumo],
                        ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatadoAno2,valor_total_formatadoAno3, valor_total_formatado, ''] ,
                        ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultadoAno2}%", f"{resultadoAno3}%", f"{resultado}%", '']   # Dados
                        # Dados
                    ]       
            
            elif  AnodaApuração == "Ano 5":


                if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                    Target = [['Consumo Unificado'],
                    ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                    ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                    [f'Target - ' + BaseCurtoFiltradoCliente.iloc[0]['Nº INTERNO'] + '(Compra e Venda)', ValorConsumoAno1Curto, ValorConsumoAno2Curto, ValorConsumoAno3Curto, ValorConsumoAno4Curto,ValorConsumoAno5Curto],  # Dados
                    [f'Target - ' + BaseLongoFiltradoCliente.iloc[0]['Nº INTERNO'] + '(Locação)', ValorConsumoAno1Longo, ValorConsumoAno2Longo, ValorConsumoAno3Longo, ValorConsumoAno4Longo,ValorConsumoAno5Longo],
                    ['Target Unificado', ValorConsumoTotalAno1, ValorConsumoTotalAno2, ValorConsumoTotalAno3, ValorConsumoTotalAno4, ValorConsumoTotalAno5],
                    ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatadoAno2, valor_total_formatadoAno3,valor_total_formatadoAno4, valor_total_formatado],
                    ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultadoAno2}%", f"{resultadoAno3}%", f"{resultadoAno4}%", f"{resultado}%"]   # Dados
                    ]
                    

                elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:  # Verifica se apenas BaseLongoFiltradoCliente não está vazio
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + BaseLongoFiltradoCliente.iloc[0]['Nº INTERNO'] + '(Locação)', ValorConsumoAno1Longo, ValorConsumoAno2Longo, ValorConsumoAno3Longo, ValorConsumoAno4Longo, ValorConsumoAno5Longo],
                        ['Target Unificado', ValorConsumoAno1Longo, ValorConsumoAno2Longo, ValorConsumoAno3Longo,ValorConsumoAno4Longo, ValorConsumoAno5Longo],
                        ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatadoAno2, valor_total_formatadoAno3,valor_total_formatadoAno4, valor_total_formatado],
                        ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultadoAno2}%", f"{resultadoAno3}%", f"{resultadoAno4}%", f"{resultado}%"]   # Dados
                    
                    ]
                    

                elif not BaseCurtoFiltradoCliente.empty and BaseLongoFiltradoCliente.empty: # Caso nenhuma das condições anteriores seja atendida
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + BaseCurtoFiltradoCliente.iloc[0]['Nº INTERNO'] + '(Compra e Venda)', ValorConsumoAno1Curto,ValorConsumoAno2Curto,ValorConsumoAno3Curto,ValorConsumoAno4Curto, ValorConsumoAno5Curto],
                        ['Target Unificado', ValorConsumoAno1Curto, ValorConsumoAno2Curto, ValorConsumoAno3Curto,ValorConsumoAno4Curto, ValorConsumoAno5Curto],
                        ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatadoAno2, valor_total_formatadoAno3,valor_total_formatadoAno4, valor_total_formatado],
                        ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultadoAno2}%", f"{resultadoAno3}%", f"{resultadoAno4}%", f"{resultado}%"]   # Dados
                    
                        # Dados
                    ]

                elif not ClientesManutenção.empty: # Caso nenhuma das condições anteriores seja atendida
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + ClientesManutenção.iloc[0]['Nº INTERNO'] + '(Manutenção)', ValorConsumoAno1Manutenção, ValorConsumoAno2Manutenção,ValorConsumoAno3Manutenção,ValorConsumoAno4Manutenção,ValorConsumoAno5Manutenção],
                        ['Target Unificado', ValorConsumoAno1Manutenção, ValorConsumoAno2Manutenção, ValorConsumoAno3Manutenção,ValorConsumoAno4Manutenção, ValorConsumoAno5Manutenção],
                        ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatadoAno2, valor_total_formatadoAno3, valor_total_formatadoAno4, valor_total_formatado] ,
                        ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultadoAno2}%",  f"{resultadoAno3}%", f"{resultadoAno4}%", f"{resultado}%"]   # Dados
                        # Dados
                    ]
                elif not ClientesNovoComodato.empty: # Caso nenhuma das condições anteriores seja atendida
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + ClientesNovoComodato.iloc[0]['Nº INTERNO'] + '(Novo Comodato)', ValorConsumoAno1NovoComodato, ValorConsumoAno2NovoComodato,ValorConsumoAno3NovoComodato,ValorConsumoAno4NovoComodato,ValorConsumoAno5NovoComodato],
                        ['Target Unificado', ValorConsumoAno1NovoComodato, ValorConsumoAno2NovoComodato, ValorConsumoAno3NovoComodato, ValorConsumoAno4NovoComodato, ValorConsumoAno5NovoComodato],
                        ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatadoAno2,valor_total_formatadoAno3, valor_total_formatadoAno4, valor_total_formatado] ,
                        ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultadoAno2}%", f"{resultadoAno3}%", f"{resultadoAno4}%", f"{resultado}%"]   # Dados
                        # Dados
                    ] 
                elif not ClientesacordodeConsumo.empty: # Caso nenhuma das condições anteriores seja atendida
                    Target = [['Consumo Unificado'],
                        ['', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5'],  # Cabeçalho superior
                        ['Meta %', '100%', '100%', '100%', '100%', '100%'],  # Linha de metas
                        [f'Target - ' + ClientesacordodeConsumo.iloc[0]['Nº INTERNO'] + '(Novo Comodato)', ValorConsumoAno1AcordoDeConsumo, ValorConsumoAno2AcordoDeConsumo,ValorConsumoAno3AcordoDeConsumo,ValorConsumoAno4AcordoDeConsumo,ValorConsumoAno5AcordoDeConsumo],
                        ['Target Unificado', ValorConsumoAno1AcordoDeConsumo, ValorConsumoAno2AcordoDeConsumo, ValorConsumoAno3AcordoDeConsumo, ValorConsumoAno4AcordoDeConsumo, ValorConsumoAno5AcordoDeConsumo],
                        ['Valor Consumido - Unificado', valor_total_formatadoAno1, valor_total_formatadoAno2,valor_total_formatadoAno3, valor_total_formatadoAno4, valor_total_formatado] ,
                        ['Percentual de Atingimento', f"{resultadoAno1}%", f"{resultadoAno2}%", f"{resultadoAno3}%", f"{resultadoAno4}%", f"{resultado}%"]   # Dados
                        # Dados
                    ]     
                            
            # Nome do arquivo
            # Nome do arquivo

            def limpar_nome_arquivo(nome):
                return re.sub(r'[<>:"/\\|?*]', '_', nome)
            
            # nome_arquivo = limpar_nome_arquivo(f"{row['SAP PRINCIPAL']} - {row['RAZÃO SOCIAL']}.pdf")

            if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                nome_arquivo = limpar_nome_arquivo(f"{row['SAP PRINCIPAL']}!{row['RAZÃO SOCIAL']}!{AnodaApuração}!{valor_total_formatado}!{resultado}!{ValorConsumoTotalAno1}!{diferencaCurtoeLongo}!{MultaFormatada}!{DataDaApuraçãoLongo}!{DataFimApuração}!{BaseLongoFiltradoCliente.iloc[0]['Nº INTERNO']}!{BaseLongoFiltradoCliente.iloc[0]['MODALIDADE']}!{BaseCurtoFiltradoCliente.iloc[0]['Nº INTERNO']}!{BaseCurtoFiltradoCliente.iloc[0]['MODALIDADE']}!{meses_passados}!{total_meses}!{MesSelecionado}!{AnoSelecionado}.pdf")
            elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:
                nome_arquivo = limpar_nome_arquivo(f"{row['SAP PRINCIPAL']}!{row['RAZÃO SOCIAL']}!{AnodaApuração}!{valor_total_formatado}!{resultado}!{ValorConsumoAno1Longo}!{diferencaCurtoeLongo}!{MultaFormatada}!{DataDaApuraçãoLongo}!{DataFimApuração}!{BaseLongoFiltradoCliente.iloc[0]['Nº INTERNO']}!{BaseLongoFiltradoCliente.iloc[0]['MODALIDADE']}!{meses_passados}!{total_meses}!{MesSelecionado}!{AnoSelecionado}.pdf")
            elif not BaseCurtoFiltradoCliente.empty and BaseLongoFiltradoCliente.empty:
                nome_arquivo = limpar_nome_arquivo(f"{row['SAP PRINCIPAL']}!{row['RAZÃO SOCIAL']}!{AnodaApuração}!{valor_total_formatado}!{resultado}!{ValorConsumoAno1Curto}!{diferencaCurtoeLongo}!{MultaFormatada}!{DataDaApuraçãoLongo}!{DataFimApuração}!{BaseCurtoFiltradoCliente.iloc[0]['Nº INTERNO']}!{BaseCurtoFiltradoCliente.iloc[0]['MODALIDADE']}!{meses_passados}!{total_mesesCurto}!{MesSelecionado}!{AnoSelecionado}.pdf")
            
            if not ClientesManutenção.empty:
                nome_arquivo = limpar_nome_arquivo(f"{row['SAP PRINCIPAL']}!{row['RAZÃO SOCIAL']}!{AnodaApuração}!{valor_total_formatado}!{resultado}!{ValorConsumoAno1Manutenção}!{diferencaCurtoeLongo}!{MultaFormatada}!{DataDaApuraçãoMANUTENÇÃO}!{DataFimApuraçãoManutenção}!{ClientesManutenção.iloc[0]['Nº INTERNO']}!{ClientesManutenção.iloc[0]['MODALIDADE']}!{meses_passados}!{total_meses_manutenção}!{MesSelecionado}!{AnoSelecionado}.pdf")
            
            if not ClientesNovoComodato.empty:
                nome_arquivo = limpar_nome_arquivo(f"{row['SAP PRINCIPAL']}!{row['RAZÃO SOCIAL']}!{AnodaApuração}!{valor_total_formatado}!{resultado}!{ValorConsumoAno1NovoComodato}!{diferencaCurtoeLongo}!{MultaFormatada}!{DataDaApuraçãoNovoComodato}!{DataFimApuraçãoNovoComodato}!{ClientesNovoComodato.iloc[0]['Nº INTERNO']}!{ClientesNovoComodato.iloc[0]['MODALIDADE']}!{meses_passados}!{total_meses_NovoComodato}!{MesSelecionado}!{AnoSelecionado}.pdf")
             
            if not ClientesacordodeConsumo.empty:
                nome_arquivo = limpar_nome_arquivo(f"{row['SAP PRINCIPAL']}!{row['RAZÃO SOCIAL']}!{AnodaApuração}!{valor_total_formatado}!{resultado}!{ValorConsumoAno1AcordoDeConsumo}!{diferencaCurtoeLongo}!{MultaFormatada}!{DataDaApuraçãoAcordodeConsumo}!{DataFimApuraçãoNacordodeconsumo}!{ClientesacordodeConsumo.iloc[0]['Nº INTERNO']}!{ClientesacordodeConsumo.iloc[0]['MODALIDADE']}!{meses_passados}!{total_meses_acordodeconsumo}!{MesSelecionado}!{AnoSelecionado}.pdf")
   
            
            largura_pagina = 620
            altura_pagina = 400
            
            buffer = io.BytesIO()
            c = pdf_canvas.Canvas(buffer,pagesize=(largura_pagina, altura_pagina))

            
            c.drawString(200,380, f"Relatório Apuração  - {TituloRelatorio}")

            def resource_path(relative_path):
                """Retorna o caminho absoluto para o recurso, funcionando com PyInstaller"""
                if hasattr(sys, '_MEIPASS'):
                    return os.path.join(sys._MEIPASS, relative_path)
                return os.path.join(os.path.abspath("."), relative_path)

            # Agora você usa assim:
            caminho_imagem = resource_path("images/logo.png")

          

            c.drawImage(caminho_imagem, 10, 370, width=20, height=20)
            
            if not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:

                InformaçõesDoContratoTotal = [['Informações da Apuração'],['Data Inicio',DataInicioApuraçãoLongoFormatada],['Data Fim',DataFimApuraçãoFormatada],["Mesês faltantes - Contrato ",str(total_meses) + ' Meses'],['Mesês Apurados - Ano Corrente ',meses_passados]]
            if BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:

                InformaçõesDoContratoTotal = [['Informações da Apuração'],['Data Inicio',DataInicioApuraçãoCurtoFormatada],['Data Fim',DataFimApuraçãoFormatadaCurto],["Mesês faltantes - Contrato ",str(total_mesesCurto) + ' Meses'],['Mesês Apurados - Ano Corrente ',meses_passados]]

            if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:

                InformaçõesDoContratoTotal = [['Informações da Apuração'],['Data Inicio',DataInicioApuraçãoLongoFormatada],['Data Fim',DataFimApuraçãoFormatada],["Mesês faltantes - Contrato ",str(total_meses) + ' Meses'],['Mesês Apurados - Ano Corrente ',meses_passados]]

            if not ClientesManutenção.empty:

                InformaçõesDoContratoTotal = [['Informações da Apuração'],['Data Inicio',DataInicioApuraçãoManutençãoFormatada],['Data Fim',DataFimMANUTENÇÃOFormatada],["Mesês faltantes - Contrato ",str(total_meses_manutenção) + ' Meses'],['Mesês Apurados - Ano Corrente ',meses_passados]]

            if not ClientesNovoComodato.empty:

                InformaçõesDoContratoTotal = [['Informações da Apuração'],['Data Inicio',DataInicioApuraçãoNovoComodatoFormatada],['Data Fim',DataFimNovoComodatoFormatada],["Mesês faltantes - Contrato ",str(total_meses_NovoComodato) + ' Meses'],['Mesês Apurados - Ano Corrente ',meses_passados]]
            
            if not ClientesacordodeConsumo.empty:

                InformaçõesDoContratoTotal = [['Informações da Apuração'],['Data Inicio',DataInicioApuraçãoacordodeconsumoFormatada],['Data Fim',DataFimAcordoDeConsumo],["Mesês faltantes - Contrato ",str(total_meses_acordodeconsumo) + ' Meses'],['Mesês Apurados - Ano Corrente ',meses_passados]]
    

            tabela_InformaçãoDoContratoTotal = Table(InformaçõesDoContratoTotal, colWidths=[100, 150])
            tabela_InformaçãoDoContratoTotal.setStyle(StyleTituloMudado)
            AlturaTabelaInformaçõesdaApuracao, LarguraTabelaInformaçõesdaApuracao = 10, altura_pagina - 180
            tabela_InformaçãoDoContratoTotal.wrapOn(c, largura_pagina, altura_pagina)
            tabela_InformaçãoDoContratoTotal.drawOn(c, AlturaTabelaInformaçõesdaApuracao, LarguraTabelaInformaçõesdaApuracao)

            print(DataDaApuraçãoFormatadaCurto)
            print(DataDaApuraçãoFormatada)
            # Tabelas de contrato longo prazo (se houver dados)
            if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty:
                InfContratosLongo = [
                    ['Informações do Contrato'],
                    ['Modalidade', BaseLongoFiltradoCliente.iloc[0]['MODALIDADE']],
                    ['Nº Contrato', BaseLongoFiltradoCliente.iloc[0]['Nº INTERNO']],
                    ['Versão Contratual', BaseLongoFiltradoCliente.iloc[0]['VERSÃO']],
                    ['Vigência Contratual', Vigencia],
                    ['Inicio da Apuração', DataDaApuraçãoFormatadaLongo]
                ]
                tabela_contratoLongo = Table(InfContratosLongo, colWidths=[100, 150])
                tabela_contratoLongo.setStyle(StyleTituloMudado)
                x_pos_contratoLongo, y_pos_contratoLongo = 350, altura_pagina - 236
                tabela_contratoLongo.wrapOn(c, largura_pagina, altura_pagina)
                tabela_contratoLongo.drawOn(c, x_pos_contratoLongo, y_pos_contratoLongo)

            elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:    
                InfContratosLongo = [
                    ['Informações do Contrato'],
                    ['Modalidade', BaseLongoFiltradoCliente.iloc[0]['MODALIDADE']],
                    ['Nº Contrato', BaseLongoFiltradoCliente.iloc[0]['Nº INTERNO']],
                    ['Versão Contratual', BaseLongoFiltradoCliente.iloc[0]['VERSÃO']],
                    ['Vigência Contratual', Vigencia],
                    ['Inicio da Apuração', DataDaApuraçãoFormatadaLongo]
                ]
                tabela_contratoLongo = Table(InfContratosLongo, colWidths=[100, 150])
                tabela_contratoLongo.setStyle(StyleTituloMudado)
                x_pos_contratoLongo, y_pos_contratoLongo = 350, altura_pagina - 118
                tabela_contratoLongo.wrapOn(c, largura_pagina, altura_pagina)
                tabela_contratoLongo.drawOn(c, x_pos_contratoLongo, y_pos_contratoLongo)

            # Criar tabelas
            tabela_cliente = Table(InfClientes, colWidths=[120, 150])
            

            tabela_target = Table(Target, colWidths=[100, 50])

            tabela_cliente.setStyle(StyleTituloMudado)

            tabela_target.setStyle(StyleTabelaTarget)

            

            # Definir dimensões da folha


            if not BaseCurtoFiltradoCliente.empty:
                InfContratos = [ ['Informações do Contrato'],
                                ['Modalidade', BaseCurtoFiltradoCliente.iloc[0]['MODALIDADE']], 
                                ['Nº Contrato', BaseCurtoFiltradoCliente.iloc[0]['Nº INTERNO']], 
                                ['Versão Contratual', BaseCurtoFiltradoCliente.iloc[0]['VERSÃO']], 
                                ['Vigência Contratual', VigenciaCurto],
                                ['Inicio da Apuração', DataDaApuraçãoFormatadaCurto],
                                ]
                
                
                tabela_contrato = Table(InfContratos, colWidths=[100, 150])
                tabela_contrato.setStyle(StyleTituloMudado)
                tabela_contrato.wrapOn(c, largura_pagina, altura_pagina)
                x_pos_contrato, y_pos_contrato = 350, altura_pagina - 118
                tabela_contrato.drawOn(c, x_pos_contrato, y_pos_contrato)

            if not ClientesManutenção.empty:
                InfContratosMANUTENÇÃO = [ ['Informações do Contrato'],
                                ['Modalidade', ClientesManutenção.iloc[0]['MODALIDADE']], 
                                ['Nº Contrato', ClientesManutenção.iloc[0]['Nº INTERNO']], 
                                ['Versão Contratual', ClientesManutenção.iloc[0]['VERSÃO']], 
                                ['Vigência Contratual', VigenciaMANUTENÇÃO],
                                ['Inicio da Apuração', DataDaApuraçãoFormatadaMANUTENÇÃO],
                                ]
                
                
                tabela_contratoMANUTENÇÃO = Table(InfContratosMANUTENÇÃO, colWidths=[100, 150])
                tabela_contratoMANUTENÇÃO.setStyle(StyleTituloMudado)
                tabela_contratoMANUTENÇÃO.wrapOn(c, largura_pagina, altura_pagina)
                x_pos_contrato, y_pos_contrato = 350, altura_pagina - 118
                tabela_contratoMANUTENÇÃO.drawOn(c, x_pos_contrato, y_pos_contrato)
            
            if not ClientesNovoComodato.empty:
                InfContratosNovoComodato = [ ['Informações do Contrato'],
                                ['Modalidade', ClientesNovoComodato.iloc[0]['MODALIDADE']], 
                                ['Nº Contrato', ClientesNovoComodato.iloc[0]['Nº INTERNO']], 
                                ['Versão Contratual', ClientesNovoComodato.iloc[0]['VERSÃO']], 
                                ['Vigência Contratual', VigenciaNovoComodato],
                                ['Inicio da Apuração', DataDaApuraçãoFormatadaNovoComodato],
                                ]
                tabela_contratoNovoComodato = Table(InfContratosNovoComodato, colWidths=[100, 150])
                tabela_contratoNovoComodato.setStyle(StyleTituloMudado)
                tabela_contratoNovoComodato.wrapOn(c, largura_pagina, altura_pagina)
                x_pos_contrato, y_pos_contrato = 350, altura_pagina - 118
                tabela_contratoNovoComodato.drawOn(c, x_pos_contrato, y_pos_contrato)

            if not ClientesacordodeConsumo.empty:
                InfContratosAcordodeConsumo = [ ['Informações do Contrato'],
                                ['Modalidade', ClientesacordodeConsumo.iloc[0]['MODALIDADE']], 
                                ['Nº Contrato', ClientesacordodeConsumo.iloc[0]['Nº INTERNO']], 
                                ['Versão Contratual', ClientesacordodeConsumo.iloc[0]['VERSÃO']], 
                                ['Vigência Contratual', VigenciaAcordoDeConsumo],
                                ['Inicio da Apuração', DataDaApuraçãoFormatadaAcordoConsumo],
                                ]   
                tabela_contratoAcordodeConsumo = Table(InfContratosAcordodeConsumo, colWidths=[100, 150])
                tabela_contratoAcordodeConsumo.setStyle(StyleTituloMudado)
                tabela_contratoAcordodeConsumo.wrapOn(c, largura_pagina, altura_pagina)
                x_pos_contrato, y_pos_contrato = 350, altura_pagina - 118
                tabela_contratoAcordodeConsumo.drawOn(c, x_pos_contrato, y_pos_contrato)
                

            # Posições das tabelas
            x_pos_cliente, y_pos_cliente = 10, altura_pagina - 105
            
            x_pos_dadosapuração, y_pos_dadosapuração = 10, altura_pagina - 180

            # Adicionar título e tabelas ao PDF
            c.setFont("Helvetica-Bold", 9)
            
            tabela_cliente.wrapOn(c, largura_pagina, altura_pagina)
            tabela_cliente.drawOn(c, x_pos_cliente, y_pos_cliente)
            
            
            tabela_target.wrapOn(c, largura_pagina, altura_pagina)
            tabela_target.drawOn(c, 10, 20)

            # Desenhar tabela de contrato longo prazo, se houver da




            
            if not ColigadosFiltrado.empty:
                
                ColigadosFiltrado = ColigadosFiltrado[["CÓDIGO SAP","RAZÃO SOCIAL"]]
                ColigadosFiltrado['CÓDIGO SAP'] = ColigadosFiltrado['CÓDIGO SAP'].astype(int)
                ColigadosTabela =  [['Coligados'],['Sap Coligado', 'Razão Social']] + ColigadosFiltrado.iloc[::-1].values.tolist()

                c.showPage()
                altura_Pagina_Coligado = calcular_altura_tabela(len(ColigadosTabela)) + 120  # Ajustar altura conforme necessário
                c.setPageSize((largura_pagina, altura_Pagina_Coligado))

                # Criar a tabela de equipamentos de longo prazo com os dados invertidos
                tabela_Coligado = Table(ColigadosTabela, colWidths=[100, 400])
                tabela_Coligado.setStyle(StyleColigados)

                AlturaTituloColigado = altura_Pagina_Coligado - 30


                # Posição da tabela de equipamentos de longo prazo, logo abaixo do título
                y_pos_Coligado = AlturaTituloColigado - calcular_altura_tabela(len(ColigadosTabela)) - 30
                tabela_Coligado.wrapOn(c, largura_pagina, altura_pagina)
                tabela_Coligado.drawOn(c, 10, y_pos_Coligado)   


            
            if not dados_lentes.empty:
            

                # Convertendo o DataFrame para a lista com formatação
                dados_tabela_lentes = [['Produtos Consumidos',' Consumo realizado até o Fechamento','Cobrança Anual'],['LENTES', 'DESCRIÇÃO CONSUMO', 'VALOR TOTAL', 'TARGET UNIFICADO', 'DIFERENÇA', 'CÁLCULO DE MULTA']] + [
                    [linha['LENTES'], linha['DESCRIÇÃO CONSUMO'], formatar_moeda(linha['VALOR TOTAL']), '', '', '']  # Colunas extras vazias para quarta, quinta e sexta
                    for _, linha in dados_lentes.iterrows()
                ]
                
                if AnodaApuração == 'Ano 1':
                    # Adicionando os textos fixos às colunas mescladas
                    if len(dados_tabela_lentes) > 1:
                        if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty: 
                            # Verifica se há dados na tabela
                            dados_tabela_lentes[2][3] = ValorConsumoTotalAno1
                        elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno1Longo 
                        elif not BaseCurtoFiltradoCliente.empty and BaseLongoFiltradoCliente.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno1Curto   

                        elif not ClientesManutenção.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno1Manutenção 

                        elif not ClientesNovoComodato.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno1NovoComodato
                        elif not ClientesacordodeConsumo.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno1AcordoDeConsumo
                            
                              # Texto da quarta coluna mesclada
                        dados_tabela_lentes[2][4] = diferencaCurtoeLongo      # Texto da quinta coluna mesclada
                        dados_tabela_lentes[2][5] = MultaFormatada         # Texto da sexta coluna mesclada
                elif AnodaApuração == 'Ano 2':
                    if len(dados_tabela_lentes) > 1:
                        if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty: 
                            # Verifica se há dados na tabela
                            dados_tabela_lentes[2][3] = ValorConsumoTotalAno2
                        elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno2Longo 
                        elif not BaseCurtoFiltradoCliente.empty and BaseLongoFiltradoCliente.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno2Curto  

                        elif not ClientesManutenção.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno2Manutenção 
                        elif not ClientesNovoComodato.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno2NovoComodato  

                        elif not ClientesacordodeConsumo.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno2AcordoDeConsumo
                              # Texto da quarta coluna mesclada
                        dados_tabela_lentes[2][4] = diferencaCurtoeLongo      # Texto da quinta coluna mesclada
                        dados_tabela_lentes[2][5] = MultaFormatada
                elif AnodaApuração == 'Ano 3':
                    if len(dados_tabela_lentes) > 1:
                        if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty: 
                            # Verifica se há dados na tabela
                            dados_tabela_lentes[2][3] = ValorConsumoTotalAno3
                        elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno3Longo 
                        elif not BaseCurtoFiltradoCliente.empty and BaseLongoFiltradoCliente.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno3Curto  
                            
                        elif not ClientesManutenção.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno3Manutenção 
                            
                        elif not ClientesNovoComodato.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno3NovoComodato 
                        elif not ClientesacordodeConsumo.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno3AcordoDeConsumo   # Texto da quarta coluna mesclada
                        dados_tabela_lentes[2][4] = diferencaCurtoeLongo      # Texto da quinta coluna mesclada
                        dados_tabela_lentes[2][5] = MultaFormatada  
                elif AnodaApuração == 'Ano 4':
                    if len(dados_tabela_lentes) > 1:
                        if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty: 
                            # Verifica se há dados na tabela
                            dados_tabela_lentes[2][3] = ValorConsumoTotalAno4
                        elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno4Longo 
                        elif not BaseCurtoFiltradoCliente.empty and BaseLongoFiltradoCliente.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno4Curto
                        elif not ClientesManutenção.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno4Manutenção   
                        elif not ClientesNovoComodato.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno4NovoComodato   
                        elif not ClientesacordodeConsumo.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno4AcordoDeConsumo # Texto da quarta coluna mesclada
                        dados_tabela_lentes[2][4] = diferencaCurtoeLongo      # Texto da quinta coluna mesclada
                        dados_tabela_lentes[2][5] = MultaFormatada 
                elif AnodaApuração == 'Ano 5':
                    if len(dados_tabela_lentes) > 1:
                        if not BaseLongoFiltradoCliente.empty and not BaseCurtoFiltradoCliente.empty: 
                            # Verifica se há dados na tabela
                            dados_tabela_lentes[2][3] = ValorConsumoTotalAno5
                        elif not BaseLongoFiltradoCliente.empty and BaseCurtoFiltradoCliente.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno5Longo 
                        elif not BaseCurtoFiltradoCliente.empty and BaseLongoFiltradoCliente.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno5Curto
                        elif not ClientesManutenção.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno5Manutenção 
                        elif not ClientesNovoComodato.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno5NovoComodato 
                        elif not ClientesacordodeConsumo.empty:

                            dados_tabela_lentes[2][3] = ValorConsumoAno5AcordoDeConsumo     # Texto da quarta coluna mesclada
                        dados_tabela_lentes[2][4] = diferencaCurtoeLongo      # Texto da quinta coluna mesclada
                        dados_tabela_lentes[2][5] = MultaFormatada 
                # Configurando o estilo da tabela

                if resultado < 100:
                    styleConsumo = TableStyle([
                            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                            ('FONTSIZE', (0, 0), (-1, -1), 4),
                            ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
                            ('TOPPADDING', (0, 0), (-1, -1), 1),
                            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),

                            # Mesclagem de células
                            ('SPAN', (0, 0), (1, 0)),  # Mescla colunas 1 e 2 na primeira linha
                            ('SPAN', (2, 0), (4, 0)),  # Mescla colunas 3 a 5 na primeira linha
                            ('SPAN', (3, 2), (3, -1)),  # Mescla a quarta coluna a partir da terceira linha
                            ('SPAN', (4, 2), (4, -1)),  # Mescla a quinta coluna a partir da terceira linha
                            ('SPAN', (5, 2), (5, -1)),  # Mescla a sexta coluna a partir da terceira linha

                            # Ajuste de tamanho de fonte nas colunas mescladas
                            ('FONTSIZE', (3, 2), (3, -1), 9),
                            ('FONTSIZE', (4, 2), (4, -1), 9),
                            ('FONTSIZE', (5, 2), (5, -1), 9),

                            # Formatação do cabeçalho
                            ('BACKGROUND', (0, 0), (-1, 0), (68/255, 83/255, 106/255)),  # Fundo azul escuro
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),  # Texto branco
                            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                            ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),
                            ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),
                            ('FONTSIZE', (0, 0), (-1, 0), 6),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),

                            # Formatação da segunda linha
                            ('BACKGROUND', (0, 1), (-1, 1), colors.lightgrey),
                            ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),

                            # Fundo azul claro para colunas 3 a 5 na primeira linha
                            ('BACKGROUND', (2, 0), (4, 0), (132/255, 150/255, 175/255)),

                            # Bordas e grade interna
                            ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
                            ('BOX', (0, 0), (-1, -1), 0.5, colors.lightgrey),

                            # Ajustes na coluna 6
                            ('ALIGN', (5, 0), (5, 0), 'CENTER'),
                            ('VALIGN', (5, 0), (5, 0), 'MIDDLE'),
                            ('FONTSIZE', (5, 0), (5, 0), 6),
                            ('COLWIDTH', (5, 0), (5, -1), 50),

                                # Texto vermelho e negrito apenas na coluna 5 e 6 a partir da terceira linha
                            ('TEXTCOLOR', (4, 2), (5, -1), colors.red),
                            ('FONTNAME', (4, 2), (5, -1), 'Helvetica-Bold'),
                        ])
                else:
                    styleConsumo = TableStyle([
                        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 0), (-1, -1), 4),  # Tamanho da fonte para todas as células
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
                        ('TOPPADDING', (0, 0), (-1, -1), 1),
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                        
                        # Mescla as colunas 1 e 2 na primeira linha
                        ('SPAN', (0, 0), (1, 0)),  

                        # Mescla as colunas 3, 4 e 5 na primeira linha
                        ('SPAN', (2, 0), (4, 0)),

                        # Mescla a quarta, quinta e sexta colunas a partir da terceira linha
                        ('SPAN', (3, 2), (3, -1)),  
                        ('SPAN', (4, 2), (4, -1)),  
                        ('SPAN', (5, 2), (5, -1)),  
                        
                        # Aumenta o tamanho da fonte das colunas mescladas a partir da terceira linha
                        ('FONTSIZE', (3, 2), (3, -1), 9),  
                        ('FONTSIZE', (4, 2), (4, -1), 9),  
                        ('FONTSIZE', (5, 2), (5, -1), 9),  

                        # Formatação da primeira linha (cabeçalho)
                        ('BACKGROUND', (0, 0), (-1, 0), (68/255, 83/255, 106/255)),  # Fundo azul-escuro
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),  # Texto branco
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Alinhamento centralizado
                        ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),  # Linha acima da primeira linha
                        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),  # Linha abaixo da primeira linha
                        ('FONTSIZE', (0, 0), (-1, 0), 6),  # Aumenta o tamanho da fonte na primeira linha
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Negrito na primeira linha

                        # Formatação da segunda linha
                        ('BACKGROUND', (0, 1), (-1, 1), colors.lightgrey),  # Fundo cinza claro
                        ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),  # Texto em negrito na segunda linha

                        # Fundo azul claro para as colunas 3 a 5 na primeira linha
                        ('BACKGROUND', (2, 0), (4, 0), (132/255, 150/255, 175/255)),  # Fundo azul claro para as colunas 3 a 5

                        # Divisões da tabela com cinza claro
                        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Grade interna
                        ('BOX', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Borda externa

                        ('ALIGN', (5, 0), (5, 0), 'CENTER'),  # Alinhamento centralizado da coluna 6
                        ('VALIGN', (5, 0), (5, 0), 'MIDDLE'),  # Alinhamento vertical no meio
                        ('FONTSIZE', (5, 0), (5, 0), 6),  # Tamanho da fonte ajustado
                        ('COLWIDTH', (5, 0), (5, -1), 50),
                        ('TEXTCOLOR', (4, 3), (5, -1), colors.green),  
                        ('FONTNAME', (4, 3), (5, -1), 'Helvetica-Bold'),
                    ])
                    
                # Criando e configurando a tabela
                TabelaConsumo = Table(dados_tabela_lentes, colWidths=[50, 100, 80, 120, 120, 120])  # Ajustando largura da nova coluna
                TabelaConsumo.setStyle(styleConsumo)

                # Renderizando a tabela
                c.showPage()
                c.setPageSize((largura_pagina, altura_pagina_Consumo))
                y_alturatituloConsumo = 320

                y_pos_ConsumoEquipamento = y_alturatituloConsumo - calcular_altura_tabela(len(dados_lentes)) - 100
                TabelaConsumo.wrapOn(c, largura_pagina, altura_pagina_Consumo)
                TabelaConsumo.drawOn(c, 10, 15)
                
            
                
            if not EquipamentosGeraisFiltrado.empty: 
                altura_pagina_3 = calcular_altura_tabela(len(equipamentos_longo_com_cabecalho)) + 100
                c.showPage()
                c.setPageSize((largura_pagina, altura_pagina_3)) 

            # Posição do título "Equipamentos de Longo Prazo" com deslocamento fixo de 10
                y_pos_titulo_EquipamentosLongo = calcular_altura_tabela(len(equipamentos_longo_com_cabecalho)) + 100 - 20
                c.setFont("Helvetica-Bold", 9)


                # Inverter os dados da tabela de equipamentos longo prazo com a remoção de duplicatas
                equipamentos_longo_com_cabecalho_invertido = [['Equipamentos'],['SKU Equipamento', 'Descrição','N INTERNO','Série']] + EquipamentosGeraisFiltrado.iloc[::-1].values.tolist()

                # Criar a tabela de equipamentos de longo prazo com os dados invertidos
                tabela_equipamentos_longo = Table(equipamentos_longo_com_cabecalho_invertido, colWidths=[100, 150])
                tabela_equipamentos_longo.setStyle(StyleColigados)

                # Posição da tabela de equipamentos de longo prazo, logo abaixo do título
                y_pos_EquipamentosLongo = y_pos_titulo_EquipamentosLongo - calcular_altura_tabela(len(equipamentos_longo_com_cabecalho_invertido)) - 30
                tabela_equipamentos_longo.wrapOn(c, largura_pagina, altura_pagina)
                tabela_equipamentos_longo.drawOn(c, 10, y_pos_EquipamentosLongo)       

            if not BaseHistoricaFiltradaCompleta.empty:
                # Agrupar os dados por 'Mês' e 'Codigo_PN' e somar as colunas 'Quantidade' e 'Valor_Unitario'
                BaseHistoricaFiltrada = (
                    BaseHistoricaFiltrada
                    .groupby(['Ano','Mês', 'Codigo_PN', 'Item 2', 'Nome_PN'], as_index=False)
                    .agg({
                        'Descricao_Item': 'first',  # Mantém a descrição do produto
                        'Quantidade': 'sum',       # Soma a quantidade
                        'Ano': 'first',            # Mantém o ano (se único por agrupamento)
                        'Total Gross': 'sum'        # Soma os valores unitários
                    })
                )
                
                BaseHistoricaFiltradaCompleta = (
                    BaseHistoricaFiltradaCompleta
                    .groupby(['Ano','Mês', 'Codigo_PN', 'Item 2', 'Nome_PN'], as_index=False)
                    .agg({
                        'Descricao_Item': 'first',  # Mantém a descrição do produto (primeiro valor encontrado)
                        'Quantidade': 'sum',        # Soma a quantidade
                        'Total Gross': 'sum'         # Soma os valores unitários
                    })
                )
               

                # Ordenar o DataFrame por 'Mês' (decrescente) e 'Ano' (decrescente)
                BaseHistoricaFiltrada = BaseHistoricaFiltrada.sort_values(by=['Ano','Mês'], ascending=[False, False])
                
                BaseHistoricaFiltradaCompleta = BaseHistoricaFiltradaCompleta.sort_values(by=['Ano','Mês'], ascending=[False, False])

                # Converter o valor total para formato de moeda BRL
                BaseHistoricaFiltrada['Total Gross'] = BaseHistoricaFiltrada['Total Gross'].apply(
                    lambda x: format_currency(x, 'BRL', locale='pt_BR')
                )
                BaseHistoricaFiltradaCompleta['Total Gross'] = BaseHistoricaFiltradaCompleta['Total Gross'].apply(
                    lambda x: format_currency(x, 'BRL', locale='pt_BR')
                )

                # Renomear colunas, se necessário
                BaseHistoricaFiltrada.rename(columns={
                    'Codigo_PN': 'SAP Principal',
                    'Nome_PN': 'Razão Social',
                    'Item 2': 'SKU Produto',
                    'Descricao_Item': 'Descrição Produto',
                    'Quantidade': 'Quantidade',
                    'Total Gross': 'Valor Total',
                    'Mês': 'Mês',
                    'Ano': 'Ano'
                }, inplace=True)
                
                BaseHistoricaFiltradaCompleta.rename(columns={
                    'Codigo_PN': 'SAP Principal',
                    'Nome_PN': 'Razão Social',
                    'Item 2': 'SKU Produto',
                    'Descricao_Item': 'Descrição Produto',
                    'Quantidade': 'Quantidade',
                    'Total Gross': 'Valor Total',
                    'Mês': 'Mês',
                    'Ano': 'Ano'
                }, inplace=True)

                # Reordenar colunas na sequência desejada
                BaseHistoricaFiltrada = BaseHistoricaFiltrada[[
                    'SAP Principal', 'Razão Social', 'SKU Produto', 'Descrição Produto',
                    'Quantidade', 'Valor Total', 'Mês', 'Ano'
                ]]
                
                BaseHistoricaFiltradaCompleta = BaseHistoricaFiltradaCompleta[[
                    'SAP Principal', 'Razão Social', 'SKU Produto', 'Descrição Produto',
                    'Quantidade', 'Valor Total', 'Mês', 'Ano'
                ]]

                # Adicionar cabeçalho à tabela
                historico_com_cabecalho = [['Extrato de Consumo - Visão Geral'],['SAP Principal', 'Razão Social', 'SKU Produto', 'Descrição Produto',
                                            'Quantidade', 'Valor Total', 'Mês', 'Ano']] + BaseHistoricaFiltradaCompleta.values.tolist()

                # Criar a tabela de histórico
                tabela_historico = Table(historico_com_cabecalho, colWidths=[50, 170, 50, 170, 50, 50, 20, 20])

                # Estilo da tabela (você deve definir 'style' e 'calcular_altura_tabela' no seu código)
                tabela_historico.setStyle(StyleColigados)

                # Criar nova página para a tabela de histórico (se necessário)
                c.showPage()
                altura_pagina_historico = calcular_altura_tabela(len(historico_com_cabecalho)) + 120  # Ajustar altura conforme necessário
                c.setPageSize((largura_pagina, altura_pagina_historico))

                # Título da tabela de histórico
                y_pos_titulo_historico = altura_pagina_historico - 10
                c.setFont("Helvetica-Bold", 9)

                # Altura da tabela (estimativa do total necessário)
                altura_tabela = calcular_altura_tabela(len(historico_com_cabecalho))

                # Calcular a posição inicial da tabela logo abaixo do título
                y_pos_historico = y_pos_titulo_historico - altura_tabela - 10  # Ajuste de 30 unidades para espaçamento entre título e tabela

                # Desenhar a tabela
                tabela_historico.wrapOn(c, largura_pagina, altura_pagina_historico)
                tabela_historico.drawOn(c, 10, y_pos_historico)

            c.save()        

            try:
                buffer.seek(0)

                def buscar_arquivo_existente(nome_arquivo):
                    headers = {
                        "Authorization": f"Bearer {access_token_global}"
                    }

                    url = f"https://api.box.com/2.0/folders/{FOLDER_ID}/items"
                    params = {
                        "fields": "id,name",
                        "limit": 1000
                    }

                    response = requests.get(url, headers=headers, params=params)
                    if response.status_code == 200:
                        itens = response.json().get("entries", [])
                        for item in itens:
                            if item["name"] == nome_arquivo:
                                return item["id"]
                    else:
                        print(f"❌ Erro ao buscar arquivos: {response.text}")
                    return None

                def fazer_upload(buffer, nome_arquivo):
                    headers = {
                        "Authorization": f"Bearer {access_token_global}"
                    }

                    arquivo_id = buscar_arquivo_existente(nome_arquivo)

                    if arquivo_id:
                        print(f"🚫 Arquivo '{nome_arquivo}' já existe no Box (ID: {arquivo_id}). Pulando upload.")
                        return None  # Não faz upload
                    else:
                        print(f"📤 Enviando novo arquivo '{nome_arquivo}' para a pasta...")
                        files = {
                            "attributes": (None, f'{{"name": "{nome_arquivo}", "parent": {{"id": "{FOLDER_ID}"}}}}', 'application/json'),
                            "file": ('file', buffer, 'application/pdf')
                        }

                        response = requests.post(UPLOAD_URL, headers=headers, files=files)

                        if response.status_code == 401:
                            print("⚠️ Token expirado, tentando atualizar...")
                            refresh_access_token()
                            headers["Authorization"] = f"Bearer {access_token_global}"
                            response = requests.post(UPLOAD_URL, headers=headers, files=files)

                        response.raise_for_status()
                        return response

                response = fazer_upload(buffer, nome_arquivo)

                if response is not None:
                    if response.status_code in [200, 201]:
                        print(f"✅ Upload do arquivo '{nome_arquivo}' feito com sucesso!")
                    else:
                        print(f"❌ Erro no upload: {response.text}")
                else:
                    print(f"🔔 Upload do arquivo '{nome_arquivo}' não realizado pois ele já existe.")

            except requests.exceptions.HTTPError as http_err:
                messagebox.showerror("Erro HTTP", f"Erro HTTP:\n{http_err}\n\nResposta: {response.text}")
            except Exception as e:
                messagebox.showerror("Erro", f"Ocorreu um erro inesperado:\n{str(e)}")

        messagebox.showinfo("Finalizado", "Processo Finalizado. PDFs criados.")

        window.destroy()   
                    
    else:
        messagebox.showerror("Erro", "Selecione os arquivos.")

  
def selecionar_arquivo1():
    caminho_arquivo1.set(
        filedialog.askopenfilename(
            initialdir="./",
            title="Selecione a planilha de Base Unificada",
            filetypes = [
    ("Excel files", "*.xlsx"), 
    ("CSV files", "*.csv"), 
    ("All files", "*.*")
]
        )
    )

def selecionar_arquivo2():
    caminho_arquivo2.set(
        filedialog.askopenfilename(
            initialdir="./",
            title="Selecione a planilha de Base de Longo",
            filetypes = [
    ("Excel files", "*.xlsx"), 
    ("CSV files", "*.csv"), 
    ("All files", "*.*")
]
        )
    )
    
def selecionar_arquivo3():
    caminho_arquivo3.set(
        filedialog.askopenfilename(
            initialdir="./",
            title="Selecione a planilha de Coligados",
            filetypes = [
    ("Excel files", "*.xlsx"), 
    ("CSV files", "*.csv"), 
    ("All files", "*.*")
]
        )
    )
    
def selecionar_arquivo4():
    caminho_arquivo4.set(
        filedialog.askopenfilename(
            initialdir="./",
            title="Selecione a planilha de Base Historico",
            filetypes = [("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
    )

def abrir_poupup_serial():
    popup = tk.Toplevel(window)
    popup.geometry("300x400")
    popup.title("Digite o Access Key")
    popup.resizable(False, False)

    frame = tk.Frame(popup)
    frame.pack(expand=True, fill="both", pady=10)

    lista_meses = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]

    tk.Label(frame, text="Título do Documento:", font=("OpenSansRoman Bold", 13 * -1)).grid(row=2, column=0, pady=(2, 2), sticky="w")
    title_input = tk.Entry(frame, textvariable=TitleInput, font=("OpenSansRoman Bold", 13 * -1))
    title_input.grid(row=3, column=0, pady=(2, 10), padx=10, sticky="we")

    month_label = tk.Label(frame, text="Mês:", font=("OpenSansRoman Bold", 13 * -1))
    month_combobox = ttk.Combobox(frame, values=lista_meses, state="readonly", width=20)
    month_combobox.set(lista_meses[0])

    year_label = tk.Label(frame, text="Ano:", font=("OpenSansRoman Bold", 13 * -1))
    year_combobox = ttk.Combobox(frame, values=[str(year) for year in range(2000, datetime.now().year + 5)], state="readonly", width=10)
    year_combobox.set(str(datetime.now().year))

    # Inicialmente escondidos
    month_label.grid(row=4, column=0, pady=(2, 2), sticky="w")
    month_combobox.grid(row=5, column=0, pady=(2, 5), padx=10, sticky="we")
    year_label.grid(row=6, column=0, pady=(2, 2), sticky="w")
    year_combobox.grid(row=7, column=0, pady=(2, 10), padx=10, sticky="we")

    month_label.grid_remove()
    month_combobox.grid_remove()
    year_label.grid_remove()
    year_combobox.grid_remove()

    def toggle_data_entry():
        if CheckboxValue.get():
            month_label.grid()
            month_combobox.grid()
            year_label.grid()
            year_combobox.grid()
        else:
            month_label.grid_remove()
            month_combobox.grid_remove()
            year_label.grid_remove()
            year_combobox.grid_remove()

    checkbox = tk.Checkbutton(frame, text="Incluir Data", variable=CheckboxValue, command=toggle_data_entry,
                              font=("OpenSansRoman Bold", 13 * -1))
    checkbox.grid(row=8, column=0, pady=(2, 10), sticky="w")

    # Novo: SAP ÚNICO
    def toggle_sap_input():
        if SapCheckboxValue.get():
            sap_label.grid()
            sap_entry.grid()
        else:
            sap_label.grid_remove()
            sap_entry.grid_remove()

    SapCheckboxValue = tk.BooleanVar()
    SapInputValue = tk.StringVar()

    sap_checkbox = tk.Checkbutton(frame, text="SAP ÚNICO", variable=SapCheckboxValue, command=toggle_sap_input,
                                  font=("OpenSansRoman Bold", 13 * -1))
    sap_checkbox.grid(row=9, column=0, pady=(2, 2), sticky="w")

    sap_label = tk.Label(frame, text="INSIRA O SAP:", font=("OpenSansRoman Bold", 13 * -1))
    sap_entry = tk.Entry(frame, textvariable=SapInputValue, font=("OpenSansRoman Bold", 13 * -1))

    sap_label.grid(row=10, column=0, pady=(2, 2), sticky="w")
    sap_entry.grid(row=11, column=0, pady=(2, 10), padx=10, sticky="we")
    sap_label.grid_remove()
    sap_entry.grid_remove()

    def validar_access_key():
        title = TitleInput.get()
        global DateValue, MesSelecionado, AnoSelecionado, SapSelecionado, ValorSAP

        if CheckboxValue.get():
            nome_mes = month_combobox.get()
            MesSelecionado = nome_mes
            AnoSelecionado = int(year_combobox.get())

            numero_mes = lista_meses.index(nome_mes) + 1
            ultimo_dia = calendar.monthrange(AnoSelecionado, numero_mes)[1]

            DateValue = datetime(AnoSelecionado, numero_mes, ultimo_dia)
        else:
            DateValue = None
            MesSelecionado = None
            AnoSelecionado = None

        SapSelecionado = SapCheckboxValue.get()
        ValorSAP = SapInputValue.get() if SapSelecionado else None

        print(f"Título: {title}")
        print(f"Incluir Data? {CheckboxValue.get()}")
        print(f"Data gerada: {DateValue}")
        print(f"SAP Único? {SapSelecionado}")
        print(f"Valor SAP: {ValorSAP}")

        processar_arquivos()

    tk.Button(frame, text="Confirmar", command=validar_access_key).grid(row=12, column=0, pady=(10, 10))


    


window = tk.Tk()
window.geometry("509x250")
window.configure(bg="#FFFFFF")

caminho_arquivo1 = tk.StringVar()
caminho_arquivo2 = tk.StringVar()
caminho_arquivo3 = tk.StringVar()
caminho_arquivo4 = tk.StringVar()
caminho_arquivo5 = tk.StringVar()
CheckboxValue = tk.BooleanVar()
TitleInput = tk.StringVar()
SapCheckboxValue = tk.BooleanVar()
SapInputValue = tk.StringVar()
DateValue = None
MesSelecionado = None
AnoSelecionado = None
SapSelecionado = None
ValorSAP = None

canvas = TkCanvas(window, bg="#FFFFFF", height=400, width=509, bd=0, highlightthickness=0, relief="ridge")
canvas.place(x=0, y=0)
canvas.create_rectangle(0.0, 0.0, 509.0, 500.0, fill="#D9D9D9", outline="")
canvas.create_rectangle(0.0, 0.0, 237.0, 500.0, fill="#972323", outline="")

canvas.create_text(289.0, 31.0, anchor="nw", text="Selecione a Base Unificada", fill="#0F0F0F", font=("OpenSansRoman Bold", 13 * -1))
canvas.create_text(16.0, 70.0, anchor="nw", text="Gerador de Apuração", fill="#FFFFFF", font=("OpenSansRoman Bold", 22 * -1))
canvas.create_text(260.0, 120.0, anchor="nw", text="Selecione a planilha de Base Historica", fill="#0F0F0F", font=("OpenSansRoman Bold", 13 * -1))

button_1 = Button(text="Base Unificada", borderwidth=1, highlightthickness=0, command=selecionar_arquivo1, relief="flat")
button_1.place(x=299.0, y=67.0, width=147.0, height=28.0)

button_4 = Button(text="Base Historica", borderwidth=0, highlightthickness=0, command=selecionar_arquivo4, relief="flat")
button_4.place(x=299.0, y=150.0, width=147.0, height=28.0)

button_5 = Button(text="Gerar arquivo", borderwidth=0, highlightthickness=0, command=abrir_poupup_serial, relief="flat")
button_5.place(x=275.0, y=200.0, width=190.0, height=33.0)

canvas.create_text(100.0, 130.0, anchor="nw", text="J&J", fill="#FFFFFF", font=("Roboto Mono", 20 * -1))

window.resizable(False, False)
window.title("Gerador de Apuração")
window.mainloop()
