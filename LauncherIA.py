
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
import math
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

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
    """
    Função refatorada para processar arquivos e gerar relatórios em PDF.
    Layout corrigido com base na foto de referência.
    """
    
    # ============================================================================
    # SEÇÃO 1: OBTENÇÃO DOS PARÂMETROS E CONFIGURAÇÕES INICIAIS
    # ============================================================================
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
    
    # Validação inicial dos arquivos
    if not (BaseCurtoCaminho and BaseHistoricaCaminho and TituloRelatorio):
        messagebox.showerror("Erro", "Arquivos obrigatórios não selecionados!")
        return
        
    # ============================================================================
    # SEÇÃO 2: CARREGAMENTO E VALIDAÇÃO DOS DADOS
    # ============================================================================
    try:
        # Validação das colunas
        df_validacao = pd.read_excel(BaseCurtoCaminho)
        validar_colunas(df_validacao)
        
        # Definição da data atual baseada no checkbox
        if ValorCheckBox:
            data_atual = DateValue
        else:
            data_atual = datetime.today()
            
        messagebox.showinfo("Processo Iniciado", "Arquivos selecionados, iniciando o processo")
        
        # Carregamento das bases de dados
        Base = pd.read_excel(BaseCurtoCaminho)
        BaseConsumo = pd.read_excel(BaseCurtoCaminho, engine='openpyxl', sheet_name='Bases - Consumo')
        BaseHistorica = pd.read_excel(BaseHistoricaCaminho, engine='openpyxl', sheet_name='Sheet1')
        BaseHistoricaCompleta = pd.read_excel(BaseHistoricaCaminho, engine='openpyxl', sheet_name='Sheet1')
        
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar arquivos: {str(e)}")
        return
    
    # ============================================================================
    # SEÇÃO 3: PROCESSAMENTO E FILTRAGEM DOS DADOS
    # ============================================================================
    # Processamento da base principal
    Base['INICIO DA APURAÇÃO'] = pd.to_numeric(Base['INICIO DA APURAÇÃO'], errors='coerce')
    Base['DATA INICIAL'] = pd.to_datetime('1899-12-30') + pd.to_timedelta(Base['INICIO DA APURAÇÃO'], unit='D')
    Base['PRAZO APURACAO'] = pd.to_numeric(Base['PRAZO APURACAO'], errors='coerce').fillna(0).astype(int)
    Base['DATA FINAL'] = Base['DATA INICIAL'] + Base['PRAZO APURACAO'].apply(lambda x: DateOffset(months=int(x)))
    Base = Base[Base['DATA INICIAL'] < data_atual]
    Base = Base[Base['ATIVO OU INATIVO'] == 'ATIVO']
    
    # Filtros para coligados
    ColigadosFiltros = BaseConsumo[
        (BaseConsumo['SAP'] == 'Coligado') &
        (BaseConsumo['ATIVO OU INATIVO'] == 'ATIVO')
    ][['CÓDIGO SAP', 'RAZÃO SOCIAL', 'SAP PRINCIPAL']].drop_duplicates()
    
    # Filtros por modalidade
    filtro_base = (
        (Base['ATIVO OU INATIVO'] == 'ATIVO') & 
        (Base['LINHA DO CONTRATO'] == 'Principal') & 
        (Base['INICIO DA APURAÇÃO'].notna())
    )
    
    # Filtros para cada modalidade
    df_filtrado = Base[filtro_base & (Base['MODALIDADE'] == 'Compra e Venda com consumo')].drop_duplicates(subset='SAP PRINCIPAL')
    BaseLongoFiltrado = Base[filtro_base & (Base['MODALIDADE'] == 'NOVA LOCAÇÃO')].drop_duplicates(subset='SAP PRINCIPAL')
    MANUTENÇÃO = Base[filtro_base & (Base['MODALIDADE'] == 'MANUTENÇÃO') & (Base['CONSUMO ANO 1'].notna())].drop_duplicates(subset='SAP PRINCIPAL')
    NovoComodato = Base[filtro_base & (Base['MODALIDADE'] == 'NOVO COMODATO')].drop_duplicates(subset='SAP PRINCIPAL')
    acordodeconsumo = Base[filtro_base & (Base['MODALIDADE'] == 'Acordo de Consumo')].drop_duplicates(subset='SAP PRINCIPAL')
    
    # Equipamentos gerais
    EquipamentosGerais = Base[(Base['ATIVO OU INATIVO'] == 'ATIVO') & (Base['EQUIPAMENTO'].notnull())]
    
    # Concatenação dos dados filtrados
    df_concatenado = pd.concat([BaseLongoFiltrado, df_filtrado, MANUTENÇÃO, NovoComodato, acordodeconsumo], ignore_index=True)
    df_concatenado = df_concatenado.drop_duplicates(subset=["SAP PRINCIPAL"], keep="first")
    
    # Filtro por SAP se selecionado
    print(SapSelecionado)
    print(ValorSAP)
    
    if SapSelecionado == True:
        ValorSapInt = int(ValorSAP)
        df_concatenado = df_concatenado[df_concatenado["SAP PRINCIPAL"] == ValorSapInt]

    df_concatenado = df_concatenado.drop_duplicates(subset=["SAP PRINCIPAL"], keep="first")
    df_concatenado.to_excel("excelfiltrado.xlsx")
    
    # ============================================================================
    # SEÇÃO 4: CONFIGURAÇÃO DE ESTILOS PARA TABELAS
    # ============================================================================
    try:
        locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')
    except:
        locale.setlocale(locale.LC_TIME, 'Portuguese')

    # Estilo para informações gerais (cliente/contrato)
    StyleInformacoes = TableStyle([
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
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('SPAN', (0, 0), (-1, 0)),
        ('BACKGROUND', (0, 0), (-1, 0), (68/255, 83/255, 106/255)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
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
            ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Grade interna
            ('BOX', (0, 0), (-1, -1), 0.5, colors.lightgrey),  # Borda externa

            ])
    
      # Style para equipamentos
    StyleEquipamentos = TableStyle([
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                
                # Cabeçalho principal
                ('SPAN', (0, 0), (-1, 0)),
                ('BACKGROUND', (0, 0), (-1, 0), (68/255, 83/255, 106/255)),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                
                # Sub-cabeçalho
                ('BACKGROUND', (0, 1), (-1, 1), colors.lightgrey),
                ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
                ('ALIGN', (0, 1), (-1, 1), 'CENTER'),
                
                # Bordas
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('INNERGRID', (0, 1), (-1, -1), 0.5, colors.lightgrey),
            ])


    messagebox.showinfo("Salvar Arquivos", "Iniciar o processo de salvar!")

    # ============================================================================
    # SEÇÃO 5: FUNÇÕES AUXILIARES
    # ============================================================================
    def formatar_moeda(valor):
        if pd.isna(valor) or valor == 0:
            return "R$ 0,00"
        return f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

    def validar_formatar_consumo(valor):
        if pd.notna(valor):
            return formatar_moeda(valor)
        return "R$ 0,00"

    def calcular_ano_referencia(data_inicio):
        data_inicio = datetime.strptime(data_inicio, "%d/%m/%Y")
        if ValorCheckBox == True:
            data_atual_calc = DateValue
        else:
            data_atual_calc = datetime.today()
        
        diferenca_meses = (data_atual_calc.year - data_inicio.year) * 12 + (data_atual_calc.month - data_inicio.month)
        
        if data_atual_calc.day < data_inicio.day:
            diferenca_meses -= 1
        
        ano = (diferenca_meses // 12) + 1
        
        # Limitar a 10 anos
        if ano > 10:
            ano = 10
            
        return f"Ano {ano}"

    # ============================================================================
    # SEÇÃO 6: PROCESSAMENTO POR CLIENTE
    # ============================================================================
    contador = 0
    data_base_excel = datetime(1899, 12, 30)
    
    for index, row in df_concatenado.iterrows():
        contador += 1
        sap_principal_filtro = row['SAP PRINCIPAL']
        Ninterno = row['Nº INTERNO']
        Versão = row['VERSÃO']

        print(f"Processando cliente {contador}: {sap_principal_filtro}")

        # ============================================================================
        # SEÇÃO 7: BUSCAR TODAS AS MODALIDADES DO CLIENTE
        # ============================================================================
        todas_modalidades_cliente = Base[
            (Base['SAP PRINCIPAL'] == sap_principal_filtro) & 
            (Base['ATIVO OU INATIVO'] == 'ATIVO') & 
            (Base['LINHA DO CONTRATO'] == 'Principal') & 
            (Base['INICIO DA APURAÇÃO'].notna()) &
            (Base['DATA FINAL'] > data_atual)
        ].copy()

        # Agrupar por modalidade
        modalidades_dict = {}
        for _, modal_row in todas_modalidades_cliente.iterrows():
            modalidade = modal_row['MODALIDADE']
            if modalidade not in modalidades_dict:
                modalidades_dict[modalidade] = []
            modalidades_dict[modalidade].append(modal_row)

        print(f"Cliente {sap_principal_filtro} possui modalidades: {list(modalidades_dict.keys())}")

        # ============================================================================
        # SEÇÃO 8: PREPARAÇÃO DE DADOS CONSOLIDADOS
        # ============================================================================
        FiltrandoLentes = BaseConsumo[(BaseConsumo['CÓDIGO SAP'] == sap_principal_filtro) & 
                                     (BaseConsumo['Nº INTERNO'] == Ninterno) & 
                                     (BaseConsumo['ATIVO OU INATIVO'] == "ATIVO")]

        lentesFiltroHistorico = FiltrandoLentes['SKU PRODUTO'].dropna().unique().tolist()

        if not lentesFiltroHistorico:
            lentesFiltroHistorico = [
                "ICB00", "PCB00", "ZCB00", "ZCT00", "ZFR00", "ZKB00", "ZLB00",
                "ZMA00", "ZMB00", "ZMT00", "ZXR00", "ZXT00", "DFW00", "DCB00",
                "DIB00", "DIU00", "DFR00", "DET00", "DEN00"
            ]

        ColigadosFiltrado = ColigadosFiltros[ColigadosFiltros['SAP PRINCIPAL'] == sap_principal_filtro]

        # ============================================================================
        # SEÇÃO 9: CALCULAR DATA DE REFERÊNCIA (CONTRATO MAIS ANTIGO)
        # ============================================================================
        
        data_inicio_mais_antiga = None
        contrato_referencia = None
        
        for modalidade, contratos in modalidades_dict.items():
            for contrato in contratos:
                data_inicio_contrato = data_base_excel + timedelta(contrato['INICIO DA APURAÇÃO'])
                if data_inicio_mais_antiga is None or data_inicio_contrato < data_inicio_mais_antiga:
                    data_inicio_mais_antiga = data_inicio_contrato
                    contrato_referencia = contrato

        if contrato_referencia is not None:
            dados_referencia = contrato_referencia
        else:
            dados_referencia = row

        # ============================================================================
        # SEÇÃO 10: PROCESSAMENTO DE DATAS E APURAÇÃO
        # ============================================================================
        DataDaApuração = data_base_excel + timedelta(dados_referencia.get('INICIO DA APURAÇÃO'))
        DataFimApuração = DataDaApuração + relativedelta(months=int(dados_referencia.get('PRAZO APURACAO', 0)))
        
        mes_extenso = meses_portugues[DataDaApuração.month]
        anodeApuracaosemfromatar = DataDaApuração.year
        
        DataDaApuraçãoFormatada = f"{mes_extenso} de {anodeApuracaosemfromatar}"
        DataFimApuraçãoFormatada = DataFimApuração.strftime('%d/%m/%Y')

        dif_anos = DataFimApuração.year - data_atual.year
        dif_meses = DataFimApuração.month - data_atual.month
        total_meses = (dif_anos * 12) + dif_meses - 1

        DataInicioApuraçãoFormatada = DataDaApuração.strftime('%d/%m/%Y')

        # Processamento das datas de vigência
        try:
            if isinstance(dados_referencia['DT. INÍCIO'], (int, float)) and dados_referencia['DT. INÍCIO'] > 60:
                DataInicio = data_base_excel + timedelta(dados_referencia['DT. INÍCIO'] - pd.Timedelta(days=2))
            else:
                DataInicio = pd.to_datetime(dados_referencia['DT. INÍCIO'])
                
            if isinstance(dados_referencia['DT. FINAL'], (int, float)) and dados_referencia['DT. FINAL'] > 60:
                DataFim = data_base_excel + timedelta(dados_referencia['DT. FINAL'] - pd.Timedelta(days=2))
            else:
                DataFim = pd.to_datetime(dados_referencia['DT. FINAL'])
                
        except (TypeError, ValueError):
            serial_inicio = int(dados_referencia['DT. INÍCIO'])
            serial_fim = int(dados_referencia['DT. FINAL'])
            DataInicio = datetime(1900, 1, 1) + timedelta(days=serial_inicio - 2)
            DataFim = datetime(1900, 1, 1) + timedelta(days=serial_fim - 2)

        DataInicioFormatada = DataInicio.strftime('%d/%m/%Y')
        DataFimFormatada = DataFim.strftime('%d/%m/%Y')
        Vigencia = f"{DataInicioFormatada} - {DataFimFormatada}"

        AnodaApuração = calcular_ano_referencia(DataInicioFormatada)

        # ============================================================================
        # AJSUTANDO STYLE DO CONSUMO UNIFICADO 
        # ============================================================================

        coluna_destacada = {
            'Ano 1': 1,
            'Ano 2': 2,
            'Ano 3': 3,
            'Ano 4': 4,
            'Ano 5': 5
        }.get(AnodaApuração, 1)
        StyleConsumoUnificado = TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 5),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
        ('TOPPADDING', (0, 0), (-1, -1), 1),
        # Cabeçalho
        ('SPAN', (0, 0), (-1, 0)),
        ('BACKGROUND', (0, 0), (-1, 0), (192/255, 0/255, 10/255)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),
        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),
        ('FONTSIZE', (0, 0), (-1, 0), 6),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        # Segunda linha
        ('BACKGROUND', (0, 1), (-1, 1), (192/255, 0/255, 10/255)),
        ('TEXTCOLOR', (0, 1), (-1, 1), colors.white),
        ('ALIGN', (0, 1), (-1, 1), 'CENTER'),
        # Terceira linha
        ('BACKGROUND', (0, 2), (-1, 2), (230/255, 230/255, 230/255)),
        ('TEXTCOLOR', (0, 2), (-1, 2), colors.black),
        ('ALIGN', (0, 2), (-1, 2), 'LEFT'),
        # Primeira coluna da terceira linha pra baixo
        ('BACKGROUND', (0, 2), (0, -1), (230/255, 230/255, 230/255)),
        ('TEXTCOLOR', (0, 2), (0, -1), colors.black),
        # Títulos das linhas restantes
        ('TEXTCOLOR', (0, 3), (0, -1), colors.black),
        ('FONTNAME', (0, 3), (0, -1), 'Helvetica-Bold'),
        # Fundo restante
        ('BACKGROUND', (1, 3), (-1, -1), colors.white),
        # Coluna destacada em negrito
        ('FONTNAME', (coluna_destacada, 2), (coluna_destacada, -1), 'Helvetica-Bold'),
        # Grade
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.lightgrey),
    ])
    
        

        # ============================================================================
        # SEÇÃO 11: CONSUMO UNIFICADO CORRIGIDO (TARGET + VALOR CONSUMIDO)
        # ============================================================================
        # Target Unificado: somar consumo anos 1-10 de todos os contratos
        target_unificado = {}
        for ano in range(1, 11):
            target_unificado[f'ano_{ano}'] = 0
            
        # Somar consumo de todas as modalidades para target
        for modalidade, contratos in modalidades_dict.items():
            for contrato in contratos:
                for ano in range(1, 11):
                    coluna_consumo = f'CONSUMO ANO {ano}'
                    valor_consumo = pd.to_numeric(contrato.get(coluna_consumo), errors='coerce')
                    if pd.notna(valor_consumo):
                        target_unificado[f'ano_{ano}'] += valor_consumo

        # ============================================================================
        # SEÇÃO 12: VALOR CONSUMIDO - BUSCAR NA BASE HISTÓRICA
        # ============================================================================
        BaseHistorica['DataApuração'] = pd.to_datetime(BaseHistorica['Ano'].astype(str) + '-' + BaseHistorica['Mês'].astype(str).str.zfill(2))
        
        # Valor consumido: buscar na base histórica baseado no ano do contrato
        valor_consumido = {}
        for ano in range(1, 11):
            # Calcular período do ano baseado no contrato mais antigo
            inicio_periodo = DataDaApuração + relativedelta(months=(ano-1)*12)
            fim_periodo = DataDaApuração + relativedelta(months=ano*12-1)
            
            # Filtrar histórico para este período
            historico_ano = BaseHistorica[
                (
                    (BaseHistorica['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistorica['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistorica['Item 2'].isin(lentesFiltroHistorico)) &
                (BaseHistorica['DataApuração'] >= inicio_periodo.strftime('%Y-%m')) &
                (BaseHistorica['DataApuração'] <= fim_periodo.strftime('%Y-%m'))
            ]
            
            # Somar valores do histórico
            valor_consumido[f'ano_{ano}'] = historico_ano['Total Gross'].sum() if not historico_ano.empty else 0

        # Calcular ano atual da apuração
        ano_atual = int(AnodaApuração.split()[1])
        data_inicio_ano_atual = DataDaApuração + relativedelta(months=(ano_atual-1)*12)
        meses_apurados = (data_atual.year - data_inicio_ano_atual.year) * 12 + (data_atual.month - data_inicio_ano_atual.month)

        if data_atual.day >= data_inicio_ano_atual.day:
            meses_apurados += 1

        meses_passados = (meses_apurados % 12) or 12

        # ============================================================================
        # SEÇÃO 13: PROCESSAMENTO DE EQUIPAMENTOS
        # ============================================================================
        EquipamentosGeraisFiltrado = EquipamentosGerais[
            (EquipamentosGerais['SAP PRINCIPAL'] == sap_principal_filtro)
        ][['EQUIPAMENTO', 'DESCRIÇÃO EQUIPAMENTO', 'Nº INTERNO', 'SÉRIE']].drop_duplicates(
            subset=['EQUIPAMENTO', 'DESCRIÇÃO EQUIPAMENTO', 'Nº INTERNO', 'SÉRIE']
        )

        # ============================================================================
        # SEÇÃO 14: PREPARAÇÃO DOS DADOS DO CLIENTE
        # ============================================================================
        RazaoSocialCompleta = f"{row['SAP PRINCIPAL']} - {row['RAZÃO SOCIAL']}"
        
        # Obter dados das lentes para produtos consumidos
        skus = FiltrandoLentes['SKU PRODUTO'].dropna().astype(str).tolist()
        descricoes = FiltrandoLentes['DESCRIÇÃO CONSUMO'].dropna().astype(str).tolist()

        if not skus:
            skus = ["ZCB00", "PCB00", "ZCT00", "ICB00", "DCB00", "DIU00", "DIB00", "ZMA00", "ZKB00", "ZLB00", "ZMB00", "ZMT00", "ZXR00", "ZXT00", "ZFR00", "DFW00", "DFR00", "DEN00", "DET00"]
            descricoes = ["LIO TECNIS ONE - peça única", "LIO TECNIS ITEC PRELOADED - peça única", "LIO TECNIS ONE TÓRICA - peça única", "LIO TECNIS Eyhance", "TECNIS SIMPLICITY DCB", "EYHANCE TORIC II SIMPLICITY", "TECNIS EYHANCE SIMPLICITY", "LIO TECNIS MF - 3 peças", "TECNIS ONE MF Low Add - Peça Única +2.75", "TECNIS ONE MF Low Add - Peça Única +3.25", "TECNIS ONE MF- Peça Única +4.00", "TECNIS ONE TÓRICA MF - Peça Única +4.00", "TECNIS SYMFONY - Peça Única", "TECNIS SYMFONY TÓRICA - Peça Únic", "LIO TECNIS Synergy", "TECNIS SYNERGY TORIC SIMPLICITY - Peça Única", "TECNIS SYNERGY SIMPLICITY - Peça Única", "TECNIS PURESEE Simplicity - peça única", "TECNIS PURESEE Simplicity Tórica - peça única"]

        # ============================================================================
        # SEÇÃO 15: GERAÇÃO DO PDF COM LAYOUT CORRIGIDO
        # ============================================================================
        tem_multiplas_modalidades = len(modalidades_dict) > 1
        
        if tem_multiplas_modalidades:
            nome_arquivo = f"Relatório_Unificado_{sap_principal_filtro}_{AnodaApuração.replace(' ', '_')}.pdf"
        else:
            modalidade_unica = list(modalidades_dict.keys())[0]
            nome_arquivo = f"Relatório_{modalidade_unica.replace(' ', '_')}_{sap_principal_filtro}_{AnodaApuração.replace(' ', '_')}.pdf"
        
        # Criar PDF com ReportLab
        c = pdf_canvas.Canvas(nome_arquivo, pagesize=letter)
        width, height = letter
        
        # ============================================================================
        # SEÇÃO 16: CABEÇALHO DO RELATÓRIO
        # ============================================================================
        # Logo da J&J (simulado)
        c.setFillColor(colors.red)
        c.rect(30, height - 60, 40, 30, fill=1)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 16)
        c.drawString(35, height - 50, "J&J")
        
        # Título do relatório
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 16)
        titulo_relatorio = f"Relatório Apuração - {DataDaApuraçãoFormatada}"
        c.drawCentredString(width/2, height - 40, titulo_relatorio)
        
        y_position = height - 80

        # ============================================================================
        # SEÇÃO 17: LAYOUT CORRIGIDO - CLIENTE AO LADO DO CONTRATO
        # ============================================================================

        # Preparar dados do cliente
        if pd.isnull(row['SAM']) or row['SAM'] == '':
            dados_cliente = [
                ['Informações do Cliente'],
                ['Sap Principal', RazaoSocialCompleta], 
                ['Consultor', str(row['CONSULTOR'])], 
                ['Distrital', str(row['DISTRITAL'])], 
                ['Sam', '-']
            ]
        else:
            dados_cliente = [
                ['Informações do Cliente'],
                ['Sap Principal', RazaoSocialCompleta], 
                ['Consultor', str(row['CONSULTOR'])], 
                ['Distrital', str(row['DISTRITAL'])], 
                ['Sam', str(row['SAM'])]
            ]

        # =============================================
        # Cliente (sempre à esquerda)
        # =============================================
        tabela_cliente = Table(dados_cliente, colWidths=[90, 160])
        tabela_cliente.setStyle(StyleInformacoes)

        altura_cliente = tabela_cliente.wrap(width, height)[1]
        tabela_cliente.drawOn(c, 10, y_position - altura_cliente)

        # Posição inicial para os contratos (à direita)
        x_contrato = 320
        y_contrato = y_position

        # =============================================
        # Modalidades em cascata (à direita)
        # =============================================
        for modalidade, contratos in modalidades_dict.items():
            for contrato in contratos:
                dados_contrato = [
                    ['Informações do Contrato'],
                    ['Modalidade', modalidade],
                    ['Nº Contrato', str(contrato.get('Nº INTERNO', 'N/A'))],
                    ['Versão Contratual', str(contrato.get('VERSÃO', 'Contrato Raiz'))],
                    ['Vigência Contratual', Vigencia],
                    ['Inicio da Apuração', DataDaApuraçãoFormatada]
                ]

            tabela_contrato = Table(dados_contrato, colWidths=[100, 160])
            tabela_contrato.setStyle(StyleInformacoes)

            altura_contrato = tabela_contrato.wrap(width, height)[1]

            # Quebra de página se não couber
            if y_contrato - altura_contrato < 50:
                c.showPage()
                y_contrato = height - 50

            tabela_contrato.drawOn(c, x_contrato, y_contrato - altura_contrato)

            # Próxima tabela um pouco abaixo (cascata)o
            y_contrato -= altura_contrato + 10

        # =============================================
        # Ajusta y_position geral
        # =============================================
        y_position = min(y_position - altura_cliente, y_contrato) - 20

        # ============================================================================
        # SEÇÃO 18: INFORMAÇÕES DA APURAÇÃO (EMBAIXO)
        # ============================================================================
        dados_apuracao = [
            ['Informações da Apuração'],
            ['Data Inicio', DataInicioApuraçãoFormatada],
            ['Data Fim', DataFimApuraçãoFormatada],
            ['Meses faltantes - Contrato', f"{total_meses} Meses"],
            ['Meses Apurados - Ano Corrente', str(meses_passados)]
        ]
        
        tabela_apuracao = Table(dados_apuracao, colWidths=[150, 200])
        tabela_apuracao.setStyle(StyleInformacoes)
        
        altura_tabela_apuracao = tabela_apuracao.wrap(350, height)[1]
        tabela_apuracao.drawOn(c, 10, y_position - altura_tabela_apuracao)
        y_position -= altura_tabela_apuracao + 30

        # ============================================================================
        # SEÇÃO 19: CONSUMO UNIFICADO (10 ANOS, SUPORTE A MÚLTIPLAS MODALIDADES)
        # ============================================================================
        # Descobre o maior prazo em anos (arredondado para cima)
        max_anos = 0
        for contratos in modalidades_dict.values():
            for contrato in contratos:
                prazo_meses = contrato.get('PRAZO APURACAO', 0)  # em meses
                prazo_anos = math.ceil(prazo_meses / 12)          # converte para anos
                if prazo_anos > max_anos:
                    max_anos = prazo_anos
        if y_position < 300:
            c.showPage()
            y_position = height - 50

        consumo_data = [['Consumo Unificado']]

        # Cabeçalhos para 10 anos
        anos_headers = [''] + [f'Ano {i}' for i in range(1, max_anos + 1)]
        consumo_data.append(anos_headers)

        # Meta %
        meta_row = ['Meta %'] + ['100%'] * max_anos
        consumo_data.append(meta_row)

        # Totais unificados (vão acumulando todas as modalidades)
        total_target = {f'ano_{i}': 0 for i in range(1, max_anos + 1)}
        total_consumido = {f'ano_{i}': 0 for i in range(1, max_anos + 1)}

        # 🔹 Loop em todas as modalidades do SAP
        for modalidade, contratos in modalidades_dict.items():
            for contrato in contratos:
                target_row = [f"Target - {contrato.get('Nº INTERNO', '')}({modalidade})"]

                for ano in range(1, max_anos + 1):
                    valor_target = pd.to_numeric(contrato.get(f'CONSUMO ANO {ano}'), errors='coerce') or 0
                    valor_consumo = valor_consumido.get(f'ano_{ano}', 0)

                    target_row.append(formatar_moeda(valor_target))

                    total_target[f'ano_{ano}'] += valor_target
                    total_consumido[f'ano_{ano}'] += valor_consumo

                consumo_data.append(target_row)

        # 🔹 Linha Target Unificado
        target_unif_row = ['Target Unificado']
        for ano in range(1, max_anos + 1):
            target_unif_row.append(formatar_moeda(total_target[f'ano_{ano}']))
        consumo_data.append(target_unif_row)

        valor_consumido_row = ['Valor Consumido - Unificado']
        for ano in range(1, max_anos + 1):
            valor_consumido_row.append(formatar_moeda(total_consumido[f'ano_{ano}']))
        consumo_data.append(valor_consumido_row)

        # 🔹 Linha Percentual de Atingimento
        percentual_row = ['Percentual de Atingimento']
        percentuais = {}  # salva como número
        for ano in range(1, max_anos + 1):
            t = total_target[f'ano_{ano}']
            c_val = total_consumido[f'ano_{ano}']
            if t > 0 and c_val > 0:
                perc = (c_val / t) * 100
                percentuais[f'ano_{ano}'] = perc  # salva como número
                percentual_row.append(f'{perc:.2f}%')
            else:
                percentuais[f'ano_{ano}'] = 0
                percentual_row.append('0.00%')

        consumo_data.append(percentual_row)

        # Ajustar largura das colunas (1 label + 10 anos)
        col_width = (width - 180) / (1 + max_anos)
        col_widths = [150] + [col_width] * max_anos

        tabela_consumo = Table(consumo_data, colWidths=col_widths)
        tabela_consumo.setStyle(StyleConsumoUnificado)

        altura_tabela_consumo = tabela_consumo.wrap(width, height)[1]
        if y_position - altura_tabela_consumo < 50:
            c.showPage()
            y_position = height - 50

        tabela_consumo.drawOn(c, 10, y_position - altura_tabela_consumo)
        y_position -= altura_tabela_consumo + 30
        # ============================================================================
        # SEÇÃO 20: EQUIPAMENTOS DO CLIENTE
        # ============================================================================
        c.showPage()
        if y_position < 300:
            y_position = height - 50

        # Verificar se há equipamentos para este cliente
        if not EquipamentosGeraisFiltrado.empty:
            # Preparar dados dos equipamentos
            equipamentos_data = [['Equipamentos Vinculados ao Cliente']]
            equipamentos_data.append(['Equipamento', 'Descrição', 'Nº Interno', 'Série'])
            
            # Adicionar dados dos equipamentos
            for _, equip in EquipamentosGeraisFiltrado.iterrows():
                equipamentos_data.append([
                    str(equip.get('EQUIPAMENTO', '')),
                    str(equip.get('DESCRIÇÃO EQUIPAMENTO', '')),
                    str(equip.get('Nº INTERNO', '')),
                    str(equip.get('SÉRIE', ''))
                ])
            
            # Criar tabela de equipamentos
            tabela_equipamentos = Table(equipamentos_data, colWidths=[100, 200, 100, 100])
            tabela_equipamentos.setStyle(StyleEquipamentos)
            
            altura_tabela_equipamentos = tabela_equipamentos.wrap(width, height)[1]
            if y_position - altura_tabela_equipamentos < 50:

                y_position = height - 50
            
            tabela_equipamentos.drawOn(c, 10, y_position - altura_tabela_equipamentos)
            y_position -= altura_tabela_equipamentos + 30
        
        # ============================================================================
        # SEÇÃO 21: COLIGADOS DO CLIENTE (APENAS SE EXISTIREM)
        # ============================================================================
        c.showPage()
        

        # Verificar se há coligados para este cliente
        if not ColigadosFiltrado.empty:
            # Preparar dados dos coligados
            coligados_data = [['Coligados Vinculados ao Cliente']]
            coligados_data.append(['Código SAP', 'Razão Social', 'SAP Principal'])
            
            # Adicionar dados dos coligados
            for _, coligado in ColigadosFiltrado.iterrows():
                coligados_data.append([
                    str(coligado.get('CÓDIGO SAP', '')),
                    str(coligado.get('RAZÃO SOCIAL', '')),
                    str(coligado.get('SAP PRINCIPAL', ''))
                ])
            
            # Criar tabela de coligados
            tabela_coligados = Table(coligados_data, colWidths=[100, 300, 100])
            tabela_coligados.setStyle(StyleColigados)
            
            altura_tabela_coligados = tabela_coligados.wrap(width, height)[1]
            if y_position - altura_tabela_coligados < 50:
                c.showPage()
                y_position = height - 50
            
            tabela_coligados.drawOn(c, 10, y_position - altura_tabela_coligados)
            y_position -= altura_tabela_coligados + 30
        c.showPage()
        # ============================================================================
        # SEÇÃO 22: PRODUTOS CONSUMIDOS COBRANÇA ANUAL
        # ============================================================================
        
        
        # Calcular valores para o ano atual
        ano_atual_num = min(ano_atual, max_anos)  
        valor_total_ano_atual = valor_consumido.get(f'ano_{ano_atual_num}', 0)
        target_ano_atual = target_unificado.get(f'ano_{ano_atual_num}', 0)
        percentual_atual = percentuais.get(f'ano_{ano_atual_num}', 0)
        
        diferenca = valor_total_ano_atual - target_ano_atual
        # Calcular multa
        multa = 0
        if percentual_atual < 100:  # só aplica se não bateu 100%
            if 91 < percentual_atual <= 99:
                multa = target_ano_atual * 0.06
            elif 81 < percentual_atual <= 90:
                multa = target_ano_atual * 0.14
            elif 71 < percentual_atual <= 80:
                multa = target_ano_atual * 0.21
            elif 61 < percentual_atual <= 70:
                multa = target_ano_atual * 0.28
            elif 51 < percentual_atual <= 60:
                multa = target_ano_atual * 0.34
            elif 41 < percentual_atual <= 50:
                multa = target_ano_atual * 0.40
            elif 31 < percentual_atual <= 40:
                multa = target_ano_atual * 0.45
            elif 21 < percentual_atual <= 30:
                multa = target_ano_atual * 0.48
            elif 0 <= percentual_atual <= 20:
                multa = target_ano_atual * 0.50
        
        produtos_data = [['Produtos Consumidos',' Consumo realizado até o Fechamento','Cobrança Anual']]
        produtos_data.append(['LENTES', 'DESCRIÇÃO', 'VALOR TOTAL', 'TARGET UNIFICADO', 'DIFERENÇA', 'CÁLCULO DE MULTA'])

        # Buscar valores específicos por SKU para o ano atual
        inicio_periodo_atual = DataDaApuração + relativedelta(months=(ano_atual_num-1)*12)
        fim_periodo_atual = DataDaApuração + relativedelta(months=ano_atual_num*12-1)

        for i, sku in enumerate(skus):
            descricao = descricoes[i] if i < len(descricoes) else ""

            # Buscar consumo específico do SKU
            historico_sku = BaseHistorica[
                (
                    (BaseHistorica['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistorica['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistorica['Item 2'] == sku) &
                (BaseHistorica['DataApuração'] >= inicio_periodo_atual.strftime('%Y-%m')) &
                (BaseHistorica['DataApuração'] <= fim_periodo_atual.strftime('%Y-%m'))
            ]

            valor_sku = historico_sku['Total Gross'].sum() if not historico_sku.empty else 0

            # Para a primeira linha de produto, adicionar os valores centralizados
            if i == 0:
                produtos_data.append([
                    sku,
                    descricao,
                    formatar_moeda(valor_sku),
                    formatar_moeda(target_ano_atual),  # TARGET na primeira linha
                    formatar_moeda(diferenca),         # DIFERENÇA na primeira linha
                    formatar_moeda(multa)              # MULTA na primeira linha
                ])
            else:
                # Para as demais linhas, deixar as últimas 3 colunas vazias (serão mescladas)
                produtos_data.append([
                    sku,
                    descricao,
                    formatar_moeda(valor_sku),
                    "",  # Vazio para mesclar
                    "",  # Vazio para mesclar
                    ""   # Vazio para mesclar
                ])

        # Ajuste colWidths: removendo a coluna "CONSUMO"
        styleConsumo = TableStyle([
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
            ('TOPPADDING', (0, 0), (-1, -1), 1),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            
            # Mescla as colunas 0 e 1 na primeira linha
            ('SPAN', (0, 0), (1, 0)),
            
            # Mescla as colunas 2, 3, 4 e 5 na primeira linha
            ('SPAN', (2, 0), (5, 0)),
            
            # IMPORTANTE: Mesclar as colunas TARGET, DIFERENÇA e MULTA verticalmente
            # da linha 2 (primeira linha de produtos) até a última linha
            ('SPAN', (3, 2), (3, -1)),  # TARGET UNIFICADO mesclado verticalmente
            ('SPAN', (4, 2), (4, -1)),  # DIFERENÇA mesclada verticalmente
            ('SPAN', (5, 2), (5, -1)),  # CÁLCULO DE MULTA mesclado verticalmente
            
            # Formatação da primeira linha (cabeçalho principal)
            ('BACKGROUND', (0, 0), (1, 0), (68/255, 83/255, 106/255)),  # Produtos Consumidos
            ('BACKGROUND', (2, 0), (5, 0), (132/255, 150/255, 175/255)),  # Cobrança Anual
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTSIZE', (0, 0), (-1, 0), 6),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            
            # Formatação da segunda linha (sub-cabeçalhos)
            ('BACKGROUND', (0, 1), (-1, 1), colors.lightgrey),
            ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 1), (-1, 1), 5),
            
            # Aumentar fonte das colunas mescladas (valores centralizados)
            ('FONTSIZE', (3, 2), (3, -1), 9),  # TARGET
            ('FONTSIZE', (4, 2), (4, -1), 9),  # DIFERENÇA  
            ('FONTSIZE', (5, 2), (5, -1), 9),  # MULTA
            ('FONTNAME', (3, 2), (5, -1), 'Helvetica-Bold'),
            
            # Cores especiais para diferença e multa
            ('TEXTCOLOR', (4, 2), (4, -1), colors.red if diferenca < 0 else colors.green),
            ('TEXTCOLOR', (5, 2), (5, -1), colors.red),
            
            # Bordas
            ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ])
        tabela_produtos = Table(produtos_data, colWidths=[50, 100, 80, 120, 120, 120])
        tabela_produtos.setStyle(styleConsumo)

        altura_tabela_produtos = tabela_produtos.wrap(width, height)[1]
        if y_position - altura_tabela_produtos < 50:
            c.showPage()
            y_position = height - 50

        tabela_produtos.drawOn(c, 10, y_position - altura_tabela_produtos)
        y_position -= altura_tabela_produtos + 30

        # ============================================================================
        # SEÇÃO 23: EXTRATO DE CONSUMO - VISÃO GERAL
        # ============================================================================
        if y_position < 300:
            c.showPage()
            y_position = height - 50

     

        # Buscar dados do histórico
        historico_extrato = BaseHistorica[
            (
                (BaseHistorica['Codigo_PN'] == sap_principal_filtro) |
                (BaseHistorica['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
            ) &
            (BaseHistorica['Item 2'].isin(lentesFiltroHistorico))
        ].copy()

        # Agregar por Codigo_PN, RAZÃO SOCIAL, Item 2, Mês e Ano somando Quantidade e Total Gross
        historico_agrupado = historico_extrato.groupby(
            ['Codigo_PN', 'Nome_PN', 'Item 2', 'Mês', 'Ano'],
            as_index=False
        ).agg({
            'Descricao_Item': 'first',
            'Quantidade': 'sum',
            'Total Gross': 'sum'
        })

        # Ordenar por ano e mês (mais recente primeiro)
        historico_agrupado = historico_agrupado.sort_values(['Ano', 'Mês'], ascending=[False, False])

        # Montar apenas os dados
        extrato_dados = []
        for _, registro in historico_agrupado.iterrows():
            quantidade = int(registro.get('Quantidade', 0)) if pd.notna(registro.get('Quantidade', 0)) else 0
            valor_total = formatar_moeda(registro.get('Total Gross', 0))
            razao_social = str(registro['Nome_PN'])

            extrato_dados.append([
                str(registro['Codigo_PN']),
                razao_social,
                str(registro.get('Item 2', '')),
                str(quantidade),
                valor_total,
                str(registro.get('Mês', '')),
                str(registro.get('Ano', ''))
            ])

        # =====================================================================
        # DESENHAR A TABELA EM PARTES (para caber nas páginas)
        # =====================================================================
        if extrato_dados:  # Se há dados
            
            # Cabeçalho fixo da tabela
            cabecalho = ['SAP Principal', 'Razão Social', 'SKU', 'Quantidade', 'Valor', 'Mês', 'Ano']
            linha_titulo = ['Extrato de Consumo - Visão Geral'] + [''] * (len(cabecalho) - 1)
            max_linhas_por_pagina = 100  # ajusta conforme necessário

            
            for i in range(0, len(extrato_dados), max_linhas_por_pagina):
                bloco = extrato_dados[i:i + max_linhas_por_pagina]

                # 2. JUNTE AS TRÊS PARTES: TÍTULO, CABEÇALHO DE COLUNAS E DADOS
                dados_completos = [linha_titulo, cabecalho] + bloco

                # 3. CRIE A TABELA USANDO A NOVA ESTRUTURA DE DADOS
                tabela_extrato = Table(dados_completos, colWidths=[50, 260, 40, 50, 100, 50, 40])
                
                # Aplique seu estilo, que agora funcionará perfeitamente
                tabela_extrato.setStyle(StyleBaseHistorica)

                # ... o resto do seu código para desenhar a tabela ...
                altura_tabela_extrato = tabela_extrato.wrap(width, height)[1]

                if y_position - altura_tabela_extrato < 50:
                    c.showPage()
                    y_position = height - 50

                tabela_extrato.drawOn(c, 10, y_position - altura_tabela_extrato)
                y_position -= altura_tabela_extrato + 30

        # ============================================================================
        # SEÇÃO 24: FINALIZAR PDF E UPLOAD
        # ============================================================================
        c.save()
        print(f"PDF gerado: {nome_arquivo}")

        # Upload para Box
        try:
            if access_token:
                headers = {'Authorization': f'Bearer {access_token}'}
                
                files = {
                    'attributes': (None, f'{{"name": "{nome_arquivo}", "parent": {{"id": "{FOLDER_ID}"}}}}', 'application/json'),
                    'file': (nome_arquivo, open(nome_arquivo, 'rb'), 'application/pdf')
                }
                
                response = requests.post(UPLOAD_URL, headers=headers, files=files)
                files['file'][1].close()
                
                if response.status_code in [200, 201]:
                    print(f"✅ Arquivo {nome_arquivo} enviado com sucesso para o Box!")
                else:
                    print(f"❌ Erro ao enviar {nome_arquivo}: {response.text}")
                    
                    # Tentar refresh do token
                    try:
                        refresh_access_token()
                        headers = {'Authorization': f'Bearer {access_token_global}'}
                        
                        files = {
                            'attributes': (None, f'{{"name": "{nome_arquivo}", "parent": {{"id": "{FOLDER_ID}"}}}}', 'application/json'),
                            'file': (nome_arquivo, open(nome_arquivo, 'rb'), 'application/pdf')
                        }
                        
                        response = requests.post(UPLOAD_URL, headers=headers, files=files)
                        files['file'][1].close()
                        
                        if response.status_code in [200, 201]:
                            print(f"✅ Arquivo {nome_arquivo} enviado com sucesso após refresh do token!")
                        else:
                            print(f"❌ Erro persistente ao enviar {nome_arquivo}: {response.text}")
                            
                    except Exception as refresh_error:
                        print(f"❌ Erro ao fazer refresh do token: {refresh_error}")
                        
        except Exception as upload_error:
            print(f"❌ Erro no upload: {upload_error}")

    print("Processamento concluído!")
    messagebox.showinfo("Concluído", f"Foram processados {contador} clientes com sucesso!")

                    
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
        
        popup.destroy()
        processar_arquivos()

    tk.Button(frame, text="Confirmar", command=validar_access_key).grid(row=12, column=0, pady=(10, 10))


# ============================================================================
# INTERFACE GRÁFICA
# ============================================================================
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
