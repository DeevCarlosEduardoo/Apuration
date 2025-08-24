
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
    """
    Função refatorada para processar arquivos e gerar relatórios em PDF.
    Mantém o formato original do PDF, mas com suporte a múltiplas modalidades
    e extensão para 10 anos.
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

    # Configuração de estilo das tabelas
    style = TableStyle([
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
        ('TOPPADDING', (0, 0), (-1, -1), 1),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
    ])

    StyleTituloMudado = TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 5),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
        ('TOPPADDING', (0, 0), (-1, -1), 1),
        ('SPAN', (0, 0), (-1, 0)),
        ('BACKGROUND', (0, 0), (-1, 0), (68/255, 83/255, 106/255)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),
        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),
        ('FONTSIZE', (0, 0), (-1, 0), 6),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('TEXTCOLOR', (0, 1), (0, -1), colors.black),
        ('FONTNAME', (0, 1), (0, -1), 'Helvetica-Bold'),
        ('BACKGROUND', (1, 1), (1, -1), colors.white),
        ('INNERGRID', (0, 1), (-1, -1), 0, colors.white),
        ('BOX', (0, 1), (-1, -1), 0, colors.white),
    ])

    StyleColigados = TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 5),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
        ('TOPPADDING', (0, 0), (-1, -1), 1),
        ('SPAN', (0, 0), (-1, 0)),
        ('BACKGROUND', (0, 0), (-1, 0), (68/255, 83/255, 106/255)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),
        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),
        ('FONTSIZE', (0, 0), (-1, 0), 6),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
        ('TEXTCOLOR', (0, 2), (0, -1), colors.black),
        ('FONTNAME', (0, 2), (-1, -1), 'Helvetica'),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.lightgrey),
    ])

    StyleBaseHistorica = TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 5),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
        ('TOPPADDING', (0, 0), (-1, -1), 1),
        ('SPAN', (0, 0), (-1, 0)),
        ('BACKGROUND', (0, 0), (-1, 0), (68/255, 83/255, 106/255)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),
        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),
        ('FONTSIZE', (0, 0), (-1, 0), 6),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
        ('TEXTCOLOR', (0, 2), (0, -1), colors.black),
        ('FONTNAME', (0, 2), (-1, -1), 'Helvetica'),
        ('BACKGROUND', (1, 1), (1, -1), colors.white),
        ('INNERGRID', (0, 1), (-1, -1), 0, colors.white),
        ('BOX', (0, 1), (-1, -1), 0, colors.white),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.lightgrey),
    ])

    def calcular_altura_tabela(num_linhas):
        altura_linha = 15
        return num_linhas * altura_linha

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
        return ""

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
        # Buscar TODAS as modalidades do cliente no Base original (não só no concatenado)
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
        # Filtragem de lentes para o cliente
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
        # Encontrar o contrato mais antigo entre todas as modalidades
        data_inicio_mais_antiga = None
        contrato_referencia = None
        
        for modalidade, contratos in modalidades_dict.items():
            for contrato in contratos:
                data_inicio_contrato = data_base_excel + timedelta(contrato['INICIO DA APURAÇÃO'])
                if data_inicio_mais_antiga is None or data_inicio_contrato < data_inicio_mais_antiga:
                    data_inicio_mais_antiga = data_inicio_contrato
                    contrato_referencia = contrato

        # Usar o contrato mais antigo como referência para datas
        if contrato_referencia is not None:
            dados_referencia = contrato_referencia
        else:
            dados_referencia = row

        # ============================================================================
        # SEÇÃO 10: PROCESSAMENTO DE DATAS E APURAÇÃO
        # ============================================================================
        # Processamento das datas de apuração baseado no contrato mais antigo
        DataDaApuração = data_base_excel + timedelta(dados_referencia.get('INICIO DA APURAÇÃO'))
        DataFimApuração = DataDaApuração + relativedelta(months=int(dados_referencia.get('PRAZO APURACAO', 0)))
        
        mes_extenso = meses_portugues[DataDaApuração.month]
        anodeApuracaosemfromatar = DataDaApuração.year
        
        DataDaApuraçãoFormatada = f"{mes_extenso} de {anodeApuracaosemfromatar}"
        DataDaApuraçãoFiltro = DataDaApuração.strftime('%Y-%m')
        DataFimApuraçãoFormatada = DataFimApuração.strftime('%d/%m/%Y')

        dif_anos = DataFimApuração.year - data_atual.year
        dif_meses = DataFimApuração.month - data_atual.month
        total_meses = (dif_anos * 12) + dif_meses - 1

        DataInicioApuraçãoFormatada = DataDaApuração.strftime('%d/%m/%Y')
        DataFimApuraçãoFormatada = DataFimApuração.strftime('%d/%m/%Y')

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

        InicioApuração = DataDaApuração
        AnodaApuração = calcular_ano_referencia(DataInicioFormatada)

        # ============================================================================
        # SEÇÃO 11: CONSUMO UNIFICADO (SOMAR TODAS AS MODALIDADES - 10 ANOS)
        # ============================================================================
        # Inicializar arrays para consumo unificado (10 anos)
        consumo_unificado = {}
        for ano in range(1, 11):  # 10 anos
            consumo_unificado[f'ano_{ano}'] = 0

        # Somar consumo de todas as modalidades
        for modalidade, contratos in modalidades_dict.items():
            for contrato in contratos:
                for ano in range(1, 11):  # 10 anos
                    coluna_consumo = f'CONSUMO ANO {ano}'
                    valor_consumo = pd.to_numeric(contrato.get(coluna_consumo), errors='coerce')
                    if pd.notna(valor_consumo):
                        consumo_unificado[f'ano_{ano}'] += valor_consumo

        # Formatar valores de consumo
        consumo_formatado = {}
        for ano in range(1, 11):  # 10 anos
            valor = consumo_unificado[f'ano_{ano}']
            consumo_formatado[f'ano_{ano}'] = formatar_moeda(valor) if valor > 0 else "R$ 0,00"

        # ============================================================================
        # SEÇÃO 12: PROCESSAMENTO DE HISTÓRICO PARA ATÉ 10 ANOS
        # ============================================================================
        BaseHistorica['DataApuração'] = pd.to_datetime(BaseHistorica['Ano'].astype(str) + '-' + BaseHistorica['Mês'].astype(str).str.zfill(2))
        BaseHistoricaCompleta['DataApuração'] = pd.to_datetime(BaseHistoricaCompleta['Ano'].astype(str) + '-' + BaseHistoricaCompleta['Mês'].astype(str).str.zfill(2))

        # Processamento para cada ano (1 a 10) baseado no contrato mais antigo
        BaseHistoricaFiltradas = {}
        
        for ano in range(1, 11):  # Agora vai até ano 10
            inicio_periodo = DataDaApuração + relativedelta(months=(ano-1)*12)
            fim_periodo = DataDaApuração + relativedelta(months=ano*12-1)
            
            BaseHistoricaFiltradas[f'ano_{ano}'] = BaseHistorica[
                (
                    (BaseHistorica['Codigo_PN'] == sap_principal_filtro) |
                    (BaseHistorica['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
                ) &
                (BaseHistorica['Item 2'].isin(lentesFiltroHistorico)) &
                (BaseHistorica['DataApuração'] >= inicio_periodo.strftime('%Y-%m')) &
                (BaseHistorica['DataApuração'] <= fim_periodo.strftime('%Y-%m'))
            ]

        BaseHistoricaFiltradaCompleta = BaseHistoricaCompleta[
            (
                (BaseHistoricaCompleta['Codigo_PN'] == sap_principal_filtro) |
                (BaseHistoricaCompleta['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
            ) &
            (BaseHistoricaCompleta['Item 2'].isin(lentesFiltroHistorico)) &
            (BaseHistoricaCompleta['DataApuração'] >= DataDaApuração)
        ]

        # Cálculo dos meses passados baseado no ano atual
        ano_atual = int(AnodaApuração.split()[1])
        data_inicio_ano_atual = DataDaApuração + relativedelta(months=(ano_atual-1)*12)
        meses_apurados = (data_atual.year - data_inicio_ano_atual.year) * 12 + (data_atual.month - data_inicio_ano_atual.month)

        if data_atual.day >= data_inicio_ano_atual.day:
            meses_apurados += 1

        meses_passados = (meses_apurados % 12) or 12

        # ============================================================================
        # SEÇÃO 13: PROCESSAMENTO DE LENTES E VALORES HISTÓRICOS
        # ============================================================================
        # Obter dados das lentes
        skus = FiltrandoLentes['SKU PRODUTO'].dropna().astype(str).tolist()
        descricoes = FiltrandoLentes['DESCRIÇÃO CONSUMO'].dropna().astype(str).tolist()

        lentes_dados = [f"{sku} {desc}" for sku, desc in zip(skus, descricoes)]

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

        # Criando DataFrames para cada ano
        dados_lentes_anos = {}
        for ano in range(1, 11):
            dados_lentes_anos[f'ano_{ano}'] = pd.DataFrame({
                "LENTES": lentes,
                "DESCRIÇÃO CONSUMO": descricao
            })

        # Função para obter valor somado por ano
        def obter_valor_somado_por_ano(lente, ano):
            if f'ano_{ano}' in BaseHistoricaFiltradas:
                valores_correspondentes = BaseHistoricaFiltradas[f'ano_{ano}'][BaseHistoricaFiltradas[f'ano_{ano}']['Item 2'] == lente]['Total Gross']
                return valores_correspondentes.sum() if not valores_correspondentes.empty else 0
            return 0

        # Calcular valores para cada ano
        valores_totais_anos = {}
        for ano in range(1, 11):
            if f'ano_{ano}' in dados_lentes_anos:
                dados_lentes_anos[f'ano_{ano}']['VALOR TOTAL'] = dados_lentes_anos[f'ano_{ano}']['LENTES'].apply(lambda x: obter_valor_somado_por_ano(x, ano))
                dados_lentes_anos[f'ano_{ano}'] = dados_lentes_anos[f'ano_{ano}'].drop_duplicates(subset=['LENTES'])
                valores_totais_anos[f'ano_{ano}'] = {
                    'valor': dados_lentes_anos[f'ano_{ano}']['VALOR TOTAL'].sum(),
                    'formatado': formatar_moeda(dados_lentes_anos[f'ano_{ano}']['VALOR TOTAL'].sum())
                }

        # ============================================================================
        # SEÇÃO 14: PROCESSAMENTO DE EQUIPAMENTOS
        # ============================================================================
        EquipamentosGeraisFiltrado = EquipamentosGerais[
            (EquipamentosGerais['SAP PRINCIPAL'] == sap_principal_filtro)
        ][['EQUIPAMENTO', 'DESCRIÇÃO EQUIPAMENTO', 'Nº INTERNO', 'SÉRIE']].drop_duplicates(
            subset=['EQUIPAMENTO', 'DESCRIÇÃO EQUIPAMENTO', 'Nº INTERNO', 'SÉRIE']
        )

        equipamentos_longo_com_cabecalho = [['SKU Equipamento', 'Descrição', 'N INTERNO', 'Série']] + EquipamentosGeraisFiltrado.iloc[::-1].values.tolist()
        tabela_equipamentos_longo = Table(equipamentos_longo_com_cabecalho, colWidths=[100, 250])
        tabela_equipamentos_longo.setStyle(StyleColigados)

        # ============================================================================
        # SEÇÃO 15: PREPARAÇÃO DOS DADOS DO CLIENTE
        # ============================================================================
        RazaoSocialCompleta = f"{row['SAP PRINCIPAL']} - {row['RAZÃO SOCIAL']}"
        
        if pd.isnull(row['SAM']) or row['SAM'] == '':
            InfClientes = [['Informações do Cliente'],
                          ['Sap Principal', RazaoSocialCompleta], 
                          ['Consultor', row['CONSULTOR']], 
                          ['Distrital', row['DISTRITAL']], 
                          ['Sam', '']]
        else:
            InfClientes = [['Informações do Cliente'],
                          ['Sap Principal', RazaoSocialCompleta], 
                          ['Consultor', row['CONSULTOR']], 
                          ['Distrital', row['DISTRITAL']], 
                          ['Sam', row['SAM']]]

        # ============================================================================
        # SEÇÃO 16: GERAÇÃO DO PDF NO FORMATO ORIGINAL
        # ============================================================================
        # Determinar se é PDF unificado ou individual
        tem_multiplas_modalidades = len(modalidades_dict) > 1
        
        if tem_multiplas_modalidades:
            nome_arquivo = f"Relatório_Unificado_{sap_principal_filtro}_{AnodaApuração}.pdf"
            print(f"Gerando PDF unificado para cliente {sap_principal_filtro} com modalidades: {', '.join(modalidades_dict.keys())}")
        else:
            modalidade_unica = list(modalidades_dict.keys())[0]
            nome_arquivo = f"Relatório_{modalidade_unica.replace(' ', '_')}_{sap_principal_filtro}_{AnodaApuração}.pdf"
            print(f"Gerando PDF individual para cliente {sap_principal_filtro} - modalidade: {modalidade_unica}")
        
        # Criar PDF com ReportLab
        c = pdf_canvas.Canvas(nome_arquivo, pagesize=letter)
        width, height = letter
        
        # Cabeçalho do relatório
        c.setFont("Helvetica-Bold", 16)
        c.drawString(50, height - 50, f"Relatório Apuração - {DataDaApuraçãoFormatada}")
        
        y_position = height - 80
        
        # ============================================================================
        # SEÇÃO 17: INFORMAÇÕES DA APURAÇÃO
        # ============================================================================
        c.setFont("Helvetica-Bold", 12)
        c.drawString(50, y_position, "Informações da Apuração")
        y_position -= 20
        
        c.setFont("Helvetica", 10)
        c.drawString(50, y_position, f"Data Inicio {DataInicioApuraçãoFormatada}")
        y_position -= 15
        c.drawString(50, y_position, f"Data Fim {DataFimApuraçãoFormatada}")
        y_position -= 15
        c.drawString(50, y_position, f"Meses faltantes - Contrato {total_meses} Meses")
        y_position -= 15
        c.drawString(50, y_position, f"Meses Apurados - Ano Corrente {meses_passados}")
        y_position -= 30
        
        # ============================================================================
        # SEÇÃO 18: INFORMAÇÕES DO CONTRATO (MÚLTIPLAS MODALIDADES)
        # ============================================================================
        c.setFont("Helvetica-Bold", 12)
        c.drawString(50, y_position, "Informações do Contrato")
        y_position -= 20
        
        # Para cada modalidade, criar uma seção
        for modalidade, contratos in modalidades_dict.items():
            contrato_principal = contratos[0]  # Pegar o primeiro contrato da modalidade
            
            c.setFont("Helvetica", 10)
            c.drawString(50, y_position, f"Modalidade {modalidade}")
            y_position -= 15
            c.drawString(50, y_position, f"Nº Contrato {contrato_principal.get('Nº CONTRATO', 'N/A')}")
            y_position -= 15
            c.drawString(50, y_position, f"Versão Contratual {contrato_principal.get('VERSÃO', 'N/A')}")
            y_position -= 15
            
            # Calcular vigência específica desta modalidade
            try:
                if isinstance(contrato_principal['DT. INÍCIO'], (int, float)) and contrato_principal['DT. INÍCIO'] > 60:
                    data_inicio_modal = data_base_excel + timedelta(contrato_principal['DT. INÍCIO'] - pd.Timedelta(days=2))
                else:
                    data_inicio_modal = pd.to_datetime(contrato_principal['DT. INÍCIO'])
                    
                if isinstance(contrato_principal['DT. FINAL'], (int, float)) and contrato_principal['DT. FINAL'] > 60:
                    data_fim_modal = data_base_excel + timedelta(contrato_principal['DT. FINAL'] - pd.Timedelta(days=2))
                else:
                    data_fim_modal = pd.to_datetime(contrato_principal['DT. FINAL'])
                    
            except (TypeError, ValueError):
                serial_inicio = int(contrato_principal['DT. INÍCIO'])
                serial_fim = int(contrato_principal['DT. FINAL'])
                data_inicio_modal = datetime(1900, 1, 1) + timedelta(days=serial_inicio - 2)
                data_fim_modal = datetime(1900, 1, 1) + timedelta(days=serial_fim - 2)
            
            vigencia_modal = f"{data_inicio_modal.strftime('%d/%m/%Y')} - {data_fim_modal.strftime('%d/%m/%Y')}"
            c.drawString(50, y_position, f"Vigência Contratual {vigencia_modal}")
            y_position -= 15
            
            # Início da apuração específica desta modalidade
            data_apuracao_modal = data_base_excel + timedelta(contrato_principal.get('INICIO DA APURAÇÃO'))
            mes_modal = meses_portugues[data_apuracao_modal.month]
            ano_modal = data_apuracao_modal.year
            c.drawString(50, y_position, f"Inicio da Apuração {mes_modal} de {ano_modal}")
            y_position -= 25
            
            # Se há múltiplas modalidades, adicionar separador
            if len(modalidades_dict) > 1:
                c.line(50, y_position, width - 50, y_position)
                y_position -= 15
        
        # ============================================================================
        # SEÇÃO 19: INFORMAÇÕES DO CLIENTE
        # ============================================================================
        if y_position < 200:
            c.showPage()
            y_position = height - 50
            
        tabela_info_cliente = Table(InfClientes, colWidths=[120, 300])
        tabela_info_cliente.setStyle(StyleTituloMudado)
        tabela_info_cliente.wrapOn(c, width, height)
        tabela_info_cliente.drawOn(c, 50, y_position - tabela_info_cliente.wrap(width, height)[1])
        y_position -= tabela_info_cliente.wrap(width, height)[1] + 30
        
        # ============================================================================
        # SEÇÃO 20: CONSUMO UNIFICADO (10 ANOS)
        # ============================================================================
        if y_position < 300:
            c.showPage()
            y_position = height - 50
            
        # Criar tabela de consumo unificado (10 anos)
        consumo_data = [['Consumo Unificado']]
        consumo_headers = []
        consumo_meta = []
        consumo_target = []
        consumo_valores = []
        consumo_percentual = []
        
        # Cabeçalhos e dados para 10 anos
        for ano in range(1, 11):
            consumo_headers.append(f'Ano {ano}')
            consumo_meta.append('100%')  # Meta sempre 100%
            
            # Target - usar do primeiro contrato encontrado ou calcular
            if f'TARGET ANO {ano}' in contrato_referencia:
                target_valor = pd.to_numeric(contrato_referencia.get(f'TARGET ANO {ano}'), errors='coerce')
                if pd.notna(target_valor):
                    consumo_target.append(formatar_moeda(target_valor))
                else:
                    consumo_target.append('R$ 0,00')
            else:
                consumo_target.append('R$ 0,00')
            
            # Valor consumido unificado
            consumo_valores.append(consumo_formatado[f'ano_{ano}'])
            
            # Calcular percentual de atingimento
            target_num = pd.to_numeric(contrato_referencia.get(f'TARGET ANO {ano}'), errors='coerce') if f'TARGET ANO {ano}' in contrato_referencia else 0
            consumo_num = consumo_unificado[f'ano_{ano}']
            
            if target_num > 0 and consumo_num > 0:
                percentual = (consumo_num / target_num) * 100
                consumo_percentual.append(f'{percentual:.2f}%')
            else:
                consumo_percentual.append('0.00%')
        
        # Construir tabela
        consumo_data.append(consumo_headers)
        consumo_data.append(['Meta %'] + consumo_meta)
        consumo_data.append(['Target - Unificado'] + consumo_target)
        consumo_data.append(['Valor Consumido - Unificado'] + consumo_valores)
        consumo_data.append(['Percentual de Atingimento'] + consumo_percentual)
        
        # Ajustar largura das colunas para 10 anos
        col_width = (width - 200) / 11  # 11 colunas (1 label + 10 anos)
        col_widths = [120] + [col_width] * 10
        
        tabela_consumo = Table(consumo_data, colWidths=col_widths)
        tabela_consumo.setStyle(StyleTituloMudado)
        tabela_consumo.wrapOn(c, width, height)
        
        if y_position - tabela_consumo.wrap(width, height)[1] < 50:
            c.showPage()
            y_position = height - 50
        
        tabela_consumo.drawOn(c, 50, y_position - tabela_consumo.wrap(width, height)[1])
        y_position -= tabela_consumo.wrap(width, height)[1] + 30
        
        # ============================================================================
        # SEÇÃO 21: COLIGADOS
        # ============================================================================
        if not ColigadosFiltrado.empty:
            if y_position < 200:
                c.showPage()
                y_position = height - 50
            
            coligados_data = [['Coligados']]
            coligados_data.append(['Sap Coligado', 'Razão Social'])
            
            for _, coligado in ColigadosFiltrado.iterrows():
                coligados_data.append([str(coligado['CÓDIGO SAP']), coligado['RAZÃO SOCIAL']])
            
            tabela_coligados = Table(coligados_data, colWidths=[100, 300])
            tabela_coligados.setStyle(StyleColigados)
            tabela_coligados.wrapOn(c, width, height)
            
            if y_position - tabela_coligados.wrap(width, height)[1] < 50:
                c.showPage()
                y_position = height - 50
            
            tabela_coligados.drawOn(c, 50, y_position - tabela_coligados.wrap(width, height)[1])
            y_position -= tabela_coligados.wrap(width, height)[1] + 30
        
        # ============================================================================
        # SEÇÃO 22: PRODUTOS CONSUMIDOS COBRANÇA ANUAL
        # ============================================================================
        if y_position < 400:
            c.showPage()
            y_position = height - 50
        
        # Usar dados do ano atual da apuração
        ano_atual_num = int(AnodaApuração.split()[1])
        dados_lentes_ano_atual = dados_lentes_anos.get(f'ano_{ano_atual_num}', dados_lentes_anos['ano_1'])
        valor_total_ano_atual = valores_totais_anos.get(f'ano_{ano_atual_num}', valores_totais_anos['ano_1'])
        
        # Target unificado do ano atual
        target_ano_atual = 0
        for modalidade, contratos in modalidades_dict.items():
            for contrato in contratos:
                target_valor = pd.to_numeric(contrato.get(f'CONSUMO ANO {ano_atual_num}'), errors='coerce')
                if pd.notna(target_valor):
                    target_ano_atual += target_valor
        
        target_formatado = formatar_moeda(target_ano_atual)
        diferenca = valor_total_ano_atual['valor'] - target_ano_atual
        diferenca_formatada = formatar_moeda(diferenca)
        
        # Cálculo da multa (exemplo simples)
        multa = abs(diferenca) * 0.1 if diferenca < 0 else 0
        multa_formatada = formatar_moeda(multa)
        
        produtos_data = [['Produtos Consumidos Cobrança Anual']]
        produtos_data.append(['LENTES', 'DESCRIÇÃO CONSUMO', 'VALOR TOTAL', 'TARGET UNIFICADO', 'DIFERENÇA', 'CÁLCULO DE MULTA'])
        
        for _, produto in dados_lentes_ano_atual.iterrows():
            valor_produto = formatar_moeda(produto['VALOR TOTAL'])
            produtos_data.append([
                produto['LENTES'], 
                produto['DESCRIÇÃO CONSUMO'], 
                valor_produto,
                '', '', ''  # Valores específicos por produto podem ser adicionados aqui
            ])
        
        # Linha de totais
        produtos_data.append(['', '', valor_total_ano_atual['formatado'], target_formatado, diferenca_formatada, multa_formatada])
        
        tabela_produtos = Table(produtos_data, colWidths=[60, 200, 80, 80, 80, 80])
        tabela_produtos.setStyle(StyleBaseHistorica)
        tabela_produtos.wrapOn(c, width, height)
        
        if y_position - tabela_produtos.wrap(width, height)[1] < 50:
            c.showPage()
            y_position = height - 50
        
        tabela_produtos.drawOn(c, 50, y_position - tabela_produtos.wrap(width, height)[1])
        y_position -= tabela_produtos.wrap(width, height)[1] + 30
        
        # ============================================================================
        # SEÇÃO 23: EQUIPAMENTOS
        # ============================================================================
        if not EquipamentosGeraisFiltrado.empty:
            if y_position < 200:
                c.showPage()
                y_position = height - 50
            
            c.setFont("Helvetica-Bold", 12)
            c.drawString(50, y_position, "Equipamentos")
            y_position -= 20
            
            tabela_equipamentos_longo.wrapOn(c, width, height)
            if y_position - tabela_equipamentos_longo.wrap(width, height)[1] < 50:
                c.showPage()
                y_position = height - 50
            
            tabela_equipamentos_longo.drawOn(c, 50, y_position - tabela_equipamentos_longo.wrap(width, height)[1])
            y_position -= tabela_equipamentos_longo.wrap(width, height)[1] + 30
        
        # ============================================================================
        # SEÇÃO 24: EXTRATO DE CONSUMO - VISÃO GERAL
        # ============================================================================
        if y_position < 300:
            c.showPage()
            y_position = height - 50
        
        c.setFont("Helvetica-Bold", 12)
        c.drawString(50, y_position, "Extrato de Consumo - Visão Geral")
        y_position -= 30
        
        # Buscar dados do histórico para o extrato
        extrato_data = [['SAP Principal', 'Razão Social', 'SKU Produto', 'Descrição Produto', 'Quantidade', 'Valor Total', 'Mês', 'Ano']]
        
        # Obter dados do BaseHistoricaCompleta
        historico_detalhado = BaseHistoricaCompleta[
            (
                (BaseHistoricaCompleta['Codigo_PN'] == sap_principal_filtro) |
                (BaseHistoricaCompleta['Codigo_PN'].isin(ColigadosFiltrado['CÓDIGO SAP']))
            ) &
            (BaseHistoricaCompleta['Item 2'].isin(lentesFiltroHistorico))
        ].copy()
        
        # Ordenar por ano e mês (mais recente primeiro)
        historico_detalhado = historico_detalhado.sort_values(['Ano', 'Mês'], ascending=[False, False])
        
        # Pegar apenas os últimos registros para não sobrecarregar o PDF
        historico_detalhado = historico_detalhado.head(50)  # Limitar a 50 registros
        
        for _, registro in historico_detalhado.iterrows():
            quantidade = registro.get('Quantity', 0)
            valor_total = formatar_moeda(registro.get('Total Gross', 0))
            razao_social = row['RAZÃO SOCIAL'] if registro['Codigo_PN'] == sap_principal_filtro else 'COLIGADO'
            
            extrato_data.append([
                str(registro['Codigo_PN']),
                razao_social,
                registro.get('Item 2', ''),
                registro.get('Item', ''),
                str(int(quantidade)) if pd.notna(quantidade) else '0',
                valor_total,
                str(registro.get('Mês', '')),
                str(registro.get('Ano', ''))
            ])
        
        # Se há dados no extrato, criar tabela
        if len(extrato_data) > 1:
            tabela_extrato = Table(extrato_data, colWidths=[70, 120, 60, 120, 50, 70, 30, 40])
            tabela_extrato.setStyle(StyleBaseHistorica)
            tabela_extrato.wrapOn(c, width, height)
            
            if y_position - tabela_extrato.wrap(width, height)[1] < 50:
                c.showPage()
                y_position = height - 50
            
            tabela_extrato.drawOn(c, 50, y_position - tabela_extrato.wrap(width, height)[1])

        # ============================================================================
        # SEÇÃO 25: UPLOAD PARA BOX
        # ============================================================================
            c.save()
            print(f"PDF gerado: {nome_arquivo}")

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

        print("PDF criados")

      
                    
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
