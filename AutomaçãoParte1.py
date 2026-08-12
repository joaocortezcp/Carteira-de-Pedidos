import os
import sys
import time
import subprocess
import requests
import csv
import win32com.client
import pythoncom
import asyncio
from io import StringIO
from datetime import datetime
from playwright.sync_api import sync_playwright
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

# --- IMPORTAÇÕES DA INTERFACE GRÁFICA ---
import tkinter as tk
from tkinter import ttk
import threading

# =====================================================================#
# ==================== CONSTANTES DO EXCEL ============================#
# =====================================================================#
xlCalculationManual = -4135
xlCalculationAutomatic = -4105
xlPart = 2
xlDelimited = 1
xlRight = -4152
xlLeft = -4131
xlUp = -4162
xlCellTypeVisible = 12
xlDoubleQuote = 1
xlPasteFormulas = -4123 
xlOr = 2
xlCenter = -4108 # Nova constante adicionada para centralizar o texto

# =====================================================================#
# ======================= FUNÇÕES AUXILIARES ==========================#
# =====================================================================#

def limpar_valor(valor):
    """Garante a leitura perfeita dos dados em texto ao puxar da memória."""
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor).strip()

def obter_ano_mes(valor):
    """Extrai o ano e o mês da célula para validação, contornando o formato de texto."""
    if valor is None:
        return None, None
    if hasattr(valor, 'year') and hasattr(valor, 'month'):
        return valor.year, valor.month
    
    texto = str(valor).strip().replace(".", "/")
    if not texto or texto.lower() == "none":
        return None, None
    try:
        parte_data = texto.split(" ")[0]
        pedacos = parte_data.split("/")
        if len(pedacos) >= 3:
            mes = int(pedacos[1])
            ano = int(pedacos[2])
            if ano < 100:
                ano += 2000
            return ano, mes
    except Exception:
        pass
    return None, None


# =====================================================================#
# =================== 1. EXCEL: PREPARAÇÃO E LIMPEZA ==================#
# =====================================================================#

def AbrirUltimoRelatorio():
    caminho_base = r"C:\Users\Joao Cortez\Desktop\1 - Status de Pedidos"
    if not os.path.exists(caminho_base):
        raise FileNotFoundError(f"Erro: A pasta principal não foi encontrada em {caminho_base}")

    pastas = [os.path.join(caminho_base, nome) for nome in os.listdir(caminho_base) if os.path.isdir(os.path.join(caminho_base, nome))]
    if not pastas:
        raise FileNotFoundError("Nenhuma pasta encontrada.")
        
    pasta_mais_recente = max(pastas, key=os.path.getmtime)
    arquivos = [os.path.join(pasta_mais_recente, nome) for nome in os.listdir(pasta_mais_recente) if nome.endswith((".xlsb", ".xlsx"))]
    if not arquivos:
        raise FileNotFoundError("Nenhum arquivo Excel encontrado.")

    arquivo_mais_recente = max(arquivos, key=os.path.getmtime)
    
    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = True 
    workbook = excel_app.Workbooks.Open(arquivo_mais_recente)
    return excel_app, workbook

def LimparAbas(excel_app, workbook):
    abas_linhas = {
        "Base OT - Consolidados": 3,
        "CRÉDITO": 2,
        "TRACKING": 2,
        "SCP": 2
    }
    
    for aba, linha_inicio in abas_linhas.items():
        try:
            ws = workbook.Sheets(aba)
            ws.Visible = -1 
            ultima_linha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            if ultima_linha >= linha_inicio:
                ws.Range(f"{linha_inicio}:{ultima_linha}").Delete()
        except Exception:
            pass

    try:
        workbook.Sheets("CRÉDITO").Activate()
        workbook.Sheets("CRÉDITO").Range("A2").Select()
    except:
        pass


# =====================================================================#
# ================== 2. SAP: AUTOMAÇÃO (ÚNICA SESSÃO) =================#
# =====================================================================#

def IniciarSAP():
    SapApp = r"C:\Program Files (x86)\SAP\FrontEnd\SAPGUI\saplogon.exe"
    os.startfile(SapApp)
    time.sleep(10)

    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.OpenConnection("Global S/4 Prod [finance users] LAP", True)
    session = connection.Children(0)
    return session

def ExtrairCreditoSAP(session):
    data_hora = datetime.now().strftime("%d%mas%Hi%M")
    nome_arquivo = f"CreditoDia{data_hora}"
    caminho_diretorio = r"C:\Users\Joao Cortez\Desktop\Locais\Crédito"
    
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nY_LAD_65000039"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[1]/btn[17]").press()

    session.findById("wnd[1]/usr/txtENAME-LOW").text = "GP00211480"
    session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
    session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 10
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[33]").press()

    grid_layout = session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell")
    grid_layout.currentCellRow = 257
    grid_layout.selectedRows = "257"
    grid_layout.contextMenu()
    grid_layout.selectContextMenuItem("&FIND")

    session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").selected = True
    session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = "/BRDDVM"
    session.findById("wnd[2]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
    session.findById("wnd[2]/tbar[0]/btn[0]").press()
    session.findById("wnd[2]/tbar[0]/btn[12]").press()

    grid_layout.selectedRows = "60"
    grid_layout.clickCurrentCell()

    main_grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")
    main_grid.selectedRows = "0"
    main_grid.contextMenu()
    main_grid.selectContextMenuItem("&XXL")

    campo_nome_arquivo = session.findById("wnd[1]/usr/ssubSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/txtGS_EXPORT-FILE_NAME")
    campo_nome_arquivo.text = nome_arquivo
    session.findById("wnd[1]/tbar[0]/btn[20]").press()

    session.findById("wnd[1]/usr/ctxtDY_PATH").text = caminho_diretorio
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(5) 
    
    return nome_arquivo, caminho_diretorio

def ExtrairOTSAP(session):
    data_hora = datetime.now().strftime("%d%mas%Hi%M")
    nome_arquivo = f"OtDia{data_hora}"
    caminho_diretorio = r"C:\Users\Joao Cortez\Desktop\Locais\OT"

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nY_LAD_65000605"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[1]/btn[17]").press()

    session.findById("wnd[1]/usr/txtENAME-LOW").text = "GP00211480"
    session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
    session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 10
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    main_container = session.findById("wnd[0]/usr/cntlMY_CONTAINER/shellcont/shell")
    main_container.pressToolbarContextButton("&MB_VARIANT")
    main_container.selectContextMenuItem("&LOAD")

    layout_container = session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell")
    layout_container.contextMenu()
    layout_container.selectContextMenuItem("&FIND")

    session.findById("wnd[2]/usr/chkGS_SEARCH-SHOW_HITS").selected = True
    session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = "/BRBBTH"
    session.findById("wnd[2]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
    session.findById("wnd[2]/tbar[0]/btn[0]").press()
    session.findById("wnd[2]/tbar[0]/btn[12]").press()

    layout_container.selectedRows = "46"
    layout_container.clickCurrentCell()

    main_container.currentCellRow = 3
    main_container.contextMenu()
    main_container.selectContextMenuItem("&XXL")

    session.findById("wnd[1]/tbar[0]/btn[20]").press()

    session.findById("wnd[1]/usr/ctxtDY_PATH").text = caminho_diretorio
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = f"{nome_arquivo}.xlsx"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(8) 
    
    return nome_arquivo, caminho_diretorio

def FecharSAP():
    os.system("taskkill /f /im saplogon.exe >nul 2>&1")
    time.sleep(2)


# =====================================================================#
# =================== 3. EXCEL: CÓPIA E FORMATAÇÃO ====================#
# =====================================================================#

def ProcessarCredito(workbook_principal, nome_arquivo_sap, caminho_diretorio, excel_app):
    wb_sap = None
    for wb in excel_app.Workbooks:
        if nome_arquivo_sap in wb.Name:
            wb_sap = wb
            break
            
    if not wb_sap:
        caminho_completo = os.path.join(caminho_diretorio, f"{nome_arquivo_sap}.xlsx")
        wb_sap = excel_app.Workbooks.Open(caminho_completo)
        
    ws_sap = wb_sap.ActiveSheet
    ws_sap.AutoFilterMode = False
    TotalLinhas = ws_sap.Cells(ws_sap.Rows.Count, 1).End(xlUp).Row

    if TotalLinhas > 1:
        ws_sap.Range(f"A1:B{TotalLinhas}").AutoFilter(Field=2, Criteria1="=")
        try:
            ws_sap.Range(f"A2:A{TotalLinhas}").SpecialCells(xlCellTypeVisible).EntireRow.Delete()
        except:
            pass 
        ws_sap.AutoFilterMode = False

        TotalLinhas = ws_sap.Cells(ws_sap.Rows.Count, 1).End(xlUp).Row 
        ws_sap.Range(f"A1:B{TotalLinhas}").AutoFilter(Field=1, Criteria1="=")
        try:
            ws_sap.Range(f"A2:A{TotalLinhas}").SpecialCells(xlCellTypeVisible).EntireRow.Delete()
        except:
            pass 
        ws_sap.AutoFilterMode = False 

        TotalLinhas = ws_sap.Cells(ws_sap.Rows.Count, 1).End(xlUp).Row

    if TotalLinhas > 1:
        ws_sap.Range(f"A1:A{TotalLinhas}").TextToColumns(
            Destination=ws_sap.Range("A1"), DataType=xlDelimited,
            TextQualifier=xlDoubleQuote, ConsecutiveDelimiter=False, 
            Tab=True, Semicolon=False, Comma=False, Space=False, Other=False
        )

        valores_extraidos = ws_sap.Range(f"A2:B{TotalLinhas}").Value
        ws_credito = workbook_principal.Sheets("CRÉDITO")
        ws_credito.Range(f"A2:B{TotalLinhas}").Value = valores_extraidos

    excel_app.CutCopyMode = False
    excel_app.DisplayAlerts = False 
    wb_sap.Close(SaveChanges=False)
    excel_app.DisplayAlerts = True 


def ProcessarOT(workbook_principal, nome_arquivo_sap, caminho_diretorio, excel_app):
    wb_sap = None
    for wb in excel_app.Workbooks:
        if nome_arquivo_sap in wb.Name:
            wb_sap = wb
            break
            
    if not wb_sap:
        caminho_completo = os.path.join(caminho_diretorio, f"{nome_arquivo_sap}.xlsx")
        wb_sap = excel_app.Workbooks.Open(caminho_completo)
        
    ws_sap = wb_sap.ActiveSheet
    ws_destino = workbook_principal.Sheets("Base OT - Consolidados")

    ultima_linha = ws_sap.Cells(ws_sap.Rows.Count, 1).End(xlUp).Row
    ultima_coluna = ws_sap.Cells(1, ws_sap.Columns.Count).End(-4159).Column

    if ultima_linha >= 2:
        valores_extraidos = ws_sap.Range(ws_sap.Cells(2, 1), ws_sap.Cells(ultima_linha, ultima_coluna)).Value
        ws_destino.Range(ws_destino.Cells(3, 1), ws_destino.Cells(ultima_linha + 1, ultima_coluna)).Value = valores_extraidos

    excel_app.CutCopyMode = False
    excel_app.DisplayAlerts = False 
    wb_sap.Close(SaveChanges=False)
    excel_app.DisplayAlerts = True 

    TratarDadosOT(ws_destino, excel_app)


def TratarDadosOT(ws, excel_app):
    excel_app.ScreenUpdating = False
    excel_app.Calculation = xlCalculationManual
    
    try:
        try:
            ws.Unprotect(Password="")
        except:
            pass
            
        ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        if ultimaLinha < 3:
            ultimaLinha = 3 

        dados = ws.Range(f"A1:R{ultimaLinha}").Value

        hoje = datetime.now()
        meses_permitidos = []
        for i in range(3):
            m = hoje.month - i
            a = hoje.year
            if m <= 0:
                m += 12
                a -= 1
            meses_permitidos.append((a, m))

        linhas_para_deletar = []
        
        for i in range(ultimaLinha, 2, -1):
            idx = i - 1 
            
            val_I = limpar_valor(dados[idx][8])
            val_K = limpar_valor(dados[idx][10])
            val_M = limpar_valor(dados[idx][12])
            val_N = limpar_valor(dados[idx][13])
            val_O = dados[idx][14]

            deletar = False
            
            if val_I.startswith("1104") or val_M.startswith("1143"):
                deletar = True
            elif val_N != "" and val_K == "":
                deletar = True
            else:
                ano_celula, mes_celula = obter_ano_mes(val_O)
                if ano_celula is not None and mes_celula is not None:
                    if (ano_celula, mes_celula) not in meses_permitidos:
                        deletar = True

            if deletar:
                linhas_para_deletar.append(i)

        if linhas_para_deletar:
            inicio_bloco = linhas_para_deletar[0]
            fim_bloco = linhas_para_deletar[0]
            
            for linha in linhas_para_deletar[1:]:
                if linha == fim_bloco - 1:
                    fim_bloco = linha
                else:
                    ws.Range(f"{fim_bloco}:{inicio_bloco}").EntireRow.Delete()
                    inicio_bloco = linha
                    fim_bloco = linha
            
            ws.Range(f"{fim_bloco}:{inicio_bloco}").EntireRow.Delete()

        ws.Columns("C:C").Replace(What="BE6", Replacement="BD3", LookAt=xlPart)

        colunas_texto = ["D", "G", "I", "K", "M", "N"]
        for col in colunas_texto:
            ws.Columns(f"{col}:{col}").TextToColumns(
                Destination=ws.Range(f"{col}1"), 
                DataType=xlDelimited, 
                FieldInfo=(1, 1), 
                TrailingMinusNumbers=True
            )

        ws.Columns("Q:Q").TextToColumns(
            Destination=ws.Range("Q1"), 
            DataType=xlDelimited, 
            FieldInfo=(1, 1)
        )
        ws.Columns("Q:Q").NumberFormatLocal = "0,00"
        ws.Columns("Q:Q").Replace(What=".", Replacement=",", LookAt=xlPart)

        colunas_data = ["E", "J", "L", "O", "P", "R"] 
        for col in colunas_data:
            ws.Columns(f"{col}:{col}").Replace(What=".", Replacement="/", LookAt=xlPart)
            ws.Columns(f"{col}:{col}").HorizontalAlignment = xlRight
            ws.Columns(f"{col}:{col}").TextToColumns(
                Destination=ws.Range(f"{col}1"), 
                DataType=xlDelimited, 
                FieldInfo=(1, 4)
            )
            ws.Columns(f"{col}:{col}").NumberFormatLocal = "dd/mm/aaaa"

        ws.Rows(2).HorizontalAlignment = xlLeft
        ws.Cells.EntireColumn.AutoFit()
        ws.Activate()           
        ws.Cells(1, 1).Select()
        
        ws.Calculate()

    finally:
        excel_app.Calculation = xlCalculationAutomatic
        excel_app.ScreenUpdating = True


# =====================================================================#
# =================== 4. SIRIUS: WEB EXTRACTION =======================#
# =====================================================================#

def ExtrairEProcessarSirius(usuario, senha, excel_app, workbook_principal):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        page.goto("https://sirius.fiorde.com.br/")
        page.fill('#user-input', usuario)
        page.fill('#password-input', senha)
        page.click('#submit-button')

        page.wait_for_selector('text="Relatórios"', timeout=15000)
        page.locator('text="Relatórios"').click()
        page.locator('text="Relatórios Exportados"').click()

        page.wait_for_selector('.MuiSelect-select', timeout=15000) 
        page.locator('.MuiSelect-select').click()
        page.locator('li[data-value="/pendencias"]').click()

        page.locator('.ant-select-selection__rendered').click()
        page.locator('text="Colgate"').click()

        with page.expect_download(timeout=90000) as download_info:
            page.locator('text="Extrair"').click()
        
        download = download_info.value
        nome_original = download.suggested_filename
        
        pasta_destino = r"C:\Users\Joao Cortez\Desktop\Locais\Fiorde Pendencias" 
        caminho_arquivo = os.path.join(pasta_destino, nome_original)
        
        download.save_as(caminho_arquivo)
        browser.close()

    time.sleep(1)
    
    if not os.path.exists(caminho_arquivo):
        raise FileNotFoundError("ERRO: O arquivo Sirius não foi encontrado após o download.")

    wb_sirius = excel_app.Workbooks.Open(caminho_arquivo)
    ws_sirius = wb_sirius.ActiveSheet
    ws_tracking = workbook_principal.Sheets("TRACKING")
    
    ultima_linha_sirius = ws_sirius.Cells(ws_sirius.Rows.Count, 1).End(xlUp).Row

    ws_sirius.Columns("J:J").TextToColumns(
        Destination=ws_sirius.Range("J1"), 
        DataType=xlDelimited, 
        FieldInfo=(1, 1), 
        TrailingMinusNumbers=True
    )
    
    ws_sirius.UsedRange.AutoFilter(Field=2, Criteria1="=BRDD", Operator=xlOr, Criteria2="=BREE")

    colunas_origem = ["J", "AG", "AH", "AJ", "AK"]
    colunas_destino = ["A", "B", "C", "D", "E"]
    
    for col_origem, col_destino in zip(colunas_origem, colunas_destino):
        ws_sirius.Range(f"{col_origem}1:{col_origem}{ultima_linha_sirius}").Copy()
        ws_tracking.Paste(Destination=ws_tracking.Range(f"{col_destino}1"))
    
    excel_app.CutCopyMode = False

    excel_app.DisplayAlerts = False
    wb_sirius.Close(SaveChanges=False)
    excel_app.DisplayAlerts = True


# =====================================================================#
# ================= 5. GOOGLE SHEETS: SELENIUM + REQUESTS =============#
# =====================================================================#

def ExtrairGoogleSheets(excel_app, workbook_principal):
    # Força o encerramento de qualquer Chrome "fantasma"
    os.system("taskkill /f /im chrome.exe /T >nul 2>&1")
    time.sleep(3)
    
    caminho_chrome = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    pasta_perfil = r"C:\Users\Joao Cortez\Desktop\Sheets Automação"
    
    comando = (
        f'"{caminho_chrome}" '
        f'--remote-debugging-port=9222 '
        f'--user-data-dir="{pasta_perfil}" '
        f'--profile-directory="Profile 1" '
        f'--disable-features=ProfilePicker'
    )
    subprocess.Popen(comando)
    time.sleep(5)

    opcoes = Options()
    opcoes.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

    driver = webdriver.Chrome(options=opcoes)
    driver.switch_to.new_window('tab')
    
    driver.get("https://docs.google.com/spreadsheets/d/1nSdYBMjI0FOarvCrgEqau8tvrfv1yOO_MffVfztR588/edit?gid=0#gid=0")
    
    # SISTEMA DE RETRY PARA PREVENIR ERRO 401 E LENTIDÃO NA REDE
    sucesso_download = False
    tentativas = 0
    max_tentativas = 3
    url_exportacao = "https://docs.google.com/spreadsheets/d/1nSdYBMjI0FOarvCrgEqau8tvrfv1yOO_MffVfztR588/export?format=csv&gid=0"
    
    while tentativas < max_tentativas and not sucesso_download:
        time.sleep(8) # Espera inicial (ou entre tentativas)
        
        sessao = requests.Session()
        sessao.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        })

        # Pega os cookies atuais da página (se estiver demorando, os cookies de auth podem não estar prontos na 1ª vez)
        for cookie in driver.get_cookies():
            sessao.cookies.set(cookie['name'], cookie['value'])

        resposta = sessao.get(url_exportacao)
        
        # Se for 200 (Sucesso), saímos do loop
        if resposta.status_code == 200:
            sucesso_download = True
        else:
            tentativas += 1
            print(f"Tentativa {tentativas} falhou com status {resposta.status_code}. Atualizando página e tentando novamente...")
            driver.refresh() # Atualiza a página para forçar o carregamento dos cookies
    
    if not sucesso_download:
        driver.quit()
        raise Exception(f"Erro ao baixar a planilha após {max_tentativas} tentativas. O Google retornou: {resposta.status_code}")

    resposta.encoding = 'utf-8'
    leitor_csv = csv.reader(StringIO(resposta.text))
    
    try:
        cabecalho = next(leitor_csv) 
    except StopIteration:
        driver.quit()
        raise Exception("A planilha retornou vazia. O Google pode estar bloqueando a requisição ou a autenticação falhou.")
    
    linhas_filtradas = []
    for linha in leitor_csv:
        if len(linha) > 6 and linha[6] in ["BRDD", "BREE"]:
            linhas_filtradas.append(linha)
            
    if linhas_filtradas:
        max_colunas = max(len(linha) for linha in linhas_filtradas)
        for linha in linhas_filtradas:
            linha.extend([""] * (max_colunas - len(linha)))

        num_linhas = len(linhas_filtradas)
        ws_scp = workbook_principal.Sheets("SCP")
        range_destino = ws_scp.Range(ws_scp.Cells(2, 1), ws_scp.Cells(num_linhas + 1, max_colunas))
        range_destino.Value = linhas_filtradas

    driver.quit()
    os.system("taskkill /f /im chrome.exe /T >nul 2>&1")


# =====================================================================#
# =================== 6. LIMPEZA FINAL DA BASE OT =====================#
# =====================================================================#

def LimpezaFinalBaseOT(workbook_principal, excel_app):
    ws = workbook_principal.Sheets("Base OT - Consolidados")
    
    excel_app.ScreenUpdating = False
    excel_app.Calculation = xlCalculationManual
    
    try:
        if ws.AutoFilterMode:
            ws.AutoFilterMode = False

        ultima_linha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        # PASSO 1: Fórmulas
        if ultima_linha >= 3:
            ws.Range("W1:AQ1").Copy()
            ws.Range(f"W3:AQ{ultima_linha}").PasteSpecial(Paste=xlPasteFormulas)
            excel_app.CutCopyMode = False
        
        ws.Calculate()

        # PASSO 2: Limpeza (de baixo para cima)
        for i in range(ultima_linha, 2, -1):
            val_AG = ws.Range(f"AG{i}").Value
            val_AG = str(val_AG).strip() if val_AG is not None else ""
            
            val_AH = ws.Range(f"AH{i}").Value
            val_AH = str(val_AH).strip().lower() if val_AH is not None else ""
            
            val_AI = ws.Range(f"AI{i}").Value
            val_AI = str(val_AI).strip().lower() if val_AI is not None else ""
            
            texto_AM = ws.Range(f"AM{i}").Text

            if val_AG == "Bloqueado por vendas/CS" and (val_AH != "" or val_AI != ""):
                ws.Rows(i).Delete()
            else:
                if val_AI == "sim":
                    ws.Range(f"AG{i}").ClearContents()
                    ws.Range(f"AH{i}").ClearContents()
                    val_AG = ""
                    val_AH = ""

                if val_AH == "sim":
                    ws.Range(f"AG{i}").ClearContents()
                    val_AG = ""

                if val_AG not in ["Bloqueado por vendas/CS", "MOQ", "OOS", ""]:
                    ws.Range(f"AG{i}").ClearContents()

                if texto_AM in ["#N/A", "#N/D"]:
                    ws.Range(f"R{i}").FormulaR1C1 = ws.Range("R1").FormulaR1C1

        # PASSO FINAL: Filtro, Centralização e AutoFit
        ultima_linha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        if not ws.AutoFilterMode:
            ws.Range(f"A2:AQ{ultima_linha}").AutoFilter()

        used_range = ws.UsedRange
        used_range.HorizontalAlignment = xlCenter
        used_range.VerticalAlignment = xlCenter
        used_range.EntireColumn.AutoFit()
        
        workbook_principal.Save()
        
    finally:
        excel_app.Calculation = xlCalculationAutomatic
        excel_app.ScreenUpdating = True


# =====================================================================#
# =================== 7. LIMPEZA FINAL DE PROCESSOS ===================#
# =====================================================================#

def FecharExcelsSecundarios(excel_app, nome_principal):
    """Fecha todos os relatórios do Excel abertos que não sejam o arquivo final principal."""
    for wb in excel_app.Workbooks:
        if wb.Name != nome_principal:
            excel_app.DisplayAlerts = False
            try:
                wb.Close(SaveChanges=False)
            except:
                pass
            excel_app.DisplayAlerts = True


# =====================================================================#
# ====================== INTERFACE GRÁFICA (GUI) ======================#
# =====================================================================#

em_execucao = False
tempo_inicial = 0
total_passos_menu = 7 # ATUALIZADO PARA 7 PASSOS

def AtualizarRelogio():
    """Faz o cronômetro girar em tempo real na tela, segundo a segundo."""
    if em_execucao:
        tempo_decorrido = int(time.time() - tempo_inicial)
        minutos, segundos = divmod(tempo_decorrido, 60)
        lbl_tempo.config(text=f"⏱️ Tempo Decorrido: {minutos:02d}:{segundos:02d}")
        janela.after(1000, AtualizarRelogio)

def AtualizarProgresso(passo, descricao):
    """Envia a atualização para a janela de forma segura (Thread-safe)."""
    porcentagem = int((passo / total_passos_menu) * 100)
    
    def atualizar_tela():
        barra_progresso['value'] = porcentagem
        lbl_porcentagem.config(text=f"{porcentagem}%")
        lbl_status.config(text=f"▶️ Passo {passo}/{total_passos_menu}: {descricao}")
    
    # Manda a janela principal atualizar os dados com segurança
    janela.after(0, atualizar_tela)
    print(f"\n[Progresso: {porcentagem}%] Passo {passo}: {descricao}")

def FluxoPrincipal():
    """Esta é a função que vai rodar em segundo plano (Thread)."""
    global em_execucao, tempo_inicial
    
    try:
        # 1. Inicializa o ambiente COM para o Excel/SAP (Evita o erro CoInitialize)
        pythoncom.CoInitialize() 
        # 2. Inicializa o motor do Playwright para rodar liso na Thread
        asyncio.set_event_loop(asyncio.new_event_loop()) 
        
        # PASSO 1
        AtualizarProgresso(1, "Preparando Relatório Principal...")
        meu_excel, workbook_principal = AbrirUltimoRelatorio()
        nome_relatorio = workbook_principal.Name
        LimparAbas(meu_excel, workbook_principal)

        # PASSO 2
        AtualizarProgresso(2, "Conectando e Extraindo dados do SAP...")
        sessao_sap = IniciarSAP()
        nome_credito, dir_credito = ExtrairCreditoSAP(sessao_sap)
        nome_ot, dir_ot = ExtrairOTSAP(sessao_sap)
        FecharSAP()

        # PASSO 3
        AtualizarProgresso(3, "Formatando dados do SAP no Excel...")
        ProcessarCredito(workbook_principal, nome_credito, dir_credito, meu_excel)
        ProcessarOT(workbook_principal, nome_ot, dir_ot, meu_excel)

        # PASSO 4
        AtualizarProgresso(4, "Extraindo portal Sirius (Playwright)...")
        ExtrairEProcessarSirius("joao_cortez@colpal.com", "Colg@te1", meu_excel, workbook_principal)

        # PASSO 5
        AtualizarProgresso(5, "Extraindo Google Sheets (Selenium)...")
        ExtrairGoogleSheets(meu_excel, workbook_principal)
        
        # PASSO 6 - NOVO PASSO
        AtualizarProgresso(6, "Limpeza final e formatação (Base OT)...")
        LimpezaFinalBaseOT(workbook_principal, meu_excel)

        # PASSO 7
        AtualizarProgresso(7, "Limpando memória e fechando processos...")
        FecharExcelsSecundarios(meu_excel, nome_relatorio)

        # CONCLUÍDO
        def sucesso_tela():
            lbl_status.config(text="✅ AUTOMAÇÃO FINALIZADA COM SUCESSO!", fg="green")
            btn_iniciar.config(state=tk.NORMAL)
        janela.after(0, sucesso_tela)
        
    except Exception as e:
        def erro_tela():
            lbl_status.config(text=f"❌ Ocorreu um erro. Verifique o terminal.", fg="red")
            btn_iniciar.config(state=tk.NORMAL)
        janela.after(0, erro_tela)
        print(f"Erro na execução: {e}")
        
    finally:
        em_execucao = False

def IniciarAutomacao():
    """Inicia a thread do robô e o relógio para não travar a janela."""
    global em_execucao, tempo_inicial
    
    btn_iniciar.config(state=tk.DISABLED)
    barra_progresso['value'] = 0
    lbl_porcentagem.config(text="0%")
    lbl_status.config(text="Iniciando...", fg="black")
    
    em_execucao = True
    tempo_inicial = time.time()
    AtualizarRelogio()
    
    thread_robo = threading.Thread(target=FluxoPrincipal)
    thread_robo.daemon = True
    thread_robo.start()


# =====================================================================#
# ==================== CONSTRUÇÃO DA JANELA ===========================#
# =====================================================================#

if __name__ == "__main__":
    janela = tk.Tk()
    janela.title("Robô de Automação - Status de Pedidos")
    janela.geometry("450x250")
    janela.resizable(False, False)
    
    frame = tk.Frame(janela, padx=20, pady=20)
    frame.pack(expand=True, fill=tk.BOTH)

    lbl_titulo = tk.Label(frame, text="Automação de Relatórios", font=("Arial", 14, "bold"))
    lbl_titulo.pack(pady=(0, 15))

    lbl_tempo = tk.Label(frame, text="⏱️ Tempo Decorrido: 00:00", font=("Arial", 10))
    lbl_tempo.pack()

    barra_progresso = ttk.Progressbar(frame, orient="horizontal", length=350, mode="determinate")
    barra_progresso.pack(pady=10)

    lbl_porcentagem = tk.Label(frame, text="0%", font=("Arial", 10, "bold"))
    lbl_porcentagem.pack()

    lbl_status = tk.Label(frame, text="Aguardando início...", font=("Arial", 9), fg="gray")
    lbl_status.pack(pady=5)

    btn_iniciar = tk.Button(frame, text="🚀 INICIAR AUTOMAÇÃO", font=("Arial", 10, "bold"), 
                            bg="#4CAF50", fg="white", width=25, height=2, command=IniciarAutomacao)
    btn_iniciar.pack(pady=(10, 0))

    janela.mainloop()
