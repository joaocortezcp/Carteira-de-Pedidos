import win32com.client as win32
import os
import time
from datetime import datetime

# ==========================================
# CONSTANTES GERAIS DO EXCEL
# ==========================================
XL_UP = -4162
XL_TO_LEFT = -4159
XL_FILTER_FONT_COLOR = 13 
XL_CELL_TYPE_VISIBLE = 12
XL_DELIMITED = 1
XL_DOUBLE_QUOTE = 1
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105
XL_PASTE_VALUES = -4163 
XL_AND = 1 # Usado para combinar dois critérios no filtro

# ==========================================
# FUNÇÕES DE APOIO
# ==========================================
def obter_filtro_mes_atual():
    """Identifica o mês atual do sistema e retorna a string de filtro correta."""
    meses = {
        1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL",
        5: "MAIO", 6: "JUNHO", 7: "JULHO", 8: "AGOSTO",
        9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO"
    }
    mes_atual = meses[datetime.now().month]
    return f"FATURADO {mes_atual}"

def obter_meses_passados():
    """Identifica os dois meses anteriores ao atual considerando virada de ano."""
    meses = {
        1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL",
        5: "MAIO", 6: "JUNHO", 7: "JULHO", 8: "AGOSTO",
        9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO"
    }
    mes_atual = datetime.now().month
    
    idx_passado = mes_atual - 1 if mes_atual > 1 else 12
    idx_retrasado = idx_passado - 1 if idx_passado > 1 else 12
    
    return f"FATURADO {meses[idx_passado]}", f"FATURADO {meses[idx_retrasado]}"

# ==========================================
# 1. CONEXÃO E LIMPEZA DO EXCEL BASE (JÁ ABERTO)
# ==========================================
def conectar_e_limpar_excel():
    print("Procurando o Excel e a planilha 'BRDD' já abertos na tela...")
    
    try:
        excel = win32.GetActiveObject("Excel.Application")
    except Exception:
        print("Erro: O Excel não está aberto. Por favor, abra o arquivo primeiro.")
        return None, None
        
    wb_inicial = None
    for wb in excel.Workbooks:
        if wb.Name.startswith("Relatório de Status de Pedidos - BRDD"):
            wb_inicial = wb
            break
            
    if not wb_inicial:
        print("Erro: O arquivo 'Relatório de Status de Pedidos - BRDD...' não foi encontrado aberto.")
        return None, None

    # Otimização de Performance
    excel.DisplayAlerts = False 
    excel.ScreenUpdating = False # Desliga a atualização visual (deixa MUITO mais rápido)
    
    print("Processando aba 'Item a item' (Limpando base antiga)...")
    ws_item = wb_inicial.Sheets("Item a item")
    ws_item.AutoFilterMode = False
    
    ultima_linha_item = ws_item.Cells(ws_item.Rows.Count, 1).End(XL_UP).Row
    if ultima_linha_item >= 2:
        rng_item = ws_item.Range(f"A1:AA{ultima_linha_item}")
        
        str_mes_passado, str_mes_retrasado = obter_meses_passados()
        print(f"-> Mantendo APENAS: '{str_mes_passado}' e '{str_mes_retrasado}' (Coluna Y)")
        
        try:
            # Filtra tudo que for DIFERENTE do mês passado E DIFERENTE do retrasado
            rng_item.AutoFilter(
                Field=25, 
                Criteria1=f"<>{str_mes_passado}", 
                Operator=XL_AND, 
                Criteria2=f"<>{str_mes_retrasado}"
            )
            
            # Deleta as linhas filtradas (excluindo o cabeçalho)
            linhas_para_deletar = ws_item.Range(f"A2:AA{ultima_linha_item}").SpecialCells(XL_CELL_TYPE_VISIBLE)
            linhas_para_deletar.EntireRow.Delete()
        except Exception:
            pass # Se não houver nada no filtro, segue o fluxo
            
    ws_item.AutoFilterMode = False
    return excel, wb_inicial

# ==========================================
# 2. FUNÇÃO UNIVERSAL DE FILTRO E CÓPIA
# ==========================================
def copiar_dados_por_status(wb_inicial, status, coluna_copia="D"):
    print(f"\nPreparando dados: Filtrando por '{status}' e copiando coluna {coluna_copia}...")
    ws_base = wb_inicial.Sheets("Base OT - Consolidados")
    ws_base.AutoFilterMode = False
    
    ultima_linha_base = ws_base.Cells(ws_base.Rows.Count, 1).End(XL_UP).Row
    rng_base = ws_base.Range(f"A2:AQ{ultima_linha_base}")
    
    rng_base.AutoFilter(Field=36, Criteria1=status)
    
    try:
        copia_range = ws_base.Range(f"{coluna_copia}3:{coluna_copia}{ultima_linha_base}").SpecialCells(XL_CELL_TYPE_VISIBLE)
        copia_range.Copy()
        print(f"SUCESSO! Dados de '{status}' copiados para a área de transferência.")
        return True
    except Exception:
        print(f"Aviso: Nenhuma informação para '{status}' foi encontrada.")
        ws_base.AutoFilterMode = False
        return False

# ==========================================
# 3. CONEXÃO COM O SAP
# ==========================================
def conectar_sap():
    print("Iniciando SAP Logon...")
    SapApp = r"C:\Program Files (x86)\SAP\FrontEnd\SAPGUI\saplogon.exe"
    os.startfile(SapApp)
    time.sleep(10) 

    try:
        SapGuiAuto = win32.GetObject("SAPGUI")
    except Exception as e:
        print("Erro: O SAP não carregou a tempo. Aumente o time.sleep().")
        raise e

    application = SapGuiAuto.GetScriptingEngine
    connection = application.OpenConnection("Global S/4 Prod [finance users] LAP", True)
    time.sleep(3) 
    
    return connection.Children(0)

# ==========================================
# 4. EXTRAÇÕES NO SAP
# ==========================================
def extrair_sap_carteira(session, diretorio):
    print("[SAP] Extraindo: EM CARTEIRA...")
    data_hora = datetime.now().strftime("%d%mas%Hi%M")
    nome_arquivo = f"EmCarteiraDia{data_hora}.xlsx"
    caminho_completo = os.path.join(diretorio, nome_arquivo)

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nY_LAD_65000038"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    
    session.findById("wnd[1]/usr/txtENAME-LOW").text = "GP00211480"
    session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
    session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 10
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    
    session.findById("wnd[0]/usr/ctxtERDAT-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtERDAT-LOW").caretPosition = 0
    session.findById("wnd[0]/usr/ctxtERDAT-LOW").showContextMenu()
    session.findById("wnd[0]/usr").selectContextMenuItem("DELACTX")
    session.findById("wnd[0]/usr/btn%_VBELN_%_APP_%-VALU_PUSH").press()
    
    # SEGURANÇA: Limpa a lista antes de colar os novos para evitar acúmulos de rodadas antigas
    try: session.findById("wnd[1]/tbar[0]/btn[16]").press() 
    except: pass
    session.findById("wnd[1]/tbar[0]/btn[24]").press() # Botão Colar
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    
    grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")
    grid.currentCellColumn = "VBELN"
    grid.selectedRows = "0"
    grid.contextMenu()
    grid.selectContextMenuItem("&XXL")
    
    session.findById("wnd[1]/tbar[0]/btn[20]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = diretorio
    try:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo
    except: pass 
    
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = len(diretorio)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    return caminho_completo

def extrair_sap_faturamento(session, diretorio):
    print("[SAP] Extraindo: EM FATURAMENTO...")
    data_hora = datetime.now().strftime("%d%mas%Hi%M")
    nome_arquivo = f"EmFaturamentoDia{data_hora}.xlsx" 
    caminho_completo = os.path.join(diretorio, nome_arquivo)

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nY_LAD_65000687"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()
    
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtENAME-LOW").text = "GP00201400"
    session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
    session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 10
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    
    session.findById("wnd[0]/usr/ctxtSP$00012-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtSP$00012-LOW").caretPosition = 0
    session.findById("wnd[0]/usr/ctxtSP$00012-LOW").showContextMenu()
    session.findById("wnd[0]/usr").selectContextMenuItem("DELACTX")
    
    session.findById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH").press() 
    
    # SEGURANÇA: Limpa a lista antes de colar
    try: session.findById("wnd[1]/tbar[0]/btn[16]").press() 
    except: pass
    session.findById("wnd[1]/tbar[0]/btn[24]").press() 
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    
    session.findById("wnd[0]/tbar[1]/btn[33]").press()
    layout_grid = session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell")
    layout_grid.currentCellRow = 136
    layout_grid.selectedRows = "136"
    layout_grid.contextMenu()
    layout_grid.selectContextMenuItem("&FIND")
    
    session.findById("wnd[2]/usr/chkGS_SEARCH-SHOW_HITS").selected = True
    session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = "/THAINÁ"
    session.findById("wnd[2]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
    session.findById("wnd[2]/usr/chkGS_SEARCH-SHOW_HITS").setFocus()
    session.findById("wnd[2]/tbar[0]/btn[0]").press()
    session.findById("wnd[2]").close()
    
    layout_grid.selectedRows = "139"
    layout_grid.clickCurrentCell()
    
    main_grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    main_grid.setCurrentCell(5, "KNA1-NAME1")
    main_grid.selectedRows = "5"
    main_grid.contextMenu()
    main_grid.selectContextMenuItem("&XXL")
    
    session.findById("wnd[1]/tbar[0]/btn[20]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = diretorio
    try:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo
    except: pass 
    
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = len(diretorio)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    return caminho_completo

def extrair_sap_faturado_mes_atual(session, diretorio):
    print("[SAP] Extraindo: FATURADO MÊS ATUAL...")
    data_hora = datetime.now().strftime("%d%mas%Hi%M")
    nome_arquivo = f"FaturadoMesAtualDia{data_hora}.xlsx" 
    caminho_completo = os.path.join(diretorio, nome_arquivo) 

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nY_LAD_65000695"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()
    
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtENAME-LOW").text = "GP00201400"
    session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
    session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 10
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    
    session.findById("wnd[0]/usr/ctxtS_VKBUR-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtS_VKBUR-LOW").caretPosition = 0
    session.findById("wnd[0]/usr/ctxtS_VKBUR-LOW").showContextMenu()
    session.findById("wnd[0]/usr").selectContextMenuItem("DELACTX")
    session.findById("wnd[0]/usr/btn%_S_VBELN_%_APP_%-VALU_PUSH").press()
    
    # SEGURANÇA: Limpa a lista antes de colar
    try: session.findById("wnd[1]/tbar[0]/btn[16]").press() 
    except: pass
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    
    tabela_principal = session.findById("wnd[0]/usr/cntlMY_CONTAINER/shellcont/shell")
    tabela_principal.pressToolbarContextButton("&MB_VARIANT")
    tabela_principal.selectContextMenuItem("&LOAD")
    
    layout_shell = session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell")
    layout_shell.contextMenu()
    layout_shell.selectContextMenuItem("&FIND")
    
    session.findById("wnd[2]/usr/chkGS_SEARCH-SHOW_HITS").selected = True 
    session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = "/THAINÁBRBB"
    session.findById("wnd[2]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
    session.findById("wnd[2]/tbar[0]/btn[0]").press()
    session.findById("wnd[2]/tbar[0]/btn[12]").press()
    
    layout_shell.selectedRows = "254"
    layout_shell.clickCurrentCell()
    
    tabela_principal.setCurrentCell(3, "NAMEC") 
    tabela_principal.contextMenu()
    tabela_principal.selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[20]").press()
    
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = diretorio
    try:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo
    except: pass 
    
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = len(diretorio)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    return caminho_completo

def extrair_sap_atp(session, diretorio):
    print("[SAP] Extraindo: ATP...")
    data_hora = datetime.now().strftime("%d%mas%Hi%M")
    nome_arquivo = f"ATPDia{data_hora}.xlsx"
    caminho_completo = os.path.join(diretorio, nome_arquivo)

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nY_LAD_65000026S4"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()
    
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtENAME-LOW").text = "GP00201400"
    session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
    session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 10
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Prepara o campo e abre seleção múltipla
    session.findById("wnd[0]/usr/ctxtSP$00007-LOW").showContextMenu()
    session.findById("wnd[0]/usr").selectContextMenuItem("DELACTX")
    session.findById("wnd[0]/usr/btn%_SP$00004_%_APP_%-VALU_PUSH").press()
    
    # Limpa a lista existente (segurança) e cola os novos dados da área de transferência
    try: session.findById("wnd[1]/tbar[0]/btn[16]").press() 
    except: pass
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    
    # Executa relatório
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    
    # Selecionar Layout //JOAO
    session.findById("wnd[0]/tbar[1]/btn[33]").press()
    layout_grid = session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell")
    layout_grid.currentCellRow = 4
    layout_grid.selectedRows = "4"
    layout_grid.contextMenu()
    layout_grid.selectContextMenuItem("&FIND")

    session.findById("wnd[2]/usr/chkGS_SEARCH-SHOW_HITS").selected = True
    session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = "//JOAO"
    session.findById("wnd[2]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
    session.findById("wnd[2]/tbar[0]/btn[0]").press()
    session.findById("wnd[2]").close()

    layout_grid.selectedRows = "14"
    layout_grid.clickCurrentCell()

    # Exportar XXL
    main_grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    main_grid.setCurrentCell(5, "VBAK-AUDAT")
    main_grid.selectedRows = "5"
    main_grid.contextMenu()
    main_grid.selectContextMenuItem("&XXL")

    session.findById("wnd[1]/tbar[0]/btn[20]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = diretorio
    try:
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo
    except: pass 
    
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = len(diretorio)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    return caminho_completo

# ==========================================
# 5. FORMATAÇÃO DAS EXTRAÇÕES (EXCEL)
# ==========================================
def formatar_planilha_carteira(excel, caminho_arquivo_sap):
    print("Formatando planilha 'Em Carteira'...")
    excel.Calculation = XL_CALCULATION_MANUAL
    wb_sap = excel.Workbooks.Open(caminho_arquivo_sap)
    ws_sap = wb_sap.ActiveSheet
    
    try:
        ultima_linha = ws_sap.Cells(ws_sap.Rows.Count, 1).End(XL_UP).Row
        if ultima_linha >= 2:
            ws_sap.Range(f"A2:A{ultima_linha}").TextToColumns(Destination=ws_sap.Range("A2"), DataType=XL_DELIMITED)
            ws_sap.Range(f"C2:C{ultima_linha}").TextToColumns(Destination=ws_sap.Range("C2"), DataType=XL_DELIMITED)
            ws_sap.Range(f"E2:E{ultima_linha}").TextToColumns(Destination=ws_sap.Range("E2"), DataType=XL_DELIMITED)

            for col in ["G", "H", "I", "J", "K"]:
                rng = ws_sap.Range(f"{col}2:{col}{ultima_linha}")
                rng.TextToColumns(rng, XL_DELIMITED, XL_DOUBLE_QUOTE, False, False, False, False, False, False, "", None, ".", ",")

            ws_sap.AutoFilterMode = False
            ws_sap.Range(f"A1:L{ultima_linha}").AutoFilter(Field=12, Criteria1="Overweight Order")
            try:
                ws_sap.Range(f"A2:A{ultima_linha}").SpecialCells(XL_CELL_TYPE_VISIBLE).EntireRow.Delete()
            except Exception: pass
            ws_sap.AutoFilterMode = False
            
            ultima_linha = ws_sap.Cells(ws_sap.Rows.Count, 1).End(XL_UP).Row
            if ultima_linha >= 2:
                ws_sap.Range(f"N2:N{ultima_linha}").Formula = "=H2/100"
                ws_sap.Range(f"H2:H{ultima_linha}").Value = ws_sap.Range(f"N2:N{ultima_linha}").Value
                ws_sap.Range(f"N2:N{ultima_linha}").ClearContents() 

        return wb_sap, ws_sap
    finally:
        excel.Calculation = XL_CALCULATION_AUTOMATIC

def formatar_planilha_faturamento(excel, caminho_arquivo_sap):
    print("Formatando planilha 'Em Faturamento'...")
    excel.Calculation = XL_CALCULATION_MANUAL
    wb_sap = excel.Workbooks.Open(caminho_arquivo_sap)
    ws_sap = wb_sap.ActiveSheet
    
    try:
        ultima_linha = ws_sap.Cells(ws_sap.Rows.Count, 1).End(XL_UP).Row
        if ultima_linha >= 2:
            ws_sap.Range(f"A2:A{ultima_linha}").TextToColumns(Destination=ws_sap.Range("A2"), DataType=XL_DELIMITED)
            ws_sap.Range(f"C2:C{ultima_linha}").TextToColumns(Destination=ws_sap.Range("C2"), DataType=XL_DELIMITED)
            ws_sap.Range(f"E2:E{ultima_linha}").TextToColumns(Destination=ws_sap.Range("E2"), DataType=XL_DELIMITED)

            rng_g = ws_sap.Range(f"G2:G{ultima_linha}")
            rng_g.NumberFormat = "0"
            rng_g.Value = rng_g.Value 
            
            rng_h = ws_sap.Range(f"H2:H{ultima_linha}")
            rng_h.NumberFormatLocal = "0,00"
            rng_h.Value = rng_h.Value 

        return wb_sap, ws_sap
    finally:
        excel.Calculation = XL_CALCULATION_AUTOMATIC

def formatar_planilha_faturado_mes_atual(excel, caminho_arquivo_sap):
    print("Formatando planilha 'Faturado Mês Atual'...")
    excel.Calculation = XL_CALCULATION_MANUAL
    wb_sap = excel.Workbooks.Open(caminho_arquivo_sap)
    ws_sap = wb_sap.ActiveSheet
    
    try:
        ultima_linha = ws_sap.Cells(ws_sap.Rows.Count, 1).End(XL_UP).Row
        if ultima_linha >= 2:
            ws_sap.Range(f"A1:A{ultima_linha}").TextToColumns(Destination=ws_sap.Range("A1"), DataType=XL_DELIMITED)
            ws_sap.Range(f"C1:C{ultima_linha}").TextToColumns(Destination=ws_sap.Range("C1"), DataType=XL_DELIMITED)
            ws_sap.Range(f"D1:D{ultima_linha}").TextToColumns(Destination=ws_sap.Range("D1"), DataType=XL_DELIMITED)
            
            rng_f = ws_sap.Range(f"F2:F{ultima_linha}")
            rng_f.NumberFormat = "0"
            rng_f.Value = rng_f.Value

            rng_g = ws_sap.Range(f"G2:G{ultima_linha}")
            rng_g.NumberFormatLocal = "0,00"
            rng_g.Value = rng_g.Value

            rng_h = ws_sap.Range(f"H2:H{ultima_linha}")
            rng_h.NumberFormatLocal = "0,00"
            rng_h.Value = rng_h.Value

        return wb_sap, ws_sap
    finally:
        excel.Calculation = XL_CALCULATION_AUTOMATIC

def formatar_planilha_atp(excel, caminho_arquivo_sap):
    print("Formatando planilha 'ATP'...")
    excel.Calculation = XL_CALCULATION_MANUAL
    wb_sap = excel.Workbooks.Open(caminho_arquivo_sap)
    ws_sap = wb_sap.ActiveSheet
    
    try:
        # Text to columns nas colunas C, E e G
        for col in ["C", "E", "G"]:
            coluna_origem = ws_sap.Columns(f"{col}:{col}")
            celula_destino = ws_sap.Range(f"{col}1")
            coluna_origem.TextToColumns(Destination=celula_destino, DataType=XL_DELIMITED)
            
        return wb_sap, ws_sap
    finally:
        excel.Calculation = XL_CALCULATION_AUTOMATIC

# ==========================================
# 6. CONSOLIDAÇÃO DOS DADOS
# ==========================================
def consolidar_carteira(wb_inicial, wb_sap, ws_sap):
    print("Consolidando 'Em Carteira'...")
    ws_item = wb_inicial.Sheets("Item a item")
    ultima_linha_sap = ws_sap.Cells(ws_sap.Rows.Count, 1).End(XL_UP).Row
    if ultima_linha_sap < 2: 
        wb_sap.Close(SaveChanges=False)
        return

    linha_inicio = ws_item.Cells(ws_item.Rows.Count, 6).End(XL_UP).Row + 1
    linha_fim = linha_inicio + (ultima_linha_sap - 2)

    ws_item.Range(f"F{linha_inicio}:P{linha_fim}").Value = ws_sap.Range(f"A2:K{ultima_linha_sap}").Value
    ws_item.Range(f"O{linha_inicio}:R{linha_fim}").Value = ws_sap.Range(f"I2:L{ultima_linha_sap}").Value
    ws_item.Range(f"N{linha_inicio}:N{linha_fim}").Formula = f"=L{linha_inicio}*O{linha_inicio}"
    ws_item.Range(f"N{linha_inicio}:N{linha_fim}").Value = ws_item.Range(f"N{linha_inicio}:N{linha_fim}").Value
    
    wb_sap.Close(SaveChanges=False)

def consolidar_faturamento(wb_inicial, wb_sap, ws_sap):
    print("Consolidando 'Em Faturamento'...")
    ws_item = wb_inicial.Sheets("Item a item")
    ultima_linha_sap = ws_sap.Cells(ws_sap.Rows.Count, 1).End(XL_UP).Row
    if ultima_linha_sap < 2: 
        wb_sap.Close(SaveChanges=False)
        return

    ultima_coluna_sap = ws_sap.Cells(1, ws_sap.Columns.Count).End(XL_TO_LEFT).Column
    linha_inicio = ws_item.Cells(ws_item.Rows.Count, 6).End(XL_UP).Row + 1
    
    range_origem = ws_sap.Range(ws_sap.Cells(2, 1), ws_sap.Cells(ultima_linha_sap, ultima_coluna_sap))
    range_destino = ws_item.Range(ws_item.Cells(linha_inicio, 6), ws_item.Cells(linha_inicio + ultima_linha_sap - 2, 6 + ultima_coluna_sap - 1))
    
    range_destino.Value = range_origem.Value
    wb_sap.Close(SaveChanges=False)

def consolidar_faturado_mes_atual(wb_inicial, wb_sap, ws_sap):
    print("Consolidando 'Faturado Mês Atual' com VLOOKUP...")
    ws_item = wb_inicial.Sheets("Item a item")
    ultima_linha_sap = ws_sap.Cells(ws_sap.Rows.Count, 1).End(XL_UP).Row
    if ultima_linha_sap < 2: 
        wb_sap.Close(SaveChanges=False)
        return

    linha_inicio = ws_item.Cells(ws_item.Rows.Count, 6).End(XL_UP).Row + 1
    linha_fim = linha_inicio + (ultima_linha_sap - 2)

    ws_item.Range(f"F{linha_inicio}:H{linha_fim}").Value = ws_sap.Range(f"A2:C{ultima_linha_sap}").Value
    ws_item.Range(f"J{linha_inicio}:N{linha_fim}").Value = ws_sap.Range(f"D2:H{ultima_linha_sap}").Value
    
    # Prevenção: Copiar e Colar Valores nativo do Excel para evitar erro do '#N/D' no win32com
    rng_formulas = ws_item.Range(f"I{linha_inicio}:I{linha_fim}")
    rng_formulas.Formula = f"=VLOOKUP(H{linha_inicio},'Base OT - Consolidados'!D:F,3,0)"
    
    rng_formulas.Copy()
    rng_formulas.PasteSpecial(Paste=XL_PASTE_VALUES)
    wb_inicial.Application.CutCopyMode = False 
    
    wb_sap.Close(SaveChanges=False)

def consolidar_atp(wb_inicial, wb_sap, ws_sap):
    print("Consolidando 'ATP'...")
    ws_atp = wb_inicial.Sheets("ATP")
    
    # 1. Desoculta a aba
    ws_atp.Visible = -1 
    
    # 2. Apaga informações da linha 3 para baixo (mantém a linha 1 de cabeçalho e a linha 2 com as fórmulas)
    ultima_linha_atp = ws_atp.Cells(ws_atp.Rows.Count, 1).End(XL_UP).Row
    if ultima_linha_atp >= 3:
        # Deleta as linhas antigas da tabela (da 3 em diante)
        ws_atp.Rows(f"3:{ultima_linha_atp}").Delete()
        
    # 3. Copia e Cola os dados a partir da A2
    ultima_linha_sap = ws_sap.Cells(ws_sap.Rows.Count, 1).End(XL_UP).Row
    ultima_coluna_sap = ws_sap.Cells(1, ws_sap.Columns.Count).End(XL_TO_LEFT).Column
    
    if ultima_linha_sap >= 2:
        # Pega os dados do SAP (sem o cabeçalho)
        range_origem = ws_sap.Range(ws_sap.Cells(2, 1), ws_sap.Cells(ultima_linha_sap, ultima_coluna_sap))
        
        # Cola no destino a partir da linha 2
        # As colunas P e Q ficarão intactas e a tabela arrastará as fórmulas automaticamente
        range_destino = ws_atp.Range(ws_atp.Cells(2, 1), ws_atp.Cells(ultima_linha_sap, ultima_coluna_sap))
        range_destino.Value = range_origem.Value
        
    wb_sap.Close(SaveChanges=False)
    
# ==========================================
# 7. CORREÇÃO INVOICE E FÓRMULA COLUNA Q
# ==========================================
def corrigir_invoice_em_faturamento(wb_inicial):
    print("\nIniciando correção de Invoice Amt...")
    
    try:
        ws_correcao = wb_inicial.Sheets("Correção Invoice Amt - Em fat")
        ws_correcao.Visible = -1 
        print("-> Atualizando conexões da aba 'Correção Invoice Amt - Em fat'...")
        
        for lo in ws_correcao.ListObjects:
            try: lo.QueryTable.Refresh(BackgroundQuery=False)
            except: pass
        for qt in ws_correcao.QueryTables:
            try: qt.Refresh(BackgroundQuery=False)
            except: pass
                
        ws_correcao.Visible = 0 
    except Exception as e:
        print(f"-> Aviso: Erro na aba de correção: {e}")

    ws_base = wb_inicial.Sheets("Base OT - Consolidados")
    ws_base.AutoFilterMode = False
    
    ultima_linha = ws_base.Cells(ws_base.Rows.Count, 1).End(XL_UP).Row
    if ultima_linha >= 3:
        print("-> Filtrando 'EM FATURAMENTO' (Col AJ) e aplicando fórmula da Q1...")
        ws_base.Range(f"A2:AQ{ultima_linha}").AutoFilter(Field=36, Criteria1="EM FATURAMENTO")
        
        try:
            intervalo_visivel = ws_base.Range(f"Q3:Q{ultima_linha}").SpecialCells(XL_CELL_TYPE_VISIBLE)
            
            formula_q1 = ws_base.Range("Q1").FormulaR1C1
            intervalo_visivel.FormulaR1C1 = formula_q1
            
            for area in intervalo_visivel.Areas:
                area.Copy()
                area.PasteSpecial(Paste=XL_PASTE_VALUES)
                
            wb_inicial.Application.CutCopyMode = False 
            print("-> Fórmula da coluna Q aplicada e convertida em valores com sucesso!")
        except Exception as e:
            print(f"-> Aviso: Erro na coluna Q: {e}")
            
    ws_base.AutoFilterMode = False

# ==========================================
# 8. ARRASTAR FÓRMULAS FINAIS (DUPLO CLIQUE)
# ==========================================
def arrastar_formulas(wb_inicial):
    print("\nAplicando fórmulas (A:E e S:AA) até a última linha preenchida...")
    ws_item = wb_inicial.Sheets("Item a item")
    
    ultima_linha_f = ws_item.Cells(ws_item.Rows.Count, 6).End(XL_UP).Row
    
    if ultima_linha_f > 2:
        origem_ae = ws_item.Range("A2:E2")
        destino_ae = ws_item.Range(f"A2:E{ultima_linha_f}")
        origem_ae.AutoFill(Destination=destino_ae)
        
        origem_saa = ws_item.Range("S2:AA2")
        destino_saa = ws_item.Range(f"S2:AA{ultima_linha_f}")
        origem_saa.AutoFill(Destination=destino_saa)
        
    print("-> Fórmulas finais aplicadas com sucesso!")

# ==========================================
# 9. MOTOR PRINCIPAL (FLUXO)
# ==========================================
if __name__ == "__main__":
    dir_carteira = r"C:\Users\Joao Cortez\Desktop\Locais\EmCarteira"
    dir_faturamento = r"C:\Users\Joao Cortez\Desktop\Locais\EmFaturamento"
    dir_faturado_mes = r"C:\Users\Joao Cortez\Desktop\Locais\FaturadoMesAtual"
    dir_atp = r"C:\Users\Joao Cortez\Desktop\Locais\ATP"
    
    excel, wb_inicial = conectar_e_limpar_excel()
    
    if wb_inicial:
        sessao_sap = None
        
        try:
            # --- FLUXO 1: EM CARTEIRA ---
            if not sessao_sap: sessao_sap = conectar_sap()
            
            if copiar_dados_por_status(wb_inicial, "EM CARTEIRA", coluna_copia="D"):
                arq_carteira = extrair_sap_carteira(sessao_sap, dir_carteira)
                time.sleep(5)
                wb_sap_cart, ws_sap_cart = formatar_planilha_carteira(excel, arq_carteira)
                consolidar_carteira(wb_inicial, wb_sap_cart, ws_sap_cart)
                
            # --- FLUXO 2: EM FATURAMENTO ---
            if copiar_dados_por_status(wb_inicial, "EM FATURAMENTO", coluna_copia="D"):
                arq_fat = extrair_sap_faturamento(sessao_sap, dir_faturamento)
                time.sleep(5)
                wb_sap_fat, ws_sap_fat = formatar_planilha_faturamento(excel, arq_fat)
                consolidar_faturamento(wb_inicial, wb_sap_fat, ws_sap_fat)

            # --- FLUXO 3: FATURADO MÊS ATUAL ---
            status_mes = obter_filtro_mes_atual()
            if copiar_dados_por_status(wb_inicial, status_mes, coluna_copia="M"):
                arq_fat_mes = extrair_sap_faturado_mes_atual(sessao_sap, dir_faturado_mes)
                time.sleep(5)
                wb_sap_fat_mes, ws_sap_fat_mes = formatar_planilha_faturado_mes_atual(excel, arq_fat_mes)
                consolidar_faturado_mes_atual(wb_inicial, wb_sap_fat_mes, ws_sap_fat_mes)

            # --- FLUXO 4: EXTRAÇÃO ATP ---
            # Reutiliza o filtro "EM CARTEIRA" (Coluna D) para a seleção múltipla do ATP
            if copiar_dados_por_status(wb_inicial, "EM CARTEIRA", coluna_copia="D"):
                arq_atp = extrair_sap_atp(sessao_sap, dir_atp)
                time.sleep(5)
                wb_sap_atp, ws_sap_atp = formatar_planilha_atp(excel, arq_atp)
                consolidar_atp(wb_inicial, wb_sap_atp, ws_sap_atp)

            # --- FLUXO 5: CORREÇÃO INVOICE AMT E COLUNA Q ---
            corrigir_invoice_em_faturamento(wb_inicial)

            # --- FLUXO 6: PREENCHER FÓRMULAS FINAIS ---
            arrastar_formulas(wb_inicial)

        finally:
            # Garante que o Excel volte a ser visível/atualizado
            excel.ScreenUpdating = True
            excel.DisplayAlerts = True
            
            # --- FINALIZAÇÃO: VOLTAR PARA A PRIMEIRA ABA E ATUALIZAR TUDO (DUPLO REFRESH) ---
            try:
                print("\nVoltando para a primeira aba e executando Atualização de Dados (Refresh All)...")
                wb_inicial.Sheets(1).Activate() # Volta para a primeira aba
                wb_inicial.RefreshAll()         # 1º Refresh All
                time.sleep(3)                   # Breve pausa para o Excel processar tabelas/conexões em cascata
                wb_inicial.RefreshAll()         # 2º Refresh All
                print("Duplo Refresh All executado com sucesso.")
            except Exception as e:
                print(f"Aviso ao tentar atualizar as conexões do Excel: {e}")
            
        print("\nAUTOMAÇÃO FINALIZADA! Todos os processos foram executados, consolidados e os arquivos extras foram fechados.")
