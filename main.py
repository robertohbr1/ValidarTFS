import re
import sys
from pathlib import Path
import glob
import webbrowser

import pandas as pd
import pdfplumber

PERMITE_ABRIR_TASK = False
PERMITE_ABRIR_PBI = False
PERMITE_ABRIR_REDMINE = False

PERGUNTA_ABRIR_TASK = "Deseja abrir as Tasks? (s/n)"
PERGUNTA_ABRIR_PBI = "Deseja abrir os PBIs? (s/n)"
PERGUNTA_ABRIR_REDMINE = "Deseja abrir os Redmines? (s/n)"

TARGET_SETOR = "AR1"

NETWORK_DIR = r"\\profs01\documentos\PROCERGS\Relatorios_PPR"
NETWORK_DIR_DEDODURO = r"\Apropriação de Horas"
FILE_PATTERN_DEDODURO = "DedoDuro*.xlsx"
SHEET_NAME_DEDODURO = "H.Apropriadas X H.Ponto - Setor"
LIMITE_MIN_DEDODURO = 90
LIMITE_MAX_DEDODURO = 110

NETWORK_DIR_ENTREGAVEIS = r"\Entregaveis"
FILE_PATTERN_ENTREGAVEIS = "Previa_Entregaveis*.xlsx"
SHEET_NAME_ENTREGAVEIS_1 = "DFR - Projeto"
SHEET_NAME_ENTREGAVEIS_2 = "DFR - Manutenção"
TARGET_APROVADO_ENTREGAVEIS = "PENDENTE"

NETWORK_DIR_ITAD = r"\ITAD"
FILE_PATTERN_ITAD = "Previa_ITAD*.xlsx"
SHEET_NAME_ITAD = "Não Conformidades"
TARGET_ITAD = "DFR.AR1"

NETWORK_DIR_RPM = r"\Validacao_RPM"
FILE_PATTERN_RPM = "InconsistenciasRPM_Setores.pdf"
TARGET_RPM = "AR1"

def find_file_more_recent(directory: str, pattern: str) -> Path | None:
    """Recursively search for files matching ``pattern`` under ``directory``.

    When multiple candidates exist the most recently modified file is returned.
    Returns a :class:`pathlib.Path` or ``None`` if nothing found.
    """
    root = Path(directory)
    if not root.exists():
        return None

    candidates: list[Path] = []
    # use rglob to recursively search; pattern may include wildcards
    for path in root.rglob(pattern):
        if path.is_file():
            candidates.append(path)

    if not candidates:
        return None

    # pick the newest modified file
    newest = max(candidates, key=lambda p: p.stat().st_mtime)
    return newest


def read_and_filter_dedoduro(file_path: Path) -> pd.DataFrame:
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    # read the specified sheet into a DataFrame, then filter by the "Setor" column. Column names start in line 5, so we skip the first 4 rows
    df = pd.read_excel(file_path, sheet_name=SHEET_NAME_DEDODURO, header=4)
    # if df.iloc[3, 1] != "Setor":
    #     raise KeyError("'Setor' not found in expected position (row 4, column 2)")
    df.rename(columns={df.columns[0]: "Divisão"}, inplace=True)
    df.rename(columns={df.columns[1]: "Setor"}, inplace=True)
    df.rename(columns={df.columns[2]: "Matricula"}, inplace=True)
    df.rename(columns={df.columns[3]: "Nome"}, inplace=True)

    df.rename(columns={df.columns[4]: "Normais_PES"}, inplace=True)
    df.rename(columns={df.columns[5]: "Normais_RPM"}, inplace=True)
    df.rename(columns={df.columns[6]: "Normais_Perc"}, inplace=True)

    df.rename(columns={df.columns[7]: "Extras_PES"}, inplace=True)
    df.rename(columns={df.columns[8]: "Extras_RPM"}, inplace=True)
    df.rename(columns={df.columns[9]: "Extras_Perc"}, inplace=True)

    df.rename(columns={df.columns[10]: "BIP_PES"}, inplace=True)
    df.rename(columns={df.columns[11]: "BIP_RPM"}, inplace=True)
    df.rename(columns={df.columns[12]: "BIP_Perc"}, inplace=True)

    filtered = df[df["Setor"] == TARGET_SETOR]
    filtered = filtered[((filtered["Normais_Perc"].notna()) & ((filtered["Normais_Perc"] < LIMITE_MIN_DEDODURO) | (filtered["Normais_Perc"] > LIMITE_MAX_DEDODURO))) 
                        | ((filtered["Extras_Perc"].notna()) & ((filtered["Extras_Perc"] < LIMITE_MIN_DEDODURO) | (filtered["Extras_Perc"] > LIMITE_MAX_DEDODURO))) 
                        | ((filtered["BIP_Perc"].notna()) & ((filtered["BIP_Perc"] < LIMITE_MIN_DEDODURO) | (filtered["BIP_Perc"] > LIMITE_MAX_DEDODURO)))]

    filtered["Normais_PES"] = filtered["Normais_PES"].round(2)
    filtered["Normais_RPM"] = filtered["Normais_RPM"].round(2)
    filtered["Normais_Perc"] = filtered["Normais_Perc"].round(2)
    filtered["Extras_PES"] = filtered["Extras_PES"].round(2)
    filtered["Extras_RPM"] = filtered["Extras_RPM"].round(2)
    filtered["Extras_Perc"] = filtered["Extras_Perc"].round(2)
    filtered["BIP_PES"] = filtered["BIP_PES"].round(2)  
    filtered["BIP_RPM"] = filtered["BIP_RPM"].round(2)
    filtered["BIP_Perc"] = filtered["BIP_Perc"].round(2)
    return filtered

def Show_DedoDuro():
    excel_file = find_file_more_recent(NETWORK_DIR + NETWORK_DIR_DEDODURO, FILE_PATTERN_DEDODURO)
    if excel_file is None:
        print(f"No file matching {FILE_PATTERN_DEDODURO} found in {NETWORK_DIR + NETWORK_DIR_DEDODURO}")
        sys.exit(1)
    
    print(f"Abrindo {excel_file}")
    try:
        result = read_and_filter_dedoduro(excel_file)
    except Exception as exc:  # pylint: disable=broad-except        
        return
    
    if result.empty:
        return
    else:
        print(f"Encontrados {len(result)} registros para o setor {TARGET_SETOR}:")
        print(result.to_string(index=False))
        print(f"Verifique os dados acima. O Percentuais < {LIMITE_MIN_DEDODURO} ou > {LIMITE_MAX_DEDODURO} estão fora do limite.")
        print(f"==> A correção provável é ajustar as horas no RPM.")
    
def read_and_filter_PreviaEntregaveis(file_path: Path, sheet_name: str, tipo: str) -> pd.DataFrame:
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    if tipo == 'Válidos e Pendentes':
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=2)
        df = df[df["Aprovado Alteração"] == TARGET_APROVADO_ENTREGAVEIS]
    else:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        header_row = df[df.iloc[:, 0] == "Épicos Inválidos"].index[0]
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row + 2) 

    filtered = df[df["Setor"] == TARGET_SETOR]
    filtered = filtered.drop(columns=["Tipo Entregável"], errors='ignore')
    filtered = filtered.drop(columns=["Target Date"], errors='ignore')
    filtered = filtered.drop(columns=["Data Entrega Real"], errors='ignore')
    filtered = filtered.drop(columns=["IAP"], errors='ignore')
    return filtered

def Show_PreviaEntregaveis():
    excel_file = find_file_more_recent(NETWORK_DIR + NETWORK_DIR_ENTREGAVEIS, FILE_PATTERN_ENTREGAVEIS)
    if excel_file is None:
        print(f"No file matching {FILE_PATTERN_ENTREGAVEIS} found in {NETWORK_DIR + NETWORK_DIR_ENTREGAVEIS}")
        return

    print(f"Abrindo {excel_file}")

    for sheet in [SHEET_NAME_ENTREGAVEIS_1, SHEET_NAME_ENTREGAVEIS_2]:
        for tipo in ['Válidos e Pendentes', 'Inválidos']:
            print(f"Planilha {sheet} - Buscando Épicos {tipo}")
            try:
                result = read_and_filter_PreviaEntregaveis(excel_file, sheet, tipo)
            except Exception as exc:  # pylint: disable=broad-except
                continue
        
            if result.empty:
                continue
            else:
                print(f"Encontrados {len(result)} registros:")
                # display rows as a table with column names, one record per line
                print(result.to_string(index=False))
                if tipo == 'Válidos e Pendentes':
                    print(f"Verifique os dados acima. Os Épicos com status 'PENDENTE' estão aguardando aprovação.")
                    print(f"==> A correção provável é solicitar, no próprio Épico, a aprovação da Chefia, que deve ser citada com @nome.")


def read_and_filter_ITAD(file_path: Path) -> pd.DataFrame:
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    df = pd.read_excel(file_path, sheet_name=SHEET_NAME_ITAD)
    filtered = df[df["TeamProject"] == TARGET_ITAD]

    return filtered

def buscaPBI(mensagem: str) -> str | None:
    match = re.search(r'#(\d+)', mensagem)
    if match:
        return match.group(1)
    return None

def abreSite(pbis: set):
    for pbi in pbis:
        url = f"https://dev.azure.com/Procergs/DFR.AR1/_workitems/edit/{pbi}/"
        webbrowser.open(url)

def abreRedmine(redmines: set):
    for redmine in redmines:
        url = f"https://redmine.intra.rs.gov.br/issues/{redmine}"
        webbrowser.open(url)

def Show_ITAD():
    excel_file = find_file_more_recent(NETWORK_DIR + NETWORK_DIR_ITAD, FILE_PATTERN_ITAD)
    if excel_file is None:
        print(f"No file matching {FILE_PATTERN_ITAD} found in {NETWORK_DIR + NETWORK_DIR_ITAD}")
        sys.exit(1)

    print(f"Abrindo {excel_file} - {SHEET_NAME_ITAD} - Buscando Não Conformidades para {TARGET_ITAD}")
    try:
        result = read_and_filter_ITAD(excel_file)
    except Exception as exc:  # pylint: disable=broad-except
        return

    if result.empty:
        return
    else:
        # Cria variável como SET para evitar duplicidade de Redmine, e lista para os demais casos
        retRedmine = set()
        retPBI_RedmineErro = set()
        retPBI_RedmineVazio = set()
        retPBI_MultiRef = set()
        retTask = set()
        outrasMensagens = set()
        for line in result.itertuples(index=False):
            if "Task sem esforço" in line.Mensagem:
                retTask.add(f"{(line.Mensagem[1:8])}")
            elif "Demanda em situação de Rascunho" in line.Mensagem:
                #busca o número do Redmine, considerando que o formato é "[DemandaId=12345]"                
                match = re.search(r'\[DemandaId=(\d+)\]', line.Mensagem)
                if match:
                    retRedmine.add(match.group(1))
            elif "Erro na busca da demanda no Redmine" in line.Mensagem:
                retPBI_RedmineErro.add(buscaPBI(line.Mensagem))
            elif "sem sistema ou incorreto" in line.Mensagem:
                retPBI_RedmineVazio.add(buscaPBI(line.Mensagem))
            elif "PBI com multipla referência a Entregável" in line.Mensagem:
                retPBI_MultiRef.add(buscaPBI(line.Mensagem))
            else:
                outrasMensagens.add(f"{line.Mensagem}")

        if retRedmine:
            print(f"Redmine ainda definido com Rascunho: {retRedmine}")
            if perguntaAbrirRedmine() == 's':
                abreRedmine(retRedmine)
        if retPBI_RedmineErro:
            print(f"PBIs com erro na busca no Redmine: {retPBI_RedmineErro}")
            if perguntaAbrirPBI() == 's':
                abreSite(retPBI_RedmineErro)
        if retPBI_RedmineVazio:
            print(f"PBIs sem sistema ou com sistema incorreto: {retPBI_RedmineVazio}")
            if perguntaAbrirPBI() == 's':
                abreSite(retPBI_RedmineVazio)
        if retPBI_MultiRef:
            print(f"PBIs com múltipla referência a Entregável: {retPBI_MultiRef}")
            if perguntaAbrirPBI() == 's':
                abreSite(retPBI_MultiRef)
        if retTask:
            print(f"Tasks sem esforço registrado: {retTask}")
            if perguntaAbrirTask() == 's':
                abreSite(retTask)
        if outrasMensagens:
            print(f"Outras mensagens:")
            for line in outrasMensagens:
                print(f"   {line}")

def perguntaAbrirTask():
    if not PERMITE_ABRIR_TASK:
        return 'n'
    return input(PERGUNTA_ABRIR_TASK).lower()

def perguntaAbrirPBI():
    if not PERMITE_ABRIR_PBI:
        return 'n'
    return input(PERGUNTA_ABRIR_PBI).lower()

def perguntaAbrirRedmine():
    if not PERMITE_ABRIR_REDMINE:
        return 'n'
    return input(PERGUNTA_ABRIR_REDMINE).lower()

def Show_RPM():
    """Lê e processa o PDF de inconsistências RPM por setor."""
    pdf_path = find_file_more_recent(NETWORK_DIR + NETWORK_DIR_RPM, FILE_PATTERN_RPM)
    
    if not Path(pdf_path).exists():
        print(f"Arquivo PDF não encontrado: {pdf_path}")
        return
    
    print(f"Abrindo {pdf_path}")

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if TARGET_RPM in text:
                    print("=" * 80)
                    print(f"Texto da página {page.page_number}:")  
                    for line in text.splitlines()[3:-1]: # Ignora as linhas iniciais e a última linha que contém o rodapé
                        print(f"{line}")
    except Exception as exc:
        print(f"Erro ao processar PDF: {exc}")

def main():
    print(f"Verifica Pendências para o setor {TARGET_SETOR}")
    print("-" * 80)
    Show_DedoDuro()
    print("-" * 80)
    Show_PreviaEntregaveis()
    print("-" * 80)
    Show_RPM()
    print("-" * 80)
    Show_ITAD()
    print("-" * 80)

if __name__ == "__main__":
    main()

