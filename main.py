import os
from pickle import GLOBAL
import re
import subprocess
import sys
from pathlib import Path
import pandas as pd
import pdfplumber
import shutil

targetSetor = "AR1"

NETWORK_DIR = r"\\profs01\documentos\PROCERGS\Relatorios_PPR"
NETWORK_DIR_DEDODURO = r"\Apropriação de Horas"
FILE_PATTERN_DEDODURO = "DedoDuro*.xlsx"
SHEET_NAME_DEDODURO = "H.Apropriadas X H.Ponto - Setor"
LIMITE_MIN_DEDODURO = 99
LIMITE_MAX_DEDODURO = 101

NETWORK_DIR_ENTREGAVEIS = r"\Entregaveis"
FILE_PATTERN_ENTREGAVEIS = "Previa_Entregaveis*.xlsx"
SHEET_NAME_ENTREGAVEIS_1 = "DFR - Projeto"
SHEET_NAME_ENTREGAVEIS_2 = "DFR - Manutenção"
TARGET_APROVADO_ENTREGAVEIS = "PENDENTE"

NETWORK_DIR_ITAD = r"\ITAD"
FILE_PATTERN_ITAD = "Previa_ITAD*.xlsx"
SHEET_NAME_ITAD = "Não Conformidades"

NETWORK_DIR_RPM = r"\Validacao_RPM"
FILE_PATTERN_RPM = "InconsistenciasRPM_Setores.pdf"

retRedmine = set()

def find_file_more_recent(directory: str, pattern: str) -> Path | None:
    """Recursively search for files matching ``pattern`` under ``directory``.

    When multiple candidates exist the most recently modified file is returned.
    Returns a :class:`pathlib.Path` or ``None`` if nothing found.
    """
    root = Path(directory)
    if not root.exists():
        printGrava(f"Diretório não existe: {directory}")
        sys.exit(1)

    candidates: list[Path] = []
    # use rglob to recursively search; pattern may include wildcards
    for path in root.rglob(pattern):
        if path.is_file():
            candidates.append(path)

    if not candidates:
        printGrava(f"Nenhum arquivo encontrado com {pattern} em {directory}")
        sys.exit(1)

    # pick the newest modified file
    newest = max(candidates, key=lambda p: p.stat().st_mtime)
    return newest


def read_and_filter_dedoduro(file_path: Path, coluna: int) -> pd.DataFrame:
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    # read the specified sheet into a DataFrame, then filter by the "Setor" column. Column names start in line 5, so we skip the first 4 rows
    df = pd.read_excel(file_path, sheet_name=SHEET_NAME_DEDODURO, header=4)
    df.rename(columns={df.columns[0]: "Divisão"}, inplace=True)
    df.rename(columns={df.columns[1]: "Setor"}, inplace=True)
    df.rename(columns={df.columns[2]: "Matricula"}, inplace=True)
    df.rename(columns={df.columns[3]: "Nome"}, inplace=True)

    df = df[df["Setor"] == targetSetor]
    if coluna == 1:
        df.rename(columns={df.columns[4]: "Normais_PES"}, inplace=True)
        df.rename(columns={df.columns[5]: "Normais_RPM"}, inplace=True)
        df.rename(columns={df.columns[6]: "Normais_Perc"}, inplace=True)
        df = df.drop(df.columns[7:13], axis=1)

        df = df[(df["Normais_Perc"] <= LIMITE_MIN_DEDODURO) | (df["Normais_Perc"] >= LIMITE_MAX_DEDODURO)]
       
        df["Normais_PES"] = df["Normais_PES"].round(2)
        df["Normais_RPM"] = df["Normais_RPM"].round(2)
        df["Normais_Perc"] = df["Normais_Perc"].round(2)

    elif coluna == 2:
        df.rename(columns={df.columns[7]: "Extras_PES"}, inplace=True)
        df.rename(columns={df.columns[8]: "Extras_RPM"}, inplace=True)
        df.rename(columns={df.columns[9]: "Extras_Perc"}, inplace=True)
        df = df.drop(df.columns[10:13], axis=1)
        df = df.drop(df.columns[4:7], axis=1)

        df = df[(df["Extras_Perc"] <= LIMITE_MIN_DEDODURO) | (df["Extras_Perc"] >= LIMITE_MAX_DEDODURO)]
        
        df["Extras_PES"] = df["Extras_PES"].round(2)
        df["Extras_RPM"] = df["Extras_RPM"].round(2)
        df["Extras_Perc"] = df["Extras_Perc"].round(2)
    elif coluna == 3:
        df.rename(columns={df.columns[10]: "BIP_PES"}, inplace=True)
        df.rename(columns={df.columns[11]: "BIP_RPM"}, inplace=True)
        df.rename(columns={df.columns[12]: "BIP_Perc"}, inplace=True)
        df = df.drop(df.columns[4:10], axis=1)

        df = df[(df["BIP_Perc"] <= LIMITE_MIN_DEDODURO) | (df["BIP_Perc"] >= LIMITE_MAX_DEDODURO)]
        
        df["BIP_PES"] = df["BIP_PES"].round(2)  
        df["BIP_RPM"] = df["BIP_RPM"].round(2)
        df["BIP_Perc"] = df["BIP_Perc"].round(2)

    return df

def printGrava(texto: str, modo: str = "a"):
    print(texto)
    with open(f"AR.txt", modo, encoding="utf-8") as f:
        f.write(texto + "\n")

def Show_DedoDuro(excel_file: Path | None = None):
    MostraArquivo = True

    for x in [1, 2, 3]:
        try:
            result = read_and_filter_dedoduro(excel_file, x)
        except Exception as exc:  # pylint: disable=broad-except        
            return
        
        if not result.empty:
            if MostraArquivo:
                MostraArquivo = False
                printGrava(f"Arquivo {excel_file}")

            printGrava("```")
            printGrava(result.to_string(index=False))
            printGrava("```")

    if not MostraArquivo:
        printGrava(f"Verifique os dados acima. O Percentuais < {LIMITE_MIN_DEDODURO} ou > {LIMITE_MAX_DEDODURO} estão fora do limite.")
        printGrava(f"==> A correção provável é ajustar as horas no RPM.")
        Separador()

    
def read_and_filter_PreviaEntregaveis(file_path: Path, sheet_name: str, tipo: str) -> pd.DataFrame:
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    if tipo == 'Pendente:':
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=2)
        df = df[df["Aprovado Alteração"] == TARGET_APROVADO_ENTREGAVEIS]
    else:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        header_row = df[df.iloc[:, 0] == "Épicos Inválidos"].index[0]
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row + 2) 

    filtered = df[df["Setor"] == targetSetor]
    filtered = filtered.drop(columns=["Tipo Entregável"], errors='ignore')
    filtered = filtered.drop(columns=["Target Date"], errors='ignore')
    filtered = filtered.drop(columns=["Data Entrega Real"], errors='ignore')
    filtered = filtered.drop(columns=["IAP"], errors='ignore')
    return filtered

def Show_PreviaEntregaveis(excel_file: Path | None = None):
    MostraArquivo = True

    for sheet in [SHEET_NAME_ENTREGAVEIS_1, SHEET_NAME_ENTREGAVEIS_2]:
        for tipo in ['Pendente:', 'Inválido:']:
            try:
                result = read_and_filter_PreviaEntregaveis(excel_file, sheet, tipo)
            except Exception as exc:  # pylint: disable=broad-except
                continue
        
            if result.empty:
                continue
            else:
                if MostraArquivo:
                    MostraArquivo = False
                    printGrava(f"Arquivo {excel_file}")
                printGrava(f"Planilha {sheet} - Épico {tipo}")

                if tipo == 'Pendente:':
                    # Altera result para a coluna ID ser o conteúdo atual começando com o símbolo #
                    result["ID"] = result["ID"].astype(str).apply(lambda x: f"#{x}" if not x.startswith("#") else x)                  

                printGrava(result.to_string(index=False))

                if tipo == 'Pendente:':
                    printGrava(f"Verifique os dados acima. Os Épicos com status 'PENDENTE' estão aguardando aprovação.")
                    printGrava(f"==> A correção provável é solicitar, no próprio Épico, a aprovação da Chefia, que deve ser citada com @nome.")
                else: # tipo == "Inválido"                
                    pass

            Separador()

def read_and_filter_ITAD(file_path: Path) -> pd.DataFrame:
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    df = pd.read_excel(file_path, sheet_name=SHEET_NAME_ITAD)
    filtered = df[df["TeamProject"] == "DFR." + targetSetor]

    return filtered

def abreRedmine(redmines: set):
    if redmines:
        printGrava(f"Redmine ainda definido com Rascunho:")
        for redmine in redmines:
            url = f"https://redmine.intra.rs.gov.br/issues/{redmine}"
            printGrava(url)
            # webbrowser.open(url)
        Separador()

def Show_ITAD(excel_file: Path | None = None):
    MostraArquivo = True
    try:
        result = read_and_filter_ITAD(excel_file)
    except Exception as exc:  # pylint: disable=broad-except
        return

    if result.empty:
        return
    else:
        # Cria variável como SET para evitar duplicidade de Redmine, e lista para os demais casos
        for line in result.itertuples(index=False):
            if MostraArquivo:
                MostraArquivo = False
                printGrava(f"Arquivo {excel_file} - {SHEET_NAME_ITAD} - Buscando Não Conformidades para {"DFR." + targetSetor}")

            printGrava(line.Mensagem)
            Separador()
            if "Demanda em situação de Rascunho" in line.Mensagem:
                #busca o número do Redmine, considerando que o formato é "[DemandaId=12345]"                
                match = re.search(r'\[DemandaId=(\d+)\]', line.Mensagem)
                if match:
                    retRedmine.add(match.group(1))
    
    if not MostraArquivo: # Se apresentou erro
        Separador()

def Show_RPM(pdf_path: Path | None = None):
    """Lê e processa o PDF de inconsistências RPM por setor."""
    MostraArquivo = True

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if targetSetor in text:
                    if MostraArquivo:
                        MostraArquivo = False
                        printGrava(f"Arquivo {pdf_path}")
                    Separador()
                    printGrava(f"Texto da página {page.page_number}:")  
                    for line in text.splitlines()[3:-1]: # Ignora as linhas iniciais e a última linha que contém o rodapé
                        printGrava(f"{line}")
                    Separador()
    except Exception as exc:
        printGrava(f"Erro ao processar PDF: {exc}")


def Separador():
    printGrava("")
    printGrava("_" * 3)

def atualizaWiki():
    destino = r"C:\Users\rb65847\Source\repos\DFR.AR1.wiki"
    arquivo_destino = destino + r"\DFR.AR1\Relatório-TFS%2DPES%2DRPM.md"
    shutil.copy("AR.txt", arquivo_destino)



    os.chdir(destino) 
    subprocess.run(["git", "pull"], check=True)
    subprocess.run(["git", "add", "*"], check=True)

    data_atual = pd.Timestamp.now().strftime("%d-%m-%Y")
    mensagem = f"Atualizado em {data_atual}"
    subprocess.run(["git", "commit", "-m", mensagem], check=True)

    subprocess.run(["git", "push"], check=True)


def main():
    global targetSetor

    excel_file_dedoduro = find_file_more_recent(NETWORK_DIR + NETWORK_DIR_DEDODURO, FILE_PATTERN_DEDODURO)
    excel_file_previa = find_file_more_recent(NETWORK_DIR + NETWORK_DIR_ENTREGAVEIS, FILE_PATTERN_ENTREGAVEIS)
    excel_file_itad = find_file_more_recent(NETWORK_DIR + NETWORK_DIR_ITAD, FILE_PATTERN_ITAD)

    pdf_path_rpm = find_file_more_recent(NETWORK_DIR + NETWORK_DIR_RPM, FILE_PATTERN_RPM)

    printGrava(f"**Problemas encontrados nos Arquivos {NETWORK_DIR}**", modo="w")
    printGrava(f"Gerado em {pd.Timestamp.now().strftime("%d/%m/%Y %H:%M:%S")}")
    
    Separador()

    for busca in ['AR1', 'AR2', 'AR3', 'AR4', 'AR5']:
        targetSetor = busca
        retRedmine.clear()
        
        # Receber parâmetros da chamada do programa para definir o targetSetor
        printGrava(f"**Pendências para o setor {targetSetor}**")
        Separador()
        Show_RPM(pdf_path_rpm)
        Show_DedoDuro(excel_file_dedoduro)
        Show_PreviaEntregaveis(excel_file_previa)
        Show_ITAD(excel_file_itad)

        abreRedmine(retRedmine)

    print(f"Arquivos gerados: AR.txt")

    atualizaWiki()

if __name__ == "__main__":
    main()

