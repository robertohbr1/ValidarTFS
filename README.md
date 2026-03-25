# monitortfs

This project includes a simple routine to locate and read an Excel workbook from a network share,\
filtering records by a specific sector value.

## Usage

1. Ensure dependencies are installed (pandas, openpyxl) via pip install -r requirements.txt or\
   the project wheel/installer.
2. Run the script with python main.py.
3. The code searches for the first file matching DedoDuro*.xlsx in the\
   \\profs01\\documentos\\PROCERGS\\Relatorios_PPR\\Apropria��o de Horas directory and opens the\
   worksheet named H.Apropriadas X H.Ponto - Setor.
4. Rows where the Setor column equals AR1, AR2, AR3, AR4 and AR5 are printed to the console.
