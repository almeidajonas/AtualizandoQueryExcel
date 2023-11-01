import win32com.client
import time


# **Variaveis**
caminho = "Caminho\\Arquivo"
nomeArquivo = "NomeExcel.xlsx"


# **Abrindo Excel**
#Criar uma instância do Excel
File = win32com.client.Dispatch("Excel.Application")

# Definir a visibilidade do Excel (0 para execução em segundo plano, 1 para abrir na máquina)
File.Visible = 1

# Abrir o arquivo do Excel
Workbook = File.Workbooks.open(caminho+nomeArquivo)

#Atualiza a query do excel
Workbook.RefreshAll()
#Espera o arquivo terminar de atualizar
File.CalculateUntilAsyncQueriesDone()
time.sleep(10)

Workbook.Save()
File.quit()

