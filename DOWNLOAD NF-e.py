import tkinter as tk
from tkinter import messagebox
import subprocess
import shlex
import time
import os
import win32com.client
import psutil
import shutil


# Obter o diretório do usuário atual
user_dir = os.path.expanduser("~")

# Criar o caminho dinâmico para a pasta SAP\SAP GUI
sap_gui_path = os.path.join(user_dir, "Documents", "SAP", "SAP GUI")



def fazer_login():
    close_process("saplogon.exe")
    caminho_executavel_sap = r"C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe"
    if not os.path.exists(caminho_executavel_sap):
        caminho_executavel_sap = (
            r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        )

    print(f"Caminho do executável: {caminho_executavel_sap}")

    if not os.path.exists(caminho_executavel_sap):
        print("O arquivo não foi encontrado. Verifique o caminho.")
        messagebox.showerror(
            "Erro", "O executável do SAP GUI não foi encontrado. Verifique o caminho."
        )
        return

    # Use shlex.split para tratar corretamente o caminho com espaços
    comando = shlex.split(f'"{caminho_executavel_sap}"')
    print(f"Comando para execução: {comando}")

    try:
        subprocess.Popen(comando)
    except Exception as e:
        print(f"Ocorreu um erro ao iniciar o SAP GUI: {e}")
        messagebox.showerror("Erro", f"Ocorreu um erro ao iniciar o SAP GUI: {e}")
        return

    time.sleep(3)

    usuario_sap = entry_usuario.get()
    senha_sap = entry_senha.get()

    try:
        subprocess.Popen(comando)
    except Exception as e:
        print(f"Ocorreu um erro ao iniciar o SAP GUI: {e}")

    time.sleep(3)
    sapguiauto = win32com.client.GetObject("SAPGUI")
    application = sapguiauto.GetScriptingEngine
    try:
        connection = application.OpenConnection("NOVO PVR - Produção SAP ECC", True)
    except Exception as e:
        print(
            f"Conexão 'PVR - Produção (Externo)' não encontrada. Tentando 'PVR - Produção (Interno)'..."
        )
        try:
            connection = application.OpenConnection("PVR - Produção SAP ECC", True)
        except Exception as e:
            print(f"Ocorreu um erro ao abrir a conexão: {e}")

    session = connection.Children(0)

    try:
        # Preencher as informações de cliente, usuário e senha
        session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "300"
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = usuario_sap
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = senha_sap
        session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "PT"
        session.findById("wnd[0]").sendVKey(0)

    except:
        messagebox.showerror("Erro", f"Usuário ou senha incorretos")
        close_process("saplogon.exe")
        return



# In[3]:


# FUNÇÃO ENCERRAR O SAP
def close_process(nome_processo):
    for proc in psutil.process_iter(['pid', 'name']):
        if nome_processo.lower() in proc.info['name'].lower():
            try:
                processo = psutil.Process(proc.info['pid'])
                processo.terminate()  # ou processo.kill() para forçar o fechamento
                print(f'{proc.info["name"]} (PID: {proc.info["pid"]}) foi fechado.')
            except psutil.NoSuchProcess:
                print(f'Erro: Processo {proc.info["name"]} não encontrado.')
            except psutil.AccessDenied:
                print(f'Erro: Acesso negado para fechar {proc.info["name"]}.')

def processar_arquivos(pasta_origem, pasta_destino):
    # Verifica se as pastas existem
    if not os.path.exists(pasta_origem):
        print(f"A pasta de origem '{pasta_origem}' não foi encontrada.")
        return

    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)
        print(f"A pasta de destino '{pasta_destino}' não existia, então foi criada.")

    # Itera sobre os arquivos na pasta de origem
    for arquivo in os.listdir(pasta_origem):
        caminho_arquivo_origem = os.path.join(pasta_origem, arquivo)

        # Verifica se é um arquivo
        if os.path.isfile(caminho_arquivo_origem):
            caminho_arquivo_destino = os.path.join(pasta_destino, arquivo)

            # Se o arquivo já existir no destino, remove o arquivo antigo
            if os.path.exists(caminho_arquivo_destino):
                os.remove(caminho_arquivo_destino)
                print(f"Arquivo {arquivo} existente no destino removido.")

            # Move o arquivo para a pasta de destino
            shutil.move(caminho_arquivo_origem, caminho_arquivo_destino)
            print(f"Arquivo {arquivo} movido para {pasta_destino}")


def executar_rotina():
    # Armazena informações
    # codigos_material = text_codigos_material.get("1.0", tk.END).strip()  # Obtém todos os códigos do material
    close_process("saplogon.exe")
    local  = entry_local.get()
    data0 = entry_data0.get()
    data1 = entry_data1.get()
    pastad  = entry_pastad.get()
    
    # Fazer login no SAP
    fazer_login()

    #INFORMAÇÕES SAP
    sapguiauto = win32com.client.GetObject("SAPGUI")
    application = sapguiauto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)


    session.findById("wnd[0]/tbar[0]/okcd").text = "/nzfi016"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtS_BUPLA-LOW").text = local
    session.findById("wnd[0]/usr/ctxtS_BUPLA-HIGH").text = local
    session.findById("wnd[0]/usr/ctxtS_BUDAT-LOW").text = data0
    session.findById("wnd[0]/usr/ctxtS_BUDAT-HIGH").text = data1
    session.findById("wnd[0]/usr/ctxtP_VARIA").text = "/ISSJOI"
    session.findById("wnd[0]/usr/ctxtS_BUDAT-HIGH").setFocus
    session.findById("wnd[0]/usr/ctxtS_BUDAT-HIGH").caretPosition = 6
    session.findById("wnd[0]").sendVKey(8)


    def realizar_exportacao(session, sap_gui_path):
        session.findById("wnd[0]/titl/shellcont/shell").pressContextButton("%GOS_TOOLBOX")
        session.findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem("%GOS_VIEW_ATTA")

        session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").pressToolbarButton("&MB_FILTER")
        session.findById("wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell").selectedRows = "0"
        session.findById("wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell").doubleClickCurrentCell()
        session.findById("wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/btn600_BUTTON").press()
        session.findById("wnd[3]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "@IT\QADOBE ACROBAT READER@"
        session.findById("wnd[3]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-HIGH").text = ""
        session.findById("wnd[3]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 26
        session.findById("wnd[3]").sendVKey(0)

        session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").currentCellColumn = "BITM_DESCR"
        session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").pressToolbarButton("%ATTA_EXPORT")
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sap_gui_path
        session.findById("wnd[1]").sendVKey(11)
  
        session.findById("wnd[1]").sendVKey(12)
        session.findById("wnd[0]").sendVKey(3)




    row_index = 0
    
    # Tenta descobrir o número total de linhas do grid
    try:
        total_rows = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").rowCount
    except Exception as e:
        print(f"Erro ao tentar obter número total de linhas: {e}")
        total_rows = 0
    
    while row_index < total_rows:
        print(f"\n🔄 Processando linha {row_index}...")

        try:
            grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")
            grid.selectedRows = str(row_index)
            grid.currentCellRow = row_index

            tentativa_sucesso = False

            # Tentativa com DOCNUM
            try:
                print("➡️ Tentando com 'DOCNUM'")
                grid.currentCellColumn = "DOCNUM"
                grid.clickCurrentCell()
                realizar_exportacao(session, sap_gui_path)
                tentativa_sucesso = True
            except Exception as e_doc:
                print(f"❌ Falha com 'DOCNUM': {e_doc}")

            if tentativa_sucesso:
                try:
                    processar_arquivos(sap_gui_path, pastad)
                    print(f"✅ Sucesso com 'DOCNUM' na linha {row_index}")
                except Exception as e_proc:
                    print(f"⚠️ Sucesso parcial com 'DOCNUM', mas falhou no pós-processamento: {e_proc}")
                row_index += 1
                continue  # pula pra próxima linha

            # Se falhou com DOCNUM, tenta com BELNR
            try:
                print("➡️ Tentando com 'BELNR'")
                grid.currentCellColumn = "BELNR"
                grid.clickCurrentCell()
                realizar_exportacao(session, sap_gui_path)
                processar_arquivos(sap_gui_path, pastad)
                print(f"✅ Sucesso com 'BELNR' na linha {row_index}")
            except Exception as e_belnr:
                print(f"❌ Falha com 'BELNR': {e_belnr}")
                try: session.findById("wnd[1]").sendVKey(3)
                except: pass
                try: session.findById("wnd[0]").sendVKey(3)
                except: pass

            row_index += 1

        except Exception as e:
            print(f"\n🚨 Erro inesperado na linha {row_index}: {e}")

    close_process("saplogon.exe")   


# Criando a janela principal
root = tk.Tk()
root.title("Execução de Rotinas")
root.geometry("280x330")  # Ajuste para comportar os widgets
root.configure(bg='#f2f2f2')
root.resizable(False, False)  # Impede o redimensionamento da janela

# Fonte personalizada
root.option_add("*Font", "Amiko 10")

# Título centralizado
label_titulo = tk.Label(root, text="Download de NF-s (ZFI016)", font=("Amiko", 12, "bold"), bg='#f2f2f2', fg='#b23a48')
label_titulo.pack(pady=(0, 0))
label_titulo.pack(padx=(0, 0))

# Frame principal para centralizar os elementos
frame = tk.Frame(root, bg='#f2f2f2')
frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

# Configuração de estilo para os widgets
input_style = {'font': ('Amiko', 10), 'bg': '#ffffff', 'bd': 0, 'highlightthickness': 1,
               'highlightbackground': '#d1d1d1', 'highlightcolor': '#b23a48', 'relief': 'flat'}

label_style = {'bg': '#f2f2f2', 'font': ('Amiko', 9)}
button_style = {'font': ('Amiko', 10), 'bg': '#b23a48', 'fg': 'white', 'bd': 0,
                'activebackground': '#8c2e39', 'relief': 'flat', 'cursor': 'hand2'}

# Criando rótulos e entradas manualmente
tk.Label(frame, text="Usuário (SAP)", **label_style).grid(row=0, column=0, sticky='w', pady=(2, 2))
entry_usuario = tk.Entry(frame, **input_style, width=5)
entry_usuario.grid(row=1, column=0, sticky='ew', pady=(2, 8), padx=2)

tk.Label(frame, text="Senha (SAP)", **label_style).grid(row=0, column=1, sticky='w', pady=(2, 2))
entry_senha = tk.Entry(frame, **input_style, width=5, show="*")
entry_senha.grid(row=1, column=1, sticky='ew', pady=(2, 8), padx=2)

tk.Label(frame, text="Local de Negócio", **label_style).grid(row=2, column=0, sticky='w', pady=(2, 2))
entry_local = tk.Entry(frame, **input_style, width=8)
entry_local.grid(row=3, column=0, sticky='ew', pady=(2, 8), padx=2)

tk.Label(frame, text="Data Inicial", **label_style).grid(row=2, column=1, sticky='w', pady=(2, 2))
entry_data0 = tk.Entry(frame, **input_style, width=8)
entry_data0.grid(row=3, column=1, sticky='ew', pady=(2, 8), padx=2)

tk.Label(frame, text="Data Final", **label_style).grid(row=2, column=2, sticky='w', pady=(2, 2))
entry_data1 = tk.Entry(frame, **input_style, width=8)
entry_data1.grid(row=3, column=2, sticky='ew', pady=(2, 8), padx=2)

tk.Label(frame, text="Caminho da Pasta (Destino)", **label_style).grid(row=6, column=0, columnspan=3, sticky='w', pady=(2, 2))
entry_pastad = tk.Entry(frame, **input_style)
entry_pastad.grid(row=7, column=0, columnspan=3, sticky='ew', pady=(2, 8), padx=2)

# Centralizar o botão
button_executar = tk.Button(frame, text="Executar", command=executar_rotina, **button_style, width=20)
button_executar.grid(row=11, column=0, columnspan=3, pady=(10, 20), sticky='ew')

# Configurar o comportamento de hover para os botões
def on_enter(e):
    e.widget['bg'] = '#8c2e39'

def on_leave(e):
    e.widget['bg'] = '#b23a48'

button_executar.bind("<Enter>", on_enter)
button_executar.bind("<Leave>", on_leave)

# Inicializa a interface
root.mainloop()
