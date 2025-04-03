import tkinter as tk
from tkinter import messagebox
import subprocess
import shlex
import time
import os
import win32com.client
import psutil
import shutil


def fazer_login():
    try: 
        caminho_executavel_sap = r'C:\Program Files (x86)\SAP\FrontEnd\SAPGUI\saplogon.exe'
    except:
        caminho_executavel_sap = r'C:\Program Files (x86)\SAP\FrontEnd\SAPGUI\saplogon.exe'
    
    # Verifique o caminho absoluto
    print(f"Caminho do executável: {caminho_executavel_sap}")

    if not os.path.exists(caminho_executavel_sap):
        print("O arquivo não foi encontrado. Verifique o caminho.")
        messagebox.showerror("Erro", "O executável do SAP GUI não foi encontrado. Verifique o caminho.")
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

    usuario = entry_usuario.get()
    senha = entry_senha.get()
    
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
        print(f"Conexão 'PVR - Produção (Externo)' não encontrada. Tentando 'PVR - Produção (Interno)'...")
        try:
            connection = application.OpenConnection("NOVO PVR - Produção SAP ECC", True)
        except Exception as e:
            print(f"Ocorreu um erro ao abrir a conexão: {e}")

    session = connection.Children(0)

    try:
        # Preencha as informações de cliente, usuário e senha
        session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "300"
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = usuario
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = senha
        session.findById("wnd[0]").sendVKey(0)
    except Exception as e:
        print(f"Ocorreu um erro durante a autenticação no SAP: {e}")


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


# In[4]:


def executar_rotina():
    # Armazena informações
    # codigos_material = text_codigos_material.get("1.0", tk.END).strip()  # Obtém todos os códigos do material
    close_process("saplogon.exe")
    local  = entry_local.get()
    data0 = entry_data0.get()
    data1 = entry_data1.get()
    pastat  = entry_pastat.get()
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
    session.findById("wnd[0]/usr/ctxtP_VARIA").text = "/RELFISCAL"
    session.findById("wnd[0]/usr/ctxtS_BUDAT-HIGH").setFocus
    session.findById("wnd[0]/usr/ctxtS_BUDAT-HIGH").caretPosition = 6
    session.findById("wnd[0]").sendVKey(8)


    # Inicializa o contador de linhas
    row_index = 0
    
    # Inicializa a variável de controle para continuar o loop
    has_more_rows = True
    
    while has_more_rows:
        try:
            # Atualiza a linha selecionada
            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = str(row_index)
            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellRow = row_index
            # Verifica se a linha selecionada é válida
            current_cell_column = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = "DOCNUM"
            if current_cell_column is None:
                # Se não houver mais linhas, encerra o loop
                has_more_rows = False
                continue
    
            # Executa as ações necessárias para a linha atual
            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").clickCurrentCell()
    
            try:
                session.findById("wnd[0]/titl/shellcont/shell").pressContextButton("%GOS_TOOLBOX")
                session.findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem("%GOS_VIEW_ATTA")
                session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
                session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").selectedRows = "0"
                session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
                session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").pressToolbarButton("%ATTA_EXPORT")
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = pastat
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                session.findById("wnd[1]").sendVKey(11)
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                session.findById("wnd[0]").sendVKey(3)
                row_index += 1
            except:
                session.findById("wnd[0]").sendVKey(3)
                row_index += 1
    
        except Exception as e:   
            print(f"Erro ao processar a linha {row_index}: {e}")
            has_more_rows = False
            close_process("saplogon.exe")

    # Defina os caminhos das pastas
    origem = pastat
    destino = pastad
    
    # Verifique se a pasta de destino existe, caso contrário, crie-a
    if not os.path.exists(destino):
        os.makedirs(destino)
    
    # Liste todos os arquivos na pasta de origem
    arquivos = os.listdir(origem)
    
    # Mova os arquivos para a pasta de destino
    for arquivo in arquivos:
        caminho_arquivo_origem = os.path.join(origem, arquivo)
        caminho_arquivo_destino = os.path.join(destino, arquivo)
        
        # Verifique se o arquivo já existe na pasta de destino e exclua-o se necessário
        if os.path.exists(caminho_arquivo_destino):
            os.remove(caminho_arquivo_destino)
        
        # Mova o arquivo
        shutil.move(caminho_arquivo_origem, caminho_arquivo_destino)
 
    messagebox.showinfo("Execução", "Rotina executada com sucesso!")



# In[5]:


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

tk.Label(frame, text="Caminho da Pasta (Temporária)", **label_style).grid(row=4, column=0, columnspan=3, sticky='w', pady=(2, 2))
entry_pastat = tk.Entry(frame, **input_style)
entry_pastat.grid(row=5, column=0, columnspan=3, sticky='ew', pady=(2, 8), padx=2)

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

