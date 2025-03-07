import time
import subprocess
import os
import shutil
import psutil
import shlex
from datetime import datetime
from tkinter import messagebox
import tkinter as tk
import win32com.client



# In[2]:


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

    usuario_sap = entry_usuario_sap.get()
    senha_sap = entry_senha_sap.get()

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
    for proc in psutil.process_iter(["pid", "name"]):
        if nome_processo.lower() in proc.info["name"].lower():
            try:
                processo = psutil.Process(proc.info["pid"])
                processo.terminate()  # ou processo.kill() para forçar o fechamento
                print(f'{proc.info["name"]} (PID: {proc.info["pid"]}) foi fechado.')
            except psutil.NoSuchProcess:
                print(f'Erro: Processo {proc.info["name"]} não encontrado.')
            except psutil.AccessDenied:
                print(f'Erro: Acesso negado para fechar {proc.info["name"]}.')


# In[4]:


def copiar_arquivos(origem, destino):
    try:
        if not os.path.exists(destino):
            os.makedirs(destino)  # Cria a pasta de destino se não existir

        for item in os.listdir(origem):
            origem_item = os.path.join(origem, item)
            destino_item = os.path.join(destino, item)
            if os.path.isfile(origem_item):
                shutil.copy2(origem_item, destino_item)  # Copia o arquivo
                print(f"Copiado: {origem_item} para {destino_item}")
            elif os.path.isdir(origem_item):
                shutil.copytree(origem_item, destino_item)  # Copia a pasta
                print(f"Copiado: {origem_item} para {destino_item}")
    except Exception as e:
        print(f"Ocorreu um erro ao copiar arquivos: {e}")
        messagebox.showerror("Erro", f"Ocorreu um erro ao copiar arquivos: {e}")


# In[5]:


# Obter o diretório do usuário atual
user_dir = os.path.expanduser("~")

# Criar o caminho dinâmico para a pasta SAP\SAP GUI
sap_gui_path = os.path.join(user_dir, "Documents", "SAP", "SAP GUI")


# In[6]:


def pasta_origem():
    pasta = entry_pasta.get()
    return pasta


# In[7]:


def excluir_arquivos(pasta):
    try:
        # Verifica se a pasta existe
        if os.path.exists(pasta):
            for item in os.listdir(pasta):
                caminho_item = os.path.join(pasta, item)
                if os.path.isfile(caminho_item):
                    os.remove(caminho_item)  # Exclui o arquivo
                    print(f"Arquivo excluído: {caminho_item}")
                elif os.path.isdir(caminho_item):
                    shutil.rmtree(caminho_item)  # Exclui a pasta
                    print(f"Pasta excluída: {caminho_item}")
        else:
            print("A pasta especificada não existe.")
            messagebox.showerror("Erro", "A pasta especificada não existe.")
    except Exception as e:
        print(f"Ocorreu um erro ao excluir arquivos: {e}")
        messagebox.showerror("Erro", f"Ocorreu um erro ao excluir arquivos: {e}")


# In[8]:


def SAP_NF():
    # Fazer login no SAP
    fazer_login()

    # INFORMAÇÕES SAP
    sapguiauto = win32com.client.GetObject("SAPGUI")
    application = sapguiauto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "mir5"
    session.findById("wnd[0]").sendVKey(0)

    try:
        # Verifica se a pasta existe
        if os.path.exists(sap_gui_path):
            for arquivo in os.listdir(sap_gui_path):
                caminho_arquivo = os.path.join(sap_gui_path, arquivo)
                if os.path.isfile(caminho_arquivo):
                    # Armazenando o nome do arquivo
                    nome_arquivo = arquivo

                    # Divide o nome do arquivo em partes com base no traço "-"
                    partes = arquivo.split("-")

                    # Verifica se o nome do arquivo segue o padrão esperado
                    if len(partes) >= 3:
                        nf = partes[
                            0
                        ]  # Primeiro valor, antes do primeiro traço, é a NF
                        migo = partes[
                            1
                        ]  # Segundo valor, entre o primeiro e o segundo traço, é a MIGO

                        print(f"Processando arquivo: {arquivo}")
                        print(f"NF: {nf}, MIGO: {migo}")
                        print(f"Nome do arquivo: {nome_arquivo}")

                        # Executar o código SAP GUI usando as variáveis NF e MIGO
                        session.findById("wnd[0]/usr/txtSO_BELNR-LOW").text = migo
                        session.findById("wnd[0]/usr/txtSO_XBLNR-LOW").text = nf
                        session.findById("wnd[0]").sendVKey(8)
                        session.findById(
                            "wnd[0]/usr/cntlGRID1/shellcont/shell"
                        ).clickCurrentCell()
                        session.findById(
                            "wnd[0]/titl/shellcont/shell"
                        ).pressContextButton("%GOS_TOOLBOX")
                        session.findById(
                            "wnd[0]/titl/shellcont/shell"
                        ).selectContextMenuItem("%GOS_PCATTA_CREA")
                        session.findById("wnd[1]/usr/ctxtDY_PATH").text = sap_gui_path
                        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = (
                            nome_arquivo
                        )
                        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = (
                            25
                        )
                        session.findById("wnd[1]").sendVKey(0)
                        session.findById("wnd[0]").sendVKey(12)
                        session.findById("wnd[0]").sendVKey(12)

                        print(f"Arquivo {arquivo} processado com sucesso.")
                    else:
                        print(f"O arquivo '{arquivo}' não segue o padrão esperado.")
        else:
            print("A pasta especificada não existe.")
    except Exception as e:
        print(f"Ocorreu um erro ao processar os arquivos e executar no SAP GUI: {e}")


# In[10]:


def obter_data_criacao(arquivo):
    # Obtém a data de criação do arquivo (timestamp)
    timestamp_criacao = os.path.getctime(arquivo)

    # Converte o timestamp para uma data legível
    data_criacao = datetime.fromtimestamp(timestamp_criacao)

    return data_criacao


def mover_arquivos_por_empreendimento(caminho_base, pasta_origem):
    # Lista todos os arquivos na pasta de origem
    for arquivo in os.listdir(pasta_origem):
        if arquivo.endswith(".pdf"):
            # Extrai os 4 últimos dígitos do nome do arquivo (sem a extensão .pdf)
            codigo_arquivo = arquivo[
                -8:-4
            ]  # Pega os últimos 4 dígitos antes da extensão

            # Obtém a data de criação do arquivo
            caminho_origem = os.path.join(pasta_origem, arquivo)
            data_criacao = obter_data_criacao(caminho_origem)
            ano = data_criacao.strftime("%Y")
            mes = data_criacao.strftime("%m.%Y")
            dia = data_criacao.strftime("%d.%m.%Y")

            # Percorre todas as pastas dentro do caminho base
            for pasta in os.listdir(caminho_base):
                caminho_empreendimento = os.path.join(caminho_base, pasta)

                # Verifica se a pasta começa com "00" e se os 4 primeiros dígitos da pasta correspondem ao código do arquivo
                if (
                    os.path.isdir(caminho_empreendimento)
                    and pasta.startswith("00")
                    and pasta[:4] == codigo_arquivo
                ):
                    print(f"Movendo arquivo {arquivo} para o empreendimento: {pasta}")

                    # Cria o caminho completo para a pasta com base na data de criação do arquivo
                    caminho_ano = os.path.join(caminho_empreendimento, ano)
                    caminho_mes = os.path.join(caminho_ano, mes)
                    caminho_dia = os.path.join(caminho_mes, dia)

                    # Cria a pasta do ano se não existir
                    if not os.path.exists(caminho_ano):
                        os.makedirs(caminho_ano)
                        print(f"Pasta {caminho_ano} criada.")

                    # Cria a pasta do mês se não existir
                    if not os.path.exists(caminho_mes):
                        os.makedirs(caminho_mes)
                        print(f"Pasta {caminho_mes} criada.")

                    # Cria a pasta do dia se não existir
                    if not os.path.exists(caminho_dia):
                        os.makedirs(caminho_dia)
                        print(f"Pasta {caminho_dia} criada.")

                    # Move o arquivo para a pasta do dia correspondente
                    caminho_destino = os.path.join(caminho_dia, arquivo)
                    shutil.move(caminho_origem, caminho_destino)

                    print(f"Arquivo {arquivo} movido para {caminho_destino}.")
                    break  # Sai do loop para não verificar outras pastas

    print("Processo de movimentação concluído.")


# In[11]:


def executar_rotina():
    pasta = pasta_origem()
    excluir_arquivos(sap_gui_path)
    copiar_arquivos(pasta, sap_gui_path)
    SAP_NF()
    close_process("saplogon.exe")
    excluir_arquivos(sap_gui_path)
    caminho_base = r"G:\Drives compartilhados\Fiscal_Arquivo_de_Notas\FISCAL_SERVIÇOS"
    mover_arquivos_por_empreendimento(caminho_base, pasta)
    messagebox.showinfo("Sucesso", "NF's anexadas com sucesso a MIGO no SAP!")


# In[12]:


# Criando a janela principal
root = tk.Tk()
root.title("FISCAL")
root.geometry("230x250")  # Ajuste para comportar melhor os widgets
root.configure(bg="#f2f2f2")
root.resizable(False, False)  # Impede o redimensionamento da janela

# Fonte personalizada
root.option_add("*Font", "Amiko 10")

# Título centralizado com quebra de linha
label_titulo = tk.Label(
    root,
    text="Upload NF SAP",
    font=("Amiko", 10, "bold"),
    bg="#f2f2f2",
    fg="#b23a48",
    wraplength=240,  # Ajuste a largura conforme necessário
)
label_titulo.pack(pady=(10, 10))

# Frame principal para centralizar os elementos
frame = tk.Frame(root, bg="#f2f2f2")
frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=5)

# Configuração de estilo para os widgets
input_style = {
    "font": ("Amiko", 10),
    "bg": "#ffffff",
    "bd": 0,
    "highlightthickness": 1,
    "highlightbackground": "#d1d1d1",
    "highlightcolor": "#b23a48",
    "relief": "flat",
}

label_style = {"bg": "#f2f2f2", "font": ("Amiko", 10)}
button_style = {
    "font": ("Amiko", 10),
    "bg": "#b23a48",
    "fg": "white",
    "bd": 0,
    "activebackground": "#8c2e39",
    "relief": "flat",
    "cursor": "hand2",
}

# Criando rótulos e entradas
tk.Label(frame, text="Usuário (SAP)", **label_style).grid(
    row=0, column=0, sticky="w", pady=(2, 1)
)
entry_usuario_sap = tk.Entry(frame, **input_style, width=25)
entry_usuario_sap.grid(row=1, column=0, sticky="ew", pady=(0, 5))

tk.Label(frame, text="Senha (SAP)", **label_style).grid(
    row=2, column=0, sticky="w", pady=(2, 1)
)
entry_senha_sap = tk.Entry(frame, **input_style, width=25, show="*")
entry_senha_sap.grid(row=3, column=0, sticky="ew", pady=(0, 5))

tk.Label(frame, text="Pasta de Origem", **label_style).grid(
    row=4, column=0, sticky="w", pady=(2, 1)
)
entry_pasta = tk.Entry(frame, **input_style, width=25)
entry_pasta.grid(row=5, column=0, sticky="ew", pady=(0, 5))

# Centralizar o botão
button_executar = tk.Button(
    frame, text="Realizar Upload", command=executar_rotina, **button_style, width=25
)
button_executar.grid(row=8, column=0, pady=(10, 10), sticky="ew")


# Configurar o comportamento de hover para os botões
def on_enter(e):
    e.widget["bg"] = "#8c2e39"


def on_leave(e):
    e.widget["bg"] = "#b23a48"


button_executar.bind("<Enter>", on_enter)
button_executar.bind("<Leave>", on_leave)


# Adicionando a funcionalidade de tecla Enter
def on_enter_pressed(event):
    executar_rotina()


root.bind("<Return>", on_enter_pressed)

# Inicializa a interface
root.mainloop()
