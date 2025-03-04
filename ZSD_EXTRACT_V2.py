import csv
import os
import tkinter as tk
from tkinter import ttk, StringVar, messagebox
from customtkinter import CTk, CTkButton, CTkEntry
import subprocess
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, StringVar
import customtkinter as ctk
import pandas as pd
import win32com.client
import openpyxl
import tempfile
from tkinter import Label
from pywinauto import Application



def fazer_login():
    # Verifique se o SAP já está aberto
    try:
        app = Application(backend="uia").connect(title="SAP Logon")
    except Exception as e:
        try:
            app = Application(backend="uia").connect(title="SAP Easy Access")
        except Exception as e:
            try:
                app = Application(backend="uia").connect(title="SAP Logon 750")
            except Exception as e:
                try:
                    app = Application(backend="uia").connect(title="SAP Easy Access 750")
                except Exception as e:
                    app = None

    # Se a aplicação do SAP não foi encontrada, inicie o SAP
    if app is None:
        caminho_executavel_sap = 'C:\\Program Files\\SAP\\FrontEnd\\SAPGUI\\saplogon.exe'
        if not os.path.exists(caminho_executavel_sap):
            messagebox.showerror("Erro", "O executável do SAP GUI não foi encontrado. Verifique o caminho.")
            return

        try:
            subprocess.Popen([caminho_executavel_sap])
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao iniciar o SAP GUI: {e}")
            return

        time.sleep(3)

    usuario = entry_usuario.get()
    senha = entry_senha.get()
    transacao_selecionada = transacao_var.get()

   # Substitua 'caminho_para_o_executavel' pelo caminho completo para o executável do SAP GUI
    caminho_executavel_sap = 'C:\\Program Files\\SAP\\FrontEnd\\SAPGUI\\saplogon.exe'
    
    # Verifique se o executável do SAP GUI existe
    if not os.path.exists(caminho_executavel_sap):
        print("O executável do SAP GUI não foi encontrado. Verifique o caminho.")
        return  
    
    # Execute o SAP GUI
    try:
        subprocess.Popen([caminho_executavel_sap])
    except Exception as e:
        print(f"Ocorreu um erro ao iniciar o SAP GUI: {e}")
    
    time.sleep(3)
    sapguiauto = win32com.client.GetObject("SAPGUI")
    application = sapguiauto.GetScriptingEngine
    try:
        connection = application.OpenConnection("PVR - Produção (Externo)", True)
    except Exception as e:
        print(f"Conexão 'PVR - Produção (Externo)' não encontrada. Tentando 'PVR - Produção (Interno)'...")
        try:
            connection = application.OpenConnection("PVR - Produção (Interno)", True)
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



    # INFORMAÇÕES SAP
    sapguiauto = win32com.client.GetObject("SAPGUI")
    application = sapguiauto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    
    # ZSD007
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = transacao_selecionada
    session.findById("wnd[0]").sendVKey(0)

    if transacao_selecionada == "ZSD007":
        layout = "/MARI APTO"
    if transacao_selecionada == "ZSD008":
        layout = "/VGV MARI"
    if transacao_selecionada == "ZSD009":
        layout = "/PARCELAS V2"


    tabela_dados_composicao = pd.read_csv("G:\\Drives compartilhados\\Controladoria_DC03\\BANCO DE DADOS CONTROLADORIA\\MATRIZ .csv")
    tabela = pd.DataFrame(tabela_dados_composicao)       
    for i, cod in enumerate(tabela["cod"]):
        nome = tabela.loc[i, "emp"]
    
        session.findById("wnd[0]/usr/ctxtS_EMPTO-LOW").text = cod                
        session.findById("wnd[0]/usr/ctxtP_VARIA").text = layout         
        session.findById("wnd[0]/usr/ctxtP_VARIA").setFocus
        session.findById("wnd[0]/usr/ctxtP_VARIA").caretPosition = 12
        session.findById("wnd[0]").sendVKey(8)
        session.findById("wnd[0]").sendVKey(45)
        session.findById("wnd[1]").sendVKey(0)
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "G:\\Drives compartilhados\\Controladoria_Contabilidade\\2 - CONTABILIDADE\\Fechamentos\\RELATÓRIOS FECHAMENTO - PASTA TRANSITÓRIA\\"+transacao_selecionada+"\\TXT"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome + "_"+transacao_selecionada+".txt"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 18
        session.findById("wnd[1]").sendVKey(11)
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"+transacao_selecionada
        session.findById("wnd[0]").sendVKey(0)                 

        input_file = 'G:\\Drives compartilhados\\Controladoria_Contabilidade\\2 - CONTABILIDADE\\Fechamentos\\RELATÓRIOS FECHAMENTO - PASTA TRANSITÓRIA\\'+transacao_selecionada+'\\TXT\\'+nome+'_'+transacao_selecionada+'.txt'
        output_file = 'G:\\Drives compartilhados\\Controladoria_Contabilidade\\2 - CONTABILIDADE\\Fechamentos\\RELATÓRIOS FECHAMENTO - PASTA TRANSITÓRIA\\'+transacao_selecionada+'\\EXCEL\\'+nome+'_'+transacao_selecionada+'.xlsx'
        
        wb = openpyxl.Workbook()
        ws = wb.worksheets[0]
        
        # Verifique se o arquivo de saída já existe e, se existir, exclua-o
        if os.path.exists(output_file):
           os.remove(output_file)
        
        with open(input_file, 'r') as data:
                    reader = csv.reader(data, delimiter='|')
                    for row in reader:
                        ws.append(row)
        
        wb.save(output_file)

    
    
    tabela_dados_composicao2 = pd.read_csv("G:\\Drives compartilhados\\Controladoria_DC03\\BANCO DE DADOS CONTROLADORIA\\MATRIZ FASES.csv")
    tabela2 = pd.DataFrame(tabela_dados_composicao2)       
    for i, cod2 in enumerate(tabela2["cod"]):
        nome2 = tabela2.loc[i, "emp"]

        session.findById("wnd[0]/usr/ctxtS_EMPTOF-LOW").text = cod2                
        session.findById("wnd[0]/usr/ctxtP_VARIA").text = layout         
        session.findById("wnd[0]/usr/ctxtP_VARIA").setFocus
        session.findById("wnd[0]/usr/ctxtP_VARIA").caretPosition = 12
        session.findById("wnd[0]").sendVKey(8)
        session.findById("wnd[0]").sendVKey(45)
        session.findById("wnd[1]").sendVKey(0)
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "G:\\Drives compartilhados\\Controladoria_Contabilidade\\2 - CONTABILIDADE\\Fechamentos\\RELATÓRIOS FECHAMENTO - PASTA TRANSITÓRIA\\"+transacao_selecionada+"\\TXT"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome2 + "_"+transacao_selecionada+".txt"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 18
        session.findById("wnd[1]").sendVKey(11)
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"+transacao_selecionada
        session.findById("wnd[0]").sendVKey(0)  
        

        input_file = 'G:\\Drives compartilhados\\Controladoria_Contabilidade\\2 - CONTABILIDADE\\Fechamentos\\RELATÓRIOS FECHAMENTO - PASTA TRANSITÓRIA\\'+transacao_selecionada+'\\TXT\\'+nome2+'_'+transacao_selecionada+'.txt'
        output_file = 'G:\\Drives compartilhados\\Controladoria_Contabilidade\\2 - CONTABILIDADE\\Fechamentos\\RELATÓRIOS FECHAMENTO - PASTA TRANSITÓRIA\\'+transacao_selecionada+'\\EXCEL\\'+nome2+'_'+transacao_selecionada+'.xlsx'
        
        wb = openpyxl.Workbook()
        ws = wb.worksheets[0]
        
        # Verifique se o arquivo de saída já existe e, se existir, exclua-o
        if os.path.exists(output_file):
           os.remove(output_file)
        
        with open(input_file, 'r') as data:
                    reader = csv.reader(data, delimiter='|')
                    for row in reader:
                        ws.append(row)
        
        wb.save(output_file)



# Crie a janela principal
janela = CTk()
janela.title("SAP Automation")

# Defina a cor de fundo escura
bg_color = "#242424"
janela.configure(bg=bg_color)  # Defina a cor de fundo da janela principal

# Defina o tamanho da fonte padrão
font_size = 14

# Função para criar um rótulo e caixa de diálogo com título centralizado
def criar_caixa_dialogo(titulo, row, show=None, pady_label=(20, 5), pady_entry=5):
    # Rótulo do título
    ttk.Label(janela, text=titulo, font=("Helvetica", font_size), foreground="white", background=bg_color).grid(row=row, column=0, columnspan=2, padx=10, pady=pady_label, sticky="n")

    # Campo de entrada
    entry = CTkEntry(janela, font=("Helvetica", font_size), show=show if show else None)  # Use a cor de fundo da janela principal
    entry.grid(row=row + 1, column=0, columnspan=2, padx=15, pady=pady_entry)

    return entry

# Crie as caixas de diálogo
entry_usuario = criar_caixa_dialogo("Usuário:", 0, pady_label=(10, 5))
entry_senha = criar_caixa_dialogo("Senha:", 2, show="*", pady_label=(10, 5))

# Rótulo de seleção de transação
ttk.Label(janela, text="Selecione a Transação:", font=("Helvetica", font_size), foreground="white", background=bg_color).grid(row=4, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="n")

# Caixa de seleção de transacao
transacao_var = StringVar()
transacao_combobox = ttk.Combobox(janela, textvariable=transacao_var, values=["ZSD007", "ZSD008", "ZSD009"], font=("Helvetica", font_size))
transacao_combobox.grid(row=5, column=0, columnspan=2, padx=10, pady=(10, 5))

# Rótulo da mensagem
mensagem = "Confira os empreendimentos no link abaixo:"
ttk.Label(janela, text=mensagem, font=("Helvetica", font_size), foreground="white", background=bg_color).grid(row=6, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="n")

# Link 1
link1 = "G:\\Drives compartilhados\\Controladoria_DC03\\BANCO DE DADOS CONTROLADORIA\\MATRIZ .csv"

# Descrição do link 1
descricao_link1 = "Lista de Empreendimentos"

def abrir_link1(event):
    os.startfile(link1)

# Rótulo para exibir a descrição clicável do link 1
lbl_descricao1 = ttk.Label(janela, text=descricao_link1, font=("Helvetica", font_size), foreground="blue", background=bg_color, cursor="hand2")
lbl_descricao1.grid(row=7, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="n")

# Configurar a ação de clique na descrição do link 1
lbl_descricao1.bind("<Button-1>", abrir_link1)

# Link 2
link2 = "G:\\Drives compartilhados\\Controladoria_DC03\\BANCO DE DADOS CONTROLADORIA\\MATRIZ FASES.csv"

# Descrição do link 2
descricao_link2 = "Lista de Empreendimentos Fase"

def abrir_link2(event):
    os.startfile(link2)

# Rótulo para exibir a descrição clicável do link 2
lbl_descricao2 = ttk.Label(janela, text=descricao_link2, font=("Helvetica", font_size), foreground="blue", background=bg_color, cursor="hand2")
lbl_descricao2.grid(row=9, column=0, columnspan=2, padx=10, pady=(15, 10), sticky="n")

# Configurar a ação de clique na descrição do link 2
lbl_descricao2.bind("<Button-1>", abrir_link2)


# Botão de login personalizado
btn_login = CTkButton(janela, text="Executar", command=fazer_login, corner_radius=5, font=("Helvetica", font_size), bg_color="#a0c0ff")
btn_login.grid(row=10, column=0, columnspan=2, padx=10, pady=10)

janela.mainloop()



