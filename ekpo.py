import tkinter as tk
from tkinter import messagebox
import win32com.client
import time
import datetime
import os
import pandas as pd
import traceback
import logging

# --- Configura o logger para depuração ---
logging.basicConfig(
    filename="sap_automation.log",
    level=logging.DEBUG,
    format="%(asctime)s %(levelname)s %(message)s",
)

# --- Função auxiliar para pegar usuário do Windows ---
def safe_get_username():
    try:
        return os.getlogin()
    except Exception:
        return os.environ.get("USERNAME", "unknown")

# --- Função principal de automação SAP ---
def executar_automacao_sap():
    try:
        usuario = safe_get_username()
        pasta = fr"C:\Users\{usuario}\OneDrive - Accenture\Desktop\junior"
        os.makedirs(pasta, exist_ok=True)

        data_atual = datetime.datetime.now().strftime("%d_%m_%y")
        nome_xls = f"EKPO_{data_atual}.XLS"
        caminho_xls = os.path.join(pasta, nome_xls)
        caminho_txt = os.path.splitext(caminho_xls)[0] + ".txt"

        # --- Conecta ao SAP ---
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)

        # --- Código do VBScript convertido (EKPO) ---
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "se16n"
        session.findById("wnd[0]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/usr/ctxtGD-TAB").text = "ekpo"
        session.findById("wnd[0]/usr/ctxtGD-TAB").setFocus()
        session.findById("wnd[0]/usr/ctxtGD-TAB").caretPosition = 4
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # Exportação
        time.sleep(2)
        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")

        # Seleciona formato Excel
        session.findById(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/"
            "sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[2,0]"
        ).select()
        session.findById(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/"
            "sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[2,0]"
        ).setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # Define caminho e nome do arquivo
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = pasta
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_xls
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # Confirma substituição se aparecer
        try:
            session.findById("wnd[1]/tbar[0]/btn[20]").press()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except:
            pass

        time.sleep(2)

        # --- Converte XLS → TXT ---
        df = pd.read_excel(caminho_xls, header=None)
        df.to_csv(caminho_txt, sep="\t", index=False, header=False)
        os.remove(caminho_xls)

        messagebox.showinfo("Sucesso", f"Arquivo EKPO salvo e convertido:\n{caminho_txt}")

    except Exception as e:
        logging.error("Erro na automação: " + str(e) + "\n" + traceback.format_exc())
        messagebox.showerror("Erro", f"Ocorreu um erro na automação:\n\n{e}")

# --- Interface gráfica (Tkinter) ---
def criar_interface():
    janela = tk.Tk()
    janela.title("Junior Dev - Automação SAP")
    janela.geometry("400x300")
    janela.configure(bg="#0d0d0d")  

    # --- Título ---
    label_titulo = tk.Label(
        janela,
        text="Junior Dev",
        font=("Helvetica", 24, "bold"),
        fg="#3d9df2",  # azul neon
        bg="#0d0d0d"
    )
    label_titulo.pack(pady=50)

    # --- Botões ---
    estilo_botao = {
        "font": ("Helvetica", 12, "bold"),
        "bg": "#b3b3b3",
        "fg": "#0d0d0d",
        "activebackground": "#1a1a1a",
        "activeforeground": "#00aaff",
        "width": 15,
        "height": 1,
        "relief": "ridge",
        "bd": 3,
    }

    btn_iniciar = tk.Button(
        janela,
        text="Iniciar",
        command=executar_automacao_sap,
        **estilo_botao
    )
    btn_iniciar.pack(pady=10)

    btn_voltar = tk.Button(
        janela,
        text="Voltar",
        command=janela.destroy,
        **estilo_botao
    )
    btn_voltar.pack(pady=5)

    janela.mainloop()

# --- Executa ---
if __name__ == "__main__":
    criar_interface()
