import tkinter as tk
import subprocess
from telas.secretaria.sge import tela_secretaria_sge

def showScreen_mec(root):
 #   webbrowser.open(url_mec_login)

    for widget in root.winfo_children():
        widget.destroy()

    label = tk.Label(root, text="MEC - SISTEC", wraplength=200, justify="center")
    label.config(font=("Arial", 12),bg='#034AA6', fg='white')
    label.pack(pady=10)

    script_button1 = tk.Button(root, text="Código de Autenticação", command=lambda: subprocess.run("processos\codigo-de-autenticacao-pyautogui.py"), height=2, width=20)
    script_button1.pack(pady=10)

    back_button = tk.Button(root, text="Voltar", command=tela_secretaria_sge.showScreen_sge(root), height=1, width=10)
    back_button.pack(pady=10)
    back_button.place(relx=0.5, rely=0.9, anchor="center")

#    mostrar_mensagem("Atenção", "Faça login.", erro=False)