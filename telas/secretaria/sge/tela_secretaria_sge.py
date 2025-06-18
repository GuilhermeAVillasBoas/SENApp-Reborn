import tkinter as tk
from telas.secretaria import tela_secretaria
from telas.secretaria.sge.mec import tela_secretaria_sge_mec

def showScreen_sge(root):
#    close_specific_tab(janela_mec)
 #   close_specific_tab(janela_mecLogin)
    """Exibe a segunda tela."""
    for widget in root.winfo_children():
        widget.destroy()

    label_info = tk.Label(root, text="Escolha uma opção abaixo para continuar:", wraplength=200, justify="center")
    label_info.config(font=("Arial", 12),bg='#034AA6', fg='white')
    label_info.pack(pady=10)

    # Botão Código de Autenticação
    button1 = tk.Button(root, text="MEC - SISTEC", command=lambda: tela_secretaria_sge_mec.showScreen_mec(root), height=2, width=25)
    button1.pack(pady=5)

    # Botão Tela 2
    button2 = tk.Button(root, text="Em breve...", height=2, width=25)
    button2.pack(pady=5)

    back_button = tk.Button(root, text="Voltar", command=lambda: tela_secretaria.showScreen_secretaria(root), height=1, width=10)
    back_button.pack(pady=10)
    back_button.place(relx=0.5, rely=0.9, anchor="center")

#    bring_or_open_window_fullscreen(janela_sge, "C:\Totvs\RM.NET\RM.exe")