import tkinter as tk
from telas import tela_principal
from telas.secretaria.sge import tela_secretaria_sge
from telas.secretaria.relatorios import tela_secretaria_relatorios
from telas.secretaria.comunicacao import tela_secretaria_comunicacao

def showScreen_secretaria(root):
    for widget in root.winfo_children():
        widget.destroy()

    label_info = tk.Label(root, text="Selecione a área", wraplength=200, justify="center")
    label_info.pack(pady=10)
    label_info.config(font=("Arial Black", 12),bg='#034AA6', fg='white')

    # Botão SGE
    button1 = tk.Button(root, text="SGE", command=lambda: tela_secretaria_sge.showScreen_sge(root), height=2, width=25)
    button1.pack(pady=5)

    # Botão Relatórios
    button2 = tk.Button(root, text="Relatórios", command=lambda: tela_secretaria_relatorios.showScreen_relatorios(root), height=2, width=25)
    button2.pack(pady=5)

    # Botão Comunicação
    button3 = tk.Button(root, text="Comunicação", command=lambda: tela_secretaria_comunicacao.showScreen_comunicacao(root=root), height=2, width=25)
    button3.pack(pady=5)

    back_button = tk.Button(root, text="Voltar", command=lambda: tela_principal.showScreen_principal(root), height=1, width=10)
    back_button.pack(pady=10)
    back_button.place(relx=0.5, rely=0.9, anchor="center")