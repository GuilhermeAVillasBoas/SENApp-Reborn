import tkinter as tk
from telas.secretaria import tela_secretaria

def showScreen_relatorios(root):
    for widget in root.winfo_children():
        widget.destroy()

    label_info = tk.Label(root, text="Escolha uma opção abaixo para continuar:", wraplength=200, justify="center")
    label_info.config(font=("Arial", 12),bg='#034AA6', fg='white')
    label_info.pack(pady=10)
    
    # Botão 1
    button1 = tk.Button(root, text="Banco de horas",command=None, height=2, width=25)
    button1.pack(pady=10)

    back_button = tk.Button(root, text="Voltar", command = lambda:tela_secretaria.showScreen_secretaria(root), height=1, width=10)
    back_button.pack(pady=10)
    back_button.place(relx=0.5, rely=0.9, anchor="center")