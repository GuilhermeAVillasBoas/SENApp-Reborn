import tkinter as tk
import subprocess
from telas.secretaria.comunicacao import tela_secretaria_comunicacao


def showScreen_whatsapp(root):
    for widget in root.winfo_children():
        widget.destroy()

    turma = tk.Entry(root, width=25, font=("Arial", 10))
    turma.insert(0, "Turma (ex: 1A)")
    turma.pack(pady=20)

    mensagem = tk.Text(root, width=25, height=3, wrap="word", font=("Arial", 10))
    mensagem.insert("1.0", "Mensagem a ser enviada")
    mensagem.pack(pady=10)

    # Botão 1
    button1 = tk.Button(root, text="Enviar mensagem",command= lambda: subprocess.run(["python", "processos\enviarPywhatkit.py", turma.get(), mensagem.get("1.0", "end")]), height=2, width=25)
    button1.pack(pady=10)

    back_button = tk.Button(root, text="Voltar", command=lambda: tela_secretaria_comunicacao.showScreen_comunicacao(root=root), height=1, width=10)
    back_button.pack(pady=10)
    back_button.place(relx=0.5, rely=0.9, anchor="center")
