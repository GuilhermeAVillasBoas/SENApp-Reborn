import tkinter as tk
from telas.secretaria import tela_secretaria as tela

def showScreen_principal(root):
    # Limpa a tela atual
    for widget in root.winfo_children():
        widget.destroy()

    label_status = tk.Label(root, text="", bg=root.cget("bg"))
    label_status.pack(side="bottom", fill="x")


        # FRAME do título
    frame_titulo = tk.Frame(root, bg='#034AA6')
    frame_titulo.pack(pady=(20, 10))

    label_logo = tk.Label(frame_titulo, text="≡ SENAI ≡", bg='#034AA6', fg='white',
                        font=("Arial Black", 24, 'bold italic'))
    label_logo.pack()

        # FRAME do subtítulo
    frame_info = tk.Frame(root, bg='#034AA6')
    frame_info.pack(pady=5)

    label_info = tk.Label(frame_info, text="Escolha um setor", wraplength=200,
                        justify="center", font=("Arial", 12), bg='#034AA6', fg='white')
    label_info.pack()

        # FRAME 1
    frame1 = tk.Frame(root, bg='#034AA6')
    frame1.pack(pady=(10, 5))

    button1 = tk.Button(frame1, text="Secretaria", command=lambda: tela.showScreen_secretaria(root),
                        height=1, width=15)
    button1.pack(side="left", padx=5)

    button2 = tk.Button(frame1, text="Financeiro", command=None,
                        height=1, width=15)
    button2.pack(side="left", padx=5)

    # FRAME 2
    frame2 = tk.Frame(root, bg='#034AA6')
    frame2.pack(pady=5)

    button3 = tk.Button(frame2, text="Atendimento", command=None,
                        height=1, width=15)
    button3.pack(side="left", padx=5)

    button4 = tk.Button(frame2, text="Em Breve", command=None,
                        height=1, width=15)
    button4.pack(side="left", padx=5)