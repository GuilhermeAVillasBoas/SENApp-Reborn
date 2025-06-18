import tkinter as tk
import subprocess
from telas.secretaria.comunicacao import tela_secretaria_comunicacao


def showScreen_whatsapp(root):
    for widget in root.winfo_children():
        widget.destroy()

    label_info = tk.Label(root
                          , text="WhatsApp"
                          , wraplength=200
                          , justify="center")
    label_info.config(font=("Arial Black", 12)
                      , bg='#034AA6'
                      , fg='white')
    label_info.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

    label_turma = tk.Label(root
                           , text="Código da turma"
                           , wraplength=200
                           , justify="center"
                           , font=("Arial Black", 9)
                           , bg='#034AA6'
                           , fg='white')
    label_turma.grid(row=1, column=0, padx=5, pady=10)

    turma = tk.Entry(root
                     , width=20
                     , font=("Arial", 10))
    turma.grid(row=1, column=1, padx=10, pady=10)
    turma.insert(0, "ex: QUA01232025U123")

    label_mensagem = tk.Label(root
                              , text="Mensagem"
                              , wraplength=200
                              , justify="center"
                              , font=("Arial Black", 9)
                              , bg='#034AA6'
                              , fg='white')
    label_mensagem.grid(row=2, column=0, padx=5, pady=10)

    mensagem = tk.Text(root
                       , width=20
                       , height=3
                       , wrap="word"
                       , font=("Arial", 10))
    mensagem.grid(row=2, column=1, padx=10, pady=10)
    mensagem.insert("1.0", "Mensagem a ser enviada")

    # Botão 1
    button1 = tk.Button(root
                        , text="Enviar mensagem"
                        ,command= lambda: subprocess.run(["python"
                                                          , "processos\enviarPywhatkit.py"
                                                          , turma.get()
                                                          , mensagem.get("1.0", "end")])
                        , height=2
                        , width=25)
    button1.grid(row=3, 
                 column=0, 
                 columnspan=2, 
                 padx=10, 
                 pady=5)

    back_button = tk.Button(root
                            , text="Voltar"
                            , command=lambda: tela_secretaria_comunicacao.showScreen_comunicacao(root=root)
                            , height=1
                            , width=5)
    back_button.place(row=4
                      , column=0
                      , columnspan=2
                      , padx=10
                      , pady=5)
