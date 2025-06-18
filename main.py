import tkinter as tk
from telas import tela_principal

root = tk.Tk()
root.title("SENApp")  # Título da janela
root.iconbitmap("img/icone.ico")  # Ícone da janela
root.geometry("300x250+1600+600")  # Posição inicial da janela
root.resizable(False, False)  # Janela não redimensionável
root.attributes("-topmost", True)  # Janela sempre no topo
root.configure(bg='#034AA6')  # Cor de fundo da janela
tela_principal.showScreen_principal(root)
root.mainloop()
