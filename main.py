import tkinter as tk
from tkinter import ttk
from pgr import main
from PIL import Image, ImageTk
import os
import threading
import subprocess
import time

USERNAME = os.getenv("USERNAME")
stop_threads = False  # Global control variable to stop threads

####### FILE PATHS #######
file_base_rtf = f"C:\\Users\\{USERNAME}\\Downloads\\teste-pgr\\files\\teste-novo\\NOVEMBRO 2024 - PGR - JUMA ALIMENTOS LTDA.rtf"
pgr_modelo = f"C:\\Users\\{USERNAME}\\Downloads\\teste-pgr\\files\\teste-novo\\NOVO MODELO PGR.docx"
pgr_destino = f"C:\\Users\\{USERNAME}\\Downloads\\teste-pgr\\files\\teste-novo\\nome_arquivo_novo.docx"
pdf_path = f"C:\\Users\\{USERNAME}\\Downloads\\teste-pgr\\files\\27.11.2024 - AEP - T.J. ALIMENTOS LTDA - COZINHA GERAL.pdf"

# Function to kill Word processes
def matar_word():
    try:
        subprocess.call(["taskkill", "/F", "/IM", "WINWORD.EXE"])
        print("Processo do Word encerrado com sucesso.")
    except Exception as e:
        print(f"Erro ao tentar encerrar o Word: {str(e)}")

# Function to start the process in a thread
def thread_iniciar_processo():
    global stop_threads
    stop_threads = False  # Reset flag when starting the process
    threading.Thread(target=iniciar_processo).start()
    mover_barra_progresso()  # Start the progress bar

# Function to stop the process
def parar_robo():
    global stop_threads
    stop_threads = True  # Signal to stop the threads
    progress_label.config(text="Processo interrompido!")
    progress_bar.stop()

# Function to start the document processing
def iniciar_processo():
    global stop_threads
    if stop_threads:
        return  # If stop signal is triggerezd, exit early

    # Update the progress status
    progress_label.config(text="Iniciando o processamento...")
    root.update()  # Update the Tkinter interface

    try:
        # Step 1: Process the documents
        progress_label.config(text="Processando documentos PGR...")
        root.update()  # Update the Tkinter interface
        main(file_base_rtf, pgr_modelo, pgr_destino, pdf_path)  # Call the main function
        
        if stop_threads:
            return  # Stop the process if needed

        progress_label.config(text="Processamento concluído!")
    except Exception as e:
        progress_label.config(text=f"Erro: {str(e)}")
    finally:
        progress_bar.stop()
        root.update()  # Ensure the Tkinter interface is updated at the end

# Function to move the progress bar while the process is active
def mover_barra_progresso():
    if progress_label.cget("text") != "Processamento concluído!" and not stop_threads:
        progress_bar.step(1)  # Move the progress bar one step
        root.after(100, mover_barra_progresso)  # Call again after 100ms

# Tkinter GUI setup
root = tk.Tk()
root.title("Processar Arquivos PGR")
root.geometry("400x300")

# Add the company logo
logo_image_path = f"C:\\Users\\{USERNAME}\\Desktop\\pcmso\\logo_empresa.jpg"
if os.path.exists(logo_image_path):  # Check if the file exists
    logo_image = Image.open(logo_image_path)
    logo_image = logo_image.resize((200, 100), Image.LANCZOS)
    logo_photo = ImageTk.PhotoImage(logo_image)
    logo_label = tk.Label(root, image=logo_photo)
    logo_label.pack(pady=10)
else:
    print(f"Arquivo de logo não encontrado em: {logo_image_path}")

# Button to start the process
botao_processar = tk.Button(
    root, text="Processar arquivos PGR", command=thread_iniciar_processo)
botao_processar.pack(pady=10)

# Button to stop the process
botao_parar = tk.Button(root, text="Parar o Robô", command=parar_robo)
botao_parar.pack(pady=10)

# Progress bar
progress_bar = ttk.Progressbar(root, orient="horizontal", mode="determinate", length=280)
progress_bar.pack(pady=10)

# Status label for the process
progress_label = tk.Label(root, text="Aguardando...")
progress_label.pack()

# Start the Tkinter interface
root.mainloop()
