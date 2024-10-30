import json
import hmac
import hashlib
from datetime import datetime
import tkinter as tk
from tkinter import messagebox

SECRET_KEY = b"waystermelo@"  # Deve ser a mesma chave secreta usada no sistema principal
LICENSE_FILE = "license.txt"

def create_signature(data):
    """Cria uma assinatura HMAC-SHA256 para os dados fornecidos."""
    return hmac.new(SECRET_KEY, data.encode('utf-8'), hashlib.sha256).hexdigest()

def gerar_licenca(duracao_em_minutos):
    """Gera um novo arquivo de licença com uma assinatura digital."""
    try:
        # Gerar uma nova data de ativação
        activation_time = datetime.now()
        data_str = f"{activation_time.isoformat()}|{duracao_em_minutos}"  # Inclui duração na assinatura

        # Criar assinatura digital
        signature = create_signature(data_str)

        # Salvar licença no arquivo
        with open(LICENSE_FILE, "w") as file:
            json.dump({
                "activation_time": activation_time.isoformat(),
                "duracao": duracao_em_minutos,
                "signature": signature
            }, file)

        # Exibir uma mensagem de sucesso
        root = tk.Tk()
        root.withdraw()  # Ocultar a janela principal do tkinter
        messagebox.showinfo("Licença Gerada", f"Licença gerada com sucesso!\nDuração: {duracao_em_minutos} minutos.\nArquivo de licença salvo como '{LICENSE_FILE}'.")
        root.destroy()

    except Exception as e:
        print(f"Erro ao gerar a licença: {str(e)}")

if __name__ == "__main__":
    try:
        # Solicitar ao usuário a duração da licença em minutos
        duracao_em_minutos = int(input("Informe a duração da licença em minutos: "))
        if duracao_em_minutos <= 0:
            raise ValueError("A duração deve ser um número positivo.")
        gerar_licenca(duracao_em_minutos)
    except ValueError as e:
        print(f"Erro: {str(e)}. Por favor, forneça um número inteiro positivo para a duração.")
