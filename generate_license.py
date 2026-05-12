"""
Gerador de Licença - PDF Analyzer
Gera o arquivo license.txt com assinatura HMAC-SHA256 válida.
"""

import hashlib
import hmac
import json
from datetime import datetime, timedelta

SECRET_KEY = b"waystermelo@"
LICENSE_FILE = "license.txt"


def create_signature(data: str) -> str:
    """Cria uma assinatura HMAC-SHA256 para os dados fornecidos."""
    return hmac.new(SECRET_KEY, data.encode('utf-8'), hashlib.sha256).hexdigest()


def generate_license(duracao_minutos: int) -> dict:
    """Gera os dados da licença com a duração especificada em minutos."""
    activation_time = datetime.now().isoformat()
    data_str = f"{activation_time}|{duracao_minutos}"
    signature = create_signature(data_str)

    license_data = {
        "activation_time": activation_time,
        "duracao": duracao_minutos,
        "signature": signature
    }
    return license_data


def main():
    print("=" * 50)
    print("  GERADOR DE LICENÇA - PDF Analyzer")
    print("=" * 50)
    print()
    print("Atalhos de duração:")
    print("  1) 60 min      (1 hora)")
    print("  2) 1440 min    (1 dia)")
    print("  3) 10080 min   (7 dias)")
    print("  4) 43200 min   (30 dias)")
    print("  5) 525600 min  (365 dias)")
    print("  6) 5256000 min (10 anos)")
    print("  0) Digitar valor personalizado")
    print()

    opcao = input("Escolha uma opção (0-6): ").strip()

    atalhos = {
        "1": 60,
        "2": 1440,
        "3": 10080,
        "4": 43200,
        "5": 525600,
        "6": 5256000,
    }

    if opcao in atalhos:
        duracao = atalhos[opcao]
    elif opcao == "0":
        try:
            duracao = int(input("Digite a duração em minutos: ").strip())
            if duracao <= 0:
                print("ERRO: A duração deve ser maior que zero.")
                return
        except ValueError:
            print("ERRO: Valor inválido. Digite um número inteiro.")
            return
    else:
        print("ERRO: Opção inválida.")
        return

    license_data = generate_license(duracao)

    with open(LICENSE_FILE, "w") as f:
        json.dump(license_data, f)

    # Cálculo para exibição
    activation = datetime.fromisoformat(license_data["activation_time"])
    expiration = activation + timedelta(minutes=duracao)
    dias = duracao / 1440

    print()
    print("=" * 50)
    print("  LICENÇA GERADA COM SUCESSO!")
    print("=" * 50)
    print(f"  Arquivo:     {LICENSE_FILE}")
    print(f"  Ativação:    {activation.strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"  Expiração:   {expiration.strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"  Duração:     {duracao} minutos ({dias:.1f} dias)")
    print(f"  Assinatura:  {license_data['signature'][:32]}...")
    print("=" * 50)


if __name__ == "__main__":
    main()
