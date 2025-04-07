import os
import requests
from dotenv import load_dotenv


def baixar_csv(data_inicio, data_fim, output_file="relatorio.csv"):
    """Baixa o arquivo CSV da API com base nas datas fornecidas e salva localmente."""

    # Carregar variáveis do .env
    load_dotenv()
    ACCESS_TOKEN = os.getenv("ACCESS_TOKEN")

    if not ACCESS_TOKEN:
        raise ValueError("Token de autenticação não encontrado no .env!")

    # URL da API
    url = f"https://backend.metam.com.br/api/last-report/68298/export?start={data_inicio}&end={data_fim}"
    headers = {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Content-Type": "application/json"
    }

    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()  # Levanta um erro se a resposta for ruim (4xx, 5xx)
        data = response.json()

        report_url = data.get("last_report_export")
        if not report_url:
            raise ValueError("Link do relatório não encontrado no JSON.")

        csv_response = requests.get(report_url, stream=True)
        csv_response.raise_for_status()

        # Salvar o CSV
        with open(output_file, "wb") as csv_file:
            for chunk in csv_response.iter_content(chunk_size=1024):
                csv_file.write(chunk)

        return output_file  # Retorna o caminho do arquivo salvo

    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Erro ao conectar à API: {e}")
