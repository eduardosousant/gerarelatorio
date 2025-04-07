import os
import sys
import pandas as pd
from PySide6.QtGui import QIcon, QFont, QPixmap
from PySide6.QtCore import QDate
from PySide6.QtWidgets import QApplication, QWidget, QFileDialog, QProgressBar, QToolButton, QLabel, QDateEdit, QFrame, \
    QErrorMessage
from api import baixar_csv  # Importa a função de download
from relatorio import gerar_grafico, gerar_relatorio  # Importar as funções
from datetime import datetime

# Função para obter o caminho correto para arquivos, considerando a execução como executável ou ambiente de desenvolvimento
def obter_caminho_arquivo(arquivo):
    if getattr(sys, 'frozen', False):  # Se o código está rodando como executável
        return os.path.join(sys._MEIPASS, arquivo)
    else:  # Se está rodando no ambiente de desenvolvimento
        return os.path.join(os.getcwd(), arquivo)
class ReportGenerator(QWidget):
    def __init__(self):
        super().__init__()
        self.setup_ui()
        self.setWindowTitle("Gerador de Relatórios")
        self.setWindowIcon(QIcon("ico.png"))
        self.setStyleSheet("background-color: lightblue;")  # Cor de fundo azul clara

    def setup_ui(self):
        self.resize(400, 360)

        # Frame
        self.frame = QFrame(self)
        self.frame.setGeometry(70, 180, 281, 160)

        # Data Inicial
        self.label_data_inicial = QLabel("Data Inicial:", self.frame)
        self.label_data_inicial.setGeometry(20, 20, 100, 25)

        self.data_inicial = QDateEdit(self.frame)
        self.data_inicial.setGeometry(120, 20, 120, 25)
        self.data_inicial.setCalendarPopup(True)
        self.data_inicial.setDate(QDate.currentDate().addMonths(-1)) # Primeiro dia do mês anterior

        # Data Final
        self.label_data_final = QLabel("Data Final:", self.frame)
        self.label_data_final.setGeometry(20, 50, 100, 25)

        self.data_final = QDateEdit(self.frame)
        self.data_final.setGeometry(120, 50, 120, 25)
        self.data_final.setCalendarPopup(True)
        self.data_final.setDate(QDate.currentDate().addDays(-1))  # Define padrão como hoje

        # Botão Importar
        self.Importar = QToolButton(self.frame, text="Gerar Relatório...")
        self.Importar.setGeometry(40, 90, 191, 21)
        self.Importar.clicked.connect(self.importar_arquivo)

        # Barra de Progresso
        self.progresso = QProgressBar(self.frame)
        self.progresso.setGeometry(20, 120, 251, 23)
        self.progresso.setValue(0)

        # Logo
        self.label_logo = QLabel(self, pixmap=QPixmap("ico.png"))
        self.label_logo.setGeometry(155, 10, 80, 80)
        self.label_logo.setScaledContents(True)

        # Título
        self.label_title = QLabel("Gerador de Relatórios", self)
        self.label_title.setGeometry(70, 90, 271, 50)
        self.label_title.setFont(QFont("Arial", 20))

        # Mensagem informativa
        self.label_info = QLabel("Para garantir a precisão do relatório, selecione apenas datas anteriores ao dia atual.", self)
        self.label_info.setGeometry(20, 140, 400, 50)
        self.label_info.setFont(QFont("Arial", 10))
        self.label_info.setWordWrap(True)  # Permite quebra de linha caso o texto seja muito grande


        # Rodapé
        self.label_footer = QLabel("assistentecomercial01@geoeste.com.br", self)
        self.label_footer.setGeometry(176, 330, 191, 20)
        self.label_footer.setFont(QFont("Arial", 7))

    def importar_arquivo(self):
        """Baixa automaticamente o CSV e processa o arquivo."""
        data_inicio = self.data_inicial.date().toString("yyyy-MM-dd")
        data_fim = self.data_final.date().toString("yyyy-MM-dd")

        try:
            csv_path = baixar_csv(data_inicio, data_fim)
            self.processar_arquivo(csv_path)  # Processa o arquivo baixado
        except Exception as e:
            self.show_error_message("Erro", str(e))

    def show_error_message(self, title, message):
        error_dialog = QErrorMessage(self)
        error_dialog.setWindowTitle(title)
        error_dialog.showMessage(message)


    def processar_arquivo(self, file):
        try:
            df = pd.read_csv(file, header=0, sep=";", encoding="latin1")
        except Exception as e:
            self.show_error_message("Erro ao Ler Arquivo", f"Não foi possível ler o arquivo CSV: {e}")
            return

        # 🔹 Normaliza os nomes das colunas para evitar erros de compatibilidade
        df.columns = (
            df.columns.str.normalize('NFKD')
            .str.encode('ascii', errors='ignore')
            .str.decode('utf-8')
            .str.strip().str.lower().str.replace(r"[ _-]", "", regex=True)
        )

        # 🔹 Tenta acessar a coluna "Hidrômetro" pelo índice 18 ou pelo nome
        if len(df.columns) > 18:
            df["Hidrômetro"] = df.iloc[:, 18]
        else:
            col_hidrometro = next(
                (col for col in df.columns if any(p in col for p in ["hidrometro", "hidro", "hidr"])), None
            )
            if col_hidrometro:
                df.rename(columns={col_hidrometro: "Hidrômetro"}, inplace=True)
            else:
                self.show_error_message("Erro no Arquivo", "Não foi possível identificar a coluna do Hidrômetro.")
                return

        # 🔹 Converte a coluna de data e filtra os dados relevantes
        df["data_hora_dispositivo"] = pd.to_datetime(df.iloc[:, 0], errors='coerce')
        df = df[["data_hora_dispositivo", "Hidrômetro"]].dropna()

        # 🔹 Verifica se as colunas existem
        if "data_hora_dispositivo" not in df.columns or "Hidrômetro" not in df.columns:
            self.show_error_message("Arquivo CSV inválido", "Colunas obrigatórias ausentes.")
            return

        self.progresso.setValue(50)

        # Obtém as datas escolhidas pelo usuário
        data_inicio = self.data_inicial.date().toPython()
        data_fim = self.data_final.date().toPython()

        # Filtra os dados pelo intervalo de tempo selecionado
        df = df[(df["data_hora_dispositivo"].dt.date >= data_inicio) &
                (df["data_hora_dispositivo"].dt.date <= data_fim)]


        # 🔹 Cria coluna "data" sem a informação de hora
        df["data"] = df["data_hora_dispositivo"].dt.date
        df["data"] = pd.to_datetime(df["data"])
        df['dia_da_semana'] = df['data_hora_dispositivo'].dt.dayofweek
        df['final_de_semana'] = df['dia_da_semana'].apply(lambda x: 'Final de semana' if x >= 5 else 'Dia de semana')

        # 🔹 Ordena os dados para garantir cálculo correto
        df.sort_values("data_hora_dispositivo", inplace=True)

        # 🔹 Calcula o consumo diário (última leitura - primeira leitura do dia)
        consumo_diario = df.groupby("data")["Hidrômetro"].apply(lambda x: x.iloc[-1] - x.iloc[0]).reset_index()

        # 🔹 Garante que todas as datas estejam no DataFrame (preenchendo com "Sem Dados")
        datas_completas = pd.DataFrame(pd.date_range(start=df["data"].min(), end=df["data"].max()), columns=["data"])
        consumo_diario = datas_completas.merge(consumo_diario, on="data", how="left")
        consumo_diario["Hidrômetro"] = pd.to_numeric(consumo_diario["Hidrômetro"], errors="coerce")
        consumo_diario["Hidrômetro"] = consumo_diario["Hidrômetro"].fillna(0)  # Opcional: Substituir NaN por 0

        # 🔹 Calcula estatísticas (evita erro caso `consumo_diario` esteja vazio)
        if not consumo_diario.empty and consumo_diario["Hidrômetro"].dtype != object:

            maior_consumo = consumo_diario.loc[consumo_diario["Hidrômetro"].idxmax()]
            consumo_diario['mes_ano'] = pd.to_datetime(consumo_diario['data']).dt.to_period('M')
            media_mensal = consumo_diario.groupby('mes_ano')['Hidrômetro'].mean().reset_index()
            consumo_semanal = df.groupby(['final_de_semana', 'data'])['Hidrômetro'].apply(lambda x: x.iloc[-1] - x.iloc[0]).reset_index()
            consumo_semanal_summary = consumo_semanal.groupby('final_de_semana')['Hidrômetro'].sum().reset_index()
            mes_ano = consumo_diario['mes_ano'].iloc[0].strftime('%m/%Y')
            total_consumo = consumo_diario["Hidrômetro"].sum()
        else:
            maior_consumo = "Sem Dados"
            media_mensal = pd.DataFrame()
            consumo_semanal = "Sem Dados"
            consumo_semanal_summary = "Sem Dados"
            mes_ano = "Sem Dados"
            total_consumo = "Sem Dados"

        self.progresso.setValue(70)

        # Obter a data atual
        data_atual = datetime.now()

       # Chama a função do relatorio.py para gerar o gráfico
        gerar_grafico(consumo_diario, mes_ano, data_inicio, data_fim)
        nome_arquivo = f"relatorio_{data_atual.strftime('%d%m%y')}.pdf"

         # Chama a função do relatorio.py para gerar o relatório
        gerar_relatorio(nome_arquivo, mes_ano, consumo_diario, total_consumo, maior_consumo, media_mensal, consumo_semanal_summary, data_inicio, data_fim)
        self.progresso.setValue(100)
        os.startfile(nome_arquivo)


if __name__ == "__main__":
    app = QApplication([])
    window = ReportGenerator()
    window.show()
    app.exec()
