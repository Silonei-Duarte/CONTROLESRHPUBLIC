from collections import defaultdict
from textwrap import wrap
import pandas as pd
import time
from PyQt6.QtCore import Qt
import plotly.graph_objects as go
import plotly.io as pio
import tempfile
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QComboBox, QHBoxLayout, QScrollArea,
    QMessageBox, QPushButton, QFileDialog
)
from PyQt6.QtCore import QUrl  # Import necessário
from PyQt6.QtGui import QIcon
from PyQt6.QtWebEngineWidgets import QWebEngineView
import os
from PyQt6.QtWidgets import QSizePolicy

class AbaGraficoAssinuidade(QWidget):
    def __init__(self, df, periodo):
        super().__init__()
        self.df = df.copy()
        self.periodo = periodo
        self.filtered_df = df.copy()

        self.updating_filters = False
        self.resize(1900, 900)  # Largura: 1900px, Altura: 900px

        # Configurar o ícone da janela
        icon_path = os.path.join(os.path.dirname(__file__), 'icone.ico')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        # Cores fixas para cada local
        self.COLOR_MAP = {
            "Fabrica De Máquinas": "#003756",
            "Fabrica De Transportadores": "#ffc62e",
            "Adm": "#009c44",
            "Comercial": "#919191"
        }

        self.SECTOR_COLOR_MAP = {
            "Fabrica De Máquinas": "#0072cb",
            "Fabrica De Transportadores": "#ffed2d",
            "Adm": "#35b96f",
            "Comercial": "#cfcfcf"
        }

        self.NAME_COLOR = "#ff6600"  # Cor para os nomes

        # Layout principal
        self.main_layout = QVBoxLayout(self)

        # Layout fixo para filtros (linha superior)
        self.top_layout = QHBoxLayout()
        filters_container = QHBoxLayout()

        # Combobox para Local
        local_label = QLabel("Locais:")
        self.local_combobox = QComboBox()
        self.local_combobox.setFixedSize(300, 30)
        self.local_combobox.addItem("Todas")
        locais_filtrados = [local for local in self.df["Local"].dropna().unique() if local != "Total"]
        self.local_combobox.addItems(locais_filtrados)
        self.local_combobox.currentTextChanged.connect(self.update_filters)

        # Combobox para Setor
        setor_label = QLabel("Setor:")
        self.setor_combobox = QComboBox()
        self.setor_combobox.setFixedSize(300, 30)
        self.setor_combobox.addItem("Não listar")
        self.setor_combobox.addItem("Todas")
        self.setor_combobox.currentTextChanged.connect(self.update_filters)

        # Combobox para Colaborador
        nome_label = QLabel("Nomes:")
        self.nome_combobox = QComboBox()
        self.nome_combobox.setFixedSize(300, 30)
        self.nome_combobox.addItem("Não listar")
        self.nome_combobox.addItem("Todos")
        self.nome_combobox.setEnabled(False)
        self.nome_combobox.currentTextChanged.connect(self.update_filters)

        # Adicionar filtros ao container e alinhá-los à esquerda
        filters_container.addWidget(local_label)
        filters_container.addWidget(self.local_combobox)
        filters_container.addWidget(setor_label)
        filters_container.addWidget(self.setor_combobox)
        filters_container.addWidget(nome_label)
        filters_container.addWidget(self.nome_combobox)
        filters_container.addStretch()  # Adiciona espaço à direita

        # Adicionar o container de filtros ao layout superior
        self.top_layout.addLayout(filters_container)
        self.main_layout.addLayout(self.top_layout)

        # Layout inferior (Situação à esquerda, botões à direita)
        self.bottom_layout = QHBoxLayout()
        situacao_layout = QHBoxLayout()

        # Combobox para Situação
        situacao_label = QLabel("Situação:")
        self.situacao_combobox = QComboBox()
        self.situacao_combobox.setFixedSize(300, 30)
        self.situacao_combobox.currentTextChanged.connect(self.update_filters)
        situacao_layout.addWidget(situacao_label)
        situacao_layout.addWidget(self.situacao_combobox)
        situacao_layout.addStretch()
        self.bottom_layout.addLayout(situacao_layout)

        # Botões (PDF e Anexar ao Email)
        self.pdf_button = QPushButton("Gerar PDF")
        self.pdf_button.setFixedSize(120, 30)
        self.pdf_button.clicked.connect(self.generate_pdf)

        self.email_button = QPushButton("Anexar ao Email")
        self.email_button.setFixedSize(120, 30)
        self.email_button.clicked.connect(self.attach_graph_to_email)

        # Layout de botões alinhados à direita
        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()
        buttons_layout.addWidget(self.pdf_button)
        buttons_layout.addWidget(self.email_button)
        self.bottom_layout.addLayout(buttons_layout)

        # Adicionar a linha inferior ao layout principal
        self.main_layout.addLayout(self.bottom_layout)

        # Área de rolagem para gráficos
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)

        # WebEngineView para exibir o gráfico
        self.web_view = QWebEngineView()

        # Ajustar a política de expansão corretamente
        self.web_view.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        # Adicionar o web_view como widget dentro da área de rolagem
        self.scroll_area.setWidget(self.web_view)

        # Garantir que o QScrollArea controla todas as barras de rolagem
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        # Adicionar a área de rolagem ao layout principal
        self.main_layout.addWidget(self.scroll_area)

        # Atualizar filtros e gerar gráficos iniciais
        self.update_filters()

    def update_grafico(self, df_grafico):
        """Recebe o DataFrame do frequencia.py tbm e gera o gráfico com base nele"""
        self.df = df_grafico  # Atualiza o DataFrame com os dados passados
        self.filtered_df = df_grafico  # Atualiza o filtro também
        self.plot_graph()  # Chama a função para gerar o gráfico novamente

    def generate_pdf(self):
        """Gera um arquivo PDF do gráfico atual com escala reduzida."""
        try:
            file_path, _ = QFileDialog.getSaveFileName(self, "Salvar Gráfico como PDF", "", "Arquivos PDF (*.pdf)")
            if file_path:
                # Verificar se o gráfico existe (se a função plot_graph foi chamada antes)
                if not hasattr(self, 'current_figure'):
                    QMessageBox.critical(self, "Erro", "O gráfico não foi gerado ainda.")
                    return

                # Acessar o gráfico gerado
                fig = self.current_figure

                # Reduzir a escala do gráfico (se necessário)
                original_width = fig.layout.width
                original_height = fig.layout.height

                fig.update_layout(
                    width=original_width,
                    height=original_height
                )

                # Salvar o gráfico como PDF com a orientação correta
                pio.write_image(fig, file_path, format='pdf', scale=1, width=original_width, height=original_height,
                                engine="kaleido")

                # Restaurar o tamanho original após salvar
                fig.update_layout(
                    width=original_width,
                    height=original_height
                )

                QMessageBox.information(self, "Sucesso", f"Gráfico salvo como PDF em:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao salvar o PDF: {e}")

    def set_current_email(self, email):
        """Recebe o e-mail gerado e armazena na aba de gráficos."""
        self.current_mail = email

    def attach_graph_to_email(self):
        """Anexa o gráfico ao e-mail corretamente, garantindo que fique no início do corpo do e-mail, antes de tudo."""

        # Se necessário, armazene o e-mail gerado, usando set_current_email
        if hasattr(self, "current_mail") and self.current_mail:
            self.set_current_email(self.current_mail)  # Certifica-se de armazenar o e-mail antes de usar

        # Tenta pegar o e-mail armazenado dentro da aba de histórico
        if hasattr(self, "current_mail") and self.current_mail:
            email = self.current_mail
        elif hasattr(self.parentWidget(), "current_mail") and self.parentWidget().current_mail:
            email = self.parentWidget().current_mail
        elif hasattr(self.parentWidget(), "janela_principal") and self.parentWidget().janela_principal.current_mail:
            email = self.parentWidget().janela_principal.current_mail
        else:
            QMessageBox.warning(self, "Aviso",
                                "Nenhum e-mail foi gerado ainda. Por favor, gere o e-mail antes de anexar o gráfico.")
            return

        try:
            # Verificar se o gráfico foi gerado (se a função plot_graph foi chamada antes)
            if not hasattr(self, 'current_figure'):
                QMessageBox.critical(self, "Erro", "O gráfico não foi gerado ainda.")
                return

            # Acessar o gráfico gerado
            fig = self.current_figure

            # Criar um arquivo temporário para salvar a imagem do gráfico
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            image_path = temp_file.name
            temp_file.close()

            # Salvar o gráfico como imagem
            pio.write_image(fig, image_path, format="png", engine="kaleido")

            # Criar Content-ID único para a imagem (para exibição inline)
            cid = f"graph_image_{int(time.time())}"

            # Adicionar a imagem ao e-mail como anexo inline
            attachment = email.Attachments.Add(image_path)
            attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", cid)

            # Pegar o corpo atual do e-mail
            email_body = email.HTMLBody

            # Inserir a frase e o gráfico no início do corpo do e-mail
            frase = f"<p style='font-size:16px;'>Segue relatório de frequência do {self.periodo}:</p><br>"

            # Modificar o corpo do e-mail: coloca a frase e o gráfico antes do conteúdo atual
            email.HTMLBody = f"{frase}<p><img src='cid:{cid}' width='1800'></p><br><br>{email_body}"

            QMessageBox.information(self, "Sucesso", "Gráfico inserido corretamente no início do e-mail!")

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao anexar a imagem ao corpo do e-mail: {e}")

    def update_filters(self):
        """Atualiza os filtros em cascata: Local -> Setor -> Nome -> Situação"""
        if self.updating_filters:
            return

        self.updating_filters = True

        try:
            # Remover qualquer linha 'Total' antes de processar
            self.df = self.df[self.df["Local"] != "Total"]

            local = self.local_combobox.currentText()
            setor = self.setor_combobox.currentText()
            nome = self.nome_combobox.currentText()

            # Começar sempre com todo o DataFrame
            filtered_df = self.df.copy()

            # Aplicar filtro de Local primeiro
            if local != "Todas":
                filtered_df = filtered_df[filtered_df["Local"] == local]

            # Aplicar filtro de Setor (se não for "Não listar")
            if setor != "Não listar":
                if setor != "Todas":
                    filtered_df = filtered_df[filtered_df["Setor"] == setor]

            # Aplicar filtro de Nome (se não for "Não listar")
            if nome != "Não listar":
                if nome != "Todos":
                    filtered_df = filtered_df[filtered_df["Colaborador"] == nome]

            # Atualizar apenas os filtros ABAIXO do que foi selecionado
            self.filtered_df = filtered_df.copy()

            # Se um filtro acima foi alterado, redefinir os abaixo para "Não listar"
            sender = self.sender()
            if sender == self.local_combobox:
                self.setor_combobox.setCurrentIndex(0)
                self.nome_combobox.setCurrentIndex(0)

            if sender == self.setor_combobox:
                self.nome_combobox.setCurrentIndex(0)

            # Atualizar filtros disponíveis
            self.update_setor_filter()
            self.update_nome_filter()
            self.update_situacao_filter()  # Sempre chama a função corrigida!

            # Atualizar o gráfico
            self.plot_graph()

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Ocorreu um erro ao atualizar os filtros:\n{e}")

        finally:
            self.updating_filters = False

    def update_local_filter(self):
        """Atualiza os dados filtrados com base no local e ajusta os locais disponíveis."""
        locais_filtrados = [local for local in self.df["Local"].dropna().unique() if local != "Total"]

        local_selecionado = self.local_combobox.currentText()

        self.local_combobox.blockSignals(True)
        self.local_combobox.clear()
        self.local_combobox.addItem("Todas")
        self.local_combobox.addItems(locais_filtrados)
        self.local_combobox.blockSignals(False)

        if local_selecionado in locais_filtrados:
            self.local_combobox.setCurrentText(local_selecionado)
        else:
            self.local_combobox.setCurrentIndex(0)

    def update_setor_filter(self):
        """Atualiza os setores disponíveis com base no local selecionado, garantindo que 'Todas' funcione corretamente."""
        local = self.local_combobox.currentText()

        # Filtrar apenas os setores disponíveis no Local selecionado
        if local == "Todas":
            setores_disponiveis = self.df["Setor"].dropna().unique()
        else:
            setores_disponiveis = self.df[self.df["Local"] == local]["Setor"].dropna().unique()

        setor_selecionado = self.setor_combobox.currentText()

        self.setor_combobox.blockSignals(True)
        self.setor_combobox.clear()
        self.setor_combobox.addItem("Não listar")
        self.setor_combobox.addItem("Todas")
        self.setor_combobox.addItems(setores_disponiveis)
        self.setor_combobox.blockSignals(False)

        # Se "Todas" foi selecionado, manter a seleção
        if setor_selecionado == "Todas" or setor_selecionado in setores_disponiveis:
            self.setor_combobox.setCurrentText(setor_selecionado)
        else:
            self.setor_combobox.setCurrentIndex(0)

    def update_nome_filter(self):
        """Atualiza os colaboradores disponíveis com base no setor selecionado, sem alterar os filtros acima."""
        local = self.local_combobox.currentText()
        setor = self.setor_combobox.currentText()

        # Filtrar apenas os nomes disponíveis no Local e Setor selecionados
        if local == "Todas" and setor == "Todas":
            nomes_disponiveis = self.df["Colaborador"].dropna().unique()
        elif local == "Todas":
            nomes_disponiveis = self.df[self.df["Setor"] == setor]["Colaborador"].dropna().unique()
        elif setor == "Todas":
            nomes_disponiveis = self.df[self.df["Local"] == local]["Colaborador"].dropna().unique()
        else:
            nomes_disponiveis = self.df[
                (self.df["Local"] == local) & (self.df["Setor"] == setor)
                ]["Colaborador"].dropna().unique()

        nome_selecionado = self.nome_combobox.currentText()

        self.nome_combobox.blockSignals(True)
        self.nome_combobox.clear()
        self.nome_combobox.addItem("Não listar")
        self.nome_combobox.addItem("Todos")  # Mantém a opção "Todos"
        self.nome_combobox.addItems(nomes_disponiveis)
        self.nome_combobox.blockSignals(False)

        # Se "Todos" estava selecionado, mantém "Todos"
        if nome_selecionado == "Todos" or nome_selecionado in nomes_disponiveis:
            self.nome_combobox.setCurrentText(nome_selecionado)
        else:
            self.nome_combobox.setCurrentIndex(0)

        self.nome_combobox.setEnabled(True)

    def update_situacao_filter(self):
        """Atualiza o filtro para mostrar apenas colunas de cálculo (%)."""
        if self.filtered_df.empty:
            self.situacao_combobox.clear()
            return

        # Pegando apenas colunas que terminam com '(%)' e possuem pelo menos um valor numérico maior que zero
        situacoes_disponiveis = [
            col for col in self.filtered_df.columns
            if col.endswith("(%)") and self.filtered_df[col].apply(pd.to_numeric, errors='coerce').gt(0).any()
        ]

        situacao_anterior = self.situacao_combobox.currentText()

        self.situacao_combobox.blockSignals(True)
        self.situacao_combobox.clear()

        # Adiciona "Todas as Situações"
        self.situacao_combobox.addItem("Todas as Situações")

        # Adiciona apenas Assiduidade e Absenteísmo (%)
        self.situacao_combobox.addItems(situacoes_disponiveis)
        self.situacao_combobox.blockSignals(False)

        # Mantém a seleção anterior se possível
        if situacao_anterior in situacoes_disponiveis or situacao_anterior == "Todas as Situações":
            self.situacao_combobox.setCurrentText(situacao_anterior)
        elif situacoes_disponiveis:
            self.situacao_combobox.setCurrentText(situacoes_disponiveis[0])
        else:
            self.situacao_combobox.setCurrentIndex(0)

    def plot_graph_data(fig, grouped_data, offsets, bar_height, color, label_prefix=""):
        """Função para plotar dados no gráfico de ocorrências usando Plotly."""
        annotations = []

        for i, (label, valor) in enumerate(grouped_data.items()):
            # Nome formatado no lado esquerdo da barra
            wrapped_label = "\n".join(wrap(label_prefix + label, width=30))

            # Adicionar barra ao gráfico
            fig.add_trace(go.Bar(
                y=[offsets + i],
                x=[valor],
                orientation='h',
                marker=dict(color=color),
                hoverinfo="x+y",
                name=wrapped_label
            ))

            # Adicionar anotação do valor ao lado da barra
            annotations.append(dict(
                x=valor + 0.5,
                y=offsets + i,
                text=str(valor),
                showarrow=False,
                font=dict(size=10),
                xanchor='left',
                yanchor='middle'
            ))

            # Adicionar anotação do rótulo à esquerda da barra
            annotations.append(dict(
                x=0,
                y=offsets + i,
                text=wrapped_label,
                showarrow=False,
                font=dict(size=10),
                xanchor='right',
                yanchor='middle'
            ))

        # Aplicar anotações ao gráfico
        fig.update_layout(annotations=annotations)

        return offsets + len(grouped_data)

    def formatar_horas(self, total_horas):
        """Converte valores de horas para o formato HH:MM:SS, lidando corretamente com timedelta, float e strings."""
        if pd.isna(total_horas):
            return "Erro"

        try:
            if isinstance(total_horas, (int, float)):  # Se for um número decimal (exemplo: 4.5 → 4h30min)
                horas = int(total_horas)
                minutos = round((total_horas - horas) * 60)
                return f"{horas:02}:{minutos:02}:00"

            elif isinstance(total_horas, pd.Timedelta):  # Se for um timedelta (tempo acumulado)
                total_seconds = int(total_horas.total_seconds())
                horas = total_seconds // 3600
                minutos = (total_seconds % 3600) // 60
                segundos = total_seconds % 60
                return f"{horas:02}:{minutos:02}:{segundos:02}"

            elif isinstance(total_horas, str):  # Se já for string (caso de planilhas importadas)
                return total_horas.strip()

        except Exception:
            return "Erro"

        return "Erro"

    def plot_graph(self):
        """Plota os gráficos de Assiduidade ou Absenteísmo com cores agrupadas por local, setor e nome."""
        if self.filtered_df.empty:
            self.web_view.setHtml("<h3 style='text-align:center;'>Nenhum dado disponível</h3>")
            return

        situacao = self.situacao_combobox.currentText()

        if situacao == "Todas as Situações":
            self.plot_all_situations()
            return

        local = self.local_combobox.currentText()
        setor = self.setor_combobox.currentText()
        nome = self.nome_combobox.currentText()

        # Agrupamento por local
        df_local = self.df if local == "Todas" else self.df[self.df["Local"] == local]
        grouped_local = df_local.groupby("Local")[situacao].mean().round(2)
        grouped_local = grouped_local[grouped_local > 0].sort_values(ascending=False)

        # Agrupamento por setor
        df_setor = self.filtered_df if setor == "Todas" else self.filtered_df[self.filtered_df["Setor"] == setor]
        grouped_setor = df_setor.groupby(["Local", "Setor"])[situacao].mean().round(2).reset_index()
        grouped_setor = grouped_setor[grouped_setor[situacao] > 0].sort_values(by=situacao, ascending=False)

        # Agrupamento por nome
        df_nome = self.filtered_df if nome == "Todos" else self.filtered_df[self.filtered_df["Colaborador"] == nome]
        grouped_nome = df_nome.groupby("Colaborador")[situacao].mean().round(2)
        grouped_nome = grouped_nome[grouped_nome > 0].sort_values(ascending=False)

        if "Total" in grouped_local.index:
            grouped_local = grouped_local.drop("Total")

        # Preparar dados
        local_data = [(val, local, self.COLOR_MAP.get(local, "#000000")) for local, val in grouped_local.items()]
        setor_data = [(row[situacao], f"{row['Local']} - {row['Setor']}",
                       self.SECTOR_COLOR_MAP.get(f"{row['Local']} - {row['Setor']}",
                                                 self.SECTOR_COLOR_MAP.get(row['Local'], "#000000")))
                      for _, row in grouped_setor.iterrows()]

        nome_data = [(val, nome, self.NAME_COLOR) for nome, val in grouped_nome.items()]

        all_data = sorted(local_data, reverse=True) + sorted(setor_data, reverse=True) + sorted(nome_data, reverse=True)
        values, labels, colors = zip(*all_data) if all_data else ([], [], [])

        # Criar gráfico
        fig = go.Figure()

        # Criar lista de textos incluindo horas
        textos = []
        for nome, val in zip(labels, values):
            if " - " in nome:
                local, setor = nome.split(" - ", 1)
                linha = self.df[(self.df["Local"] == local) & (self.df["Setor"] == setor)]
            else:
                linha = self.df[
                    (self.df["Colaborador"] == nome) |
                    (self.df["Setor"] == nome) |
                    (self.df["Local"] == nome)
                    ]

            trabalhado = linha["Trabalhando"].sum() if "Trabalhando" in linha else pd.Timedelta(0)
            faltas = linha["Atraso/Faltas"].sum() if "Atraso/Faltas" in linha else pd.Timedelta(0)

            def format_td(td):
                if isinstance(td, pd.Timedelta):
                    total_seconds = int(td.total_seconds())
                    h = total_seconds // 3600
                    m = (total_seconds % 3600) // 60
                    s = total_seconds % 60
                    return f"{h:02}:{m:02}:{s:02}"
                return "00:00:00"

            if ("Assiduidade" in situacao and val < 98) or ("Absenteísmo" in situacao and val > 2):
                texto = f"<span style='color:red;'>{val:.2f}% | Trabalhando: {format_td(trabalhado)} | Atraso/Faltas: {format_td(faltas)}</span>"
            else:
                texto = f"{val:.2f}% | Trabalhando: {format_td(trabalhado)} | Atraso/Faltas: {format_td(faltas)}"
            textos.append(texto)

        fig.add_trace(go.Bar(
            y=labels,
            x=values,
            orientation='h',
            marker=dict(color=colors),
            text=textos,
            textposition='outside'
        ))

        max_value = max(values) if values else 0
        fig.update_layout(
            title=f"{situacao} - Período: {self.periodo}",
            xaxis_title="Percentual (%)",
            yaxis=dict(autorange="reversed"),
            height=max(500, len(labels) * 30),
            width=1800,  # aumentamos de 1000 para 1400
            margin=dict(l=100, r=300, t=50, b=50),  # aumentamos a margem direita de 150 para 300
            xaxis=dict(range=[0, max_value + 10]),
            plot_bgcolor='white',
            paper_bgcolor='white'
        )

        fig.update_traces(cliponaxis=False)

        if not values:
            self.web_view.setHtml("<h3 style='text-align:center;'>Nenhum dado disponível</h3>")
            return

        self.current_figure = fig
        temp_path = os.path.join(tempfile.gettempdir(), "graph_assiduidade_absenteismo.html")
        pio.write_html(fig, temp_path)
        self.web_view.setUrl(QUrl.fromLocalFile(temp_path))

    def plot_all_situations(self):
        """Plota o gráfico com Assiduidade (%) e Absenteísmo (%) por Local, Setor e Colaborador (médias corretas)."""
        if self.filtered_df.empty:
            self.web_view.setHtml("<h3 style='text-align:center;'>Nenhum dado disponível</h3>")
            return

        situacoes_permitidas = ["Assiduidade (%)", "Absenteísmo (%)"]

        local_filtro = self.local_combobox.currentText()
        setor_filtro = self.setor_combobox.currentText()
        nome_filtro = self.nome_combobox.currentText()

        df = self.filtered_df.copy()

        if local_filtro != "Todas":
            df = df[df["Local"] == local_filtro]
        if setor_filtro != "Todas" and setor_filtro != "Não listar":
            df = df[df["Setor"] == setor_filtro]
        if nome_filtro != "Todos" and nome_filtro != "Não listar":
            df = df[df["Colaborador"] == nome_filtro]

        # Agrupamento correto por média
        grouped_data = {}

        # Locais
        df_locais = df.groupby("Local")[situacoes_permitidas].mean().round(2)
        grouped_data["Locais"] = df_locais.to_dict(orient="index")

        # Setores
        if setor_filtro != "Não listar":
            df_setores = df.groupby(["Local", "Setor"])[situacoes_permitidas].mean().round(2).reset_index()
            df_setores["Setor_Completo"] = df_setores["Local"] + " - " + df_setores["Setor"]
            grouped_data["Setores"] = df_setores.set_index("Setor_Completo")[situacoes_permitidas].to_dict(
                orient="index")
        else:
            grouped_data["Setores"] = {}

        # Colaboradores
        if nome_filtro != "Não listar":
            df_colabs = df.groupby("Colaborador")[situacoes_permitidas].mean().round(2)
            grouped_data["Colaboradores"] = df_colabs.to_dict(orient="index")
        else:
            grouped_data["Colaboradores"] = {}

        # Preparar dados do gráfico
        y_labels, values, colors, text_labels = [], [], [], []
        index = 0

        for category in ["Locais", "Setores", "Colaboradores"]:
            sorted_items = sorted(grouped_data[category].items(), key=lambda x: sum(x[1].values()), reverse=True)
            for group_name, situacoes in sorted_items:
                primeiro = True
                for sit, val in sorted(situacoes.items(), key=lambda x: x[1], reverse=True):
                    if val <= 0:
                        continue

                    # Buscar horas reais
                    if category == "Locais":
                        linha = df[df["Local"] == group_name]
                    elif category == "Setores":
                        if " - " in group_name:
                            local, setor = group_name.split(" - ", 1)
                            linha = df[(df["Local"] == local) & (df["Setor"] == setor)]
                        else:
                            linha = pd.DataFrame()
                    else:
                        linha = df[df["Colaborador"] == group_name]

                    trabalhado = linha["Trabalhando"].sum() if "Trabalhando" in linha else pd.Timedelta(0)
                    faltas = linha["Atraso/Faltas"].sum() if "Atraso/Faltas" in linha else pd.Timedelta(0)

                    def format_td(td):
                        if isinstance(td, pd.Timedelta):
                            total_seconds = int(td.total_seconds())
                            h = total_seconds // 3600
                            m = (total_seconds % 3600) // 60
                            s = total_seconds % 60
                            return f"{h:02}:{m:02}:{s:02}"
                        return "00:00:00"

                    if ("Assiduidade" in sit and val < 98) or ("Absenteísmo" in sit and val > 2):
                        texto = f"<span style='color:red;'>{sit} {val:.2f}% | Trabalhando: {format_td(trabalhado)} | Atraso/Faltas: {format_td(faltas)}</span>"
                    else:
                        texto = f"{sit} {val:.2f}% | Trabalhando: {format_td(trabalhado)} | Atraso/Faltas: {format_td(faltas)}"
                    text_labels.append(texto)

                    y_labels.append(group_name if primeiro else "\u200B" * (index + 1))
                    values.append(val)

                    # Cor correta
                    if category == "Locais":
                        colors.append(self.COLOR_MAP.get(group_name, "#000000"))
                    elif category == "Setores":
                        local_name = group_name.split(" - ")[0]
                        colors.append(
                            self.SECTOR_COLOR_MAP.get(group_name, self.SECTOR_COLOR_MAP.get(local_name, "#000000")))
                    else:
                        colors.append(self.NAME_COLOR)

                    primeiro = False
                    index += 1

        if not values:
            self.web_view.setHtml("<h3 style='text-align:center;'>Nenhum dado disponível</h3>")
            return

        fig = go.Figure()
        fig.add_trace(go.Bar(
            y=y_labels,
            x=values,
            orientation='h',
            marker=dict(color=colors),
            text=text_labels,
            textposition='outside'
        ))

        max_text_length = max(len(txt) for txt in text_labels) if text_labels else 10
        extra_space = max_text_length * 11
        max_value = max(values) if values else 0
        dynamic_width = min(1800, max(1800, max_value + extra_space))

        fig.update_layout(
            title=f"Assiduidade / Absenteísmo (%) | Período: {self.periodo}",
            xaxis_title="Percentual (%)",
            yaxis=dict(autorange="reversed"),
            height=max(500, len(y_labels) * 30),
            width=1800,  # aumentamos de 1000 para 1400
            margin=dict(l=100, r=400, t=50, b=50),  # aumentamos a margem direita de 150 para 300
            xaxis=dict(range=[0, max_value + 10]),
            plot_bgcolor='white',
            paper_bgcolor='white'
        )
        fig.update_traces(cliponaxis=False)

        self.current_figure = fig
        temp_path = os.path.join(tempfile.gettempdir(), "graph_all_situations.html")
        pio.write_html(fig, temp_path)
        self.web_view.setUrl(QUrl.fromLocalFile(temp_path))

