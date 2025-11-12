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

class GraficoAfastamento(QWidget):
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
        self.local_combobox.addItem("Todos")
        locais_filtrados = [local for local in self.df["Local"].dropna().unique() if local != "Total"]
        self.local_combobox.addItems(locais_filtrados)
        self.local_combobox.currentTextChanged.connect(self.update_filters)

        # Combobox para Setor
        setor_label = QLabel("Setor:")
        self.setor_combobox = QComboBox()
        self.setor_combobox.setFixedSize(300, 30)
        self.setor_combobox.addItem("Não listar")
        self.setor_combobox.addItem("Todos")
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
        situacao_label = QLabel("Filtrar Situações por:")
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

        # Garantir que o QScrollArea controla Todos as barras de rolagem
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
            if local != "Todos":
                filtered_df = filtered_df[filtered_df["Local"] == local]

            # Aplicar filtro de Setor (se não for "Não listar")
            if setor != "Não listar":
                if setor != "Todos":
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
        self.local_combobox.addItem("Todos")
        self.local_combobox.addItems(locais_filtrados)
        self.local_combobox.blockSignals(False)

        if local_selecionado in locais_filtrados:
            self.local_combobox.setCurrentText(local_selecionado)
        else:
            self.local_combobox.setCurrentIndex(0)

    def update_setor_filter(self):
        """Atualiza os setores disponíveis com base no local selecionado, garantindo que 'Todos' funcione corretamente."""
        local = self.local_combobox.currentText()

        # Filtrar apenas os setores disponíveis no Local selecionado
        if local == "Todos":
            setores_disponiveis = self.df["Setor"].dropna().unique()
        else:
            setores_disponiveis = self.df[self.df["Local"] == local]["Setor"].dropna().unique()

        setor_selecionado = self.setor_combobox.currentText()

        self.setor_combobox.blockSignals(True)
        self.setor_combobox.clear()
        self.setor_combobox.addItem("Não listar")
        self.setor_combobox.addItem("Todos")
        self.setor_combobox.addItems(setores_disponiveis)
        self.setor_combobox.blockSignals(False)

        # Se "Todos" foi selecionado, manter a seleção
        if setor_selecionado == "Todos" or setor_selecionado in setores_disponiveis:
            self.setor_combobox.setCurrentText(setor_selecionado)
        else:
            self.setor_combobox.setCurrentIndex(0)

    def update_nome_filter(self):
        """Atualiza os colaboradores disponíveis com base no setor selecionado, sem alterar os filtros acima."""
        local = self.local_combobox.currentText()
        setor = self.setor_combobox.currentText()

        # Filtrar apenas os nomes disponíveis no Local e Setor selecionados
        if local == "Todos" and setor == "Todos":
            nomes_disponiveis = self.df["Colaborador"].dropna().unique()
        elif local == "Todos":
            nomes_disponiveis = self.df[self.df["Setor"] == setor]["Colaborador"].dropna().unique()
        elif setor == "Todos":
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
        """Atualiza as situações disponíveis, garantindo que só apareçam as disponíveis no filtro atual."""
        if self.filtered_df.empty:
            self.situacao_combobox.clear()
            return

        # Pegando apenas colunas que começam com 'Qtd' e que possuem pelo menos um valor maior que zero
        situacoes_disponiveis = [
            col for col in self.filtered_df.columns
            if col.startswith("Qtd") and self.filtered_df[col].sum() > 0
        ]

        situacao_anterior = self.situacao_combobox.currentText()

        self.situacao_combobox.blockSignals(True)
        self.situacao_combobox.clear()

        # Adicionar opção "Todos os Dados"
        self.situacao_combobox.addItem("Todos os Dados")

        # Adicionar situações individuais
        self.situacao_combobox.addItems(situacoes_disponiveis)
        self.situacao_combobox.blockSignals(False)

        # Se a situação anterior ainda estiver disponível, manter a seleção
        if situacao_anterior in situacoes_disponiveis or situacao_anterior == "Todos os Dados":
            self.situacao_combobox.setCurrentText(situacao_anterior)
        elif situacoes_disponiveis:
            # Caso contrário, definir automaticamente a primeira situação disponível
            self.situacao_combobox.setCurrentText(situacoes_disponiveis[0])
        else:
            self.situacao_combobox.setCurrentIndex(0)  # Se não houver nenhuma situação, resetar

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
        """Plota os gráficos ajustando cores e agrupando por quantidade da situação filtrada com Plotly."""
        if self.filtered_df.empty:
            html_content = "<h3 style='text-align:center;'>Nenhum dado disponível</h3>"
            self.web_view.setHtml(html_content)
            return

        situacao = self.situacao_combobox.currentText()

        # Criar um dicionário mapeando cada "Qtd. [Situação]" para sua respectiva coluna de horas
        colunas_horas = {
            col: col.replace("Qtd. ", "") for col in self.filtered_df.columns if col.startswith("Qtd. ")
        }
        colunas_horas = {qtd: horas for qtd, horas in colunas_horas.items() if horas in self.filtered_df.columns}

        # Selecione a coluna de horas correspondente à situação atual
        coluna_horas = colunas_horas.get(situacao, None)

        # Chama a função separada se for "Todos os Dados"
        if situacao == "Todos os Dados":
            self.plot_all_situations()
            return

        local = self.local_combobox.currentText()
        setor = self.setor_combobox.currentText()
        nome = self.nome_combobox.currentText()

        # Agrupar os dados
        if local == "Todos":
            grouped_local = self.df.groupby("Local")[situacao].sum().sort_values(ascending=False)
            grouped_local_horas = self.df.groupby("Local")[
                coluna_horas].sum() if coluna_horas and coluna_horas in self.df.columns else pd.Series(
                dtype="timedelta64[ns]")
        else:
            grouped_local = self.df[self.df["Local"] == local].groupby("Local")[situacao].sum().sort_values(
                ascending=False)
            grouped_local_horas = self.df[self.df["Local"] == local].groupby("Local")[
                coluna_horas].sum() if coluna_horas and coluna_horas in self.df.columns else pd.Series(
                dtype="timedelta64[ns]")

        grouped_local = grouped_local[grouped_local > 0]  # Remove locais sem valor

        if setor == "Todos":
            grouped_setor = self.filtered_df.groupby(["Local", "Setor"])[situacao].sum().reset_index()
            grouped_setor_horas = self.filtered_df.groupby(["Local", "Setor"])[
                coluna_horas].sum().reset_index() if coluna_horas and coluna_horas in self.filtered_df.columns else pd.DataFrame()
        else:
            grouped_setor = self.filtered_df[self.filtered_df["Setor"] == setor].groupby(["Local", "Setor"])[
                situacao].sum().reset_index()
            grouped_setor_horas = self.filtered_df[self.filtered_df["Setor"] == setor].groupby(["Local", "Setor"])[
                coluna_horas].sum().reset_index() if coluna_horas and coluna_horas in self.filtered_df.columns else pd.DataFrame()

        grouped_setor[situacao] = pd.to_numeric(grouped_setor[situacao], errors="coerce").fillna(0).astype(int)
        grouped_setor = grouped_setor[grouped_setor[situacao] > 0]
        grouped_setor = grouped_setor.sort_values(by=situacao, ascending=False)

        if nome == "Todos":
            grouped_nome = self.filtered_df.groupby("Colaborador")[situacao].sum().sort_values(ascending=False)
            grouped_nome_horas = self.filtered_df.groupby("Colaborador")[
                coluna_horas].sum() if coluna_horas and coluna_horas in self.filtered_df.columns else pd.Series(
                dtype="timedelta64[ns]")
        else:
            grouped_nome = self.filtered_df[self.filtered_df["Colaborador"] == nome].groupby("Colaborador")[
                situacao].sum().sort_values(ascending=False)
            grouped_nome_horas = self.filtered_df[self.filtered_df["Colaborador"] == nome].groupby("Colaborador")[
                coluna_horas].sum() if coluna_horas and coluna_horas in self.filtered_df.columns else pd.Series(
                dtype="timedelta64[ns]")

        grouped_nome = grouped_nome[grouped_nome > 0]

        if "Total" in grouped_local.index:
            grouped_local = grouped_local.drop("Total")

        # Criar listas para gráficos
        local_data = [
            (val, local, self.COLOR_MAP[local], grouped_local_horas.get(local, pd.Timedelta(0)))
            for local, val in grouped_local.items()]

        setor_data = [
            (
                row[situacao],
                f"{row['Local']} - {row['Setor']}",
                self.SECTOR_COLOR_MAP.get(f"{row['Local']} - {row['Setor']}",
                                          self.SECTOR_COLOR_MAP.get(row['Local'], "#000000")),
                grouped_setor_horas[grouped_setor_horas["Local"] == row["Local"]][
                    grouped_setor_horas["Setor"] == row["Setor"]][
                    coluna_horas].sum() if not grouped_setor_horas.empty else pd.Timedelta(0)
            )
            for _, row in grouped_setor.iterrows()
        ]

        nome_data = [
            (val, nome, self.NAME_COLOR, grouped_nome_horas.get(nome, pd.Timedelta(0)))
            for nome, val in grouped_nome.items()]

        # Ordenar os dados
        local_data.sort(reverse=True, key=lambda x: x[0])
        setor_data.sort(reverse=True, key=lambda x: x[0])
        nome_data.sort(reverse=True, key=lambda x: x[0])

        # Juntar os dados na ordem correta (Locais -> Setores -> Nomes)
        sorted_data = local_data + setor_data + nome_data

        # Desempacotar os dados em y_labels e text_labels
        values, labels, colors, horas = zip(*sorted_data) if sorted_data else ([], [], [], [])

        # Formatar as horas corretamente
        horas_formatadas = [self.formatar_horas(h) if isinstance(h, (pd.Timedelta, int, float)) else "00:00:00" for h in
                            horas]

        # Agora, definindo y_labels e text_labels com as horas formatadas
        y_labels = labels
        text_labels = [f"{valor} | {horas}" if horas != "00:00:00" else f"{valor}" for valor, horas in
                       zip(values, horas_formatadas)]

        # Criar gráfico com Plotly
        fig = go.Figure()
        fig.add_trace(go.Bar(
            y=y_labels,
            x=values,
            orientation='h',
            marker=dict(color=colors),
            text=text_labels,
            textposition='outside'  # Garante que o texto será mostrado ao lado da barra
        ))

        # Calcular o comprimento máximo do texto para definir a largura da área de plotagem
        max_text_length = max(len(txt) for txt in text_labels) if text_labels else 10
        extra_space = max_text_length * 10  # Ajuste o multiplicador para garantir o espaço necessário para o texto

        # Ajustar largura do gráfico considerando o maior valor e o texto
        max_value = max(values) if values else 0  # Valor máximo das barras
        dynamic_width = min(1800, max(1800, max_value + extra_space))


        titulo_bruto = f"{situacao} - {self.periodo}"
        titulo_formatado = self.quebrar_titulo_auto(titulo_bruto)

        fig.update_layout(
            title=dict(
                text=titulo_formatado,
                font=dict(size=16),
            ),
            xaxis_title="Quantidade",
            yaxis=dict(autorange="reversed"),  # Inverte a ordem do eixo Y (barras de cima para baixo)
            height=max(500, len(y_labels) * 30),  # Ajusta a altura dinamicamente
            width=dynamic_width,  # Ajusta a largura do gráfico
            margin=dict(l=100, r=10, t=50, b=50),
            xaxis=dict(range=[0, max(10, max_value + 10)])  # Define um mínimo de 10, mas ajusta conforme o necessário
        )

        # Se values estiver vazio, exibir uma mensagem de "Nenhum dado disponível"
        if not values:
            html_content = "<h3 style='text-align:center;'>Nenhum dado disponível</h3>"
            self.web_view.setHtml(html_content)
            return

        self.current_figure = fig

        # Salvar HTML temporário e carregar no WebView
        temp_path = os.path.join(tempfile.gettempdir(), "graph_all_situations.html")
        pio.write_html(fig, temp_path)

        self.web_view.setUrl(QUrl.fromLocalFile(temp_path))

    def plot_all_situations(self):
        """Plota o gráfico exibindo Todos os Dados separadas por local, setor e colaborador, usando Plotly."""
        if self.filtered_df.empty:
            html_content = "<h3 style='text-align:center;'>Nenhum dado disponível</h3>"
            self.web_view.setHtml(html_content)
            return

        situacoes_disponiveis = [
            col for col in self.filtered_df.columns
            if col.startswith("Qtd") and self.filtered_df[col].sum() > 0
        ]

        # Criar um dicionário mapeando cada "Qtd. [Situação]" para sua respectiva coluna de horas
        colunas_horas = {
            col: col.replace("Qtd. ", "") for col in self.filtered_df.columns if col.startswith("Qtd. ")
        }
        colunas_horas = {qtd: horas for qtd, horas in colunas_horas.items() if horas in self.filtered_df.columns}


        if not situacoes_disponiveis:
            html_content = "<h3 style='text-align:center;'>Nenhum dado disponível</h3>"
            self.web_view.setHtml(html_content)
            return

        local_filtro = self.local_combobox.currentText()
        setor_filtro = self.setor_combobox.currentText()
        nome_filtro = self.nome_combobox.currentText()

        grouped_data = {
            "Locais": defaultdict(lambda: defaultdict(int)),
            "Setores": defaultdict(lambda: defaultdict(int)),
            "Colaboradores": defaultdict(lambda: defaultdict(int)),
            "Horas_Locais": defaultdict(lambda: defaultdict(lambda: pd.Timedelta(0))),
            "Horas_Setores": defaultdict(lambda: defaultdict(lambda: pd.Timedelta(0))),
            "Horas_Colaboradores": defaultdict(lambda: defaultdict(lambda: pd.Timedelta(0)))

        }

        for _, row in self.filtered_df.iterrows():
            local_name = row["Local"]
            setor_name = row["Setor"]
            colaborador_name = row["Colaborador"]

            # Aplicar os filtros corretamente
            if local_filtro != "Todos" and local_name != local_filtro:
                continue
            if setor_filtro != "Todos" and setor_filtro != "Não listar" and setor_name != setor_filtro:
                continue
            if nome_filtro != "Todos" and nome_filtro != "Não listar" and colaborador_name != nome_filtro:
                continue

            for sit in situacoes_disponiveis:
                valor = row[sit]
                coluna_horas = colunas_horas.get(sit, None)  # Obter a coluna correspondente de horas
                horas = pd.to_timedelta(row[coluna_horas], errors="coerce") if (
                            coluna_horas and pd.notna(row[coluna_horas])) else pd.Timedelta(0)
                if pd.isna(horas) or not isinstance(horas, pd.Timedelta):
                    horas = pd.Timedelta(0)

                if valor > 0:
                    grouped_data["Locais"][local_name][sit] += valor
                    grouped_data["Horas_Locais"][local_name][sit] += horas if pd.notna(horas) else pd.Timedelta(0)

                    if setor_filtro != "Não listar":
                        setor_unico = f"{local_name} - {setor_name}"
                        grouped_data["Setores"][setor_unico][sit] += valor
                        grouped_data["Horas_Setores"][setor_unico][sit] += horas if pd.notna(horas) else pd.Timedelta(0)

                    if nome_filtro != "Não listar":
                        grouped_data["Colaboradores"][colaborador_name][sit] += valor
                        grouped_data["Horas_Colaboradores"][colaborador_name][sit] += horas if pd.notna(
                            horas) else pd.Timedelta(0)

        y_labels = []  # Lista de rótulos únicos no eixo Y
        values = []  # Lista de valores das barras
        colors = []  # Lista de cores das barras
        text_labels = []
        index = 0  # Índice global para forçar separação das barras

        for category in ["Locais", "Setores", "Colaboradores"]:  # Agora inclui "Colaboradores"
            sorted_items = sorted(grouped_data[category].items(), key=lambda x: sum(x[1].values()), reverse=True)
            for group_name, situacoes in sorted_items:
                primeiro = True  # Flag para exibir o nome apenas uma vez dentro do grupo
                for sit, val in sorted(situacoes.items(), key=lambda x: x[1], reverse=True):
                    # Pega as horas corretamente dentro do loop
                    horas_dict_map = {
                        "Locais": "Horas_Locais",
                        "Setores": "Horas_Setores",
                        "Colaboradores": "Horas_Colaboradores"
                    }

                    horas_categoria = horas_dict_map.get(category, None)

                    if horas_categoria and group_name in grouped_data[horas_categoria]:
                        horas = grouped_data[horas_categoria][group_name].get(sit, pd.Timedelta(0))
                    else:
                        horas = pd.Timedelta(0)

                    if pd.isna(horas) or not isinstance(horas, pd.Timedelta):
                        horas = pd.Timedelta(0)

                    # Formata corretamente o rótulo incluindo as horas somadas
                    text_labels.append(f"{val} | {self.formatar_horas(horas)} | {sit}" if self.formatar_horas(
                        horas) != "00:00:00" else f"{val} | {sit}")

                    # Adiciona os rótulos do eixo Y
                    y_labels.append(group_name if primeiro else "\u200B" * (index + 1))
                    values.append(val)

                    # Aplica as cores corretamente
                    if category == "Locais":
                        colors.append(self.COLOR_MAP[group_name])
                    elif category == "Setores":
                        local_name = group_name.split(" - ")[0]
                        colors.append(self.SECTOR_COLOR_MAP.get(group_name, self.SECTOR_COLOR_MAP.get(local_name)))
                    elif category == "Colaboradores":
                        colors.append(self.NAME_COLOR)  # Define a cor do colaborador

                    primeiro = False
                    index += 1

        # Criar gráfico com Plotly
        fig = go.Figure()
        fig.add_trace(go.Bar(
            y=y_labels,
            x=values,
            orientation='h',
            marker=dict(color=colors),
            text=text_labels,
            textposition='outside'  # Garante que o texto será mostrado ao lado da barra
        ))

        # Calcular o comprimento máximo do texto para definir a largura da área de plotagem
        max_text_length = max(len(txt) for txt in text_labels) if text_labels else 10
        extra_space = max_text_length * 10  # Ajuste o multiplicador para garantir o espaço necessário para o texto

        # Ajustar largura do gráfico considerando o maior valor e o texto
        max_value = max(values) if values else 0  # Valor máximo das barras
        dynamic_width = min(1800, max(1800, max_value + extra_space))

        titulo_bruto = f"{self.periodo}"
        titulo_formatado = self.quebrar_titulo_auto(titulo_bruto)

        fig.update_layout(
            title=dict(
                text=titulo_formatado,
                font=dict(size=16),
            ),
            xaxis_title="Quantidade",
            yaxis=dict(autorange="reversed"),
            height=max(500, len(y_labels) * 30),
            width=dynamic_width,
            margin=dict(l=100, r=10, t=80, b=50),
            xaxis=dict(range=[0, max(10, max_value * 1.15)])
        )

        # Se `values` estiver vazio, exibir uma mensagem de "Nenhum dado disponível"
        if not values:
            html_content = "<h3 style='text-align:center;'>Nenhum dado disponível</h3>"
            self.web_view.setHtml(html_content)
            return

        self.current_figure = fig

        # Salvar HTML temporário e carregar no WebView
        temp_path = os.path.join(tempfile.gettempdir(), "graph_all_situations.html")
        pio.write_html(fig, temp_path)

        self.web_view.setUrl(QUrl.fromLocalFile(temp_path))

    def quebrar_titulo_auto(self, texto):
        max_linha = 220
        if len(texto) <= max_linha:
            return texto
        # Tenta quebrar no espaço mais próximo da posição max_linha
        pos_quebra = texto.rfind(' ', 0, max_linha)
        if pos_quebra == -1:
            # Se não houver espaço antes, tenta depois
            pos_quebra = texto.find(' ', max_linha)
            if pos_quebra == -1:
                # Se não achar espaço nenhum, quebra bruto
                pos_quebra = max_linha
        return texto[:pos_quebra] + '<br>' + texto[pos_quebra + 1:]

