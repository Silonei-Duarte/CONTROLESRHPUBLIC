from textwrap import wrap
import pandas as pd
from PyQt6.QtCore import Qt, QUrl
import plotly.graph_objects as go
import plotly.io as pio
import tempfile
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QComboBox, QHBoxLayout, QScrollArea,
    QMessageBox, QPushButton, QFileDialog
)
from PyQt6.QtGui import QIcon
from PyQt6.QtWebEngineWidgets import QWebEngineView
import os
from PyQt6.QtWidgets import QSizePolicy

class GraficoSetoresApp(QWidget):
    def __init__(self, df, periodo):
        super().__init__()
        self.df = df.copy()
        self.periodo = periodo
        self.filtered_df = df.copy()

        self.updating_filters = False
        self.resize(1900, 900)

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

        self.NAME_COLOR = "#ff6600"

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


        filters_container.addWidget(local_label)
        filters_container.addWidget(self.local_combobox)
        filters_container.addWidget(setor_label)
        filters_container.addWidget(self.setor_combobox)
        filters_container.addStretch()

        self.top_layout.addLayout(filters_container)
        self.main_layout.addLayout(self.top_layout)

        self.bottom_layout = QHBoxLayout()
        situacao_layout = QHBoxLayout()

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

        # Layout de botões alinhados à direita
        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()
        buttons_layout.addWidget(self.pdf_button)
        self.bottom_layout.addLayout(buttons_layout)

        # Adicionar a linha inferior ao layout principal
        self.main_layout.addLayout(self.bottom_layout)

        # 🔹 Área de rolagem para gráficos
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)

        # 🔹 WebEngineView para exibir o gráfico
        self.web_view = QWebEngineView()

        # 🔹 Ajustar a política de expansão corretamente
        self.web_view.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        # 🔹 Adicionar o web_view como widget dentro da área de rolagem
        self.scroll_area.setWidget(self.web_view)

        # 🔹 Garantir que o QScrollArea controla Todos as barras de rolagem
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        # 🔹 Adicionar a área de rolagem ao layout principal
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

    def update_filters(self):
        """Atualiza os filtros em cascata: Local -> Setor -> Nome -> Situação"""
        if self.updating_filters:
            return

        self.updating_filters = True

        try:
            # 🔹 Remover qualquer linha 'Total' antes de processar
            self.df = self.df[self.df["Local"] != "Total"]

            local = self.local_combobox.currentText()
            setor = self.setor_combobox.currentText()
            nome = "Não listar"  # Nome fixado como não listar já que o filtro foi removido

            # 🔹 Começar sempre com todo o DataFrame
            filtered_df = self.df.copy()

            # 🔹 Aplicar filtro de Local primeiro
            if local != "Todos":
                filtered_df = filtered_df[filtered_df["Local"] == local]

            # 🔹 Aplicar filtro de Setor (se não for "Não listar")
            if setor != "Não listar":
                if setor != "Todos":
                    filtered_df = filtered_df[filtered_df["Setor"] == setor]

            # 🔹 Atualizar apenas os filtros ABAIXO do que foi selecionado
            self.filtered_df = filtered_df.copy()

            # 🔹 Se um filtro acima foi alterado, redefinir os abaixo para "Não listar"
            sender = self.sender()
            if sender == self.local_combobox:
                self.setor_combobox.setCurrentIndex(0)

            # 🔹 Atualizar filtros disponíveis
            self.update_setor_filter()
            self.update_situacao_filter()  # 🔹 Sempre chama a função corrigida!

            # 🔹 Atualizar o gráfico
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

        # 🔹 Filtrar apenas os setores disponíveis no Local selecionado
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

        # 🔹 Se "Todos" foi selecionado, manter a seleção
        if setor_selecionado == "Todos" or setor_selecionado in setores_disponiveis:
            self.setor_combobox.setCurrentText(setor_selecionado)
        else:
            self.setor_combobox.setCurrentIndex(0)

    def update_situacao_filter(self):
        """Atualiza as situações disponíveis, incluindo colunas numéricas e percentual formatado como string."""
        if self.filtered_df.empty:
            self.situacao_combobox.clear()
            return

        situacoes_disponiveis = []

        for col in self.filtered_df.columns:
            if col in ["Local", "Setor"]:
                continue

            serie = self.filtered_df[col]

            # Coluna numérica normal
            if pd.api.types.is_numeric_dtype(serie):
                if serie.sum() > 0:
                    situacoes_disponiveis.append(col)

            # Coluna percentual em string, ex: "2.83%"
            elif serie.dtype == object and serie.str.contains('%').any():
                try:
                    valores = serie.str.replace('%', '', regex=False).str.replace(',', '.', regex=False).astype(float)
                    if valores.sum() > 0:
                        situacoes_disponiveis.append(col)
                except Exception:
                    pass

        situacao_anterior = self.situacao_combobox.currentText()

        self.situacao_combobox.blockSignals(True)
        self.situacao_combobox.clear()
        self.situacao_combobox.addItems(situacoes_disponiveis)
        self.situacao_combobox.blockSignals(False)

        if situacao_anterior in situacoes_disponiveis:
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

    def plot_graph(self):
        """Plota os gráficos ajustando cores e agrupando por quantidade da situação filtrada com Plotly."""
        if self.filtered_df.empty:
            html_content = "<h3 style='text-align:center;'>Nenhum dado disponível</h3>"
            self.web_view.setHtml(html_content)
            return

        situacao = self.situacao_combobox.currentText()

        if situacao not in self.filtered_df.columns:
            QMessageBox.critical(self, "Erro", f"A coluna '{situacao}' não foi encontrada nos dados.")
            return

        local = self.local_combobox.currentText()
        setor = self.setor_combobox.currentText()

        # Agrupar os dados
        if local == "Todos":
            # Sempre pega todos os dados para a parte de Local
            df_local = self.df.copy()
            if df_local[situacao].dtype == object and df_local[situacao].str.contains('%').any():
                df_local[situacao] = pd.to_numeric(
                    df_local[situacao].str.replace('%', '', regex=False).str.replace(',', '.', regex=False),
                    errors='coerce')

            grouped_local = df_local.groupby("Local")[situacao].sum().sort_values(ascending=False)
        else:
            df_local = self.df[self.df["Local"] == local].copy()
            if df_local[situacao].dtype == object and df_local[situacao].str.contains('%').any():
                df_local[situacao] = pd.to_numeric(
                    df_local[situacao].str.replace('%', '', regex=False).str.replace(',', '.', regex=False),
                    errors='coerce')

            grouped_local = df_local.groupby("Local")[situacao].sum().sort_values(ascending=False)

        if situacao == "Turnover (%)":
            deslig_col = "Desligados no Intervalo"
            media_col = "Média Colab. no Intervalo"

            # ⚠️ Local usa todos os dados da empresa (não filtrado por setor)
            df_base_local = self.df[self.df["Local"] == local].copy() if local != "Todos" else self.df.copy()
            grouped_local_df = df_base_local.groupby("Local")[[deslig_col, media_col]].sum()
            grouped_local_df = grouped_local_df[grouped_local_df[media_col] > 0]
            grouped_local = (grouped_local_df[deslig_col] / grouped_local_df[media_col]) * 100
            grouped_local = grouped_local[grouped_local > 0].sort_values(ascending=False)

            # ✅ Setor continua com base no filtro aplicado
            if setor == "Não listar":
                grouped_setor = pd.DataFrame(columns=["Local", "Setor", situacao])
            else:
                df_base_setor = self.filtered_df.copy()
                grouped_setor_df = df_base_setor.groupby(["Local", "Setor"])[
                    [deslig_col, media_col]].sum().reset_index()
                grouped_setor_df = grouped_setor_df[grouped_setor_df[media_col] > 0].copy()
                grouped_setor_df[situacao] = (grouped_setor_df[deslig_col] / grouped_setor_df[media_col]) * 100
                grouped_setor_df = grouped_setor_df[grouped_setor_df[situacao] > 0]
                grouped_setor = grouped_setor_df[["Local", "Setor", situacao]]
                grouped_setor = grouped_setor.sort_values(by=situacao, ascending=False)

        else:
            grouped_local = pd.to_numeric(grouped_local, errors="coerce")
            grouped_local = grouped_local[grouped_local > 0]

            if setor == "Não listar":
                grouped_setor = pd.DataFrame(columns=["Local", "Setor", situacao])
            else:
                df_filt = self.filtered_df.copy()
                if df_filt[situacao].dtype == object and df_filt[situacao].str.contains('%').any():
                    df_filt[situacao] = pd.to_numeric(
                        df_filt[situacao].str.replace('%', '', regex=False).str.replace(',', '.', regex=False),
                        errors='coerce')

                grouped_setor = df_filt.groupby(["Local", "Setor"])[situacao].sum().reset_index()

                grouped_setor[situacao] = pd.to_numeric(grouped_setor[situacao], errors="coerce").fillna(0).astype(
                    float)
                grouped_setor = grouped_setor[grouped_setor[situacao] > 0]
                grouped_setor = grouped_setor.sort_values(by=situacao, ascending=False)

        if "Total" in grouped_local.index:
            grouped_local = grouped_local.drop("Total")

        # Criar listas para gráficos
        local_data = [
            (val, local, self.COLOR_MAP.get(local, "#000000"))
            for local, val in grouped_local.items()
            if val > 0
        ]

        setor_data = [
            (
                row[situacao],
                f"{row['Local']} - {row['Setor']}",
                self.SECTOR_COLOR_MAP.get(f"{row['Local']} - {row['Setor']}",
                                          self.SECTOR_COLOR_MAP.get(row['Local'], "#000000"))
            )
            for _, row in grouped_setor.iterrows()
            if row[situacao] > 0
        ]

        # Ordenar os dados
        local_data.sort(reverse=True, key=lambda x: x[0])
        setor_data.sort(reverse=True, key=lambda x: x[0])

        # Juntar os dados na ordem correta
        sorted_data = local_data + setor_data
        values, labels, colors = zip(*sorted_data) if sorted_data else ([], [], [])

        text_labels = [f"{valor:.2f}" for valor in values]
        y_labels = labels

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
        extra_space = max_text_length * 10
        max_value = max(values) if values else 0
        dynamic_width = min(1800, max(1800, max_value + extra_space))

        titulo_bruto = f"{situacao} - {self.periodo}"
        titulo_formatado = self.quebrar_titulo_auto(titulo_bruto)

        fig.update_layout(
            title=dict(text=titulo_formatado, font=dict(size=16)),
            xaxis_title="Quantidade",
            yaxis=dict(autorange="reversed"),
            height=max(500, len(y_labels) * 30),
            width=dynamic_width,
            margin=dict(l=100, r=10, t=50, b=50),
            xaxis=dict(range=[0, max(10, max_value + 10)])
        )

        if not values:
            html_content = "<h3 style='text-align:center;'>Nenhum dado disponível</h3>"
            self.web_view.setHtml(html_content)
            return

        self.current_figure = fig

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