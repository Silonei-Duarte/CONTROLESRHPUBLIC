import time
import os
import tempfile
import plotly.graph_objects as go
import plotly.io as pio
from textwrap import wrap
from PyQt6.QtCore import QUrl
import json
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QPushButton,
    QScrollArea, QSizePolicy, QMessageBox, QFileDialog, QDialog, QLineEdit, QFormLayout
)
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtCore import Qt
import pandas as pd

class AbaGraficos(QWidget):
    def __init__(self, data):
        super().__init__()
        self.updating_filters = False

        self.data = pd.DataFrame(data, columns=[
            "local", "setor", "Nome", "Código", "Rotina", "Horas", "Salário Base",
            "Salário por Hora", "Valor Final"
        ])
        self.data["Horas"] = pd.to_timedelta(self.data["Horas"], errors='coerce')

        self.filtered_data = self.data.copy()

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

        self.config_path = os.path.join(os.path.dirname(__file__), "extrasprevistos.json")
        self.PrevistoS_MAP = self.carregar_Previstos()

        # Layout principal (vertical)
        self.main_layout = QVBoxLayout()  # Inicializar o layout principal

        # Layout horizontal para filtros e botão
        top_layout = QHBoxLayout()

        # Combobox para Locais
        self.local_combobox = QComboBox()
        self.local_combobox.setFixedSize(300, 30)
        self.local_combobox.addItem("Todas")
        self.local_combobox.addItems(self.data["local"].dropna().unique())
        self.local_combobox.currentTextChanged.connect(self.update_filters)
        top_layout.addWidget(QLabel("Locais:"))
        top_layout.addWidget(self.local_combobox)

        # Combobox para Setores
        self.setor_combobox = QComboBox()
        self.setor_combobox.setFixedSize(600, 30)
        self.setor_combobox.addItem("Não listar")
        self.setor_combobox.addItem("Todas")
        self.setor_combobox.setEnabled(True)
        self.setor_combobox.currentTextChanged.connect(self.update_filters)
        top_layout.addWidget(QLabel("Setor:"))
        top_layout.addWidget(self.setor_combobox)

        # Combobox para Nomes
        self.nome_combobox = QComboBox()
        self.nome_combobox.setFixedSize(300, 30)
        self.nome_combobox.addItem("Não listar")
        self.nome_combobox.addItem("Todos")
        self.nome_combobox.setEnabled(False)
        self.nome_combobox.currentTextChanged.connect(self.update_filters)
        top_layout.addWidget(QLabel("Nomes:"))
        top_layout.addWidget(self.nome_combobox)

        # Botão para gerar PDF
        self.pdf_button = QPushButton("Gerar PDF")
        self.pdf_button.setFixedSize(120, 30)
        self.pdf_button.clicked.connect(self.generate_pdf)
        top_layout.addStretch()
        top_layout.addWidget(self.pdf_button)

        # Botão para anexar o gráfico ao e-mail existente
        self.attach_email_button = QPushButton("Anexar ao E-mail")
        self.attach_email_button.setFixedSize(150, 30)
        self.attach_email_button.clicked.connect(self.attach_graph_to_email)
        top_layout.addWidget(self.attach_email_button)

        # Botão para configurar horas previstos
        self.config_button = QPushButton("Editar Teto Previsto")
        self.config_button.setFixedSize(150, 30)
        self.config_button.clicked.connect(self.abrir_config_previstos)
        top_layout.addWidget(self.config_button)

        # Adicionar o layout horizontal ao layout principal
        self.main_layout.addLayout(top_layout)

        #  Área de rolagem para gráficos
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)

        #  WebEngineView para exibir o gráfico
        self.web_view = QWebEngineView()

        #  Ajustar a política de expansão corretamente
        self.web_view.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        #  Adicionar o web_view como widget dentro da área de rolagem
        self.scroll_area.setWidget(self.web_view)

        #  Remover rolagem horizontal e permitir apenas rolagem vertical
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        #  Adicionar a área de rolagem ao layout principal
        self.main_layout.addWidget(self.scroll_area)

        self.setLayout(self.main_layout)  # Definir o layout principal

        # Gerar o gráfico inicial
        self.update_filters()

    def carregar_Previstos(self):
        """Carrega os valores de horas extras previstos do arquivo JSON ou salva padrão se não existir."""
        padrao = {
            "Fabrica De Máquinas": 0.0,
            "Fabrica De Transportadores": 0.0,
            "Adm": 0.0,
            "Comercial": 0.0
        }

        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                QMessageBox.warning(self, "Erro", f"Falha ao ler o arquivo 'extrasprevistos.json':\n{e}")
                return padrao
        else:
            try:
                with open(self.config_path, "w", encoding="utf-8") as f:
                    json.dump(padrao, f, indent=4)
            except Exception as e:
                QMessageBox.warning(self, "Erro", f"Não foi possível criar o arquivo padrão:\n{e}")
            return padrao

    def abrir_config_previstos(self):
        """Abre janela para editar horas previstos por local."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Configurar Teto R$ Previsto")
        layout = QFormLayout(dialog)

        campos = {}
        for local in self.PrevistoS_MAP:
            campo = QLineEdit(str(self.PrevistoS_MAP[local]))
            campos[local] = campo
            layout.addRow(QLabel(local), campo)

        salvar_btn = QPushButton("Salvar")
        salvar_btn.clicked.connect(lambda: self.salvar_previstos(campos, dialog))
        layout.addRow(salvar_btn)

        dialog.setLayout(layout)
        dialog.exec()

    def salvar_previstos(self, campos, dialog):
        """Salva os valores editados no arquivo JSON."""
        try:
            novos_valores = {}
            for local, campo in campos.items():
                valor = float(campo.text().replace(",", "."))
                novos_valores[local] = valor
            self.PrevistoS_MAP = novos_valores

            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(self.PrevistoS_MAP, f, indent=4)

            QMessageBox.information(self, "Salvo", "Teto Previsto salvo com sucesso!")
            dialog.accept()
            self.plot_graph()  # Atualiza o gráfico
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao salvar: {e}")

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
        """Recebe o e-mail gerado na tela principal e armazena na aba de gráficos."""
        self.current_mail = email

    def attach_graph_to_email(self):
        """Anexa o gráfico ao e-mail corretamente, garantindo que fique no início do corpo do e-mail, antes de tudo."""

        #  Certificar que o e-mail gerado está armazenado
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
            #  Verificar se o gráfico foi gerado (se a função plot_graph foi chamada antes)
            if not hasattr(self, 'current_figure'):
                QMessageBox.critical(self, "Erro", "O gráfico não foi gerado ainda.")
                return

            #  Acessar o gráfico gerado
            fig = self.current_figure

            #  Criar um arquivo temporário para salvar a imagem do gráfico
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            image_path = temp_file.name
            temp_file.close()

            #  Salvar o gráfico como imagem PNG
            fig.write_image(image_path, format="png", engine="kaleido")

            #  Criar Content-ID único para a imagem (para exibição inline)
            cid = f"graph_image_{int(time.time())}"

            #  Adicionar a imagem ao e-mail como anexo inline
            attachment = email.Attachments.Add(image_path)
            attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", cid)

            #  Frase antes do gráfico (AGORA COM O MESMO ESTILO DA TABELA)

            frase = f"<p style='font-size:16px;'>Segue controle de horas extras:</p><br>"

            #  Modificar o corpo do e-mail: coloca a frase e o gráfico antes do conteúdo atual
            email.HTMLBody = f"{frase}<p><img src='cid:{cid}' width='1800'></p><br><br>{email.HTMLBody}"

            QMessageBox.information(self, "Sucesso", "Gráfico inserido corretamente no início do e-mail!")

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao anexar a imagem ao corpo do e-mail: {e}")

    def update_filters(self):
        """Atualiza os filtros em cascata: local -> setor -> Nome."""
        if self.updating_filters:
            return  # Evitar chamadas recursivas

        self.updating_filters = True  # Iniciar bloqueio de eventos
        try:
            # Filtrar por local
            local = self.local_combobox.currentText()
            if local == "Todas":
                self.filtered_data = self.data.copy()
            else:
                self.filtered_data = self.data[self.data["local"] == local]

            # Resetar Locais e Nomes ao mudar local
            if self.sender() == self.local_combobox:
                self.setor_combobox.setCurrentIndex(0)  # Voltar para "Não listar"
                self.nome_combobox.setCurrentIndex(0)  # Voltar para "Não listar"

            # Atualizar Locais com base no filtro de local
            self.update_setor_filter()

            # Resetar Nomes ao mudar setor
            if self.sender() == self.setor_combobox:
                self.nome_combobox.setCurrentIndex(0)  # Voltar para "Não listar"

            # Atualizar Nomes com base no filtro de setor
            self.update_nome_filter()

            # Atualizar o gráfico
            self.plot_graph()
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Ocorreu um erro ao atualizar os filtros:\n{e}")
        finally:
            self.updating_filters = False  # Liberar bloqueio de eventos

    def update_local_filter(self):
        """Atualiza os dados filtrados com base na local e ajusta os locais disponíveis."""
        local = self.local_combobox.currentText()
        if local == "Todas":
            self.filtered_data = self.data.copy()
        else:
            self.filtered_data = self.data[self.data["local"] == local]

        # Atualizar locais disponíveis
        self.update_setor_filter()

    def update_setor_filter(self):
        """Atualiza os locais disponíveis com base na local selecionada."""
        locais_disponiveis = self.filtered_data["setor"].dropna().unique()
        setor_selecionado = self.setor_combobox.currentText()

        self.setor_combobox.blockSignals(True)
        self.setor_combobox.clear()
        self.setor_combobox.addItem("Não listar")
        self.setor_combobox.addItem("Todas")
        self.setor_combobox.addItems(locais_disponiveis)
        self.setor_combobox.blockSignals(False)

        # Restaurar a seleção do setor, se for válido
        if setor_selecionado in [self.setor_combobox.itemText(i) for i in range(self.setor_combobox.count())]:
            self.setor_combobox.setCurrentText(setor_selecionado)
        else:
            self.setor_combobox.setCurrentIndex(0)  # Voltar para "Não listar" se inválido

    def update_nome_filter(self):
        """Atualiza os nomes disponíveis com base na local e setor selecionados."""
        setor = self.setor_combobox.currentText()

        # Filtrar por setor
        if setor == "Não listar":
            self.nome_combobox.setEnabled(False)
            return  # Não listar nomes, sair da função
        elif setor == "Todas":
            nomes_disponiveis = self.filtered_data["Nome"].dropna().unique()
        else:
            self.filtered_data = self.filtered_data[self.filtered_data["setor"] == setor]
            nomes_disponiveis = self.filtered_data["Nome"].dropna().unique()

        # Atualizar nomes no ComboBox
        nome_selecionado = self.nome_combobox.currentText()
        self.nome_combobox.blockSignals(True)
        self.nome_combobox.clear()
        self.nome_combobox.addItem("Não listar")
        self.nome_combobox.addItem("Todos")
        self.nome_combobox.addItems(nomes_disponiveis)
        self.nome_combobox.blockSignals(False)

        # Restaurar a seleção do nome, se for válido
        if nome_selecionado in [self.nome_combobox.itemText(i) for i in range(self.nome_combobox.count())]:
            self.nome_combobox.setCurrentText(nome_selecionado)
        else:
            self.nome_combobox.setCurrentIndex(0)  # Voltar para "Não listar" se inválido

        # Habilitar o ComboBox de nomes
        self.nome_combobox.setEnabled(True)

    def plot_graph_data(self, fig, grouped_data, offsets, bar_height, color, label_prefix=""):
        """Função para plotar dados no gráfico usando Plotly."""
        bars = []
        annotations = []

        for i, (label, row) in enumerate(grouped_data.iterrows()):
            valor_final = row["Valor Final"]
            horas = row["Horas"]

            total_horas_timedelta = pd.to_timedelta(horas, errors="coerce")
            if pd.isna(total_horas_timedelta):
                total_horas_timedelta = pd.Timedelta("0 days 00:00:00")

            total_horas_str = self.convert_timedelta_to_dias(total_horas_timedelta)

            # Criar barra horizontal
            bars.append(go.Bar(
                x=[valor_final],
                y=[i + offsets],
                orientation='h',
                marker=dict(color=color),
                font=dict(size=10),
                text=f"R$ {valor_final:.2f} / {total_horas_str}",
                textposition="outside"
            ))

            # Adicionar rótulo no eixo Y
            wrapped_label = "<br>".join(wrap(label_prefix + label, width=30))
            annotations.append(dict(
                x=0, y=i + offsets, text=wrapped_label,
                xanchor="right", showarrow=False, font=dict(size=10)
            ))

        # Adicionar as barras ao gráfico
        for bar in bars:
            fig.add_trace(bar)

        # Adicionar anotações
        fig.update_layout(annotations=annotations)

        return offsets + len(grouped_data)

    def convert_timedelta_to_dias(self, total_horas):
        """Converte timedelta para string no formato 'HHH:MM:SS'."""
        if pd.isna(total_horas) or total_horas.total_seconds() == 0:
            return ""

        total_seconds = int(total_horas.total_seconds())
        total_hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60

        return f"{total_hours:02}:{minutes:02}:{seconds:02}"

    def plot_graph(self):
        """Plota o gráfico ajustando cores por local, setor e nome corretamente com Plotly."""

        # Verifica se há dados filtrados
        if self.filtered_data.empty:
            html_content = "<h3 style='text-align:center;'>Nenhum dado disponível</h3>"
            self.web_view.setHtml(html_content)
            return

        local = self.local_combobox.currentText()
        setor = self.setor_combobox.currentText()
        nome = self.nome_combobox.currentText()

        #  Inicializa a lista de labels antes de usá-la
        y_labels = []
        values = []
        colors = []
        text_labels = []

        #  Criar a instância do gráfico antes de qualquer update_layout()
        fig = go.Figure()

        #  CALCULAR OS TOTAIS GERAIS ANTES DE FILTRAR SETOR/NOME PARA NÃO ALTERAR OS VALORES DO LOCAL
        grouped_local = self.data.groupby("local").agg({"Valor Final": "sum", "Horas": "sum"})
        grouped_local = grouped_local.sort_values("Valor Final", ascending=False)  # 🔄 Inverter ordem

        #  FILTRAGEM DINÂMICA
        filtered_data = self.data.copy()

        # Aplicar filtro de local se selecionado
        if local != "Todas":
            filtered_data = filtered_data[filtered_data["local"] == local]
            grouped_local = grouped_local.loc[[local]]

        # Aplicar filtro de setor se selecionado
        grouped_setor = None
        if setor != "Não listar":
            if setor != "Todas":
                filtered_data = filtered_data[filtered_data["setor"] == setor]
            grouped_setor = filtered_data.groupby(["local", "setor"]).agg({"Valor Final": "sum", "Horas": "sum"})
            grouped_setor = grouped_setor.sort_values("Valor Final", ascending=False)

        # Aplicar filtro de nome se selecionado
        grouped_nome = None
        if nome != "Não listar":
            if nome != "Todos":
                filtered_data = filtered_data[filtered_data["Nome"] == nome]
            grouped_nome = filtered_data.groupby("Nome").agg({"Valor Final": "sum", "Horas": "sum"})
            grouped_nome = grouped_nome.sort_values("Valor Final", ascending=False)

        #  PLOTAR LOCAIS PRIMEIRO + Previsto e Saldo
        if not grouped_local.empty:
            for local_nome in grouped_local.index:
                local_color = self.COLOR_MAP.get(local_nome, "#808080")
                row = grouped_local.loc[local_nome]
                valor_final = row["Valor Final"]
                horas_str = self.convert_timedelta_to_dias(pd.to_timedelta(row["Horas"], errors="coerce"))

                # 1️⃣ Valor Real
                y_labels.append(local_nome)
                values.append(valor_final)
                colors.append(local_color)
                texto_horas = f" / {horas_str}" if horas_str else ""
                text_labels.append(f"Realizado: R$ {valor_final:.2f}{texto_horas}")

                valor_ref = self.PrevistoS_MAP.get(local_nome, 0)

                if valor_ref > 0:
                    # 2️⃣ Previsto - em laranja
                    y_labels.append(f"{local_nome} (Previsto)")
                    values.append(valor_ref)
                    colors.append("#ff9900")
                    text_labels.append(
                        f"Previsto: R$ {valor_ref:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    )

                    # 3️⃣ Saldo
                    saldo = valor_ref - valor_final
                    saldo_color = "#007f01" if saldo > 0 else "red"
                    y_labels.append(f"{local_nome} (Saldo)")
                    values.append(saldo)
                    colors.append(saldo_color)
                    text_labels.append(
                        f"Saldo: R$ {saldo:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    )

        #  PLOTAR SETORES DEPOIS
        if grouped_setor is not None and not grouped_setor.empty:
            for (local_nome, setor_nome), row in grouped_setor.iterrows():
                setor_unico = f"{str(local_nome)} - {str(setor_nome)}"
                setor_color = self.SECTOR_COLOR_MAP.get(setor_unico, self.SECTOR_COLOR_MAP.get(local_nome))
                valor_final = row["Valor Final"]
                horas_str = self.convert_timedelta_to_dias(pd.to_timedelta(row["Horas"], errors="coerce"))
                y_labels.append(setor_unico)
                values.append(valor_final)
                colors.append(setor_color)
                texto_horas = f" / {horas_str}" if horas_str else ""
                text_labels.append(f"Realizado: R$ {valor_final:.2f}{texto_horas}")

        #  PLOTAR NOMES POR ÚLTIMO
        if grouped_nome is not None and not grouped_nome.empty:
            for i, (label, row) in enumerate(grouped_nome.iterrows()):
                valor_final = row["Valor Final"]
                horas_str = self.convert_timedelta_to_dias(pd.to_timedelta(row["Horas"], errors="coerce"))
                y_labels.append(label)
                values.append(valor_final)
                colors.append(self.NAME_COLOR)
                texto_horas = f" / {horas_str}" if horas_str else ""
                text_labels.append(f"Realizado: R$ {valor_final:.2f}{texto_horas}")

        #  Criar gráfico com Plotly
        fig.add_trace(go.Bar(
            y=y_labels,
            x=values,
            orientation='h',
            marker=dict(color=colors),
            text=text_labels,
            textposition='outside'
        ))

        # Calcular o comprimento máximo do texto para definir a largura da área de plotagem
        max_text_length = max(len(txt) for txt in text_labels) if text_labels else 10
        extra_space = max_text_length * 16  # Ajuste o multiplicador para garantir o espaço necessário para o texto

        # Ajustar largura do gráfico considerando o maior valor e o texto
        max_value_barras = max(values) if values else 0
        max_value_metas = max(self.PrevistoS_MAP.values())
        max_value = max(max_value_barras, max_value_metas)

        dynamic_width = min(1800, max(1800, max_value + extra_space))

        x_min = min(values)
        x_max = max(10, max_value + 10000)

        # Se houver valor negativo, adiciona margem de 5000 à esquerda do menor valor
        x_range_min = x_min - 5000 if x_min < 0 else 0

        fig.update_layout(
            title="Gráfico",
            xaxis_title="Valor Final",
            yaxis=dict(autorange="reversed"),
            height=max(500, len(y_labels) * 30),
            width=dynamic_width,
            margin=dict(l=200, r=10, t=50, b=50),
            xaxis=dict(
                range=[x_range_min, x_max],
                fixedrange=True
            )
        )

        #  Se `values` estiver vazio, exibir uma mensagem de "Nenhum dado disponível"
        if not values:
            html_content = "<h3 style='text-align:center;'>Nenhum dado disponível</h3>"
            self.web_view.setHtml(html_content)
            return

        self.current_figure = fig

        #  Salvar HTML temporário e carregar no WebView
        temp_path = os.path.join(tempfile.gettempdir(), "graph_all_situations.html")
        pio.write_html(fig, temp_path)

        self.web_view.setUrl(QUrl.fromLocalFile(temp_path))

