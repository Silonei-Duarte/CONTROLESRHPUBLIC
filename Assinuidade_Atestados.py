import os
import json
import pandas as pd
from datetime import datetime, time, timedelta
from main import ControlesRH
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QLabel, QPushButton, QLineEdit,
                             QHBoxLayout, QFileDialog, QTableWidget, QTableWidgetItem, QMessageBox, QDialog)
import win32com.client
from Assinuidade_Atestados_Grafico import AbaGraficoAssinuidade

class AbaAssiduidade_Atestado(QWidget):
    def __init__(self, janela_principal):
        super().__init__()
        self.janela_principal = janela_principal  # Referência para controlefrequencia
        self.current_mail = None  # Inicialmente sem e-mail

        # Arquivo de configuração
        self.config_file = os.path.join(os.path.dirname(__file__), "frequencia.json")
        self.folder_path = ""

        # Layout principal
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        # Seleção de pasta única
        folder_layout = QHBoxLayout()
        self.hrap_line_edit = QLineEdit()
        self.hrap_line_edit.setPlaceholderText("Selecione a pasta das planilhas")
        self.hrap_line_edit.setReadOnly(True)

        folder_button = QPushButton("...")
        folder_button.setFixedSize(120, 30)
        folder_button.clicked.connect(self.select_folder)

        folder_layout.addWidget(QLabel("Pasta das Planilhas"))
        folder_layout.addWidget(self.hrap_line_edit)
        folder_layout.addWidget(folder_button)
        layout.addLayout(folder_layout)

        # Layout dos botões múltiplos (substitua o antigo bloco do botão único)
        botoes_layout = QHBoxLayout()

        # Botão para HRAP604
        self.calculate_hrap_btn = QPushButton("Calcular HRAP604")
        self.calculate_hrap_btn.setFixedSize(150, 30)
        self.calculate_hrap_btn.clicked.connect(lambda: self.carregar_dados_historico("HRAP604.xlsx"))
        botoes_layout.addWidget(self.calculate_hrap_btn)

        # Botão para AUSENCIAS DIARIO
        self.calculate_diario_btn = QPushButton("Cálculo Ontem")
        self.calculate_diario_btn.setFixedSize(150, 30)
        self.calculate_diario_btn.clicked.connect(lambda: self.carregar_dados_historico("AUSENCIAS DIARIO.xlsx"))
        botoes_layout.addWidget(self.calculate_diario_btn)

        # Botão para AUSENCIAS SEMANAL
        self.calculate_semanal_btn = QPushButton("Cálculo Semanal")
        self.calculate_semanal_btn.setFixedSize(150, 30)
        self.calculate_semanal_btn.clicked.connect(lambda: self.carregar_dados_historico("AUSENCIAS SEMANAL.xlsx"))
        botoes_layout.addWidget(self.calculate_semanal_btn)

        # Botão para AUSENCIAS MENSAL
        self.calculate_mensal_btn = QPushButton("Cálculo Mensal")
        self.calculate_mensal_btn.setFixedSize(150, 30)
        self.calculate_mensal_btn.clicked.connect(lambda: self.carregar_dados_historico("AUSENCIAS MENSAL.xlsx"))
        botoes_layout.addWidget(self.calculate_mensal_btn)

        # Espaço para alinhamento
        botoes_layout.addStretch()

        self.email_detalhado = QPushButton("E-mail Detalhado")
        self.email_detalhado.setFixedSize(120, 30)
        self.email_detalhado.clicked.connect(self.enviar_email_detalhado)
        botoes_layout.addWidget(self.email_detalhado)

        # Botão para abrir gráficos
        self.grafico_button = QPushButton("Gráficos")
        self.grafico_button.setFixedSize(120, 30)
        self.grafico_button.clicked.connect(self.abrir_graficos)
        botoes_layout.addWidget(self.grafico_button)

        # Botão Voltar (já existente)
        self.btn_voltar = QPushButton("Voltar ao Menu")
        self.btn_voltar.setFixedSize(120, 30)
        self.btn_voltar.clicked.connect(self.voltar_menu)
        botoes_layout.addWidget(self.btn_voltar)

        layout.addLayout(botoes_layout)

        # Layout horizontal para status e período lado a lado
        status_periodo_layout = QHBoxLayout()

        # Alinha os labels totalmente à esquerda
        status_periodo_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)

        # Label do Status
        self.status_label = QLabel("Nenhum cálculo realizado ainda")
        self.status_label.setStyleSheet("color: red; font-size: 12px;")
        status_periodo_layout.addWidget(self.status_label)

        # Label do Período Analisado
        self.periodo_label = QLabel("Período Analisado: -")
        status_periodo_layout.addWidget(self.periodo_label)

        # Adiciona o layout horizontal ao layout principal
        layout.addLayout(status_periodo_layout)

        # Campo de filtro
        self.search_field = QLineEdit()
        self.search_field.setPlaceholderText("Filtrar usando / entre dados! Tambem pode digitar COLUNA:DADO para filtrar pela expressão exata na coluna")
        self.search_field.textChanged.connect(self.apply_global_filter)
        layout.addWidget(self.search_field)

        # Tabela para exibir os dados processados
        self.tableWidget = QTableWidget()
        layout.addWidget(self.tableWidget)

        # Evento de duplo clique na tabela
        self.tableWidget.cellDoubleClicked.connect(self.abrir_detalhes_colaborador)
        self.df_original = None  # Guarda a planilha carregada

        # Carrega a configuração ao iniciar
        self.load_config()

        # Criar o dicionário de locais diretamente na `__init__`
        self.dicionario_locais = {}

        # Define o layout da aba
        self.setLayout(layout)

    def select_folder(self):
        """Abre o diálogo para selecionar a pasta onde as planilhas estão localizadas."""
        folder_path = QFileDialog.getExistingDirectory(self, "Selecionar Pasta com Planilhas")
        if folder_path:
            self.folder_path = folder_path
            self.hrap_line_edit.setText(folder_path)
            self.save_config()

    def save_config(self):
        """Salva o caminho da pasta das planilhas em um arquivo JSON."""
        config = {"folder_path": self.folder_path}
        with open(self.config_file, "w") as f:
            json.dump(config, f, indent=4)

    def load_config(self):
        """Carrega o caminho da pasta das planilhas de um arquivo JSON."""
        if os.path.exists(self.config_file):
            with open(self.config_file, "r") as f:
                config = json.load(f)
                self.folder_path = config.get("folder_path", "")

                if self.folder_path:
                    self.hrap_line_edit.setText(self.folder_path)

    def open_save_with_excel(self, file_path):
        """Abre e salva a planilha usando o Excel via win32com sem fechar outras instâncias."""
        excel = win32com.client.Dispatch("Excel.Application")

        # Verifica se o Excel já estava rodando antes
        was_excel_open = excel.Workbooks.Count > 0

        excel.Visible = False  # Não exibir a interface do Excel

        try:
            workbook = excel.Workbooks.Open(file_path)  # Usar caminho completo
            workbook.Save()  # Salva a planilha
            workbook.Close(SaveChanges=False)
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao abrir/salvar a planilha no Excel: {e}")
        finally:
            # Fecha o Excel apenas se ele não estava aberto antes do script rodar
            if not was_excel_open:
                excel.Quit()

    def carregar_dados_historico(self, filename):
        """Carrega a planilha especificada, calcula horas Trabalhando/Atestados e também Dias Trabalhando/Dias Atestado."""
        if not self.folder_path:
            self.hrap_line_edit.setText("Selecione a pasta primeiro!")
            return

        file_path = os.path.join(self.folder_path, filename)
        if not os.path.exists(file_path):
            self.status_label.setText(f"Arquivo: {filename} não encontrado!")
            self.status_label.setStyleSheet("color: red; font-size: 12px; ")
            self.periodo_label.setText(f"00:00:00")
            return

        self.open_save_with_excel(file_path)

        locais_path = os.path.join(self.folder_path, "LOCAIS.xlsx")
        if os.path.exists(locais_path):
            locais_df = pd.read_excel(locais_path, header=None, dtype=str)
            self.dicionario_locais = dict(zip(locais_df[0].str.strip(), locais_df[1].str.strip()))
        else:
            self.dicionario_locais = {}

        df = pd.read_excel(file_path, header=None, dtype=str)
        self.df_original = df.copy()

        periodo_inicio = df.iloc[0, 10]
        periodo_fim = df.iloc[0, 11]

        mapa_situacoes = {
            "Atestado Medico": "Dias Atestado",
            "Atestado Medico Noturno": "Dias Atestado",
            "Licença Médica 15 dias": "Dias Atestado",
            "Trabalhando": "Dias Trabalhando",
            "Trabalho Noturno": "Dias Trabalhando",
            "Férias": "Dias Trabalhando",
            "Férias Noturnas": "Dias Trabalhando",
            "Acidente Trabalho": "Dias Trabalhando",
            "Acidente Trabalho Noturno": "Dias Trabalhando",
            "Licença Paternidade": "Dias Trabalhando",
            "Ferias Coletivas": "Dias Trabalhando",
            "Férias Coletivas Noturnas": "Dias Trabalhando",
            "Aviso Previo Trab.": "Dias Trabalhando",
            "Aviso Prévio Trab. Noturno": "Dias Trabalhando",
            "Viagem a Servico": "Dias Trabalhando",
            "Ausencia p/Cursos/Treinamentos": "Dias Trabalhando",
            "Saída Intermediária BH": "Dias Trabalhando",
            "Saída Antecipada BH": "Dias Trabalhando",
            "Saída Antecipada Noturna BH": "Dias Trabalhando",
            "Saída Intermediária Noturna BH": "Dias Trabalhando",
            "Atraso BH": "Dias Trabalhando",
            "Atraso Noturno BH": "Dias Trabalhando",
            "Faltas Noturnas BH": "Dias Trabalhando",
        }

        dados_finais = {}
        dias_registrados = {}
        nome_colaborador = None
        coletando = False

        for _, row in df.iterrows():
            if pd.notna(row[0]) and str(row[0]).strip().lower() == "total":
                coletando = False
                continue

            if pd.notna(row[2]):  # Novo colaborador
                nome_colaborador = str(row[2]).strip()
                coletando = True
                numloc = str(row[3]).strip() if pd.notna(row[3]) else ""
                local, setor = self.buscar_setor_e_local(numloc, self.dicionario_locais)
                if nome_colaborador not in dados_finais:
                    dados_finais[nome_colaborador] = {"Local": local, "Setor": setor}
                if nome_colaborador not in dias_registrados:
                    dias_registrados[nome_colaborador] = {"Dias Trabalhando": set(), "Dias Atestado": set()}
                continue

            if not coletando or not nome_colaborador:
                continue

            if pd.notna(row[6]):
                descricao_situacao = str(row[6]).strip()
                if descricao_situacao not in mapa_situacoes:
                    continue

                tipo_dia = mapa_situacoes[descricao_situacao]
                data_str = str(row[7]).strip() if pd.notna(row[7]) else ""

                if tipo_dia == "Dias Atestado":
                    if data_str != "-" and data_str != "":
                        try:
                            dia = pd.to_datetime(row[7]).date()
                            dias_registrados[nome_colaborador][tipo_dia].add(dia)
                        except:
                            dias_registrados[nome_colaborador][tipo_dia].add("data_indefinida")
                    else:
                        # Conta 1 dia genérico para atestados sem data
                        dias_registrados[nome_colaborador][tipo_dia].add("data_indefinida")
                else:
                    # Para Dias Trabalhando conta só datas válidas
                    if data_str != "" and data_str != "-":
                        try:
                            dia = pd.to_datetime(row[7]).date()
                            dias_registrados[nome_colaborador][tipo_dia].add(dia)
                        except:
                            pass

                # Soma de horas
                if pd.notna(row[9]):
                    valor = self.ajustar_horas_excel(row[9])
                    tipo_hora = "Trabalhando" if mapa_situacoes[
                                                     descricao_situacao] == "Dias Trabalhando" else "Atestados"
                    if tipo_hora in dados_finais[nome_colaborador]:
                        dados_finais[nome_colaborador][tipo_hora] = self.somar_tempos(
                            dados_finais[nome_colaborador][tipo_hora], valor
                        )
                    else:
                        dados_finais[nome_colaborador][tipo_hora] = valor

        # Converte para DataFrame
        df_resultado = pd.DataFrame.from_dict(dados_finais, orient="index").reset_index()
        df_resultado.rename(columns={"index": "Colaborador"}, inplace=True)

        # Garante colunas de horas
        for c in ["Trabalhando", "Atestados"]:
            if c not in df_resultado.columns:
                df_resultado[c] = "00:00:00"

        # Adiciona colunas de dias
        df_resultado["Dias Trabalhando"] = [
            len(dias_registrados[n]["Dias Trabalhando"]) for n in df_resultado["Colaborador"]
        ]
        df_resultado["Dias Atestado"] = [
            len(dias_registrados[n]["Dias Atestado"]) for n in df_resultado["Colaborador"]
        ]

        # Reorganiza colunas
        colunas_ordenadas = ["Local", "Setor", "Colaborador", "Trabalhando", "Atestados", "Dias Trabalhando",
                             "Dias Atestado"]
        df_resultado = df_resultado[colunas_ordenadas]

        # Mostra período
        periodo_inicio_fmt = pd.to_datetime(periodo_inicio).strftime("%d/%m/%Y")
        periodo_fim_fmt = pd.to_datetime(periodo_fim).strftime("%d/%m/%Y")
        self.periodo_label.setText(f"Período Analisado: {periodo_inicio_fmt} a {periodo_fim_fmt}")

        for col in ["Trabalhando", "Atestados"]:
            df_resultado[col] = df_resultado[col].fillna("00:00:00")

        # Exibe na tabela
        self.mostrar_resultado(df_resultado)

        # Status
        if filename == "HRAP604.xlsx":
            texto = f"Cálculo HRAP604"
        elif filename == "AUSENCIAS DIARIO.xlsx":
            texto = f"Cálculo Diario - Atualizado a cada 1 hora"
        elif filename == "AUSENCIAS SEMANAL.xlsx":
            texto = f"Cálculo Semanal - Atualizado a cada 1 hora"
        elif filename == "AUSENCIAS MENSAL.xlsx":
            texto = f"Cálculo Mensal - Atualizado a cada 1 hora"

        self.status_label.setText(texto)
        self.status_label.setStyleSheet("color: red; font-size: 12px; ")

    def adicionar_colunas_indicadores(self, df):
        """Cria colunas de Dias Trabalhando e Dias Atestado no lugar das colunas de percentual."""

        dias_trabalhando = []
        dias_atestado = []

        for _, row in df.iterrows():
            colaborador = row["Colaborador"]
            df_colab = self.df_original[self.df_original[2] == colaborador]

            trab_set = set()
            ates_set = set()

            for _, linha in df_colab.iterrows():
                if pd.isna(linha[7]) or pd.isna(linha[6]):
                    continue
                try:
                    dia = pd.to_datetime(linha[7]).date()
                except Exception:
                    continue

                situacao = str(linha[6]).strip()
                if situacao in ["Trabalhando", "Trabalho Noturno"]:
                    trab_set.add(dia)
                elif situacao in ["Atestado Medico", "Atestado Medico Noturno", "Licença Médica 15 dias"]:
                    ates_set.add(dia)

            dias_trabalhando.append(len(trab_set))
            dias_atestado.append(len(ates_set))

        # Cria as novas colunas
        df["Dias Trabalhando"] = dias_trabalhando
        df["Dias Atestado"] = dias_atestado

        return df

    def mostrar_resultado(self, df):
        """Exibe o DataFrame na tabela da interface garantindo reset completo, inclusive na ordenação."""

        self.tableWidget.setSortingEnabled(False)
        self.tableWidget.clear()
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)

        self.tableWidget.setRowCount(df.shape[0])
        self.tableWidget.setColumnCount(df.shape[1])
        self.tableWidget.setHorizontalHeaderLabels(df.columns)
        self.tableWidget.verticalHeader().setVisible(False)
        self.tableWidget.setStyleSheet("QTableWidget {gridline-color: black;}")

        self.tableWidget.horizontalHeader().setSortIndicator(-1, Qt.SortOrder.AscendingOrder)

        cores = [QColor("#FFFFFF"), QColor("#D3D3D3")]
        coluna_cor = {}
        cor_idx = 0
        for col in range(3, df.shape[1]):
            if (col - 3) % 2 == 0:
                cor_idx = (cor_idx + 1) % 2
            coluna_cor[col] = cores[cor_idx]
        for col in range(3):
            coluna_cor[col] = QColor("#FFFFFF")

        # índice da coluna Dias Atestado
        col_idx_atestado = df.columns.get_loc("Dias Atestado")

        for row in range(df.shape[0]):
            valor_atestado = 0
            try:
                valor_atestado = int(df.iat[row, col_idx_atestado])
            except:
                pass

            for col in range(df.shape[1]):
                col_name = df.columns[col]
                valor = df.iat[row, col]
                valor_str = str(valor)

                if col in [0, 1, 2]:
                    item = QTableWidgetItem(valor_str)
                else:
                    if ":" in valor_str and valor_str.count(":") == 2:
                        try:
                            h, m, s = map(int, valor_str.split(":"))
                            segundos_totais = h * 3600 + m * 60 + s
                            item = QTableWidgetItem()
                            item.setData(Qt.ItemDataRole.EditRole, segundos_totais)
                            item.setText(valor_str)
                        except:
                            item = QTableWidgetItem(valor_str)
                    elif col_name.endswith("(%)"):
                        try:
                            item = QTableWidgetItem(f"{float(valor):.2f}")
                            item.setData(Qt.ItemDataRole.EditRole, float(valor))
                        except:
                            item = QTableWidgetItem("0.00")
                    else:
                        item = QTableWidgetItem(valor_str)

                item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                item.setForeground(QColor("#000000"))

                # cor padrão de fundo
                if col in coluna_cor:
                    item.setBackground(coluna_cor[col])

                #  se tiver Dias Atestado > 0 pinta Local, Setor, Colaborador e Dias Atestado
                if valor_atestado > 0 and col_name in ["Local", "Setor", "Colaborador", "Dias Atestado"]:
                    item.setBackground(QColor("#FFCCCC"))

                self.tableWidget.setItem(row, col, item)

        fonte = self.tableWidget.font()
        fonte.setPointSize(8)
        self.tableWidget.setFont(fonte)
        self.tableWidget.resizeColumnsToContents()

        largura_extra = 15
        for col in range(self.tableWidget.columnCount()):
            largura_atual = self.tableWidget.columnWidth(col)
            if col > 2:
                self.tableWidget.setColumnWidth(col, largura_atual + largura_extra)
            else:
                self.tableWidget.setColumnWidth(col, largura_atual)

        self.tableWidget.setSortingEnabled(True)
        self.atualizar_total()

    def somar_tempos(self, tempo1, tempo2):
        """Soma dois tempos no formato HH:MM:SS"""
        try:
            t1 = pd.to_timedelta(tempo1)
            t2 = pd.to_timedelta(tempo2)
            total = t1 + t2
            horas = int(total.total_seconds() // 3600)
            minutos = int((total.total_seconds() % 3600) // 60)
            return f"{horas:02}:{minutos:02}:00"
        except:
            return tempo1  # fallback se erro

    def apply_global_filter(self, text):
        """Filtra as linhas da tabela com base na pesquisa parcial e permite busca por coluna específica (ex: Setor:"Produção")."""

        terms = [term.strip() for term in text.split('/') if term.strip()]  # Normaliza os termos

        if not terms:  # Se não houver termos, exibe todas as linhas
            for row in range(self.tableWidget.rowCount()):
                self.tableWidget.setRowHidden(row, False)
            self.atualizar_total()
            return

        # Criar um dicionário com os nomes das colunas para facilitar a busca específica
        headers = {}
        for col in range(self.tableWidget.columnCount()):
            header_item = self.tableWidget.horizontalHeaderItem(col)
            if header_item:  # Apenas adiciona colunas com nome
                headers[header_item.text().strip().lower()] = col

        # Iterar por todas as linhas da tabela
        for row in range(self.tableWidget.rowCount()):
            row_matches_all_terms = True  # Assume que a linha deve aparecer

            for term in terms:
                term_match = False  # Assume que o termo **não** foi encontrado na linha

                # Verifica se o termo segue o formato "coluna:valor" para busca exata
                if ":" in term:
                    coluna_nome, valor_busca = map(str.strip, term.split(":", 1))
                    coluna_nome = coluna_nome.lower()

                    # Se a coluna existir, busca SOMENTE nessa coluna
                    if coluna_nome in headers:
                        col_idx = headers[coluna_nome]
                        item = self.tableWidget.item(row, col_idx)

                        if item and item.text().strip().lower() == valor_busca.lower():
                            term_match = True  # Encontrou o termo exato na coluna correta

                else:
                    # Busca geral (qualquer coluna até a coluna "Colaborador" inclusa)
                    for col in range(3):  # Busca apenas nas colunas 0 a 2 (Local, Setor, Colaborador)
                        item = self.tableWidget.item(row, col)
                        if item and term.lower() in item.text().strip().lower():
                            term_match = True  # Encontrou o termo em alguma dessas colunas
                            break  # Já encontrou, não precisa verificar o resto

                if not term_match:
                    row_matches_all_terms = False
                    break  # Se um termo não foi encontrado, a linha não será exibida

            self.tableWidget.setRowHidden(row,
                                          not row_matches_all_terms)  # Esconde apenas se **não** corresponde a todos os termos

        self.atualizar_total()  # Atualiza a linha de total após o filtro

    def atualizar_total(self):
        """Atualiza a linha de total com base nos dados visíveis na tabela,
        somando colunas de horas (HHH:MM:SS) e colunas de dias (inteiros)."""

        if self.tableWidget.rowCount() == 0:
            return

        num_cols = self.tableWidget.columnCount()
        totais_tempo = {}
        totais_dias = {}

        # Inicializa os acumuladores
        for col in range(num_cols):
            totais_tempo[col] = timedelta()
            totais_dias[col] = 0

        # Soma os valores visíveis da tabela
        for row in range(self.tableWidget.rowCount()):
            if self.tableWidget.isRowHidden(row):
                continue

            for col in range(num_cols):
                item = self.tableWidget.item(row, col)
                if item:
                    texto = item.text().strip()

                    # Colunas de hora (HH:MM:SS)
                    if ":" in texto and texto.count(":") == 2:
                        try:
                            h, m, s = map(int, texto.split(":"))
                            totais_tempo[col] += timedelta(hours=h, minutes=m, seconds=s)
                        except:
                            pass
                    # Colunas de dias
                    elif self.tableWidget.horizontalHeaderItem(col).text() in ["Dias Trabalhando", "Dias Atestado"]:
                        try:
                            totais_dias[col] += int(texto)
                        except:
                            pass

        # Remove linha de total antiga, se existir
        if self.tableWidget.rowCount() > 0 and self.tableWidget.item(self.tableWidget.rowCount() - 1, 0) and \
                self.tableWidget.item(self.tableWidget.rowCount() - 1, 0).text() == "Total":
            self.tableWidget.removeRow(self.tableWidget.rowCount() - 1)

        # Adiciona linha de totais
        row_total = self.tableWidget.rowCount()
        self.tableWidget.insertRow(row_total)

        for col in range(num_cols):
            if col == 0:
                item = QTableWidgetItem("Total")
            elif self.tableWidget.horizontalHeaderItem(col).text() in ["Dias Trabalhando", "Dias Atestado"]:
                item = QTableWidgetItem(str(totais_dias[col]))
            elif totais_tempo[col] != timedelta():
                total = totais_tempo[col]
                total_segundos = int(total.total_seconds())
                h = total_segundos // 3600
                m = (total_segundos % 3600) // 60
                s = total_segundos % 60
                item = QTableWidgetItem(f"{h:02}:{m:02}:{s:02}")
            else:
                item = QTableWidgetItem("")

            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item.setBackground(Qt.GlobalColor.yellow)
            item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
            self.tableWidget.setItem(row_total, col, item)

    def ajustar_horas_excel(self, valor):
        """ Ajusta corretamente as horas acumuladas considerando dias acumulados como múltiplos de 24h. """
        if pd.isna(valor) or valor is None or valor == "nan":
            return "-"

        if isinstance(valor, datetime):  # Se for um datetime (com data completa)
            dia_real = valor.day  # Captura o dia corretamente
            horas = valor.hour
            minutos = valor.minute
            segundos = round(valor.second + valor.microsecond / 1_000_000)  # Arredonda corretamente os segundos
            horas_totais = (dia_real * 24) + horas  # Cada dia equivale a 24 horas

        elif isinstance(valor, time):  # Se for apenas um objeto `time`
            horas_totais = valor.hour
            minutos = valor.minute
            segundos = round(valor.second + valor.microsecond / 1_000_000)

        elif isinstance(valor, str):  # Se for string "YYYY-MM-DD HH:MM:SS" ou "HH:MM:SS"
            try:
                if "." in valor:
                    valor_dt = datetime.strptime(valor, "%Y-%m-%d %H:%M:%S.%f")  # Se tem microssegundos
                else:
                    valor_dt = datetime.strptime(valor, "%Y-%m-%d %H:%M:%S")  # Se NÃO tem microssegundos

                dia_real = valor_dt.day
                horas_totais = (dia_real * 24) + valor_dt.hour
                minutos = valor_dt.minute
                segundos = round(valor_dt.second + valor_dt.microsecond / 1_000_000)

            except ValueError:
                partes = valor.split(":")
                horas_totais, minutos = map(int, partes[:2])
                segundos = round(float(partes[2]))

        else:
            return "-"

        # Ajuste para sempre mostrar segundos como 00
        if segundos > 0:
            minutos += 1
        segundos = 0

        if minutos >= 60:
            horas_totais += 1
            minutos = 0

        return f"{horas_totais:02}:{minutos:02}:00"

    def buscar_setor_e_local(self, numloc, dicionario_locais):
        """Busca o setor e local baseado no NUMLOC diretamente no dicionário de locais."""
        codloc = numloc.strip()

        # Define valores padrão caso não encontre
        local, setor_final = "Não encontrado", "Não encontrado"

        # Se não encontrar o codloc no dicionário, já retorna
        if not codloc:
            return local, setor_final

            # Tenta encontrar o CODLOC exato na planilha, reduzindo se necessário
        while codloc:
            for codloc_planilha, descricao in dicionario_locais.items():
                if codloc_planilha.startswith(codloc):  # Verifica se a linha começa com o CODLOC procurado
                    partes_descricao = descricao.split(",", 1)  # Divide apenas na primeira vírgula
                    local = partes_descricao[0].strip()  # A primeira parte define o Local

                    # O setor final será a última parte do CODLOC na planilha
                    setor_final = codloc_planilha.rsplit(",", 1)[-1].strip() if "," in codloc_planilha else local
                    return local, setor_final # Retorna imediatamente se encontrou

            # Se não encontrou, tenta reduzir removendo a última parte do CODLOC
            if "." in codloc:
                codloc = ".".join(codloc.split(".")[:-1])  # Remove a última parte do código
            else:
                break  # Sai do loop se não puder mais reduzir

        return local, setor_final  # Agora o "Não encontrado" só aparece uma única vez

    def voltar_menu(self):
        from main import ControlesRH  # Garante que não haverá problemas de importação circular
        self.menu_principal = ControlesRH()
        self.menu_principal.show()
        self.window().close()

    def enviar_email_detalhado(self):
        """Gera e abre um e-mail no Outlook com base no detalhamento do relatório filtrado na tela."""

        if self.tableWidget.rowCount() == 0:
            QMessageBox.warning(self, "Aviso", "Nenhum dado carregado para enviar por e-mail.")
            return

        periodo_analisado = self.periodo_label.text().replace("Período Analisado: ", "").strip()

        cores_locais = {
            "Fabrica De Máquinas": "#003756",
            "Fabrica De Transportadores": "#ffc62e",
            "Adm": "#009c44",
            "Comercial": "#919191"
        }

        dados_agrupados = {}  # {Local: {Setor: [linhas]}}

        colunas_exibidas = []
        for col in range(self.tableWidget.columnCount()):
            nome_coluna = self.tableWidget.horizontalHeaderItem(col).text()
            if nome_coluna not in ["Local", "Setor", "Colaborador"]:  # exclui também Colaborador
                colunas_exibidas.append((col, nome_coluna))

        if not colunas_exibidas:
            QMessageBox.warning(self, "Aviso", "Nenhuma coluna válida encontrada para envio de e-mail.")
            return

        num_rows = self.tableWidget.rowCount()
        if num_rows > 0 and self.tableWidget.item(num_rows - 1, 0) and \
                self.tableWidget.item(num_rows - 1, 0).text() == "Total":
            num_rows -= 1  # Ignora linha de total

        for row in range(num_rows):
            if self.tableWidget.isRowHidden(row):
                continue

            # pega valor da coluna "Dias Atestado"
            col_idx_atestado = None
            for c in range(self.tableWidget.columnCount()):
                if self.tableWidget.horizontalHeaderItem(c).text() == "Dias Atestado":
                    col_idx_atestado = c
                    break

            if col_idx_atestado is None:
                continue

            item_atestado = self.tableWidget.item(row, col_idx_atestado)
            valor_atestado = int(item_atestado.text()) if item_atestado and item_atestado.text().isdigit() else 0

            # pula se não tem atestado
            if valor_atestado <= 0:
                continue

            local = self.tableWidget.item(row, 0).text().strip()
            setor = self.tableWidget.item(row, 1).text().strip()
            colaborador = self.tableWidget.item(row, 2).text().strip()

            if local not in dados_agrupados:
                dados_agrupados[local] = {}
            if setor not in dados_agrupados[local]:
                dados_agrupados[local][setor] = []

            linha = [(colaborador, "#ffffff")]

            for col, nome_coluna in colunas_exibidas:
                item = self.tableWidget.item(row, col)
                valor = item.text() if item else "-"
                if not valor or valor.strip() in ["", "0", "00:00:00", "0.00"]:
                    valor = "-"
                cor_fundo = item.background().color().name() if item else "#ffffff"
                linha.append((valor, cor_fundo))

            dados_agrupados[local][setor].append(linha)

        # Monta corpo HTML do e-mail
        corpo_email = f"<p style='font-size:16px;'>Detalhamento do relatório de frequência de {periodo_analisado}:</p>"

        for local in sorted(dados_agrupados.keys()):
            cor_local = cores_locais.get(local, "#000000")
            corpo_email += f'<h3 style="color:{cor_local}; font-size: 16px;">{local}</h3>'

            for setor in sorted(dados_agrupados[local].keys()):
                corpo_email += f'<h4 style="margin-left:10px; font-size:14px;">Setor: {setor}</h4>'
                corpo_email += """
                <table border="1" cellspacing="0" cellpadding="3"
                       style="border-collapse: collapse; width: auto; font-size: 14px; margin-left:10px;">
                    <tr style="background-color: #f2f2f2; text-align: center;">
                        <th>Colaborador</th>"""
                for _, nome_coluna in colunas_exibidas:
                    corpo_email += f'<th>{nome_coluna}</th>'
                corpo_email += "</tr>"

                for linha in dados_agrupados[local][setor]:
                    corpo_email += "<tr>"
                    for valor, cor in linha:
                        corpo_email += f'<td style="padding: 3px; background-color: {cor}; text-align: center;">{valor}</td>'
                    corpo_email += "</tr>"

                corpo_email += "</table><br>"

        # Envia via Outlook
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            email = outlook.CreateItem(0)
            email.Subject = f"Relatório Detalhado de Histórico de Frequência - {periodo_analisado}"
            signature = email.HTMLBody
            email.HTMLBody = corpo_email + "<br><br>" + signature

            self.janela_principal.current_mail = email
            self.current_mail = email
            email.Display()
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao abrir o Outlook: {e}")

    def abrir_detalhes_colaborador(self, row, _):
        """Abre uma nova janela com os dados originais do colaborador ao dar duplo clique em qualquer célula da linha."""

        if self.df_original is None:
            return

        colaborador = self.tableWidget.item(row, 2).text().strip()  # Coluna "Colaborador"

        # Localiza a linha do colaborador na planilha original
        dados_colaborador = []
        encontrou_nome = False

        for _, linha in self.df_original.iterrows():
            if not encontrou_nome:
                if pd.notna(linha[2]) and str(linha[2]).strip() == colaborador:
                    encontrou_nome = True  # Começa a coletar os dados
            else:
                if pd.notna(linha[0]) and str(linha[0]).strip().lower() == "total":
                    break  # Parar quando encontrar "Total" na coluna 0

                # Coleta e formata os dados
                escala = linha[4] if pd.notna(linha[4]) else "-"
                situacao = linha[5] if pd.notna(linha[5]) else "-"
                descricao = linha[6] if pd.notna(linha[6]) else "-"

                # Formatar Data para "25/02"
                if pd.notna(linha[7]):
                    try:
                        data = pd.to_datetime(linha[7]).strftime("%d/%m")  # Formato correto
                    except Exception:
                        data = str(linha[7])  # Se der erro, mantém original
                else:
                    data = "-"

                marcacao = linha[8] if pd.notna(linha[8]) else "-"

                # Formatar Horas Geradas para "08:50:00"
                if pd.notna(linha[9]):
                    try:
                        horas_geradas = str(pd.to_datetime(linha[9]).strftime("%H:%M:%S"))  # Removendo milissegundos
                    except Exception:
                        horas_geradas = str(linha[9])  # Se der erro, mantém original
                else:
                    horas_geradas = "-"

                dados_colaborador.append([escala, situacao, descricao, data, marcacao, horas_geradas])

        if not dados_colaborador:
            QMessageBox.warning(self, "Aviso", f"Não foram encontrados dados na planilha para {colaborador}.")
            return

        self.mostrar_detalhes_colaborador(colaborador, dados_colaborador)

    def mostrar_detalhes_colaborador(self, colaborador, dados):
        """Exibe uma nova janela com os dados do colaborador extraídos da planilha."""

        janela = QDialog(self)
        janela.setWindowTitle(f"Detalhes de {colaborador}")
        janela.resize(700, 500)

        layout = QVBoxLayout()

        tabela = QTableWidget()
        tabela.setRowCount(len(dados))
        tabela.setColumnCount(6)
        tabela.setHorizontalHeaderLabels(
            ["Escala", "Situação", "Desc. Situação", "Data", "Marcação às", "Horas Geradas"]
        )

        for row, linha in enumerate(dados):
            for col, valor in enumerate(linha):
                item = QTableWidgetItem(str(valor))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                tabela.setItem(row, col, item)

        tabela.resizeColumnsToContents()
        tabela.horizontalHeader().setStretchLastSection(True)

        layout.addWidget(tabela)
        janela.setLayout(layout)

        janela.exec()

    def abrir_graficos(self):
        """Abre a aba de gráficos usando colunas de horas e dias."""

        colunas = [self.tableWidget.horizontalHeaderItem(col).text() for col in range(self.tableWidget.columnCount())]
        dados = []

        for row in range(self.tableWidget.rowCount()):
            if self.tableWidget.item(row, 0) and self.tableWidget.item(row, 0).text() == "Total":
                continue
            if self.tableWidget.isRowHidden(row):
                continue

            linha = [self.tableWidget.item(row, col).text() if self.tableWidget.item(row, col) else '-'
                     for col in range(self.tableWidget.columnCount())]
            dados.append(linha)

        if not dados:
            QMessageBox.warning(self, "Aviso", "Nenhum dado visível. Ajuste os filtros ou calcule primeiro.")
            return

        df = pd.DataFrame(dados, columns=colunas)

        # Convertendo colunas de horas para timedelta
        colunas_horas = ["Trabalhando", "Atestados"]
        for col in colunas_horas:
            if col in df.columns:
                df[col] = df[col].replace("-", "00:00:00").apply(pd.to_timedelta, errors='coerce').fillna(
                    pd.Timedelta(seconds=0))

        # Convertendo colunas de dias para inteiro
        colunas_dias = ["Dias Trabalhando", "Dias Atestado"]
        for col in colunas_dias:
            if col in df.columns:
                df[col] = df[col].replace("-", "0").replace("", "0").apply(pd.to_numeric, errors='coerce').fillna(
                    0).astype(int)

        self.aba_grafico = AbaGraficoAssinuidade(df, self.periodo_label.text())
        self.aba_grafico.show()

        if hasattr(self, "current_mail") and self.current_mail:
            self.aba_grafico.set_current_email(self.current_mail)


