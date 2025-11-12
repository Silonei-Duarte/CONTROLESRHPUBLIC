import re
import sys
import os
import json
import pandas as pd
from workalendar.america import Brazil
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget,
    QPushButton, QLabel, QFileDialog, QLineEdit, QHBoxLayout, QTableView, QMessageBox, QComboBox
)
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from PyQt6.QtGui import QStandardItemModel, QStandardItem, QIcon
from PyQt6.QtCore import Qt
import win32com.client
from datetime import time, datetime
from PyQt6.QtWidgets import QTabWidget  # Certifique-se de importar o QTabWidget
from horaextragrafico import AbaGraficos  # Importar a classe para a aba de gráficos
import calendar
from main import ControlesRH
from openpyxl.styles import PatternFill

class ControleHorasExtras(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Controle de Horas Extras")
        self.resize(1900, 900)


        # Configurar o ícone da janela
        icon_path = os.path.join(os.path.dirname(__file__), 'icone.ico')  # Ajuste o caminho conforme necessário
        if os.path.exists(icon_path):  # Verifica se o arquivo existe
            self.setWindowIcon(QIcon(icon_path))

        # Criar arquivo de feriados se não existir
        feriados_path = os.path.join(os.path.dirname(__file__), 'feriados.json')
        if not os.path.exists(feriados_path):
            with open(feriados_path, "w", encoding="utf-8") as f:
                json.dump([], f, indent=4, ensure_ascii=False)

        # Criar um QTabWidget para gerenciar abas
        self.tabs = QTabWidget()
        self.resultados_tab = QWidget()
        self.graficos_tab = None  # Será inicializada ao clicar no botão "Calcular"

        # Layout da aba de resultados
        layout = QVBoxLayout()

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

        # Botões para calcular, exportar e gerar e-mail
        button_layout = QHBoxLayout()

        # Combo de ano
        self.combo_ano_dsr = QComboBox()
        anos = [str(a) for a in range(2020, datetime.today().year + 2)]
        self.combo_ano_dsr.addItems(anos)
        self.combo_ano_dsr.setCurrentText(str(datetime.today().year))

        # Combo de mês
        self.combo_mes_dsr = QComboBox()
        meses = [f"{i:02}" for i in range(1, 13)]
        self.combo_mes_dsr.addItems(meses)
        self.combo_mes_dsr.setCurrentText(f"{datetime.today().month:02}")

        # Adicionar ao layout
        button_layout.addWidget(QLabel("Mês DSR:"))
        button_layout.addWidget(self.combo_mes_dsr)
        button_layout.addWidget(QLabel("Ano DSR:"))
        button_layout.addWidget(self.combo_ano_dsr)

        # Botão para calcular HRAP601
        self.calculate_hrap_btn = QPushButton("Calcular HRAP601")
        self.calculate_hrap_btn.setFixedSize(120, 30)
        self.calculate_hrap_btn.clicked.connect(lambda: self.on_calculate_clicked("HRAP601.xlsx"))
        button_layout.addWidget(self.calculate_hrap_btn)

        # Botão para calcular HE SEMANAL
        self.calculate_semanal_btn = QPushButton("Cálculo Semanal")
        self.calculate_semanal_btn.setFixedSize(120, 30)
        self.calculate_semanal_btn.clicked.connect(lambda: self.on_calculate_clicked("HE SEMANAL.xlsx"))
        button_layout.addWidget(self.calculate_semanal_btn)

        # Botão para calcular HE MENSAL
        self.calculate_mensal_btn = QPushButton("Cálculo Mensal")
        self.calculate_mensal_btn.setFixedSize(120, 30)
        self.calculate_mensal_btn.clicked.connect(lambda: self.on_calculate_clicked("HE MENSAL.xlsx"))
        button_layout.addWidget(self.calculate_mensal_btn)

        self.export_excel_btn = QPushButton("Exportar para Excel")
        self.export_excel_btn.setFixedSize(120, 30)
        self.export_excel_btn.clicked.connect(self.export_to_excel)
        button_layout.addWidget(self.export_excel_btn)

        self.generate_email_btn = QPushButton("E-mail")
        self.generate_email_btn.setFixedSize(120, 30)
        self.generate_email_btn.clicked.connect(self.generate_email)
        button_layout.addWidget(self.generate_email_btn)



        button_layout.addStretch()

        # Botão para editar feriados
        self.edit_feriados_btn = QPushButton("Editar Feriados")
        self.edit_feriados_btn.setFixedSize(120, 30)
        self.edit_feriados_btn.clicked.connect(self.editar_feriados)
        button_layout.addWidget(self.edit_feriados_btn)

        #  Criar botão "Voltar ao Menu"
        self.btn_voltar = QPushButton("Voltar ao Menu", self)
        self.btn_voltar.setFixedSize(120, 30)
        self.btn_voltar.clicked.connect(self.voltar_menu)
        button_layout.addWidget(self.btn_voltar)


        #  Adicionar um espaçamento para empurrar os botões à esquerda
        layout.addLayout(button_layout)

        # Label para mostrar qual cálculo foi realizado e o horário da última atualização
        self.status_label = QLabel("Nenhum cálculo realizado ainda")
        self.status_label.setStyleSheet(
            "color: red; font-size: 12px;")  # Estiliza o texto em vermelho e negrito
        layout.addWidget(self.status_label)

        # Campo de filtro único
        self.global_filter = QLineEdit()
        self.global_filter.setPlaceholderText("Filtrar usando / entre dados! Tambem pode digitar COLUNA:DADO para filtrar pela expressão exata na coluna")
        self.global_filter.textChanged.connect(self.apply_global_filter)
        layout.addWidget(self.global_filter)

        # Tabela para exibir os resultados
        self.table_view = QTableView()
        self.table_view.verticalHeader().setVisible(False)  #  Oculta a numeração das linhas
        self.table_view.setStyleSheet("QTableView { gridline-color: black; }")
        layout.addWidget(self.table_view)

        # Inicializar o modelo para QTableView
        self.model = QStandardItemModel()
        self.table_view.setModel(self.model)
        self.table_view.setSortingEnabled(False)

        # Adicionar o layout de resultados à aba "Resultados"
        self.resultados_tab.setLayout(layout)

        # Adicionar a aba de resultados ao QTabWidget
        self.tabs.addTab(self.resultados_tab, "Resultados")

        # Definir o widget principal como o QTabWidget
        self.setCentralWidget(self.tabs)

        # Caminhos das planilhas e configuração
        self.file1_path = None
        self.file2_path = None
        self.config_file = os.path.join(os.path.dirname(__file__), "horaextra.json")
        self.load_config()

    def editar_feriados(self):
        from PyQt6.QtWidgets import QDialog, QVBoxLayout, QPlainTextEdit, QPushButton

        feriados_path = os.path.join(os.path.dirname(__file__), 'feriados.json')

        dialog = QDialog(self)
        dialog.setWindowTitle("Editar Feriados (um por linha - formato dd/mm/aaaa)")
        dialog.resize(400, 300)

        layout = QVBoxLayout(dialog)

        text_edit = QPlainTextEdit(dialog)
        layout.addWidget(text_edit)

        try:
            with open(feriados_path, "r", encoding="utf-8") as f:
                feriados = json.load(f)
                linhas = "\n".join(feriados)
                text_edit.setPlainText(linhas)
        except Exception:
            text_edit.setPlainText("")

        salvar_btn = QPushButton("Salvar", dialog)
        salvar_btn.clicked.connect(lambda: self.salvar_feriados(text_edit.toPlainText(), dialog))
        layout.addWidget(salvar_btn)

        dialog.exec()

    def salvar_feriados(self, texto, dialog):
        feriados_path = os.path.join(os.path.dirname(__file__), 'feriados.json')
        linhas = [linha.strip() for linha in texto.strip().splitlines() if linha.strip()]

        # Validação de datas
        datas_validas = []
        for linha in linhas:
            try:
                datetime.strptime(linha, "%d/%m/%Y")  # Valida formato
                datas_validas.append(linha)
            except ValueError:
                QMessageBox.critical(self, "Erro", f"Data inválida: {linha}. Use formato dd/mm/aaaa.")
                return

        try:
            with open(feriados_path, "w", encoding="utf-8") as f:
                json.dump(datas_validas, f, indent=4, ensure_ascii=False)
            QMessageBox.information(self, "Sucesso", "Feriados salvos com sucesso.")
            dialog.accept()
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao salvar o arquivo: {e}")

    def voltar_menu(self):
        self.menu_principal = ControlesRH()
        self.menu_principal.show()
        self.close()

    def on_calculate_clicked(self, file1_name):
        """Processa os arquivos e cria a aba de gráficos, dependendo do botão pressionado."""
        if not self.folder_path:
            QMessageBox.warning(self, "Erro", "Por favor, selecione a pasta das planilhas antes de calcular.")
            return

        # Define os caminhos corretos dependendo do botão pressionado
        self.file1_path = os.path.join(self.folder_path, file1_name)
        self.file2_path = os.path.join(self.folder_path, "FPRE905.xlsx")

        # Verifica se os arquivos existem
        if not os.path.exists(self.file1_path) or not os.path.exists(self.file2_path):
            QMessageBox.warning(self, "Erro",
                                f"Os arquivos {file1_name} ou FPRE905.xlsx não foram encontrados na pasta selecionada.")
            return

        # Definir o texto correto no QLabel com o horário atualizado
        hora_atual = datetime.now().strftime("%H:%M")
        if file1_name == "HRAP601.xlsx":
            texto = f"Calculo HRAP601"
        elif file1_name == "HE SEMANAL.xlsx":
            texto = f"Calculo Semanal - Atualizado a cada 1 hora"
        elif file1_name == "HE MENSAL.xlsx":
            texto = f"Calculo Mensal - Atualizado a cada 1 hora"

        self.status_label.setText(texto)
        self.status_label.setStyleSheet("color: red;font-size: 12px;")

        # Processar os arquivos
        self.process_files()
        self.create_graficos_tab()

    def create_graficos_tab(self):
        """Recria a aba de gráficos para garantir que os dados sejam sempre atualizados."""
        if self.graficos_tab:
            # Remove a aba antiga antes de recriar
            index = self.tabs.indexOf(self.graficos_tab)
            if index != -1:
                self.tabs.removeTab(index)

        # Criar uma nova instância da AbaGraficos com os dados atualizados
        self.graficos_tab = AbaGraficos(self.resultados)
        self.tabs.addTab(self.graficos_tab, "Gráficos")

    def select_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Selecionar Pasta com Planilhas")
        if folder_path:
            self.folder_path = folder_path
            self.hrap_line_edit.setText(folder_path)  # Atualiza o campo de texto com a pasta selecionada
            self.save_config()

    def save_config(self):
        """Salva o caminho da pasta das planilhas em um arquivo JSON."""
        config = {"folder_path": self.folder_path}
        with open(self.config_file, "w") as f:
            json.dump(config, f)

    def load_config(self):
        """Carrega o caminho da pasta das planilhas de um arquivo JSON."""
        if os.path.exists(self.config_file):
            with open(self.config_file, "r") as f:
                config = json.load(f)
                self.folder_path = config.get("folder_path")

                if self.folder_path:
                    self.hrap_line_edit.setText(self.folder_path)  # Atualiza o campo da pasta

    def process_files(self):
        # Garante que self.file1_path já foi definido pelo botão correto no on_calculate_clicked()
        if not self.file1_path:
            QMessageBox.warning(self, "Erro", "Arquivo principal não definido corretamente.")
            return

        self.file2_path = os.path.join(self.folder_path, "FPRE905.xlsx")
        locais_path = os.path.join(self.folder_path, "LOCAIS.xlsx")

        # Verifica se os arquivos existem
        if not os.path.exists(self.file1_path) or not os.path.exists(self.file2_path):
            QMessageBox.warning(
                self,
                "Erro",
                "Os arquivos HRAP601.xlsx ou FPRE905.xlsx não foram encontrados na pasta selecionada."
            )
            return

        try:
            # Abre e salva os arquivos no Excel sem modificar o conteúdo
            self.open_save_with_excel(self.file1_path)
            self.open_save_with_excel(self.file2_path)

            # Carregar as planilhas no pandas
            planilha1 = pd.read_excel(self.file1_path, header=None)
            planilha2 = pd.read_excel(self.file2_path, header=None)

            def ajustar_horas_excel(valor):
                """ Ajusta corretamente as horas acumuladas considerando dias acumulados como múltiplos de 24h. """
                if isinstance(valor, datetime):
                    dia_real = valor.day
                    horas = valor.hour
                    minutos = valor.minute
                    segundos = valor.second
                    horas_totais = (dia_real * 24) + horas

                    if segundos >= 30:
                        minutos += 1
                    segundos = 0

                    if minutos >= 60:
                        horas_totais += 1
                        minutos = 0

                    return f"{horas_totais:02}:{minutos:02}:00"

                elif isinstance(valor, time):
                    horas, minutos, segundos = valor.hour, valor.minute, valor.second
                    if segundos >= 30:
                        minutos += 1
                    segundos = 0
                    if minutos >= 60:
                        horas += 1
                        minutos = 0
                    return f"{horas:02}:{minutos:02}:00"

                return valor

            # LOCAIS.xlsx (dicionário de locais)
            if os.path.exists(locais_path):
                locais_df = pd.read_excel(locais_path, header=None)
                locais_dict = dict(zip(locais_df[0], locais_df[1]))
            else:
                locais_dict = {}

            resultados = []

            for _, row1 in planilha1.iterrows():
                nome = row1[1]  # Coluna B
                if pd.isna(nome) or str(nome).strip() == "":
                    continue

                numcad = row1[0]
                numemp = row1[6]
                codigo = int(row1[3])
                rotina = row1[4]
                horas = ajustar_horas_excel(row1[5])

                hh, mm, ss = map(int, horas.split(':')) if isinstance(horas, str) else (0, 0, 0)
                horas_em_decimais = Decimal(str(hh)) + (Decimal(str(mm)) / Decimal("60")) + (
                        Decimal(str(ss)) / Decimal("3600"))
                horas_formatadas = f"{hh:02}:{mm:02}:{ss:02}"

                rows = planilha2[(planilha2[0] == numcad) & (planilha2[1] == numemp)]

                if not rows.empty:
                    salario_base_str = str(rows.iloc[0, 5]).strip()
                    salario_base_str = salario_base_str.replace('.', '').replace(',', '.')
                    try:
                        salario_base = Decimal(salario_base_str)
                    except InvalidOperation:
                        salario_base = Decimal("0.0")

                    salario_por_hora = (salario_base / Decimal("220")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

                    index_linha = rows.index[0]

                    setor = ""
                    for i in range(index_linha, -1, -1):
                        valor_celula = str(planilha2.iloc[i, 0]).strip()
                        if valor_celula.startswith("01."):
                            setor = valor_celula
                            break

                    # separa partes
                    # setor vem de algo como "01.01.06.02 , 2012, Administração Geral, Suprimentos, Almoxarifado"
                    codigo_setor = setor.split(",")[0].strip()  # -> "01.01.06.02"

                    # procurar local no LOCAIS.xlsx pela sequência de código
                    local = "Local não encontrado"
                    if locais_dict:
                        for k, v in locais_dict.items():
                            k_str = str(k).strip()
                            if k_str.startswith(codigo_setor):
                                local = v
                                break
                        # se não achou, tenta reduzir o código (tirando o último ponto)
                        if local == "Local não encontrado" and "." in codigo_setor:
                            base = ".".join(codigo_setor.split(".")[:-1])
                            for k, v in locais_dict.items():
                                if str(k).strip().startswith(base):
                                    local = v
                                    break

                    # partes depois da vírgula continuam servindo pra exibição
                    partes = [p.strip() for p in setor.split(",")]
                    setor_exibicao = partes[-1] if len(partes) >= 1 else setor

                    # debug
                    print(f"\nFuncionário: {nome}")
                    print(f"Setor bruto: {setor}")
                    print(f"Código setor: {codigo_setor}")
                    print(f"Setor exibicao: {setor_exibicao}")
                    print(f"Local encontrado: {local}")
                    print("-" * 80)



                else:
                    salario_base = Decimal("0.0")
                    salario_por_hora = Decimal("0.0")
                    setor = "Setor não encontrado"
                    setor_exibicao = setor
                    local = setor

                valor_final = salario_por_hora * horas_em_decimais
                if codigo in [17, 80]:
                    valor_final *= Decimal("1.5")
                elif codigo == 303:
                    valor_final *= Decimal("2.0")
                elif codigo == 302:
                    valor_final = (horas_em_decimais * salario_por_hora * Decimal("1.5")) + (
                            (horas_em_decimais * salario_por_hora * Decimal("1.5")) * Decimal("0.25")
                    )
                elif codigo == 304:
                    valor_final = (horas_em_decimais * salario_por_hora * Decimal("2.0")) + (
                            (horas_em_decimais * salario_por_hora * Decimal("2.0")) * Decimal("0.25")
                    )

                valor_final = valor_final.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

                resultados.append([
                    local, setor_exibicao, nome, codigo, rotina, horas_formatadas,
                    float(salario_base), float(salario_por_hora), float(valor_final)
                ])

            resultados = self.calcular_dsr(resultados)
            self.display_results(resultados)
            self.resultados = resultados

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao processar as planilhas: {e}")

    def calcular_dsr(self, resultados):
        """Calcula e insere a linha do DSR abaixo de cada funcionário."""
        dsr_resultados = []
        temp_grupo = []
        nome_atual = None

        # Pegando mês e ano selecionados manualmente
        ano_atual = int(self.combo_ano_dsr.currentText())
        mes_atual = int(self.combo_mes_dsr.currentText())

        # 📥 Carregar feriados do JSON
        feriados_path = os.path.join(os.path.dirname(__file__), "feriados.json")
        feriados = []

        if os.path.exists(feriados_path):
            with open(feriados_path, "r", encoding="utf-8") as f:
                feriados = [
                    datetime.strptime(data_str.strip(), "%d/%m/%Y").date()
                    for data_str in json.load(f)
                ]

        # 📅 Dias úteis (segunda a sábado)
        dias_uteis_lista = []
        for dia in range(1, calendar.monthrange(ano_atual, mes_atual)[1] + 1):
            dt = datetime(ano_atual, mes_atual, dia)
            if dt.weekday() < 6:
                dias_uteis_lista.append(dt.date())

        # 📅 Domingos
        domingos_lista = [
            dia for dia in range(1, calendar.monthrange(ano_atual, mes_atual)[1] + 1)
            if datetime(ano_atual, mes_atual, dia).weekday() == 6
        ]

        # 🎯 Feriados no mês, apenas segunda a sexta (0 a 4)
        feriados_mes = [
            f for f in feriados
            if f.year == ano_atual and f.month == mes_atual and f.weekday() < 5
        ]

        # ✅ Remover apenas feriados úteis (segunda a sexta)
        dias_uteis = sum(1 for d in dias_uteis_lista if d not in feriados_mes)

        # Total DSR = domingos + feriados úteis
        total_dsr_dias = len(domingos_lista) + len(feriados_mes)

        # Agrupar por funcionário
        for linha in resultados:
            nome = linha[2]  # Coluna "Nome"

            if nome_atual and nome != nome_atual:
                dsr_linha = self.calcular_linha_dsr(temp_grupo, nome_atual, dias_uteis, total_dsr_dias)
                if dsr_linha:
                    dsr_resultados.append(dsr_linha)
                temp_grupo = []

            temp_grupo.append(linha)
            dsr_resultados.append(linha)
            nome_atual = nome

        if temp_grupo:
            dsr_linha = self.calcular_linha_dsr(temp_grupo, nome_atual, dias_uteis, total_dsr_dias)
            if dsr_linha:
                dsr_resultados.append(dsr_linha)

        return dsr_resultados

    def calcular_linha_dsr(self, grupo, nome, dias_uteis, total_dsr_dias):
        """Calcula o valor do DSR para um grupo de registros de um funcionário."""
        total_50 = Decimal("0.00")
        total_100 = Decimal("0.00")
        total_adicional_noturno_50 = Decimal("0.00")
        total_adicional_noturno_100 = Decimal("0.00")

        local, rotina, setor = "", "", ""

        for linha in grupo:
            codigo = linha[3]  # Código
            valor_final = Decimal(str(linha[8]))  # Valor final

            if codigo in [17, 80]:  # Hora extra 50%
                total_50 += valor_final
            elif codigo == 303:  # Hora extra 100%
                total_100 += valor_final
            elif codigo == 302:  # Adicional noturno 50%
                total_adicional_noturno_50 += valor_final
            elif codigo == 304:  # Adicional noturno 100%
                total_adicional_noturno_100 += valor_final

            local = linha[0]
            rotina = linha[3]
            setor = linha[1]

        # Soma todos os adicionais
        total_horas_extras = total_50 + total_100 + total_adicional_noturno_50 + total_adicional_noturno_100

        if total_horas_extras > 0 and dias_uteis > 0 and total_dsr_dias > 0:
            # Fórmula correta do DSR ajustada para o mês atual
            dsr_valor = (total_horas_extras / dias_uteis) * total_dsr_dias
            dsr_valor = dsr_valor.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            # Cria a linha do DSR e retorna
            return [local, setor, nome, "DSR", "DSR", "-", "-","-", float(dsr_valor)]

        return None

    def open_save_with_excel(self, file_path):
        """Abre e salva a planilha usando o Excel via win32com sem fechar outras instâncias."""
        excel = win32com.client.Dispatch("Excel.Application")

        # Verifica se o Excel já estava rodando antes
        was_excel_open = excel.Workbooks.Count > 0

        excel.Visible = False  # Não exibir a interface do Excel

        try:
            workbook = excel.Workbooks.Open(file_path)
            workbook.Save()  # Salva a planilha
            workbook.Close(SaveChanges=False)
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao abrir/salvar a planilha no Excel: {e}")
        finally:
            # Fecha o Excel apenas se ele não estava aberto antes do script rodar
            if not was_excel_open:
                excel.Quit()

    def display_results(self, resultados):
        """Exibe os resultados alternando cores por grupo de nomes, com linhas DSR destacadas."""
        self.model.clear()
        self.model.setHorizontalHeaderLabels([
            "Local", "Setor", "Nome", "Código", "Rotina", "Horas",
            "Salário Base", "Salário por Hora", "Valor Final"
        ])

        current_name = None
        alternate_color = False
        group_color_1 = Qt.GlobalColor.white
        group_color_2 = Qt.GlobalColor.lightGray

        for resultado in resultados:
            items = []
            nome = resultado[2]
            is_dsr = resultado[3] == "DSR"

            if nome != current_name and not is_dsr:
                alternate_color = not alternate_color
                current_name = nome

            background_color = group_color_1 if alternate_color else group_color_2

            for col_index, value in enumerate(resultado):
                item = QStandardItem(str(value))

                # Formatação correta dos valores numéricos
                if isinstance(value, float):
                    item.setText(f"{value:.2f}")

                # Aplicar a cor de fundo mantendo o agrupamento correto
                item.setBackground(background_color)

                item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
                items.append(item)

            self.model.appendRow(items)

        self.update_totals()
        self.table_view.resizeColumnsToContents()

    def apply_global_filter(self, text):
        """Filtra as linhas da tabela com base na pesquisa parcial e permite busca exata por coluna (ex: Setor:"Produção")."""

        terms = [term.strip() for term in text.split('/') if term.strip()]  # Normaliza os termos

        if not terms:  # Se não houver termos, exibe todas as linhas
            for row in range(self.model.rowCount()):
                self.table_view.setRowHidden(row, False)
            self.update_totals()
            return

        #  Criar um dicionário com os nomes das colunas para facilitar a busca específica
        headers = {}
        for col in range(self.model.columnCount()):
            header_item = self.model.headerData(col, Qt.Orientation.Horizontal, Qt.ItemDataRole.DisplayRole)
            if header_item:  # Apenas adiciona colunas com nome
                headers[header_item.strip().lower()] = col

        #  Iterar por todas as linhas da tabela
        for row in range(self.model.rowCount()):
            row_matches_all_terms = True  # Assume que a linha deve aparecer

            for term in terms:
                term_match = False  # Assume que o termo **não** foi encontrado na linha

                #  Verifica se o termo segue o formato "coluna:valor" para busca exata
                if ":" in term:
                    coluna_nome, valor_busca = map(str.strip, term.split(":", 1))
                    coluna_nome = coluna_nome.lower()

                    #  Se a coluna existir, busca SOMENTE nessa coluna
                    if coluna_nome in headers:
                        col_idx = headers[coluna_nome]
                        index = self.model.index(row, col_idx)
                        cell_text = self.model.data(index, Qt.ItemDataRole.DisplayRole) or ""

                        if valor_busca.lower() in cell_text.strip().lower():
                            term_match = True  # Encontrou o termo na coluna correta

                else:
                    #  Busca geral (qualquer coluna)
                    for col in range(self.model.columnCount()):
                        index = self.model.index(row, col)
                        cell_text = self.model.data(index, Qt.ItemDataRole.DisplayRole) or ""

                        if term.lower() in cell_text.strip().lower():
                            term_match = True
                            break  # Já encontrou, não precisa verificar o resto

                if not term_match:
                    row_matches_all_terms = False
                    break  # Se um termo não foi encontrado, a linha não será exibida

            self.table_view.setRowHidden(row,
                                         not row_matches_all_terms)  # Esconde apenas se **não** corresponde a todos os termos

        self.update_totals()  # Atualiza os totais após o filtro

    def export_to_excel(self):
        """Exporta os dados filtrados atualmente exibidos no QTableView para um arquivo Excel com alternância de cores."""
        file_path, _ = QFileDialog.getSaveFileName(self, "Salvar Arquivo Excel", "", "Arquivos Excel (*.xlsx)")
        if file_path:
            try:
                # Preparar os dados filtrados para exportação
                data = []
                row_colors = []  # Para guardar as cores das linhas
                last_name = None  # Variável para guardar o nome da última pessoa
                alternate_color = True  # Para alternar entre branco e cinza claro

                for row in range(self.model.rowCount()):
                    if not self.table_view.isRowHidden(row):  # Verifica se a linha está visível (não filtrada)
                        row_data = []
                        nome_index = self.model.index(row, 1)  # Coluna "Nome"
                        nome = self.model.data(nome_index, Qt.ItemDataRole.DisplayRole)  # Obter o nome

                        for col in range(self.model.columnCount()):
                            index = self.model.index(row, col)
                            value = self.model.data(index, Qt.ItemDataRole.DisplayRole)

                            # Verifica se a coluna é uma das que precisam de formatação de número
                            if col in [5, 6, 7]:  # Colunas Salário Base, Salário por Hora, Valor Final
                                # Troca ponto por vírgula
                                value = str(value).replace('.', ',')

                            row_data.append(value)

                        data.append(row_data)

                        # Determina a cor da linha com base no nome
                        if nome != last_name:  # Se o nome for diferente do anterior
                            alternate_color = not alternate_color  # Alterna a cor
                        row_colors.append('D3D3D3' if alternate_color else 'FFFFFF')  # lightgray and white

                        last_name = nome  # Atualiza o último nome

                # Criar o DataFrame com os dados
                df = pd.DataFrame(data, columns=[
                    self.model.headerData(i, Qt.Orientation.Horizontal) for i in range(self.model.columnCount())
                ])

                # Criar o arquivo Excel usando pandas com openpyxl
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Controle de Horas Extras')

                    # Acessar o arquivo Excel criado e aplicar as cores
                    workbook = writer.book
                    worksheet = workbook['Controle de Horas Extras']

                    # Aplicar as cores alternadas nas linhas
                    for row_num, row_color in enumerate(row_colors, start=2):  # Começar na linha 2
                        fill = PatternFill(start_color=row_color, end_color=row_color, fill_type="solid")
                        for cell in worksheet[row_num]:
                            cell.fill = fill

                QMessageBox.information(self, "Sucesso", f"Dados exportados para: {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao exportar para Excel: {e}")

    def generate_email(self):
        """Gera um e-mail com os dados filtrados atualmente na tabela, agrupando por Local e Setor no estilo do frequencia.py."""
        global total_horas_formatado
        try:
            #  Definir cores dos locais (somente no fundo da célula Local)
            cores_locais = {
                "Fabrica De Máquinas": "#003756",
                "Fabrica De Transportadores": "#ffc62e",
                "Adm": "#009c44",
                "Comercial": "#919191"
            }

            #  Criar dicionário agrupando por LOCAL e SETOR
            agrupados = {}
            total_valor_final = 0.0

            for row in range(self.model.rowCount()):
                if not self.table_view.isRowHidden(row):  # Apenas linhas visíveis
                    local = self.model.data(self.model.index(row, 0), Qt.ItemDataRole.DisplayRole) or "Sem local"
                    setor = self.model.data(self.model.index(row, 1), Qt.ItemDataRole.DisplayRole) or "Sem setor"
                    nome = self.model.data(self.model.index(row, 2), Qt.ItemDataRole.DisplayRole) or "Sem Nome"
                    salario_base = self.model.data(self.model.index(row, 6), Qt.ItemDataRole.DisplayRole) or "0.00"
                    valor_final = self.model.data(self.model.index(row, 8), Qt.ItemDataRole.DisplayRole) or "0.00"
                    horas = self.model.data(self.model.index(row, 5), Qt.ItemDataRole.DisplayRole) or "00:00:00"
                    codigo = self.model.data(self.model.index(row, 3), Qt.ItemDataRole.DisplayRole) or ""

                    if local == "TOTAL":
                        total_horas_formatado = horas
                        total_valor_final = float(valor_final)
                        continue

                    try:
                        valor_final = float(valor_final)
                    except ValueError:
                        valor_final = 0.0

                    try:
                        salario_base = float(salario_base)
                    except ValueError:
                        salario_base = 0.0

                    # Converter horas para segundos apenas se não for a linha do DSR
                    horas_em_segundos = 0
                    if codigo != "DSR":
                        try:
                            partes_horas = list(map(int, horas.split(":")))
                            while len(partes_horas) < 3:
                                partes_horas.append(0)  # Garante que temos HH:MM:SS
                            hh, mm, ss = partes_horas
                            horas_em_segundos = (hh * 3600) + (mm * 60) + ss
                        except ValueError:
                            horas_em_segundos = 0

                    #  Agrupar primeiro por LOCAL
                    if local not in agrupados:
                        agrupados[local] = {}

                    #  Depois, agrupar por SETOR dentro do LOCAL
                    if setor not in agrupados[local]:
                        agrupados[local][setor] = {}

                    #  Depois, agrupar por FUNCIONÁRIO dentro do SETOR (somando horas e valores)
                    if nome not in agrupados[local][setor]:
                        agrupados[local][setor][nome] = {
                            "salario_base": salario_base,
                            "horas_em_segundos": 0,
                            "valor_final": 0.0
                        }

                    # Acumular valores
                    agrupados[local][setor][nome]["horas_em_segundos"] += horas_em_segundos
                    agrupados[local][setor][nome]["valor_final"] += valor_final

            html_body = """
            <html>
            <head>
                <style>
                    table {
                        border-collapse: collapse;
                        width: auto;
                        font-size: 14px;
                    }
                    th, td {
                        border: 1px solid black;
                        padding: 6px;
                        white-space: nowrap;
                    }
                    th {
                        background-color: #f2f2f2;
                    }
                </style>
            </head>
            <body>
                <p>Controle de horas extras detalhado.</p>

                <table>
                    <tr>
                        <th>Local</th>
                        <th>Setor</th>
                        <th>Nome</th>
                        <th>Horas</th>
                        <th>Salário Base</th>
                        <th>Valor Final HREs</th>
                    </tr>
            """

            #  Adicionar os dados agrupados corretamente na **única tabela**
            for local, setores in agrupados.items():
                cor_local = cores_locais.get(local, "#000000")  # Cor apenas no fundo da célula do Local

                for setor, funcionarios in setores.items():
                    qtd_funcionarios_setor = len(funcionarios)
                    primeira_linha_setor = True  # Para mesclar corretamente o Setor

                    for nome, dados in funcionarios.items():
                        # Converter total de horas de segundos para HH:MM:SS
                        total_horas = dados["horas_em_segundos"]
                        hh = total_horas // 3600
                        mm = (total_horas % 3600) // 60
                        ss = total_horas % 60
                        horas_formatadas = f"{hh:02}:{mm:02}:{ss:02}"

                        html_body += "<tr>"

                        #  O LOCAL aparece em todas as linhas (não mescla mais)
                        html_body += f'<td style="background-color:{cor_local}; color:#fff; font-weight:bold;">{local}</td>'

                        #  O SETOR ainda é mesclado corretamente
                        if primeira_linha_setor:
                            html_body += f'<td rowspan="{qtd_funcionarios_setor}">{setor}</td>'
                            primeira_linha_setor = False  # Evita repetir

                        #  Adicionar os dados do funcionário
                        html_body += f"""
                            <td>{nome}</td>
                            <td>{horas_formatadas}</td>
                            <td>R$ {dados["salario_base"]:.2f}</td>
                            <td>R$ {dados["valor_final"]:.2f}</td>
                        </tr>
                        """

            #  Adicionar a linha TOTAL no final, sem afetar a assinatura
            html_body += f"""
                <tr>
                    <td colspan="3" style="font-weight:bold; text-align:center; background-color:#f2f2f2;">TOTAL GERAL</td>
                    <td style="text-align:center;">{total_horas_formatado}</td>
                    <td colspan="2" style="text-align:center;">R$ {total_valor_final:.2f}</td>
                </tr>
            """

            html_body += "</table><br></table>"


            #  Criar o e-mail no Outlook e adicionar a assinatura corretamente
            outlook = win32com.client.Dispatch("Outlook.Application")
            self.current_mail = outlook.CreateItem(0)
            self.current_mail.Subject = "Controle de Horas Extras"
            inspector = self.current_mail.GetInspector
            signature = self.current_mail.HTMLBody  # Obtém a assinatura original do Outlook
            self.current_mail.HTMLBody = html_body + "<br><br><style>th, td {border: none;}</style>" + signature

            self.current_mail.Display()  # Exibir o e-mail no Outlook

            #  Atualizar a aba de gráficos, se existir
            if self.graficos_tab:
                self.graficos_tab.set_current_email(self.current_mail)

            QMessageBox.information(self, "Sucesso", "E-mail gerado com sucesso no Outlook!")

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao gerar o e-mail: {e}")

    def update_totals(self):
        """Atualiza a linha totalizadora com a soma das colunas de Horas e Valor Final, respeitando os filtros."""
        total_horas_segundos = 0
        total_valor_final = Decimal("0.00")

        #  Remover a linha total anterior (caso já exista)
        last_row = self.model.rowCount() - 1
        if last_row >= 0:
            last_item = self.model.index(last_row, 0).data(Qt.ItemDataRole.DisplayRole)
            if last_item == "TOTAL":
                self.model.removeRow(last_row)

        for row in range(self.model.rowCount()):
            if not self.table_view.isRowHidden(row):  # Apenas somar linhas visíveis (não filtradas)
                horas_index = self.model.index(row, 5)  # Coluna "Horas"
                valor_final_index = self.model.index(row, 8)  # Coluna "Valor Final"

                horas_str = self.model.data(horas_index, Qt.ItemDataRole.DisplayRole) or "00:00:00"
                valor_final_str = self.model.data(valor_final_index, Qt.ItemDataRole.DisplayRole) or "0.00"

                # Converter tempo HH:MM:SS para segundos
                try:
                    hh, mm, ss = map(int, horas_str.split(":"))
                    total_horas_segundos += (hh * 3600) + (mm * 60) + ss
                except ValueError:
                    pass  # Ignorar caso não seja um formato válido de tempo

                # Converter valor final para Decimal e somar
                try:
                    valor_final = Decimal(str(valor_final_str))
                    total_valor_final += valor_final
                except ValueError:
                    pass  # Ignorar caso não seja um número válido

        # Converter total de segundos para HH:MM:SS
        hh = total_horas_segundos // 3600
        mm = (total_horas_segundos % 3600) // 60
        ss = total_horas_segundos % 60
        total_horas_formatado = f"{hh:02}:{mm:02}:{ss:02}"

        #  Adicionar novamente a linha total atualizada
        self.add_total_row(total_horas_formatado, total_valor_final)

        #  Garantir que a tabela seja redesenhada corretamente após atualização
        self.table_view.viewport().update()

    def add_total_row(self, total_horas, total_valor_final):
        """Adiciona ou atualiza a linha de totalização no final da tabela."""
        # Criar a linha do total
        total_row = ["TOTAL", "", "", "", "", total_horas, "", "", f"{total_valor_final:.2f}"]

        # Verificar se já existe a linha total e remover antes de adicionar
        last_row = self.model.rowCount() - 1
        if last_row >= 0:
            last_item = self.model.index(last_row, 0).data(Qt.ItemDataRole.DisplayRole)
            if last_item == "TOTAL":
                self.model.removeRow(last_row)

        # Criar os itens da linha totalizadora
        items = []
        for col_index, value in enumerate(total_row):
            item = QStandardItem(str(value))
            item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)  # Impedir edição
            item.setBackground(Qt.GlobalColor.yellow)  # Destacar a linha em amarelo
            item.setFont(self.model.item(0, 0).font())  # Manter a fonte padrão
            items.append(item)

        # Adicionar a linha totalizadora no final
        self.model.appendRow(items)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = ControleHorasExtras()
    window.show()
    sys.exit(app.exec())