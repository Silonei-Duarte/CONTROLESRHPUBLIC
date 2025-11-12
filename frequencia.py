import sys
import oracledb
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QWidget, QLineEdit, QDateEdit, QPushButton, QHBoxLayout, QLabel, QMessageBox, QTabWidget
)
import os
import re
from PyQt6.QtGui import QColor, QIcon, QFont
from PyQt6.QtCore import Qt, QDate
import win32com.client
import pandas as pd
from datetime import datetime, timedelta
from main import ControlesRH
from contextlib import contextmanager
from hisfrequenciagrafico import AbaGraficoFrequencia
from historicofrequencia import AbaHistorico
from Assiduidade import AbaAssiduidade
from Assinuidade_Atestados import AbaAssiduidade_Atestado
from Database import get_connection


SQL_CONSIDERAPONTO = """
WITH UltimosRegistros AS (
    SELECT 
        NUMEMP, 
        NUMCAD, 
        INIAPU, 
        APUPON,
        ROW_NUMBER() OVER (PARTITION BY NUMEMP, NUMCAD ORDER BY INIAPU DESC) AS rn
    FROM R038APU
    WHERE NUMEMP IN (10, 16, 17)
      AND TIPCOL = 1
)
SELECT U.NUMEMP, U.NUMCAD
FROM UltimosRegistros U
JOIN R034FUN F ON F.NUMEMP = U.NUMEMP AND F.NUMCAD = U.NUMCAD AND F.TIPCOL = 1
WHERE U.rn = 1 
  AND U.APUPON = 1
  AND F.DATADM <= SYSDATE
  AND NOT EXISTS (
      SELECT 1
      FROM R038AFA A
      WHERE A.NUMEMP = U.NUMEMP
        AND A.NUMCAD = U.NUMCAD
        AND A.TIPCOL = 1
        AND A.SITAFA <> 1
        AND A.DATTER > SYSDATE
        AND A.DATAFA < SYSDATE
  )

"""

SQL_BATIDAS_QUERY = """
SELECT 
    f.NUMEMP, 
    f.NUMCAD,
    f.NumLoc, 
    f.NomFun, 
    f.CodEsc,
    COALESCE(TO_CHAR(a.DatAcc, 'DD/MM/YYYY'), :selected_date) AS DataFormatada,
    COALESCE(
        LISTAGG(
            TO_CHAR(TO_TIMESTAMP('1970-01-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS') 
            + INTERVAL '1' MINUTE * a.HorAcc, 'HH24:MI'), ' / '
        ) WITHIN GROUP (ORDER BY a.HorAcc),
        ''
    ) AS HorariosFormatados
FROM R034FUN f
LEFT JOIN R070ACC a 
    ON a.NUMEMP = f.NUMEMP 
    AND a.NUMCAD = f.NUMCAD
    AND a.DatAcc = TO_DATE(:selected_date, 'DD/MM/YYYY')
WHERE 
    f.NUMEMP IN (10, 16, 17)
    AND f.SitAfa = '001'
    AND f.TipCol = 1
GROUP BY 
    f.NUMEMP, f.NUMCAD, f.NumLoc, f.NomFun, f.CodEsc, a.DatAcc
ORDER BY f.NUMEMP, f.NUMCAD
"""

SQL_QUERY2 = """
SELECT 
    f.NUMEMP, 
    f.NUMCAD,
    f.NumLoc, 
    f.NomFun, 
    f.CodEsc,
    TO_CHAR(TO_DATE(:selected_date, 'DD/MM/YYYY'), 'DD/MM/YYYY') AS DataFormatada,
    LISTAGG(
        TO_CHAR(
            TO_TIMESTAMP('1970-01-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS') 
            + INTERVAL '1' MINUTE * a.HorAcc, 'HH24:MI'
        ), ' / '
    ) WITHIN GROUP (ORDER BY a.DatAcc ASC, a.HorAcc ASC) AS HorariosFormatados 
FROM R034FUN f
LEFT JOIN R070ACC a 
    ON a.NUMEMP = f.NUMEMP 
    AND a.NUMCAD = f.NUMCAD
    AND (a.DatAcc + (a.HorAcc / 1440)) BETWEEN 
        (TO_DATE(:selected_date, 'DD/MM/YYYY') + INTERVAL '12' HOUR) 
        AND (TO_DATE(:selected_date, 'DD/MM/YYYY') + INTERVAL '1' DAY + INTERVAL '12' HOUR)
WHERE 
    f.CodEsc = 62
    AND f.NUMEMP IN (10, 16, 17, 19, 11)
    AND f.SitAfa = '001'
    AND f.TipCol = 1
GROUP BY 
    f.NUMEMP, f.NUMCAD, f.NumLoc, f.NomFun, f.CodEsc
ORDER BY 
    f.NUMEMP, f.NUMCAD

"""

SQL_Setor ="""SELECT CODLOC, NUMLOC  FROM R016HIE"""

SQL_CONSIDERABANCO ="""SELECT POSSBH
FROM R038HSI
WHERE NUMEMP = 17
  AND TIPCOL = 1
  AND NUMCAD = 77
ORDER BY DATALT DESC
FETCH FIRST 1 ROW ONLY"""

class controlefrequencia(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Controle de Frequência")
        self.resize(1900, 900)

        # Criando o TabWidget para múltiplas abas
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # Criando a aba de Frequência normalmente
        self.aba_frequencia = QWidget()
        self.setup_aba_frequencia()
        self.tabs.addTab(self.aba_frequencia, "Frequência")

        #  Adicionando a aba "Histórico" diretamente com a classe importada
        self.aba_historico = AbaHistorico(self)  # Criando o objeto da classe importada
        self.tabs.addTab(self.aba_historico, "Histórico")

        #  Adicionando a aba "assiduidade" diretamente com a classe importada
        self.aba_assiduidade = AbaAssiduidade(self)  # Criando o objeto da classe importada
        self.tabs.addTab(self.aba_assiduidade, "Assiduidade")

        #  Adicionando a aba "Atestados" diretamente com a classe importada
        self.aba_assiduidade_atestados = AbaAssiduidade_Atestado(self)  # Criando o objeto da classe importada
        self.tabs.addTab(self.aba_assiduidade_atestados, "Atestados")

        # Configurar o ícone da janela
        icon_path = os.path.join(os.path.dirname(__file__), 'icone.ico')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        #  Carregar setores do banco de dados {NUMLOC: CODLOC}
        self.setores_dict = carregar_setores()

        #  Criar o dicionário de locais diretamente na `__init__`
        locais_path = "T:/SC-CNV  BI-BF/RH/Rotinas/RELATORIOS/LOCAIS.xlsx"
        if os.path.exists(locais_path):
            locais_df = pd.read_excel(locais_path, header=None)
            self.dicionario_locais = dict(zip(locais_df[0], locais_df[1]))
        else:
            self.dicionario_locais = {}

        #  Carregar os dados iniciais (com a data de hoje)
        self.update_data_with_date()

    def setup_aba_frequencia(self):
        """Configura a aba principal de frequência"""
        layout = QVBoxLayout()

        #  Layout horizontal para a data e o botão de busca
        date_layout = QHBoxLayout()

        self.label_data = QLabel("Expediente:")
        date_layout.addWidget(self.label_data)

        self.date_picker = QDateEdit()
        self.date_picker.setCalendarPopup(True)
        self.date_picker.setDate(QDate.currentDate())
        self.date_picker.setFixedWidth(120)
        date_layout.addWidget(self.date_picker)

        self.label_dia_semana = QLabel("")
        date_layout.addWidget(self.label_dia_semana)

        self.date_picker.dateChanged.connect(self.atualizar_nome_dia)
        self.atualizar_nome_dia()

        self.search_button = QPushButton("Buscar")
        self.search_button.setFixedSize(120, 30)
        self.search_button.clicked.connect(self.update_data_with_date)
        date_layout.addWidget(self.search_button)

        self.email_button = QPushButton("E-mail Atrasos")
        self.email_button.setFixedSize(120, 30)
        self.email_button.clicked.connect(self.enviar_email_frequencia)
        date_layout.addWidget(self.email_button)

        self.delete_button = QPushButton("Excluir Linha")
        self.delete_button.setFixedSize(120, 30)
        self.delete_button.clicked.connect(self.excluir_linha_selecionada)

        self.graph_button = QPushButton("Gráficos")
        self.graph_button.setFixedSize(120, 30)
        self.graph_button.clicked.connect(self.abrir_graficos)

        self.btn_voltar = QPushButton("Voltar ao Menu")
        self.btn_voltar.setFixedSize(120, 30)
        self.btn_voltar.clicked.connect(self.voltar_menu)

        date_layout.addStretch()
        date_layout.addWidget(self.delete_button)
        date_layout.addWidget(self.graph_button)
        date_layout.addWidget(self.btn_voltar)

        self.search_field = QLineEdit()
        self.search_field.setPlaceholderText(
            "Filtrar usando / entre dados! Tambem pode digitar COLUNA:DADO para filtrar pela expressão exata na coluna")
        self.search_field.textChanged.connect(self.apply_global_filter)

        layout.addLayout(date_layout)
        layout.addWidget(self.search_field)

        #  Legenda visual de cores (sem HTML)
        legenda_layout = QHBoxLayout()

        def adicionar_legenda(cor_hex, texto):
            cor = QLabel()
            cor.setFixedSize(20, 20)
            cor.setStyleSheet(f"background-color: {cor_hex}; border: 1px solid black;")
            label = QLabel(texto)
            label.setStyleSheet("margin-left: 4px; margin-right: 15px;")
            legenda_layout.addWidget(cor)
            legenda_layout.addWidget(label)

        adicionar_legenda("#ff0000", "Falta")
        adicionar_legenda("#ffff00", "Atraso")
        adicionar_legenda("#20b812", "Hora Extra")
        adicionar_legenda("#00a2ff", "Considera Banco de Horas")
        adicionar_legenda("#ffffff", "Dentro do Horário")

        layout.addLayout(legenda_layout)

        self.tableWidget = QTableWidget()

        #  Definir tamanho da fonte da tabela aqui
        font = self.tableWidget.font()
        font.setPointSize(8)
        self.tableWidget.setFont(font)

        layout.addWidget(self.tableWidget)

        self.aba_frequencia.setLayout(layout)

    def abrir_graficos(self):
        """Abre a aba de gráficos passando os dados da tabela atual, incluindo faltas e atrasos de banco de horas."""
        dados = []
        colunas = ['Local', 'Setor', 'Colaborador', 'Qtd. Faltas', 'Qtd. Atrasos', 'Qtd. Faltas BH', 'Qtd. Atrasos BH']

        for row in range(self.tableWidget.rowCount()):
            if not self.tableWidget.isRowHidden(row):  # Apenas linhas visíveis
                local = self.tableWidget.item(row, 0).text() if self.tableWidget.item(row, 0) else '-'
                setor = self.tableWidget.item(row, 1).text() if self.tableWidget.item(row, 1) else '-'
                colaborador = self.tableWidget.item(row, 3).text() if self.tableWidget.item(row, 3) else '-'

                qtd_faltas = 0
                qtd_atrasos = 0
                qtd_faltas_bh = 0
                qtd_atrasos_bh = 0

                contou_falta = False
                contou_falta_bh = False

                for col in range(6, self.tableWidget.columnCount(), 2):
                    batida_item = self.tableWidget.item(row, col + 1)
                    if batida_item:
                        cor = batida_item.background().color().name().lower()
                        texto = batida_item.text().strip().lower()

                        # Falta
                        if cor == "#ff0000" and texto == "falta" and not contou_falta:
                            qtd_faltas += 1
                            contou_falta = True

                        # Atraso
                        elif cor == "#ffff00":
                            qtd_atrasos += 1

                        # Falta BH
                        elif cor == "#00a2ff" and texto == "falta" and not contou_falta_bh:
                            qtd_faltas_bh += 1
                            contou_falta_bh = True

                        # Atraso BH
                        elif cor == "#00a2ff" and texto != "falta":
                            qtd_atrasos_bh += 1

                if any([qtd_faltas, qtd_atrasos, qtd_faltas_bh, qtd_atrasos_bh]):
                    dados.append([local, setor, colaborador, qtd_faltas, qtd_atrasos, qtd_faltas_bh, qtd_atrasos_bh])

        if not dados:
            QMessageBox.warning(self, "Aviso",
                                "Nenhum dado visível com atraso ou falta. Ajuste os filtros ou calcule primeiro.")
            return

        df = pd.DataFrame(dados, columns=colunas)
        periodo = self.date_picker.date().toString("dd/MM/yyyy")
        self.aba_grafico = AbaGraficoFrequencia(df, periodo)
        self.aba_grafico.show()

        if hasattr(self, "current_mail") and self.current_mail:
            self.aba_grafico.set_current_email(self.current_mail)

    def atualizar_nome_dia(self):
        """Atualiza o nome do dia da semana com base na data selecionada."""
        data_selecionada = self.date_picker.date().toPyDate()
        nome_dia = data_selecionada.strftime("%A")  # Obtém o nome do dia da semana
        dias_traduzidos = {
            "Monday": "Segunda-feira",
            "Tuesday": "Terça-feira",
            "Wednesday": "Quarta-feira",
            "Thursday": "Quinta-feira",
            "Friday": "Sexta-feira",
            "Saturday": "Sábado",
            "Sunday": "Domingo"
        }
        self.label_dia_semana.setText(dias_traduzidos.get(nome_dia, nome_dia))  # Atualiza o label


    def update_data_with_date(self):
        """Atualiza os dados com base na data selecionada pelo usuário."""
        selected_date = self.date_picker.date().toString("dd/MM/yyyy")  # ✅ Ajuste no formato da data
        validado, max_batidas, escala_horarios = validar_batidas(selected_date)  # Agora retorna escala_horarios também
        self.reconstruir_tabela(validado, max_batidas, escala_horarios)

    def reconstruir_tabela(self, validado, max_batidas, escala_horarios):
        """ Remove todos os dados e recria a tabela do zero, garantindo que o layout seja gerado corretamente. """

        self.tableWidget.clear()  # Remove todo o conteúdo da tabela (cabeçalhos inclusos)
        self.tableWidget.setRowCount(0)  # Limpa as linhas para evitar problemas
        self.tableWidget.setColumnCount(0)  # Limpa as colunas para recriar corretamente

        #  Chama novamente a função de exibição com os dados atualizados
        self.exibir_tabela(validado, max_batidas, escala_horarios)

    def exibir_tabela(self, validado, max_batidas, escala_horarios):
        """ Exibe os dados na tabela com colunas dinâmicas, recriando corretamente. """

        max_colunas = max(len(escala_horarios), max_batidas)

        header_labels = ['Local', 'Setor', 'Cadastro', 'Colaborador', 'Escala', 'Data']
        for i in range(max_colunas):
            header_labels.append(f'Escala {i + 1}')
            header_labels.append(f'Batida {i + 1}')

        self.tableWidget.setColumnCount(len(header_labels))
        self.tableWidget.setHorizontalHeaderLabels(header_labels)
        self.tableWidget.setRowCount(len(validado))

        font = self.tableWidget.font()
        font.setPointSize(8)
        self.tableWidget.setFont(font)

        for row, dados in enumerate(validado):
            numloc = dados[2]
            codemp = dados[0]
            numcad = dados[1]

            linha = list(dados)  # Usa todos os dados sem remover nada

            selected_date = self.date_picker.date().toString("dd/MM/yyyy")
            selected_date_dt = datetime.strptime(selected_date, "%d/%m/%Y")

            data_atual_dt = datetime.today()
            hora_atual = datetime.now().strftime("%H:%M")

            entradas_saidas = []
            escalas = []

            for i in range(6, len(linha), 2):
                escala = linha[i]
                escalas.append(escala)

                if i == 6:
                    entradas_saidas.append("entrada")
                else:
                    if escala == linha[i - 2]:
                        entradas_saidas.append(entradas_saidas[-1])
                    else:
                        entradas_saidas.append("saida" if entradas_saidas[-1] == "entrada" else "entrada")

            for col, valor in enumerate(linha):
                item = QTableWidgetItem(str(valor))
                item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)

                if col >= 6 and (col - 6) % 2 == 0:
                    item.setBackground(QColor(80, 80, 80))
                    item.setForeground(QColor(255, 255, 255))

                elif col >= 7 and (col - 6) % 2 == 1:
                    escala_idx = (col - 7) // 2
                    escala = escalas[escala_idx] if escala_idx < len(escalas) else "Falta"
                    entrada_saida = entradas_saidas[escala_idx] if escala_idx < len(entradas_saidas) else "entrada"

                    batida = linha[col]

                    try:
                        escala_time = datetime.strptime(escala, "%H:%M").time() if escala != "Falta" else None
                        batida_time = datetime.strptime(batida, "%H:%M").time() if batida != "Falta" else None
                        hora_atual_time = datetime.strptime(hora_atual, "%H:%M").time()
                    except ValueError:
                        escala_time = None
                        batida_time = None
                        hora_atual_time = None

                    if batida == "Falta":
                        if selected_date_dt > data_atual_dt:
                            item.setText("-")
                            item.setBackground(QColor(255, 255, 255))
                            item.setForeground(QColor(0, 0, 0))
                            encontrou_traco = True
                        elif selected_date_dt.date() == data_atual_dt.date():
                            if escala_time and escala_time > hora_atual_time:
                                item.setText("-")
                                item.setBackground(QColor(255, 255, 255))
                                item.setForeground(QColor(0, 0, 0))
                                encontrou_traco = True
                            else:
                                item.setText("Falta")
                                item.setBackground(QColor(255, 0, 0))
                                item.setForeground(QColor(255, 255, 255))
                        else:
                            item.setText("Falta")
                            item.setBackground(QColor(255, 0, 0))
                            item.setForeground(QColor(255, 255, 255))

                    elif escala_time and batida_time:
                        diferenca = calcular_diferenca_minutos(escala, batida)

                        if diferenca >= 10:
                            if entrada_saida == "entrada" and batida_time > escala_time:
                                item.setBackground(QColor(255, 255, 0))
                            elif entrada_saida == "entrada" and batida_time < escala_time:
                                item.setBackground(QColor(32, 184, 18))
                            elif entrada_saida == "saida" and batida_time > escala_time:
                                item.setBackground(QColor(32, 184, 18))
                            elif entrada_saida == "saida" and batida_time < escala_time:
                                item.setBackground(QColor(255, 255, 0))
                        else:
                            item.setBackground(QColor(255, 255, 255))

                        #  Define fonte preta se fundo não for vermelho
                        if item.background().color().name().lower() != "#ff0000":
                            item.setForeground(QColor(0, 0, 0))

                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tableWidget.setItem(row, col, item)

            # Tratar propagação de "-"
            encontrou_traco = False
            for col in range(6, self.tableWidget.columnCount(), 2):
                item = self.tableWidget.item(row, col + 1)
                if item:
                    if item.text() == "-":
                        encontrou_traco = True
                    if encontrou_traco:
                        item.setText("-")
                        item.setBackground(QColor(255, 255, 255))
                        item.setForeground(QColor(0, 0, 0))

            self.verificar_banco_horas(codemp, numcad, row, self.tableWidget)

            # Recupera a linha já exibida
            linha_atual = [self.tableWidget.item(row, col).text() if self.tableWidget.item(row, col) else ""
                           for col in range(self.tableWidget.columnCount())]

            # Remove CodEmp (índice 0) e CodLoc (índice 2) da linha original
            dados_sem_codemp_codloc = [linha_atual[1]] + linha_atual[3:]  # 1 = numcad, remove 0 e 2

            # Adiciona local e setor corretamente
            local, setor = buscar_setor_e_local(numloc, self.setores_dict, self.dicionario_locais)
            linha_final = [local, setor] + dados_sem_codemp_codloc

            # Atualiza a linha na tabela
            for col, valor in enumerate(linha_final):
                item = self.tableWidget.item(row, col)
                if item:
                    item.setText(str(valor))
                else:
                    novo_item = QTableWidgetItem(str(valor))
                    novo_item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
                    novo_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.tableWidget.setItem(row, col, novo_item)

        fonte = self.tableWidget.font()
        fonte.setPointSize(8)
        self.tableWidget.setFont(fonte)
        self.tableWidget.resizeColumnsToContents()

        largura_extra = 15
        for col in range(self.tableWidget.columnCount()):
            largura_atual = self.tableWidget.columnWidth(col)
            self.tableWidget.setColumnWidth(col, largura_atual + largura_extra)

        self.tableWidget.verticalHeader().setVisible(False)
        self.tableWidget.setStyleSheet("QTableWidget {gridline-color: black;}")
        self.tableWidget.setSortingEnabled(True)
        self.tableWidget.sortItems(3, Qt.SortOrder.AscendingOrder)

    def verificar_banco_horas(self, codemp, numcad, row, tableWidget):
        """Verifica se o colaborador possui banco de horas e pinta as células da linha que estariam em vermelho ou amarelo de verde"""
        query = """
        SELECT POSSBH
        FROM R038HSI
        WHERE NUMEMP = :codemp
          AND TIPCOL = 1
          AND NUMCAD = :numcad
        ORDER BY DATALT DESC
        FETCH FIRST 1 ROW ONLY
        """
        with get_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(query, {'codemp': codemp, 'numcad': numcad})
                result = cursor.fetchone()
                if result and result[0] == "S":
                    for col in range(tableWidget.columnCount()):
                        item = tableWidget.item(row, col)
                        if item:
                            cor = item.background().color().name().lower()
                            if cor in ["#ff0000", "#ffff00"]:  # vermelho ou amarelo
                                item.setBackground(QColor("#00a2ff"))  # verde
                                item.setForeground(QColor(255, 255, 255))

    def excluir_linha_selecionada(self):
        """Exclui a linha selecionada da tabela"""
        selected_row = self.tableWidget.currentRow()

        if selected_row == -1:  # Nenhuma linha foi selecionada
            QMessageBox.warning(self, "Aviso", "Selecione uma linha para excluir.")
            return  # Sai da função para evitar execução desnecessária

        resposta = QMessageBox.question(
            self, "Confirmação", "Tem certeza que deseja excluir esta linha?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No
        )

        if resposta == QMessageBox.StandardButton.Yes:
            self.tableWidget.removeRow(selected_row)
            self.tableWidget.clearSelection()  # Remove qualquer seleção residual
            self.tableWidget.setCurrentCell(-1, -1)  # Remove a célula ativa

    def enviar_email_frequencia(self):
        """Gera e abre um e-mail no Outlook com resumo simplificado de atrasos e faltas."""

        data_selecionada = self.date_picker.date().toString("dd/MM/yyyy")

        dados_frequencia = {}
        contagem_faltas = {}
        contagem_atrasos = {}
        contagem_faltas_bh = {}
        contagem_atrasos_bh = {}
        max_colunas_por_grupo = {}

        cores_locais = {
            "Fabrica De Máquinas": "#003756",
            "Fabrica De Transportadores": "#ffc62e",
            "Adm": "#009c44",
            "Comercial": "#919191"
        }

        for row in range(self.tableWidget.rowCount()):
            if self.tableWidget.isRowHidden(row):
                continue

            local = self.tableWidget.item(row, 0).text()
            setor = self.tableWidget.item(row, 1).text()
            colaborador = self.tableWidget.item(row, 3).text()
            data = self.tableWidget.item(row, 5).text()

            chave_setor = (local, setor)

            escalas = []
            batidas = []
            cores_batidas = []
            houve_problema = False
            contou_falta = False
            contou_falta_bh = False
            colunas_utilizadas = 0

            # Garante que o dicionário tenha as chaves
            contagem_faltas.setdefault(chave_setor, 0)
            contagem_atrasos.setdefault(chave_setor, 0)
            contagem_faltas_bh.setdefault(chave_setor, 0)
            contagem_atrasos_bh.setdefault(chave_setor, 0)
            max_colunas_por_grupo.setdefault(chave_setor, 0)

            for col in range(6, self.tableWidget.columnCount(), 2):
                escala_item = self.tableWidget.item(row, col)
                batida_item = self.tableWidget.item(row, col + 1)

                escala = escala_item.text() if escala_item else ""
                batida = batida_item.text() if batida_item else ""

                cor_batida = batida_item.background().color().name().lower() if batida_item else "#ffffff"
                texto_batida = batida.strip().lower()

                if not houve_problema and cor_batida in ["#ff0000", "#ffff00", "#00a2ff"]:
                    houve_problema = True

                # Falta
                if cor_batida == "#ff0000" and texto_batida == "falta" and not contou_falta:
                    contagem_faltas[chave_setor] += 1
                    contou_falta = True

                # Atraso
                elif cor_batida == "#ffff00":
                    contagem_atrasos[chave_setor] += 1

                # Falta BH
                elif cor_batida == "#00a2ff" and texto_batida == "falta" and not contou_falta_bh:
                    contagem_faltas_bh[chave_setor] += 1
                    contou_falta_bh = True

                # Atraso BH
                elif cor_batida == "#00a2ff" and texto_batida != "falta":
                    contagem_atrasos_bh[chave_setor] += 1

                if escala or batida:
                    escalas.append(escala)
                    batidas.append(batida)
                    cores_batidas.append(cor_batida)
                    colunas_utilizadas += 1

            max_colunas_por_grupo[chave_setor] = max(max_colunas_por_grupo[chave_setor], colunas_utilizadas)

            if houve_problema:
                dados_frequencia.setdefault(chave_setor, []).append({
                    "nome": colaborador,
                    "data": data,
                    "escalas": escalas,
                    "batidas": batidas,
                    "cores_batidas": cores_batidas
                })

        if not dados_frequencia:
            QMessageBox.information(self, "Aviso", "Nenhum colaborador com atraso ou falta. Nenhum e-mail será gerado.")
            return

        corpo_email = f"<span style='font-size:16px;'><b>Relatório de frequência do dia {data_selecionada}:</b></span><br><br>"

        corpo_email += """
        <b>Legenda:</b><br>
        <table border="1" cellspacing="0" cellpadding="3" style="border-collapse: collapse; width: auto; font-size: 14px;">
            <tr>
                <td style="background-color: #ff0000; color: white; text-align: center; white-space: nowrap;">Falta</td>
                <td style="background-color: #ffff00; text-align: center; white-space: nowrap;">Atraso</td>
                <td style="background-color: #20b812; text-align: center; white-space: nowrap;">Hora Extra</td>
                <td style="background-color: #00a2ff; text-align: center; white-space: nowrap;">Considera BH</td>
                <td style="background-color: #ffffff; text-align: center; white-space: nowrap;">Dentro do Horário</td>
            </tr>
        </table>
        <br>
        """

        corpo_email += "<b>Resumo:</b><br>"

        setores_ordenados = sorted(dados_frequencia.keys(), key=lambda x: (x[0], x[1]))

        for local, setor in setores_ordenados:
            f = contagem_faltas.get((local, setor), 0)
            a = contagem_atrasos.get((local, setor), 0)
            f_bh = contagem_faltas_bh.get((local, setor), 0)
            a_bh = contagem_atrasos_bh.get((local, setor), 0)

            if any([f, a, f_bh, a_bh]):
                corpo_email += f"- {local} - {setor} ("
                partes = []
                if f > 0:
                    partes.append(f"{f} faltas")
                if a > 0:
                    partes.append(f"{a} atrasos")
                if f_bh > 0:
                    partes.append(f"{f_bh} faltas BH")
                if a_bh > 0:
                    partes.append(f"{a_bh} atrasos BH")

                corpo_email += " / ".join(partes) + ")<br>"

        corpo_email += '<table border="1" cellspacing="0" cellpadding="3" style="border-collapse: collapse; width: 100%; font-size: 14px; text-align: center;">'
        corpo_email += """
        <tr style="background-color: #f2f2f2;">
            <th>Local</th>
            <th>Setor</th>
            <th>Nome</th>
            <th>Data</th>
        """

        num_colunas_max = max(max_colunas_por_grupo.values(), default=0)
        for i in range(num_colunas_max):
            corpo_email += f"<th>Escala {i + 1}</th><th>Batida {i + 1}</th>"
        corpo_email += "</tr>"

        for (local, setor) in sorted(dados_frequencia.keys(), key=lambda x: (x[0], x[1])):
            colaboradores = dados_frequencia[(local, setor)]
            cor_setor = cores_locais.get(local, "#000000")

            for idx, colaborador in enumerate(colaboradores):
                corpo_email += "<tr>"
                if idx == 0:
                    corpo_email += f'<td rowspan="{len(colaboradores)}" style="background-color:{cor_setor}; color:#fff; padding: 3px;">{local}</td>'
                    corpo_email += f'<td rowspan="{len(colaboradores)}" style="padding: 3px;">{setor}</td>'

                corpo_email += f"<td>{colaborador['nome']}</td><td>{colaborador['data']}</td>"

                for i in range(num_colunas_max):
                    escala = colaborador["escalas"][i] if i < len(colaborador["escalas"]) else "-"
                    batida = colaborador["batidas"][i] if i < len(colaborador["batidas"]) else "-"
                    cor = colaborador["cores_batidas"][i] if i < len(colaborador["cores_batidas"]) else "#ffffff"

                    corpo_email += f'<td style="background-color: #505050; color: #FFFFFF; padding: 3px;">{escala}</td>'
                    corpo_email += f'<td style="background-color: {cor}; padding: 3px;">{batida}</td>'

                corpo_email += "</tr>"

        corpo_email += "</table><br>"

        outlook = win32com.client.Dispatch("Outlook.Application")
        email = outlook.CreateItem(0)
        email.Subject = f"Relatório de Frequência - Atrasos e Faltas ({data_selecionada})"
        signature = email.HTMLBody
        email.HTMLBody = corpo_email + "<br><br>" + signature

        self.current_mail = email
        if self.aba_historico:
            self.aba_historico.current_mail = self.current_mail

        email.Display()
        QMessageBox.information(self, "Sucesso", "E-mail gerado com sucesso no Outlook!")

    def apply_global_filter(self, text):
        """Filtra as linhas da tabela com base na pesquisa parcial e permite busca por coluna específica (ex: Setor:"Produção")."""

        current_tab = self.tabs.currentWidget()
        table_widget = None

        # Localiza a QTableWidget dentro da aba ativa
        for i in range(current_tab.layout().count()):
            widget = current_tab.layout().itemAt(i).widget()
            if isinstance(widget, QTableWidget):
                table_widget = widget
                break

        if table_widget is None:
            return

        terms = [term.strip() for term in text.split('/') if term.strip()]

        # Se não houver termos, exibe todas as linhas
        if not terms:
            for row in range(table_widget.rowCount()):
                table_widget.setRowHidden(row, False)
            return

        #  Criar um dicionário com os nomes das colunas para facilitar a busca específica
        headers = {}
        for col in range(table_widget.columnCount()):
            header_item = table_widget.horizontalHeaderItem(col)
            if header_item:  # Somente adiciona colunas que possuem nome
                headers[header_item.text().strip().lower()] = col

        # Percorre todas as linhas da tabela
        for row in range(table_widget.rowCount()):
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
                        item = table_widget.item(row, col_idx)

                        if item and item.text().strip().lower() == valor_busca.lower():
                            term_match = True  # Encontrou o termo exato na coluna correta

                else:
                    #  Busca geral (qualquer coluna)
                    for col in range(table_widget.columnCount()):
                        item = table_widget.item(row, col)
                        if item and term.lower() in item.text().strip().lower():
                            term_match = True  # Encontrou o termo em alguma coluna
                            break  # Já encontrou, não precisa verificar o resto

                if not term_match:
                    row_matches_all_terms = False
                    break  # Se um termo não foi encontrado, a linha não será exibida

            table_widget.setRowHidden(row,
                                      not row_matches_all_terms)  # Esconde apenas se **não** corresponde a todos os termos

    def voltar_menu(self):
        self.menu_principal = ControlesRH()
        self.menu_principal.show()
        self.close()

def obter_dados(selected_date):
    """Obtém as batidas de ponto do banco para a data selecionada, considerando apenas quem deve bater ponto."""

    with get_connection() as connection:

        batidas_result = []
        with connection.cursor() as cursor:
            # Buscar colaboradores que devem bater ponto
            cursor.execute(SQL_CONSIDERAPONTO)
            colaboradores_com_ponto = {f"{row[0]}-{row[1]}" for row in cursor.fetchall()}

            # Executar consulta de batidas normais
            cursor.execute(SQL_BATIDAS_QUERY, {'selected_date': selected_date})
            for row in cursor.fetchall():
                chave = f"{row[0]}-{row[1]}"
                if chave in colaboradores_com_ponto and row[4] != 62:
                    batidas_result.append(row)

            # Executar consulta de batidas para código de escala 62
            cursor.execute(SQL_QUERY2, {'selected_date': selected_date})
            for row in cursor.fetchall():
                chave = f"{row[0]}-{row[1]}"
                if chave in colaboradores_com_ponto and row[4] == 62:
                    batidas_result.append(row)

    return batidas_result



# Consulta para pegar os horários da escala
SQL_ESCALA_QUERY = """
SELECT 
    CODESC, 
    DESSIM
FROM R006ESC
WHERE CodEsc = :codesc
"""

def carregar_todas_escalas():
    """Carrega todas as escalas do banco de dados e armazena em um dicionário."""
    escalas_dict = {}

    with get_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute("SELECT CODESC, DESSIM FROM R006ESC")  # Busca todas as escalas de uma vez
            for codesc, dessim in cursor.fetchall():
                escalas_dict[codesc] = dessim

    return escalas_dict


def obter_escala(codesc):
    """Obtém a escala de um código específico."""
    with get_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(SQL_ESCALA_QUERY, {'codesc': codesc})
            escala_result = cursor.fetchone()

    return escala_result[1] if escala_result else None  # Retorna apenas DESSIM

def extrair_horarios_escala(descricao_escala, data_selecionada):
    """Extrai os horários da escala para o dia selecionado e retorna uma lista de horários."""
    if not descricao_escala:
        return ["Sem Escala"]

    descricao_upper = descricao_escala.upper()

    # Função para capturar os horários corretamente
    def capturar_horarios(texto):
        horarios = re.findall(r"\d{1,2}[:.]?\d{2}|\d{4}", texto)

        def formatar_horario(h):
            if ":" in h:
                horas, minutos = h.split(":")
                return f"{int(horas):02}:{minutos}"
            elif len(h) == 3:  # Exemplo: "730" -> "07:30"
                return f"0{h[0]}:{h[1:]}"
            elif len(h) == 4:  # Exemplo: "0730" -> "07:30"
                return f"{h[:2]}:{h[2:]}"
            return h

        return [formatar_horario(h) for h in horarios]

    # Obter o dia da semana a partir da data selecionada
    dia_da_semana = datetime.strptime(data_selecionada, "%d/%m/%Y").strftime("%A").lower()
    dias_traduzidos = {
        "monday": "segunda",
        "tuesday": "terça",
        "wednesday": "quarta",
        "thursday": "quinta",
        "friday": "sexta",
        "saturday": "sábado",
        "sunday": "domingo"
    }
    dia_selecionado = dias_traduzidos.get(dia_da_semana, "domingo")

    # Ordem dos dias para suportar intervalos
    dias_ordem = ["segunda", "terça", "quarta", "quinta", "sexta", "sábado", "domingo"]

    horarios_escala = ["Sem Escala"] * 4

    #  Verifica se a descrição contém intervalo (ex: "TERÇA A SÁBADO")
    match_intervalo = re.search(
        r"(SEGUNDA|TERÇA|QUARTA|QUINTA|SEXTA|SÁBADO|DOMINGO)\s*A\s*(SEGUNDA|TERÇA|QUARTA|QUINTA|SEXTA|SÁBADO|DOMINGO)",
        descricao_upper
    )
    if match_intervalo:
        dia_inicio, dia_fim = match_intervalo.groups()
        dia_inicio = dia_inicio.lower()
        dia_fim = dia_fim.lower()
        idx_inicio = dias_ordem.index(dia_inicio)
        idx_fim = dias_ordem.index(dia_fim)

        if idx_inicio <= idx_fim:
            intervalo = dias_ordem[idx_inicio:idx_fim+1]
        else:
            intervalo = dias_ordem[idx_inicio:] + dias_ordem[:idx_fim+1]

        if dia_selecionado in intervalo:
            restante_texto = descricao_upper[match_intervalo.end():]
            horarios_extraidos = capturar_horarios(restante_texto)[:4]
            return horarios_extraidos if horarios_extraidos else ["Sem Escala"]

    #  Dias isolados (sexta, sábado, domingo) seguem igual
    if dia_selecionado == "sexta":
        match_sexta = re.search(r"SEXTA", descricao_upper)
        if match_sexta:
            restante_texto = descricao_upper[match_sexta.end():]
            horarios_extraidos = capturar_horarios(restante_texto)[:4]
            return horarios_extraidos if horarios_extraidos else ["Sem Escala"]
        return ["Sem Escala"]

    if dia_selecionado == "sábado":
        match_sabado = re.search(r"SABADO|SÁBADO", descricao_upper)
        if match_sabado:
            restante_texto = descricao_upper[match_sabado.end():]
            horarios_extraidos = capturar_horarios(restante_texto)[:4]
            return horarios_extraidos if horarios_extraidos else ["Sem Escala"]
        return ["Sem Escala"]

    if dia_selecionado == "domingo":
        match_domingo = re.search(r"DOMINGO", descricao_upper)
        if match_domingo:
            restante_texto = descricao_upper[match_domingo.end():]
            horarios_extraidos = capturar_horarios(restante_texto)[:4]
            return horarios_extraidos if horarios_extraidos else ["Sem Escala"]
        return ["Sem Escala"]

    return horarios_escala if any(h != "Sem Escala" for h in horarios_escala) else ["Sem Escala"]


def validar_batidas(selected_date):
    """Processa os dados e junta as escalas com as batidas, ajustando corretamente as batidas para os horários certos."""
    batidas = obter_dados(selected_date)
    escalas_dict = carregar_todas_escalas()  # Agora carregamos todas as escalas uma vez

    resultados = []  # Lista onde armazenamos os resultados
    max_batidas = 0  # Para rastrear o maior número de batidas registradas
    escala_horarios = []

    for numemp, numcad, numloc, nomfun, codesc, dataformatada, horarios_batidos in batidas:
        # Se o código for 62, exibe a data como "Ontem (DD/MM/YYYY) e Hoje (DD/MM/YYYY)"
        if codesc == 62:
            ontem_str = dataformatada.split(" a ")[0].strip()  # Pegamos apenas a primeira data (ontem)
            ontem = datetime.strptime(ontem_str, "%d/%m/%Y")  # Convertendo corretamente para datetime
            hoje = ontem + timedelta(days=1)
            dataformatada = f"{ontem.strftime('%d/%m/%Y')} a {hoje.strftime('%d/%m/%Y')}"  # Exibe ambos os dias corretamente

        # Pega a escala do colaborador
        descricao_escala = escalas_dict.get(codesc, None)
        escala_horarios = extrair_horarios_escala(descricao_escala, selected_date)

        #  Se não tem escala, ignora esse colaborador de vez
        if not escala_horarios or escala_horarios == ["Sem Escala"]:
            continue

        # Quebrar horários batidos em lista diretamente do SQL
        batidas_lista = horarios_batidos.split(" / ") if horarios_batidos else []

        # Gera a escala ajustada com base nas batidas ANTES de alinhar
        escala_horarios = gerar_escala_ajustada(escala_horarios, batidas_lista)

        # Alinhamento correto das batidas com a escala
        batidas_alinhadas = alinhar_batidas(escala_horarios, batidas_lista)

        # Atualizar o número máximo de batidas encontradas (ANTES de alinhar)
        max_batidas = max(max_batidas, len(batidas_lista), len(escala_horarios))

        # Monta a linha para exibição, garantindo que tenha pares suficientes de escala/batida
        linha = [numemp, numcad, numloc, nomfun, codesc, dataformatada]
        for i in range(len(batidas_alinhadas)):
            linha.append(escala_horarios[i] if i < len(escala_horarios) else "Falta")
            linha.append(batidas_alinhadas[i])
        # Obtém a data e hora atuais
        data_atual = datetime.now().strftime("%d/%m/%Y")
        hora_atual = datetime.now().strftime("%H:%M")

        # Converte a data e hora para objetos datetime para comparação
        try:
            data_atual_dt = datetime.strptime(data_atual, "%d/%m/%Y")
            hora_atual_time = datetime.strptime(hora_atual, "%H:%M").time()

            # Verifica se a escala tem um intervalo (DD/MM/YYYY a DD/MM/YYYY)
            if " a " in dataformatada:
                partes_data = dataformatada.split(" a ")
                escala_data_dt = datetime.strptime(partes_data[0].strip(), "%d/%m/%Y")  # Primeira data
            else:
                escala_data_dt = datetime.strptime(dataformatada, "%d/%m/%Y")

            # Se a data da escala for hoje e a primeira escala ainda não chegou, não incluir nos resultados
            if escala_data_dt == data_atual_dt and escala_horarios and escala_horarios[0] != "Falta":
                primeiro_horario_escala = datetime.strptime(escala_horarios[0], "%H:%M").time()

                if primeiro_horario_escala > hora_atual_time:
                    continue  # Pula este colaborador, pois a escala ainda não começou
        except ValueError:
            pass  # Caso alguma conversão falhe, simplesmente não faz essa verificação

        # Se passou pela verificação, adiciona a linha nos resultados
        resultados.append(linha)

    return resultados, max_batidas, escala_horarios


def gerar_escala_ajustada(escala_horarios, batidas_lista):
    if not escala_horarios:
        return ["Escala não encontrada"] * len(batidas_lista)  # Garante que haja pelo menos colunas suficientes

    nova_escala = []  # Escala ajustada com repetições corretas

    for batida in batidas_lista:
        melhor_escala = None
        menor_diferenca = float('inf')

        for escala in escala_horarios:
            diferenca = calcular_diferenca_minutos(batida, escala)
            if diferenca < menor_diferenca:
                menor_diferenca = diferenca
                melhor_escala = escala

        if melhor_escala:
            nova_escala.append(melhor_escala)  #  Agora adiciona cada ocorrência necessária

    # Garantir que os horários da escala original sejam mantidos caso ainda não tenham sido adicionados
    for escala in escala_horarios:
        if escala not in nova_escala:
            nova_escala.append(escala)

    #  Mantém a ordem correta da escala original
    nova_escala.sort(key=lambda x: escala_horarios.index(x) if x in escala_horarios else float('inf'))

    return nova_escala

def alinhar_batidas(escala_horarios, batidas_lista):
    """ Aloca corretamente as batidas nos horários da escala, garantindo alinhamento adequado. """
    total_escala = len(escala_horarios)
    total_batidas = len(batidas_lista)

    # Inicializa a lista de batidas alinhadas com "Falta"
    batidas_alinhadas = ["Falta"] * total_escala

    index_batida = 0  # Índice para percorrer a lista de batidas

    for i in range(total_escala):
        horario_escala = escala_horarios[i]

        while index_batida < total_batidas:
            batida_atual = batidas_lista[index_batida]

            # Se há um próximo horário de escala, comparar qual está mais perto
            if i + 1 < total_escala:
                proximo_horario_escala = escala_horarios[i + 1]
                diferenca_atual = calcular_diferenca_minutos(horario_escala, batida_atual)
                diferenca_proximo = calcular_diferenca_minutos(proximo_horario_escala, batida_atual)

                if diferenca_proximo < diferenca_atual:
                    break

            # Se a batida for antes da primeira escala e estiver dentro de 10 minutos, alocar
            if i == 0 and calcular_diferenca_minutos(batida_atual, horario_escala) < 10:
                batidas_alinhadas[i] = batida_atual
                index_batida += 1
                continue

            # Verifica se já tem batida alocada nesse índice
            if batidas_alinhadas[i] == "Falta":
                batidas_alinhadas[i] = batida_atual
            else:
                # Se já tem batida aqui, pulamos para o próximo índice livre para colocar a nova batida
                for j in range(i + 1, total_escala):
                    if batidas_alinhadas[j] == "Falta":
                        batidas_alinhadas[j] = batida_atual
                        break

            index_batida += 1
            break  # Sai do loop para manter a estrutura da escala
    return batidas_alinhadas


def calcular_diferenca_minutos(horario1, horario2):
    """ Calcula a diferença em minutos entre dois horários no formato HH:MM """
    try:
        h1 = datetime.strptime(horario1, "%H:%M")
        h2 = datetime.strptime(horario2, "%H:%M")
        return abs((h1 - h2).total_seconds() / 60)  # Retorna a diferença absoluta em minutos
    except ValueError:
        return float("inf")  # Caso não seja possível converter, retorna infinito

def carregar_setores():
    """Carrega todos os setores do banco de dados em um dicionário {NUMLOC: CODLOC} para busca rápida."""
    setores_dict = {}

    with get_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(SQL_Setor)
            setores_dict = {str(row[1]).strip(): str(row[0]).strip() for row in cursor.fetchall()}  # {NUMLOC: CODLOC}

    return setores_dict


def buscar_setor_e_local(numloc, setores_dict, dicionario_locais):
    """Busca o setor e local baseado no NUMLOC, cruzando com a planilha Locais.xlsx."""

    codloc = setores_dict.get(str(numloc).strip())  # Obtém o CODLOC correspondente ao NUMLOC do funcionário

    # Define valores padrão caso não encontre
    local, setor_final = "Não encontrado", "Não encontrado"

    # Se não encontrar o codloc no dicionário, já retorna
    if not codloc:
        return local, setor_final

        #  Tenta encontrar o CODLOC exato na planilha, reduzindo se necessário
    while codloc:
        for codloc_planilha, descricao in dicionario_locais.items():
            if codloc_planilha.startswith(codloc):  # Verifica se a linha começa com o CODLOC procurado
                partes_descricao = descricao.split(",", 1)  # Divide apenas na primeira vírgula
                local = partes_descricao[0].strip()  # A primeira parte define o Local

                # O setor final será a última parte do CODLOC na planilha
                setor_final = codloc_planilha.rsplit(",", 1)[-1].strip() if "," in codloc_planilha else local
                return local, setor_final  # Retorna imediatamente se encontrou

        #  Se não encontrou, tenta reduzir removendo a última parte do CODLOC
        if "." in codloc:
            codloc = ".".join(codloc.split(".")[:-1])  # Remove a última parte do código
        else:
            break  # Sai do loop se não puder mais reduzir

    return local, setor_final  # Agora o "Não encontrado" só aparece uma única vez


# Função principal para rodar o aplicativo
def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = controlefrequencia()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()