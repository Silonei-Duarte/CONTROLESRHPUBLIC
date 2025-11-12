import colorsys
import sys
from PyQt6.QtCore import QCoreApplication, Qt, QDate
QCoreApplication.setAttribute(Qt.ApplicationAttribute.AA_ShareOpenGLContexts)
from PyQt6.QtWidgets import (QApplication, QWidget, QHBoxLayout, QMessageBox, QCheckBox,
                             QTabWidget, QDateEdit, QComboBox, QFileDialog)
import fitz
from graphviz import Digraph
import re
import pandas as pd
from datetime import datetime
import random
import json
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QTableWidget, QTableWidgetItem, QLabel, QSizePolicy, QSpacerItem, QPushButton
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon, QFont, QIntValidator
import os
from PyQt6.QtWidgets import QLineEdit, QTextEdit
from Painel_Gestores import PainelGestores
from Painel_Setores import PainelSetores
from Database import get_connection
import pandas as pd
from PyQt6.QtWidgets import QFileDialog, QMessageBox
from PyQt6.QtCore import Qt
from Painel_Setores_Grafico import GraficoSetoresApp


class PainelConsultaFuncionarios(QWidget):
    def __init__(self):
        super().__init__()
        self.df_funcionarios = pd.DataFrame()
        self.df_ativos = pd.DataFrame()
        self.df_desligados = pd.DataFrame()
        self.janela_detalhes = None
        self.sinal_setores_conectado = False
        self.carregar_dicionario_locais_com_busca()
        self.carregar_nomes_filtrados()
        self.initUI()
        QApplication.processEvents()
        self.consultar_funcionarios()

        # Inicializa painel de setores e conecta evento de clique duplo
        self.painel_setores = PainelSetores(self)

        self.showMaximized()

    def initUI(self):
        self.setWindowTitle("Consulta de Funcionários e Gestores")

        icon_path = os.path.join(os.path.dirname(__file__), 'icone.ico')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        botoes_layout = QHBoxLayout()
        botoes_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)

        spacer = QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
        botoes_layout.addItem(spacer)

        btn_grafico = QPushButton("Abrir Gráfico")
        btn_grafico.setFixedSize(120, 30)
        btn_grafico.clicked.connect(self.abrir_graficos)
        botoes_layout.addWidget(btn_grafico)

        self.export_excel_btn = QPushButton("Exportar para Excel")
        self.export_excel_btn.setFixedSize(120, 30)
        self.export_excel_btn.clicked.connect(self.export_to_excel)
        botoes_layout.addWidget(self.export_excel_btn)

        btn_config_filtros = QPushButton("Configurar Filtros")
        btn_config_filtros.setFixedSize(140, 30)
        btn_config_filtros.clicked.connect(self.configurar_filtros)
        botoes_layout.addWidget(btn_config_filtros)

        btn_organograma = QPushButton("Organograma")
        btn_organograma.setFixedSize(140, 30)
        btn_organograma.clicked.connect(self.abrir_organograma)
        botoes_layout.addWidget(btn_organograma)

        self.btn_voltar = QPushButton("Voltar ao Menu")
        self.btn_voltar.setFixedSize(140, 30)
        self.btn_voltar.clicked.connect(self.voltar_menu)
        botoes_layout.addWidget(self.btn_voltar)

        layout.addLayout(botoes_layout)

        data_layout = QHBoxLayout()

        self.date_inicio_picker = QDateEdit()
        self.date_inicio_picker.setDate(QDate(2000, 1, 1))
        self.date_inicio_picker.setCalendarPopup(True)
        self.date_inicio_picker.setDisplayFormat("dd/MM/yyyy")
        data_layout.addWidget(QLabel("Admitidos Desde:"))
        data_layout.addWidget(self.date_inicio_picker)

        self.date_fim_picker = QDateEdit()
        self.date_fim_picker.setDate(QDate.currentDate())
        self.date_fim_picker.setCalendarPopup(True)
        self.date_fim_picker.setDisplayFormat("dd/MM/yyyy")
        data_layout.addWidget(QLabel("Admitidos/Desligados até:"))
        data_layout.addWidget(self.date_fim_picker)

        self.dias_desligados = QLineEdit()
        self.dias_desligados.setFixedWidth(40)
        self.dias_desligados.setText("30")
        self.dias_desligados.setValidator(QIntValidator(1, 360, self))
        data_layout.addWidget(QLabel("Voltando Intervalo de Dias:"))
        data_layout.addWidget(self.dias_desligados)

        btn_buscar = QPushButton("Buscar")
        btn_buscar.setFixedWidth(100)
        btn_buscar.clicked.connect(self.consultar_funcionarios)
        data_layout.addWidget(btn_buscar)

        data_layout.addStretch()
        layout.addLayout(data_layout)

        self.label_resumo = QLabel()
        fonte = QFont()
        fonte.setPointSize(9)
        self.label_resumo.setFont(fonte)
        layout.addWidget(self.label_resumo)

        self.search_field = QLineEdit()
        self.search_field.setPlaceholderText("Filtrar usando / entre dados. Também pode digitar COLUNA:DADO")
        self.search_field.textChanged.connect(self.apply_global_filter)
        layout.addWidget(self.search_field)

        self.tab_widget = QTabWidget()

        # Aba Funcionários
        self.tab_funcionarios = QWidget()
        layout_funcionarios = QVBoxLayout(self.tab_funcionarios)

        checkbox_layout = QHBoxLayout()
        checkbox_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)

        self.checkbox_desligados = QCheckBox("Exibir desligados")
        self.checkbox_desligados.setChecked(False)
        self.checkbox_desligados.stateChanged.connect(self.atualizar_exibicao_df)
        checkbox_layout.addWidget(self.checkbox_desligados)

        self.checkbox_admitidos = QCheckBox("Exibir admitidos")
        self.checkbox_admitidos.setChecked(False)
        self.checkbox_admitidos.stateChanged.connect(self.atualizar_exibicao_df)
        checkbox_layout.addWidget(self.checkbox_admitidos)

        layout_funcionarios.addLayout(checkbox_layout)

        self.tableWidget = QTableWidget()
        layout_funcionarios.addWidget(self.tableWidget)
        self.tableWidget.verticalHeader().setVisible(False)

        # Aba Setores
        self.tab_setores = QWidget()
        layout_setores = QVBoxLayout(self.tab_setores)

        self.setores_table = QTableWidget()
        self.setores_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.setores_table.verticalHeader().setVisible(False)
        layout_setores.addWidget(self.setores_table)

        # Aba Gestores
        self.tab_gestores = QWidget()
        layout_gestores = QVBoxLayout(self.tab_gestores)

        self.label_info_gestores = QLabel("Lista de Gestores e Quantidade de Colaboradores")
        font_info = QFont()
        font_info.setPointSize(10)
        self.label_info_gestores.setFont(font_info)
        layout_gestores.addWidget(self.label_info_gestores)

        self.gestores_table = QTableWidget()
        self.gestores_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.gestores_table.verticalHeader().setVisible(False)
        layout_gestores.addWidget(self.gestores_table)

        # Tabs
        self.tab_widget.addTab(self.tab_funcionarios, "Funcionários")
        self.tab_widget.addTab(self.tab_setores, "Setores")
        self.tab_widget.addTab(self.tab_gestores, "Gestores")
        self.tab_widget.currentChanged.connect(self.on_tab_changed)

        layout.addWidget(self.tab_widget)
        self.setLayout(layout)

        self.current_mail = None

    def consultar_funcionarios(self):
        # Armazenar o texto do filtro atual
        texto_filtro_atual = ""
        if hasattr(self, 'search_field'):
            texto_filtro_atual = self.search_field.text()

        data_ref = self.date_fim_picker.date().toString("dd/MM/yyyy")
        data_inicio = self.date_inicio_picker.date().toString("dd/MM/yyyy")
        data_fim = self.date_fim_picker.date().toString("dd/MM/yyyy")
        try:
            dias_desligado = int(self.dias_desligados.text())
        except:
            dias_desligado = 30

        query = f"""
        SELECT 
            C.NUMEMP,
            H.CODLOC AS CODLOC_FUNCIONARIO,
            C.TIPCOL,
            C.NUMCAD,
            FUNC.NOMFUN AS NOM_FUNCIONARIO,
            CAR.TITCAR AS CARGO,
            TO_CHAR(
                CASE 
                    WHEN HSA.TIPSAL = 2 THEN HSA.VALSAL * 220
                    ELSE HSA.VALSAL
                END, 
                '9999990.99'
            ) AS SALARIO,
            TO_CHAR(FUNC.DATADM, 'DD/MM/YYYY') AS DATADM,
            TO_CHAR(A.DATAFA, 'DD/MM/YYYY') AS DATAFA,
            C.USU_empGes,
            HGESTOR.CODLOC AS CODLOC_GESTOR,
            C.USU_TipGestor,
            GESTOR.NOMFUN AS NOM_GESTOR,
            HSA.DATALT
        FROM R034CPL C
        LEFT JOIN R034FUN FUNC 
               ON FUNC.NUMEMP = C.NUMEMP 
              AND FUNC.TIPCOL = C.TIPCOL 
              AND FUNC.NUMCAD = C.NUMCAD
        LEFT JOIN R024CAR CAR 
               ON FUNC.CODCAR = CAR.CODCAR 
              AND FUNC.ESTCAR = CAR.ESTCAR
        LEFT JOIN R016HIE H 
               ON FUNC.NUMLOC = H.NUMLOC
        LEFT JOIN R034FUN GESTOR 
               ON GESTOR.NUMEMP = C.USU_empGes
              AND GESTOR.TIPCOL = C.USU_TipGestor
              AND GESTOR.NUMCAD = C.USU_CadGestor
        LEFT JOIN R016HIE HGESTOR 
               ON GESTOR.NUMLOC = HGESTOR.NUMLOC
        LEFT JOIN R038AFA A 
               ON A.NUMEMP = C.NUMEMP
              AND A.TIPCOL = C.TIPCOL
              AND A.NUMCAD = C.NUMCAD
              AND A.SITAFA = '007'
        LEFT JOIN (
            SELECT NUMEMP, TIPCOL, NUMCAD, DATALT, TIPSAL, VALSAL
            FROM (
                SELECT HSA.*,
                       ROW_NUMBER() OVER (PARTITION BY NUMEMP, TIPCOL, NUMCAD ORDER BY DATALT DESC) AS RN
                FROM R038HSA HSA
            )
            WHERE RN = 1
        ) HSA
               ON HSA.NUMEMP = C.NUMEMP
              AND HSA.TIPCOL = C.TIPCOL
              AND HSA.NUMCAD = C.NUMCAD
        WHERE 
            C.NUMEMP IN (10, 16, 17, 19, 11)
        AND (
            (
                FUNC.DATADM BETWEEN TO_DATE('{data_inicio}', 'DD/MM/YYYY') AND TO_DATE('{data_fim}', 'DD/MM/YYYY')
                AND (A.DATAFA IS NULL OR A.DATAFA >= TO_DATE('{data_fim}', 'DD/MM/YYYY'))
                AND FUNC.SITAFA NOT IN ('003', '024', '913')
            )
            OR (
                FUNC.SITAFA = '007'
                AND A.DATAFA >= TO_DATE('{data_ref}', 'DD/MM/YYYY') - {dias_desligado}
            )
        )
        """

        with get_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(query)
                colunas = [desc[0] for desc in cursor.description]
                registros = [dict(zip(colunas, row)) for row in cursor.fetchall()]

        df = pd.DataFrame(registros)

        if df.empty:
            self.df_ativos = pd.DataFrame()
            self.df_desligados = pd.DataFrame()
            self.df_admitidos = pd.DataFrame()
            self.tableWidget.setRowCount(0)
            self.tableWidget.setColumnCount(0)
            self.setores_table.setRowCount(0)
            self.setores_table.setColumnCount(0)
            return

        df = df.fillna("").astype(str)
        for col in df.columns:
            df[col] = df[col].astype(str).str.strip()

        if hasattr(self, 'nomes_filtrados') and self.nomes_filtrados:
            df = df[~df["NOM_FUNCIONARIO"].isin(self.nomes_filtrados)]

        tipos_dict = {"1": "Empregado", "2": "Terceiro", "3": "Parceiro",
                      "1.0": "Empregado", "2.0": "Terceiro", "3.0": "Parceiro"}
        df["TIPCOL"] = df["TIPCOL"].replace(tipos_dict)
        df["USU_TIPGESTOR"] = df["USU_TIPGESTOR"].replace(tipos_dict)

        df.rename(columns={
            "NUMEMP": "Empresa",
            "CODLOC_FUNCIONARIO": "Setor Funcionário",
            "TIPCOL": "Tipo",
            "NUMCAD": "Cadastro",
            "NOM_FUNCIONARIO": "Colaborador",
            "CARGO": "Cargo",
            "SALARIO": "Salário Mensal",
            "DATADM": "Data Admissão",
            "DATAFA": "Data Desligamento",
            "USU_EMPGES": "Empresa Gestor",
            "CODLOC_GESTOR": "Setor Gestor",
            "USU_TIPGESTOR": "Tipo Gestor",
            "NOM_GESTOR": "Gestor"
        }, inplace=True)

        df[["Local", "Setor"]] = df["Setor Funcionário"].apply(
            lambda codloc: pd.Series(self.buscar_local_setor(codloc))
        )
        df[["Local Gestor", "Setor Gestor"]] = df["Setor Gestor"].apply(
            lambda codloc: pd.Series(self.buscar_local_setor(codloc))
        )

        colunas_ordenadas = [
            "Local", "Setor", "Tipo", "Cadastro",
            "Colaborador", "Cargo", "Salário Mensal",
            "Data Admissão", "Data Desligamento",
            "Local Gestor", "Setor Gestor", "Tipo Gestor", "Gestor"
        ]
        df = df[colunas_ordenadas]

        df["Data Admissão_dt"] = pd.to_datetime(df["Data Admissão"], format="%d/%m/%Y", errors="coerce")
        df["Data Desligamento_dt"] = pd.to_datetime(df["Data Desligamento"], format="%d/%m/%Y", errors="coerce")

        data_ref_dt = pd.to_datetime(self.date_fim_picker.date().toString("dd/MM/yyyy"), format="%d/%m/%Y")
        limite_inferior = data_ref_dt - pd.Timedelta(days=int(self.dias_desligados.text() or "30"))

        df_ativos = df[
            (df["Data Admissão_dt"] <= data_ref_dt) &
            (df["Data Desligamento_dt"].isna() | (df["Data Desligamento_dt"] >= data_ref_dt))
            ].copy()

        df_desligados = df[
            (df["Data Desligamento_dt"] >= limite_inferior) &
            (df["Data Desligamento_dt"] < data_ref_dt)
            ].copy()

        df_admitidos = df[
            (df["Data Admissão_dt"] >= limite_inferior) &
            (df["Data Admissão_dt"] <= data_ref_dt)
            ].copy()

        for d in (df, df_ativos, df_desligados, df_admitidos):
            d.drop(columns=["Data Admissão_dt", "Data Desligamento_dt"], inplace=True, errors="ignore")

        self.df_funcionarios = df.copy()
        self.df_ativos = df_ativos
        self.df_desligados = df_desligados
        self.df_admitidos = df_admitidos

        if self.tab_widget.currentIndex() == 1:
            self.painel_setores.atualizar_tabela_setores()

        self.atualizar_exibicao_df()
        if texto_filtro_atual:
            self.apply_global_filter(texto_filtro_atual)

    def atualizar_exibicao_df(self):
        if hasattr(self, "checkbox_desligados") and self.checkbox_desligados.isChecked():
            df = self.df_desligados.copy()
        elif hasattr(self, "checkbox_admitidos") and self.checkbox_admitidos.isChecked():
            df = self.df_admitidos.copy()
        else:
            df = self.df_ativos.copy()

        texto_filtro_atual = self.search_field.text()

        self.tableWidget.clearContents()
        self.tableWidget.setRowCount(0)
        self.tableWidget.setColumnCount(0)

        if df.empty:
            return

        colunas = list(df.columns)
        self.tableWidget.setColumnCount(len(colunas))
        self.tableWidget.setHorizontalHeaderLabels(colunas)
        self.tableWidget.setRowCount(len(df))

        fonte = self.tableWidget.font()
        fonte.setPointSize(8)
        self.tableWidget.setFont(fonte)

        for row_idx in range(len(df)):
            for col_idx, col in enumerate(colunas):
                valor = df.iloc[row_idx, col_idx]
                item = QTableWidgetItem(str(valor))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tableWidget.setItem(row_idx, col_idx, item)

        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.resizeRowsToContents()
        self.tableWidget.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        self.df_exibido = df.copy()
        self.atualizar_labels_resumo()

        if texto_filtro_atual:
            self.apply_global_filter(texto_filtro_atual)

    def atualizar_labels_resumo(self):
        total_ativos = len(self.df_ativos)
        total_desligados = len(self.df_desligados)

        # Usar a data de corte para desligamentos
        data_selecionada = self.date_fim_picker.date().toString("dd/MM/yyyy")
        data_referencia = pd.to_datetime(data_selecionada, format="%d/%m/%Y").normalize()
        dias = int(self.dias_desligados.text())
        limite_inferior = data_referencia - pd.Timedelta(days=dias)

        try:
            datas_adm = pd.to_datetime(self.df_ativos["Data Admissão"], format="%d/%m/%Y", errors="coerce")
            admitidos_ultimo_mes = datas_adm.between(limite_inferior, data_referencia).sum()
        except Exception:
            admitidos_ultimo_mes = 0

        self.label_resumo.setText(
            f"<b>Total de Colaboradores Ativos:</b> {total_ativos} &nbsp;&nbsp;&nbsp;&nbsp; "
            f"<b>Colaboradores Desligados no Intervalo:</b> {total_desligados} &nbsp;&nbsp;&nbsp;&nbsp; "
            f"<b>Admitidos no Intervalo:</b> {admitidos_ultimo_mes}"
        )

    def on_tab_changed(self, index):
        if index == 1:  # Aba de Setores
            self.painel_setores.atualizar_tabela_setores()
        elif index == 2:  # Aba de Gestores
            PainelGestores.atualizar_tabela_gestores(
                self.df_ativos, self.gestores_table, self.label_info_gestores
            )

        # Reaplicar o filtro global ao mudar de aba
        texto_filtro = self.search_field.text()
        self.apply_global_filter(texto_filtro)

    def carregar_dicionario_locais_com_busca(self):
        """Carrega LOCAIS.xlsx com base em frequencia.json e define função de busca."""

        self.dicionario_locais = {}
        self.folder_path = ""

        config_path = os.path.join(os.path.dirname(__file__), "frequencia.json")
        if os.path.exists(config_path):
            with open(config_path, "r") as f:
                config = json.load(f)
                self.folder_path = config.get("folder_path", "")

        locais_path = os.path.join(self.folder_path, "LOCAIS.xlsx")
        if os.path.exists(locais_path):
            locais_df = pd.read_excel(locais_path, header=None, dtype=str)
            self.dicionario_locais = dict(zip(locais_df[0].str.strip(), locais_df[1].str.strip()))
        else:
            QMessageBox.warning(self, "Arquivo não encontrado",
                                f"⚠️ O arquivo 'LOCAIS.xlsx' não foi encontrado no caminho:\n{locais_path}")

        # Função interna de busca
        def buscar(codloc):
            codloc = str(codloc).strip()
            local, setor_final = "Não encontrado", "Não encontrado"
            if not codloc:
                return local, setor_final
            while codloc:
                for codloc_planilha, descricao in self.dicionario_locais.items():
                    if codloc_planilha.startswith(codloc):
                        partes = descricao.split(",", 1)
                        local = partes[0].strip()
                        setor_final = codloc_planilha.rsplit(",", 1)[-1].strip() if "," in codloc_planilha else local
                        return local, setor_final
                if "." in codloc:
                    codloc = ".".join(codloc.split(".")[:-1])
                else:
                    break
            return local, setor_final

        self.buscar_local_setor = buscar

    def apply_global_filter(self, text):
        terms = [term.strip() for term in text.split('/') if term.strip()]

        # Determinar qual tabela está ativa
        current_tab_index = self.tab_widget.currentIndex()
        if current_tab_index == 0:  # Aba de Funcionários
            table_widget = self.tableWidget
        elif current_tab_index == 1:  # Aba de Setores
            table_widget = self.setores_table
        elif current_tab_index == 2:  # Aba de Gestores
            table_widget = self.gestores_table
        else:
            return  # Evita erros caso existam outras abas no futuro
            
        # Se não houver termos, exibe todas as linhas
        if not terms:
            for row in range(table_widget.rowCount()):
                table_widget.setRowHidden(row, False)
            return
    
        headers = {}
        for col in range(table_widget.columnCount()):
            header_item = table_widget.horizontalHeaderItem(col)
            if header_item:
                headers[header_item.text().strip().lower()] = col
    
        for row in range(table_widget.rowCount()):
            row_matches_all_terms = True
            for term in terms:
                term_match = False
                if ":" in term:
                    coluna_nome, valor_busca = map(str.strip, term.split(":", 1))
                    coluna_nome = coluna_nome.lower()
                    if coluna_nome in headers:
                        col_idx = headers[coluna_nome]
                        item = table_widget.item(row, col_idx)
                        if item and item.text().strip().lower() == valor_busca.lower():
                            term_match = True
                else:
                    for col in range(table_widget.columnCount()):
                        item = table_widget.item(row, col)
                        if item and term.lower() in item.text().strip().lower():
                            term_match = True
                            break
                if not term_match:
                    row_matches_all_terms = False
                    break
            table_widget.setRowHidden(row, not row_matches_all_terms)

    def carregar_nomes_filtrados(self):
        """Carrega a lista de nomes a serem filtrados do arquivo no mesmo diretório do script."""
        self.nomes_filtrados = []
        try:
            con_path = os.path.join(os.path.dirname(__file__), "nomes_filtrados.txt")

            if os.path.exists(con_path):
                with open(con_path, "r", encoding="utf-8") as f:
                    self.nomes_filtrados = [linha.strip() for linha in f.readlines() if linha.strip()]
            else:
                # Criar com padrão inicial
                self.nomes_filtrados = ["Denilce Jung", "Felipe Mello", "Robinson Dresch", "Clomar Francisco Milani"]
                self.salvar_nomes_filtrados()
        except Exception as e:
            QMessageBox.warning(self, "Erro ao carregar nomes filtrados",
                                f"Não foi possível carregar o arquivo de nomes filtrados: {str(e)}")

    def salvar_nomes_filtrados(self):
        try:
            con_path = os.path.join(os.path.dirname(__file__), "nomes_filtrados.txt")
            with open(con_path, "w", encoding="utf-8") as f:
                for nome in self.nomes_filtrados:
                    f.write(f"{nome}\n")
        except Exception as e:
            QMessageBox.warning(self, "Erro ao salvar nomes filtrados",
                                f"Não foi possível salvar o arquivo: {str(e)}")

    def export_to_excel(self):
        """Exporta os dados da aba atualmente selecionada para um arquivo Excel."""
        file_path, _ = QFileDialog.getSaveFileName(self, "Salvar Arquivo Excel", "", "Arquivos Excel (*.xlsx)")
        if file_path:
            try:
                # Descobre qual aba está ativa
                aba_index = self.tab_widget.currentIndex()
                if aba_index == 0:
                    tabela = self.tableWidget
                    nome_planilha = "Funcionarios"
                elif aba_index == 1:
                    tabela = self.setores_table
                    nome_planilha = "Setores"
                elif aba_index == 2:
                    tabela = self.gestores_table
                    nome_planilha = "Gestores"
                else:
                    QMessageBox.warning(self, "Erro", "Aba não reconhecida.")
                    return

                if tabela.rowCount() == 0 or tabela.columnCount() == 0:
                    QMessageBox.warning(self, "Atenção", "Tabela vazia. Nada para exportar.")
                    return

                data = []
                for row in range(tabela.rowCount()):
                    row_data = []
                    for col in range(tabela.columnCount()):
                        item = tabela.item(row, col)
                        row_data.append(item.text() if item else "")
                    data.append(row_data)

                headers = [tabela.horizontalHeaderItem(i).text() for i in range(tabela.columnCount())]

                df = pd.DataFrame(data, columns=headers)

                df.to_excel(file_path, index=False, sheet_name=nome_planilha)

                QMessageBox.information(self, "Sucesso", f"Dados exportados para: {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao exportar para Excel: {e}")

    def voltar_menu(self):
        from main import ControlesRH
        self.menu_principal = ControlesRH()
        self.menu_principal.show()
        self.window().close()


    def configurar_filtros(self):
        """Abre janela para configurar a lista de nomes filtrados."""
        dialog = JanelaFiltrosNomes(self.nomes_filtrados, self)
        if dialog.exec():
            self.nomes_filtrados = dialog.nomes_filtrados
            self.salvar_nomes_filtrados()
            self.consultar_funcionarios()  # Atualiza a consulta com os novos filtros

    def abrir_organograma(self):
        if self.df_ativos.empty:
            QMessageBox.information(self, "Organograma", "Nenhum dado disponível.")
            return

        dialog = JanelaOrganograma(self.df_ativos, self)
        dialog.exec()

    def abrir_graficos(self):
        if self.tab_widget.currentIndex() != 1:
            QMessageBox.information(self, "Atenção", "Vá até a aba 'Setores' para abrir o gráfico.")
            return

        tabela = self.setores_table
        if tabela.rowCount() == 0 or tabela.columnCount() == 0:
            QMessageBox.warning(self, "Tabela vazia", "Não há dados para exibir o gráfico.")
            return

        colunas = [tabela.horizontalHeaderItem(i).text() for i in range(tabela.columnCount())]

        # Remover linhas onde 'Setor' contém 'Todos'
        linhas_validas = []
        for row in range(tabela.rowCount()):
            item_setor = tabela.item(row, colunas.index("Setor"))
            if item_setor and "Todos" in item_setor.text():
                continue
            linha = []
            for col in range(tabela.columnCount()):
                item = tabela.item(row, col)
                linha.append(item.text() if item else "")
            linhas_validas.append(linha)

        df = pd.DataFrame(linhas_validas, columns=colunas)

        # Converter colunas numéricas
        for col in df.columns:
            if "Quantidade" in col or "Admitidos" in col or "Desligados" in col or "Média" in col:
                df[col] = pd.to_numeric(df[col].str.replace(",", ".").str.replace("%", ""), errors="coerce")

        data_ref = self.date_fim_picker.date().toString("dd/MM/yyyy")
        self.grafico_window = GraficoSetoresApp(df, data_ref)
        self.grafico_window.show()


class JanelaFiltrosNomes(QDialog):
    def __init__(self, nomes_filtrados, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Configurar Nomes Filtrados")
        self.resize(500, 400)
        self.nomes_filtrados = nomes_filtrados.copy()
        
        layout = QVBoxLayout(self)
        
        # Instruções
        label_instrucoes = QLabel("Adicione, edite ou remova nomes que devem ser filtrados da consulta.\nCada nome deve estar em uma linha separada.")
        layout.addWidget(label_instrucoes)
        
        # Editor de texto para os nomes
        self.text_edit = QTextEdit()
        self.text_edit.setPlainText("\n".join(self.nomes_filtrados))
        layout.addWidget(self.text_edit)
        
        # Botões
        botoes_layout = QHBoxLayout()
        
        btn_cancelar = QPushButton("Cancelar")
        btn_cancelar.clicked.connect(self.reject)
        botoes_layout.addWidget(btn_cancelar)
        
        btn_salvar = QPushButton("Salvar")
        btn_salvar.clicked.connect(self.aceitar)
        botoes_layout.addWidget(btn_salvar)
        
        layout.addLayout(botoes_layout)
        
    def aceitar(self):
        # Atualiza a lista de nomes filtrados
        texto = self.text_edit.toPlainText()
        self.nomes_filtrados = [linha.strip() for linha in texto.split('\n') if linha.strip()]
        self.accept()



class JanelaOrganograma(QDialog):
    def __init__(self, df_ativos, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Gerar Organograma por Local")
        self.resize(400, 100)
        self.df_ativos = df_ativos.copy()

        layout = QVBoxLayout(self)

        # Agrupamento desejado
        agrupamentos_personalizados = {
            "Comercial": "Adm e Comercial",
            "Adm": "Adm e Comercial",
            "Fabrica De Máquinas": "Fabrica De Máquinas",
            "Fabrica De Transportadores": "Fabrica De Transportadores"
        }

        # Criar coluna agrupada
        self.df_ativos["Local Agrupado"] = self.df_ativos["Local"].apply(
            lambda x: agrupamentos_personalizados.get(x.strip(), x.strip())
        )

        # Adiciona agrupamento "Todos os Locais"
        self.df_ativos["Local Agrupado"] = self.df_ativos["Local Agrupado"].fillna("")
        locais = sorted(set(self.df_ativos["Local Agrupado"]) - {""})
        locais.insert(0, "Todos os Locais")  # opção agrupada geral

        self.combo_local = QComboBox()
        self.combo_local.addItems(locais)
        layout.addWidget(QLabel("Selecione o Local:"))
        layout.addWidget(self.combo_local)

        btn_gerar = QPushButton("Gerar Organograma")
        btn_gerar.clicked.connect(self.tratar_gerar_organograma)
        layout.addWidget(btn_gerar)

    def tratar_gerar_organograma(self):
        local = self.combo_local.currentText()

        if local == "Todos os Locais":
            df = self.df_ativos[self.df_ativos["Local Agrupado"].notna()].copy()
            df["Local Agrupado"] = "Todos os Locais"
            self.df_filtrados = df
        else:
            self.df_filtrados = self.df_ativos[self.df_ativos["Local Agrupado"] == local].copy()

        self.gerar_organograma()

    def gerar_organograma(self):
        def gerar_id(nome):
            return nome.strip().lower().replace(" ", "_").replace(".", "").replace(",", "").replace("-", "")

        def destacar_parenteses(texto, cor="#ff0000"):
            return re.sub(r"\(([^)]+)\)", rf"(<FONT COLOR='{cor}'>\1</FONT>)", texto)

        def gerar_cor_hex():
            while True:
                r, g, b = [random.choice([0, 128, 255]) for _ in range(3)]

                if abs(r - g) < 60 and abs(g - b) < 60: continue  # cinzas
                if sum(x >= 200 for x in [r, g, b]) >= 2: continue  # tons claros
                if 255 in [r, g, b] and sum(x <= 50 for x in [r, g, b]) >= 2: continue  # neon puro
                if g >= 200 and (r + b) < 300: continue  # verde claro ou neon

                h, l, s = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
                if l > 0.45: continue  # muito claro

                if (r + g + b) / 3 > 200: continue  # cor média muito clara

                return "#{:02x}{:02x}{:02x}".format(r, g, b)

        local = self.combo_local.currentText()
        df = self.df_filtrados.copy()

        if df.empty:
            QMessageBox.information(self, "Organograma", "Nenhum dado para este local.")
            return

        dot = Digraph(engine='dot', comment='Organograma', format='pdf')
        dot.attr(rankdir='TB', nodesep='2.0', ranksep='2.0', fontsize='12')
        dot.attr(bgcolor="#eeeeee")  # fundo cinza claro
        titulo_html = f"""<<FONT POINT-SIZE="24" COLOR="red"><B>Organograma - {local}</B></FONT>>"""
        dot.attr(label=titulo_html, labelloc="t", labeljust="c", fontname="Arial")

        gestores = df["Gestor"].dropna().unique()
        cores_gestores = {gestor: gerar_cor_hex() for gestor in gestores}
        gestores_blocos = {}

        for gestor in gestores:
            setor = df.loc[df["Colaborador"] == gestor, "Setor"].values
            if len(setor) > 0 and setor[0].strip():
                setor = setor[0]
            else:
                setor_global = self.df_ativos.loc[self.df_ativos["Colaborador"] == gestor, "Setor"].values
                setor = setor_global[0] if len(setor_global) > 0 else "Setor Desconhecido"

            bloco_nome = f"{setor} ({gestor})"
            bloco_id = gerar_id(bloco_nome)
            gestores_blocos[gestor] = {"id": bloco_id, "nome": bloco_nome, "colaborador": gestor}

            cor_gestor = cores_gestores.get(gestor, "#000000")
            bloco_nome_formatado = destacar_parenteses(bloco_nome, cor=cor_gestor)
            label_html = f"""<<TABLE BORDER="0" CELLBORDER="1" CELLSPACING="0" COLOR="black">
            <TR><TD BGCOLOR="white"><B>{bloco_nome_formatado}</B></TD></TR>
            </TABLE>>"""

            dot.node(bloco_id, label=label_html, shape="plaintext")

        for gestor, dados in gestores_blocos.items():
            gestor_superior = df.loc[df["Colaborador"] == gestor, "Gestor"].values
            if len(gestor_superior) > 0 and gestor_superior[0] in gestores_blocos:
                origem = gestores_blocos[gestor_superior[0]]["id"]
                destino = dados["id"]
                if origem != destino:
                    cor = cores_gestores.get(gestor_superior[0], "#000000")
                    dot.edge(origem, destino, color=cor)

        colaboradores_df = df[~df["Colaborador"].isin(gestores)]
        blocos_colaboradores = {}

        for (_, linha) in colaboradores_df.iterrows():
            colaborador = linha["Colaborador"].strip()
            setor = linha["Setor"].strip()
            gestor = linha["Gestor"].strip()
            if not gestor:
                continue
            bloco_nome = f"{setor} ({gestor})"
            bloco_id = gerar_id(bloco_nome)

            blocos_colaboradores.setdefault(bloco_id, {
                "nome": bloco_nome,
                "colaboradores": [],
                "gestor": gestor
            })
            blocos_colaboradores[bloco_id]["colaboradores"].append(colaborador)

        for bloco_id, dados in blocos_colaboradores.items():
            lista_colab = ""
            for col in sorted(dados["colaboradores"]):
                col_formatado = destacar_parenteses(col)
                lista_colab += f"<TR><TD ALIGN='LEFT'>• {col_formatado}</TD></TR>"

            cor_gestor = cores_gestores.get(dados["gestor"], "#000000")
            nome_formatado = destacar_parenteses(dados["nome"], cor=cor_gestor)
            setor_label = f"""<<TABLE BORDER="0" CELLBORDER="1" CELLSPACING="0" COLOR="black">
                <TR><TD BGCOLOR="white"><B>{nome_formatado}</B></TD></TR>
                {lista_colab}
            </TABLE>>"""

            dot.node(bloco_id, label=setor_label, shape="plaintext")

            if dados["gestor"] in gestores_blocos:
                gestor_id = gestores_blocos[dados["gestor"]]["id"]
                if gestor_id != bloco_id:
                    cor = cores_gestores.get(dados["gestor"], "#000000")
                    dot.edge(gestor_id, bloco_id, color=cor)

        # Selecionar pasta para salvar
        pasta = QFileDialog.getExistingDirectory(self, "Selecionar pasta para salvar o organograma")
        if not pasta:
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Organograma_{local}_{timestamp}"
        filepath = os.path.join(pasta, filename)

        output_pdf = dot.render(filepath, format="pdf", cleanup=True)
        self.adicionar_rodape_pdf(output_pdf, texto="By: Silonei Duarte")
        os.startfile(output_pdf)
        self.accept()

    def adicionar_rodape_pdf(self, path_pdf, texto="By: Silonei Duarte"):

        doc = fitz.open(path_pdf)
        page = doc[-1]
        largura = page.rect.width
        altura = page.rect.height

        page.insert_text(
            point=(largura - 120, altura - 20),
            text=texto,
            fontsize=8,
            fontname="helv",
            color=(0, 0, 0)
        )

        # Salvar em arquivo temporário
        temp_path = path_pdf.replace(".pdf", "_temp.pdf")
        doc.save(temp_path)

        # Substituir o original
        doc.close()
        os.replace(temp_path, path_pdf)



if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = PainelConsultaFuncionarios()
    window.show()
    sys.exit(app.exec())