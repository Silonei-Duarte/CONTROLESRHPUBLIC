import sys
from unidecode import unidecode
from PyQt6.QtWidgets import (
    QApplication, QWidget, QHBoxLayout, QPushButton, QLineEdit, QMessageBox, QDateEdit, QHeaderView
)
from PyQt6.QtWidgets import QListWidget, QListWidgetItem
import pandas as pd
from PyQt6.QtCore import QDate
from datetime import datetime, timedelta
import win32com.client
import os
from contextlib import contextmanager
import oracledb
import json
from decimal import Decimal, InvalidOperation
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QTableWidget, QTableWidgetItem, QLabel
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon, QFont
from afastamentosgrafico import GraficoAfastamento
from Database import get_connection

class Appatestados(QWidget):
    def __init__(self):
        super().__init__()
        self.df_original = None
        self.df_completo = None
        self.carregar_dicionario_locais_com_busca()

        self.initUI()

    def initUI(self):
        self.setWindowTitle("Consulta de Afastamentos")
        self.resize(1900, 900)

        # Configurar o ícone da janela
        icon_path = os.path.join(os.path.dirname(__file__), 'icone.ico')  # Ajuste o caminho conforme necessário
        if os.path.exists(icon_path):  # Verifica se o arquivo existe
            self.setWindowIcon(QIcon(icon_path))

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        # Caixa lateral com as situações
        self.lista_sitafas = QListWidget()
        self.lista_sitafas.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.lista_sitafas.setFixedWidth(200)
        self.lista_sitafas.setFixedHeight(180)

        self.sitafas = {
            2: "Férias", 3: "Auxílio Doença", 4: "Acidente de Trabalho",
            6: "Licença Maternidade", 11: "Licença Paternidade",
            14: "Atestado Médico", 20: "Licença médica 15 dias",
            24: "Pensão Vitalícia", 25: "Falecimento Familiar",
            52: "Férias Noturnas", 53: "Auxílio Doença Noturno",
            54: "Acidente de Trabalho Noturno", 56: "Licença Maternidade Noturna",
            61: "Licença Paternidade Noturna", 64: "Atestado Médico Noturno",
            72: "Licença Casamento", 73: "Doação de Sangue", 201: "Saída Médico", 202: "Saída Médico Noturna",
            74: "Justiça Eleitoral", 75: "Licença Amamentação", 913: "Aposentadoria por Invalidez",
            912: "Suspensão Disciplinar",
            918: "Suspensão Disciplinar", 979: "Suspensão Disciplinar Noturna"
        }

        for cod, nome in self.sitafas.items():
            item = QListWidgetItem(nome)
            item.setData(Qt.ItemDataRole.UserRole, cod)
            self.lista_sitafas.addItem(item)
            if cod in (14, 64):
                item.setSelected(True)

        # Bloco de botões de consulta
        blocos_layout = QVBoxLayout()

        linha_correntes = QHBoxLayout()
        linha_correntes.setAlignment(Qt.AlignmentFlag.AlignLeft)
        linha_correntes.addWidget(QLabel("🔹 Afastamentos Correntes"))
        btn_correntes = QPushButton("Consultar")
        btn_correntes.setFixedSize(120, 30)
        btn_correntes.clicked.connect(self.consultar_correntes)
        linha_correntes.addWidget(btn_correntes)
        blocos_layout.addLayout(linha_correntes)

        linha_iniciados = QHBoxLayout()
        linha_iniciados.setAlignment(Qt.AlignmentFlag.AlignLeft)
        linha_iniciados.addWidget(QLabel("🔹 Afastamentos Iniciados Desde:"))
        self.data_iniciados = QDateEdit()
        self.data_iniciados.setCalendarPopup(True)
        self.data_iniciados.setDisplayFormat("dd/MM/yyyy")
        self.data_iniciados.setDate(QDate.currentDate().addMonths(-2))
        self.data_iniciados.setFixedSize(120, 30)
        btn_iniciados = QPushButton("Consultar")
        btn_iniciados.setFixedSize(120, 30)
        btn_iniciados.clicked.connect(self.consultar_iniciados)
        linha_iniciados.addWidget(self.data_iniciados)
        linha_iniciados.addWidget(btn_iniciados)
        blocos_layout.addLayout(linha_iniciados)

        linha_periodo = QHBoxLayout()
        linha_periodo.setAlignment(Qt.AlignmentFlag.AlignLeft)
        linha_periodo.addWidget(QLabel("🔹 Afastamentos Iniciados em:"))
        self.data_inicio_periodo = QDateEdit()
        self.data_inicio_periodo.setCalendarPopup(True)
        self.data_inicio_periodo.setDisplayFormat("dd/MM/yyyy")
        self.data_inicio_periodo.setDate(QDate.currentDate().addMonths(-1))
        self.data_inicio_periodo.setFixedSize(120, 30)

        self.data_fim_periodo = QDateEdit()
        self.data_fim_periodo.setCalendarPopup(True)
        self.data_fim_periodo.setDisplayFormat("dd/MM/yyyy")
        self.data_fim_periodo.setDate(QDate.currentDate())
        self.data_fim_periodo.setFixedSize(120, 30)

        linha_periodo.addWidget(self.data_inicio_periodo)
        linha_periodo.addWidget(QLabel("Até iniciados em:"))
        linha_periodo.addWidget(self.data_fim_periodo)

        btn_periodo = QPushButton("Consultar")
        btn_periodo.setFixedSize(120, 30)
        btn_periodo.clicked.connect(self.consultar_periodo)
        linha_periodo.addWidget(btn_periodo)
        blocos_layout.addLayout(linha_periodo)

        # Bloco de botões extras (à direita dos botões de consulta)
        botoes_extras_layout = QVBoxLayout()
        botoes_extras_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        btn_grafico = QPushButton("Abrir Gráfico")
        btn_grafico.setFixedSize(120, 30)
        btn_grafico.clicked.connect(self.abrir_graficos)
        botoes_extras_layout.addWidget(btn_grafico)

        btn_email = QPushButton("E-mail")
        btn_email.setFixedSize(120, 30)
        btn_email.clicked.connect(self.enviar_email)
        botoes_extras_layout.addWidget(btn_email)

        btn_custo = QPushButton("Custo Aproximado")
        btn_custo.setFixedSize(120, 30)
        btn_custo.clicked.connect(self.exibir_custos_afastamentos)
        botoes_extras_layout.addWidget(btn_custo)

        self.btn_voltar = QPushButton("Voltar ao Menu")
        self.btn_voltar.setFixedSize(120, 30)
        self.btn_voltar.clicked.connect(self.voltar_menu)
        botoes_extras_layout.addWidget(self.btn_voltar)

        # Agrupar filtros (lista lateral, botões e extras)
        filtros_layout = QHBoxLayout()
        filtros_layout.addWidget(self.lista_sitafas)

        layout_botoes_total = QHBoxLayout()
        layout_botoes_total.addLayout(blocos_layout)
        layout_botoes_total.addSpacing(50)
        layout_botoes_total.addLayout(botoes_extras_layout)

        filtros_layout.addLayout(layout_botoes_total)
        layout.addLayout(filtros_layout)

        # Label status
        self.status_label = QLabel("Nenhuma consulta realizada ainda.")
        self.status_label.setWordWrap(True)
        self.status_label.setMaximumWidth(1800)
        self.status_label.setStyleSheet("color: red; font-weight: bold;")
        fonte = QFont()
        fonte.setPointSize(10)
        fonte.setFamily("Arial")
        self.status_label.setFont(fonte)
        layout.addWidget(self.status_label)

        self.label_instrucao_detalhes = QLabel(
            "Resultados agrupados por colaborador. "
            "Dê dois cliques em uma linha para visualizar todos os afastamentos do colaborador dentro da consulta realizada."
        )
        self.label_instrucao_detalhes.setStyleSheet("color: red;")
        fonte_instrucao = QFont("Arial", 8)
        self.label_instrucao_detalhes.setFont(fonte_instrucao)
        self.label_instrucao_detalhes.setWordWrap(False)
        layout.addWidget(self.label_instrucao_detalhes)

        # Campo de filtro
        self.search_field = QLineEdit()
        self.search_field.setPlaceholderText("Filtrar usando / entre dados. Também pode digitar COLUNA:DADO")
        self.search_field.textChanged.connect(self.apply_global_filter)
        layout.addWidget(self.search_field)

        # Tabela
        self.tableWidget = QTableWidget()
        layout.addWidget(self.tableWidget)

        self.setLayout(layout)
        self.tableWidget.cellDoubleClicked.connect(self.carregar_detalhes)

        self.current_mail = None

    def get_sitafas_selecionadas(self):
        codigos = []
        nomes = []
        for item in self.lista_sitafas.selectedItems():
            cod = item.data(Qt.ItemDataRole.UserRole)
            nome = item.text()  # Agora já é só o nome
            codigos.append(str(cod))
            nomes.append(nome)
        return codigos, nomes

    def consultar_correntes(self):
        sitafa, _ = self.get_sitafas_selecionadas()
        if not sitafa:
            QMessageBox.warning(self, "Aviso", "Selecione ao menos uma Situação de Afastamento.")
            return
        query = f"""
        SELECT 
            A.NUMEMP, 
            H.CODLOC, 
            A.NUMCAD, 
            F.NOMFUN, 
            S.DESSIT,        
            A.DATAFA, 
            A.DATTER,
            A.CODDOE, 
            D.DESDOE, 
            A.NOMATE, 
            A.OBSAFA
        FROM R038AFA A
        JOIN R034FUN F ON A.NUMEMP = F.NUMEMP AND A.NUMCAD = F.NUMCAD
        JOIN R192DOE D ON A.CODDOE = D.CODDOE
        JOIN R016HIE H ON F.NUMLOC = H.NUMLOC
        JOIN R010SIT S ON A.SITAFA = S.CODSIT
        WHERE A.NUMEMP IN (10, 16, 17, 18, 19)
          AND F.TIPCOL = 1
          AND A.SITAFA IN ({",".join(sitafa)})
        AND A.DATAFA <= TRUNC(SYSDATE)
        AND A.DATTER >= TRUNC(SYSDATE)
        """
        self.carregar_dados_sql(query, "Afastamentos Correntes")

    def consultar_iniciados(self):
        sitafa, _ = self.get_sitafas_selecionadas()
        if not sitafa:
            QMessageBox.warning(self, "Aviso", "Selecione ao menos uma Situação de Afastamento.")
            return
        data = self.data_iniciados.date().toString("dd/MM/yyyy")
        query = f"""
        SELECT 
            A.NUMEMP, 
            H.CODLOC, 
            A.NUMCAD, 
            F.NOMFUN, 
            S.DESSIT,        
            A.DATAFA, 
            A.DATTER,
            A.CODDOE, 
            D.DESDOE, 
            A.NOMATE, 
            A.OBSAFA
        FROM R038AFA A
        JOIN R034FUN F ON A.NUMEMP = F.NUMEMP AND A.NUMCAD = F.NUMCAD
        JOIN R192DOE D ON A.CODDOE = D.CODDOE
        JOIN R016HIE H ON F.NUMLOC = H.NUMLOC
        JOIN R010SIT S ON A.SITAFA = S.CODSIT
        WHERE A.NUMEMP IN (10, 16, 17, 18, 19)
          AND F.TIPCOL = 1
          AND A.SITAFA IN ({",".join(sitafa)})
          AND A.DATAFA >= TO_DATE('{data}', 'DD/MM/YYYY')
        """
        self.carregar_dados_sql(query, f"Afastamentos Iniciados a partir de {data}")

    # Realiza consulta dos iniciados em uma data até iniciados em outra data
    def consultar_periodo(self):
        sitafa, _ = self.get_sitafas_selecionadas()
        if not sitafa:
            QMessageBox.warning(self, "Aviso", "Selecione ao menos uma Situação de Afastamento.")
            return
        data_ini = self.data_inicio_periodo.date().toString("dd/MM/yyyy")
        data_fim = self.data_fim_periodo.date().toString("dd/MM/yyyy")
        query = f"""
        SELECT 
            A.NUMEMP, 
            H.CODLOC, 
            A.NUMCAD, 
            F.NOMFUN, 
            S.DESSIT,        
            A.DATAFA, 
            A.DATTER,
            A.CODDOE, 
            D.DESDOE, 
            A.NOMATE, 
            A.OBSAFA
        FROM R038AFA A
        JOIN R034FUN F ON A.NUMEMP = F.NUMEMP AND A.NUMCAD = F.NUMCAD
        JOIN R192DOE D ON A.CODDOE = D.CODDOE
        JOIN R016HIE H ON F.NUMLOC = H.NUMLOC
        JOIN R010SIT S ON A.SITAFA = S.CODSIT
        WHERE A.NUMEMP IN (10, 16, 17, 18, 19)
          AND F.TIPCOL = 1
          AND A.SITAFA IN ({",".join(sitafa)})
        AND A.DATAFA BETWEEN TO_DATE('{data_ini}', 'DD/MM/YYYY') AND TO_DATE('{data_fim}', 'DD/MM/YYYY')

        """
        self.carregar_dados_sql(query, f"Afastamentos Iníciados em  {data_ini} até Iniciados em {data_fim}")

    def carregar_dados_sql(self, query_base, status_titulo):
        sitafas, nomes = self.get_sitafas_selecionadas()
        if not sitafas:
            QMessageBox.warning(self, "Atenção", "Selecione ao menos uma Situação (SITAFA) para consultar.")
            return

        sitafa_str = ", ".join(sitafas)
        status_texto = f"{status_titulo} | Situações: {', '.join(nomes)}"

        query = query_base.replace("{SITAFA_LIST}", sitafa_str)
        self.status_label.setText(status_texto)

        with get_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(query)
                colunas = [desc[0] for desc in cursor.description]
                registros = [dict(zip(colunas, row)) for row in cursor.fetchall()]

        df = pd.DataFrame(registros)

        if df.empty:
            self.df_completo = pd.DataFrame()
            self.df_original = pd.DataFrame()
            self.tableWidget.setRowCount(0)
            self.tableWidget.setColumnCount(0)
            self.status_label.setText(f"{status_titulo} | Nenhum dado encontrado para: {', '.join(nomes)}")
            return

        # **Limpeza dos dados** (remove espaços extras, quebras de linha e trata valores vazios)
        df = df.map(lambda x: x.replace("\n", " ").replace("\r", " ").strip() if isinstance(x, str) else x)
        df.fillna("", inplace=True)  # Substitui valores nulos por string vazia

        # Renomear colunas para exibição
        df.rename(columns={
            "NUMEMP": "Local",
            "CODLOC": "Setor",
            "NUMCAD": "Cadastro",
            "NOMFUN": "Colaborador",
            "DESSIT": "Situação",
            "DATAFA": "Data Afastamento",
            "DATTER": "Data Fim",
            "CODDOE": "CID",
            "DESDOE": "CID Descrição",
            "NOMATE": "Médico",
            "OBSAFA": "Observação"
        }, inplace=True)

        # Converter datas corretamente
        df["Data Afastamento"] = pd.to_datetime(df["Data Afastamento"], format="%d/%m/%Y", errors='coerce')
        df["Data Fim"] = pd.to_datetime(df["Data Fim"], format="%d/%m/%Y", errors='coerce')

        # Remover registros onde Data Fim (DATTER) < Data Afastamento (DATAFA)
        df = df[df["Data Fim"] >= df["Data Afastamento"]]

        # Calcular Dias Ult. Afastamento
        df["Dias Ult. Afastamento"] = (df["Data Fim"] - df["Data Afastamento"]).dt.days + 1

        # **Reorganizar colunas para manter a ordem desejada**
        colunas_ordenadas = [
            "Local", "Setor", "Cadastro", "Colaborador", "Situação",
            "Data Afastamento", "Data Fim", "Dias Ult. Afastamento",  # Dias Ult. Afastamento logo após Data Fim
            "CID", "CID Descrição", "Médico", "Observação"
        ]

        self.df_para_salario = df.copy()  # ← salvar antes de traduzir Local/Setor

        # Substituir Local e Setor usando o dicionário
        df[["Local", "Setor"]] = df["Setor"].apply(
            lambda codloc: pd.Series(self.buscar_local_setor(codloc))
        )

        # Aplicando a nova ordem no DataFrame
        df = df[colunas_ordenadas]

        # **Salvar df_completo SEM colunas agregadas** (dados puros do banco)
        self.df_completo = df.copy()


        # Criar um DataFrame SEPARADO para o agrupamento
        df_para_agrupamento = df.copy()

        # Criar colunas de contagem e soma para agrupamento
        df_para_agrupamento["Qtd. Afastamentos"] = 1  # Cada linha é um afastamento individual
        df_para_agrupamento["Qtd. Total Dias Afastados"] = df_para_agrupamento["Dias Ult. Afastamento"]  # Cópia da coluna Dias Ult. Afastamento
        df_para_agrupamento["Dias Corridos"] = 0  # Inicializa a coluna antes do agrupamento
        df_para_agrupamento["Dias para INSS"] = 0

        df_para_agrupamento = df_para_agrupamento.sort_values(by="Data Fim", ascending=False, na_position='last').reset_index(drop=True)

        # Agrupar por Local, Cadastro e Situação
        agrupado = df_para_agrupamento.groupby(["Local", "Cadastro", "Situação"], as_index=False).agg({
            "Setor": "first",
            "Colaborador": "first",
            "Qtd. Afastamentos": "sum",  # Contagem de afastamentos por colaborador
            "Qtd. Total Dias Afastados": "sum",  # Soma total de dias afastados
            "Dias Corridos": "max",  # Maior sequência de dias afastados consecutivos
            "Data Afastamento": "first",  # Primeira data de afastamento
            "Data Fim": "first",  # Última data de retorno
            "Dias Ult. Afastamento": "first",  # Soma total de dias afastados
            "CID": "first",
            "Dias para INSS": "first",
            "CID Descrição": "first",
            "Médico": "first",
            "Observação": "first"
        })

        # **Reorganizar colunas para manter a ordem desejada**
        colunas_ordenadas = [
            "Local", "Setor", "Cadastro", "Colaborador", "Situação", "Qtd. Afastamentos",
            "Qtd. Total Dias Afastados",
            "Dias Corridos", "Data Afastamento", "Data Fim", "Dias Ult. Afastamento",
            "CID", "Dias para INSS", "CID Descrição", "Médico", "Observação"
        ]
        agrupado = agrupado[colunas_ordenadas]

        # Aplicar a função para calcular os dias corridos APENAS no DataFrame do agrupamento
        agrupado = self.calcular_dias_corridos(agrupado, self.df_completo)
        agrupado = self.calcular_cid_60dias(agrupado, self.df_completo)

        # Verifica se todas as colunas CID, CID Descrição, Dias para INSS e Médico estão vazias
        colunas_para_verificar = ["CID", "Dias para INSS", "CID Descrição", "Médico"]

        if all(col in agrupado.columns for col in colunas_para_verificar):
            todas_vazias = True
            for col in colunas_para_verificar:
                col_series = agrupado[col]
                # Considera "vazia" se todos os valores forem nulos ou strings vazias
                if col_series.dropna().astype(str).str.strip().ne('').any():
                    todas_vazias = False
                    break

            if todas_vazias:
                agrupado.drop(columns=colunas_para_verificar, inplace=True)

        # **Ordenação correta pela Data Fim**
        agrupado = agrupado.sort_values(by="Data Fim", ascending=False, na_position='last').reset_index(drop=True)

        # Converter datas para string ANTES de atualizar a tabela
        agrupado["Data Afastamento"] = agrupado["Data Afastamento"].dt.strftime("%d/%m/%Y")
        agrupado["Data Fim"] = agrupado["Data Fim"].dt.strftime("%d/%m/%Y")

        # Atualizar df_original para manter os dados iniciais da exibição
        self.df_original = agrupado.copy()

        # **🔹 Preencher a TableWidget corretamente**
        colunas = list(agrupado.columns)
        self.tableWidget.setColumnCount(len(colunas))
        self.tableWidget.setRowCount(len(agrupado))
        self.tableWidget.setHorizontalHeaderLabels(colunas)

        # Define a fonte uma única vez para não repetir em cada célula
        fonte = self.tableWidget.font()
        fonte.setPointSize(8)
        self.tableWidget.setFont(fonte)

        for row_idx, row in agrupado.iterrows():
            pintar_amarelo = False
            try:
                valor_inss = int(row.get("Dias para INSS", 0))
                pintar_amarelo = valor_inss > 15
            except:
                pass

            for col_idx, col in enumerate(colunas):
                item = QTableWidgetItem(str(row[col]) if row[col] is not None else "")
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

                if pintar_amarelo:
                    item.setBackground(Qt.GlobalColor.yellow)

                self.tableWidget.setItem(row_idx, col_idx, item)

        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.resizeRowsToContents()
        self.tableWidget.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        # Oculta a coluna "Dias Corridos" se existir
        coluna_ocultar = "Dias Corridos"
        if coluna_ocultar in colunas:
            idx = colunas.index(coluna_ocultar)
            self.tableWidget.setColumnHidden(idx, True)



    def calcular_dias_corridos(self, df_agrupado, df_completo):
        """
        Passo inicial: Criar listas com todas as datas de afastamento, retorno e dias afastados
        para cada Local / Cadastro, mas apenas se a última Data Fim for >= hoje.
        """

        # Criar nova coluna com valor 0 (para ser atualizado posteriormente)
        df_agrupado["Dias Corridos"] = 0

        # Obter a data de hoje
        data_hoje = datetime.today().date()

        # Percorrer cada grupo e coletar os dados
        for index, row in df_agrupado.iterrows():
            nome_empresa = row["Local"]
            numero_cadastro = row["Cadastro"]

            # Filtrar os registros do colaborador específico
            registros_individuais = df_completo[
                (df_completo["Local"] == nome_empresa) &
                (df_completo["Cadastro"] == numero_cadastro)
                ].sort_values(by="Data Fim", ascending=False)

            # Criar listas para armazenar os dados de afastamento
            lista_data_retorno = []
            lista_data_afastamento = []
            lista_dias_afastamento = []

            # Verificar se o último retorno é maior ou igual a hoje
            if not registros_individuais.empty and registros_individuais.iloc[0]["Data Fim"].date() >= data_hoje:
                for _, registro in registros_individuais.iterrows():
                    lista_data_retorno.append(registro["Data Fim"].date())
                    lista_data_afastamento.append(registro["Data Afastamento"].date())
                    lista_dias_afastamento.append(registro["Dias Ult. Afastamento"])

                # Calcular os dias corridos
                dias_corridos = lista_dias_afastamento[0]  # Começa com o primeiro afastamento

                for i in range(len(lista_data_afastamento) - 1):
                    afastamento_atual = lista_data_afastamento[i]
                    retorno_proximo = lista_data_retorno[i + 1]

                    # Se o afastamento atual não for contínuo ao retorno anterior, parar a contagem
                    if afastamento_atual > retorno_proximo + timedelta(days=1):
                        break
                    else:
                        dias_corridos += lista_dias_afastamento[i + 1]  # Soma os dias afastados apenas se for contínuo


                # Atualizar o DataFrame com o valor calculado
                df_agrupado.at[index, "Dias Corridos"] = dias_corridos

        return df_agrupado

    def calcular_cid_60dias(self, df_agrupado, df_completo):
        """
        Para cada colaborador no df_agrupado, calcula a soma de dias afastados
        para o mesmo CID nos últimos 60 dias no df_completo,
        considerando apenas situações específicas e apenas os dias
        efetivos dentro da janela de 60 dias.
        """
        df_agrupado["Dias para INSS"] = ""  # Inicializa com vazio

        hoje = datetime.today().date()
        data_limite = hoje - timedelta(days=60)

        # Situações válidas para o cálculo
        situacoes_validas = [
            "Auxilio Doenca", "Acidente de Trabalho", "Atestado Medico",
            "Licença médica 15 dias", "Pensão Vitalícia",
            "Auxílio Doenca Noturno", "Acidente de Trabalho Noturno", "Atestado Médico Noturno"
        ]

        for index, row in df_agrupado.iterrows():
            nome_empresa = row["Local"]
            numero_cadastro = row["Cadastro"]

            # Filtrar registros do colaborador com situações válidas
            registros = df_completo[
                (df_completo["Local"] == nome_empresa) &
                (df_completo["Cadastro"] == numero_cadastro) &
                (df_completo["Situação"].isin(situacoes_validas))
                ]

            if registros.empty:
                continue

            # Garantir datas como tipo date
            registros = registros.copy()
            registros["Data Afastamento"] = registros["Data Afastamento"].dt.date
            registros["Data Fim"] = registros["Data Fim"].dt.date

            # Calcular dias dentro da janela de 60 dias
            registros["Dias No Periodo"] = registros.apply(
                lambda r: (
                    (r["Data Fim"] - r["Data Afastamento"]).days + 1
                    if r["Data Afastamento"] >= data_limite
                    else (r["Data Fim"] - data_limite).days
                    if r["Data Fim"] >= data_limite else 0
                ),
                axis=1
            )

            # Remover registros com 0 dias úteis no período
            registros = registros[registros["Dias No Periodo"] > 0]

            if not registros.empty:
                soma_por_cid = registros.groupby("CID")["Dias No Periodo"].sum()
                maior_soma = soma_por_cid.max()
                df_agrupado.at[index, "Dias para INSS"] = int(maior_soma)

        return df_agrupado

    def apply_global_filter(self, text):
        terms = [term.strip() for term in text.split('/') if term.strip()]

        # Se não houver termos, exibe todas as linhas
        if not terms:
            for row in range(self.tableWidget.rowCount()):
                self.tableWidget.setRowHidden(row, False)
            return

        headers = {}
        for col in range(self.tableWidget.columnCount()):
            header_item = self.tableWidget.horizontalHeaderItem(col)
            if header_item:
                headers[header_item.text().strip().lower()] = col

        for row in range(self.tableWidget.rowCount()):
            row_matches_all_terms = True
            for term in terms:
                term_match = False
                if ":" in term:
                    coluna_nome, valor_busca = map(str.strip, term.split(":", 1))
                    coluna_nome = coluna_nome.lower()
                    if coluna_nome in headers:
                        col_idx = headers[coluna_nome]
                        item = self.tableWidget.item(row, col_idx)
                        if item and item.text().strip().lower() == valor_busca.lower():
                            term_match = True
                else:
                    for col in range(self.tableWidget.columnCount()):
                        item = self.tableWidget.item(row, col)
                        if item and term.lower() in item.text().strip().lower():
                            term_match = True
                            break
                if not term_match:
                    row_matches_all_terms = False
                    break
            self.tableWidget.setRowHidden(row, not row_matches_all_terms)

    def carregar_detalhes(self, row, column):
        """Exibe os detalhes dos registros agrupados ao dar dois cliques em uma linha."""

        if self.df_completo.empty:
            return

        nome_empresa = self.tableWidget.item(row, 0).text().strip()
        numero_cadastro = self.tableWidget.item(row, 2).text().strip()  # ← Corrigido para coluna 2 (Cadastro)
        situacao_afastamento = self.tableWidget.item(row, 4).text().strip()

        try:
            numero_cadastro = int(numero_cadastro)
        except ValueError:
            pass

        # 🔍 Verificação robusta: tipo + strip + lower
        registros_filtrados = self.df_completo[
            (self.df_completo["Local"].astype(str).str.strip() == nome_empresa) &
            (self.df_completo["Cadastro"].astype(str) == str(numero_cadastro)) &
            (self.df_completo["Situação"].astype(str).str.strip().str.lower() == situacao_afastamento.lower())
            ]

        if registros_filtrados.empty:
            QMessageBox.information(self, "Nenhum dado", "Nenhum registro encontrado para este agrupamento.")
            return

        registros_filtrados = registros_filtrados.sort_values(by="Data Fim", ascending=False, na_position='last')

        dialog = QDialog(self)
        dialog.setWindowTitle(f"Detalhes - {nome_empresa} ({numero_cadastro}) - {situacao_afastamento}")
        dialog.resize(1600, 600)

        layout = QVBoxLayout()
        layout.addWidget(
            QLabel(f"Registros agrupados para: {nome_empresa} - {numero_cadastro} - {situacao_afastamento}"))

        tabela_detalhes = QTableWidget()
        colunas = list(registros_filtrados.columns)
        tabela_detalhes.setColumnCount(len(colunas))
        tabela_detalhes.setRowCount(len(registros_filtrados))
        tabela_detalhes.setHorizontalHeaderLabels(colunas)

        fonte = tabela_detalhes.font()
        fonte.setPointSize(8)
        tabela_detalhes.setFont(fonte)

        for row_idx, (_, registro) in enumerate(registros_filtrados.iterrows()):
            for col_idx, col in enumerate(colunas):
                valor = registro[col]
                if isinstance(valor, pd.Timestamp):
                    valor = valor.strftime("%d/%m/%Y")
                item = QTableWidgetItem(str(valor) if valor is not None else "")
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                tabela_detalhes.setItem(row_idx, col_idx, item)

        tabela_detalhes.resizeColumnsToContents()
        tabela_detalhes.resizeRowsToContents()

        layout.addWidget(tabela_detalhes)
        dialog.setLayout(layout)
        dialog.exec()

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
        else: QMessageBox.warning(self, "Arquivo não encontrado",
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

    def abrir_graficos(self):
        """Abre a aba de gráficos passando os dados da tabela atual (visível)."""

        # Obter as colunas da tabela
        colunas = [self.tableWidget.horizontalHeaderItem(col).text() for col in range(self.tableWidget.columnCount())]

        # Coletar apenas as linhas visíveis
        dados = []
        for row in range(self.tableWidget.rowCount()):
            if not self.tableWidget.isRowHidden(row):
                linha = [self.tableWidget.item(row, col).text() if self.tableWidget.item(row, col) else '-'
                         for col in range(self.tableWidget.columnCount())]
                dados.append(linha)

        if not dados:
            QMessageBox.warning(self, "Aviso", "Nenhum dado visível. Ajuste os filtros ou realize uma consulta.")
            return

        # Criar DataFrame
        df = pd.DataFrame(dados, columns=colunas)

        # Converter colunas numéricas (ex: "Qtd. Afastamentos") para inteiros
        colunas_qtd = [col for col in df.columns if col.startswith("Qtd.")]
        df[colunas_qtd] = df[colunas_qtd].replace("-", 0).apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)

        # Título: usa o mesmo label exibido na interface
        titulo_grafico = self.status_label.text()

        # Abrir janela de gráfico passando df e título
        self.aba_grafico = GraficoAfastamento(df, titulo_grafico)
        self.aba_grafico.show()

        # Se um e-mail já foi gerado, passar para a aba de gráficos
        if hasattr(self, "current_mail") and self.current_mail:
            self.aba_grafico.set_current_email(self.current_mail)

    def enviar_email(self):
        df = self.df_original

        if df.empty or "Qtd. Afastamentos" not in df.columns:
            QMessageBox.warning(self, "Erro", "É necessário realizar uma consulta antes de gerar o e-mail.")
            return

        status_texto = self.status_label.text()
        if " | " in status_texto:
            parte_periodo, parte_situacoes = status_texto.split(" | ", 1)
            if "Nenhum dado encontrado para:" in parte_situacoes:
                parte_situacoes = parte_situacoes.replace("Nenhum dado encontrado para:", "").strip()
            label_periodo = parte_periodo.strip()
            situacoes_descritivas = [s.strip() for s in parte_situacoes.split(",") if s.strip()]
        else:
            label_periodo = "Período não informado"
            situacoes_descritivas = []

        cores_locais = {
            "Fabrica De Máquinas": "#003756",
            "Fabrica De Transportadores": "#ffc62e",
            "Adm": "#009c44",
            "Comercial": "#919191"
        }

        texto_situacoes = ", ".join(situacoes_descritivas)
        corpo_email = f"<span style='font-size:16px;'><b>Relatório {label_periodo}. {texto_situacoes}</b></span><br><br>"

        corpo_email += "<b>Resumo:</b><br>"

        resumo_df = df.groupby(["Local", "Setor", "Situação"]).agg({
            "Qtd. Afastamentos": "sum",
            "Qtd. Total Dias Afastados": "sum"
        }).reset_index()

        for _, row in resumo_df.iterrows():
            corpo_email += (
                f"- {row['Local']} - {row['Setor']} - {row['Situação']} "
                f"({row['Qtd. Afastamentos']} afastamentos / {row['Qtd. Total Dias Afastados']} dias)<br>"
            )

        corpo_email += "<br><table border='1' cellspacing='0' cellpadding='3' style='border-collapse: collapse; width: 100%; font-size: 14px; text-align: center;'>"
        corpo_email += """
        <tr style="background-color: #f2f2f2;">
            <th>Data de Afastamento</th>
            <th>Colaborador</th>
            <th>Setor</th>
            <th>Qtd. Total Dias Afastados</th>
            <th>Data Fim</th>
            <th>Dias para INSS</th>
            <th>CID</th>
            <th>CID Descrição</th>
            <th>Médico</th>
        </tr>
        """

        for _, row in df.iterrows():
            corpo_email += "<tr>"
            corpo_email += (
                f"<td>{row.get('Data Afastamento', '')}</td>"
                f"<td>{row.get('Colaborador', '')}</td>"
                f"<td>{row.get('Setor', '')}</td>"
                f"<td>{row.get('Qtd. Total Dias Afastados', '')}</td>"
                f"<td>{row.get('Data Fim', '')}</td>"
                f"<td>{row.get('Dias para INSS', '')}</td>"
                f"<td>{row.get('CID', '')}</td>"
                f"<td>{row.get('CID Descrição', '')}</td>"
                f"<td>{row.get('Médico', '')}</td>"
            )
            corpo_email += "</tr>"

        corpo_email += "</table>"

        outlook = win32com.client.Dispatch("Outlook.Application")
        email = outlook.CreateItem(0)
        email.Subject = f"Relatório de Afastamentos - {label_periodo}"
        email.HTMLBody = corpo_email + "<br><br>" + email.HTMLBody

        self.current_mail = email
        email.Display()

        QMessageBox.information(self, "Sucesso", "E-mail de afastamentos gerado com sucesso no Outlook!")

    def exibir_custos_afastamentos(self):
        if self.df_original is None or self.df_original.empty:
            QMessageBox.warning(self, "Aviso", "Nenhuma consulta realizada.")
            return
        if not hasattr(self, "df_para_salario") or self.df_para_salario.empty:
            QMessageBox.warning(self, "Aviso", "Dados brutos com códigos não disponíveis.")
            return

        dialog = DialogCustoAfastamentos(self.df_original, self.df_para_salario, self)
        dialog.exec()

    def voltar_menu(self):
        from main import ControlesRH
        self.menu_principal = ControlesRH()
        self.menu_principal.show()
        self.window().close()

class DialogCustoAfastamentos(QDialog):
    def __init__(self, df_agrupado, df_para_salario, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Custo Aproximado Colaboradores Afastados")
        self.resize(1100, 600)
        layout = QVBoxLayout()

        # Carregar JSON
        config_path = os.path.join(os.path.dirname(__file__), "frequencia.json")
        with open(config_path, "r") as f:
            folder_path = json.load(f).get("folder_path", "")
        path_planilha = os.path.join(folder_path, "FPRE905.xlsx")

        # Corrige erro de leitura
        self.open_save_with_excel(path_planilha)

        try:
            df_salarios = pd.read_excel(path_planilha, header=None, dtype=str)
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao ler FPRE905.xlsx: {e}")
            return

        df_salarios[0] = df_salarios[0].astype(str).str.strip().str.lstrip("0")
        df_salarios[1] = df_salarios[1].astype(str).str.strip().str.lstrip("0")

        df_codigos = df_para_salario[["Cadastro", "Local", "Setor"]].copy()
        df_codigos = df_codigos.rename(columns={"Local": "CodEmp", "Setor": "CodLoc"})
        df_codigos[["Cadastro", "CodEmp"]] = df_codigos[["Cadastro", "CodEmp"]].astype(str).apply(lambda x: x.str.strip().str.lstrip("0"))
        df_codigos = df_codigos.drop_duplicates()

        df_base = df_agrupado.copy()
        df_base["Cadastro"] = df_base["Cadastro"].astype(str).str.strip()
        df_final = pd.merge(df_base, df_codigos, on="Cadastro", how="left")

        self.df_exportar = []  # Para exportar depois
        total_geral = Decimal("0.00")
        jornada_dia = Decimal("8.83")

        for _, linha in df_final.iterrows():
            nome = linha.get("Colaborador", "")
            local = linha.get("Local", "")
            setor = linha.get("Setor", "")
            cadastro = str(linha.get("Cadastro", "")).strip().lstrip("0")
            empresa = str(linha.get("CodEmp", "")).strip().lstrip("0")
            try:
                dias = int(linha.get("Qtd. Total Dias Afastados", 0))
            except:
                dias = 0

            rows = df_salarios[(df_salarios[0] == cadastro) & (df_salarios[1] == empresa)]
            if rows.empty:
                continue

            salario_str = str(rows.iloc[0, 5]).strip().replace('.', '').replace(',', '.')
            try:
                salario_base = Decimal(salario_str)
                situacao = unidecode(str(linha.get("Situação", "")).strip().lower())
                if "ferias" in situacao:
                    salario_base *= Decimal("1.33")
                    nome += " (*)"
            except InvalidOperation:
                salario_base = Decimal("0.00")

            salario_dia = (salario_base / Decimal("30")).quantize(Decimal("0.01"))
            salario_hora = (salario_dia / jornada_dia).quantize(Decimal("0.01"))
            custo_total = (salario_dia * dias).quantize(Decimal("0.01"))
            total_geral += custo_total

            self.df_exportar.append([
                local, setor, cadastro, nome,
                dias,
                float(salario_base),
                float(salario_hora),
                float(custo_total)
            ])

        if not self.df_exportar:
            QMessageBox.information(self, "Custo dos Afastamentos", "Nenhum custo pôde ser calculado.")
            return

        # Layout superior com botão e aviso
        topo_layout = QHBoxLayout()

        info_label = QLabel(
            "Os valores representam uma média aproximada. Devido aos diversos fatores que compõem a folha de pagamento, "
            "podem ocorrer pequenas variações em relação ao valor final."
        )
        info_label.setStyleSheet("font-size: 8pt; color: red;")
        info_label.setWordWrap(True)
        topo_layout.addWidget(info_label, stretch=1)

        btn_exportar = QPushButton("Exportar para Excel")
        btn_exportar.setFixedSize(150, 30)
        btn_exportar.clicked.connect(self.exportar_para_excel)
        topo_layout.addWidget(btn_exportar)

        layout.addLayout(topo_layout)

        colunas = ["Local", "Setor", "Cadastro", "Colaborador", "Dias Afastado",
                   "Salário Base", "Salário/hora", "Custo Total"]

        tabela = QTableWidget()
        tabela.setColumnCount(len(colunas))
        tabela.setRowCount(len(self.df_exportar))
        tabela.setHorizontalHeaderLabels(colunas)
        tabela.verticalHeader().setVisible(False)

        fonte = tabela.font()
        fonte.setPointSize(8)
        tabela.setFont(fonte)

        for row_idx, linha_dados in enumerate(self.df_exportar):
            for col_idx, valor in enumerate(linha_dados):
                item = QTableWidgetItem(f"R$ {valor:.2f}" if isinstance(valor, float) else str(valor))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                tabela.setItem(row_idx, col_idx, item)

        header = tabela.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        header.setStretchLastSection(True)

        layout.addWidget(tabela)

        legenda_label = QLabel("(*) Colaboradores em férias - salário com acréscimo de 33.3% para calculo.")
        legenda_label.setStyleSheet("font-size: 8pt; color: black; margin-top: 5px;")
        layout.addWidget(legenda_label)

        total_label = QLabel(f"<b>Total Geral Aproximado dos Afastamentos: R$ {total_geral:.2f}</b>")
        total_label.setStyleSheet("font-size: 14px; color: red; margin-top: 10px;")
        layout.addWidget(total_label)

        self.setLayout(layout)

    def exportar_para_excel(self):
        import pandas as pd
        from PyQt6.QtWidgets import QFileDialog

        if not self.df_exportar:
            QMessageBox.warning(self, "Erro", "Nenhum dado para exportar.")
            return

        df = pd.DataFrame(self.df_exportar, columns=[
            "Local", "Setor", "Cadastro", "Colaborador", "Dias Afastado",
            "Salário Base", "Salário/hora", "Custo Total"
        ])

        path, _ = QFileDialog.getSaveFileName(self, "Salvar como Excel", "", "Excel Files (*.xlsx)")
        if path:
            df.to_excel(path, index=False)
            QMessageBox.information(self, "Exportado", "Arquivo exportado com sucesso!")

    def open_save_with_excel(self, file_path):
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        was_excel_open = excel.Workbooks.Count > 0
        excel.Visible = False

        try:
            workbook = excel.Workbooks.Open(file_path)
            workbook.Save()
            workbook.Close(SaveChanges=False)
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao abrir/salvar a planilha no Excel: {e}")
        finally:
            if not was_excel_open:
                excel.Quit()



if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = Appatestados()
    window.show()
    sys.exit(app.exec())