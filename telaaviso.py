import sys
from PyQt6.QtCore import QCoreApplication, Qt
QCoreApplication.setAttribute(Qt.ApplicationAttribute.AA_ShareOpenGLContexts)

from PyQt6.QtWidgets import (
    QApplication, QWidget, QHBoxLayout, QPushButton, QLineEdit, QMessageBox, QGridLayout
)
import pandas as pd
from datetime import datetime, timedelta
import win32com.client
import os
import json
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QTableWidget, QTableWidgetItem, QLabel, QSizePolicy, QSpacerItem
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon, QFont
from email.message import EmailMessage
import smtplib
from Database import get_connection


class AppTelaAvisos(QWidget):
    def __init__(self):
        super().__init__()
        self.df_original = pd.DataFrame()
        self.carregar_dicionario_locais_com_busca()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Tela de Avisos")
        self.resize(700, 120)

        icon_path = 'icone.ico'
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.layout_principal = QVBoxLayout()
        self.layout_principal.setAlignment(Qt.AlignmentFlag.AlignTop)

        self.cards_layout = QGridLayout()
        self.layout_principal.addLayout(self.cards_layout)

        self.current_mail = None
        self.criar_dashboard_inicial()

        self.setLayout(self.layout_principal)


    def criar_dashboard_inicial(self):
        alertas = {}

        # Cria tableWidget temporário apenas para usar nas consultas
        self.tableWidget = QTableWidget()

        for nome, func, cor in [
            ("Termino Experiência", self.consultar_experiencia, "#007acc"),
            ("Retornos", self.consultar_retornos, "#009e60"),
            ("Férias Vencendo", self.consultar_ferias, "#ff9900"),
            ("Vistos Vencendo", self.consultar_vistos, "#800000")
        ]:
            func()
            contagem = 0
            for row in range(self.table_temporaria.rowCount()):
                for col in range(self.table_temporaria.columnCount()):
                    item = self.table_temporaria.item(row, col)
                    if item:
                        cor_item = item.background().color()
                        if cor_item in (Qt.GlobalColor.yellow, Qt.GlobalColor.red):
                            contagem += 1
                            break
            alertas[nome] = (contagem, cor, func)

        del self.tableWidget


        row, col = 0, 0
        for nome, (qtd, cor, callback) in alertas.items():
            botao = QPushButton(f"{nome}\n{qtd}")
            botao.setFixedSize(220, 100)
            botao.setStyleSheet(
                f"background-color: {cor}; color: white; font-weight: bold; font-size: 18px;"
            )
            botao.clicked.connect(
                lambda _, titulo=nome, cb=callback: (
                    cb(),
                    self.abrir_detalhamento(titulo, cb, self.table_temporaria),
                )
            )

            self.cards_layout.addWidget(botao, row, col)

            col += 1
            if col == 3:  # quebra linha a cada 3 cards
                col = 0
                row += 1


    def executar_consulta_simples(self, funcao_consulta):
        funcao_consulta()
        return self.df_original.copy()

    def consultar_experiencia(self):
        query = """
        SELECT 
            f.NumEmp,
            H.CODLOC,
            f.NUMCAD,
            f.NomFun, 
            TO_CHAR(f.DatAdm, 'DD/MM/YYYY') AS DatAdm
        FROM R034FUN f
        JOIN R034CPL c 
            ON f.NumEmp = c.NumEmp 
            AND f.TipCol = c.TipCol 
            AND f.NumCad = c.NumCad
        JOIN R016HIE H ON f.NUMLOC = H.NUMLOC
        WHERE f.NumEmp IN (10, 16, 17, 19, 11)
          AND f.TipCol IN (1, 2)
          AND f.SitAfa = '001'
          AND TRUNC(f.DatAdm) >= TRUNC(SYSDATE - 60)
        """

        with get_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(query)
                colunas = [desc[0] for desc in cursor.description]
                registros = [dict(zip(colunas, row)) for row in cursor.fetchall()]

        df = pd.DataFrame(registros)

        if df.empty:
            self.tableWidget.setRowCount(0)
            self.tableWidget.setColumnCount(0)
            return

        df = df.map(lambda x: x.replace("\n", " ").replace("\r", " ").strip() if isinstance(x, str) else x)
        df.fillna("", inplace=True)

        df["DatAdm_dt"] = pd.to_datetime(df["DATADM"], format="%d/%m/%Y", errors='coerce')
        df["Termino_dt"] = df["DatAdm_dt"] + pd.to_timedelta(60, unit='D')
        df["Termino"] = df["Termino_dt"].dt.strftime("%d/%m/%Y")

        df.rename(columns={
            "NUMEMP": "Local",
            "CODLOC": "Setor",
            "NUMCAD": "Cadastro",
            "NOMFUN": "Colaborador"
        }, inplace=True)

        colunas_ordenadas = ["Local", "Setor", "Cadastro", "Colaborador", "Termino"]
        df[["Local", "Setor"]] = df["Setor"].apply(lambda codloc: pd.Series(self.buscar_local_setor(codloc)))
        df.sort_values("Termino_dt", inplace=True)
        df.reset_index(drop=True, inplace=True)
        df = df[colunas_ordenadas + ["Termino_dt"]]  # Termino_dt mantida temporariamente

        table = QTableWidget()
        colunas = ["Local", "Setor", "Cadastro", "Colaborador", "Termino"]
        table.setColumnCount(len(colunas))
        table.setRowCount(len(df))
        table.setHorizontalHeaderLabels(colunas)
        hoje = datetime.now()

        for row_idx, row in df.iterrows():
            dias_restantes = (row["Termino_dt"] - hoje).days
            cor = Qt.GlobalColor.yellow if dias_restantes <= 15 else None

            for col_idx, col in enumerate(colunas):
                item = QTableWidgetItem(str(row[col]) if row[col] is not None else "")
                if cor:
                    item.setBackground(cor)
                    if cor == Qt.GlobalColor.red:
                        item.setForeground(Qt.GlobalColor.white)
                    elif cor == Qt.GlobalColor.yellow:
                        item.setForeground(Qt.GlobalColor.black)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                table.setItem(row_idx, col_idx, item)

        self.df_original = df[colunas].copy()
        self.table_temporaria = table

    def consultar_retornos(self):
        query = """
        SELECT 
            A.NUMEMP, 
            H.CODLOC, 
            A.NUMCAD, 
            F.NOMFUN, 
            S.DESSIT,        
            TO_CHAR(A.DATTER, 'DD/MM/YYYY') AS DATTER
        FROM R038AFA A
        JOIN R034FUN F ON A.NUMEMP = F.NUMEMP AND A.NUMCAD = F.NUMCAD
        JOIN R192DOE D ON A.CODDOE = D.CODDOE
        JOIN R016HIE H ON F.NUMLOC = H.NUMLOC
        JOIN R010SIT S ON A.SITAFA = S.CODSIT
        WHERE A.NUMEMP IN (10, 16, 17, 18, 19)
          AND F.TIPCOL = 1
          AND A.SITAFA IN (2, 3, 4, 6, 11, 14, 20, 24, 25, 52, 53, 54, 56, 61, 64, 72, 73, 74, 75, 201, 912, 913, 918, 979)
          AND TRUNC(A.DATTER) >= TRUNC(SYSDATE)
        """

        with get_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(query)
                colunas = [desc[0] for desc in cursor.description]
                registros = [dict(zip(colunas, row)) for row in cursor.fetchall()]

        df = pd.DataFrame(registros)
        if df.empty:
            self.df_original = pd.DataFrame()
            self.table_temporaria = QTableWidget()
            return

        df = df.map(lambda x: x.replace("\n", " ").replace("\r", " ").strip() if isinstance(x, str) else x)
        df.fillna("", inplace=True)
        df["DatTer_dt"] = pd.to_datetime(df["DATTER"], format="%d/%m/%Y", errors='coerce')

        df.rename(columns={
            "NUMEMP": "Local",
            "CODLOC": "Setor",
            "NUMCAD": "Cadastro",
            "NOMFUN": "Colaborador",
            "DESSIT": "Situação",
            "DATTER": "Termino"
        }, inplace=True)

        colunas_ordenadas = ["Local", "Setor", "Cadastro", "Colaborador", "Situação", "Termino"]
        df[["Local", "Setor"]] = df["Setor"].apply(lambda codloc: pd.Series(self.buscar_local_setor(codloc)))
        df.sort_values("DatTer_dt", inplace=True)
        df.reset_index(drop=True, inplace=True)
        df = df[colunas_ordenadas + ["DatTer_dt"]]

        table = QTableWidget()
        table.setColumnCount(len(colunas_ordenadas))
        table.setRowCount(len(df))
        table.setHorizontalHeaderLabels(colunas_ordenadas)

        hoje = datetime.now()
        for row_idx, row in df.iterrows():
            dias_restantes = (row["DatTer_dt"] - hoje).days
            cor = Qt.GlobalColor.yellow if dias_restantes <= 15 else None

            for col_idx, col in enumerate(colunas_ordenadas):
                item = QTableWidgetItem(str(row[col]) if row[col] is not None else "")
                if cor:
                    item.setBackground(cor)
                    if cor == Qt.GlobalColor.red:
                        item.setForeground(Qt.GlobalColor.white)
                    elif cor == Qt.GlobalColor.yellow:
                        item.setForeground(Qt.GlobalColor.black)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                table.setItem(row_idx, col_idx, item)

        self.df_original = df[colunas_ordenadas].copy()
        self.table_temporaria = table

    def consultar_ferias(self):
        query = """
        SELECT 
            P.NUMEMP,
            H.CODLOC,
            P.NUMCAD,
            F.NOMFUN,
            TO_CHAR(P.FIMPER, 'DD/MM/YYYY') AS FIMPER,
            TO_CHAR(P.QTDSLD, 'FM99990.00') AS QTDSLD,
            TO_CHAR(P.QTDDIR, 'FM99990.00') AS QTDDIR,
            TO_CHAR(P.LIMCON, 'DD/MM/YYYY') AS LIMCON    
        FROM R040PER P
        JOIN R034FUN F
          ON F.NUMEMP = P.NUMEMP
         AND F.TIPCOL = P.TIPCOL
         AND F.NUMCAD = P.NUMCAD
        JOIN R016HIE H ON F.NUMLOC = H.NUMLOC
        WHERE F.NUMEMP IN (10, 16, 17, 19, 11)
          AND F.TIPCOL IN (1)
          AND F.SITAFA = 001
          AND P.SITPER = 0
          AND P.QTDDIR > 0
          AND TRUNC(P.FIMPER) <= TRUNC(SYSDATE + 60)
        """

        with get_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(query)
                colunas = [desc[0] for desc in cursor.description]
                registros = [dict(zip(colunas, row)) for row in cursor.fetchall()]

        df = pd.DataFrame(registros)
        if df.empty:
            self.df_original = pd.DataFrame()
            self.table_temporaria = QTableWidget()
            return

        df = df.map(lambda x: x.replace("\n", " ").replace("\r", " ").strip() if isinstance(x, str) else x)
        df.fillna("", inplace=True)

        df["Fim_dt"] = pd.to_datetime(df["FIMPER"], format="%d/%m/%Y", errors='coerce')
        df["Limite_dt"] = pd.to_datetime(df["LIMCON"], format="%d/%m/%Y", errors='coerce')

        df.rename(columns={
            "NUMEMP": "Local",
            "CODLOC": "Setor",
            "NUMCAD": "Cadastro",
            "NOMFUN": "Colaborador",
            "FIMPER": "Fim",
            "QTDSLD": "Saldo",
            "QTDDIR": "Direito",
            "LIMCON": "Limite"
        }, inplace=True)

        colunas_ordenadas = ["Local", "Setor", "Cadastro", "Colaborador", "Fim", "Saldo", "Direito", "Limite"]
        df[["Local", "Setor"]] = df["Setor"].apply(lambda codloc: pd.Series(self.buscar_local_setor(codloc)))
        df.sort_values("Fim_dt", ascending=True, inplace=True)
        df.reset_index(drop=True, inplace=True)

        table = QTableWidget()
        table.setColumnCount(len(colunas_ordenadas))
        table.setRowCount(len(df))
        table.setHorizontalHeaderLabels(colunas_ordenadas)

        hoje = datetime.now()
        for row_idx, row in df.iterrows():
            dias_limite = (row["Limite_dt"] - hoje).days if pd.notnull(row["Limite_dt"]) else None
            cor = Qt.GlobalColor.red if dias_limite is not None and dias_limite <= 60 else None

            for col_idx, col in enumerate(colunas_ordenadas):
                valor = str(row[col]) if row[col] is not None else ""
                item = QTableWidgetItem(valor)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if cor:
                    item.setBackground(cor)
                    if cor == Qt.GlobalColor.red:
                        item.setForeground(Qt.GlobalColor.white)
                    elif cor == Qt.GlobalColor.yellow:
                        item.setForeground(Qt.GlobalColor.black)
                    item.setForeground(Qt.GlobalColor.white)
                table.setItem(row_idx, col_idx, item)

        df["Qtd.Contagem"] = 1
        self.df_original = df.copy()
        self.table_temporaria = table

    def consultar_vistos(self):
        query = """
        SELECT 
            F.NUMEMP, 
            H.CODLOC, 
            F.NUMCAD, 
            F.NOMFUN,
            N.DESNAC,
            TO_CHAR(F.DVLEST, 'DD/MM/YYYY') AS DVLEST,
            TO_CHAR(E.DATTER, 'DD/MM/YYYY') AS DATTER,
            E.VISEST
        FROM R034FUN F
        JOIN R016HIE H 
            ON F.NUMLOC = H.NUMLOC
        JOIN R023NAC N 
            ON F.CODNAC = N.CODNAC
        JOIN R034EST E 
            ON F.NUMEMP = E.NUMEMP
           AND F.TIPCOL = E.TIPCOL
           AND F.NUMCAD = E.NUMCAD
        WHERE F.NUMEMP IN (10, 16, 17, 18, 19)
          AND F.SITAFA NOT IN ('007')
          AND F.CODNAC NOT IN (10)
        """

        with get_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(query)
                colunas = [desc[0] for desc in cursor.description]
                registros = [dict(zip(colunas, row)) for row in cursor.fetchall()]

        df = pd.DataFrame(registros)
        if df.empty:
            self.df_original = pd.DataFrame()
            self.table_temporaria = QTableWidget()
            return

        df = df.map(lambda x: x.replace("\n", " ").replace("\r", " ").strip() if isinstance(x, str) else x)
        df.fillna("", inplace=True)

        # converter datas
        df["DVLEST_dt"] = pd.to_datetime(df["DVLEST"], format="%d/%m/%Y", errors='coerce')
        df["DATTER_dt"] = pd.to_datetime(df["DATTER"], format="%d/%m/%Y", errors='coerce')

        # mapear condição estrangeiro
        condicoes = {
            "1": "Visto Permanente",
            "2": "Visto Temporário",
            "3": "Asilado",
            "4": "Refugiado",
            "5": "Solicitante de Refúgio",
            "6": "Fora do Brasil"
        }
        df["Condição"] = df["VISEST"].astype(str).map(condicoes).fillna("Não definido")

        # renomear colunas
        df.rename(columns={
            "NUMEMP": "Local",
            "CODLOC": "Setor",
            "NUMCAD": "Cadastro",
            "NOMFUN": "Colaborador",
            "DESNAC": "Nacionalidade",
            "DVLEST": "Vencimento do Visto",
            "DATTER": "Termino de Residência"
        }, inplace=True)

        # limpar datas inválidas (30/12/1899 e 31/12/1900)
        datas_invalidas = [pd.Timestamp("1899-12-30"), pd.Timestamp("1900-12-31")]
        df.loc[df["DVLEST_dt"].isin(datas_invalidas), ["DVLEST_dt", "Vencimento do Visto"]] = [pd.NaT, ""]
        df.loc[df["DATTER_dt"].isin(datas_invalidas), ["DATTER_dt", "Termino de Residência"]] = [pd.NaT, ""]

        colunas_ordenadas = [
            "Local", "Setor", "Cadastro", "Colaborador",
            "Nacionalidade", "Vencimento do Visto", "Termino de Residência", "Condição"
        ]

        # substituir local e setor
        df[["Local", "Setor"]] = df["Setor"].apply(lambda codloc: pd.Series(self.buscar_local_setor(codloc)))
        df.sort_values("DVLEST_dt", inplace=True)
        df.reset_index(drop=True, inplace=True)

        # montar tabela
        table = QTableWidget()
        table.setColumnCount(len(colunas_ordenadas))
        table.setRowCount(len(df))
        table.setHorizontalHeaderLabels(colunas_ordenadas)

        hoje = datetime.now()
        for row_idx, row in df.iterrows():
            # checar se alguma data < 90 dias
            amarelo = False
            for campo in ["DVLEST_dt", "DATTER_dt"]:
                if pd.notnull(row[campo]):
                    dias = (row[campo] - hoje).days
                    if dias <= 90:
                        amarelo = True
                        break

            for col_idx, col in enumerate(colunas_ordenadas):
                item = QTableWidgetItem(str(row[col]) if row[col] is not None else "")
                if amarelo:
                    item.setBackground(Qt.GlobalColor.yellow)
                    item.setForeground(Qt.GlobalColor.black)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                table.setItem(row_idx, col_idx, item)

        self.df_original = df[colunas_ordenadas].copy()
        self.table_temporaria = table

    def abrir_detalhamento(self, titulo, callback, tabela):
        callback()
        janela = JanelaConsultaDetalhada(titulo, tabela)
        janela.exec()

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

    def voltar_menu(self):
        from main import ControlesRH
        self.menu_principal = ControlesRH()
        self.menu_principal.show()
        self.window().close()

class JanelaConsultaDetalhada(QDialog):
    def __init__(self, titulo, source_table, parent=None):
        super().__init__(parent)
        self.setWindowTitle(titulo)
        self.resize(1000, 700)
        self.titulo = titulo  # usado no e-mail

        icon_path = 'icone.ico'
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        layout = QVBoxLayout()

        #  Botão de E-mail
        self.botao_email = QPushButton("Enviar E-mail")
        self.botao_email.setFixedWidth(150)
        self.botao_email.clicked.connect(self.enviar_email)
        layout.addWidget(self.botao_email, alignment=Qt.AlignmentFlag.AlignRight)

        #  Campo de busca
        self.search_field = QLineEdit()
        self.search_field.setPlaceholderText("Filtrar usando / entre dados. Também pode digitar COLUNA:DADO")
        layout.addWidget(self.search_field)

        #  Tabela
        self.table = QTableWidget()
        layout.addWidget(self.table)

        self.setLayout(layout)

        self.mostrar_dados(source_table)
        self.search_field.textChanged.connect(self.aplicar_filtro)


    def mostrar_dados(self, source_table):
        self.table.setRowCount(source_table.rowCount())
        self.table.setColumnCount(source_table.columnCount())
        self.table.setHorizontalHeaderLabels(
            [source_table.horizontalHeaderItem(i).text() for i in range(source_table.columnCount())])

        for i in range(source_table.rowCount()):
            for j in range(source_table.columnCount()):
                orig_item = source_table.item(i, j)
                if orig_item:
                    new_item = QTableWidgetItem(orig_item.text())
                    new_item.setFlags(orig_item.flags())
                    new_item.setTextAlignment(orig_item.textAlignment())
                    new_item.setBackground(orig_item.background())
                    new_item.setForeground(orig_item.foreground())
                    self.table.setItem(i, j, new_item)

        self.table.resizeColumnsToContents()
        self.table.resizeRowsToContents()

    def aplicar_filtro(self, text):
        terms = [term.strip() for term in text.split('/') if term.strip()]
        if not terms:
            for row in range(self.table.rowCount()):
                self.table.setRowHidden(row, False)
            return

        headers = {self.table.horizontalHeaderItem(i).text().lower(): i for i in range(self.table.columnCount())}
        for row in range(self.table.rowCount()):
            row_matches = True
            for term in terms:
                match = False
                if ":" in term:
                    col_name, val = map(str.strip, term.split(":", 1))
                    col_idx = headers.get(col_name.lower())
                    if col_idx is not None:
                        item = self.table.item(row, col_idx)
                        if item and item.text().lower() == val.lower():
                            match = True
                else:
                    for col in range(self.table.columnCount()):
                        item = self.table.item(row, col)
                        if item and term.lower() in item.text().lower():
                            match = True
                            break
                if not match:
                    row_matches = False
                    break
            self.table.setRowHidden(row, not row_matches)

    def enviar_email(self):
        linhas = []
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item and item.background().color().name() in ['#ffff00', '#ff0000']:
                    linha = {}
                    for c in range(self.table.columnCount()):
                        header = self.table.horizontalHeaderItem(c).text()
                        item_val = self.table.item(row, c)
                        linha[header] = item_val.text() if item_val else ""
                    linhas.append(linha)
                    break

        if not linhas:
            QMessageBox.information(self, "Aviso", "Nenhum dado em amarelo ou vermelho para enviar.")
            return

        df = pd.DataFrame(linhas)
        if df.empty:
            QMessageBox.information(self, "Aviso", "Nenhum dado válido para e-mail.")
            return

        cores_locais = {
            "Fabrica De Máquinas": "#003756",
            "Fabrica De Transportadores": "#ffc62e",
            "Adm": "#009c44",
            "Comercial": "#919191"
        }

        corpo_email = f"<span style='font-size:16px;'>Segue <b>{self.titulo}</b> dos próximos dias:</span><br><br>"
        corpo_email += "<table border='1' cellspacing='0' cellpadding='3' style='border-collapse: collapse; width: 100%; font-size: 14px; text-align: center;'>"
        corpo_email += "<tr style='background-color: #f2f2f2;'><th>Local</th><th>Setor</th>"

        colunas_restantes = [col for col in df.columns if col not in ["Local", "Setor"]]
        for col in colunas_restantes:
            corpo_email += f"<th>{col}</th>"
        corpo_email += "</tr>"

        for local, df_local in df.groupby("Local"):
            cor_local = cores_locais.get(local, "#000000")
            setores = list(df_local.groupby("Setor"))
            total_linhas_local = len(df_local)
            local_renderizado = False  # controle para mesclar local uma única vez

            for setor, grupo in setores:
                setor_renderizado = False
                for _, row in grupo.iterrows():
                    corpo_email += "<tr>"
                    if not local_renderizado:
                        corpo_email += f"<td rowspan='{total_linhas_local}' style='background-color:{cor_local}; color:white;'>{local}</td>"
                        local_renderizado = True
                    if not setor_renderizado:
                        corpo_email += f"<td rowspan='{len(grupo)}'>{setor}</td>"
                        setor_renderizado = True
                    corpo_email += "".join(f"<td>{row[col]}</td>" for col in colunas_restantes)
                    corpo_email += "</tr>"


        corpo_email += "</table>"

        outlook = win32com.client.Dispatch("Outlook.Application")
        email = outlook.CreateItem(0)
        email.Subject = f"Avisos - {self.titulo}"
        email.HTMLBody = corpo_email + "<br><br>" + email.HTMLBody
        email.Display()

        QMessageBox.information(self, "Sucesso", "E-mail gerado com sucesso no Outlook!")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = AppTelaAvisos()
    window.show()
    sys.exit(app.exec())