import sys
from PyQt6.QtWidgets import (
    QApplication, QWidget, QHBoxLayout, QPushButton, QLineEdit, QMessageBox, QDateEdit, QScrollArea, QCheckBox
)
import pandas as pd
from datetime import datetime, timedelta
import win32com.client
import os
from PyQt6.QtGui import QColor
from contextlib import contextmanager
import oracledb
import json
from afastamentosgrafico import GraficoAfastamento
from functools import partial
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QTableWidget, QTableWidgetItem, QLabel
from PyQt6.QtGui import QIcon, QFont
from PyQt6.QtCore import Qt, QDate
from Database import get_connection

class AppAdvertencias(QWidget):
    def __init__(self):
        super().__init__()
        self.df_cnh = pd.DataFrame()
        self.df_rg = pd.DataFrame()
        self.df_original = None
        self.df_completo = None
        self.carregar_dicionario_locais_com_busca()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Consulta de Advertências")
        self.resize(1100, 900)

        # Configurar o ícone da janela
        icon_path = os.path.join(os.path.dirname(__file__), 'icone.ico')  # Ajuste o caminho conforme necessário
        if os.path.exists(icon_path):  # Verifica se o arquivo existe
            self.setWindowIcon(QIcon(icon_path))

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        # Linha 2: filtros de data e botão periodo
        data_hoje = QDate.currentDate()
        primeiro_dia_mes = QDate(data_hoje.year(), data_hoje.month(), 1)

        linha1 = QHBoxLayout()
        linha1.setAlignment(Qt.AlignmentFlag.AlignLeft)

        self.date_inicio = QDateEdit()
        self.date_inicio.setDate(primeiro_dia_mes)
        self.date_inicio.setCalendarPopup(True)
        self.date_inicio.setDisplayFormat("dd/MM/yyyy")
        linha1.addWidget(QLabel("Início:"))
        linha1.addWidget(self.date_inicio)

        self.date_fim = QDateEdit()
        self.date_fim.setDate(data_hoje)
        self.date_fim.setCalendarPopup(True)
        self.date_fim.setDisplayFormat("dd/MM/yyyy")
        linha1.addWidget(QLabel("Fim:"))
        linha1.addWidget(self.date_fim)

        btn_periodo = QPushButton("Consultar Advertências")
        btn_periodo.setFixedSize(140, 30)
        btn_periodo.clicked.connect(self.advertencias_periodo)
        linha1.addSpacing(20)
        linha1.addWidget(btn_periodo)

        linha1.addStretch()  # espaço entre esquerda e direita

        btn_email = QPushButton("E-mail")
        btn_email.setFixedSize(120, 30)
        btn_email.clicked.connect(self.enviar_email)
        linha1.addWidget(btn_email)

        btn_grafico = QPushButton("Abrir Gráfico")
        btn_grafico.setFixedSize(120, 30)
        btn_grafico.clicked.connect(self.abrir_graficos)
        linha1.addWidget(btn_grafico)

        self.btn_voltar = QPushButton("Voltar ao Menu")
        self.btn_voltar.setFixedSize(140, 30)
        self.btn_voltar.clicked.connect(self.voltar_menu)
        linha1.addWidget(self.btn_voltar)

        layout.addLayout(linha1)

        # Linha 2: checkbox afastados
        self.checkbox_afastados = QCheckBox("Listar Afastados/Demitidos")
        layout.addWidget(self.checkbox_afastados)

        # Linha 3: status
        self.status_label = QLabel("Nenhuma consulta realizada ainda.")
        self.status_label.setWordWrap(True)
        self.status_label.setMaximumWidth(1800)
        self.status_label.setStyleSheet("color: red; font-weight: bold;")
        layout.addWidget(self.status_label)

        # Linha 4: filtro de texto
        self.search_field = QLineEdit()
        self.search_field.setPlaceholderText("Filtrar usando / entre dados. Também pode digitar COLUNA:DADO")
        self.search_field.textChanged.connect(self.apply_global_filter)
        layout.addWidget(self.search_field)

        # Linha 6: tabela
        self.tableWidget = QTableWidget()
        layout.addWidget(self.tableWidget)

        self.setLayout(layout)
        self.current_mail = None

        # Conectar duplo clique uma única vez
        self.tableWidget.cellDoubleClicked.connect(self.abrir_detalhes_colaborador)

    def advertencias_periodo(self):
        sitafa_excluir = ("007", "024", "916", "913", "053", "003")
        incluir_afastados = self.checkbox_afastados.isChecked()
        sitafa_condicao = "" if incluir_afastados else f"AND F.SITAFA NOT IN {sitafa_excluir}"

        data_inicio = self.date_inicio.date().toPyDate()
        data_fim = self.date_fim.date().toPyDate()

        query = f"""
        SELECT 
            N.NUMEMP,
            H.CODLOC,
            N.NUMCAD,
            F.NOMFUN,
            TO_CHAR(N.DATNOT, 'DD/MM/YYYY') AS DATNOT,
            N.TIPNOT,
            N.NOTFIC,
            F.SITAFA
        FROM 
            R038NOT N
        JOIN 
            R034FUN F ON F.NUMEMP = N.NUMEMP 
                      AND F.TIPCOL = N.TIPCOL 
                      AND F.NUMCAD = N.NUMCAD
        JOIN 
            R016HIE H ON H.NUMLOC = F.NUMLOC
        WHERE 
            N.TIPCOL = 1
            AND N.NUMEMP IN (10, 11, 16, 17, 19)
            AND N.TIPNOT IN (7, 11, 12)
            {sitafa_condicao}
        """

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
            self.status_label.setText("Consulta de Advertências no Período — Nenhum resultado encontrado.")
            return

        df = df.map(lambda x: x.replace("\n", " ").replace("\r", " ").strip() if isinstance(x, str) else x)
        df.fillna("", inplace=True)

        df.rename(columns={
            "NUMEMP": "Local",
            "CODLOC": "Setor",
            "NUMCAD": "Cadastro",
            "NOMFUN": "Colaborador",
            "DATNOT": "Data",
            "TIPNOT": "Tipo",
            "NOTFIC": "Anotacão"
        }, inplace=True)

        tipo_map = {
            7: "Advertência de Segurança",
            11: "Advertência Disciplinar",
            12: "Suspensão Disciplinar"
        }
        df["Tipo"] = df["Tipo"].astype(int).map(tipo_map)
        df["Data"] = pd.to_datetime(df["Data"], format="%d/%m/%Y")
        df.sort_values("Data", ascending=False, inplace=True)
        df.reset_index(drop=True, inplace=True)

        df[["Local", "Setor"]] = df["Setor"].apply(lambda codloc: pd.Series(self.buscar_local_setor(codloc)))
        df = df[["Local", "Setor", "Cadastro", "Colaborador", "Data", "Tipo", "Anotacão"]]
        self.df_original = df.copy()

        agrupado = df.groupby(["Local", "Setor", "Cadastro", "Colaborador"]).agg({
            "Data": "max"
        }).reset_index()
        agrupado.rename(columns={"Data": "Última Advertência"}, inplace=True)

        # Calcular a quantidade total de advertências
        agrupado["Qtd. Total"] = df.groupby(["Local", "Setor", "Cadastro", "Colaborador"]).size().values
        
        # Calcular a quantidade de advertências por tipo
        df_adv = df[df["Tipo"].isin(["Advertência de Segurança", "Advertência Disciplinar"])]
        qtd_adv = df_adv.groupby(["Local", "Setor", "Cadastro", "Colaborador"]).size().reset_index(name="Qtd. Advertências")
        
        # Calcular a quantidade de suspensões
        df_susp = df[df["Tipo"] == "Suspensão Disciplinar"]
        qtd_susp = df_susp.groupby(["Local", "Setor", "Cadastro", "Colaborador"]).size().reset_index(name="Qtd. Suspensões")
        
        # Adicionar as contagens ao DataFrame agrupado
        agrupado = agrupado.merge(qtd_adv, on=["Local", "Setor", "Cadastro", "Colaborador"], how="left")
        agrupado = agrupado.merge(qtd_susp, on=["Local", "Setor", "Cadastro", "Colaborador"], how="left")
        
        # Preencher valores nulos com zero
        agrupado["Qtd. Advertências"] = agrupado["Qtd. Advertências"].fillna(0).astype(int)
        agrupado["Qtd. Suspensões"] = agrupado["Qtd. Suspensões"].fillna(0).astype(int)

        filtro_periodo = (df["Data"] >= pd.to_datetime(data_inicio)) & (df["Data"] <= pd.to_datetime(data_fim))
        df_periodo = df[filtro_periodo]
        qtd_periodo = df_periodo.groupby(["Local", "Setor", "Cadastro", "Colaborador"]).size().reset_index(
            name="Qtd. Advertências no Período")

        agrupado = agrupado.merge(qtd_periodo, on=["Local", "Setor", "Cadastro", "Colaborador"], how="left")
        agrupado["Qtd. Advertências no Período"] = agrupado["Qtd. Advertências no Período"].fillna(0).astype(int)

        agrupado = agrupado[
            (agrupado["Última Advertência"] >= pd.to_datetime(data_inicio)) &
            (agrupado["Última Advertência"] <= pd.to_datetime(data_fim))
            ]

        if agrupado.empty:
            self.tableWidget.setRowCount(0)
            self.tableWidget.setColumnCount(0)
            self.status_label.setText("Consulta de Advertências no Período — Nenhum resultado encontrado.")
            return

        agrupado.sort_values(by=["Última Advertência", "Qtd. Total"], ascending=[False, False],
                             inplace=True)
        agrupado.reset_index(drop=True, inplace=True)

        # reorganizar ordem das colunas
        agrupado = agrupado[[
            "Local", "Setor", "Cadastro", "Colaborador",
            "Última Advertência", "Qtd. Advertências no Período", 
            "Qtd. Advertências", "Qtd. Suspensões", "Qtd. Total"
        ]]
        
        # Criar versão para visualização, sem a coluna do período
        agrupado_visual = agrupado[[
            "Local", "Setor", "Cadastro", "Colaborador",
            "Última Advertência", "Qtd. Advertências", "Qtd. Suspensões", "Qtd. Total"
        ]]
        colunas = agrupado_visual.columns.tolist()

        self.tableWidget.setColumnCount(len(colunas))
        self.tableWidget.setRowCount(len(agrupado_visual))
        self.tableWidget.setHorizontalHeaderLabels(colunas)
        
        fonte = self.tableWidget.font()
        fonte.setPointSize(8)
        self.tableWidget.setFont(fonte)
        
        for row_idx, row in agrupado_visual.iterrows():
            for col_idx, value in enumerate(row):
                if isinstance(value, pd.Timestamp):
                    value = value.strftime("%d/%m/%Y")
                item = QTableWidgetItem(str(value))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)

                if colunas[col_idx] == "Qtd. Total" and int(value) >= 3:
                    item.setBackground(QColor("yellow"))
                elif colunas[col_idx] == "Qtd. Advertências" and int(value) >= 3:
                    item.setBackground(QColor("yellow"))
                elif colunas[col_idx] == "Qtd. Suspensões" and int(value) >= 1:
                    item.setBackground(QColor("yellow"))

                self.tableWidget.setItem(row_idx, col_idx, item)

        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.resizeRowsToContents()
        self.tableWidget.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        # Mantemos o DataFrame completo para uso futuro, mas exibimos a versão filtrada
        self.df_cnh = agrupado.copy()

        status_info = " (Listando Colaboradores Afastados/Demitidos)" if incluir_afastados else " (Somente Colaboradores Ativos)"
        periodo = f"{self.date_inicio.date().toString('dd/MM/yyyy')} até {self.date_fim.date().toString('dd/MM/yyyy')}"
        self.status_label.setText(
            f"Últimas Advertências entre {periodo} {status_info}"
        )

    def abrir_detalhes_colaborador(self, row, column):
        local = self.tableWidget.item(row, 0).text().strip()
        setor = self.tableWidget.item(row, 1).text().strip()
        cadastro = self.tableWidget.item(row, 2).text().strip()
        colaborador = self.tableWidget.item(row, 3).text().strip()

        df = self.df_original.copy()
        for col in ["Local", "Setor", "Cadastro", "Colaborador"]:
            df[col] = df[col].astype(str).str.strip()

        detalhes = df[
            (df["Local"] == local) &
            (df["Setor"] == setor) &
            (df["Cadastro"] == cadastro) &
            (df["Colaborador"] == colaborador)
            ]

        if detalhes.empty:
            return

        # Separar colunas sem a anotação
        colunas_sem_anotacao = [col for col in detalhes.columns if col != "Anotacão"]
        colunas_exibir = colunas_sem_anotacao + ["Anotacão"]

        self.detalhes_janela = QWidget()
        self.detalhes_janela.setWindowTitle(f"Detalhes de {colaborador}")
        layout = QVBoxLayout(self.detalhes_janela)

        tabela = QTableWidget()
        tabela.setColumnCount(len(colunas_exibir))
        tabela.setRowCount(len(detalhes))
        tabela.setHorizontalHeaderLabels(colunas_exibir)

        fonte = tabela.font()
        fonte.setPointSize(8)
        tabela.setFont(fonte)

        for row_idx in range(len(detalhes)):
            for col_idx, col_nome in enumerate(colunas_exibir):
                if col_nome == "Anotacão":
                    btn = QPushButton("Exibir Anotação")
                    anotacao = detalhes.iloc[row_idx]["Anotacão"]
                    btn.clicked.connect(lambda _, txt=anotacao: self.exibir_anotacao(txt))
                    tabela.setCellWidget(row_idx, col_idx, btn)
                else:
                    valor = detalhes.iloc[row_idx][col_nome]
                    if isinstance(valor, pd.Timestamp):
                        valor = valor.strftime("%d/%m/%Y")
                    item = QTableWidgetItem(str(valor))
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    tabela.setItem(row_idx, col_idx, item)

        tabela.resizeColumnsToContents()
        tabela.resizeRowsToContents()
        layout.addWidget(tabela)

        self.detalhes_janela.resize(800, 600)
        self.detalhes_janela.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        self.detalhes_janela.show()

    def exibir_anotacao(self, texto):
        self.janela_anotacao = QWidget()
        self.janela_anotacao.setWindowTitle("Anotação")
        layout = QVBoxLayout(self.janela_anotacao)

        # Limpa espaços extras e quebras de linha duplicadas
        texto_limpo = "\n".join([linha.strip() for linha in texto.strip().splitlines() if linha.strip()])

        # Cria QLabel e aplica alinhamento à direita
        label = QLabel(texto_limpo)
        label.setWordWrap(True)
        label.setAlignment(Qt.AlignmentFlag.AlignLeft)

        # Adiciona label em área com rolagem
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(label)

        layout.addWidget(scroll)

        self.janela_anotacao.resize(500, 300)
        self.janela_anotacao.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        self.janela_anotacao.show()

    def enviar_email(self):
        if self.df_cnh is None or self.df_cnh.empty:
            QMessageBox.warning(self, "Erro", "Nenhuma consulta foi realizada ainda.")
            return

        linhas_visiveis = []
        for row in range(self.tableWidget.rowCount()):
            if not self.tableWidget.isRowHidden(row):
                linha = {}
                for col in range(self.tableWidget.columnCount()):
                    header = self.tableWidget.horizontalHeaderItem(col).text()
                    item = self.tableWidget.item(row, col)
                    linha[header] = item.text() if item else ""
                linhas_visiveis.append(linha)

        df = pd.DataFrame(linhas_visiveis)
        if df.empty:
            QMessageBox.information(self, "Aviso", "Nenhum dado visível para enviar no e-mail.")
            return

        if "Cadastro" in df.columns:
            df.drop(columns=["Cadastro"], inplace=True)

        # Remover coluna oculta do email
        if "Qtd. Advertências no Período" in df.columns:
            df.drop(columns=["Qtd. Advertências no Período"], inplace=True)
        
        df.sort_values(by=["Local", "Setor", "Colaborador"], inplace=True)

        cores_locais = {
            "Fabrica De Máquinas": "#003756",
            "Fabrica De Transportadores": "#ffc62e",
            "Adm": "#009c44",
            "Comercial": "#919191"
        }

        texto_status = self.status_label.text()
        corpo_email = f"<span style='font-size:16px;'>Segue o relatório de {texto_status}.</span><br><br>"

        corpo_email += "<table border='1' cellspacing='0' cellpadding='4' style='border-collapse: collapse; font-size: 13px; width: 100%; text-align: center;'>"

        colunas = df.columns.tolist()
        corpo_email += "<tr style='background-color:#f2f2f2;'>" + "".join(
            f"<th>{col}</th>" for col in colunas) + "</tr>"

        # Agrupar por Local, depois Setor dentro de Local
        for local, df_local in df.groupby("Local"):
            cor_local = cores_locais.get(local, "#000000")
            for setor, df_setor in df_local.groupby("Setor"):
                for i, (_, row) in enumerate(df_setor.iterrows()):
                    corpo_email += "<tr>"
                    for col in colunas:
                        valor = row[col]
                        style = ""

                        if col == "Qtd. Total" and valor.isdigit() and int(valor) >= 3:
                            style = " style='background-color:yellow;'"
                        elif col == "Qtd. Advertências" and valor.isdigit() and int(valor) >= 3:
                            style = " style='background-color:yellow;'"
                        elif col == "Qtd. Suspensões" and valor.isdigit() and int(valor) >= 1:
                            style = " style='background-color:yellow;'"

                        if col == "Local":
                            if i == 0 and df_setor.index[0] == df_local.index[0]:
                                valor = f"<span style='color:white;'>{valor}</span>"
                                style = f" style='background-color:{cor_local}; color:white;'"
                                rowspan = len(df_local)
                                corpo_email += f"<td rowspan='{rowspan}'{style}>{valor}</td>"
                            continue

                        elif col == "Setor":
                            if i == 0:
                                rowspan = len(df_setor)
                                corpo_email += f"<td rowspan='{rowspan}'>{valor}</td>"
                            continue

                        corpo_email += f"<td{style}>{valor}</td>"
                    corpo_email += "</tr>"

        corpo_email += "</table>"

        try:
            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application")
            email = outlook.CreateItem(0)
            email.Subject = f"Relatório de Advertências - {texto_status}"
            email.HTMLBody = corpo_email + "<br><br>" + email.HTMLBody
            self.current_mail = email
            email.Display()
            QMessageBox.information(self, "Sucesso", "E-mail gerado com sucesso no Outlook!")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao gerar e-mail: {str(e)}")

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


    def voltar_menu(self):
        from main import ControlesRH
        self.menu_principal = ControlesRH()
        self.menu_principal.show()
        self.window().close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = AppAdvertencias()
    window.show()
    sys.exit(app.exec())