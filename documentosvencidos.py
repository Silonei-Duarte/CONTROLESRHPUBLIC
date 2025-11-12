import sys
from PyQt6.QtCore import QCoreApplication, Qt
QCoreApplication.setAttribute(Qt.ApplicationAttribute.AA_ShareOpenGLContexts)

from PyQt6.QtWidgets import (
    QApplication, QWidget, QHBoxLayout, QPushButton, QLineEdit, QMessageBox
)
import pandas as pd
from datetime import datetime, timedelta
import win32com.client
import os
from contextlib import contextmanager
import oracledb
import json
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QTableWidget, QTableWidgetItem, QLabel, QSizePolicy, QSpacerItem
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon, QFont
from email.message import EmailMessage
import smtplib
import os
from PyQt6.QtWidgets import QInputDialog, QLineEdit
from Database import get_connection


class AppConsultaDocumentos(QWidget):
    def __init__(self):
        super().__init__()
        self.df_cnh = pd.DataFrame()
        self.df_rg = pd.DataFrame()
        self.df_original = None
        self.df_completo = None
        self.carregar_dicionario_locais_com_busca()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Consulta de Documentos")
        self.resize(1100, 900)

        # Configurar o ícone da janela
        icon_path = os.path.join(os.path.dirname(__file__), 'icone.ico')  # Ajuste o caminho conforme necessário
        if os.path.exists(icon_path):  # Verifica se o arquivo existe
            self.setWindowIcon(QIcon(icon_path))

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        botoes_layout = QHBoxLayout()
        botoes_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)

        btn_cnh = QPushButton("Consultar CNH")
        btn_cnh.setFixedSize(140, 30)
        btn_cnh.clicked.connect(self.consultar_cnh)
        botoes_layout.addWidget(btn_cnh)

        btn_rg = QPushButton("Consultar RG/CIN")
        btn_rg.setFixedSize(140, 30)
        btn_rg.clicked.connect(self.consultar_rg)
        botoes_layout.addWidget(btn_rg)

        #  Adiciona espaço flexível antes do botão "Voltar ao Menu"
        spacer = QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
        botoes_layout.addItem(spacer)

        btn_email = QPushButton("E-mail Agrupado")
        btn_email.setFixedSize(120, 30)
        btn_email.clicked.connect(self.enviar_email)
        botoes_layout.addWidget(btn_email)

        btn_email_colab = QPushButton("E-mail Colaboradores")
        btn_email_colab.setFixedSize(180, 30)
        btn_email_colab.clicked.connect(self.enviar_email_colaboradores)
        botoes_layout.addWidget(btn_email_colab)

        self.btn_voltar = QPushButton("Voltar ao Menu")
        self.btn_voltar.setFixedSize(140, 30)
        self.btn_voltar.clicked.connect(self.voltar_menu)
        botoes_layout.addWidget(self.btn_voltar)

        layout.addLayout(botoes_layout)

        self.status_label = QLabel("Nenhuma consulta realizada ainda.")
        self.status_label.setWordWrap(True)
        self.status_label.setMaximumWidth(1800)
        self.status_label.setStyleSheet("color: red; font-weight: bold;")
        layout.addWidget(self.status_label)

        self.search_field = QLineEdit()
        self.search_field.setPlaceholderText("Filtrar usando / entre dados. Também pode digitar COLUNA:DADO")
        self.search_field.textChanged.connect(self.apply_global_filter)
        layout.addWidget(self.search_field)

        self.tableWidget = QTableWidget()
        layout.addWidget(self.tableWidget)

        self.setLayout(layout)
        self.current_mail = None

    def consultar_cnh(self):
        query = """
        SELECT 
            f.NumEmp,
            H.CODLOC,
            f.NUMCAD,
            f.NomFun,
            CASE 
                WHEN f.USU_AutVe2 = 1 THEN 'Interno e Externo'
                WHEN f.USU_AutVe2 = 2 THEN 'Somente Interno'
            END AS USU_AutVe2,
            TO_CHAR(c.VenCnh, 'DD/MM/YYYY') AS VenCnh,
            c.EmaPar
        FROM R034FUN f
        JOIN R034CPL c 
            ON f.NumEmp = c.NumEmp 
            AND f.TipCol = c.TipCol 
            AND f.NumCad = c.NumCad
        JOIN R016HIE H ON f.NUMLOC = H.NUMLOC
        WHERE f.NumEmp IN (10, 16, 17, 19, 11)
        AND f.TipCol IN (1, 2)
        AND f.SitAfa = '001'
        AND f.USU_AutVe2 IN (1, 2)
        AND c.VenCnh IS NOT NULL 
        AND c.VenCnh > DATE '0001-01-01'
        ORDER BY c.VenCnh ASC
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
            self.status_label.setText("Consulta CNH — nenhum resultado encontrado.")
            return

        df = df.map(lambda x: x.replace("\n", " ").replace("\r", " ").strip() if isinstance(x, str) else x)
        df.fillna("", inplace=True)

        df.rename(columns={
            "NUMEMP": "Local",
            "CODLOC": "Setor",
            "NUMCAD": "Cadastro",
            "NOMFUN": "Colaborador",
            "USU_AUTVE2": "Autorização",
            "VENCNH": "Vencimento",
            "EMAPAR": "E-mail Particular",
        }, inplace=True)

        df["Vencimento_dt"] = pd.to_datetime(df["Vencimento"], format="%d/%m/%Y", errors='coerce')
        df.sort_values("Vencimento_dt", inplace=True)
        df.drop(columns="Vencimento_dt", inplace=True)

        colunas_ordenadas = ["Local", "Setor", "Cadastro", "Colaborador", "Autorização", "Vencimento", "E-mail Particular"]
        df[["Local", "Setor"]] = df["Setor"].apply(lambda codloc: pd.Series(self.buscar_local_setor(codloc)))
        df = df[colunas_ordenadas]

        colunas = list(df.columns)
        self.tableWidget.setColumnCount(len(colunas))
        self.tableWidget.setRowCount(len(df))
        self.tableWidget.setHorizontalHeaderLabels(colunas)

        fonte = self.tableWidget.font()
        fonte.setPointSize(8)
        self.tableWidget.setFont(fonte)

        hoje = datetime.now()

        for row_idx, row in df.iterrows():
            venc = datetime.strptime(row["Vencimento"], "%d/%m/%Y")
            cor = None
            if venc < hoje:
                cor = Qt.GlobalColor.red
            elif venc < hoje + timedelta(days=180):
                cor = Qt.GlobalColor.yellow

            for col_idx, col in enumerate(colunas):
                item = QTableWidgetItem(str(row[col]) if row[col] is not None else "")
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if cor:
                    item.setBackground(cor)
                    if cor == Qt.GlobalColor.red:
                        item.setForeground(Qt.GlobalColor.white)
                    elif cor == Qt.GlobalColor.yellow:
                        item.setForeground(Qt.GlobalColor.black)

                self.tableWidget.setItem(row_idx, col_idx, item)

        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.resizeRowsToContents()
        self.tableWidget.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.df_cnh = df.copy()
        self.status_label.setText("CNH")

    def consultar_rg(self):
        query = """
        SELECT 
            f.NumEmp,
            H.CODLOC,
            f.NUMCAD,
            f.NomFun, 
            TO_CHAR(c.DexCid, 'DD/MM/YYYY') AS DexCid,
            c.EmaPar
        FROM R034FUN f
        JOIN R034CPL c 
            ON f.NumEmp = c.NumEmp 
            AND f.TipCol = c.TipCol 
            AND f.NumCad = c.NumCad
        JOIN R016HIE H ON f.NUMLOC = H.NUMLOC
        JOIN R034CPL c 
            ON f.NumEmp = c.NumEmp 
            AND f.TipCol = c.TipCol 
            AND f.NumCad = c.NumCad
        WHERE f.NumEmp IN (10, 16, 17, 19, 11)
        AND f.TipCol IN (1, 2)
        AND f.SitAfa = '001'
        AND c.DexCid IS NOT NULL 
        AND c.DexCid > DATE '0001-01-01'
        ORDER BY c.DexCid ASC
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
            self.status_label.setText("Consulta RG/CIN — nenhum resultado encontrado.")
            return

        df = df.map(lambda x: x.replace("\n", " ").replace("\r", " ").strip() if isinstance(x, str) else x)
        df.fillna("", inplace=True)

        df.rename(columns={
            "NUMEMP": "Local",
            "CODLOC": "Setor",
            "NUMCAD": "Cadastro",
            "NOMFUN": "Colaborador",
            "DEXCID": "DexCid",
            "EMAPAR": "E-mail Particular"
        }, inplace=True)

        def calcular_vencimento(data_str):
            if data_str == "31/12/1900":
                return "Não Preenchido"
            try:
                data = datetime.strptime(data_str, "%d/%m/%Y")
                return (data + timedelta(days=365 * 10)).strftime("%d/%m/%Y")
            except:
                return "Não Preenchido"

        df["Vencimento"] = df["DexCid"].apply(calcular_vencimento)

        df["Vencimento_dt"] = pd.to_datetime(
            df["Vencimento"].replace("Não Preenchido", pd.NaT), format="%d/%m/%Y", errors='coerce'
        )
        df.sort_values("Vencimento_dt", inplace=True)
        df.drop(columns=["DexCid", "Vencimento_dt"], inplace=True)

        colunas_ordenadas = ["Local", "Setor", "Cadastro", "Colaborador", "Vencimento", "E-mail Particular"]
        df[["Local", "Setor"]] = df["Setor"].apply(lambda codloc: pd.Series(self.buscar_local_setor(codloc)))
        df = df[colunas_ordenadas]

        colunas = list(df.columns)
        self.tableWidget.setColumnCount(len(colunas))
        self.tableWidget.setRowCount(len(df))
        self.tableWidget.setHorizontalHeaderLabels(colunas)

        fonte = self.tableWidget.font()
        fonte.setPointSize(8)
        self.tableWidget.setFont(fonte)

        hoje = datetime.now()

        for row_idx, row in df.iterrows():
            vencimento_str = row["Vencimento"]
            cor = None

            if vencimento_str == "Não Preenchido":
                cor = Qt.GlobalColor.red
            else:
                try:
                    venc = datetime.strptime(vencimento_str, "%d/%m/%Y")
                    if venc < hoje:
                        cor = Qt.GlobalColor.red
                    elif venc < hoje + timedelta(days=180):
                        cor = Qt.GlobalColor.yellow
                except:
                    cor = Qt.GlobalColor.red

            for col_idx, col in enumerate(colunas):
                item = QTableWidgetItem(str(row[col]) if row[col] is not None else "")
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if cor:
                    item.setBackground(cor)
                    if cor == Qt.GlobalColor.red:
                        item.setForeground(Qt.GlobalColor.white)
                    elif cor == Qt.GlobalColor.yellow:
                        item.setForeground(Qt.GlobalColor.black)

                self.tableWidget.setItem(row_idx, col_idx, item)

        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.resizeRowsToContents()
        self.tableWidget.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.df_rg = df.copy()
        self.status_label.setText("RG/CIN")

    def enviar_email(self):
        status = self.status_label.text().lower()

        if "cnh" in status:
            df_base = self.df_cnh
            tipo_doc = "CNH"
        elif "rg" in status or "cin" in status:
            df_base = self.df_rg
            tipo_doc = "RG/CIN"
        else:
            QMessageBox.warning(self, "Erro", "Nenhuma consulta foi realizada ainda.")
            return

        if df_base.empty or "Vencimento" not in df_base.columns:
            QMessageBox.warning(self, "Erro", f"Nenhum dado disponível para envio ({tipo_doc}).")
            return

        #  Gerar DataFrame com apenas as linhas visíveis
        linhas_visiveis = []
        for row in range(self.tableWidget.rowCount()):
            if not self.tableWidget.isRowHidden(row):
                linha = {}
                for col in range(self.tableWidget.columnCount()):
                    col_name = self.tableWidget.horizontalHeaderItem(col).text()
                    item = self.tableWidget.item(row, col)
                    linha[col_name] = item.text() if item else ""
                linhas_visiveis.append(linha)

        df = pd.DataFrame(linhas_visiveis)

        if df.empty:
            QMessageBox.information(self, "Aviso", "Nenhum dado visível para enviar no e-mail.")
            return

        hoje = datetime.now()

        def verificar_vencimento(v):
            try:
                data = datetime.strptime(v, "%d/%m/%Y")
                if data < hoje:
                    return "vencido"
                elif data < hoje + timedelta(days=180):
                    return "proximo"
            except:
                return "vencido"
            return None

        df["StatusVencimento"] = df["Vencimento"].apply(verificar_vencimento)
        df_filtrado = df[df["StatusVencimento"].isin(["vencido", "proximo"])].copy()

        if df_filtrado.empty:
            QMessageBox.information(self, "Aviso", f"Não há {tipo_doc}s vencidas ou próximas do vencimento.")
            return

        #  Cores de fundo por local
        cores_locais = {
            "Fabrica De Máquinas": "#003756",
            "Fabrica De Transportadores": "#ffc62e",
            "Adm": "#009c44",
            "Comercial": "#919191"
        }

        corpo_email = f"<span style='font-size:16px;'>Segue colaboradores com {tipo_doc}s vencidas ou próximas ao vencimento.</span><br><br>"
        corpo_email += "<table border='1' cellspacing='0' cellpadding='3' style='border-collapse: collapse; width: 100%; font-size: 14px; text-align: center;'>"
        corpo_email += """
        <tr style="background-color: #f2f2f2;">
            <th>Local</th>
            <th>Setor</th>
            <th>Colaborador</th>
            <th>Vencimento</th>
        </tr>
        """

        for (local, setor), grupo in df_filtrado.groupby(["Local", "Setor"]):
            cor_local = cores_locais.get(local, "#000000")
            for i, row in grupo.iterrows():
                cor_venc = ""
                if row["StatusVencimento"] == "vencido":
                    cor_venc = " style='background-color:red; color:white;'"
                elif row["StatusVencimento"] == "proximo":
                    cor_venc = " style='background-color:yellow;'"

                corpo_email += "<tr>"
                if i == grupo.index[0]:
                    corpo_email += f"<td rowspan='{len(grupo)}' style='background-color:{cor_local}; color:white;'>{local}</td>"
                    corpo_email += f"<td rowspan='{len(grupo)}'>{setor}</td>"
                corpo_email += f"<td>{row['Colaborador']}</td><td{cor_venc}>{row['Vencimento']}</td></tr>"

        corpo_email += "</table>"

        outlook = win32com.client.Dispatch("Outlook.Application")
        email = outlook.CreateItem(0)
        email.Subject = f" Relatório Documentos Vencidos ou Próximo ao Vencimento - {tipo_doc}"
        email.HTMLBody = corpo_email + "<br><br>" + email.HTMLBody

        self.current_mail = email
        email.Display()

        QMessageBox.information(self, "Sucesso", f"E-mail  gerado com sucesso no Outlook!")

    def enviar_email_colaboradores(self):

        status = self.status_label.text().lower()

        if "cnh" in status:
            df_base = self.df_cnh
            tipo_doc = "CNH"
        elif "rg" in status or "cin" in status:
            df_base = self.df_rg
            tipo_doc = "RG/CIN"
        else:
            QMessageBox.warning(self, "Erro", "Nenhuma consulta foi realizada ainda.")
            return

        if df_base.empty or "Vencimento" not in df_base.columns:
            QMessageBox.warning(self, "Erro", f"Nenhum dado disponível para envio ({tipo_doc}).")
            return

        # 🔍 Apenas linhas visíveis
        linhas_visiveis = []
        for row in range(self.tableWidget.rowCount()):
            if not self.tableWidget.isRowHidden(row):
                linha = {}
                for col in range(self.tableWidget.columnCount()):
                    col_name = self.tableWidget.horizontalHeaderItem(col).text()
                    item = self.tableWidget.item(row, col)
                    linha[col_name] = item.text() if item else ""
                linhas_visiveis.append(linha)

        df = pd.DataFrame(linhas_visiveis)

        if df.empty:
            QMessageBox.information(self, "Aviso", "Nenhum dado visível para enviar.")
            return

        hoje = datetime.now()

        def vencido_ou_proximo(v):
            try:
                if v.strip() in ("Não Preenchido", "31/12/1900"):
                    return True
                data = datetime.strptime(v, "%d/%m/%Y")
                return data < hoje + timedelta(days=180)
            except:
                return True

        df_filtrado = df[df["Vencimento"].apply(vencido_ou_proximo)].copy()

        if df_filtrado.empty:
            QMessageBox.information(self, "Aviso",
                                    f"Não há {tipo_doc}s vencidas ou próximas para os colaboradores de teste.")
            return

        msg = QMessageBox(self)
        msg.setWindowTitle("Confirmação")
        msg.setText(
            f"Deseja enviar e-mails para {len(df_filtrado)} colaborador(es) com documentação de {tipo_doc} vencida ou próxima?")
        btn_sim = msg.addButton("Sim", QMessageBox.ButtonRole.YesRole)
        btn_nao = msg.addButton("Não", QMessageBox.ButtonRole.NoRole)
        msg.exec()

        if msg.clickedButton() != btn_sim:
            return


        
        smtp_server = "mail.bruno.com.br"
        smtp_port = 465
        # Obtém o usuário logado no sistema e formata o e-mail
        usuario_logado = os.getlogin().lower()  
        smtp_user = f"{usuario_logado}@bruno.com.br"
        
        # Solicita a senha através de uma janela de diálogo
        senha_input, ok = QInputDialog.getText(
            self, "Autenticação de E-mail", 
            f"Digite a senha para {smtp_user}:", 
            QLineEdit.EchoMode.Password
        )
        
        if not ok:
            QMessageBox.warning(self, "Cancelado", "Envio de e-mail cancelado pelo usuário.")
            return
            
        smtp_password = senha_input

        enviados = 0
        for _, row in df_filtrado.iterrows():
            email_dest = row.get("E-mail Particular", "").strip()
            if not email_dest:
                continue

            colaborador = row["Colaborador"]
            vencimento = row["Vencimento"]

            msg = EmailMessage()
            msg['Subject'] = f"Documentação de {tipo_doc} vencida ou próxima - RH Bruno"
            msg['From'] = smtp_user
            msg['To'] = email_dest

            corpo = f"""
    Olá {colaborador},

    O RH da Bruno informa que sua documentação de {tipo_doc} encontra-se vencida ou próxima da data de vencimento ({vencimento}).

    Solicitamos que compareça ao RH o mais breve possível e apresente o documento renovado para atualização.

    Atenciosamente,
    Recursos Humanos
            """.strip()

            msg.set_content(corpo)

            try:
                with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                    server.login(smtp_user, smtp_password)
                    server.send_message(msg)
                    enviados += 1
            except Exception as e:
                print(f"Erro ao enviar para {colaborador}: {e}")

        QMessageBox.information(self, "Envio Concluído",
                                f"E-mails enviados com sucesso para {enviados} colaborador(es).")

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


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = AppConsultaDocumentos()
    window.show()
    sys.exit(app.exec())