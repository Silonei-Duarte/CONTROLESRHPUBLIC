import locale
import sys
from PyQt6.QtCore import QCoreApplication, Qt

QCoreApplication.setAttribute(Qt.ApplicationAttribute.AA_ShareOpenGLContexts)

from PyQt6.QtWidgets import (
    QApplication, QWidget, QHBoxLayout, QPushButton, QLineEdit, QMessageBox, QTabWidget
)
import pandas as pd
import json
from PyQt6.QtWidgets import QVBoxLayout, QTableWidget, QTableWidgetItem, QLabel, QSizePolicy, QSpacerItem
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon
import os
from PyQt6.QtWidgets import QLineEdit
from Database import get_connection
from horaextragrafico import AbaGraficos  # Importar a classe para a aba de gráficos


class AppConsultaEventos(QWidget):
    def __init__(self):
        super().__init__()
        self.df_cnh = pd.DataFrame()
        self.df_rg = pd.DataFrame()
        self.df_original = None
        self.df_completo = None

        # mover para cá
        self.tabs = QTabWidget()
        self.resultados_tab = QWidget()
        self.graficos_tab = None

        self.carregar_dicionario_locais_com_busca()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Consulta de Documentos")
        self.resize(1900, 900)

        # Configurar o ícone da janela
        icon_path = os.path.join(os.path.dirname(__file__), 'icone.ico')  # Ajuste o caminho conforme necessário
        if os.path.exists(icon_path):  # Verifica se o arquivo existe
            self.setWindowIcon(QIcon(icon_path))

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        botoes_layout = QHBoxLayout()
        botoes_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)

        label_eventos = QLabel("Eventos:")
        label_eventos.setAlignment(Qt.AlignmentFlag.AlignVCenter)
        botoes_layout.addWidget(label_eventos)

        self.eventos_field = QLineEdit("34,35,36,37,52,65")
        self.eventos_field.setFixedWidth(320)
        self.eventos_field.setPlaceholderText("Códigos de evento separados por vírgula (ex: 34,35,36)")
        botoes_layout.addWidget(self.eventos_field)

        btn_cnh = QPushButton("Consultar Eventos")
        btn_cnh.setFixedSize(140, 30)
        btn_cnh.clicked.connect(self.consultar_eventos)
        botoes_layout.addWidget(btn_cnh)

        #  Adiciona espaço flexível antes do botão "Voltar ao Menu"
        spacer = QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
        botoes_layout.addItem(spacer)

        btn_email = QPushButton("Gerar E-mail")
        btn_email.setFixedSize(140, 30)
        btn_email.clicked.connect(self.generate_email_eventos)
        botoes_layout.addWidget(btn_email)

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
        resultados_layout = QVBoxLayout()
        resultados_layout.addWidget(self.tableWidget)
        self.resultados_tab.setLayout(resultados_layout)

        self.tabs.addTab(self.resultados_tab, "Resultados")
        layout.addWidget(self.tabs)

        self.setLayout(layout)
        self.current_mail = None

    def create_graficos_tab(self):
        """Recria a aba de gráficos para garantir que os dados sejam sempre atualizados."""
        if self.graficos_tab:
            index = self.tabs.indexOf(self.graficos_tab)
            if index != -1:
                self.tabs.removeTab(index)

        self.graficos_tab = AbaGraficos(self.df_eventos)
        self.tabs.addTab(self.graficos_tab, "Gráficos")

    def consultar_eventos(self):
        eventos_texto = self.eventos_field.text().strip()
        if not eventos_texto:
            self.status_label.setText("Informe ao menos um código de evento.")
            return

        try:
            codigos_evento = [int(c.strip()) for c in eventos_texto.split(",") if c.strip().isdigit()]
        except ValueError:
            self.status_label.setText("Erro: informe apenas números separados por vírgula.")
            return

        if not codigos_evento:
            self.status_label.setText("Nenhum código de evento válido informado.")
            return

        placeholders = ",".join(str(c) for c in codigos_evento)

        query = f"""
        SELECT 
            V.NUMEMP,
            H.CODLOC,
            V.TIPCOL,
            V.NUMCAD,
            F.NOMFUN,
            V.CODEVE,
            TO_CHAR(V.VALEVE, 'FM99999990.00') AS VALEVE,
            TO_CHAR(C.PERREF, 'DD/MM/YYYY') AS PERREF
        FROM 
            R046VER V
        JOIN 
            R044CAL C ON C.CODCAL = V.CODCAL 
                     AND C.NUMEMP = V.NUMEMP
        JOIN 
            R034FUN F ON F.TIPCOL = V.TIPCOL 
                     AND F.NUMEMP = V.NUMEMP 
                     AND F.NUMCAD = V.NUMCAD
        JOIN 
            R016HIE H ON F.NUMLOC = H.NUMLOC
        LEFT JOIN 
            R038AFA A ON A.NUMEMP = F.NUMEMP
                     AND A.TIPCOL = F.TIPCOL
                     AND A.NUMCAD = F.NUMCAD
                     AND A.SITAFA = '007'
        WHERE 
            V.TIPCOL = 1
          AND V.NUMEMP IN (10, 16, 17, 19)
          AND V.CODEVE IN ({placeholders})
          AND C.PERREF = TRUNC(ADD_MONTHS(SYSDATE, -1), 'MM')
          AND TRUNC(C.DATPAG, 'MM') = ADD_MONTHS(TRUNC(C.PERREF, 'MM'), 1)
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
            self.status_label.setText("Eventos — nenhum resultado encontrado.")
            return

        df = df.map(lambda x: x.replace("\n", " ").replace("\r", " ").strip() if isinstance(x, str) else x)
        df.fillna("", inplace=True)

        df.rename(columns={
            "NUMEMP": "Local",
            "CODLOC": "Setor",
            "NUMCAD": "Cadastro",
            "NOMFUN": "Colaborador",
            "CODEVE": "Evento",
            "VALEVE": "Valor R$",
            "PERREF": "Referência Calculo"
        }, inplace=True)

        # Mapeia os códigos de evento para os nomes descritivos
        df["Evento"] = df["Evento"].replace({
            1: "Horas Normais",
            2: "Horas Normais Noturnas",
            4: "Horas Repouso Rem.Diurno",
            6: "Horas Repouso Rem.Noturno",
            8: "Horas Just.Diurnas",
            10: "Horas Just.Noturnas",
            11: "Horas Viagem",
            12: "Horas Férias Diurnas",
            13: "Horas Lic.Remuner.Noturna",
            14: "Horas Férias Noturnas",
            15: "Horas Lic.Remuner.Diurna",
            18: "Horas Faltas",
            19: "Horas Faltas Noturnas",
            22: "Horas Faltas DSR",
            23: "Horas Faltas DSR Noturno",
            26: "Horas Acidente Trabalho",
            28: "Horas Auxílio Doença",
            34: "Horas Extras c/ 50%",
            35: "Horas Extras c/50% Noturn",
            36: "Horas Extras c/ 100%",
            37: "Horas Extras c/100% Notur",
            52: "Horas Inden.Inter c/50%",
            56: "Horas Atestado Medico",
            57: "Horas Atestado Noturnas",
            58: "Horas Licença Paternidade",
            59: "Horas Licença Patern.Not.",
            60: "Adicional Noturno",
            64: "Periculosidade",
            65: "DSR Reflexo H.Extras",
            70: "Adic.Noturno s/ Férias",
            88: "Hrs.Atest.sem INSS Diurno",
            100: "Aviso Prévio Indenizado",
            102: "Média Horas Extras A.P.I.",
            110: "Adic.Noturno A.P.I",
            112: "Aviso Prévio Reavido",
            126: "Saldo de Salário",
            134: "Média Horas Extras Férias",
            140: "1/3 Férias",
            142: "Diferença de Férias",
            148: "Abono Pecuniário Férias",
            150: "Média H.Extras Abono Pec.",
            158: "Adic.Noturno Abono Pec.",
            160: "1/3 s/ Abono Pecuniário",
            162: "Diferença de Abono",
            170: "Férias Vencidas Rescisäo",
            172: "Férias Proporc.Rescisäo",
            174: "Média H.Extra Férias Resc",
            177: "Ind.Térm.Contr.Antec.Colb",
            180: "Peric.Férias Rescisäo",
            182: "Adic.Noturno Férias Resc.",
            184: "1/3 Férias Rescisäo",
            222: "13º Salário Proporc.Resc.",
            224: "Média H.Extras 13º Prop.",
            230: "Periculosidade 13º Prop.",
            232: "Adic.Noturno 13º Prop.",
            236: "13º Indenizado Rescisäo",
            238: "Média H.Extras 13º Inden.",
            250: "Desc.Adto Salarial",
            256: "Estouro do Mês",
            258: "Estouro Mês Anterior",
            264: "Líquido Rescisäo",
            276: "Pensão Judicial",
            281: "Desconto Adto Férias",
            292: "Devolução de Convênios",
            295: "Pró-Labore",
            300: "FGTS",
            301: "INSS s/ Férias",
            302: "INSS",
            303: "INSS s/ 13º Salário",
            304: "IRRF",
            306: "IRRF s/ 13º Salário",
            308: "IRRF s/ Férias",
            310: "IRRF s/ Adto Salarial",
            393: "FGTS 13º Salário",
            400: "Plano VIVO Celulares",
            401: "Outros Descontos",
            444: "Restaurantes",
            454: "FGTS 40% Rescisäo GRFP",
            455: "FGTS 13ºSal.Rescisäo GRFP",
            456: "FGTS Rescisäo GRFP",
            513: "Associação Oscar Bruno",
            570: "Jaquetas/Blusas/Camisetas",
            577: "Emprestimo Banco Bradesco",
            581: "Emprestimo Sicoob",
            591: "Pensão Vitalícia",
            594: "Pgto Art. 62 Inciso II",
            599: "Horas Faltas SD",
            600: "Associação Oscar Bruno",
            632: "Diferença de Remuneração",
            636: "Cont.Cota Part. Negocial",
            703: "Farmacias",
            704: "DSR Indenizado",
            707: "Seguro de Vida - Tokio",
            711: "Horas Sobre Aviso",
            713: "Unimed Dependentes",
            714: "Unimed Empresa",
            715: "Unimed Coparticipaçao",
            718: "Seguro Empresa BI",
            719: "Util/Ticket Alimentação",
            720: "Refeitorio Empresa",
            721: "Plano Odontologico",
            722: "Transporte",
            727: "Associação Consumo FACISC",
            728: "Descontos Unimed Parcel.",
            729: "1º Desc.Cred. Trabalhador",
            731: "Gratificação de Turno",
            821: "Horas Faltas BH",
            834: "Horas Extras c/ 50% BH",
            836: "Horas Extras c/ 100% BH"
        })

        df[["Local", "Setor"]] = df["Setor"].apply(lambda codloc: pd.Series(self.buscar_local_setor(codloc)))

        colunas_ordenadas = ["Local", "Setor", "Cadastro", "Colaborador", "Evento", "Valor R$", "Referência Calculo"]
        df = df[colunas_ordenadas]

        colunas = list(df.columns)
        self.tableWidget.setColumnCount(len(colunas))
        self.tableWidget.setRowCount(len(df))
        self.tableWidget.setHorizontalHeaderLabels(colunas)

        fonte = self.tableWidget.font()
        fonte.setPointSize(8)
        self.tableWidget.setFont(fonte)

        for row_idx, row in df.iterrows():
            for col_idx, col in enumerate(colunas):
                item = QTableWidgetItem(str(row[col]) if row[col] is not None else "")
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tableWidget.setItem(row_idx, col_idx, item)

        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.resizeRowsToContents()
        self.tableWidget.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.df_eventos = df.copy()
        self.status_label.setText("Eventos carregados.")
        self.atualizar_linha_total()

        self.df_eventos = pd.DataFrame(columns=[
            "local", "setor", "Nome", "Código", "Rotina", "Horas",
            "Salário Base", "Salário por Hora", "Valor Final"
        ])

        for _, row in df.iterrows():
            self.df_eventos.loc[len(self.df_eventos)] = [
                row["Local"],
                row["Setor"],
                row["Colaborador"],
                row["Cadastro"],
                row["Evento"],
                "00:00:00",  # Horas
                0,  # Salário Base
                0,  # Salário por Hora
                float(row["Valor R$"].replace(",", ".")) if row["Valor R$"] else 0  # Valor Final
            ]

        self.create_graficos_tab()

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
        self.atualizar_linha_total()

    def atualizar_linha_total(self):
        colunas = [self.tableWidget.horizontalHeaderItem(i).text() for i in range(self.tableWidget.columnCount())]
        if "Valor R$" not in colunas:
            return

        idx_valor = colunas.index("Valor R$")
        total = 0.0

        for row in range(self.tableWidget.rowCount()):
            if not self.tableWidget.isRowHidden(row):
                item = self.tableWidget.item(row, idx_valor)
                if item:
                    try:
                        total += float(item.text().replace(",", "").replace("R$", "").strip())
                    except ValueError:
                        continue

        # Remove linha total anterior, se existir
        if self.tableWidget.rowCount() > 0 and self.tableWidget.item(self.tableWidget.rowCount() - 1,
                                                                     0).text() == "Total":
            self.tableWidget.removeRow(self.tableWidget.rowCount() - 1)

        row_idx = self.tableWidget.rowCount()
        self.tableWidget.insertRow(row_idx)

        for col_idx in range(self.tableWidget.columnCount()):
            texto = "Total" if col_idx == 0 else ""
            if colunas[col_idx] == "Valor R$":
                texto = f"{total:,.2f}".replace(".", "@").replace(",", ".").replace("@", ",")  # formato pt-BR

            item = QTableWidgetItem(texto)
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item.setBackground(Qt.GlobalColor.yellow)
            self.tableWidget.setItem(row_idx, col_idx, item)
            self.tableWidget.resizeColumnsToContents()

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

    def generate_email_eventos(self):
        """Gera um e-mail com os dados filtrados atualmente na tabela de eventos (sem horas)."""
        try:
            # Verifica se há dados visíveis e ignora linha "Total"
            linhas_visiveis = [
                row for row in range(self.tableWidget.rowCount())
                if not self.tableWidget.isRowHidden(row)
                   and self.tableWidget.item(row, 0).text().strip().upper() != "TOTAL"
            ]
            if not linhas_visiveis:
                QMessageBox.warning(self, "Aviso", "Nenhum dado disponível para gerar o e-mail.")
                return

            # Definir cores dos locais
            cores_locais = {
                "Fabrica De Máquinas": "#003756",
                "Fabrica De Transportadores": "#ffc62e",
                "Adm": "#009c44",
                "Comercial": "#919191"
            }

            agrupados = {}
            total_valor = 0.0

            for row in linhas_visiveis:
                local = self.tableWidget.item(row, 0).text().strip() or "Sem local"
                setor = self.tableWidget.item(row, 1).text().strip() or "Sem setor"
                nome = self.tableWidget.item(row, 3).text().strip() or "Sem nome"
                locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

                valor_texto = self.tableWidget.item(row, 5).text().strip().replace("R$", "").strip()
                try:
                    valor = float(valor_texto.replace(",", "."))
                except ValueError:
                    valor = 0.0

                total_valor += valor
                agrupados.setdefault(local, {}).setdefault(setor, {}).setdefault(nome, 0.0)
                agrupados[local][setor][nome] += valor

            html = """
            <html>
            <head>
                <style>
                    table {
                        border-collapse: collapse;
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
                <p>Controle de eventos agrupado por local e setor.</p>
                <table>
                    <tr>
                        <th>Local</th>
                        <th>Setor</th>
                        <th>Nome</th>
                        <th>Valor R$</th>
                    </tr>
            """

            for local, setores in agrupados.items():
                cor = cores_locais.get(local, "#000000")
                for setor, funcionarios in setores.items():
                    primeira = True
                    qtd = len(funcionarios)
                    for nome, valor in funcionarios.items():
                        html += "<tr>"
                        html += f'<td style="background-color:{cor};color:#fff;font-weight:bold;">{local}</td>'
                        if primeira:
                            html += f'<td rowspan="{qtd}">{setor}</td>'
                            primeira = False
                        html += f"<td>{nome}</td>"
                        html += f"<td>R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") + "</td>"
                        html += "</tr>"

            html += f"""
                <tr>
                    <td colspan="3" style="font-weight:bold; text-align:center; background-color:#f2f2f2;">TOTAL GERAL</td>
                    <td style="text-align:center;">R$ {total_valor:,.2f}</td>
                </tr>
                </table>
            </body>
            </html>
            """

            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application")
            self.current_mail = outlook.CreateItem(0)
            self.current_mail.Subject = "Relatório de Eventos"
            signature = self.current_mail.HTMLBody
            self.current_mail.HTMLBody = html + "<br><br>" + signature
            self.current_mail.Display()

            # Atualizar a aba de gráficos, se existir
            if self.graficos_tab:
                self.graficos_tab.set_current_email(self.current_mail)


            QMessageBox.information(self, "Sucesso", "E-mail gerado com sucesso!")

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao gerar e-mail: {e}")

    def voltar_menu(self):
        from main import ControlesRH
        self.menu_principal = ControlesRH()
        self.menu_principal.show()
        self.window().close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = AppConsultaEventos()
    window.show()
    sys.exit(app.exec())