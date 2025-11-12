import pandas as pd
from PyQt6.QtWidgets import QTableWidgetItem, QDialog, QVBoxLayout, QLabel, QTableWidget
from PyQt6.QtCore import Qt
from Database import get_connection


class PainelSetores:
    def __init__(self, parent):
        self.parent = parent

    def calcular_media_colaboradores_30dias(self, data_ref_str: str):
        query = f"""
        SELECT 
            FUNC.NUMEMP,
            H.CODLOC,
            FUNC.TIPCOL,
            FUNC.NUMCAD,
            FUNC.NOMFUN,
            dias.DATA_DIA
        FROM (
            SELECT TO_DATE('{data_ref_str}', 'DD/MM/YYYY') - 29 + LEVEL - 1 AS DATA_DIA
            FROM DUAL
            CONNECT BY LEVEL <= 30
        ) dias
        JOIN R034CPL C ON 1 = 1
        JOIN R034FUN FUNC ON FUNC.NUMEMP = C.NUMEMP 
                        AND FUNC.TIPCOL = C.TIPCOL 
                        AND FUNC.NUMCAD = C.NUMCAD
        JOIN R016HIE H ON FUNC.NUMLOC = H.NUMLOC
        LEFT JOIN R038AFA A ON A.NUMEMP = C.NUMEMP
                           AND A.TIPCOL = C.TIPCOL
                           AND A.NUMCAD = C.NUMCAD
                           AND A.SITAFA = '007'
        WHERE 
            C.NUMEMP IN (10, 16, 17, 19, 11)
            AND FUNC.DATADM <= dias.DATA_DIA
            AND (A.DATAFA IS NULL OR A.DATAFA >= dias.DATA_DIA)
            AND FUNC.SITAFA NOT IN ('003', '024', '913')
        """

        with get_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(query)
                colunas = [desc[0] for desc in cursor.description]
                registros = cursor.fetchall()

        df = pd.DataFrame(registros, columns=colunas)

        if df.empty:
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

        if hasattr(self.parent, 'nomes_filtrados') and self.parent.nomes_filtrados:
            df = df[~df["NOMFUN"].isin(self.parent.nomes_filtrados)]

        df.rename(columns={
            "NUMEMP": "Empresa",
            "CODLOC": "Setor Funcionário",
            "TIPCOL": "Tipo",
            "NUMCAD": "Cadastro",
            "NOMFUN": "Colaborador",
            "DATA_DIA": "Data consulta"
        }, inplace=True)

        df[["Local", "Setor"]] = df["Setor Funcionário"].apply(
            lambda codloc: pd.Series(self.parent.buscar_local_setor(codloc))
        )

        df = df[["Local", "Setor", "Tipo", "Cadastro", "Colaborador", "Data consulta"]]

        df_por_dia = df.groupby(["Local", "Setor", "Data consulta"]).size().reset_index(name="QtdColaboradores")
        df_media_setor = df_por_dia.groupby(["Local", "Setor"])["QtdColaboradores"].sum().reset_index()
        df_media_setor["Média Colab. no Intervalo"] = (df_media_setor["QtdColaboradores"] / 30).round(2)
        df_media_setor.drop(columns="QtdColaboradores", inplace=True)

        df_por_local_dia = df.groupby(["Local", "Data consulta"]).size().reset_index(name="QtdColaboradores")
        df_media_local = df_por_local_dia.groupby("Local")["QtdColaboradores"].sum().reset_index()
        df_media_local["Média Colab. no Intervalo"] = (df_media_local["QtdColaboradores"] / 30).round(2)
        df_media_local.drop(columns="QtdColaboradores", inplace=True)

        return df_por_dia, df_media_setor, df_media_local

    def atualizar_tabela_setores(self):
        if self.parent.df_ativos.empty:
            return

        data_str = self.parent.date_fim_picker.date().toString("dd/MM/yyyy")
        data_ref = pd.to_datetime(data_str, format="%d/%m/%Y").normalize()
        dias = int(self.parent.dias_desligados.text())
        limite_inferior = data_ref - pd.Timedelta(days=dias)

        df_por_dia, df_media_setor, df_media_local = self.calcular_media_colaboradores_30dias(data_str)

        resumo_setores = []
        resumo_locais = {}
        locais_setores = set()
        todos_locais = set()

        for _, row in self.parent.df_ativos.iterrows():
            locais_setores.add((row["Local"], row["Setor"]))
            todos_locais.add(row["Local"])

        for _, row in self.parent.df_desligados.iterrows():
            locais_setores.add((row["Local"], row["Setor"]))
            todos_locais.add(row["Local"])

        for local in todos_locais:
            resumo_locais[local] = {
                "Local": local,
                "Setor": "Todos",
                "Quantidade de Colaboradores": 0,
                "Admitidos no Intervalo": 0,
                "Desligados no Intervalo": 0,
                "Média Colab. no Intervalo": 0,
                "Turnover (%)": "0.00%"
            }

        for local, setor in locais_setores:
            ativos = self.parent.df_ativos[(self.parent.df_ativos["Local"] == local) & (self.parent.df_ativos["Setor"] == setor)]
            desligados = self.parent.df_desligados[(self.parent.df_desligados["Local"] == local) & (self.parent.df_desligados["Setor"] == setor)]

            qtd = len(ativos)
            admitidos = pd.to_datetime(ativos["Data Admissão"], format="%d/%m/%Y", errors="coerce").between(limite_inferior, data_ref).sum()
            Desligados = pd.to_datetime(desligados["Data Desligamento"], format="%d/%m/%Y", errors="coerce").between(limite_inferior, data_ref).sum()

            media = df_media_setor[(df_media_setor["Local"] == local) & (df_media_setor["Setor"] == setor)]
            media_valor = media["Média Colab. no Intervalo"].iloc[0] if not media.empty else 0
            turnover = (Desligados / media_valor * 100) if media_valor > 0 else 0

            resumo_setores.append({
                "Local": local,
                "Setor": setor,
                "Quantidade de Colaboradores": qtd,
                "Admitidos no Intervalo": int(admitidos),
                "Desligados no Intervalo": int(Desligados),
                "Média Colab. no Intervalo": round(media_valor, 2),
                "Turnover (%)": f"{turnover:.2f}%"
            })

            resumo_locais[local]["Quantidade de Colaboradores"] += qtd
            resumo_locais[local]["Admitidos no Intervalo"] += int(admitidos)
            resumo_locais[local]["Desligados no Intervalo"] += int(Desligados)

        for local, resumo in resumo_locais.items():
            linha = df_media_local[df_media_local["Local"] == local]
            media = linha["Média Colab. no Intervalo"].iloc[0] if not linha.empty else 0
            resumo["Média Colab. no Intervalo"] = round(media, 2)
            if media > 0:
                turnover = (resumo["Desligados no Intervalo"] / media) * 100
                resumo["Turnover (%)"] = f"{turnover:.2f}%"

        resumo_setores = sorted(resumo_setores, key=lambda x: (x["Local"], "" if x["Setor"] == "Todos" else x["Setor"]))
        resumo_final = [resumo_locais[local] for local in sorted(resumo_locais.keys())] + resumo_setores

        self.parent.setores_table.clearContents()

        if not resumo_final:
            self.parent.setores_table.setRowCount(0)
            self.parent.setores_table.setColumnCount(0)
            return

        if not self.parent.sinal_setores_conectado:
            self.parent.setores_table.cellDoubleClicked.connect(self.mostrar_detalhes_setor)
            self.parent.sinal_setores_conectado = True

        colunas = ["Local", "Setor", "Quantidade de Colaboradores",
                   "Admitidos no Intervalo", "Desligados no Intervalo",
                   "Média Colab. no Intervalo", "Turnover (%)"]

        self.parent.setores_table.setColumnCount(len(colunas))
        self.parent.setores_table.setHorizontalHeaderLabels(colunas)
        self.parent.setores_table.setRowCount(len(resumo_final))
        fonte = self.parent.setores_table.font()
        fonte.setPointSize(8)
        self.parent.setores_table.setFont(fonte)

        for row_idx, info in enumerate(resumo_final):
            destacar = info["Setor"] == "Todos"
            for col_idx, coluna in enumerate(colunas):
                item = QTableWidgetItem(str(info[coluna]))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if destacar:
                    item.setBackground(Qt.GlobalColor.lightGray)
                    font_item = item.font()
                    font_item.setBold(True)
                    item.setFont(font_item)
                self.parent.setores_table.setItem(row_idx, col_idx, item)

        self.parent.setores_table.resizeColumnsToContents()
        self.parent.setores_table.resizeRowsToContents()
        self.parent.setores_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

    def mostrar_detalhes_setor(self, row, column):
        if column not in [2, 3, 4]:
            return

        local_item = self.parent.setores_table.item(row, 0)
        setor_item = self.parent.setores_table.item(row, 1)
        if not local_item or not setor_item:
            return

        local = local_item.text()
        setor = setor_item.text()
        filtro_setor = setor != "Todos"

        data_str = self.parent.date_fim_picker.date().toString("dd/MM/yyyy")
        data_ref = pd.to_datetime(data_str, format="%d/%m/%Y").normalize()
        dias = int(self.parent.dias_desligados.text())

        limite_inferior = data_ref - pd.Timedelta(days=dias)

        df_detalhes = None
        if column == 2:
            df_detalhes = self.parent.df_ativos[
                (self.parent.df_ativos["Local"] == local) &
                ((not filtro_setor) | (self.parent.df_ativos["Setor"] == setor))
            ].copy()
            titulo = f"Colaboradores Ativos - {local}"
        elif column == 3:
            df_temp = self.parent.df_ativos.copy()
            df_temp["Data Admissão"] = pd.to_datetime(df_temp["Data Admissão"], format="%d/%m/%Y", errors="coerce")
            df_detalhes = df_temp[
                (df_temp["Local"] == local) &
                ((not filtro_setor) | (df_temp["Setor"] == setor)) &
                (df_temp["Data Admissão"].between(limite_inferior, data_ref))
            ].copy()
            df_detalhes["Data Admissão"] = df_detalhes["Data Admissão"].dt.strftime("%d/%m/%Y")
            titulo = f"Colaboradores Admitidos no Intervalo - {local}"
        elif column == 4:
            df_temp = self.parent.df_desligados.copy()
            df_temp["Data Desligamento"] = pd.to_datetime(df_temp["Data Desligamento"], format="%d/%m/%Y", errors="coerce")
            df_detalhes = df_temp[
                (df_temp["Local"] == local) &
                ((not filtro_setor) | (df_temp["Setor"] == setor)) &
                (df_temp["Data Desligamento"].between(limite_inferior, data_ref))
            ].copy()
            df_detalhes["Data Desligamento"] = df_detalhes["Data Desligamento"].dt.strftime("%d/%m/%Y")
            titulo = f"Colaboradores Desligados no Último Mês - {local}"
        else:
            return

        if filtro_setor:
            titulo += f" - {setor}"

        if df_detalhes is not None and not df_detalhes.empty:
            self.parent.janela_detalhes = JanelaDetalhesColaboradores(df_detalhes, titulo, self.parent)
            self.parent.janela_detalhes.exec()

class JanelaDetalhesColaboradores(QDialog):
    def __init__(self, dataframe, titulo, parent=None):
        super().__init__(parent)
        self.setWindowTitle(titulo)
        self.resize(1500, 600)


        # Configurar layout
        layout = QVBoxLayout(self)

        # Adicionar informação sobre a quantidade
        info_label = QLabel(f"Total de registros: {len(dataframe)}")
        layout.addWidget(info_label)

        # Criar tabela
        self.tabela = QTableWidget()
        layout.addWidget(self.tabela)

        # Configurar tabela
        self.preencher_tabela(dataframe)

        # Configurar janela
        self.setLayout(layout)

    def preencher_tabela(self, df):
        if df.empty:
            return

        colunas = list(df.columns)
        self.tabela.setColumnCount(len(colunas))
        self.tabela.setHorizontalHeaderLabels(colunas)
        self.tabela.setRowCount(len(df))

        fonte = self.tabela.font()
        fonte.setPointSize(8)
        self.tabela.setFont(fonte)

        for row_idx in range(len(df)):
            for col_idx, col in enumerate(colunas):
                valor = df.iloc[row_idx, col_idx]
                item = QTableWidgetItem(str(valor))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tabela.setItem(row_idx, col_idx, item)

        self.tabela.resizeColumnsToContents()
        self.tabela.resizeRowsToContents()
        self.tabela.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        # Oculta os números das linhas
        self.tabela.verticalHeader().setVisible(False)