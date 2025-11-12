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
from hisfrequenciagrafico import AbaGraficoFrequencia

class AbaHistorico(QWidget):
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

        #  Seleção de pasta única
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

        self.email_resumo = QPushButton("E-mail Resumo")
        self.email_resumo.setFixedSize(120, 30)
        self.email_resumo.clicked.connect(self.enviar_email_resumo)
        botoes_layout.addWidget(self.email_resumo)

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

        #  Campo de filtro
        self.search_field = QLineEdit()
        self.search_field.setPlaceholderText("Filtrar usando / entre dados! Tambem pode digitar COLUNA:DADO para filtrar pela expressão exata na coluna")
        self.search_field.textChanged.connect(self.apply_global_filter)
        layout.addWidget(self.search_field)

        #  Tabela para exibir os dados processados
        self.tableWidget = QTableWidget()
        layout.addWidget(self.tableWidget)

        # Evento de duplo clique na tabela
        self.tableWidget.cellDoubleClicked.connect(self.abrir_detalhes_colaborador)
        self.df_original = None  # Guarda a planilha carregada

        # Carrega a configuração ao iniciar
        self.load_config()

        #  Criar o dicionário de locais diretamente na `__init__`
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
        """Carrega a planilha especificada e organiza os dados em uma tabela."""
        if not self.folder_path:
            self.hrap_line_edit.setText("Selecione a pasta primeiro!")
            return

        file_path = os.path.join(self.folder_path, filename)
        if not os.path.exists(file_path):
            self.status_label.setText(f"Arquivo: {filename} não encontrado!")
            self.status_label.setStyleSheet("color: red; font-size: 12px; ")
            self.periodo_label.setText(f"-")
            return

        self.open_save_with_excel(file_path)  # ABRE E FECHA ARQUIVO

        locais_path = os.path.join(self.folder_path, "LOCAIS.xlsx")
        if os.path.exists(locais_path):
            locais_df = pd.read_excel(locais_path, header=None, dtype=str)
            self.dicionario_locais = dict(zip(locais_df[0].str.strip(), locais_df[1].str.strip()))
        else:
            self.dicionario_locais = {}

        df = pd.read_excel(file_path, header=None, dtype=str)
        self.df_original = df.copy()  # Salva uma cópia do DataFrame original

        periodo_inicio = df.iloc[0, 10]
        periodo_fim = df.iloc[0, 11]

        # Primeiro dicionário completo
        mapa_situacoes_completo = {
            "Saída Intermediária",
            "Saída Antecipada",
            "Saída Antecipada Noturna",
            "Saída Intermediária Noturna",
            "Atraso",
            "Atraso Noturno",
            "Faltas Noturnas",
            "Faltas",
            "Atestado Medico",
            "Atestado Medico Noturno",
            "Falta justificada",
            "Falta justificada Noturna",
            "Licença Médica 15 dias",
            "Supensão Disciplinar",
            "Supensão Disciplinar Noturno",
            "Trabalhando",
            "Trabalho Noturno",
            "Ferias",
            "Férias Noturnas",
            "Acidente Trabalho",
            "Acidente Trabalho Noturno",
            "Licença Paternidade",
            "Ferias Coletivas",
            "Férias Coletivas Noturnas",
            "Aviso Previo Trab.",
            "Aviso Prévio Trab. Noturno",
            "Viagem a Servico",
            "Ausencia p/Cursos/Treinamentos",
            "Saída Intermediária BH",
            "Saída Antecipada BH",
            "Saída Antecipada Noturna BH",
            "Saída Intermediária Noturna BH",
            "Atraso BH",
            "Atraso Noturno BH",
            "Faltas Noturnas BH"
        }

        # Segundo dicionário (o que será mantido no código)
        mapa_situacoes = {
            "Saída Intermediária": "Atraso",
            "Saída Antecipada": "Atraso",
            "Saída Antecipada Noturna": "Atraso",
            "Saída Intermediária Noturna": "Atraso",
            "Atraso": "Atraso",
            "Falta justificada": "Faltas",
            "Falta justificada Noturna": "Faltas",
            "Atraso Noturno": "Atraso",
            "Faltas Noturnas": "Faltas",
            "Faltas": "Faltas",
            "Saída Intermediária BH": "Atraso BH",
            "Saída Antecipada BH": "Atraso BH",
            "Saída Antecipada Noturna BH": "Atraso BH",
            "Saída Intermediária Noturna BH": "Atraso BH",
            "Atraso BH": "Atraso BH",
            "Atraso Noturno BH": "Atraso BH",
            "Faltas Noturnas BH": "Faltas BH",
            "Ferias": "Ferias",
            "Férias Noturnas": "Ferias",
            "Supensão Disciplinar": "Faltas",
            "Supensão Disciplinar Noturno": "Faltas",
            "Atestado Medico": "Faltas",
            "Atestado Medico Noturno": "Faltas",
            "Ferias Coletivas": "Ferias",
            "Férias Coletivas Noturnas": "Ferias",
        }

        # Inicialização
        dados_finais = {}
        nome_colaborador = None

        # Define as situações a ignorar (fora do loop!)
        situacoes_a_ignorar = mapa_situacoes_completo - set(mapa_situacoes)

        for _, row in df.iterrows():
            if pd.notna(row[2]):
                nome_colaborador = str(row[2]).strip()
                numloc = row[3].strip()
                local, setor = self.buscar_setor_e_local(numloc, self.dicionario_locais)
                if nome_colaborador not in dados_finais:
                    dados_finais[nome_colaborador] = {"Local": local, "Setor": setor}

            if pd.notna(row[5]) and pd.notna(row[6]) and pd.notna(row[9]):
                descricao_situacao = str(row[6]).strip()

                # Ignora situações que não estão no mapa_situacoes (usando o completo como referência)
                if descricao_situacao in situacoes_a_ignorar:
                    continue

                valor = self.ajustar_horas_excel(row[9])
                if nome_colaborador:
                    dados_finais[nome_colaborador][descricao_situacao] = valor

        # Unificar e somar situações
        novas_situacoes = set()

        for colaborador, dados in dados_finais.items():
            novo_dados = {}
            for situacao, valor in dados.items():
                if situacao in ["Local", "Setor"]:
                    novo_dados[situacao] = valor
                    continue

                situacao_renomeada = mapa_situacoes.get(situacao, situacao)
                novas_situacoes.add(situacao_renomeada)

                if situacao_renomeada in novo_dados:
                    novo_dados[situacao_renomeada] = self.somar_tempos(novo_dados[situacao_renomeada], valor)
                else:
                    novo_dados[situacao_renomeada] = valor

            dados_finais[colaborador] = novo_dados

        todas_situacoes = novas_situacoes  # Garante que apenas as situações finais serão incluídas

        # Solução crucial aqui:
        df_resultado = pd.DataFrame.from_dict(dados_finais, orient="index").reset_index()
        df_resultado.rename(columns={"index": "Colaborador"}, inplace=True)

        # Colunas dinâmicas garantidas
        for situacao in todas_situacoes:
            if situacao not in df_resultado.columns:
                df_resultado[situacao] = "-"

        df_resultado = self.adicionar_colunas_contagem(df_resultado, df)

        # Ordenação final das colunas
        situacoes = sorted(set(col.replace('Qtd. ', '') for col in df_resultado.columns if col.startswith('Qtd. ')))

        colunas_ordenadas = ['Local', 'Setor', 'Colaborador']
        for situacao in situacoes:
            qtd_col = f'Qtd. {situacao}'
            horas_col = situacao
            if qtd_col in df_resultado.columns and horas_col in df_resultado.columns:
                colunas_ordenadas.extend([qtd_col, horas_col])

        df_resultado = df_resultado[colunas_ordenadas].fillna("-")

        periodo_inicio_fmt = pd.to_datetime(periodo_inicio).strftime("%d/%m/%Y")
        periodo_fim_fmt = pd.to_datetime(periodo_fim).strftime("%d/%m/%Y")
        self.periodo_label.setText(f"Período Analisado: {periodo_inicio_fmt} a {periodo_fim_fmt}")


        self.mostrar_resultado(df_resultado)

        # Atualização do status_label ao fim do carregamento
        hora_atual = datetime.now().strftime("%H:%M")
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

        #  Criar um dicionário com os nomes das colunas para facilitar a busca específica
        headers = {}
        for col in range(self.tableWidget.columnCount()):
            header_item = self.tableWidget.horizontalHeaderItem(col)
            if header_item:  # Apenas adiciona colunas com nome
                headers[header_item.text().strip().lower()] = col

        #  Iterar por todas as linhas da tabela
        for row in range(self.tableWidget.rowCount()):
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
                        item = self.tableWidget.item(row, col_idx)

                        if item and item.text().strip().lower() == valor_busca.lower():
                            term_match = True  # Encontrou o termo exato na coluna correta

                else:
                    #  Busca geral (qualquer coluna até a coluna "Colaborador" inclusa)
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
        somando colunas de quantidade e colunas de horas (HHH:MM:SS), garantindo que sempre sejam somadas em pares."""

        if self.tableWidget.rowCount() == 0:
            return  # Se não há dados, não adiciona a linha de total

        num_cols = self.tableWidget.columnCount()

        # Criar dicionários para armazenar totais por tipo de dado
        totais_qtd = {}  # Para somar valores inteiros (Qtd.)
        totais_tempo = {}  # Para somar tempo (HHH:MM:SS)

        # Identificar colunas de "Qtd." e colunas de tempo "HHH:MM:SS"
        for col in range(num_cols):
            header = self.tableWidget.horizontalHeaderItem(col)
            if header:
                nome_coluna = header.text()
                if nome_coluna.startswith("Qtd."):  # Colunas de quantidade (inteiros)
                    totais_qtd[col] = 0
                else:  # Todas as outras colunas são de tempo (HHH:MM:SS)
                    totais_tempo[col] = timedelta()

        # Percorrer todas as linhas visíveis e somar os valores corretamente
        for row in range(self.tableWidget.rowCount()):
            if self.tableWidget.isRowHidden(row):
                continue  # Ignora linhas ocultas pelo filtro

            for col in range(num_cols):
                item = self.tableWidget.item(row, col)
                if item:
                    texto = item.text().strip()

                    # Somar colunas de quantidade (inteiros)
                    if col in totais_qtd and texto.isdigit():
                        totais_qtd[col] += int(texto)

                    # Somar colunas de tempo (HHH:MM:SS)
                    elif col in totais_tempo and ":" in texto:
                        try:
                            partes = list(map(int, texto.split(":")))  # Separa HH, MM, SS
                            if len(partes) == 3:  # Se o formato for correto
                                horas, minutos, segundos = partes
                                totais_tempo[col] += timedelta(hours=horas, minutes=minutos, seconds=segundos)
                        except ValueError:
                            pass  # Ignora erros de formatação de tempo

        # Remove a linha de total antiga, se existir
        if self.tableWidget.rowCount() > 0 and self.tableWidget.item(self.tableWidget.rowCount() - 1, 0) and \
                self.tableWidget.item(self.tableWidget.rowCount() - 1, 0).text() == "Total":
            self.tableWidget.removeRow(self.tableWidget.rowCount() - 1)

        # Adiciona nova linha de total
        row_total = self.tableWidget.rowCount()
        self.tableWidget.insertRow(row_total)

        for col in range(num_cols):
            if col == 0:
                item = QTableWidgetItem("Total")  # Primeira célula com "Total"
            elif col in totais_qtd:  # Soma de quantidade
                item = QTableWidgetItem(str(totais_qtd[col]))
            elif col in totais_tempo:  # Soma de horas, convertendo de volta para HHH:MM:SS
                total_segundos = int(totais_tempo[col].total_seconds())
                horas = total_segundos // 3600
                minutos = (total_segundos % 3600) // 60
                segundos = total_segundos % 60
                item = QTableWidgetItem(f"{horas:02}:{minutos:02}:{segundos:02}")
            else:
                item = QTableWidgetItem("")  # Mantém células vazias para colunas que não precisam de soma

            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item.setBackground(Qt.GlobalColor.yellow)  #  Fundo amarelo para destacar total
            item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
            self.tableWidget.setItem(row_total, col, item)

    def contar_situacoes_colaborador(self, df):
        """
        Conta quantas vezes cada situação aparece para cada colaborador,
        contando apenas entre o nome do colaborador até a linha 'Total',
        aplicando o mesmo mapeamento de unificação.
        """
        contagem_situacoes = {}
        nome_colaborador = None
        contando = False  # Flag para começar e parar a contagem

        mapa_situacoes = {
            "Saída Intermediária": "Atraso",
            "Saída Antecipada": "Atraso",
            "Saída Antecipada Noturna": "Atraso",
            "Saída Intermediária Noturna": "Atraso",
            "Atraso": "Atraso",
            "Atraso Noturno": "Atraso",
            "Faltas Noturnas": "Faltas"
        }

        for _, row in df.iterrows():
            if pd.notna(row[2]):  # Identifica novo colaborador
                nome_colaborador = row[2].strip()
                contagem_situacoes[nome_colaborador] = {}
                contando = True
                continue

            if contando:
                if pd.notna(row[0]) and str(row[0]).strip().lower() == "total":
                    contando = False
                    continue

                if pd.notna(row[6]):
                    situacao = row[6].strip()
                    situacao_renomeada = mapa_situacoes.get(situacao, situacao)

                    contagem_situacoes[nome_colaborador][situacao_renomeada] = (
                            contagem_situacoes[nome_colaborador].get(situacao_renomeada, 0) + 1
                    )

        return contagem_situacoes

    def adicionar_colunas_contagem(self, df_resultado, df_original):
        """
        Adiciona colunas de contagem ao DataFrame com o prefixo 'Qtd.'
        """
        contagens = self.contar_situacoes_colaborador(df_original)

        # Identificar todas as situações existentes
        todas_situacoes = set()
        for situacoes in contagens.values():
            todas_situacoes.update(situacoes.keys())

        # Adiciona as colunas com prefixo 'Qtd.'
        for situacao in sorted(todas_situacoes):
            coluna_contagem = f"Qtd. {situacao}"
            df_resultado[coluna_contagem] = df_resultado["Colaborador"].map(
                lambda colab: contagens.get(colab, {}).get(situacao, 0)
            )

        return df_resultado

    def ajustar_horas_excel(self, valor):
        """ Ajusta corretamente as horas acumuladas considerando dias acumulados como múltiplos de 24h. """
        if pd.isna(valor) or valor is None or valor == "nan":
            return "-"

        if isinstance(valor, datetime):  # Se for um datetime (com data completa)
            dia_real = valor.day  # Captura o dia corretamente
            horas = valor.hour
            minutos = valor.minute
            segundos = round(valor.second + valor.microsecond / 1_000_000)  #  Arredonda corretamente os segundos
            horas_totais = (dia_real * 24) + horas  #  Cada dia equivale a 24 horas

        elif isinstance(valor, time):  # Se for apenas um objeto `time`
            horas_totais = valor.hour
            minutos = valor.minute
            segundos = round(valor.second + valor.microsecond / 1_000_000)

        elif isinstance(valor, str):  #  Se for string "YYYY-MM-DD HH:MM:SS" ou "HH:MM:SS"
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

        #  Ajuste para sempre mostrar segundos como 00
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

            #  Tenta encontrar o CODLOC exato na planilha, reduzindo se necessário
        while codloc:
            for codloc_planilha, descricao in dicionario_locais.items():
                if codloc_planilha.startswith(codloc):  # Verifica se a linha começa com o CODLOC procurado
                    partes_descricao = descricao.split(",", 1)  # Divide apenas na primeira vírgula
                    local = partes_descricao[0].strip()  # A primeira parte define o Local

                    # O setor final será a última parte do CODLOC na planilha
                    setor_final = codloc_planilha.rsplit(",", 1)[-1].strip() if "," in codloc_planilha else local
                    return local, setor_final # Retorna imediatamente se encontrou

            #  Se não encontrou, tenta reduzir removendo a última parte do CODLOC
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

    def enviar_email_resumo(self):
        """Gera e abre um e-mail no Outlook com base no resumo do relatório filtrado na tela."""

        if self.tableWidget.rowCount() == 0:
            QMessageBox.warning(self, "Aviso", "Nenhum dado carregado para enviar por e-mail.")
            return

        #  Obter o período analisado do label
        periodo_analisado = self.periodo_label.text().replace("Período Analisado: ", "").strip()

        #  Criar dicionário de cores para os locais
        cores_locais = {
            "Fabrica De Máquinas": "#003756",
            "Fabrica De Transportadores": "#ffc62e",
            "Adm": "#009c44",
            "Comercial": "#919191"
        }

        #  Dicionários para armazenar dados do resumo e do relatório
        dados_agrupados = {}
        resumo_por_local = {}

        colunas_exibidas = []
        colunas_situacoes = []

        for col in range(self.tableWidget.columnCount()):
            nome_coluna = self.tableWidget.horizontalHeaderItem(col).text()
            if nome_coluna.startswith("Qtd."):
                colunas_situacoes.append((col, nome_coluna))  # Apenas contagens de situações
            if nome_coluna in ["Setor", "Colaborador"] or nome_coluna.startswith("Qtd."):
                colunas_exibidas.append((col, nome_coluna))  # Colunas que serão exibidas

        if not colunas_exibidas:
            QMessageBox.warning(self, "Aviso", "Nenhuma coluna válida encontrada para envio de e-mail.")
            return

        #  Coletar dados visíveis na tela e montar o resumo
        num_rows = self.tableWidget.rowCount()

        if num_rows > 0 and self.tableWidget.item(num_rows - 1, 0) and \
                self.tableWidget.item(num_rows - 1, 0).text() == "Total":
            num_rows -= 1  # Remove a linha de total do e-mail

        for row in range(num_rows):  # Agora o loop ignora a última linha (Total)
            if self.tableWidget.isRowHidden(row):
                continue  # Ignorar linhas ocultas pelo filtro

            local = self.tableWidget.item(row, 0).text()  # Pegando Local da primeira coluna
            setor = self.tableWidget.item(row, 1).text()  # Pegando Setor
            colaborador = self.tableWidget.item(row, 2).text()  # Pegando Colaborador

            if local not in dados_agrupados:
                dados_agrupados[local] = []
                resumo_por_local[local] = {sit: 0 for _, sit in colunas_situacoes}
                resumo_por_local[local]["setores"] = {}
                resumo_por_local[local]["colaboradores"] = {}

            linha_dados = [(setor, "#ffffff"), (colaborador, "#ffffff")]  # Sempre mantém Setor e Colaborador

            for col, nome_coluna in colunas_situacoes:
                item = self.tableWidget.item(row, col)
                valor = int(item.text()) if item and item.text().isdigit() else 0
                cor_fundo = item.background().color().name() if item else "#ffffff"

                linha_dados.append((valor, cor_fundo))

                # Atualiza os contadores por local
                resumo_por_local[local][nome_coluna] += valor

                # Atualiza os contadores por setor
                if setor not in resumo_por_local[local]["setores"]:
                    resumo_por_local[local]["setores"][setor] = {sit: 0 for _, sit in colunas_situacoes}
                resumo_por_local[local]["setores"][setor][nome_coluna] += valor

                # Atualiza os contadores por colaborador
                if colaborador not in resumo_por_local[local]["colaboradores"]:
                    resumo_por_local[local]["colaboradores"][colaborador] = {sit: 0 for _, sit in colunas_situacoes}
                resumo_por_local[local]["colaboradores"][colaborador][nome_coluna] += valor

            dados_agrupados[local].append(linha_dados)

        #  Criar corpo do e-mail com resumo por local
        corpo_email = f"""
        <p style='font-size:16px;'>Resumo do histórico de frequência do {periodo_analisado}:</p>
        """

        for local, resumo in sorted(resumo_por_local.items()):
            cor_local = cores_locais.get(local, "#000000")  # Usa cor do dicionário ou preto

            corpo_email += f'<h3 style="color:{cor_local}; font-size: 16px;">{local}</h3>'

            corpo_email += """
            <table border="1" cellspacing="0" cellpadding="5" 
                   style="border-collapse: collapse; width: auto; font-size: 14px; text-align: center;">
                <tr style="background-color: #f2f2f2; color: red;">
                    <th>Situação</th>
                    <th>Qtd. Total</th>
                    <th>Setor(es) de Maior Ocorrência</th>
                </tr>
            """

            # Adicionar as linhas com dados compactos
            for _, situacao in colunas_situacoes:
                total_ocorrencias = resumo.get(situacao, 0)
                if total_ocorrencias > 0:
                    # Obter os setores com maior ocorrência
                    max_ocorrencia_setor = max((resumo["setores"][s].get(situacao, 0) for s in resumo["setores"]),
                                               default=0)
                    setores_mais_ocorrencias = [
                        f"{s} - {resumo['setores'][s][situacao]}" for s in resumo["setores"]
                        if resumo["setores"][s].get(situacao, 0) == max_ocorrencia_setor and max_ocorrencia_setor > 0
                    ]

                    setores_str = " / ".join(setores_mais_ocorrencias)

                    corpo_email += f"""
                    <tr>
                        <td><b>{situacao}</b></td>
                        <td>{total_ocorrencias}</td>
                        <td>{setores_str}</td>
                    </tr>
                    """

            corpo_email += "</table><br>"

            corpo_email += """
            <table border="1" cellspacing="0" cellpadding="5" 
                   style="border-collapse: collapse; width: auto; font-size: 14px; text-align: center;">
                <tr style="background-color: #f2f2f2; color: red;">
                    <th>Situação</th>
                    <th>Colaborador(es) de Maior Ocorrência Entre Todos os Setores</th>
                </tr>
            """

            # Adicionar as linhas com dados dos colaboradores
            for _, situacao in colunas_situacoes:
                total_ocorrencias = resumo.get(situacao, 0)
                if total_ocorrencias > 0:
                    # Obter os colaboradores com maior ocorrência
                    max_ocorrencia_colaborador = max(
                        (resumo["colaboradores"][c].get(situacao, 0) for c in resumo["colaboradores"]), default=0
                    )
                    colaboradores_mais_ocorrencias = [
                        f"{c} - {resumo['colaboradores'][c][situacao]}" for c in resumo["colaboradores"]
                        if resumo["colaboradores"][c].get(situacao,
                                                          0) == max_ocorrencia_colaborador and max_ocorrencia_colaborador > 0
                    ]

                    colaboradores_str = " / ".join(colaboradores_mais_ocorrencias)

                    corpo_email += f"""
                    <tr>
                        <td><b>{situacao}</b></td>
                        <td>{colaboradores_str}</td>
                    </tr>
                    """

            corpo_email += "</table><br>"

        #  Abrir o Outlook e criar um e-mail
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            email = outlook.CreateItem(0)
            email.Subject = f"Relatório Resumido Histórico de Frequência - {periodo_analisado}"
            inspector = email.GetInspector  # Abre o e-mail para edição
            signature = email.HTMLBody  # Obtém a assinatura padrão do Outlook
            email.HTMLBody = corpo_email + "<br><br>" + signature  # Adiciona a assinatura ao e-mail

            # Armazena o e-mail para ser acessado depois
            self.janela_principal.current_mail = email
            self.current_mail = email
            email.Display()

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao abrir o Outlook: {e}")

    def enviar_email_detalhado(self):
        """Gera e abre um e-mail no Outlook com base no detalhamento do relatório filtrado na tela."""

        if self.tableWidget.rowCount() == 0:
            QMessageBox.warning(self, "Aviso", "Nenhum dado carregado para enviar por e-mail.")
            return

        #  Obter o período analisado do label
        periodo_analisado = self.periodo_label.text().replace("Período Analisado: ", "").strip()

        #  Criar dicionário de cores para os locais
        cores_locais = {
            "Fabrica De Máquinas": "#003756",
            "Fabrica De Transportadores": "#ffc62e",
            "Adm": "#009c44",
            "Comercial": "#919191"
        }

        #  Dicionários para armazenar dados do resumo e do relatório
        dados_agrupados = {}

        colunas_exibidas = []
        colunas_situacoes = []

        for col in range(self.tableWidget.columnCount()):
            nome_coluna = self.tableWidget.horizontalHeaderItem(col).text()
            if nome_coluna.startswith("Qtd."):
                colunas_situacoes.append((col, nome_coluna))  # Apenas contagens de situações
            if nome_coluna in ["Setor", "Colaborador"] or nome_coluna.startswith("Qtd."):
                colunas_exibidas.append((col, nome_coluna))  # Colunas que serão exibidas

        if not colunas_exibidas:
            QMessageBox.warning(self, "Aviso", "Nenhuma coluna válida encontrada para envio de e-mail.")
            return

        #  Coletar dados visíveis na tela e montar o detalhamento
        num_rows = self.tableWidget.rowCount()

        if num_rows > 0 and self.tableWidget.item(num_rows - 1, 0) and \
                self.tableWidget.item(num_rows - 1, 0).text() == "Total":
            num_rows -= 1  # Remove a linha de total do e-mail

        for row in range(num_rows):  # Agora o loop ignora a última linha (Total)
            if self.tableWidget.isRowHidden(row):
                continue  # Ignorar linhas ocultas pelo filtro

            local = self.tableWidget.item(row, 0).text()  # Pegando Local da primeira coluna
            setor = self.tableWidget.item(row, 1).text()  # Pegando Setor
            colaborador = self.tableWidget.item(row, 2).text()  # Pegando Colaborador

            if local not in dados_agrupados:
                dados_agrupados[local] = []

            linha_dados = [(setor, "#ffffff"), (colaborador, "#ffffff")]  # Sempre mantém Setor e Colaborador

            for col, nome_coluna in colunas_situacoes:
                item = self.tableWidget.item(row, col)
                valor = int(item.text()) if item and item.text().isdigit() else 0
                cor_fundo = item.background().color().name() if item else "#ffffff"

                linha_dados.append((valor, cor_fundo))

            dados_agrupados[local].append(linha_dados)

        #  Criar corpo do e-mail com detalhamento por local
        corpo_email = f"""
        <p style='font-size:16px;'>Detalhamento do relatório de frequência de {periodo_analisado}:</p>
        """

        # Criar tabela detalhada por local
        for local, colaboradores in sorted(dados_agrupados.items()):
            cor_local = cores_locais.get(local, "#000000")  # Usa cor do dicionário ou preto

            corpo_email += f'<h3 style="color:{cor_local}; font-size: 16px;">{local}</h3>'
            corpo_email += """
            <table border="1" cellspacing="0" cellpadding="3" 
                   style="border-collapse: collapse; width: auto; font-size: 14px;">
                <tr style="background-color: #f2f2f2; text-align: center;">
            """

            # Adicionar cabeçalhos na tabela do e-mail (Setor, Colaborador e Qtd.)
            for _, nome_coluna in colunas_exibidas:
                corpo_email += f'<th style="padding: 3px;">{nome_coluna}</th>'
            corpo_email += "</tr>"

            # Adicionar os dados dos colaboradores
            for linha_dados in colaboradores:
                corpo_email += "<tr>"
                for valor, cor_fundo in linha_dados:
                    corpo_email += f'<td style="padding: 3px; background-color: {cor_fundo}; text-align: center;">{"-" if valor == 0 else valor}</td>'
                corpo_email += "</tr>"

            corpo_email += "</table><br>"

        #  Abrir o Outlook e criar um e-mail
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            email = outlook.CreateItem(0)
            email.Subject = f"Relatório Detalhado de Histórico de Frequência - {periodo_analisado}"
            inspector = email.GetInspector  # Abre o e-mail para edição
            signature = email.HTMLBody  # Obtém a assinatura padrão do Outlook
            email.HTMLBody = corpo_email + "<br><br>" + signature  # Adiciona a assinatura ao e-mail

            # Armazena o e-mail para ser acessado depois
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
        """Abre a aba de gráficos passando os dados da tabela atual."""

        # Criar DataFrame a partir da tabela atual
        dados = []
        colunas = [self.tableWidget.horizontalHeaderItem(col).text() for col in range(self.tableWidget.columnCount())]

        for row in range(self.tableWidget.rowCount()):
            if not self.tableWidget.isRowHidden(row):  # Considera apenas linhas visíveis
                linha = [self.tableWidget.item(row, col).text() if self.tableWidget.item(row, col) else '-'
                         for col in range(self.tableWidget.columnCount())]
                dados.append(linha)

        # Se não houver dados visíveis, exibir aviso e não abrir gráficos
        if not dados:
            QMessageBox.warning(self, "Aviso", "Nenhum dado visível. Ajuste os filtros ou calcule primeiro.")
            return

        # Criar DataFrame
        df = pd.DataFrame(dados, columns=colunas)

        #  Separar colunas de "Qtd." e "Horas"
        colunas_qtd = [col for col in df.columns if col.startswith("Qtd.")]
        colunas_horas = [col for col in df.columns if
                         col not in colunas_qtd and col not in ["Local", "Setor", "Colaborador"]]

        #  Converter apenas colunas de "Qtd." para números inteiros
        df[colunas_qtd] = df[colunas_qtd].replace("-", 0).apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)

        #  Converter colunas de horas para timedelta (para permitir soma depois)
        df[colunas_horas] = df[colunas_horas].replace("-", "00:00:00").apply(pd.to_timedelta, errors='coerce').fillna(
            pd.Timedelta(seconds=0))

        # Criar e abrir a aba de gráficos
        self.aba_grafico = AbaGraficoFrequencia(df,
                                                self.periodo_label.text())  # Passa o DataFrame e o período analisado
        self.aba_grafico.show()

        # Se um e-mail já foi gerado, passar para a aba de gráficos
        if hasattr(self, "current_mail") and self.current_mail:
            self.aba_grafico.set_current_email(self.current_mail)

    def mostrar_resultado(self, df):
        """Exibe o DataFrame na tabela da interface garantindo reset completo, inclusive na ordenação."""

        #  Reset total da tabela antes de recriar os dados
        self.tableWidget.setSortingEnabled(False)  # Desativa ordenação temporariamente
        self.tableWidget.clear()
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)

        #  Define a nova estrutura da tabela
        self.tableWidget.setRowCount(df.shape[0])
        self.tableWidget.setColumnCount(df.shape[1])
        self.tableWidget.setHorizontalHeaderLabels(df.columns)
        self.tableWidget.verticalHeader().setVisible(False)
        self.tableWidget.setStyleSheet("QTableWidget {gridline-color: black;}")

        #  Reseta a ordenação antes de preencher os novos dados
        self.tableWidget.horizontalHeader().setSortIndicator(-1, Qt.SortOrder.AscendingOrder)

        cores = [QColor("#FFFFFF"), QColor("#D3D3D3")]
        coluna_cor = {}

        cor_idx = 0
        for col in range(3, df.shape[1]):  # A partir da 4ª coluna (Local, Setor, Colaborador são fixos)
            if (col - 3) % 2 == 0:
                cor_idx = (cor_idx + 1) % 2
            coluna_cor[col] = cores[cor_idx]

        # Define as três primeiras colunas com cor fixa (sem agrupamento)
        for col in range(3):
            coluna_cor[col] = QColor("#FFFFFF")  # Branco fixo para as três primeiras colunas

        #  Preenche a tabela com os dados
        for row in range(df.shape[0]):
            for col in range(df.shape[1]):
                valor = df.iat[row, col]

                # Colunas iniciais fixas
                if col in [0, 1, 2]:
                    item = QTableWidgetItem(str(valor))
                else:
                    coluna = df.columns[col]
                    if coluna.startswith('Qtd.'):
                        try:
                            valor_int = int(valor)
                            if valor_int == 0:
                                item = QTableWidgetItem("-")
                            else:
                                item = QTableWidgetItem(str(valor_int))
                                item.setData(Qt.ItemDataRole.EditRole, valor_int)
                        except ValueError:
                            item = QTableWidgetItem("-")

                    else:  # Colunas de tempo no formato "HH:MM:SS"
                        try:
                            h, m, s = map(int, valor.split(':'))
                            segundos_totais = h * 3600 + m * 60 + s
                            item = QTableWidgetItem()
                            item.setData(Qt.ItemDataRole.EditRole, segundos_totais)
                            item.setText(valor)
                        except ValueError:
                            item = QTableWidgetItem("-")

                item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                item.setForeground(QColor("#000000"))

                if col in coluna_cor:
                    item.setBackground(coluna_cor[col])

                self.tableWidget.setItem(row, col, item)

        #  Ajusta fonte e tamanho das colunas
        fonte = self.tableWidget.font()
        fonte.setPointSize(8)
        self.tableWidget.setFont(fonte)

        self.tableWidget.resizeColumnsToContents()

        largura_extra = 15  # Pequeno ajuste extra nas colunas
        for col in range(self.tableWidget.columnCount()):
            largura_atual = self.tableWidget.columnWidth(col)

            #  Apenas as colunas após "Colaborador" recebem o ajuste de 15
            if col > 2:
                self.tableWidget.setColumnWidth(col, largura_atual + largura_extra)
            else:
                self.tableWidget.setColumnWidth(col, largura_atual)  # Mantém original

        #  Reativa a ordenação, mas sem seguir a ordenação anterior
        self.tableWidget.setSortingEnabled(True)
        self.atualizar_total()








