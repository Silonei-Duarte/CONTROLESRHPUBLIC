import sys
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QLabel, QTableWidget, QTableWidgetItem, QHeaderView
import pandas as pd

class JanelaDetalhesColaboradores(QDialog):
    """Janela para exibir detalhes dos colaboradores gerenciados por um gestor"""
    def __init__(self, nome_gestor, df_ativos, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Colaboradores gerenciados por: {nome_gestor}")
        self.resize(1200, 600)
        
        # Configurar layout
        layout = QVBoxLayout(self)
        
        # Filtrar colaboradores gerenciados pelo gestor
        colaboradores_df = df_ativos[df_ativos["Gestor"] == nome_gestor].copy()
        
        if colaboradores_df.empty:
            # Criar um DataFrame vazio com as colunas necessárias para evitar erro
            self.df = pd.DataFrame(columns=["Local", "Setor", "Tipo", "Colaborador", "Cargo", "Salário Mensal", "Data Admissão"])
            info_label = QLabel("Nenhum colaborador encontrado.")
        else:
            # Selecionar apenas as colunas necessárias
            colunas_detalhes = ["Local", "Setor", "Tipo", "Colaborador", "Cargo", "Salário Mensal", "Data Admissão"]
            self.df = colaboradores_df[colunas_detalhes].copy()
            
            # Ordenar por setor e nome do colaborador
            self.df = self.df.sort_values(by=["Setor", "Colaborador"])
            info_label = QLabel(f"Total de registros: {len(self.df)}")
        
        # Adicionar informação sobre a quantidade
        layout.addWidget(info_label)
        
        # Criar tabela
        self.tabela = QTableWidget()
        layout.addWidget(self.tabela)
        
        # Preencher a tabela com dados
        self.preencher_tabela()
        
        # Configurar janela
        self.setLayout(layout)

    def preencher_tabela(self):
        """Método para preencher a tabela com dados"""
        if self.df.empty:
            return

        colunas = list(self.df.columns)
        self.tabela.setColumnCount(len(colunas))
        self.tabela.setHorizontalHeaderLabels(colunas)
        self.tabela.setRowCount(len(self.df))

        fonte = self.tabela.font()
        fonte.setPointSize(8)
        self.tabela.setFont(fonte)

        for row_idx in range(len(self.df)):
            for col_idx, col in enumerate(colunas):
                valor = self.df.iloc[row_idx, col_idx]
                item = QTableWidgetItem(str(valor))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tabela.setItem(row_idx, col_idx, item)

        self.tabela.resizeRowsToContents()
        self.tabela.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        # Ajustar as colunas para dividir igualmente o espaço
        self.tabela.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        # Oculta os números das linhas
        self.tabela.verticalHeader().setVisible(False)

class JanelaDetalhesSetoresGerenciados(QDialog):
    """Janela para exibir detalhes dos setores gerenciados por um gestor"""
    def __init__(self, nome_gestor, df_ativos, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Setores gerenciados por: {nome_gestor}")
        self.resize(600, 400)  # Reduzido para metade da largura
        
        # Configurar layout
        layout = QVBoxLayout(self)
        
        # Filtrar colaboradores gerenciados pelo gestor
        colaboradores_df = df_ativos[df_ativos["Gestor"] == nome_gestor].copy()
        
        if colaboradores_df.empty:
            # Criar um DataFrame vazio com as colunas necessárias para evitar erro
            self.df = pd.DataFrame(columns=["Local", "Setor", "Colaboradores Supervisionados"])
            info_label = QLabel("Nenhum setor encontrado.")
        else:
            # Agrupar por setor e contar colaboradores
            resumo_setores = []
            setores_unicos = colaboradores_df["Setor"].unique()
            
            for setor in setores_unicos:
                # Contagem de colaboradores no setor
                qtd_colaboradores = len(colaboradores_df[colaboradores_df["Setor"] == setor])
                
                # Local do setor
                local = colaboradores_df[colaboradores_df["Setor"] == setor]["Local"].iloc[0]
                
                resumo_setores.append({
                    "Local": local,
                    "Setor": setor,
                    "Colaboradores Supervisionados": qtd_colaboradores
                })
            
            # Ordenar por setor
            resumo_setores = sorted(resumo_setores, key=lambda x: x["Setor"])
            
            # Criar DataFrame para exibição
            self.df = pd.DataFrame(resumo_setores)
            info_label = QLabel(f"Total de setores: {len(self.df)}")
        
        # Adicionar informação sobre a quantidade
        layout.addWidget(info_label)
        
        # Criar tabela
        self.tabela = QTableWidget()
        layout.addWidget(self.tabela)
        
        # Preencher a tabela com dados
        self.preencher_tabela()
        
        # Configurar janela
        self.setLayout(layout)
        
    def preencher_tabela(self):
        """Método para preencher a tabela com dados"""
        if self.df.empty:
            return

        colunas = list(self.df.columns)
        self.tabela.setColumnCount(len(colunas))
        self.tabela.setHorizontalHeaderLabels(colunas)
        self.tabela.setRowCount(len(self.df))

        fonte = self.tabela.font()
        fonte.setPointSize(8)
        self.tabela.setFont(fonte)

        for row_idx in range(len(self.df)):
            for col_idx, col in enumerate(colunas):
                valor = self.df.iloc[row_idx, col_idx]
                item = QTableWidgetItem(str(valor))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tabela.setItem(row_idx, col_idx, item)

        self.tabela.resizeRowsToContents()
        self.tabela.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        # Ajustar as colunas para dividir igualmente o espaço
        self.tabela.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        # Oculta os números das linhas
        self.tabela.verticalHeader().setVisible(False)

class PainelGestores:
    @staticmethod
    def atualizar_tabela_gestores(df_ativos, tabela, label_info):
        """Atualiza a tabela de gestores com os dados fornecidos
        
        Args:
            df_ativos: DataFrame contendo os dados dos colaboradores ativos
            tabela: QTableWidget onde os dados serão exibidos
            label_info: QLabel para mostrar informações sobre a quantidade de gestores
        """
        # Armazenar filtro visível atual
        filtro_visivel = {}
        for row in range(tabela.rowCount()):
            filtro_visivel[row] = not tabela.isRowHidden(row)
            
        if df_ativos.empty:
            tabela.setRowCount(0)
            tabela.setColumnCount(0)
            label_info.setText("Nenhum dado disponível.")
            return
            
        # Criar DataFrame com os gestores únicos
        gestores_df = df_ativos.copy()
    
        # Remover linhas onde o gestor está vazio
        gestores_df = gestores_df[gestores_df["Gestor"].notna() & (gestores_df["Gestor"] != "")]
    
        # Agrupar por gestor e contar colaboradores
        resumo_gestores = []
        gestores_unicos = gestores_df["Gestor"].unique()
    
        for gestor in gestores_unicos:
            # Dados do gestor (pode haver mais de uma linha para o mesmo gestor)
            dados_gestor = gestores_df[gestores_df["Gestor"] == gestor].iloc[0]
            
            # Contagem de colaboradores
            qtd_colaboradores = len(gestores_df[gestores_df["Gestor"] == gestor])
            
            # Obter o tipo do gestor
            tipo_gestor = dados_gestor.get("Tipo Gestor", "")
            
            # Contagem de setores gerenciados
            setores_gerenciados = gestores_df[gestores_df["Gestor"] == gestor]["Setor"].nunique()
            
            # Encontrar o gestor do gestor atual
            gestor_acima = ""
            if "Gestor" in df_ativos.columns:
                dados_pessoa = df_ativos[df_ativos["Colaborador"] == gestor]
                if not dados_pessoa.empty and dados_pessoa["Gestor"].iloc[0]:
                    gestor_acima = dados_pessoa["Gestor"].iloc[0]
            
            resumo_gestores.append({
                "Nome do Gestor": gestor,
                "Local": dados_gestor["Local Gestor"],
                "Setor": dados_gestor["Setor Gestor"],
                "Tipo": tipo_gestor,
                "Qtd. Supervisionados": qtd_colaboradores,
                "Quantidade de Setores Gerenciados": setores_gerenciados,
                "Gestor Acima": gestor_acima
            })
    
        # Ordenar por nome do gestor
        resumo_gestores = sorted(resumo_gestores, key=lambda x: x["Nome do Gestor"])
    
        # Configurar a tabela
        colunas = ["Nome do Gestor", "Local", "Setor", "Tipo", "Qtd. Supervisionados", 
                  "Quantidade de Setores Gerenciados", "Gestor Acima"]
        tabela.setColumnCount(len(colunas))
        tabela.setHorizontalHeaderLabels(colunas)
        tabela.setRowCount(len(resumo_gestores))
    
        # Ajustar tamanho da fonte
        fonte = tabela.font()
        fonte.setPointSize(9)
        tabela.setFont(fonte)
    
        # Preencher a tabela
        for row_idx, gestor_info in enumerate(resumo_gestores):
            for col_idx, coluna in enumerate(colunas):
                valor = gestor_info[coluna]
                item = QTableWidgetItem(str(valor))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                
                # Alinhar o texto ao centro
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                
                # Armazenar o nome do gestor como dado para uso posterior
                if coluna == "Qtd. Supervisionados" or coluna == "Quantidade de Setores Gerenciados":
                    item.setData(Qt.ItemDataRole.UserRole, gestor_info["Nome do Gestor"])
                
                tabela.setItem(row_idx, col_idx, item)
    
        # Ajustar o tamanho das colunas ao conteúdo
        tabela.resizeColumnsToContents()
        tabela.resizeRowsToContents()
        tabela.horizontalHeader().setStretchLastSection(True)
        
        # Atualizar o label informativo
        label_info.setText(f"Lista de Gestores e Qtd. Supervisionados (Total: {len(resumo_gestores)} gestores)")
        
        # Limpar conexões anteriores
        try:
            tabela.cellDoubleClicked.disconnect()
        except:
            pass
            
        # Configurar sinal para exibir detalhes ao clicar em uma célula
        tabela.cellDoubleClicked.connect(lambda row, col: 
            PainelGestores.clique_duplo_na_celula(row, col, tabela, df_ativos)
        )
        
    @staticmethod
    def clique_duplo_na_celula(row, col, tabela, df_ativos):
        """Processa o clique duplo em uma célula da tabela"""
        # Obter o nome da coluna clicada
        header_item = tabela.horizontalHeaderItem(col)
        if header_item is None:
            return
            
        coluna_clicada = header_item.text()
        
        # Verificar se a coluna clicada é uma das que queremos processar
        if coluna_clicada == "Qtd. Supervisionados":
            PainelGestores.abrir_detalhes_colaboradores(row, col, tabela, df_ativos)
        elif coluna_clicada == "Quantidade de Setores Gerenciados":
            PainelGestores.abrir_detalhes_setores(row, col, tabela, df_ativos)
    
    @staticmethod
    def abrir_detalhes_colaboradores(row, col, tabela, df_ativos):
        """Abre a janela de detalhes dos colaboradores"""
        item = tabela.item(row, col)
        if item is None:
            return
            
        # Recuperar o nome do gestor armazenado no item
        nome_gestor = item.data(Qt.ItemDataRole.UserRole)
        if not nome_gestor:
            return
            
        # Criar e mostrar a janela de detalhes
        janela_detalhes = JanelaDetalhesColaboradores(
            nome_gestor, 
            df_ativos, 
            tabela.window()
        )
        janela_detalhes.exec()
    
    @staticmethod
    def abrir_detalhes_setores(row, col, tabela, df_ativos):
        """Abre a janela de detalhes dos setores gerenciados"""
        item = tabela.item(row, col)
        if item is None:
            return
            
        # Recuperar o nome do gestor armazenado no item
        nome_gestor = item.data(Qt.ItemDataRole.UserRole)
        if not nome_gestor:
            return
            
        # Criar e mostrar a janela de detalhes
        janela_detalhes = JanelaDetalhesSetoresGerenciados(
            nome_gestor, 
            df_ativos, 
            tabela.window()
        )
        janela_detalhes.exec()


