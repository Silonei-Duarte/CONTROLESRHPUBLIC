import sys
import os

# --- Patch para Kaleido no executável ---
import plotly.io as pio
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
    kaleido_exe = os.path.join(base_path, 'kaleido', 'executable', 'bin', 'kaleido.exe')
    if os.path.exists(kaleido_exe):
        os.environ['PATH'] += os.pathsep + os.path.dirname(kaleido_exe)
        pio.kaleido.scope = None

import orjson
if not hasattr(orjson, "OPT_NON_STR_KEYS"):
    orjson.OPT_NON_STR_KEYS = 0

# ----------------------------------------

import json
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel, QMessageBox, QVBoxLayout, QWidget, \
    QGridLayout, QDialog, QScrollArea, QSpacerItem, QSizePolicy
from PyQt6.QtGui import QFont, QIcon
from PyQt6.QtCore import Qt
from telaaviso import AppTelaAvisos

class ControlesRH(QMainWindow):
    VERSAO_ATUAL = "1.92.0026"

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Controles RH")
        self.setGeometry(700, 300, 250, 350)

        # Configurar o ícone da janela
        icon_path = os.path.join(os.path.dirname(__file__), 'icone.ico')  # Ajuste o caminho conforme necessário
        if os.path.exists(icon_path):  # Verifica se o arquivo existe
            self.setWindowIcon(QIcon(icon_path))

        self.criar_arquivo_versao()
        self.initUI()
        self.verificar_atualizacao()

    def initUI(self):
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        layout = QGridLayout()
        layout.setHorizontalSpacing(20)
        layout.setVerticalSpacing(10)

        layout.setColumnStretch(0, 1)
        layout.setColumnStretch(1, 0)
        layout.setColumnStretch(2, 1)

        self.label = QLabel("Menu Principal", self)
        self.label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        self.label.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
        self.label.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        layout.addWidget(self.label, 0, 1, Qt.AlignmentFlag.AlignHCenter)

        self.btn_frequencia = QPushButton("Controle de Frequência", self)
        self.btn_frequencia.setFixedSize(180, 35)
        self.btn_frequencia.clicked.connect(self.abrir_frequencia)
        layout.addWidget(self.btn_frequencia, 1, 0, Qt.AlignmentFlag.AlignHCenter)

        self.btn_horaextra = QPushButton("Controle de Horas Extras", self)
        self.btn_horaextra.setFixedSize(180, 35)
        self.btn_horaextra.clicked.connect(self.abrir_horaextra)
        layout.addWidget(self.btn_horaextra, 2, 0, Qt.AlignmentFlag.AlignHCenter)

        self.btn_afastamentos = QPushButton("Controle de Afastamentos", self)
        self.btn_afastamentos.setFixedSize(180, 35)
        self.btn_afastamentos.clicked.connect(self.abrir_afastamentos)
        layout.addWidget(self.btn_afastamentos, 3, 0, Qt.AlignmentFlag.AlignHCenter)

        self.btn_advertencias = QPushButton("Controle de Advertências", self)
        self.btn_advertencias.setFixedSize(180, 35)
        self.btn_advertencias.clicked.connect(self.abrir_advertencias)
        layout.addWidget(self.btn_advertencias, 1, 1, Qt.AlignmentFlag.AlignHCenter)

        self.btn_documentosvencidos = QPushButton("Controle de Documentação", self)
        self.btn_documentosvencidos.setFixedSize(180, 35)
        self.btn_documentosvencidos.clicked.connect(self.abrir_documentosvencidos)
        layout.addWidget(self.btn_documentosvencidos, 2, 1, Qt.AlignmentFlag.AlignHCenter)

        self.btn_painel_gestor = QPushButton("Painel do Gestor", self)
        self.btn_painel_gestor.setFixedSize(180, 35)
        self.btn_painel_gestor.clicked.connect(self.abrir_painel_gestor)
        layout.addWidget(self.btn_painel_gestor, 3, 1, Qt.AlignmentFlag.AlignHCenter)

        self.btn_calculo_eventos = QPushButton("Cálculo Eventos em Folha", self)
        self.btn_calculo_eventos.setFixedSize(180, 35)
        self.btn_calculo_eventos.clicked.connect(self.abrir_calculo_eventos)
        layout.addWidget(self.btn_calculo_eventos, 1, 2, Qt.AlignmentFlag.AlignHCenter)

        # Adiciona os cards do AppTelaAvisos no centro (linha 4, coluna 0 até 2)
        self.card_widget = AppTelaAvisos()
        layout.addWidget(self.card_widget, 4, 0, 1, 3, Qt.AlignmentFlag.AlignCenter)

        # Espaço expansível abaixo dos cards
        layout.addItem(QSpacerItem(20, 40, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding), 5, 0, 1, 3)

        self.btn_sobre = QPushButton("Sobre", self)
        self.btn_sobre.setFixedSize(150, 35)
        self.btn_sobre.clicked.connect(self.mostrar_sobre)
        layout.addWidget(self.btn_sobre, 6, 0, 1, 3, Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignBottom)

        self.central_widget.setLayout(layout)

    def criar_arquivo_versao(self):
        arquivo_versao_local = os.path.join(os.path.dirname(__file__), 'versao.json')
        if not os.path.exists(arquivo_versao_local):
            dados_versao = {"versao": self.VERSAO_ATUAL}
            try:
                with open(arquivo_versao_local, 'w') as f:
                    json.dump(dados_versao, f)
            except Exception as e:
                print(f"Erro ao criar arquivo de versão: {e}")

    def verificar_atualizacao(self):
        diretorio_rede = r'\\*\Distribuição\Controles RH\_internal'
        arquivo_versao_rede = os.path.join(diretorio_rede, 'versao.json')

        if not os.path.exists(arquivo_versao_rede):
            return

        try:
            with open(arquivo_versao_rede, 'r') as f:
                dados_rede = json.load(f)
                versao_rede = dados_rede.get('versao', None)

            if versao_rede and versao_rede > self.VERSAO_ATUAL:
                resposta = QMessageBox.question(self, "Atualização disponível",
                                                f"Nova versão ({versao_rede}) disponível. Deseja atualizar agora?",
                                                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if resposta == QMessageBox.StandardButton.Yes:
                    self.criar_script_atualizacao(diretorio_rede)
        except Exception:
            pass

    def criar_script_atualizacao(self, diretorio_rede):
        """Cria um script de atualização seguro, executa e encerra imediatamente a aplicação."""
        script_path = os.path.join(os.path.dirname(__file__), 'atualizacao.bat')
        nome_executavel = os.path.basename(sys.executable)

        diretorio_completo = r'\\*\Distribuição\Controles RH'
        destino_final = os.path.dirname(os.path.dirname(__file__))
        pasta_temporaria = os.path.join(destino_final, 'temp_update')

        with open(script_path, 'w', encoding='cp850') as f:
            f.write('@echo off\n')
            f.write('chcp 850 > nul\n')
            f.write('cls\n')
            f.write('echo ====================================\n')
            f.write('echo   Atualizando a aplicação...\n')
            f.write('echo ====================================\n\n')

            # 1. Por segurança: fecha o app se ainda estiver rodando
            f.write(f'echo Fechando a aplicação...\n')
            f.write(f'taskkill /IM "{nome_executavel}" /F 2>NUL\n')
            f.write(f'if %ERRORLEVEL% NEQ 0 (echo Processo "{nome_executavel}" já estava fechado.)\n')
            f.write('timeout /t 2 >nul\n')

            # 2. Copiar para pasta temporária
            f.write(f'echo Copiando arquivos para pasta temporária...\n')
            f.write(f'robocopy "{diretorio_completo}" "{pasta_temporaria}" /E /IS /NFL /NDL /NJH /NJS /NC /NS /MT:8\n')
            f.write('if %ERRORLEVEL% GEQ 8 (\n')
            f.write('    echo ERRO ao copiar para pasta temporária. Abortando.\n')
            f.write('    pause\n')
            f.write('    exit /b\n')
            f.write(')\n\n')

            # 3. Copiar da temporária para destino final
            f.write('echo Copiando atualização final...\n')
            f.write(f'robocopy "{pasta_temporaria}" "{destino_final}" /E /IS /NFL /NDL /NJH /NJS /NC /NS /MT:8\n')
            f.write('if %ERRORLEVEL% GEQ 8 (\n')
            f.write('    echo ERRO ao copiar atualização. Abortando.\n')
            f.write('    pause\n')
            f.write('    exit /b\n')
            f.write(')\n\n')

            # 4. Limpa temporário
            f.write(f'rd /s /q "{pasta_temporaria}"\n')

            # 5. Relança
            f.write(f'echo Iniciando aplicação...\n')
            f.write(f'start "" "{nome_executavel}"\n')
            f.write('exit\n')

        # Fecha a aplicação imediatamente
        os.startfile(script_path)
        sys.exit(0)

    def mostrar_sobre(self):
        # Criar a janela de diálogo personalizada
        dialog = QDialog(self)
        dialog.setWindowTitle("Sobre")
        dialog.setFixedSize(500, 300)  # Tamanho fixo do diálogo

        # Criar o layout principal
        layout = QVBoxLayout()

        # Criar um QLabel com Rich Text para o LinkedIn e Notas de Versão
        label_texto = QLabel(
            f"""
            <p align='center'><b>Versão:</b> {self.VERSAO_ATUAL}</p>
            <p align='center'><b>Empresa:</b> Bruno Industrial</p>
            <p align='center'><b>Criador:</b> 
            <a href='https://www.linkedin.com/in/siloneiduarte/' style='color: blue; text-decoration: none;'>Silonei Duarte</a></p>
            <p align='center'><b>Notas de Versão:</b></p>
            <p align='center'><b>1.92.0019:</b><br>
            - Adicionado Aviso de vistos Estrangeiros vencendo
            <p align='center'><b>1.92.0011:</b><br>
            - Lançado Grafico para Aba de Setores do Painel de Gestor
            <p align='center'><b>1.92.0008:</b><br>
            - Lançado Calculos de folha por evento
            <p align='center'><b>1.91.0000:</b><br>
            - Lançado no Controle de Frequência calculo de assiduidade e absenteísmo</p>
             <p align='center'><b>1.9.0000:</b><br>
            - Lançado Painel do Gestor com calculos de colaboradores e setores e Turnover<br>
            - Lançado Criador de Organograma</p>
            <p align='center'><b>1.8.0000:</b><br>
            - Lançado Quadro de Avisos no Menu Principal</p>
            - Adicionado possibilidade de inserir linha de referência no Dashboard de horas Extras</p>
            <p align='center'><b>1.7.00000:</b><br>
            - Lançado Controle de Advertências com Geração de E-mails e Dashboard</p>
            <p align='center'><b>1.6.00000:</b><br>
            - Lançado Controle de Documentações Vencidas e Geração de E-mails</p>
            <p align='center'><b>1.5.00000:</b><br>
            - Lançado Controle de Afastamentos semanais com Dashboard e Geração de E-mails</p>
            <p align='center'><b>1.4.00000:</b><br>
            - Gráficos agora utilizam biblioteca mais recente, permitindo agrupamento de barras </p>
            - Controle de Frequência e Ausências realizam mesclagem de e-mails e Dashboard </p>
            <p align='center'><b>1.3.00000:</b><br>
            - Lançado Geração de Gráficos Dinâmicos de histórico de Ausências</p>
            <p align='center'><b>1.2.0000:</b><br>
            - Lançado Controle de histórico de Ausências</p>
            <p align='center'><b>1.1.0000:</b><br>
            - Lançado controle de Frequência Diário com controle faltas e atrasos diários<br>
            - Criado Atualizador Automático</p>
            <p align='center'><b>1.0.0000:</b><br>
            - Lançado controle de HEs semanais com dashboard e Geração de E-mails</p>
            """
        )
        label_texto.setOpenExternalLinks(True)  # Permite abrir links externos
        label_texto.setWordWrap(True)  # Habilita quebra de linha automática
        label_texto.setAlignment(Qt.AlignmentFlag.AlignTop)  # Alinha o texto no topo

        # Criar um QScrollArea para permitir rolagem
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)  # Permite redimensionamento
        scroll_area.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOn)  # Sempre exibe a barra de rolagem vertical
        scroll_area.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff)  # Oculta a barra de rolagem horizontal
        scroll_area.setWidget(label_texto)  # Define o QLabel dentro da área rolável

        # Adicionar o QScrollArea ao layout
        layout.addWidget(scroll_area)

        # Configurar o layout no diálogo
        dialog.setLayout(layout)

        # Exibir o diálogo (aguardando o usuário fechar)
        dialog.exec()

    def abrir_frequencia(self):
        from frequencia import controlefrequencia
        self.frequencia_window = controlefrequencia()
        self.frequencia_window.show()
        self.close()

    def abrir_horaextra(self):
        from horaextra import ControleHorasExtras
        self.horaextra_window = ControleHorasExtras()
        self.horaextra_window.show()
        self.close()

    def abrir_afastamentos(self):
        from afastamentos import Appatestados
        self.atestados_window = Appatestados()
        self.atestados_window.show()
        self.close()

    def abrir_advertencias(self):
        from advertencias import AppAdvertencias
        self.advertencias_window = AppAdvertencias()
        self.advertencias_window.show()
        self.close()

    def abrir_documentosvencidos(self):
        from documentosvencidos import AppConsultaDocumentos
        self.documentosvencido_window = AppConsultaDocumentos()
        self.documentosvencido_window.show()
        self.close()

    def abrir_painel_gestor(self):
        from Painel_Gestor import PainelConsultaFuncionarios
        self.painel_gestor_window = PainelConsultaFuncionarios()
        self.painel_gestor_window.show()
        self.close()

    def abrir_calculo_eventos(self):
        from Eventos_Folha import AppConsultaEventos
        self.calculo_eventos_window = AppConsultaEventos()
        self.calculo_eventos_window.show()
        self.close()

if __name__ == "__main__":
   # Define a flag necessária antes de criar a QApplication
    QApplication.setAttribute(Qt.ApplicationAttribute.AA_ShareOpenGLContexts)

    app = QApplication(sys.argv)
    app.setStyle("Fusion")  # Adapta ao tema do sistema

    main_window = ControlesRH()
    main_window.show()
    sys.exit(app.exec())

