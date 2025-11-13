Aplicação desktop completa (Python + PyQt6) utilizada pelo RH para consultar, processar, analisar e gerar relatórios a partir dos dados do HCM Senior.
Os dados são extraídos diretamente do banco Oracle (tabelas R034FUN, R038AFA, R070ACC, R016HIE, etc.) e também de relatórios do HCM .
O sistema gera dashboards gráficos, PDFs, emails automáticos via Outlook, tabelas dinâmicas, estatísticas, cálculos e painéis completos de gestão.
________________________________________
📌 Principais Funcionalidades
1. Frequência (Ponto Eletrônico)
Módulo principal.
Funções:
•	Consulta batidas (R070ACC).
•	Identifica atrasos, faltas, banco de horas, escalas.
•	Mostra expediente dia/hora, legenda por cores, filtros dinâmicos.
•	Permite gerar:
o	Gráficos plotly
o	PDF
o	E-mail automático com imagem inline
•	Exibe histórico completo (planilhas AUSÊNCIAS / HRAP).
Arquivos relacionados:
frequencia.py, hisfrequenciagrafico.py, historicofrequencia.py.
________________________________________
2. Assiduidade
Processa planilhas:
•	HRAP604
•	AUSÊNCIAS DIÁRIO / SEMANAL / MENSAL
Gera:
•	Total de horas trabalhadas
•	Total de horas de atestados
•	Dias trabalhados e dias de atestado
•	Gráficos interativos
•	E-mail detalhado e resumo
Arquivos:
Assiduidade.py, Assinuidade_Atestados.py, Assinuidade_Atestados_Grafico.py.
________________________________________
3. Horas Extras
Processa planilhas oficiais:
•	HRAP601 (HE diário)
•	HE SEMANAL
•	HE MENSAL
•	FPRE905 (apoio)
Funções:
•	Calcula HE por colaborador
•	Calcula DSR (mensal)
•	Exporta Excel
•	Painel gráfico completo
•	Controle de feriados
•	Geração de e-mail automático
Arquivos:
horaextra.py, horaextragrafico.py.
________________________________________
4. Afastamentos (Atestados / Licenças)
Consulta Oracle:
•	Afastamentos correntes
•	Afastamentos iniciados
•	Afastamentos por período
•	Lista completa de SITAFA (14, 64, 20, 3, 4, 61, 913, 918 etc.)
Gera:
•	Tabela agregada por colaborador
•	Tabela detalhada por duplo clique
•	Gráfico interativo
•	E-mail com dados
Arquivos:
afastamentos.py, afastamentosgrafico.py.
________________________________________
5. Documentos Vencidos (CNH, RG/CIN)
Consulta Oracle:
•	CNH (VENCNH)
•	RG/CIN (DEXCID)
•	Filtra locais e setores com dicionário LOCAIS.xlsx
•	Destaca vencidos e prestes a vencer
•	Envia e-mail (agrupado ou por colaborador)
Arquivo:
documentosvencidos.py.
________________________________________
6. Retornos, Férias Vencendo e Vistos
Dashboard rápido com cards:
•	Término de experiência
•	Retornos de afastamento
•	Férias próximas de vencer
•	Vistos vencendo
Mostra quantidade por categoria e abre detalhes.
Arquivo:
telaaviso.py.
________________________________________
7. Painel de Funcionários e Gestores
Consulta Oracle e monta um painel detalhado:
•	Funcionários ativos
•	Desligados
•	Admitidos no período
•	Setores (agrupamento)
•	Gestores (quantidade, colaboradores por gestor)
•	Organograma via GraphViz
•	Exportação para Excel
•	Abas separadas: Funcionários / Setores / Gestores
•	Gráficos por Setor / Gestor
Arquivos:
Painel_Gestor.py,
Painel_Gestores.py,
Painel_Setores.py,
Painel_Setores_Grafico.py.
________________________________________
📊 Relatórios e Gráficos
Todos os módulos possuem gráficos gerados com:
•	Plotly
•	PyQt6 + QWebEngineView
•	Exportação para PDF
•	Inserção inline em e-mail Outlook (via win32com)
________________________________________
📧 Geração Automática de E-mail (Outlook)
Integrado com win32com.client.
Todos os módulos conseguem:
•	Gerar e-mail automático
•	Anexar gráficos no corpo do e-mail (inline)
•	Inserir texto automático (período, resumo, totais)
•	Adicionar anexos Excel ou PDF
________________________________________
📂 Acesso ao Banco Senior (Oracle)
Arquivo principal:
Database.py
•	Cria SessionPool Oracle (oracledb)
•	Reaproveita conexões
•	Função get_connection() para uso em qualquer módulo
•	Conexão segura com contexto (with)
Tabelas Senior utilizadas:
•	R034FUN / R034CPL
•	R038AFA
•	R070ACC
•	R016HIE
•	R024CAR
•	R038HSA
•	R192DOE
•	R010SIT
•	Diversas dependentes de consultas específicas

O que esta aplicação entrega para o RH
•	Controle completo de frequência
•	Apoio para fechamento do ponto
•	Gestão de HE/DSR
•	Controle de documentos obrigatórios
•	Controle de afastamentos
•	Painel de colaboradores e gestores
•	Dashboards para diretoria
•	Emissão de e-mails automatizados
•	Geração de relatórios PDF e Excel

<img width="1254" height="858" alt="Captura de tela 2025-11-12 200132" src="https://github.com/user-attachments/assets/0c4b6f8d-cb14-433c-a0c4-2313108217c7" />

▶️ Como Executar
Requisitos
•	Python 3.10+
•	Oracle Instant Client
•	Instalar dependências:
pip install pyqt6 plotly pysqlite3 oracledb pandas openpyxl win32com unidecode workalendar graphviz orjson

Necessario revisar cada arquivo de rotina, pois a maior parte possuem necessidade de relátórios especificos, os quais estão em anexos junto com suas regras de geração automaticas.
Devem ser importados e parametrizados no ERP para gerarem no diretorio o qual a aplciação irá procurar.
tTambem necessario Configurar a conexão no arquivo Database.py
