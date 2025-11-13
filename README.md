<p>Aplicação desktop completa (Python + PyQt6) utilizada pelo RH para consultar, processar, analisar e gerar relatórios a partir dos dados do HCM Senior.<br>
Os dados são extraídos diretamente do banco Oracle (tabelas R034FUN, R038AFA, R070ACC, R016HIE, etc.) e também de Relatórios do HCM Senior.</p>

<p>O sistema gera dashboards gráficos, PDFs, emails automáticos via Outlook, tabelas dinâmicas, estatísticas, cálculos e painéis completos de gestão.</p>

<h2>Principais Funcionalidades</h2>

<h3>📌 1. Frequência (Ponto Eletrônico)</h3>
<p>Módulo principal.</p>
<ul>
  <li>Consulta batidas (R070ACC).</li>
  <li>Identifica atrasos, faltas, banco de horas, escalas.</li>
  <li>Mostra expediente dia/hora, legenda por cores, filtros dinâmicos.</li>
  <li>Permite gerar:
    <ul>
      <li>Gráficos plotly</li>
      <li>PDF</li>
      <li>E-mail automático com imagem inline</li>
    </ul>
  </li>
  <li>Exibe histórico completo (planilhas AUSÊNCIAS / HRAP).</li>
</ul>
<p><strong>Arquivos:</strong> frequencia.py, hisfrequenciagrafico.py, historicofrequencia.py.</p>

<h3>📌 2. Assiduidade</h3>
<p>Processa planilhas:</p>
<ul>
  <li>HRAP604</li>
  <li>AUSÊNCIAS DIÁRIO / SEMANAL / MENSAL</li>
</ul>
<p>Gera:</p>
<ul>
  <li>Total de horas trabalhadas</li>
  <li>Total de horas de atestados</li>
  <li>Dias trabalhados e dias de atestado</li>
  <li>Gráficos interativos</li>
  <li>E-mail detalhado e resumo</li>
</ul>
<p><strong>Arquivos:</strong> Assiduidade.py, Assinuidade_Atestados.py, Assinuidade_Atestados_Grafico.py.</p>

<h3>📌 3. Horas Extras</h3>
<p>Processa planilhas oficiais:</p>
<ul>
  <li>HRAP601 (HE diário)</li>
  <li>HE SEMANAL</li>
  <li>HE MENSAL</li>
  <li>FPRE905 (apoio)</li>
</ul>
<p>Funções:</p>
<ul>
  <li>Calcula HE por colaborador</li>
  <li>Calcula DSR (mensal)</li>
  <li>Exporta Excel</li>
  <li>Painel gráfico completo</li>
  <li>Controle de feriados</li>
  <li>Geração de e-mail automático</li>
</ul>
<p><strong>Arquivos:</strong> horaextra.py, horaextragrafico.py.</p>

<h3>📌 4. Afastamentos (Atestados / Licenças)</h3>
<p>Consulta Oracle:</p>
<ul>
  <li>Afastamentos correntes</li>
  <li>Afastamentos iniciados</li>
  <li>Afastamentos por período</li>
  <li>Lista completa de SITAFA (14, 64, 20, 3, 4, 61, 913, 918 etc.)</li>
</ul>
<p>Gera:</p>
<ul>
  <li>Tabela agregada por colaborador</li>
  <li>Tabela detalhada por duplo clique</li>
  <li>Gráfico interativo</li>
  <li>E-mail com dados</li>
</ul>
<p><strong>Arquivos:</strong> afastamentos.py, afastamentosgrafico.py.</p>

<h3>📌 5. Documentos Vencidos (CNH, RG/CIN)</h3>
<p>Consulta Oracle:</p>
<ul>
  <li>CNH (VENCNH)</li>
  <li>RG/CIN (DEXCID)</li>
  <li>Filtra locais e setores com dicionário LOCAIS.xlsx</li>
  <li>Destaca vencidos e prestes a vencer</li>
  <li>Envia e-mail (agrupado ou por colaborador)</li>
</ul>
<p><strong>Arquivo:</strong> documentosvencidos.py.</p>

<h3>📌 6. Retornos, Férias Vencendo e Vistos</h3>
<p>Dashboard rápido com cards:</p>
<ul>
  <li>Término de experiência</li>
  <li>Retornos de afastamento</li>
  <li>Férias próximas de vencer</li>
  <li>Vistos vencendo</li>
</ul>
<p>Mostra quantidade por categoria e abre detalhes.</p>
<p><strong>Arquivo:</strong> telaaviso.py.</p>

<h3>📌 7. Painel de Funcionários e Gestores</h3>
<p>Consulta Oracle e monta painel detalhado:</p>
<ul>
  <li>Funcionários ativos</li>
  <li>Desligados</li>
  <li>Admitidos no período</li>
  <li>Setores (agrupamento)</li>
  <li>Gestores (quantidade, colaboradores por gestor)</li>
  <li>Organograma via GraphViz</li>
  <li>Exportação para Excel</li>
  <li>Abas: Funcionários / Setores / Gestores</li>
  <li>Gráficos por setor / gestor</li>
</ul>
<p><strong>Arquivos:</strong> Painel_Gestor.py, Painel_Gestores.py, Painel_Setores.py, Painel_Setores_Grafico.py.</p>

<h3>📊 Relatórios e Gráficos</h3>
<ul>
  <li>Plotly</li>
  <li>PyQt6 + QWebEngineView</li>
  <li>Exportação para PDF</li>
  <li>Imagens inline em e-mail Outlook (win32com)</li>
</ul>

<h3>📧 Geração Automática de E-mail (Outlook)</h3>
<ul>
  <li>Geração automática de e-mail</li>
  <li>Anexo de gráficos inline</li>
  <li>Inserção de texto automático</li>
  <li>Anexo de PDF e Excel</li>
</ul>

<h3>📂 Acesso ao Banco Senior (Oracle)</h3>
<p><strong>Arquivo principal:</strong> Database.py</p>
<ul>
  <li>SessionPool Oracle (oracledb)</li>
  <li>Reaproveitamento de conexões</li>
  <li>Função get_connection() com contexto</li>
</ul>

<p><strong>Tabelas utilizadas:</strong></p>
<ul>
  <li>R034FUN / R034CPL</li>
  <li>R038AFA</li>
  <li>R070ACC</li>
  <li>R016HIE</li>
  <li>R024CAR</li>
  <li>R038HSA</li>
  <li>R192DOE</li>
  <li>R010SIT</li>
</ul>

<img width="1356" height="762" alt="Captura de tela 2025-11-12 223402" src="https://github.com/user-attachments/assets/9b2aeaf7-fbf3-4109-b966-75ff97a7a9ec" />
<img width="1910" height="797" alt="Captura de tela 2025-11-12 223438" src="https://github.com/user-attachments/assets/f929ded9-4613-4974-9fff-86fe33a64e29" />


<h3>✔️ O que esta aplicação entrega para o RH</h3>
<ul>
  <li>Controle completo de frequência</li>
  <li>Apoio no fechamento do ponto</li>
  <li>Gestão de HE/DSR</li>
  <li>Controle de documentos obrigatórios</li>
  <li>Controle de afastamentos</li>
  <li>Painel de colaboradores e gestores</li>
  <li>Dashboards para diretoria</li>
  <li>E-mails automáticos</li>
  <li>Relatórios PDF e Excel</li>
</ul>

<p>Necessário revisar cada rotina, pois a maioria depende de relatórios específicos gerados no ERP que estão na pasta 'Dependencias'. <br>
Eles devem ser importados e parametrizados para gerar os arquivos nos diretórios lidos pela aplicação. <br>
Também é necessário configurar a conexão no arquivo <strong>Database.py</strong>.</p>
