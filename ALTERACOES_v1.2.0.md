# Marine PGR · v1.2.0

Interface adaptada a partir do Marine Docs v3.1.1 fornecido como referência.

- Mesma paleta clara, marca oficial, cartões arredondados e hierarquia de títulos.
- Organização em duas colunas: compatibilidade, entrada compacta e fila à esquerda; destino e conferência à direita.
- Barra inferior com Gerar PGR, Abrir pasta, Abrir PGR, status e Limpar fila.
- Seleção de PDFs por arquivos, pasta ou arrastar e soltar (com tkinterdnd2 instalado).
- Atalhos Ctrl+O para adicionar arquivos e Ctrl+Enter para gerar.
- Rolagem na conferência e indicação de fila vazia.
- Bloqueio de geração simultânea e prevenção de duplicatas na mesma seleção.

O motor pgr_core.py e o modelo_pgr.docx foram preservados byte a byte. Configurações e títulos aprendidos continuam usando o local original.

Os módulos de faturamento, InCorp, Google Drive, banco de dados e atualização do Marine Docs não fazem parte do Gerador de PGR; as ações exibidas correspondem às funções do PGR.

Validação realizada: sintaxe dos arquivos Python, abertura da interface, inclusão de arquivo, duplicatas, seleção, limpeza e verificações básicas de dimensões. Sem uma apólice de exemplo, não foi executada uma geração completa de PGR nesta revisão.

Para executar: instale as dependências de requirements.txt e abra rodar_sem_build.bat. Para gerar o executável Windows, abra build_exe.bat. O ZIP contém o projeto com os scripts de execução e compilação; não contém executável pré-compilado.
