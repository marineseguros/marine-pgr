# Gerador de PGR · Marine Seguros — v1.2.4

Gera o PGR (Programa de Gerenciamento de Risco) no modelo timbrado da Marine a partir
da apólice, proposta ou endosso de RC-DC da Tokio Marine. O texto é copiado do PDF por
código, palavra por palavra, sem redigitar e sem IA. Depois de gerar, o app confere o
.docx contra o PDF.

## Para quem usa (Helloisy, Thayla)

1. Abra **Marine PGR.exe**.
2. Arraste o PDF da apólice, proposta ou endosso para a área Adicionar documentos, ou use Arquivos para
   escolher. Dá para colocar vários de uma vez.
3. Clique em **Gerar PGR**.
4. Veja o status na fila:
   - **✔ Conferido**: pronto para enviar à gerenciadora.
   - **⚠ Revisar**: falta algum dado do cabeçalho ou alguma seção. O painel Conferência
     mostra o que faltou.
   - **✖ Erro**: o PDF não foi reconhecido. Hoje o app reconhece só o layout da Tokio Marine.
5. Dê dois cliques no item, ou use **Abrir PGR**, para abrir o Word.

O arquivo é salvo como `PGR <SEGURADO> - MM_AAAA.docx` na mesma pasta do PDF. Se já
existir um com esse nome, o app cria `(Cópia)`. Ele nunca sobrescreve.

**Importante:** use sempre o documento mais recente do cliente. Se houve endosso, por
exemplo para incluir animais vivos, gere a partir do endosso.

**Seção com outro nome:** algumas propostas escrevem o título de outro jeito, por exemplo
"LIMITES PARA O PLANO DE GERENCIAMENTO DE RISCO" em vez de "GERENCIAMENTO DE RISCO". Nesse
caso abre uma janela para escolher o título certo na lista (são os títulos que existem no
PDF) ou digitá-lo exatamente igual. O app gera de novo no mesmo arquivo e, se a opção
"Lembrar" estiver marcada, aprende esse nome para as próximas propostas. A opção
"Não incluir esta seção" deixa a seção de fora de propósito.

**PDF sem texto:** se a proposta passou por "Microsoft Print to PDF" ou foi escaneada, o
app avisa. Nesse caso use o PDF original da Tokio. O app não faz leitura de imagem (OCR).

Em **Configurações** dá para escolher outro modelo timbrado (.docx) e uma pasta fixa
para salvar os arquivos, e também apagar os títulos de seção aprendidos.

## Para gerar o .exe (Ernane)

Requisitos: Windows com Python 3.10 ou mais novo. Marque "Add to PATH" na instalação.

1. Dê dois cliques em `build_exe.bat`. Ele instala as dependências e roda o PyInstaller.
2. O executável aparece em `dist\Marine PGR.exe` (arquivo único, tamanho variável conforme as dependências).
3. Coloque o .exe numa pasta compartilhada do Drive. Para atualizar, substitua o arquivo.

Para testar sem gerar o .exe: `pip install -r requirements.txt` e depois `rodar_sem_build.bat`.

O antivírus pode estranhar um .exe do PyInstaller sem assinatura na primeira vez. Se
acontecer, libere o arquivo ou distribua a pasta (troque `--onefile` por `--onedir` no .bat).

## Estrutura

```
main.py                 ponto de entrada
src/pgr_core.py         motor: extração, montagem do .docx e conferência literal
src/app_gui.py          interface (Tkinter + tkinterdnd2)
src/app_version.py      nome e versão
resources/              modelo timbrado, logo e ícone (vão dentro do .exe)
```

`pgr_core.py` também funciona sozinho pela linha de comando:

```
python src/pgr_core.py --proposta proposta.pdf --modelo resources/modelo_pgr.docx --saida PGR.docx
python src/pgr_core.py --proposta proposta.pdf --modelo x --saida x --listar   (lista as seções do PDF)
```

## O que entra no PGR

Limite máximo de garantia (com as notas) · Mercadorias excluídas · Importância segurada ·
Objeto do seguro · Viagens · Participação obrigatória do segurado · Gerenciamento de risco
(completo) · Sinistros.

No cabeçalho entram: vigência (da seção VIGÊNCIA), segurado (proponente), corretora e seguradora.

Para mudar a seleção ou a ordem, edite `SECOES_PADRAO` em `src/pgr_core.py`.

## Limitações atuais

- Só o layout da **Tokio Marine** (títulos em faixa cinza). Para outra seguradora, é
  preciso um par "PDF + PGR feito à mão" para adaptar.
- Validado com a Schoeler (proposta 11389237-v2): 3.140 palavras conferidas, 0 divergência.

