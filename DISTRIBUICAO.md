# Distribuição do Marine PGR

## Gerar no seu computador

1. Extraia todo o projeto. Para compilar, use Windows com Python 3.10 ou superior.
2. Abra `EXECUTADOR_DE_DISTRIBUICAO.bat`. Ele prepara um ambiente separado, instala as dependências e gera o aplicativo.
3. A pasta `distribuicao` terá o ZIP para os usuários e o arquivo SHA-256 para conferência.
4. O usuário extrai o ZIP e abre `Marine PGR.exe`. Não precisa instalar Python.

O executador de distribuição é uma ferramenta de quem mantém o projeto. O arquivo entregue aos usuários é o ZIP que ele gera.

## Publicar no GitHub

Envie o conteúdo desta pasta para a raiz do repositório, incluindo a pasta oculta `.github`. Não envie `.build-env`, `build` ou `dist`.

Crie uma Release com a tag correspondente à versão de `src/app_version.py` (por exemplo, `v1.2.4`). Anexe o ZIP de `distribuicao` e o SHA-256. Compartilhe a página dessa Release com os usuários. Subir somente o código não cria o executável.

Também está incluído um fluxo GitHub Actions. Na aba Actions, execute **Gerar distribuição Windows**. Ele gera o ZIP e o SHA-256 para baixar como artefato. Esse fluxo prepara os arquivos; a publicação da Release continua manual.

Atualização para usuários: baixar a nova versão, fechar o aplicativo antigo e abrir o novo executável. Configurações existentes continuam na pasta PGR Marine em APPDATA. Esta versão não inclui atualização automática.
