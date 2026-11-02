# Verificação local

Revisão: 2 de novembro de 2026, Windows e Python 3.12.9.

- Suíte: 64 testes descobertos; 62 aprovados, 2 ignorados por dependerem de dados privados indisponíveis nesta cópia; nenhuma falha.
- Comando: `python -m unittest discover -s tests -v`, usando ambiente com dependências fixadas nos arquivos requirements.
- `pip check`: nenhuma dependência incompatível.
- Instalação sem acesso à rede confirmou todas as versões já presentes. Não se afirma instalação limpa a partir da internet.
- Inicialização real de `gui.run()`, importando main, com Qt offscreen: código de saída 0.
- main.py, gui.py e controller.py comparados por SHA-256: idênticos à versão local de origem.
- Todos os caminhos relativos de imagens e documentos no README existem; bloco principal usa banner original, verificado por SHA-256. README convertido com o processador Markdown do PowerShell; tabelas, imagens e blocos de código preservados. Estrutura revisada. O README publicado apresenta as seções, tabelas e o diagrama Mermaid; a conferência das imagens integra a publicação.
- Manual com 12 páginas, renderizadas e revisadas visualmente, sem cadastros reais.
- Varredura da seleção textual não encontrou padrões comuns de tokens GitHub, chaves AWS ou chaves privadas. A seleção explícita exclui bases reais, histórico anterior e arquivos pessoais. Não representa garantia absoluta contra qualquer segredo.
- A cópia pública não depende dos notebooks ou do validador dos organizadores.
- A configuração onedir foi preservada com remoção apenas da cópia de planilhas privadas. Não foi gerado novo executável nesta preparação.

Autorizações registradas em PUBLICACAO.md. Publicação autorizada na conta KamillyFer-8.

- Na preparação inicial, o Git foi inicializado na branch main sem histórico anterior. Na seleção revisada, nenhum Excel, ambiente, executável, arquivo .env ou histórico anterior rastreado. Git diff --check aprovado com tolerância a CRLF e linhas finais preexistentes.
- O launcher `py -3.12` não está registrado neste computador; as instruções usam `python`, após instalação de Python 3.12.9 no PATH. Ambiente novo criado com o interpretador 3.12.9 local.
