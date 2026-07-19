# Letramento acadêmico

Material didático em formato de livro/site Quarto, desenvolvido no âmbito do **Grupo de Pesquisa em Estudos Linguísticos do Texto (GPELT)**, vinculado à Universidade Federal Rural do Semi-Árido — UFERSA.

Este repositório reúne roteiros de aula, complementos, avaliações, tutoriais, leituras e outros recursos voltados ao desenvolvimento de práticas de letramento acadêmico no ensino superior. O material foi estruturado em módulos, de modo a permitir sua organização em diferentes percursos formativos, conforme os objetivos de disciplinas, cursos, turmas e contextos institucionais, desde que observadas as restrições de uso indicadas neste repositório.

## Autoria e organização

**Prof. Mário Martins**

Grupo de Pesquisa em Estudos Linguísticos do Texto — GPELT\
Universidade Federal Rural do Semi-Árido — UFERSA

## Objetivo do projeto

O objetivo deste projeto é organizar, documentar e manter, em ambiente versionado, recursos didáticos voltados ao ensino de leitura, escrita, pesquisa e produção textual no ensino superior.

O material contempla práticas de escrita acadêmica, leitura crítica, planejamento textual, divulgação científica, elaboração de projetos de pesquisa, uso de ferramentas acadêmicas e participação em atividades universitárias de pesquisa, ensino e extensão.

Embora sua organização modular favoreça diferentes possibilidades de uso pedagógico, o conteúdo deste repositório não está automaticamente liberado para reutilização, adaptação ou redistribuição por terceiros.

## Estrutura do projeto

A organização das pastas segue uma lógica funcional, separando arquivos de configuração, aulas, avaliações, materiais complementares, tutoriais, arquivos de apoio e saída de publicação.

``` text
ANETE/
├── _arquivo/
├── aulas/
├── avaliacoes/
├── complementos/
├── docs/
├── img/
├── leituras/
├── pdfs/
├── tutoriais/
├── _quarto.yml
├── .gitignore
├── ANETE.Rproj
├── custom.css
├── custom.html
├── index.qmd
└── README.md
```

## Descrição das pastas e arquivos principais

### `_arquivo/`

Pasta destinada a materiais antigos, versões anteriores, testes, rascunhos ou arquivos que não fazem parte da versão ativa do projeto. Essa pasta funciona como área de preservação, evitando que arquivos pouco usados permaneçam soltos na raiz do projeto.

### `aulas/`

Contém os arquivos `.qmd` correspondentes aos roteiros de aula principais. Esses arquivos compõem o núcleo didático do livro/site.

### `avaliacoes/`

Contém avaliações, atividades avaliativas, provas, formulários de orientação e instrumentos relacionados ao acompanhamento da aprendizagem dos estudantes.

### `complementos/`

Contém materiais complementares às aulas, como guias, checklists, cronogramas, orientações específicas e recursos de apoio.

### `docs/`

Pasta de saída da renderização do projeto, quando configurada para publicação via GitHub Pages. Em projetos Quarto, essa pasta pode conter os arquivos HTML, CSS, imagens e demais recursos gerados automaticamente a partir dos arquivos `.qmd`.

### `img/`

Contém imagens utilizadas no projeto, como capas, figuras, ícones e demais recursos visuais.

### `leituras/`

Pasta destinada a referências, textos de apoio, bibliografias ou materiais de leitura utilizados no projeto.

### `pdfs/`

Contém arquivos em PDF utilizados como material de apoio, leitura ou documentação.

### `tutoriais/`

Contém tutoriais sobre ferramentas acadêmicas, tecnológicas ou institucionais relacionadas às práticas de leitura, escrita, pesquisa e produção textual no ensino superior.

### `_quarto.yml`

Arquivo principal de configuração do projeto Quarto. Nele são definidos aspectos como tipo de projeto, título, subtítulo, autoria, estrutura de navegação, ordem dos capítulos, tema visual, saída de renderização e demais parâmetros globais.

### `.gitignore`

Arquivo que define o que não deve ser enviado ao GitHub, como arquivos temporários do R, arquivos auxiliares de compilação e resíduos gerados automaticamente.

### `ANETE.Rproj`

Arquivo do projeto RStudio. Deve ser utilizado para abrir o projeto no RStudio, garantindo que o diretório de trabalho seja configurado corretamente.

### `custom.css`

Arquivo de customização visual do projeto. Pode conter definições de estilo, como cores, espaçamentos, fontes, caixas de destaque e ajustes de layout.

### `custom.html`

Arquivo de customização HTML, usado para incluir elementos adicionais no projeto, quando necessário.

### `index.qmd`

Arquivo inicial do projeto. Geralmente corresponde à página de apresentação ou abertura do livro/site.

## Como abrir o projeto no RStudio

Para trabalhar no projeto, abra o arquivo:

``` text
ANETE.Rproj
```

Abrir o projeto por esse arquivo ajuda a evitar problemas com caminhos relativos, renderização e localização de arquivos.

## Como renderizar o livro/site localmente

Para renderizar todo o projeto Quarto, execute no Terminal do RStudio, a partir da raiz do projeto:

``` bash
quarto render
```

Esse comando processa os arquivos `.qmd` e gera a versão final do livro/site, normalmente na pasta `docs/`, conforme configurado no arquivo `_quarto.yml`.

O comando `quarto render` apenas renderiza o projeto. Ele não abre automaticamente a visualização no navegador.

## Como pré-visualizar o livro/site localmente

Para visualizar o livro/site localmente durante a edição, execute no Terminal do RStudio, a partir da raiz do projeto:

``` bash
quarto preview
```

Esse comando abre uma versão local do livro/site no navegador e permite verificar se as páginas, links, imagens, estilos e menus estão funcionando corretamente antes da publicação.

Não é necessário usar:

``` bash
quarto preview docs
```

Neste projeto, a pasta `docs/` é a pasta de saída da renderização, não a pasta-fonte. O comando `quarto preview` deve ser executado na raiz do projeto, onde estão os arquivos `_quarto.yml`, `index.qmd` e `ANETE.Rproj`.

Caso o navegador não abra automaticamente, copie a URL exibida no Terminal, geralmente em formato semelhante a:

``` text
http://localhost:xxxx/
```

e cole manualmente no navegador.

## Fluxo recomendado de atualização

Sempre que desejar atualizar o conteúdo do livro/site, siga este procedimento:

1.  Edite os arquivos `.qmd` normalmente no RStudio.
2.  Salve as alterações.
3.  Pré-visualize localmente, se quiser acompanhar as mudanças durante a edição:

``` bash
quarto preview
```

4.  Quando a versão estiver pronta, renderize o projeto:

``` bash
quarto render
```

5.  Verifique se a renderização ocorreu sem erros.
6.  Envie as alterações ao GitHub:

``` bash
git add .
git commit -m "Atualiza conteúdo do livro"
git push origin main
```

## Atualização de caminhos no `_quarto.yml`

Sempre que um arquivo `.qmd` for movido para uma nova pasta, o caminho correspondente deve ser atualizado no arquivo `_quarto.yml`.

Exemplo:

``` yaml
- aulas/Aula_1.qmd
- complementos/complemento_painel.qmd
- avaliacoes/avaliacao_terceira_2026_1.qmd
```

Se o caminho estiver incorreto, o Quarto poderá apresentar erro durante a renderização.

## Caminhos internos em arquivos `.qmd`

Como os arquivos estão organizados em subpastas, links internos e caminhos de imagens devem considerar a localização do arquivo que faz a chamada.

Por exemplo, dentro de um arquivo localizado na pasta `aulas/`, uma imagem localizada na pasta `img/` deve ser chamada assim:

``` r
knitr::include_graphics("../img/nome_da_imagem.png")
```

Do mesmo modo, um link de uma aula para um tutorial deve considerar a subida de diretório:

``` markdown
[Consultar tutorial](../tutoriais/tutorial_zotero.qmd)
```

E um link de uma aula para um complemento deve seguir a mesma lógica:

``` markdown
[Consultar complemento](../complementos/complemento_painel.qmd)
```

## Renderização de avaliações em HTML e PDF

Para renderizar uma avaliação específica em HTML, utilize:

``` bash
quarto render caminho/do/arquivo.qmd
```

Exemplo:

``` bash
quarto render avaliacoes/prova_unidade1.qmd
```

Depois, para converter o HTML em PDF usando o pacote `pagedown`, execute no Console do R:

``` r
pagedown::chrome_print("avaliacoes/prova_unidade1.html")
```

Esse procedimento é útil para gerar versões em PDF de provas, atividades ou instrumentos avaliativos elaborados em Quarto.

## Publicação no GitHub Pages

Caso o projeto esteja configurado para gerar a saída em `docs/`, essa pasta pode ser usada para publicação via GitHub Pages.

No GitHub, a configuração deve ser feita no repositório:

``` text
https://github.com/gpeltufersa/anete
```

Em seguida, acesse:

``` text
Settings → Pages
```

Na seção de publicação, selecione:

``` text
Source: Deploy from a branch
Branch: main
Folder: /docs
```

Depois disso, o GitHub Pages publicará o conteúdo renderizado da pasta `docs/`.

A URL pública esperada do projeto é:

``` text
https://gpeltufersa.github.io/anete/
```

Antes de publicar, verifique se o arquivo abaixo existe no repositório:

``` text
docs/index.html
```

Esse arquivo funciona como página inicial do livro/site publicado.

## Cuidados antes de publicar

Antes de enviar alterações ao GitHub, recomenda-se verificar:

- se o projeto renderiza sem erros;
- se os links internos funcionam;
- se as imagens aparecem corretamente;
- se os arquivos sensíveis não foram incluídos;
- se a pasta `docs/` está atualizada, caso seja usada para publicação;
- se os caminhos no `_quarto.yml` estão corretos;
- se arquivos antigos ou rascunhos foram mantidos em `_arquivo/`, e não na raiz do projeto.

## Recomendações de manutenção

Para manter o projeto organizado ao longo dos semestres:

- mantenha na raiz apenas arquivos essenciais de configuração e documentação;
- crie novos roteiros de aula dentro da pasta `aulas/`;
- crie novas avaliações dentro da pasta `avaliacoes/`;
- coloque materiais auxiliares em `complementos/`;
- mova versões antigas para `_arquivo/`;
- evite nomes de arquivos com espaços, acentos ou caracteres especiais;
- prefira nomes curtos, descritivos e padronizados;
- atualize o `_quarto.yml` sempre que criar, remover ou mover páginas;
- renderize o projeto antes de publicar alterações.

## Convenção sugerida para nomes de arquivos

Para manter o projeto consistente, recomenda-se usar nomes em minúsculas, sem acentos e com separação por underline.

Exemplos:

``` text
aulas/aula_01_introducao.qmd
aulas/aula_02_ciencia_senso_comum.qmd
complementos/complemento_cronograma.qmd
avaliacoes/avaliacao_unidade_1.qmd
tutoriais/tutorial_zotero.qmd
```

## Observação sobre arquivos gerados automaticamente

Arquivos como os listados abaixo geralmente não precisam ser mantidos manualmente no projeto, pois são gerados durante processos de compilação:

``` text
*.aux
*.log
*.toc
*.tex
```

Esses arquivos devem ser ignorados pelo Git, salvo em situações específicas em que haja necessidade técnica de preservá-los.

## Licença e uso

Todos os direitos reservados.

O conteúdo deste repositório é disponibilizado exclusivamente para fins de consulta, organização interna, manutenção do projeto e uso autorizado pelos autores e pelo **Grupo de Pesquisa em Estudos Linguísticos do Texto (GPELT)**.

Não é permitida a reprodução, redistribuição, adaptação, remixagem, publicação, incorporação em outros materiais, uso em disciplinas externas, uso comercial ou disponibilização pública deste conteúdo, total ou parcial, sem autorização prévia e expressa dos autores responsáveis.

A ausência de um arquivo de licença aberta, como `LICENSE`, significa que este repositório **não concede permissão automática de uso, cópia, modificação ou redistribuição**.

Para solicitar autorização de uso, adaptação ou compartilhamento do material, é necessário entrar em contato com os responsáveis pelo projeto.
