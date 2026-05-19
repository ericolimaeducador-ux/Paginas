# Hub de Materiais

Landing page estatica para publicar artigos cientificos, imagens, videos,
projetos e respostas a questionamentos via GitHub Pages.

## Estrutura

- `index.html`: estrutura da pagina.
- `assets/css/styles.css`: estilos visuais e responsividade.
- `assets/js/main.js`: menu responsivo e cards iniciais.
- `Respostas a questionamentos/`: materiais da categoria de respostas tecnicas.
- `.github/workflows/pages.yml`: deploy automatico para GitHub Pages via Actions.
- `.nojekyll`: impede que o GitHub Pages processe o site com Jekyll.

## Primeiro material

O primeiro material esta em:

- `Respostas a questionamentos/parecer-tecnico-anvisa-sonda-npag.html`
- `Respostas a questionamentos/parecer-tecnico-anvisa-sonda-npag.js`

Para gerar o `.docx`, instale a dependencia `docx` e execute:

```bash
npm install docx
node "Respostas a questionamentos/parecer-tecnico-anvisa-sonda-npag.js"
```

## Publicar no GitHub Pages

No GitHub, acesse `Settings` > `Pages` e escolha `GitHub Actions` como origem
do deploy. A cada push no branch `main`, o workflow publicara a pagina.
