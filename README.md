# Excel para Recrutamento

Tutorial completo de Excel para quem trabalha com Recrutamento e Seleção, do básico ao painel de indicadores. Funciona no computador e no celular.

**Acesse:** https://kaique-fhs.github.io/tutorial-excel/

## O que tem

- 24 módulos: estrutura da base de candidatos, Tabelas, listas suspensas, limpeza de dados, fórmulas (SE, SES, CONT.SES, MÉDIASES), datas e SLA com dias úteis, PROCX, FILTRO/ÚNICO, formatação condicional, filtros e segmentação, tabela dinâmica, gráficos, indicadores (time to hire, time to fill, funil, fontes, aceite de proposta, custo por contratação), painel, scorecard de entrevista, trabalho em equipe, LGPD, Excel no celular, atalhos, erros e exercícios com gabarito.
- Fórmulas em português (separador `;`) com botão de copiar.
- Simulador de funil, quizzes, progresso salvo no navegador, busca, tema claro/escuro.
- Planilha de prática com dados fictícios: `arquivos/planilha-pratica-recrutamento.xlsx` (abas Painel, Candidatos, Vagas, Scorecard, Dados_Brutos, Exercícios, Listas).

## Estrutura

```
index.html                         página do tutorial
css/estilo.css                     estilos (cor principal azul, responsivo)
js/app.js                          interações (sem dependências)
arquivos/planilha-pratica-...xlsx  planilha de prática
ferramentas/gerar_planilha.py      gera a planilha (python + openpyxl)
```

## Atualizar a planilha de prática

```
pip install openpyxl
python ferramentas/gerar_planilha.py
```

## Publicação

GitHub Pages, branch `main`, pasta raiz. Todo push na `main` atualiza o site.
