# Conciliação Bancária

Sistema web para automatizar a conciliação entre o **extrato bancário** e o **relatório de pagamentos a fornecedores**, gerando automaticamente os lançamentos contábeis correspondentes.

## O problema que resolve

Conferir manualmente, linha por linha, se cada pagamento a fornecedor efetivamente saiu no extrato do banco — e depois lançar isso na contabilidade — é um processo repetitivo e sujeito a erro humano, especialmente com volumes grandes de transações. Esse sistema:

- Lê os dois arquivos (extrato e fornecedores), em **CSV ou Excel**, com qualquer nomenclatura comum de coluna (o app tenta reconhecer variações como "valor", "vlr", "data pagto", "razão social", etc.);
- Cruza automaticamente os itens por **documento + valor** (match forte) e, quando isso não é possível, por **valor + similaridade do nome/descrição** (match por aproximação), considerando também a proximidade de datas como critério de desempate;
- Distingue **créditos e débitos** no extrato, evitando comparar pagamentos a fornecedores com entradas de dinheiro;
- Separa o que foi conciliado automaticamente do que ficou **pendente**, permitindo revisão e **correção manual** de qualquer item conciliado;
- Permite **forçar um match manual** entre um item do extrato e um item de fornecedores que não bateram automaticamente;
- Gera os **lançamentos contábeis** (débito/crédito/participante/valor) prontos para exportação em CSV;
- Mantém um **histórico de conciliações** salvas, consultável posteriormente.

## Como funciona hoje (front-end)

Todo o processamento roda **no navegador**, sem backend:

- **`index.html`** — estrutura das telas (upload, resultado, histórico) e dos modais (correção manual e forçar conciliação);
- **`style.css`** — estilos visuais;
- **`app.js`** — toda a lógica: leitura de CSV/Excel (via [SheetJS](https://sheetjs.com/)), normalização dos dados, algoritmo de matching, renderização das tabelas e persistência do histórico em `localStorage`.

### Fluxo de uso

1. Na tela **Nova Conciliação**, envie o arquivo do extrato bancário e o do relatório de fornecedores (CSV ou Excel).
2. Clique em **Conciliar Automaticamente**.
3. Na tela de **Resultado**, revise:
   - **Conciliados**: itens que bateram automaticamente (pode corrigir qualquer campo);
   - **Pendentes**: itens de um lado que não encontraram par no outro (pode forçar um vínculo manual);
   - **Lançamentos Contábeis**: o resultado pronto para exportar em CSV.
4. Salve a conciliação no **Histórico** para consultar depois.

## Algoritmo de conciliação (resumo)

1. **Match forte** — mesmo número de documento e mesmo valor (tolerância de R$ 0,02).
2. **Match por aproximação** — valores iguais (mesma tolerância) e similaridade de nome/descrição acima de um limiar mínimo; entre todos os pares candidatos possíveis, os de maior score (similaridade + proximidade de data) são priorizados, evitando que a ordem das linhas na planilha prejudique o resultado.
3. **Pendentes** — tudo que sobrou de ambos os lados vai para revisão manual.

## Roadmap — Backend (em andamento)

O projeto está evoluindo de uma aplicação 100% client-side para uma arquitetura com backend em **Java + Spring Boot**, com os seguintes objetivos:

- API REST para persistir as conciliações (substituindo o `localStorage` atual), permitindo acesso a partir de diferentes computadores;
- Autenticação de usuários;
- Banco de dados relacional (iniciando em H2 para desenvolvimento, com caminho para PostgreSQL/MySQL em produção).

A lógica de leitura de arquivos e matching deve permanecer no front-end (JavaScript) por enquanto; o backend entra para persistência e controle de acesso.

## Tecnologias

**Front-end (atual):**
- HTML5, CSS3, JavaScript (vanilla, sem frameworks)
- [SheetJS (xlsx)](https://sheetjs.com/) para leitura de arquivos Excel

**Backend (planejado):**
- Java + Spring Boot
- Spring Security (autenticação)
- Spring Data JPA + H2 (dev) / PostgreSQL ou MySQL (produção)

## Estrutura do projeto

```
├── index.html      # Telas e modais
├── style.css       # Estilos
└── app.js          # Lógica de leitura, matching, renderização e histórico
```

## Formatos de arquivo suportados

Tanto o extrato quanto o relatório de fornecedores podem ser enviados como `.csv`, `.xlsx` ou `.xls`. O app tenta reconhecer automaticamente as colunas relevantes (data, documento, valor, descrição/fornecedor) mesmo com nomes de cabeçalho variados.
