# Relatório v19 — Coletas por autor

## Ajustes realizados

1. Foram adicionadas três novas opções na área de **Coleta**:
   - **Obras Rubem Grilo**
   - **Obras Milan Dusek**
   - **Obras Cícero Dias**

2. Foram criados três scrapers simples e específicos por autor:
   - `scrapers/rubem_griloWS_sem_miniaturas.py`
   - `scrapers/milan_dusekWS_sem_miniaturas.py`
   - `scrapers/cicero_diasWS_sem_miniaturas.py`

3. Os novos scrapers exportam o mesmo padrão de colunas usado pelo Oráculo:
   - Título
   - Autor
   - Ano
   - Técnica
   - Dimensões
   - Preço
   - Descrição
   - Link da obra
   - Link da imagem da obra

4. A coleta usa um método simples com `requests` + `BeautifulSoup`, com fallback em Playwright quando a página depende de renderização JavaScript.

5. Os scrapers antigos foram mantidos:
   - ArtSoul
   - Blombo
   - Gagosian
   - Saatchi Art

## Validação realizada

- `app.py` compilado com sucesso.
- Todos os scrapers ativos compilados com sucesso.

## Observação

A navegação online direta não pôde ser executada neste ambiente por limitação de DNS do sandbox. A estrutura foi montada com base nos arquivos enviados e nas páginas informadas pelo usuário, usando seletores e fallbacks para manter a coleta simples e resiliente.
