# Agente de Montagem de Cargas

App Streamlit que substitui a macro VBA `MontarContainers`, evitando os
problemas de execução inconsistente do VBA/Excel.

## Como rodar

1. Instale as dependências:
   ```
   pip install -r requirements.txt
   ```

2. Rode o app:
   ```
   streamlit run app.py
   ```

3. O navegador abre automaticamente em `http://localhost:8501`.

## Como usar

1. Envie o plano de embarque (.xlsx, .xlsb, .xlsm ou .csv) — o app
   procura automaticamente a aba "PLANO" (ou similar).
2. Clique em "Montar containers".
3. Veja o resumo na tela (peso, pallets, SKUs e lotes por container).
4. Baixe a planilha final pelo botão de download.

## Regras de negócio (ajustáveis em `distribuicao.py`)

- `PESO_MAX = 27800` — peso máximo por container (kg).
- `MAX_POSICOES_PISO = 21` — posições de piso por container
  (BASE não-frágil + frágil solto sem nada por baixo).
- `SKUS_FRAGEIS` — lista de SKUs que não podem receber peso por cima
  (produtos frágeis: patês, pouches, latas finas, etc.)
- Cada pallet de BASE sustenta no máximo 1 pallet de TOPO frágil
  empilhado (1-para-1). O excedente de frágil vai para o PISO
  (posição própria, sem empilhar e sem sustentar nada).

## Arquivos

- `app.py` — interface Streamlit (upload, botões, tabelas, download).
- `leitura.py` — leitura de .xlsx/.xlsb/.xlsm/.csv e localização
  automática da tabela de itens.
- `distribuicao.py` — lógica de montagem dos containers (núcleo).
- `geracao_excel.py` — geração do arquivo Excel de saída.
