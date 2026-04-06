# Buscador de Preços — Mercado Livre

Busca automaticamente para cada produto:
- **Menor preço** geral
- **Menor preço Full** (entrega rápida)
- **Mais vendido** (com quantidade de vendas)

Lê produtos e palavras negativas de uma planilha Google Sheets pública.

---

## Como gerar o .exe (Windows)

### Passo 1 — Suba este projeto no GitHub

1. Crie uma conta em [github.com](https://github.com) (grátis)
2. Crie um novo repositório (ex: `buscador-ml`)
3. Faça upload de todos os arquivos desta pasta

### Passo 2 — Execute o GitHub Actions

1. No seu repositório, clique na aba **Actions**
2. Clique em **"Build Windows EXE"**
3. Clique em **"Run workflow"** → **"Run workflow"**
4. Aguarde ~5 minutos

### Passo 3 — Baixe o .exe

1. Clique no run que foi executado
2. Em **Artifacts**, clique em **BuscadorML-Windows**
3. Baixe e extraia o `.zip`
4. Dentro está o `BuscadorML.exe`

---

## Como usar o .exe no Windows

1. Abra o `BuscadorML.exe`
2. Uma janela do Chrome abrirá automaticamente
3. O script busca os produtos da planilha Google Sheets
4. Ao terminar, salva `precos_mercadolivre.xlsx` na mesma pasta

---

## Planilha Google Sheets

Formato esperado (duas colunas):

| produto | palavras_negativas |
|---|---|
| Chlorella 500mg 60 capsulas | kit, combo, 2 potes |
| Omega 3 60 capsulas | infantil, kids |
| Creatina 300g | |

Deixe `palavras_negativas` vazio se não quiser filtrar nada.

---

## Alterar o link da planilha

Edite o arquivo `buscar_precos_ml.py` e troque a variável:

```python
GOOGLE_SHEET_URL = "https://docs.google.com/..."
```
