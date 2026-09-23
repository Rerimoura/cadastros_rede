# Cadastro de Produtos - Redes 📋

Web app em [Streamlit](https://streamlit.io/) para preenchimento automático de fichas de cadastro de produtos exigidas por redes de varejo. O usuário cola os códigos de mercadoria, e o app devolve a planilha-modelo da rede já preenchida com os dados vindos do banco.

## Como funciona

1. Escolhe a rede (hoje: SuperMaxi).
2. Cola os códigos de mercadoria (um por linha ou separados por vírgula).
3. O app busca os dados em `mercadorias` (dados fiscais/logísticos) e `categorias` (marca) no Postgres, e preenche o template Excel da rede a partir da linha 9, uma linha por produto.
4. Baixa a planilha pronta para envio.

Códigos não encontrados na base são avisados na tela e ficam de fora do arquivo gerado.

## Adicionar uma nova rede

1. Colocar o arquivo `MODELO ...xlsx` da rede em `templates/`.
2. Adicionar uma entrada em `REDE_CONFIGS` (em `CadastroRedes.py`) com:
   - `template_path`, `sheet_name`, `linha_inicial` (primeira linha de dados);
   - `colunas`: um dicionário `{letra_da_coluna: função(produto) -> valor}`.

O restante do fluxo (input, busca no banco, geração do Excel, download) é reaproveitado automaticamente.

## Configuração do banco de dados

Usa `st.secrets["postgres"]`, o mesmo banco Neon dos outros apps do grupo (Neon, Venn). Localmente, crie `.streamlit/secrets.toml`:

```toml
[postgres]
host = "..."
database = "..."
user = "..."
password = "..."
port = "5432"
```

Em produção (Streamlit Community Cloud), configure os mesmos valores em **Settings → Secrets** do app — nunca commitar esse arquivo (já está no `.gitignore`).

## Rodar localmente

```bash
pip install -r requirements.txt
streamlit run CadastroRedes.py
```
