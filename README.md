# Cálculo de Rescisão de Acessórios (com wizard de mapeamento)

App Streamlit para:
- Carregar **CSV** ou **XLSX** (você escolhe a planilha).
- Fazer **mapeamento de colunas** (campo lógico → coluna do arquivo).
- Salvar/Carregar **perfil de mapeamento** (JSON).
- Calcular valores **com** e **sem** devolução.
- Filtrar (cliente, classe, termo, serviço, status, faixas de R$).
- Exportar **CSV** e **XLSX**.

## Deploy local
```bash
python -m venv .venv
. .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run streamlit_app.py
