# FleetFlow 1.0

Primeira versão funcional do novo sistema de fechamento semanal para frotas.

## Funcionalidades
- Login por utilizador e senha.
- Dashboard semanal com indicadores e gráficos.
- Banco de motoristas pesquisável e editável.
- Importação inicial do `BANCO_DE_DADOS.xlsx`.
- Upload em lote de relatórios Uber, Uber Eats, Bolt, Bolt Food e ficheiros PRIO.
- Leitura dirigida pela aba `Arquivos_Pgto`.
- Processamento por motorista, origem e IBAN.
- Regra de dinheiro em mãos para Uber Eats.
- Reembolsos apenas para TVDE.
- Taxa de € 1,25 uma vez por IBAN quando o banco é AZUL.
- Exportação do resultado geral em Excel.
- Geração XML SEPA `pain.001.001.03`.

## Login inicial
Defina no Render as variáveis:
- `ADMIN_USER`
- `ADMIN_PASSWORD`
- `SECRET_KEY`

Em ambiente local, os padrões são `admin` / `admin123`.

## Executar localmente
```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\\Scripts\\activate
pip install -r requirements.txt
python app.py
```
Abra `http://localhost:5000`.

## Publicar no GitHub e Render
1. Envie todos os ficheiros desta pasta para um repositório GitHub.
2. No Render, crie um Web Service ligado ao repositório.
3. Use o `render.yaml` ou configure manualmente os comandos indicados.
4. Defina uma senha forte em `ADMIN_PASSWORD`.

## Observação importante
Esta é a primeira base testável. Antes de usar para pagamentos reais, compare os resultados e o XML com o programa atual e valide no portal do banco.
