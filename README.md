# TVDE Manager — versão de teste online

Arquivos prontos para publicação no Render.

## Publicação
- Build command: `pip install -r requirements.txt`
- Start command: `gunicorn app:app`

## Aviso desta primeira demonstração
A versão inicial usa SQLite. No Render, os dados de teste podem ser apagados quando o serviço reiniciar.
Antes do uso real, o banco será migrado para PostgreSQL.

## Funcionalidades atuais
- Dashboard
- Motoristas
- Viaturas
- Alertas de documentos
- Relatórios Uber e Bolt
- Dinheiro em mãos
- Cálculo líquido
- Recibos PDF
