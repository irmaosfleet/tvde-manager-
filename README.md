# Irmãos Fleet 1.0

## Login inicial
- Utilizador: `admin`
- Senha: `admin123`

Altere no Render em Environment:
- `ADMIN_USER`
- `ADMIN_PASSWORD`
- `SECRET_KEY`

## Funcionalidades
- Login protegido
- Interface responsiva
- Dashboard
- Importação automática Uber e Bolt
- Dinheiro em mãos
- Motoristas
- Viaturas
- Alertas
- Relatórios
- Recibos PDF

## Render
Build: `pip install -r requirements.txt`
Start: `gunicorn app:app`

Observação: nesta versão o SQLite é adequado para demonstração. Para uso real e vários utilizadores, migrar para PostgreSQL.
