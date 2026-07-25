# FleetFlow 1.0.3 — Correção do fechamento

Esta versão corrige o erro 500 ao clicar em **Processar semana**.

Principais mudanças:
- cálculo protegido contra linhas incompletas;
- conversão segura de valores e percentuais;
- taxa de 1,25 aplicada uma vez por IBAN AZUL;
- mensagem do erro exibida na própria página;
- fechamento marcado como ERRO no histórico quando houver falha;
- não depende de pastas templates/static.

## Render
Build: `pip install -r requirements.txt`
Start: `gunicorn app:app --workers 1 --threads 4 --timeout 180`

Login inicial: `admin` / `admin123` (altere no Render).

## Versão 1.0.6
- Corrige `UNIQUE constraint failed: imports.sha256`.
- A assinatura de duplicidade passa a usar nome do arquivo + conteúdo.
- Arquivos com conteúdos iguais e nomes diferentes são aceitos.
- Reimportação do mesmo nome com o mesmo conteúdo é ignorada sem erro 500.
