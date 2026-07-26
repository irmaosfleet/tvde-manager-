# Irmãos Fleet 1.1.7 — Estabilização

Correções desta versão:

- Reembolso Uber TVDE lido somente das quatro colunas configuradas em `Arquivos_Pgto`.
- Removida a identificação automática por palavras como “Pago a si” e “Saldo”, que podia tratar o valor bruto como reembolso.
- Área do Gestor corrigida: removida a consulta à coluna inexistente `category`.
- Dashboard e Área do Gestor passam a somar diretamente os valores gravados no fechamento.
- Tema noturno aplicado ao painel.
- Logo da Irmãos Fleet no menu, login e ícone da página.
- Combustível discriminado por ficheiro PRIO e total do fechamento preservado.

## Teste após atualizar

1. No GitHub, substitua os ficheiros pela versão deste ZIP.
2. Aguarde o deploy do Render.
3. Em **Importações**, limpe a semana anterior.
4. Importe novamente os relatórios e os ficheiros PRIO.
5. Processe um novo fechamento.
6. Confira primeiro a coluna `Reembolso` no Excel do novo fechamento.
7. Depois confira Dashboard e Área do Gestor.

Um fechamento já criado com o reembolso errado não é corrigido automaticamente.
