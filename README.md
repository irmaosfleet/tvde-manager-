IRMÃOS PLATAFORMA 1.2.1

Correções desta versão:
- Dashboard usa exatamente os totais do relatório geral do fechamento, inclusive dinheiro em mãos.
- Desconto duplicado em cadastros Uber/Bolt é consolidado por motorista e cobrado uma única vez; se houver valores diferentes, usa o maior.
- Reembolso Uber TVDE lê exatamente as colunas Reembolso_1 a Reembolso_4 configuradas em Arquivos_Pgto.
- Sem IBAN mostra apenas pagamentos positivos pendentes do último fechamento.
- Ao salvar um IBAN, todos os cadastros duplicados do mesmo motorista são atualizados e o fechamento é recalculado.
- Correção do envio do campo Banco no formulário Sem IBAN.

Após o deploy, apague as importações de teste, importe novamente os relatórios e gere um novo fechamento para reler os reembolsos e aplicar a consolidação dos descontos.
