IRMÃOS PLATAFORMA 1.2.1

Correções desta versão:
- Dashboard usa exatamente os totais do relatório geral do fechamento, inclusive dinheiro em mãos.
- Desconto duplicado em cadastros Uber/Bolt é consolidado por motorista e cobrado uma única vez; se houver valores diferentes, usa o maior.
- Reembolso Uber TVDE lê exatamente as colunas Reembolso_1 a Reembolso_4 configuradas em Arquivos_Pgto.
- Sem IBAN mostra apenas pagamentos positivos pendentes do último fechamento.
- Ao salvar um IBAN, todos os cadastros duplicados do mesmo motorista são atualizados e o fechamento é recalculado.
- Correção do envio do campo Banco no formulário Sem IBAN.

Após o deploy, apague as importações de teste, importe novamente os relatórios e gere um novo fechamento para reler os reembolsos e aplicar a consolidação dos descontos.

## Versão 1.2.2
- Recibo individual de pagamento em PDF diretamente no fechamento.
- Correção da pesquisa da tela Sem IBAN.
- Atualização de IBAN baseada no identificador externo da plataforma, evitando alterar homónimos.
- Dashboard continua a ler os totais diretamente do fechamento, sem recalcular bruto, dinheiro, combustível, descontos ou reembolsos.
- Versão da aplicação atualizada para 1.2.2.

## Versão 1.2.3
- Reembolso Uber TVDE: lê e soma exclusivamente as quatro colunas configuradas em `Arquivos_Pgto`, tolerando apenas diferenças técnicas de acentos, espaços, pontuação, quebras de linha e sufixos `.1` adicionados pelo pandas.
- Sem IBAN: reconhece também valores vazios representados por `-`, `NAN`, `NONE`, `NULL` ou `SEM IBAN` e lista cada cadastro pelo `driver_id`, sem agrupar pessoas diferentes apenas pelo nome.
