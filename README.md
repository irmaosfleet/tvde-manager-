# FleetFlow 1.0.7

Correções:
- Dinheiro em mãos da Uber Eats tratado pelo valor absoluto, pois o CSV costuma trazer esse campo negativo.
- O dinheiro entra na base da comissão e é abatido integralmente no pagamento final.
- Operações TVDE não recebem regra de dinheiro em mãos.
- Menu Parceiros visível, com Excel separado por parceiro.
- Relatório de parceiro inclui motoristas, origens, valores e comissão do parceiro.

Após atualizar no GitHub, use no Render: **Manual Deploy → Clear build cache & deploy**.
Depois limpe a semana importada e importe os relatórios novamente, para os valores de dinheiro serem normalizados.
