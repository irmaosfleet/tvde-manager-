# Irmãos Fleet - Painel de Gestão 1.1.0

Atualização financeira para GitHub e Render.

## Correções desta versão

- Dinheiro em mãos aplicado somente à Uber Eats.
- Dinheiro em mãos entra na base da comissão e é abatido integralmente no pagamento final.
- Saldos iguais ou inferiores a zero continuam fora do XML e aparecem no relatório de negativos.
- Pagamentos separados em dois grupos:
  - TVDE: Uber TVDE + Bolt TVDE.
  - Delivery: Uber Eats + Bolt Food.
- Consolidação por IBAN feita dentro de cada grupo.
- Taxa de € 1,25 aplicada uma vez por IBAN em cada grupo de pagamento.
- XMLs divididos em lotes de até € 50.000.
- Download em ZIP com nomes `PAGAMENTO_TVDE_01.xml` e `PAGAMENTO_DELIVERY_01.xml`.

## Publicação

Substitua os ficheiros do repositório pelos desta pasta e no Render use:

`Manual Deploy` → `Clear build cache & deploy`

Depois limpe a semana importada, importe novamente os relatórios e gere um novo fechamento.

## Versão 1.0.10 — Sem IBAN
- Nova opção **Sem IBAN** no menu.
- Lista de motoristas sem IBAN com pesquisa.
- Edição rápida de IBAN, banco, parceiro e comissão do parceiro.
- Botão **Salvar e recalcular** atualiza o cadastro e os fechamentos já processados.
- Recalcula somente a consolidação por IBAN e a taxa bancária, sem mudar a lógica financeira.
- Exportação Excel geral e abas separadas por parceiro.

## Versão 1.1.2
- Corrige duplicação de combustível, descontos e imediata quando o mesmo motorista possui TVDE e Delivery.
- Comissão da empresa e dos parceiros passa a ser somada uma única vez por motorista ativo.
- Totais da Dashboard e Área do Gestor revisados.


## Versão 1.1.4
- Corrige os indicadores da Dashboard e Área do Gestor.
- Comissão da empresa e do parceiro agora é a comissão calculada multiplicada pela participação de cada um no banco de dados.
- Totais são deduplicados por motorista e categoria para não dobrar combustível, descontos e dinheiro em mãos.
