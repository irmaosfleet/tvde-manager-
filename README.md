# Plataforma Irmãos Fleet 2.0 — Base de Produção

Esta atualização substitui a versão 1.4 e utiliza o banco oficial enviado pelo utilizador.

## Incluído nesta versão

- Migração do `BANCO_DE_DADOS.xlsx` para SQLite.
- 10.381 cadastros oficiais importados.
- 24 parceiros identificados e vinculados aos motoristas.
- Campos migrados: ID, nome, cidade, companhia, IBAN, banco, comissões, parceiro, cartão PRIO, percentagem, aluguel, descontos, reembolsos, imediata e observação.
- Pesquisa no cadastro por nome, ID, IBAN, cartão, cidade e companhia.
- Tela **Sem IBAN** com duas opções:
  - cadastro definitivo no banco;
  - IBAN temporário somente para uma semana e grupo.
- IBAN compartilhado permitido, com alerta e composição individual no relatório.
- Agrupamento de pagamentos por IBAN com uma taxa bancária por transferência.
- Auditoria de pagamentos por IBAN.
- Histórico de alterações do sistema.
- Cadastro de parceiros e criação de login.
- Perfil Administrador com acesso total.
- Perfil Parceiro limitado aos próprios motoristas.
- Importação Uber, Bolt e PRIO mantida.
- XML `pain.001.001.03`, divisão por partes e limite de €50.000 mantidos.

## Primeiro acesso

- Utilizador: `admin`
- Senha: `admin123`

Altere as variáveis `ADMIN_USER`, `ADMIN_PASSWORD` e `SECRET_KEY` no Render antes do uso real.

## Atualização no GitHub

Substitua todos os arquivos do repositório pelos arquivos deste pacote e aguarde o novo deploy do Render.

## Segurança obrigatória

Este pacote contém dados reais, incluindo IBANs. O repositório do GitHub deve ser **PRIVADO**. Nunca publique este ZIP ou o banco em um repositório público.

Para uso contínuo no Render, configure um Persistent Disk e defina `DB_PATH` apontando para esse disco, para que novos dados não sejam perdidos em um redeploy.

## Observação

A base de permissões de parceiros está pronta. Os acessos de cada parceiro são criados na tela **Parceiros**, informando utilizador e senha inicial. O portal detalhado do motorista será a etapa seguinte.

## Correção 2.0.1
- Corrigido erro 500 na tela de preparação/geração dos XMLs.
- Corrigido conflito do Jinja com a chave `items` dos lotes de pagamentos.
- Pacote organizado com os arquivos na raiz para facilitar o upload ao GitHub.

## Atualização 3.1.0 — importação consolidada

- Consolida várias linhas do mesmo motorista em um único relatório por plataforma, semana e grupo.
- Evita totais inflados em arquivos detalhados por viagem.
- Permite importar vários arquivos da mesma plataforma na mesma semana, somando as partes sem criar linhas duplicadas.
- Melhora a leitura de CSVs com linhas introdutórias, diferentes delimitadores e codificações.
- Amplia o reconhecimento de cabeçalhos da Uber Eats, Bolt Food e arquivos em inglês.
- Registra uma assinatura curta do arquivo no histórico de importações para facilitar auditoria.
