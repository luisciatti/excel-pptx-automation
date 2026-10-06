# Manual de uso do gerador de relatórios PPTX

## 1. Objetivo

Este projeto consolida os dados mensais dos clientes e gera apresentações PPTX preenchidas com:

- Usuários concorrentes.
- Storage de banco de dados.
- Storage total dos bancos.
- Storage do S3.
- Incidentes e minutos de indisponibilidade.
- Top 5 de sessões encerradas, quando os logs estiverem disponíveis.

O processo deve ser executado uma vez por mês, depois que os dados do período estiverem atualizados.

## 2. Pré-requisitos

Antes de começar, confirme que você tem:

- Acesso AWS SSO da equipe de Operação em Cloud: [AWS access portal](https://cloudphilips-br.awsapps.com/start/#/?tab=accounts).
- Acesso à planilha [Cloud Customers - Storage Tracking.xlsx](https://share.philips.com/:x:/r/sites/CloudOperation/Shared%20Documents/General/6%20-%20Operation%20Process/Cloud%20Customers%20-%20Storage%20Tracking.xlsx?d=w4d39b24623d04ad8959364fee9d8c65b&csf=1&web=1&e=ZzzPIa).
- Acesso ao ambiente Tasy dos clientes para consultar usuários concorrentes.
- Python e as dependências instaladas, caso a interface não seja executada pelo `.exe`.

A documentação corporativa de referência é [Processo de storage mensal](https://bionexo.atlassian.net/wiki/x/A4DcdAE).

### Segurança

Nunca coloque senhas, tokens, credenciais AWS ou dados secretos nos CSVs, no manual ou no repositório. Use o AWS SSO e o gerenciador de credenciais aprovado pela empresa.

## 3. Fluxo mensal resumido

1. Atualize a planilha com os dados de Storage e usuários.
2. Exporte ou obtenha os arquivos de Cloud Incidents do período.
3. Reúna os logs de throttling dos clientes.
4. Abra a interface do gerador.
5. Atualize os arquivos de entrada pelos botões da interface.
6. Confira o preview dos dados.
7. Gere o PPTX.
8. Confira os logs da execução e abra os arquivos gerados.

## 4. Como obter os dados de Storage

### 4.1 Abrir a conta AWS

1. Acesse o [AWS access portal](https://cloudphilips-br.awsapps.com/start/#/?tab=accounts).
2. Selecione a conta do cliente que será atualizada.
3. Confirme que está na conta e no ambiente corretos antes de consultar qualquer recurso.

### 4.2 Storage do banco de produção

1. Abra o serviço **RDS**.
2. Localize o banco de produção do cliente.
3. Abra a configuração do banco.
4. Localize o campo **Storage**.
5. Registre o valor na linha do cliente e no mês correto da planilha.
6. Confira se o cliente, o ambiente e o mês estão corretos.

### 4.3 Storage total dos bancos

Consulte os bancos relacionados ao cliente e registre os valores individuais. O total deve considerar os bancos indicados no processo:

`Storage(dbhomol) + Storage(dbprod) + Storage(dbteste) + Storage(dbtieprod)`

Anote os valores individuais antes de somar. Isso facilita a conferência e a auditoria. O valor consolidado deve ser informado no campo de Database/total de bancos da planilha, conforme o layout vigente.

### 4.4 Storage do S3

1. Abra o serviço **Amazon S3** na conta correta.
2. Localize o bucket `anexos-tasy`.
3. Abra a área de métricas/monitoramento.
4. Consulte **Total bucket size**.
5. Selecione o último dia do mês de referência.
6. Registre o valor na coluna `S3 (GB)` da planilha.
7. Confirme a unidade do valor antes de salvar.

### 4.5 Usuários concorrentes no Tasy

1. Para localizar a URL do cliente, acesse **EC2 > Load Balancers**.
2. Localize o Application Load Balancer do cliente.
3. Abra as regras do listener.
4. Identifique a URL do ambiente do cliente.
5. Acesse o ambiente usando as credenciais aprovadas.
6. Abra a função **Gerenciador de Relatórios**.
7. Filtre pelo código `4488`.
8. Com o botão direito, selecione **Testar relatório**.
9. Informe a data inicial e final do mês de referência.
10. Gere o relatório.
11. Confira o maior valor de usuários concorrentes no canto superior direito.
12. Registre esse valor na coluna de usuários da planilha.

## 5. Como preparar o Cloud Incidents

O projeto usa o Cloud Incidents como fonte única dos incidentes. Não é necessário fornecer um arquivo separado de `Incidentes`.

O CSV deve conter, no mínimo, as colunas:

- `Customer`
- `Year`
- `Month`
- `Total Time (Minutes)`
- `Consider Unavailability?`, quando disponível

O campo `Consider Unavailability?` define se o tempo deve entrar no cálculo de indisponibilidade. Registros com `Yes` entram na soma de minutos; os demais continuam contando como incidentes, mas não somam minutos de indisponibilidade.

Na interface:

1. Clique em **Atualizar Cloud Incidents**.
2. Selecione o CSV do período. É possível selecionar mais de um arquivo.
3. O sistema substitui o arquivo canônico do mesmo ano em `data/inputs/`.
4. O relatório Excel será regenerado quando o arquivo de entrada for mais novo que o relatório existente.

## 6. Como preparar os logs de throttling

Os logs são usados para preencher o total e o Top 5 de sessões encerradas.

1. Reúna os CSVs dos clientes e meses necessários.
2. Mantenha os nomes no formato esperado, por exemplo:

`einstein-throttling-log-2026-06.csv`

3. Cada arquivo deve conter as colunas `QTD` e `FUNCAO`.
4. Na interface, clique em **Adicionar arquivos CSV**.
5. Selecione um ou vários logs.
6. O sistema copia os arquivos para `data/logs/`.

O nome do cliente e o ano/mês são identificados pelo nome do arquivo. Portanto, não renomeie o arquivo de forma que remova essa informação.

## 7. Como usar a interface

Execute na raiz do projeto:

```powershell
cd C:\Repositorios\Excel
py .\interface.py
```

Na tela:

1. Escolha **Um período**, **Cliente específico** ou **Todos os períodos disponíveis**.
2. Para um período ou cliente, escolha o ano e o mês.
3. Se escolher **Cliente específico**, selecione o cliente na lista carregada para o período.
4. Marque a versão México se também precisar do PPTX com sufixo `-MEX`.
5. Atualize o Storage.
6. Atualize o Cloud Incidents.
7. Adicione os logs de throttling.
8. Abra a aba **Preview dos dados**.
9. Confirme se os clientes têm Storage, incidentes, indisponibilidade e logs esperados.
10. Clique em **Gerar apresentações**.

No modo **Cliente específico**, o sistema gera o PPTX do período somente para o
cliente selecionado. No modo **Um período**, gera os clientes do período; no
modo **Todos os períodos disponíveis**, percorre todos os templates encontrados.

## 8. Como interpretar o preview

Na aba **Preview dos dados**:

- `Pronto para preencher`: existem os campos principais de Storage para o cliente no período.
- `Sem dados Excel`: o relatório não possui algum dado principal de Storage/usuários para o cliente.
- `Incidentes`: quantidade calculada a partir do Cloud Incidents.
- `Indisp. (min)`: soma dos tempos marcados para considerar indisponibilidade.
- `Logs QTD`: total obtido dos logs de throttling encontrados para o cliente e período.

O preview é uma conferência antes da geração. Ele não substitui a validação do PPTX final.

## 9. Como verificar o resultado

Após a execução:

- Na aba **Arquivos gerados**, confira os PPTX e abra um arquivo com duplo clique.
- Na aba **Logs da execução**, confira avisos de clientes sem dados e erros de processamento.
- Verifique a pasta `output/pptx/`.
- Os relatórios Excel ficam em `output/reports/`.
- Confira alguns clientes diretamente no PPTX, especialmente os que aparecerem sem dados no preview.

## 10. Problemas comuns

### Nenhum template encontrado

Confirme se os templates estão em `data/pptx/` e seguem o formato de mês, como `202606.pptx` ou `202606 - MEX.pptx`.

### Storage não encontrado

O nome ativo precisa seguir o padrão:

`Cloud - Storage 2026.csv`

Use o botão **Atualizar Storage** para substituir o arquivo sem precisar renomeá-lo manualmente.

### Cloud Incidents não encontrado

Confirme se o nome contém `Cloud Incidents(Registro` e o ano, por exemplo:

`Cloud Incidents(Registro 2026).csv`

### Cliente aparece como sem dados

Confira o nome do cliente no CSV, o mês/ano do registro e os campos de usuários, S3 e Database no Storage. Use a aba **Preview dos dados** antes de gerar novamente.

### Logs não aparecem

Confira se o arquivo contém `QTD` e `FUNCAO` e se o nome segue o padrão `cliente-throttling-log-AAAA-MM.csv`.

### Relatório não atualiza

Verifique se o CSV atualizado foi copiado para `data/inputs/` e se a data de modificação dele é mais recente que o arquivo em `output/reports/`.

## 11. Estrutura das pastas

```text
data/inputs/   Storage e Cloud Incidents
data/logs/     Logs de throttling
data/pptx/     Templates PPTX
output/reports/Relatórios Excel consolidados
output/pptx/   PPTX finais e apresentações por cliente
```
