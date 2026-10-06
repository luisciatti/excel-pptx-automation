# Excel PPTX Automation

## Estrutura organizada

- `data/inputs/` - arquivos CSV de entrada (storage, incidentes, cloud incidents)
- `data/pptx/` - templates PPTX (somente os arquivos base usados para gerar slides)
- `data/logs/` - logs de throttling em CSV
- `output/pptx/` - apresentações geradas
- `output/reports/` - relatórios Excel gerados
- `backup/` - scripts antigos ou de referência não usados no fluxo principal

## Uso rápido

Para aprender o processo completo, abra o [MANUAL.md](MANUAL.md) ou clique em
`Abrir manual de instruções` na interface.

1. Ative o ambiente virtual:

```powershell
cd C:\Repositorios\Excel
.\.venv\Scripts\Activate.ps1
```

2. Instale dependências (se ainda não instalou):

```powershell
pip install -r dev-requirements.txt
```

3. Execute o script principal:

```powershell
py .\preencher.py
```

4. Ou abra a interface gráfica:

```powershell
py .\interface.py
```

Na interface, escolha um período ou todos os templates disponíveis. Use os botões
de atualização para substituir os CSVs de `Storage` e `Cloud Incidents` em
`data/inputs/`; os logs de throttling são adicionados em
`data/logs/`. Depois clique em `Gerar apresentações`. Os relatórios Excel são
atualizados quando os arquivos de entrada forem mais novos, e os PPTX gerados
aparecem na lista de resultados e podem ser abertos com duplo clique.

É possível gerar todos os clientes de um período ou selecionar `Cliente específico`
para gerar somente a apresentação de um cliente.

A área de resultados possui três abas: `Arquivos gerados`, `Logs da execução` e
`Preview dos dados`. O preview mostra, por cliente, Storage, incidentes,
indisponibilidade, logs e se existem dados Excel suficientes para preencher o
PPTX do período selecionado.

5. Ou gere o `.exe`:

```powershell
.\build_exe.ps1
```

6. Após gerar o `.exe`, use:

```powershell
.\dist\Preencher\Preencher.exe 202602
```

## Observações

- Os dados de entrada devem ficar em `data/inputs/`
- O storage ativo usa o formato `Cloud - Storage {ano}.csv`; os formatos `Cloud Storage` e `Storage` não são usados.
- O relatório Excel e o detalhamento dos incidentes no PPTX usam `Cloud Incidents(Registro {ano}).csv`.
- Os templates PPTX devem ficar em `data/pptx/`
- O script grava saídas em `output/pptx/`
- O `.exe` é otimizado para usar a pasta `data` externamente e não embutir todos os dados no bundle.
