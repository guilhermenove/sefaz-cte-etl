# SEFAZ CT-e ETL: frete lançado sozinho, nota a nota

Pipeline em Python que baixa da **SEFAZ** todos os CT-e (Conhecimento de Transporte Eletrônico) destinados a um CNPJ, organiza os XMLs, gera o DACTE em PDF e lança o valor do frete na planilha de controle, casando cada CT-e com a NF-e que ele transporta.

## O problema

O controle de frete por CT-e era manual: entrar no portal da SEFAZ, baixar XML por XML, achar a NF correspondente, digitar o valor e lembrar quais tinham sido cancelados. Lento, repetitivo e fácil de errar.

## Como funciona

```
Certificado A1 (.pfx)
        │  TLS mútuo
        ▼
SEFAZ CTeDistribuicaoDFe (SOAP)  ──►  XMLs por NSU incremental  ──►  parser + regras fiscais  ──►  Google Sheets
   loop até ultNSU = maxNSU          (só o que é novo)              (cancelados, dedup, soma por NF)    coluna "Frete"
```

1. **Extract:** carrega o certificado digital e consulta o webservice `CTeDistribuicaoDFe` a partir do último NSU salvo, em loop até esvaziar a fila.
2. **Transform:** descompacta cada documento (gzip + base64), separa eventos de CT-e, registra cancelamentos, extrai chave, emissão, transportadora, valor e NFs vinculadas, e filtra por data de emissão.
3. **Load:** escreve o frete na linha de cada NF (somando quando uma NF tem mais de um CT-e) e grava a chave do CT-e vinculado. Também exporta um CSV e salva XML + DACTE organizados por mês.

### Proteções

- **Incremental por NSU:** cada execução busca só o que é novo, respeitando o limite de consumo da SEFAZ.
- **Idempotente:** CT-e já salvos não duplicam, mesmo se o NSU for resetado.
- **Edição manual preservada:** uma NF que já recebeu frete da rotina nunca é sobrescrita de novo.
- **Cancelados ignorados:** eventos de cancelamento (110111) tiram o CT-e do cálculo.
- **Relógio confiável:** o filtro de "mês atual" usa a data devolvida pela SEFAZ, não a do computador.
- **Chave como texto:** evita que o Sheets transforme a chave de 44 dígitos em notação científica.

## Como usar

```bash
pip install -r requirements.txt
cp .env.example .env        # preencha CNPJ, planilha e caminhos
python rotina_cte.py        # interface gráfica (Tkinter) com log em tempo real
python rotina_cte.py --cli  # terminal / agendado
```

Para rodar agendado (Agendador de Tarefas do Windows, cron) sem digitar senha, defina `CERT_SENHA` no ambiente. Se ela não existir, o script pergunta na hora.

### Arquivos necessários (nunca versionados)

| Arquivo | O que é |
|---|---|
| `.env` | Configuração (modelo em `.env.example`) |
| `certificado.pfx` | Certificado digital A1 do CNPJ consultado |
| `credenciais_google.json` | Conta de serviço do Google com acesso de edição à planilha |

O `.gitignore` bloqueia todos eles.

### Configuração

| Variável | Descrição |
|---|---|
| `CNPJ_CONSULTADO` | CNPJ do tomador, só números |
| `CUF_AUTOR` | Código IBGE da UF (41 = PR, 35 = SP) |
| `CERT_ARQUIVO` / `CERT_SENHA` | Certificado A1 e senha (senha opcional) |
| `GOOGLE_CREDENCIAIS` / `PLANILHA_URL_OU_ID` / `ABA_SAIDAS` | Destino no Google Sheets |

No topo do script ficam os ajustes que não são segredo: ambiente (`producao` ou `homologacao`), filtro por mês e posição das colunas na planilha.

## Arquivos de estado (gerados pela rotina)

| Arquivo | Função |
|---|---|
| `ultimo_nsu.json` | Último NSU consumido; a próxima execução continua dali |
| `ctes_cancelados.json` | Chaves canceladas, para ignorar |
| `ctes_salvos.json` | Chaves já baixadas, para não duplicar |
| `nfs_frete_atualizado.json` | NFs já lançadas, para proteger edição manual |
| `cte_recebidos/` | XML + DACTE, organizados por mês |
| `cte_extraido.csv` | Resumo da última execução |

## Segurança

- Certificado, senha, credenciais e XMLs (que têm dados fiscais de parceiros) ficam fora do Git.
- A senha só existe em memória ou na variável de ambiente.
- Os PEM temporários extraídos do `.pfx` são usados apenas durante a conexão.

## Stack

Python · SOAP · certificado digital A1 · cryptography · lxml · gspread · brazilfiscalreport (DACTE) · Tkinter
