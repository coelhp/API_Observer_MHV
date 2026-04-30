# SKF Observer API – Assets, Points e NextGenSensor

Este repositório contém scripts Python para consumo direto da **API SKF Observer**, com foco em integração com **Power BI (Script Python)**, sem necessidade de armazenamento intermediário de dados.

Os scripts realizam autenticação, consulta e estruturação dos dados retornados pelos endpoints:
- `/v2/assets`
- `/v1/machines/{assetId}/points`
- `/v1/nextgensensor`

---

## 📌 Visão Geral

A API SKF Observer disponibiliza dados de monitoramento de ativos industriais.  
Cada planta é identificada por uma **porta TCP específica**, enquanto a URL base permanece a mesma.

O fluxo de dados segue a lógica:

1. Autenticação (`/token`)
2. Consulta de ativos (`/v2/assets`)
3. Consulta de pontos por ativo (`/v1/machines/{assetId}/points`)
4. Consulta de sensores de nova geração (`/v1/nextgensensor`)
5. Consolidação dos dados em DataFrames para uso no Power BI

---

## 🔐 Autenticação

Todas as requisições exigem um **Bearer Token**, obtido via:

```text
POST /token
```

## Parâmetros
- `grant_type`: `password`
- `username`: usuário SKF
- `password`: senha SKF

O token possui validade aproximada de **20 minutos**.


---

## 🌐 Configuração de Conexão

### URL Base

### Portas por Planta

| Planta | Porta |
|------|------|
| Unidade 1 | 00000 |
| Unidade 2 | 11111 |
| Unidade 3 | 22222 |
| Unidade 4 | 33333 |
| Unidade 5 | 44444 |

> A diferenciação entre plantas ocorre exclusivamente pela porta.

---

## 📍 Endpoint /v2/assets

### Descrição
Retorna a lista de ativos monitorados na planta selecionada.

### Requisição

```text
GET /v2/assets
```

### Uso
- Endpoint base do fluxo
- O campo `ID` identifica unicamente cada asset
- Este `ID` é utilizado para consultar os pontos de medição

### Exemplo de Retorno
```json
[
  {
    "ID": 123,
    "Name": "Motor Principal",
    "Area": "Moagem",
    "Status": "Online"
  }
]
```
---

## 📍 Endpoint /v1/machines/{assetId}/points
Descrição

Retorna os pontos de medição associados a um asset específico.

Requisição

```text
GET /v1/machines/{assetId}/points
```

Fluxo de Uso

Consultar /v2/assets

Extrair todos os ID dos assets

Para cada ID, consultar os respectivos points

Consolidar todos os pontos em um único DataFrame

Exemplo de Retorno
```json
[
  {
    "ID": 1092,
    "Name": "Velocidade RMS",
    "Unit": "mm/s",
    "MeasurementType": "Vibration"
  }
]
```
Observação

Nos scripts, cada ponto recebe a coluna AssetID, permitindo relacionamento direto com a tabela de assets no Power BI.

## 📍 Endpoint /v1/nextgensensor
Descrição

Retorna informações sobre sensores de nova geração instalados na planta.

Requisição

```text
GET /v1/nextgensensor
```

```json
[
  {
    "ID": 45,
    "SerialNumber": "NGS-2024-001",
    "Status": "Active"
  }
]
```

## ⚠️ Boas Práticas

Definir timeout adequado (ex.: 60 segundos)

Evitar múltiplas chamadas simultâneas

Validar retorno do token antes das requisições

Manter scripts simples para melhor estabilidade no Power BI

## 🚀 Próximos Passos

Inclusão do endpoint /v1/points/{id}/trendMeasurements

Modelagem estrela no Power BI

Evolução para dashboard online (Django ou Streamlit)


# Machine Viewer API – Assets & Workorders (Python)

Scripts em Python para consumo da SKF Machine Viewer API v3, permitindo a extração estruturada de:

Assets (ativos monitorados)

Workorders (ordens de serviço)

Os dados são retornados como pandas DataFrame, com paginação automática e prontos para uso em ETL, Power BI, dashboards ou data lakes.

## Pré-requisitos

Python

API Key válida da Machine Viewer API v3

Permissão da API Key para o shortname utilizado

Dependências:

```python
pip install requests pandas
```

## 🔐 Autenticação

A autenticação é feita via API Key, enviada no header de todas as requisições:

```python
x-api-key: SUA_API_KEY
Content-Type: application/json
```

##⚠️ Importante
A API Key deve estar explicitamente autorizada para o shortname (ex.: BRAEO00000).
Caso contrário, a API retornará HTTP 403 – Forbidden.

## Base URL

```python
https://analystapi.repcenter.skf.com
```

Todos os endpoints utilizam o shortname como parâmetro de path.

## 📍 Assets

### Endpoint

```text
POST /{shortname}/assets
```

## Descrição

Retorna a lista de ativos da hierarquia, incluindo identificação, localização funcional, criticidade, status operacional e tipo de equipamento.

Campos retornados

- `assetId`

- `assetName`

- `assetDescription`

- `functionalLocation`

- `parentId`

- `parentName`

- `criticality`

- `assetStatus`

- `assetSegment`

- `conditionIndex`

- `equipmentType`

## Observações

A API retorna até 1000 registros por requisição

O script implementa paginação automática utilizando o campo `nextCursor`

## 📍 Workorders
Endpoint

```text
POST /{shortname}/workorders
```
Descrição

Retorna ordens de serviço registradas em um período definido, contendo informações de execução, prioridade, vínculo com ativos e intervenções associadas.

Filtros obrigatórios

- `openingDateStart`

- `openingDateEnd`

Campos retornados

- `assetId`

- `id (SAM da OS)`

- `orderNumber`

- `openingDate`

- `scheduledDate`

- `deadline`

- `priority`

- `technique`

- `situation`

- `services`

- `cmms`

- `cmmsRegister`

- `reWork`

- `author`

- `intervention (campo aninhado)`

## Exemplo de Uso

```python
# Assets
df_assets = get_all_assets()
print(df_assets.head())

# Workorders
df_workorders = get_workorders(
    opening_date_start="2024-01-01 00:00:00",
    opening_date_end="2024-12-31 23:59:59"
)
print(df_workorders.head())
```
## Tratamento de Erros
### HTTP 403 – Forbidden

API Key sem permissão para o shortname

Endpoint não liberado para a chave

### HTTP 401 – Unauthorized

API Key inválida ou expirada

### HTTP 400 / 500

Erro de payload ou falha interna da API

## ⚠️ Boas Práticas

Utilizar variáveis de ambiente para armazenar API Key e shortnames

Persistir dados em:

CSV / Parquet (Data Lake)
SQLite / PostgreSQL (dashboards e BI)

Normalizar campos aninhados antes de uso analítico (`intervention`)

Relacionar tabelas pelo campo `assetId`

## 🚀 Extensões Futuras

- Suporte multi-planta (loop por shortname)

- Integração com endpoints:

    - `/lastmeasurement`

    - `/measurements`

    - `/conditions`

- Pipeline ETL automatizado

- Modelagem dimensional pronta para Power BI

## Contexto de Uso

Este projeto é indicado para:

Monitoramento preditivo

Consolidação de ativos industriais

Análise de ordens de manutenção

Dashboards operacionais e gerenciais


## 📊 Integração com Power BI

Os scripts são compatíveis com Obter Dados → Script Python

Cada script retorna um único DataFrame

Não há gravação de arquivos

Atualização ocorre sob demanda no Power BI Desktop. Podendo ser otimizada de forma que atualize com mais recorrências

Recomenda-se:

Um script por endpoint principal

Relacionamento via AssetID no modelo de dados


## 📄 Licença

Uso interno / corporativo – SKF Observer API.

## Acervo de Imagens Resultado Final

<img width="1304" height="638" alt="image" src="https://github.com/user-attachments/assets/810fcf4a-59c5-404c-a7a4-2c2bbc13cea8" />
<img width="1139" height="723" alt="image" src="https://github.com/user-attachments/assets/ab6dd7e5-040b-4184-96d9-baf72255d09a" />
<img width="1295" height="687" alt="image" src="https://github.com/user-attachments/assets/1206a23d-3eef-462a-81a8-01f2e6994b22" />

**OBS:** Imagens foram alteradas afim de respeitar e seguir com a proteção de dados da empresa.

