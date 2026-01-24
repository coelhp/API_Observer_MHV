import requests
import pandas as pd

# ===============================
# Configurações da API
# ===============================
BASE_URL = "https://analystapi.repcenter.skf.com"
SHORTNAME = "BRAEOXXX"  # Sendo 5 Shortnames respectivos a cada unidade da indústria
API_KEY = ""
TIMEOUT = 60

url = f"{BASE_URL}/{SHORTNAME}/assets"

headers = {
    "x-api-key": API_KEY,
    "Content-Type": "application/json"
}

# ===============================
# Função para coletar todos os assets
# ===============================
def get_all_assets():
    all_assets = []
    cursor = None

    while True:
        if cursor:
            payload = {
                "expression": (
                    "{filter(cursor: " + str(cursor) + "){"
                    "assetId assetName assetDescription functionalLocation "
                    "parentId parentName criticality assetStatus assetSegment "
                    "conditionIndex equipmentType}}"
                )
            }
        else:
            payload = {
                "expression": (
                    "{filter{"
                    "assetId assetName assetDescription functionalLocation "
                    "parentId parentName criticality assetStatus assetSegment "
                    "conditionIndex equipmentType}}"
                )
            }

        response = requests.post(
            url,
            headers=headers,
            json=payload,
            timeout=TIMEOUT
        )

        response.raise_for_status()
        result = response.json()

        all_assets.extend(result.get("data", []))

        cursor = result.get("nextCursor")
        if cursor is None:
            break

    return pd.DataFrame(all_assets)

# ===============================
# Execução
# ===============================
df_assets = get_all_assets()

print(df_assets.head())
print(f"\nTotal de ativos retornados: {len(df_assets)}")
