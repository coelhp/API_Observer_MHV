import requests
import pandas as pd

# ===============================
# Configurações da API
# ===============================
BASE_URL = "https://analystapi.repcenter.skf.com"
SHORTNAME = "BRAEO"  # Assim como para obter os assets, o shortnamente variam em 1 por unidade da indústria.
API_KEY = ""
TIMEOUT = 60

url = f"{BASE_URL}/{SHORTNAME}/workorders"

headers = {
    "x-api-key": API_KEY,
    "Content-Type": "application/json"
}

# ===============================
# Função para coletar workorders
# ===============================
def get_workorders(opening_date_start, opening_date_end):
    all_workorders = []
    cursor = None

    while True:
        if cursor:
            payload = {
                "expression": (
                    "{filter("
                    f"openingDateStart: \"{opening_date_start}\", "
                    f"openingDateEnd: \"{opening_date_end}\", "
                    f"cursor: {cursor}"
                    "){"
                    "assetId id orderNumber deadline priority technique "
                    "scheduledDate openingDate reWork cmmsRegister cmms "
                    "services situation author "
                    "intervention{date interventionType description isDiagnosticCorrect}"
                    "}}"
                )
            }
        else:
            payload = {
                "expression": (
                    "{filter("
                    f"openingDateStart: \"{opening_date_start}\", "
                    f"openingDateEnd: \"{opening_date_end}\""
                    "){"
                    "assetId id orderNumber deadline priority technique "
                    "scheduledDate openingDate reWork cmmsRegister cmms "
                    "services situation author "
                    "intervention{date interventionType description isDiagnosticCorrect}"
                    "}}"
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

        all_workorders.extend(result.get("data", []))

        cursor = result.get("nextCursor")
        if cursor is None:
            break

    return pd.DataFrame(all_workorders)

# ===============================
# Execução
# ===============================
df_workorders = get_workorders(
    opening_date_start="2024-01-01 00:00:00",
    opening_date_end="2028-12-31 23:59:59"
)

print(df_workorders.head())
print(f"\nTotal de workorders retornadas: {len(df_workorders)}")
