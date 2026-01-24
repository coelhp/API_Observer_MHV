import requests
import pandas as pd

# =========================
# CONFIGURAÇÕES | Etapa para definição do link, porta e usuário da API Observer
# =========================
base_url = "http://services.repcenter.skf.com"
port = ""  # 
username = ""
password = ""
timeout = 60  # segundos

# =========================
# TOKEN | Obtenção do TOKEN para requisição na API
# =========================
token_url = f"{base_url}:{port}/token"
payload = {
    "grant_type": "password",
    "username": username,
    "password": password
}
headers = {"Content-Type": "application/x-www-form-urlencoded"}

try:
    token_response = requests.post(token_url, data=payload, headers=headers, timeout=timeout)
    token = token_response.json().get("access_token") if token_response.status_code == 200 else None
except:
    token = None

# =========================
# Função GET para obtenção de dados
# =========================
def get_data(endpoint):
    try:
        url = f"{base_url}:{port}{endpoint}"
        headers = {"Authorization": f"Bearer {token}"}
        r = requests.get(url, headers=headers, timeout=timeout)
        return r.json() if r.status_code == 200 else None
    except:
        return None

# =========================
# Endpoint ASSETS
# =========================
assets = get_data("/v2/assets")
df_assets = pd.DataFrame(assets) if assets else pd.DataFrame()

# =========================
# Endpoint POINTS considerando o id de todos os assets do endpoint anterior
# =========================
points_list = []

if not df_assets.empty and "ID" in df_assets.columns:
    for asset_id in df_assets["ID"].dropna().unique():
        points = get_data(f"/v1/machines/{asset_id}/points")
        if points:
            df_tmp = pd.DataFrame(points)
            df_tmp["AssetID"] = asset_id  # relacionamento
            points_list.append(df_tmp)

# =========================
# Criando df com as respostas da requisição
# =========================
if points_list:
    df_points = pd.concat(points_list, ignore_index=True)
else:
    df_points = pd.DataFrame([{"erro": "Falha ao obter points"}])

df_points
