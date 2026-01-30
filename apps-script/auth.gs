/************** CONFIGURAÇÕES **************/
const CLIENT_ID = "SUA_CLIENT_ID";
const CLIENT_SECRET = "SEU_CLIENT_SECRET";
const REFRESH_TOKEN = "SEU_REFRESH_TOKEN";

const BLING_TOKEN_URL = "https://www.bling.com.br/Api/v3/oauth/token";

/************** TOKEN **************/
function getAccessToken() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("bling_access_token");
  if (cached) return cached;

  const auth = Utilities.base64Encode(CLIENT_ID + ":" + CLIENT_SECRET);

  const response = UrlFetchApp.fetch(BLING_TOKEN_URL, {
    method: "post",
    contentType: "application/x-www-form-urlencoded",
    headers: {
      Authorization: "Basic " + auth
    },
    payload: {
      grant_type: "refresh_token",
      refresh_token: REFRESH_TOKEN
    },
    muteHttpExceptions: true
  });

  const json = JSON.parse(response.getContentText());
  if (!json.access_token) {
    throw new Error("Erro ao obter token");
  }

  cache.put("bling_access_token", json.access_token, 3500);
  return json.access_token;
}
