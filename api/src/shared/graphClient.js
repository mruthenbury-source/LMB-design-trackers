import "isomorphic-fetch";
import { Client } from "@microsoft/microsoft-graph-client";

let cachedToken = null;
let cachedTokenExp = 0;

async function getAppToken() {
  const tenantId = process.env.AAD_TENANT_ID;
  const clientId = process.env.AAD_CLIENT_ID;
  const clientSecret = process.env.AAD_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) {
    throw new Error("Missing AAD_TENANT_ID, AAD_CLIENT_ID, or AAD_CLIENT_SECRET");
  }

  const now = Math.floor(Date.now() / 1000);
  if (cachedToken && cachedTokenExp - 60 > now) return cachedToken;

  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams();
  body.set("client_id", clientId);
  body.set("client_secret", clientSecret);
  body.set("grant_type", "client_credentials");
  body.set("scope", "https://graph.microsoft.com/.default");

  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Token request failed: ${res.status} ${text}`);
  }

  const json = await res.json();
  cachedToken = json.access_token;
  cachedTokenExp = now + (json.expires_in || 300);
  return cachedToken;
}

export async function getGraphClient() {
  const token = await getAppToken();
  return Client.init({
    authProvider: (done) => done(null, token),
  });
}
