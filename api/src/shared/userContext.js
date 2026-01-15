// Extract user identity from Azure Static Web Apps headers.
// In SWA, the platform sets 'x-ms-client-principal' (base64 JSON).
// For local dev, you can pass ?as=email@domain.com (dev-only).

function decodeClientPrincipal(headerValue) {
  try {
    const json = Buffer.from(headerValue, "base64").toString("utf8");
    return JSON.parse(json);
  } catch {
    return null;
  }
}

export function getUserContext(request) {
  const principal = request.headers.get("x-ms-client-principal");
  if (principal) {
    const p = decodeClientPrincipal(principal);
    const claims = Array.isArray(p?.claims) ? p.claims : [];
    const getClaim = (t) => claims.find((c) => c.typ === t)?.val;
    return {
      authenticated: true,
      name: p?.userDetails || getClaim("name") || "",
      email: p?.userDetails || getClaim("preferred_username") || getClaim("email") || "",
      oid: getClaim("http://schemas.microsoft.com/identity/claims/objectidentifier") || getClaim("oid") || "",
      roles: Array.isArray(p?.userRoles) ? p.userRoles : [],
    };
  }

  // dev fallback
  const url = new URL(request.url);
  const as = url.searchParams.get("as");
  if (as) {
    return { authenticated: true, name: as, email: as, oid: as, roles: ["authenticated"] };
  }

  return { authenticated: false, name: "", email: "", oid: "", roles: [] };
}
