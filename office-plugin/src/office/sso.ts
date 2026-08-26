export interface OfficeSsoResult {
  ok: boolean;
  token?: string;
  source?: "OfficeRuntime.auth" | "Office.auth";
  message?: string;
}

export async function getOfficeSsoToken(): Promise<OfficeSsoResult> {
  const officeGlobals = globalThis as any;
  const runtimeAuth = officeGlobals.OfficeRuntime?.auth;
  const officeAuth = officeGlobals.Office?.auth;
  const options = { allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true };

  try {
    if (runtimeAuth?.getAccessToken) {
      const token = await runtimeAuth.getAccessToken(options);
      if (token) return { ok: true, token, source: "OfficeRuntime.auth" };
    }
    if (officeAuth?.getAccessToken) {
      const token = await officeAuth.getAccessToken(options);
      if (token) return { ok: true, token, source: "Office.auth" };
    }
    return { ok: false, message: "Office SSO is not available in this Office runtime or manifest." };
  } catch (error: any) {
    const code = error?.code ? ` (${error.code})` : "";
    return { ok: false, message: `${error?.message || String(error)}${code}` };
  }
}

export async function primeM365OfficeSso(fetcher: typeof fetch = fetch): Promise<OfficeSsoResult> {
  const tokenResult = await getOfficeSsoToken();
  if (!tokenResult.ok || !tokenResult.token) return tokenResult;

  const response = await fetcher("/api/m365/sso", {
    method: "POST",
    headers: { "content-type": "application/json", accept: "application/json" },
    body: JSON.stringify({ accessToken: tokenResult.token, source: tokenResult.source }),
  });
  const json = await response.json().catch(() => ({}));
  if (!response.ok) return { ok: false, message: json?.error?.message || `${response.status} ${response.statusText}` };
  return { ok: true, token: tokenResult.token, source: tokenResult.source, message: json?.message || "Office SSO token accepted." };
}
