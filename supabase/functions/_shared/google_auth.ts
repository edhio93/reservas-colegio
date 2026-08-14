export type GoogleServiceAccount = {
  client_email: string;
  private_key: string;
  token_uri?: string;
};

function base64Url(input: Uint8Array | string): string {
  const bytes =
    typeof input === "string"
      ? new TextEncoder().encode(input)
      : input;

  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);

  return btoa(binary)
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
}

function pemToArrayBuffer(pem: string): ArrayBuffer {
  const clean = pem
    .replace(/-----BEGIN PRIVATE KEY-----/g, "")
    .replace(/-----END PRIVATE KEY-----/g, "")
    .replace(/\s+/g, "");

  const binary = atob(clean);
  const bytes = new Uint8Array(binary.length);

  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }

  return bytes.buffer;
}

export function loadServiceAccount(): GoogleServiceAccount {
  const raw = Deno.env.get("GOOGLE_SERVICE_ACCOUNT_JSON");

  if (!raw) {
    throw new Error(
      "Falta GOOGLE_SERVICE_ACCOUNT_JSON en Edge Function Secrets.",
    );
  }

  const parsed = JSON.parse(raw);

  if (!parsed.client_email || !parsed.private_key) {
    throw new Error(
      "GOOGLE_SERVICE_ACCOUNT_JSON no contiene client_email/private_key.",
    );
  }

  return parsed;
}

export async function getGoogleAccessToken(
  subject: string,
  scopes: string[],
): Promise<string> {
  if (!subject) {
    throw new Error("Falta usuario a impersonar en Google Workspace.");
  }

  const serviceAccount = loadServiceAccount();

  const now = Math.floor(Date.now() / 1000);
  const tokenUri =
    serviceAccount.token_uri || "https://oauth2.googleapis.com/token";

  const header = {
    alg: "RS256",
    typ: "JWT",
  };

  const claimSet = {
    iss: serviceAccount.client_email,
    sub: subject,
    scope: scopes.join(" "),
    aud: tokenUri,
    iat: now,
    exp: now + 3600,
  };

  const unsignedToken =
    `${base64Url(JSON.stringify(header))}.${base64Url(JSON.stringify(claimSet))}`;

  const key = await crypto.subtle.importKey(
    "pkcs8",
    pemToArrayBuffer(serviceAccount.private_key),
    {
      name: "RSASSA-PKCS1-v1_5",
      hash: "SHA-256",
    },
    false,
    ["sign"],
  );

  const signature = new Uint8Array(
    await crypto.subtle.sign(
      "RSASSA-PKCS1-v1_5",
      key,
      new TextEncoder().encode(unsignedToken),
    ),
  );

  const assertion = `${unsignedToken}.${base64Url(signature)}`;

  const body = new URLSearchParams({
    grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
    assertion,
  });

  const response = await fetch(tokenUri, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body,
  });

  const data = await response.json();

  if (!response.ok || !data.access_token) {
    throw new Error(
      `Google OAuth error ${response.status}: ${JSON.stringify(data)}`,
    );
  }

  return data.access_token;
}

export function base64UrlUtf8(text: string): string {
  return base64Url(new TextEncoder().encode(text));
}
