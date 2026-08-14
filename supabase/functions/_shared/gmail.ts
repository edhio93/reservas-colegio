import {
  base64UrlUtf8,
  getGoogleAccessToken,
} from "./google_auth.ts";

const GMAIL_SEND_SCOPE =
  "https://www.googleapis.com/auth/gmail.send";

function sanitizeHeader(value: string): string {
  return String(value || "").replace(/[\r\n]+/g, " ").trim();
}

export async function sendGmailHtml(args: {
  to: string;
  subject: string;
  html: string;
}): Promise<Record<string, unknown>> {
  const sender = Deno.env.get("GOOGLE_GMAIL_SENDER") || "";

  if (!sender) {
    throw new Error("Falta GOOGLE_GMAIL_SENDER.");
  }

  const token = await getGoogleAccessToken(
    sender,
    [GMAIL_SEND_SCOPE],
  );

  const mime = [
    `From: ${sanitizeHeader(sender)}`,
    `To: ${sanitizeHeader(args.to)}`,
    `Subject: ${sanitizeHeader(args.subject)}`,
    "MIME-Version: 1.0",
    'Content-Type: text/html; charset="UTF-8"',
    "Content-Transfer-Encoding: 8bit",
    "",
    args.html,
  ].join("\r\n");

  const response = await fetch(
    "https://gmail.googleapis.com/gmail/v1/users/me/messages/send",
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        raw: base64UrlUtf8(mime),
      }),
    },
  );

  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      `Gmail API ${response.status}: ${JSON.stringify(data)}`,
    );
  }

  return data;
}
