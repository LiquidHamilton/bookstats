function env(name: string): string | undefined {
  return process.env[name]?.trim() || undefined;
}

function resendApiKey(): string | undefined {
  return env("BOOKSTATS_RESEND_API_KEY");
}

function publicUrl(): string {
  return (env("BOOKSTATS_PUBLIC_URL") || "http://localhost:5173/").replace(/\/?$/, "/");
}

function sender(): string | undefined {
  return env("BOOKSTATS_EMAIL_FROM") || env("BOOKSTATS_EMAIL_SENDER");
}

function replyTo(): string | undefined {
  return env("BOOKSTATS_EMAIL_REPLY_TO");
}

export function emailConfigured(): boolean {
  return Boolean(resendApiKey() && sender());
}

export function emailProvider(): string {
  return emailConfigured() ? "resend" : "unconfigured";
}

function linkWithToken(parameter: "verify" | "reset", token: string): string {
  const url = new URL(publicUrl());
  url.searchParams.set(parameter, token);
  return url.toString();
}

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (character) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  })[character] ?? character);
}

async function sendResendMail(
  to: string,
  subject: string,
  html: string,
  category: "verify_email" | "password_reset" | "password_changed" | "feedback"
): Promise<void> {
  const apiKey = resendApiKey();
  const from = sender();
  if (!apiKey || !from) {
    throw new Error("Resend email delivery is not fully configured.");
  }

  const body: Record<string, unknown> = {
    from,
    to: [to],
    subject,
    html,
    tags: [{ name: "category", value: category }]
  };

  const configuredReplyTo = replyTo();
  if (configuredReplyTo) body.reply_to = configuredReplyTo;

  const response = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: {
      authorization: `Bearer ${apiKey}`,
      "content-type": "application/json"
    },
    body: JSON.stringify(body)
  });

  if (!response.ok) {
    const details = await response.text();
    throw new Error(`Resend email request failed (${response.status}): ${details.slice(0, 1200)}`);
  }
}

export async function sendVerificationEmail(email: string, displayName: string, token: string): Promise<void> {
  const link = linkWithToken("verify", token);
  const safeName = escapeHtml(displayName);
  await sendResendMail(
    email,
    "Verify your BookStats email",
    `<p>Hi ${safeName},</p><p>Verify your BookStats email address to enable cloud synchronization.</p><p><a href="${escapeHtml(link)}">Verify email address</a></p><p>This link expires in 24 hours. If you did not create a BookStats account, you can ignore this message.</p>`,
    "verify_email"
  );
}

export async function sendPasswordResetEmail(email: string, displayName: string, token: string): Promise<void> {
  const link = linkWithToken("reset", token);
  const safeName = escapeHtml(displayName);
  await sendResendMail(
    email,
    "Reset your BookStats password",
    `<p>Hi ${safeName},</p><p>A password reset was requested for your BookStats account.</p><p><a href="${escapeHtml(link)}">Choose a new password</a></p><p>This link expires in 1 hour. If you did not request a password reset, you can ignore this message.</p>`,
    "password_reset"
  );
}

export async function sendPasswordChangedEmail(email: string, displayName: string): Promise<void> {
  const safeName = escapeHtml(displayName);
  await sendResendMail(
    email,
    "Your BookStats password was changed",
    `<p>Hi ${safeName},</p><p>The password for your BookStats account was changed.</p><p>If you did not make this change, contact the BookStats administrator immediately.</p>`,
    "password_changed"
  );
}


function feedbackRecipient(): string | undefined {
  const configured = env("BOOKSTATS_FEEDBACK_TO") || replyTo();
  if (configured) return configured;
  const from = sender();
  if (!from) return undefined;
  return from.match(/<([^>]+)>/)?.[1] ?? from;
}

export function feedbackConfigured(): boolean {
  return emailConfigured() && Boolean(feedbackRecipient());
}

export async function sendFeedbackEmail(input: {
  kind: "bug" | "feature";
  message: string;
  contactEmail?: string;
  diagnostics: Record<string, string | number | boolean | undefined>;
}): Promise<void> {
  const to = feedbackRecipient();
  if (!to) throw new Error("BookStats feedback delivery is not configured.");
  const title = input.kind === "bug" ? "Bug report" : "Feature suggestion";
  const rows = Object.entries(input.diagnostics)
    .filter(([, value]) => value !== undefined && value !== "")
    .map(([key, value]) => `<tr><td style="padding:4px 12px 4px 0;color:#6b6b63">${escapeHtml(key)}</td><td style="padding:4px 0">${escapeHtml(String(value))}</td></tr>`)
    .join("");
  const contact = input.contactEmail ? `<p><strong>Contact:</strong> ${escapeHtml(input.contactEmail)}</p>` : "";
  await sendResendMail(
    to,
    `[BookStats] ${title} · v${escapeHtml(String(input.diagnostics.version ?? "unknown"))}`,
    `<h2>${title}</h2>${contact}<p style="white-space:pre-wrap">${escapeHtml(input.message)}</p><h3>Diagnostics</h3><table>${rows}</table>`,
    "feedback"
  );
}
