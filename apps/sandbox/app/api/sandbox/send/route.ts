import { jsonError, jsonOk, sendSandboxEmail } from "@/app/sandbox/actions"
import { getSandboxSnapshot } from "@/app/sandbox/store"
import type { SendSandboxEmailInput } from "@/app/sandbox/types"

export const runtime = "nodejs"
export const dynamic = "force-dynamic"

const cleanString = (value: unknown) =>
  typeof value === "string" && value.trim() ? value.trim() : undefined

const cleanStringArray = (value: unknown) =>
  Array.isArray(value)
    ? value.filter((entry): entry is string => typeof entry === "string" && entry.trim() !== "")
    : undefined

const cleanRecord = <T extends Record<string, unknown>>(value: unknown) =>
  value && typeof value === "object" && !Array.isArray(value) ? (value as T) : undefined

const parseSendInput = (body: unknown): SendSandboxEmailInput => {
  if (!body || typeof body !== "object") throw new Error("Request body must be an object.")
  const record = body as Record<string, unknown>
  const emailDriver = cleanString(record.emailDriver)
  const fromEmail = cleanString(record.fromEmail)
  const toEmail = cleanString(record.toEmail)
  const subject = cleanString(record.subject)
  if (!emailDriver || !fromEmail || !toEmail || !subject) {
    throw new Error("emailDriver, fromEmail, toEmail, and subject are required.")
  }

  return {
    emailDriver,
    fromEmail,
    toEmail,
    subject,
    fromName: cleanString(record.fromName),
    ccEmail: cleanString(record.ccEmail),
    bccEmail: cleanString(record.bccEmail),
    replyToEmail: cleanString(record.replyToEmail),
    inReplyToMessageId: cleanString(record.inReplyToMessageId),
    text: cleanString(record.text),
    html: cleanString(record.html),
    sendAt: cleanString(record.sendAt),
    templateId: cleanString(record.templateId),
    templateData: cleanRecord(record.templateData),
    track: cleanRecord(record.track),
    unsubscribe: cleanRecord(record.unsubscribe),
    idempotencyKey: cleanString(record.idempotencyKey),
    tenantId: cleanString(record.tenantId),
    outlookTenant: cleanString(record.outlookTenant),
    tags: cleanStringArray(record.tags),
    metadata: cleanRecord<Record<string, string>>(record.metadata),
    headers: cleanRecord<Record<string, string>>(record.headers),
    provider: cleanRecord(record.provider),
  }
}

export async function POST(request: Request) {
  try {
    const result = await sendSandboxEmail(parseSendInput(await request.json()))
    return jsonOk({ result, snapshot: getSandboxSnapshot() })
  } catch (error) {
    return jsonError(error)
  }
}
