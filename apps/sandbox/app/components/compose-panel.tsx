"use client"

import {
  ChevronDown,
  ChevronRight,
  Loader2,
  SendHorizontal,
  Settings2,
} from "lucide-react"
import { useState, type ReactNode } from "react"
import { toast } from "sonner"

import { SlidingPills } from "./sliding-pills"

import { Button } from "@/components/ui/button"
import { Checkbox } from "@/components/ui/checkbox"
import { Input } from "@/components/ui/input"
import { Textarea } from "@/components/ui/textarea"
import { RichEditor } from "./rich-editor"
import type {
  SandboxDriverInfo,
  SandboxSnapshot,
  SendSandboxEmailInput,
} from "../sandbox/types"

type Draft = {
  fromEmail: string
  fromName: string
  toEmail: string
  ccEmail: string
  bccEmail: string
  subject: string
  bodyMode: "visual" | "html" | "text"
  text: string
  html: string
  replyToEmail: string
  inReplyToMessageId: string
  trackingOpens: boolean
  trackingClicks: boolean
  sendAt: string
  unsubscribeGlobal: boolean
  templateId: string
  templateData: string
  idempotencyKey: string
  tenantId: string
  outlookTenant: string
  tags: string
  metadata: string
  headers: string
}

const parseJson = <T,>(value: string, label: string): T | undefined => {
  if (!value.trim()) return undefined
  try {
    return JSON.parse(value) as T
  } catch {
    throw new Error(`${label} must be valid JSON.`)
  }
}

const defaultDraft = (driver: SandboxDriverInfo): Draft => ({
  fromEmail: driver.defaultFromEmail,
  fromName: "",
  toEmail: driver.defaultToEmail,
  ccEmail: "",
  bccEmail: "",
  subject: `EmailKit sandbox test via ${driver.label}`,
  bodyMode: "visual",
  text: "Plain text fallback from the EmailKit sandbox.",
  html: "<p>Hello from the <strong>EmailKit sandbox</strong>.</p>",
  replyToEmail: "",
  inReplyToMessageId: "",
  trackingOpens: true,
  trackingClicks: true,
  sendAt: "",
  unsubscribeGlobal: false,
  templateId: "",
  templateData: "",
  idempotencyKey: "",
  tenantId: "",
  outlookTenant: "",
  tags: "",
  metadata: "",
  headers: "",
})

export function ComposePanel({
  driver,
  onSnapshot,
}: {
  driver: SandboxDriverInfo
  onSnapshot: (snapshot: SandboxSnapshot) => void
}) {
  const [drafts, setDrafts] = useState<Record<string, Draft>>({})
  const [optionsOpen, setOptionsOpen] = useState(false)
  const [busy, setBusy] = useState(false)
  const draft = drafts[driver.id] ?? defaultDraft(driver)
  const caps = driver.capabilities
  const supportsTrackingOpens = caps.sendTracking?.opens === true
  const supportsTrackingClicks = caps.sendTracking?.clicks === true
  const isOutlookApp = driver.id === "outlook-app"

  const update = <K extends keyof Draft>(key: K, value: Draft[K]) =>
    setDrafts((current) => ({
      ...current,
      [driver.id]: { ...draft, [key]: value },
    }))

  const csv = (value: string) =>
    value
      .split(",")
      .map((entry) => entry.trim())
      .filter(Boolean)

  const send = async () => {
    setBusy(true)
    try {
      const track =
        supportsTrackingOpens || supportsTrackingClicks
          ? {
              opens: supportsTrackingOpens ? draft.trackingOpens : undefined,
              clicks: supportsTrackingClicks ? draft.trackingClicks : undefined,
            }
          : undefined

      const payload: SendSandboxEmailInput = {
        emailDriver: driver.id,
        fromEmail: draft.fromEmail,
        fromName: draft.fromName || undefined,
        toEmail: draft.toEmail,
        ccEmail: draft.ccEmail || undefined,
        bccEmail: draft.bccEmail || undefined,
        replyToEmail: draft.replyToEmail || undefined,
        inReplyToMessageId: draft.inReplyToMessageId || undefined,
        subject: draft.subject,
        text: draft.bodyMode === "text" ? draft.text : undefined,
        html: draft.bodyMode === "text" ? undefined : draft.html,
        sendAt: caps.scheduling && draft.sendAt ? draft.sendAt : undefined,
        templateId: caps.templates ? draft.templateId || undefined : undefined,
        templateData: caps.templates
          ? parseJson(draft.templateData, "Template data")
          : undefined,
        track,
        unsubscribe:
          caps.unsubscribe && draft.unsubscribeGlobal
            ? { global: true }
            : undefined,
        idempotencyKey: caps.sendIdempotency
          ? draft.idempotencyKey || undefined
          : undefined,
        tenantId: caps.tenantRouting ? draft.tenantId || undefined : undefined,
        outlookTenant: isOutlookApp
          ? draft.outlookTenant.trim() || undefined
          : undefined,
        tags: csv(draft.tags),
        metadata: parseJson<Record<string, string>>(draft.metadata, "Metadata"),
        headers: parseJson<Record<string, string>>(
          draft.headers,
          "Custom headers"
        ),
      }

      const response = await fetch("/api/sandbox/send", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify(payload),
      })
      const data = await response.json()
      if (!response.ok || !data.ok) throw new Error(data.error || "Send failed")
      onSnapshot(data.snapshot as SandboxSnapshot)
      toast.success(`Sent via ${driver.label}`)
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "Send failed")
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="flex flex-col gap-3 p-3">
      <div className="grid grid-cols-2 gap-2">
        <Field label="From">
          <Input
            value={draft.fromEmail}
            onChange={(event) => update("fromEmail", event.target.value)}
            placeholder="sender@example.com"
            className="h-7 text-xs"
          />
        </Field>
        <Field label="Name">
          <Input
            value={draft.fromName}
            onChange={(event) => update("fromName", event.target.value)}
            placeholder="(optional)"
            className="h-7 text-xs"
          />
        </Field>
      </div>

      {isOutlookApp && (
        <Field label="Tenant">
          <Input
            value={draft.outlookTenant}
            onChange={(event) => update("outlookTenant", event.target.value)}
            placeholder="contoso.onmicrosoft.com or tenant id (optional)"
            className="h-7 text-xs"
          />
        </Field>
      )}

      <Field label="To">
        <Input
          value={draft.toEmail}
          onChange={(event) => update("toEmail", event.target.value)}
          placeholder="recipient@example.com"
          className="h-7 text-xs"
        />
      </Field>

      <div className="grid grid-cols-2 gap-2">
        <Field label="CC">
          <Input
            value={draft.ccEmail}
            onChange={(event) => update("ccEmail", event.target.value)}
            placeholder="(optional)"
            className="h-7 text-xs"
          />
        </Field>
        <Field label="BCC">
          <Input
            value={draft.bccEmail}
            onChange={(event) => update("bccEmail", event.target.value)}
            placeholder="(optional)"
            className="h-7 text-xs"
          />
        </Field>
      </div>

      <Field label="Subject">
        <Input
          value={draft.subject}
          onChange={(event) => update("subject", event.target.value)}
          className="h-7 text-xs"
        />
      </Field>

      <div className="flex flex-col gap-1.5">
        <div className="flex items-center justify-between">
          <label className="text-[10px] tracking-wider text-muted-foreground uppercase">
            Body
          </label>
          <SlidingPills
            value={draft.bodyMode}
            onChange={(mode) => update("bodyMode", mode)}
            options={[
              { value: "visual", label: "Visual" },
              { value: "html", label: "HTML" },
              { value: "text", label: "Text" },
            ]}
          />
        </div>
        {draft.bodyMode === "visual" && (
          <RichEditor
            value={draft.html}
            onChange={(value) => update("html", value)}
          />
        )}
        {draft.bodyMode === "html" && (
          <Textarea
            rows={5}
            value={draft.html}
            onChange={(event) => update("html", event.target.value)}
            className="font-mono text-xs"
            placeholder="<p>Hello</p>"
          />
        )}
        {draft.bodyMode === "text" && (
          <Textarea
            rows={5}
            value={draft.text}
            onChange={(event) => update("text", event.target.value)}
            className="text-xs"
            placeholder="Plain text body"
          />
        )}
      </div>

      <button
        onClick={() => setOptionsOpen((current) => !current)}
        className="flex items-center gap-1.5 pt-1 text-[10px] font-medium tracking-wider text-muted-foreground uppercase transition-colors hover:text-foreground"
      >
        <Settings2 className="size-3" />
        Advanced
        {optionsOpen ? (
          <ChevronDown className="size-3" />
        ) : (
          <ChevronRight className="size-3" />
        )}
      </button>

      {optionsOpen && (
        <div className="flex flex-col gap-3 rounded-lg border p-3">
          <div className="grid grid-cols-2 gap-2">
            <Field label="Reply-To">
              <Input
                value={draft.replyToEmail}
                onChange={(event) => update("replyToEmail", event.target.value)}
                className="h-7 text-xs"
              />
            </Field>
            <Field label="In-Reply-To">
              <Input
                value={draft.inReplyToMessageId}
                onChange={(event) =>
                  update("inReplyToMessageId", event.target.value)
                }
                placeholder="<msg-id@provider>"
                className="h-7 text-xs"
              />
            </Field>
          </div>

          {(supportsTrackingOpens || supportsTrackingClicks) && (
            <div className="flex items-center gap-4">
              <ToggleField
                label="Track opens"
                checked={draft.trackingOpens}
                disabled={!supportsTrackingOpens}
                onChange={(value) => update("trackingOpens", value)}
              />
              <ToggleField
                label="Track clicks"
                checked={draft.trackingClicks}
                disabled={!supportsTrackingClicks}
                onChange={(value) => update("trackingClicks", value)}
              />
            </div>
          )}

          <CapabilityGroup caps={caps} keys={["scheduling"]}>
            <Field label="Schedule (ISO)">
              <Input
                value={draft.sendAt}
                placeholder="2026-05-21T12:00:00Z"
                onChange={(event) => update("sendAt", event.target.value)}
                className="h-7 text-xs"
              />
            </Field>
          </CapabilityGroup>

          <CapabilityGroup caps={caps} keys={["unsubscribe"]}>
            <ToggleField
              label="Global unsubscribe header"
              checked={draft.unsubscribeGlobal}
              onChange={(value) => update("unsubscribeGlobal", value)}
            />
          </CapabilityGroup>

          <CapabilityGroup caps={caps} keys={["templates"]}>
            <Field label="Template ID">
              <Input
                value={draft.templateId}
                onChange={(event) => update("templateId", event.target.value)}
                className="h-7 text-xs"
              />
            </Field>
            <Field label="Template data (JSON)">
              <Textarea
                rows={2}
                value={draft.templateData}
                onChange={(event) => update("templateData", event.target.value)}
                className="font-mono text-xs"
                placeholder='{"name": "Pau"}'
              />
            </Field>
          </CapabilityGroup>

          <CapabilityGroup caps={caps} keys={["sendIdempotency"]}>
            <Field label="Idempotency key">
              <Input
                value={draft.idempotencyKey}
                onChange={(event) =>
                  update("idempotencyKey", event.target.value)
                }
                className="h-7 text-xs"
              />
            </Field>
          </CapabilityGroup>

          <CapabilityGroup caps={caps} keys={["tenantRouting"]}>
            <Field label="Tenant ID">
              <Input
                value={draft.tenantId}
                onChange={(event) => update("tenantId", event.target.value)}
                className="h-7 text-xs"
              />
            </Field>
          </CapabilityGroup>

          <Field label="Tags (comma-separated)">
            <Input
              value={draft.tags}
              onChange={(event) => update("tags", event.target.value)}
              placeholder="sandbox, test"
              className="h-7 text-xs"
            />
          </Field>
          <Field label="Metadata (JSON)">
            <Textarea
              rows={2}
              value={draft.metadata}
              onChange={(event) => update("metadata", event.target.value)}
              placeholder='{"key": "value"}'
              className="font-mono text-xs"
            />
          </Field>
          <Field label="Custom headers (JSON)">
            <Textarea
              rows={2}
              value={draft.headers}
              onChange={(event) => update("headers", event.target.value)}
              placeholder='{"X-Custom": "value"}'
              className="font-mono text-xs"
            />
          </Field>
        </div>
      )}

      <Button
        onClick={send}
        disabled={busy || !driver.ready}
        size="sm"
        className="self-start"
      >
        {busy ? (
          <Loader2
            className="size-3 animate-spin"
            style={{ animationDuration: "0.6s" }}
          />
        ) : (
          <SendHorizontal className="size-3" />
        )}
        {busy ? "Sending" : "Send"}
      </Button>
      {!driver.ready && (
        <p className="font-mono text-[10px] text-destructive">
          Missing env: {driver.missingRequiredEnv.join(", ")}
        </p>
      )}
    </div>
  )
}

function Field({ label, children }: { label: string; children: ReactNode }) {
  return (
    <label className="grid gap-1">
      <span className="text-[10px] tracking-wider text-muted-foreground uppercase">
        {label}
      </span>
      {children}
    </label>
  )
}

function ToggleField({
  label,
  checked,
  disabled,
  onChange,
}: {
  label: string
  checked: boolean
  disabled?: boolean
  onChange: (value: boolean) => void
}) {
  return (
    <label className="flex items-center gap-1.5 text-xs">
      <Checkbox
        disabled={disabled}
        checked={checked}
        onCheckedChange={(value) => onChange(Boolean(value))}
      />
      {label}
    </label>
  )
}

function CapabilityGroup({
  caps,
  keys,
  children,
}: {
  caps: SandboxDriverInfo["capabilities"]
  keys: (keyof SandboxDriverInfo["capabilities"])[]
  children: ReactNode
}) {
  const anySupported = keys.some((key) => Boolean(caps[key]))
  if (!anySupported) return null
  return <div>{children}</div>
}
