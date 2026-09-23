#!/usr/bin/env node

import dotenv from "dotenv";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawn } from "node:child_process";
import { createRequire } from "node:module";
import { fileURLToPath, pathToFileURL } from "node:url";
import { PublicClientApplication } from "@azure/msal-node";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import ExcelJS from "exceljs";
import mammoth from "mammoth";
import AdmZip from "adm-zip";

const require = createRequire(import.meta.url);
const _pdfParseRaw = require("pdf-parse");
const CFB = require("cfb");

let parsePdf;
if (typeof _pdfParseRaw === "function") {
  parsePdf = _pdfParseRaw;
} else if (_pdfParseRaw && typeof _pdfParseRaw.PDFParse === "function") {
  parsePdf = async (buffer) => {
    const instance = new _pdfParseRaw.PDFParse();
    return instance.parse(buffer);
  };
} else {
  parsePdf = async () => {
    throw new Error(
      "pdf-parse module loaded but no usable export was found. Please pin pdf-parse to ^1.1.1."
    );
  };
}

dotenv.config({ quiet: true });

const APP_NAME = "m365-attachment-reader-mcp-local";
const APP_DATA_DIR = resolveAppDataDir();

function resolveAppDataDir() {
  const moduleDir = path.dirname(fileURLToPath(import.meta.url));
  const candidates = [
    process.env.M365_MCP_DATA_DIR,
    process.platform === "win32"
      ? path.join(process.env.LOCALAPPDATA || os.homedir(), APP_NAME)
      : path.join(os.homedir(), `.${APP_NAME}`),
    path.join(moduleDir, `.${APP_NAME}`),
  ].filter(Boolean);

  for (const candidate of candidates) {
    try {
      fs.mkdirSync(candidate, { recursive: true });
      return candidate;
    } catch {
      // Try the next writable location.
    }
  }

  throw new Error(
    `Unable to create an application data directory. Tried: ${candidates.join(", ")}`
  );
}

const DEBUG_LOG = path.join(APP_DATA_DIR, "debug.log");
const CLIENT_ID = process.env.M365_CLIENT_ID;
const TENANT_ID = process.env.M365_TENANT_ID || "common";
const AUTHORITY = `https://login.microsoftonline.com/${TENANT_ID}`;
// Mail.ReadWrite subsumes Mail.Read (used by the read tools) and is required to
// create drafts; Mail.Send is required to send mail. Adding these scopes means
// users must re-consent (add the delegated permissions in Entra, then re-run
// begin_auth) — an existing token does not gain new scopes automatically.
const SCOPES = ["User.Read", "Mail.ReadWrite", "Mail.Send"];
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const AUTO_OPEN_BROWSER =
  String(process.env.M365_AUTO_OPEN_BROWSER || "true").toLowerCase() !== "false";

const MAX_RETURN_CHARS = 35000;
const MAX_EMAIL_BODY_CHUNK_CHARS = 12000;
const MAX_IMAGE_BLOCKS = 20;
const MAX_IMAGE_BLOCK_BYTES = 450 * 1024;
const MAX_TOTAL_IMAGE_BYTES = 700 * 1024;
const IMAGE_RESIZE_STEPS = [1600, 1280, 1024, 800, 640, 480];
const MAX_DOC_IMAGES = 15;
const MAX_PDF_OCR_PAGES = 5;
const MAX_PDF_VISUAL_PAGES = 3;
const MAX_ARCHIVE_ITEMS = 12;
const MAX_ARCHIVE_ENTRY_BYTES = 20 * 1024 * 1024;
const MAX_NESTED_DEPTH = 2;
const MAX_MSG_ATTACHMENTS = 5;

// Outbound send-with-attachment limits. Microsoft Graph inlines base64
// attachments only when each file is < 3 MB and the whole write request stays
// under ~4 MB (base64 inflates binary by ~33%). We keep a conservative total
// budget so multi-file sends don't blow the request cap.
const MAX_SEND_FILE_BYTES = 3 * 1024 * 1024;
const MAX_SEND_TOTAL_BYTES = Number(process.env.M365_SEND_MAX_BYTES) || 3 * 1024 * 1024;

// Optional allowlist of base directories that local attachments must live under.
// Empty (default) means any readable path is allowed. Use the OS path separator
// (":" on macOS/Linux, ";" on Windows) to list multiple roots.
const ALLOWED_SEND_DIRS = String(process.env.M365_SEND_ALLOWED_DIRS || "")
  .split(path.delimiter)
  .map((dir) => dir.trim())
  .filter(Boolean)
  // Ignore unsubstituted DXT template placeholders (e.g. "${user_config.x}")
  // that some clients pass literally when an optional config field is left blank.
  // Without this, a blank allowlist would resolve to one bogus path and reject
  // every real file. Empty/placeholder => no allowlist => any readable path.
  .filter((dir) => !dir.includes("${"))
  .map((dir) => path.resolve(dir));

const IMAGE_MIME_BY_EXT = {
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  png: "image/png",
  gif: "image/gif",
  webp: "image/webp",
  bmp: "image/bmp",
  tif: "image/tiff",
  tiff: "image/tiff",
};

// Content types for outbound attachments, keyed by file extension. Falls back to
// application/octet-stream for anything unlisted. Includes the image types so
// loadLocalAttachment can resolve every supported format from a single map.
const MIME_BY_EXT = {
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  pptm: "application/vnd.ms-powerpoint.presentation.macroEnabled.12",
  ppsx: "application/vnd.openxmlformats-officedocument.presentationml.slideshow",
  potx: "application/vnd.openxmlformats-officedocument.presentationml.template",
  ppt: "application/vnd.ms-powerpoint",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  doc: "application/msword",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  xlsm: "application/vnd.ms-excel.sheet.macroEnabled.12",
  xls: "application/vnd.ms-excel",
  pdf: "application/pdf",
  csv: "text/csv",
  txt: "text/plain",
  md: "text/markdown",
  json: "application/json",
  xml: "application/xml",
  html: "text/html",
  zip: "application/zip",
  ...IMAGE_MIME_BY_EXT,
};

let pdfOcrDepsPromise = null;
let msgReaderPromise = null;
let rarDepsPromise = null;
let sevenZipDepsPromise = null;
let sharpPromise = null;

function log(...parts) {
  const line =
    `[${new Date().toISOString()}] ` +
    parts
      .map((part) => {
        if (typeof part === "string") return part;
        if (part instanceof Error) return part.stack || part.message;
        try {
          return JSON.stringify(part);
        } catch {
          return String(part);
        }
      })
      .join(" ") +
    "\n";

  fs.appendFileSync(DEBUG_LOG, line, "utf8");
}

process.on("uncaughtException", (err) => {
  log("uncaughtException", err);
  console.error(err);
});

process.on("unhandledRejection", (err) => {
  log("unhandledRejection", err);
  console.error(err);
});

function getDeviceLoginUrlFromMessage(message) {
  const text = String(message || "");
  const match = text.match(/https?:\/\/[^\s]+/i);
  return match ? match[0] : "https://microsoft.com/devicelogin";
}

function openUrlInBrowser(url) {
  if (!AUTO_OPEN_BROWSER || !url || typeof url !== "string") {
    return false;
  }

  try {
    if (process.platform === "win32") {
      spawn("cmd", ["/c", "start", "", url], {
        detached: true,
        stdio: "ignore",
      }).unref();
      return true;
    }

    if (process.platform === "darwin") {
      spawn("open", [url], {
        detached: true,
        stdio: "ignore",
      }).unref();
      return true;
    }

    spawn("xdg-open", [url], {
      detached: true,
      stdio: "ignore",
    }).unref();
    return true;
  } catch (err) {
    log("openUrlInBrowser failed", err);
    return false;
  }
}

if (!CLIENT_ID) {
  log("Missing M365_CLIENT_ID in env");
}

const MSAL_CACHE_PATH = path.join(APP_DATA_DIR, "msal-cache.json");

// Persist MSAL's token cache (including the refresh token) to disk so auth
// survives app/process restarts. With this the user signs in once and the
// server can silently refresh access tokens going forward (refresh tokens roll
// for ~90 days), which is what makes unattended scheduled runs reliable.
// The file holds sensitive tokens, so it is written with 0600 permissions.
const cachePlugin = {
  async beforeCacheAccess(context) {
    try {
      if (fs.existsSync(MSAL_CACHE_PATH)) {
        context.tokenCache.deserialize(fs.readFileSync(MSAL_CACHE_PATH, "utf8"));
      }
    } catch (err) {
      log("MSAL cache read failed", err?.message || String(err));
    }
  },
  async afterCacheAccess(context) {
    try {
      if (context.cacheHasChanged) {
        fs.writeFileSync(MSAL_CACHE_PATH, context.tokenCache.serialize(), {
          mode: 0o600,
        });
      }
    } catch (err) {
      log("MSAL cache write failed", err?.message || String(err));
    }
  },
};

const pca = new PublicClientApplication({
  auth: {
    clientId: CLIENT_ID || "missing-client-id",
    authority: AUTHORITY,
  },
  cache: { cachePlugin },
});

let currentAccount = null;
let cachedAccessToken = null;
let cachedTokenExpiresAt = 0;
let authPromise = null;

const authState = {
  status: "not_started",
  verificationUri: null,
  userCode: null,
  message: null,
  account: null,
  error: null,
  updatedAt: null,
};

function setAuthState(patch) {
  Object.assign(authState, patch, {
    updatedAt: new Date().toISOString(),
  });
  log("authState", authState);
}

function getPublicAuthState() {
  return {
    ...authState,
    configured: !!CLIENT_ID,
    hasCachedToken:
      !!cachedAccessToken && Date.now() < cachedTokenExpiresAt - 60 * 1000,
    tokenExpiresAt: cachedTokenExpiresAt
      ? new Date(cachedTokenExpiresAt).toISOString()
      : null,
  };
}

function ensureConfigured() {
  if (!CLIENT_ID) {
    throw new Error(
      "Microsoft 365 is not configured. Set M365_CLIENT_ID in the environment before using auth tools."
    );
  }
}

function getCachedToken() {
  if (
    cachedAccessToken &&
    cachedTokenExpiresAt &&
    Date.now() < cachedTokenExpiresAt - 60 * 1000
  ) {
    return cachedAccessToken;
  }

  return null;
}

function buildBeginAuthText() {
  const parts = ["Please complete Microsoft 365 login now."];

  if (authState.verificationUri) {
    parts.push(`Open URL: ${authState.verificationUri}`);
  }

  if (authState.userCode) {
    parts.push(`Enter code: ${authState.userCode}`);
  }

  if (authState.message) {
    parts.push(`Full prompt: ${authState.message}`);
  }

  if (!AUTO_OPEN_BROWSER) {
    parts.push("Browser auto-open is disabled by M365_AUTO_OPEN_BROWSER=false.");
  }

  return parts.join("\n");
}

async function startDeviceCodeAuth() {
  if (authPromise) return authPromise;

  ensureConfigured();

  authPromise = pca
    .acquireTokenByDeviceCode({
      scopes: SCOPES,
      deviceCodeCallback: (response) => {
        log("deviceCodeCallback RAW RESPONSE", JSON.stringify(response));
        const message = response?.message || "";
        const verificationUri = getDeviceLoginUrlFromMessage(message);
        const userCode = response?.userCode || response?.user_code || null;
        const opened = openUrlInBrowser(verificationUri);

        setAuthState({
          status: "pending",
          verificationUri,
          userCode,
          message,
          error: null,
        });

        log("deviceCodeCallback", {
          verificationUri,
          userCode,
          message,
          opened,
        });
      },
    })
    .then((result) => {
      if (!result?.accessToken) {
        throw new Error("No access token returned.");
      }

      currentAccount = result.account ?? null;
      cachedAccessToken = result.accessToken;
      cachedTokenExpiresAt = result.expiresOn
        ? result.expiresOn.getTime()
        : Date.now() + 50 * 60 * 1000;

      setAuthState({
        status: "authenticated",
        account: currentAccount?.username || null,
        error: null,
      });

      return result.accessToken;
    })
    .catch((err) => {
      log("MSAL ERROR FULL DETAIL", {
        name: err?.name,
        message: err?.message,
        errorCode: err?.errorCode,
        errorMessage: err?.errorMessage,
        subError: err?.subError,
        correlationId: err?.correlationId,
        stack: err?.stack,
      });
      setAuthState({
        status: "error",
        error: String(err?.message || err),
      });
      throw err;
    })
    .finally(() => {
      authPromise = null;
    });

  return authPromise;
}

async function restoreAccountFromCache() {
  if (currentAccount) return currentAccount;
  try {
    const accounts = await pca.getTokenCache().getAllAccounts();
    if (accounts && accounts.length > 0) {
      currentAccount = accounts[0];
      setAuthState({
        status: "authenticated",
        account: currentAccount.username || null,
        error: null,
      });
      log("Restored account from persistent cache", currentAccount.username || "");
    }
  } catch (err) {
    log("restoreAccountFromCache failed", err?.message || String(err));
  }
  return currentAccount;
}

async function getValidAccessToken() {
  ensureConfigured();

  const cached = getCachedToken();
  if (cached) return cached;

  if (!currentAccount) {
    await restoreAccountFromCache();
  }

  if (currentAccount) {
    try {
      const silent = await pca.acquireTokenSilent({
        account: currentAccount,
        scopes: SCOPES,
      });

      if (silent?.accessToken) {
        cachedAccessToken = silent.accessToken;
        cachedTokenExpiresAt = silent.expiresOn
          ? silent.expiresOn.getTime()
          : Date.now() + 50 * 60 * 1000;

        setAuthState({
          status: "authenticated",
          account: currentAccount?.username || null,
          error: null,
        });

        return silent.accessToken;
      }
    } catch (err) {
      log("acquireTokenSilent failed", err?.message || String(err));
    }
  }

  if (authState.status === "pending") {
    throw new Error(
      "Microsoft 365 login is not yet complete. Finish the browser login first, then call auth_status."
    );
  }

  if (authState.status === "error") {
    throw new Error(
      `Microsoft 365 login is currently in error state: ${authState.error || "unknown error"}. Please call begin_auth again.`
    );
  }

  throw new Error("Microsoft 365 is not logged in yet. Please call begin_auth first.");
}

async function graphGetJson(url, extraHeaders = {}) {
  const token = await getValidAccessToken();

  const res = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      ...extraHeaders,
    },
  });

  if (!res.ok) {
    const text = await res.text();
    log("graphGetJson failed", { url, status: res.status, text });
    throw new Error(`Graph API ${res.status}: ${text}`);
  }

  return res.json();
}

async function graphGetBytes(url) {
  const token = await getValidAccessToken();

  const res = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });

  if (!res.ok) {
    const text = await res.text();
    log("graphGetBytes failed", { url, status: res.status, text });
    throw new Error(`Graph API ${res.status}: ${text}`);
  }

  const arrayBuffer = await res.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

async function graphSend(method, url, bodyObj) {
  const token = await getValidAccessToken();

  const res = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      Accept: "application/json",
    },
    body: bodyObj ? JSON.stringify(bodyObj) : undefined,
  });

  if (!res.ok) {
    const text = await res.text();
    log("graphSend failed", { method, url, status: res.status, text });
    throw new Error(`Graph API ${res.status}: ${text}`);
  }

  // sendMail returns 202 with no body; createMessage returns 201 with JSON.
  const text = await res.text();
  return text ? JSON.parse(text) : null;
}

async function graphDelete(url) {
  const token = await getValidAccessToken();

  const res = await fetch(url, {
    method: "DELETE",
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });

  if (!res.ok) {
    const text = await res.text();
    log("graphDelete failed", { url, status: res.status, text });
    throw new Error(`Graph API ${res.status}: ${text}`);
  }

  return null;
}

function mailboxPrefix(mailbox = "me") {
  return mailbox === "me"
    ? `${GRAPH_BASE}/me`
    : `${GRAPH_BASE}/users/${encodeURIComponent(mailbox)}`;
}

function messageCollectionUrl(mailbox = "me", folder = "inbox") {
  const base = mailboxPrefix(mailbox);
  const normalizedFolder = String(folder || "inbox").trim().toLowerCase();

  if (!normalizedFolder || normalizedFolder === "inbox") {
    return `${base}/mailFolders/inbox/messages`;
  }

  if (normalizedFolder === "all" || normalizedFolder === "messages") {
    return `${base}/messages`;
  }

  return `${base}/mailFolders/${encodeURIComponent(folder)}/messages`;
}

function normalizeNeedle(value) {
  return normalizeExtractedText(value).toLowerCase();
}

function matchesNeedle(value, needle) {
  if (!needle) return true;
  return String(value || "").toLowerCase().includes(needle);
}

function formatSender(fromName, fromAddress) {
  if (fromName && fromAddress) {
    return `${fromName} <${fromAddress}>`;
  }

  return fromName || fromAddress || "(unknown)";
}

const MESSAGE_ID_DESCRIPTION =
  "Outlook messageId copied from list_recent_messages, search_messages, list_flagged_messages, or read_email. Use this same value for read_email and reply_outlook_email.";

const MESSAGE_CHAIN_FOOTER =
  "Next steps: pass messageId to read_email to read the body, then to reply_outlook_email to reply in the same Outlook thread. Do not use send_outlook_email for replies — that starts a new conversation.";

const MESSAGE_CHAIN_NEXT_TOOLS = {
  read: "read_email",
  replyInThread: "reply_outlook_email",
  passField: "messageId",
};

const DRAFT_ID_DESCRIPTION =
  "Outlook draftId from send_outlook_email, reply_outlook_email, or list_outlook_drafts. Only unsent drafts can be edited or deleted.";

const DRAFT_NEXT_TOOLS = {
  read: "read_email",
  edit: "edit_outlook_draft",
  delete: "delete_outlook_draft",
  passField: "draftId",
};

const DRAFT_CHAIN_FOOTER =
  "Next steps: pass draftId to read_email to review, edit_outlook_draft to change it, or delete_outlook_draft to discard it. Nothing is sent until you send it from Outlook or call send/reply with send_now=true.";

function recipientAddresses(recipients) {
  return (recipients || [])
    .map((item) => item?.emailAddress?.address)
    .filter(Boolean);
}

function mapListedMessage(message) {
  const id = message.id;
  return {
    id,
    messageId: id,
    subject: message.subject,
    receivedDateTime: message.receivedDateTime,
    hasAttachments: message.hasAttachments,
    fromName: message.from?.emailAddress?.name || "",
    fromAddress: message.from?.emailAddress?.address || "",
    to: recipientAddresses(message.toRecipients),
    cc: recipientAddresses(message.ccRecipients),
    bodyPreview: message.bodyPreview || "",
    isRead: message.isRead,
    conversationId: message.conversationId || null,
  };
}

function formatListedMessageText(message, index) {
  const lines = [
    `${index + 1}. ${message.subject || "(no subject)"}`,
    `From: ${formatSender(message.fromName, message.fromAddress)}`,
  ];
  if (Array.isArray(message.to) && message.to.length) {
    lines.push(`To: ${message.to.join(", ")}`);
  }
  lines.push(`Received: ${message.receivedDateTime}`);
  lines.push(`Has Attachments: ${message.hasAttachments}`);
  if (message.bodyPreview) {
    lines.push(`Preview: ${String(message.bodyPreview).slice(0, 200)}`);
  }
  lines.push(`messageId: ${message.messageId || message.id}`);
  return lines.join("\n");
}

function formatMessageListText(messages, emptyText) {
  if (!messages.length) return emptyText;
  return `${messages
    .map((message, index) => formatListedMessageText(message, index))
    .join("\n\n")}\n\n${MESSAGE_CHAIN_FOOTER}`;
}

function safeFilename(name) {
  return (name || "attachment.bin").replace(/[<>:"/\\|?*\x00-\x1F]/g, "_");
}

function registerAliases(server, name, config, handler, aliases = []) {
  server.registerTool(name, config, handler);
  for (const alias of aliases) {
    server.registerTool(alias, config, handler);
  }
}

function createParseContext() {
  return {
    imageBlocks: [],
    maxImageBlocks: MAX_IMAGE_BLOCKS,
    imageBytesUsed: 0,
    maxTotalImageBytes: MAX_TOTAL_IMAGE_BYTES,
    imageNotes: [],
  };
}

function bufferToArrayBuffer(buffer) {
  return buffer.buffer.slice(
    buffer.byteOffset,
    buffer.byteOffset + buffer.byteLength
  );
}

function getExtension(filename = "") {
  const ext = path.extname(filename || "").toLowerCase();
  return ext.startsWith(".") ? ext.slice(1) : ext;
}

function normalizeExtractedText(text) {
  return String(text || "")
    .replace(/\u0000/g, "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function normalizeEmailBodyText(text) {
  return String(text || "")
    .replace(/\u0000/g, "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .trim();
}

function decodeXmlEntities(text) {
  return String(text || "")
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) =>
      String.fromCodePoint(parseInt(hex, 16))
    )
    .replace(/&#(\d+);/g, (_, dec) =>
      String.fromCodePoint(parseInt(dec, 10))
    )
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function extractTextFromXml(xml) {
  const withBreaks = String(xml || "")
    .replace(/<a:br\s*\/>/gi, "\n")
    .replace(/<w:br\s*\/>/gi, "\n")
    .replace(/<a:tab\s*\/>/gi, "\t")
    .replace(/<w:tab\s*\/>/gi, "\t")
    .replace(/<\/(a:p|w:p|p:txBody|text:p|p:notesTextViewPr)>/gi, "\n");

  const matches = [...withBreaks.matchAll(/>([^<>]+)</g)]
    .map((match) => decodeXmlEntities(match[1]))
    .map((value) => value.replace(/[ \t]+/g, " ").trim())
    .filter(Boolean);

  return normalizeExtractedText(matches.join("\n"));
}

function truncateExtractedText(text, maxLength = MAX_RETURN_CHARS) {
  const totalLength = String(text || "").length;

  if (totalLength <= maxLength) {
    return {
      text,
      totalLength,
      returnedLength: totalLength,
      omittedLength: 0,
      truncated: false,
    };
  }

  return {
    text: `${String(text || "").slice(0, maxLength)}\n\n... (parsed output truncated)`,
    totalLength,
    returnedLength: maxLength,
    omittedLength: totalLength - maxLength,
    truncated: true,
  };
}

function toNumberAwareSort(entries, pattern) {
  return [...entries].sort((a, b) => {
    const aNum = Number(a.entryName.match(pattern)?.[1] || 0);
    const bNum = Number(b.entryName.match(pattern)?.[1] || 0);
    return aNum - bNum;
  });
}

function addContextNote(context, note) {
  if (!note) return;
  if (!context.imageNotes.includes(note)) {
    context.imageNotes.push(note);
  }
}

async function loadSharp() {
  if (!sharpPromise) {
    sharpPromise = import("sharp").then((module) => {
      const sharp = module.default || module.sharp || module;
      if (typeof sharp !== "function") {
        throw new Error("sharp loaded but did not expose a usable default export.");
      }
      return sharp;
    });
  }

  return sharpPromise;
}

async function optimizeImageForBlock(buffer, mimeType) {
  const original = Buffer.from(buffer);

  if (original.length <= MAX_IMAGE_BLOCK_BYTES) {
    return {
      ok: true,
      buffer: original,
      mimeType,
      originalBytes: original.length,
      finalBytes: original.length,
      transformed: false,
    };
  }

  try {
    const sharp = await loadSharp();
    const metadata = await sharp(original, {
      animated: false,
      limitInputPixels: false,
    })
      .rotate()
      .metadata();

    const hasAlpha = !!metadata.hasAlpha;
    const outputMimeType = hasAlpha ? "image/webp" : "image/jpeg";

    for (const maxDimension of IMAGE_RESIZE_STEPS) {
      let pipeline = sharp(original, {
        animated: false,
        limitInputPixels: false,
      }).rotate();

      pipeline = pipeline.resize({
        width: maxDimension,
        height: maxDimension,
        fit: "inside",
        withoutEnlargement: true,
      });

      const compressed = hasAlpha
        ? await pipeline.webp({ quality: 76, effort: 4, alphaQuality: 80 }).toBuffer()
        : await pipeline.jpeg({ quality: 76, mozjpeg: true }).toBuffer();

      if (compressed.length <= MAX_IMAGE_BLOCK_BYTES) {
        return {
          ok: true,
          buffer: compressed,
          mimeType: outputMimeType,
          originalBytes: original.length,
          finalBytes: compressed.length,
          transformed: true,
        };
      }
    }

    const fallback = await sharp(original, {
      animated: false,
      limitInputPixels: false,
    })
      .rotate()
      .resize({
        width: 320,
        height: 320,
        fit: "inside",
        withoutEnlargement: true,
      })
      .webp({ quality: 60, effort: 4 })
      .toBuffer();

    if (fallback.length <= MAX_IMAGE_BLOCK_BYTES) {
      return {
        ok: true,
        buffer: fallback,
        mimeType: "image/webp",
        originalBytes: original.length,
        finalBytes: fallback.length,
        transformed: true,
      };
    }

    return {
      ok: false,
      reason: "single_image_too_large",
      originalBytes: original.length,
      finalBytes: fallback.length,
    };
  } catch (err) {
    log("Image optimization error", {
      mimeType,
      bytes: original.length,
      error: err?.message || String(err),
    });

    return {
      ok: false,
      reason: "image_optimization_failed",
      originalBytes: original.length,
      finalBytes: original.length,
    };
  }
}

async function addImageBlock(context, buffer, mimeType) {
  if (!mimeType || !buffer) {
    return { attached: false, reason: "invalid_image_payload" };
  }

  if (context.imageBlocks.length >= context.maxImageBlocks) {
    addContextNote(
      context,
      `Some images were skipped because the maximum of ${context.maxImageBlocks} visual blocks was reached.`
    );
    return { attached: false, reason: "image_block_count_limit" };
  }

  const optimized = await optimizeImageForBlock(buffer, mimeType);
  if (!optimized.ok) {
    if (optimized.reason === "single_image_too_large") {
      addContextNote(
        context,
        `Some large images could not be attached because even after downscaling they exceeded the per-image MCP payload budget of ${Math.round(MAX_IMAGE_BLOCK_BYTES / 1024)} KB.`
      );
    } else {
      addContextNote(
        context,
        "Some images could not be optimized for inline visual preview and were skipped."
      );
    }

    return {
      attached: false,
      reason: optimized.reason,
      originalBytes: optimized.originalBytes,
      finalBytes: optimized.finalBytes,
    };
  }

  if (context.imageBytesUsed + optimized.finalBytes > context.maxTotalImageBytes) {
    addContextNote(
      context,
      `Some images were skipped because the total inline image budget of ${Math.round(context.maxTotalImageBytes / 1024)} KB was reached.`
    );
    return {
      attached: false,
      reason: "total_image_budget_reached",
      originalBytes: optimized.originalBytes,
      finalBytes: optimized.finalBytes,
    };
  }

  context.imageBlocks.push({
    type: "image",
    data: optimized.buffer.toString("base64"),
    mimeType: optimized.mimeType,
  });
  context.imageBytesUsed += optimized.finalBytes;

  if (optimized.transformed) {
    addContextNote(
      context,
      "Large images were downscaled and recompressed before being attached so they stay within Claude Desktop's MCP tool-result payload budget."
    );
  }

  return {
    attached: true,
    reason: optimized.transformed ? "compressed_for_delivery" : "attached_original",
    mimeType: optimized.mimeType,
    originalBytes: optimized.originalBytes,
    finalBytes: optimized.finalBytes,
    transformed: optimized.transformed,
  };
}

function inferImageMimeType(ext, contentType = "") {
  const normalizedType = String(contentType || "").toLowerCase();
  if (normalizedType.startsWith("image/")) return normalizedType;
  return IMAGE_MIME_BY_EXT[String(ext || "").toLowerCase()] || null;
}

function isImageAttachment(ext, contentType = "") {
  return !!inferImageMimeType(ext, contentType);
}

function isProbablyTextLike(ext, contentType = "", kind = "") {
  return (
    String(contentType || "").toLowerCase().startsWith("text/") ||
    ["txt", "md", "json", "xml", "html", "csv"].includes(ext) ||
    kind === "itemAttachment"
  );
}

function isPowerPointZipExt(ext) {
  return ["pptx", "pptm", "ppsx", "potx"].includes(ext);
}

function isIgnoredArchiveEntry(name) {
  const normalized = String(name || "").replace(/\\/g, "/");
  const base = path.posix.basename(normalized);

  return (
    !base ||
    base.startsWith(".") ||
    normalized.startsWith("__MACOSX/") ||
    normalized.endsWith("/.DS_Store")
  );
}

async function collectZipImages(zip, mediaPrefix, context, maxImages) {
  let added = 0;

  for (const entry of zip.getEntries()) {
    if (entry.isDirectory || !entry.entryName.startsWith(mediaPrefix)) continue;
    if (added >= maxImages || context.imageBlocks.length >= context.maxImageBlocks) {
      break;
    }

    const ext = getExtension(entry.entryName);
    const mimeType = inferImageMimeType(ext);
    if (!mimeType) continue;

    const result = await addImageBlock(context, entry.getData(), mimeType);
    if (result.attached) {
      added += 1;
    }
  }

  return added;
}

function formatMetadataBlock(lines) {
  const visible = lines.filter(Boolean);
  return visible.length ? visible.join("\n") : "";
}

function appendSection(parts, title, body) {
  const normalized = normalizeExtractedText(body);
  if (!normalized) return;
  parts.push(`### ${title}\n${normalized}`);
}

function sanitizeSnippet(text) {
  return normalizeExtractedText(String(text || "").replace(/[^\S\n]+/g, " "));
}

function csvEscapeCell(value) {
  const normalized =
    value === null || value === undefined
      ? ""
      : value instanceof Date
        ? value.toISOString()
        : String(value);

  if (/[",\n\r]/.test(normalized)) {
    return `"${normalized.replace(/"/g, "\"\"")}"`;
  }

  return normalized;
}

function worksheetToCsv(worksheet) {
  let maxColumn = 0;

  worksheet.eachRow({ includeEmpty: true }, (row) => {
    const values = Array.isArray(row.values) ? row.values.slice(1) : [];
    maxColumn = Math.max(maxColumn, values.length);
  });

  const lines = [];
  worksheet.eachRow({ includeEmpty: true }, (row) => {
    const values = Array.isArray(row.values) ? row.values.slice(1) : [];
    const cells = [];

    for (let index = 0; index < maxColumn; index += 1) {
      const cell = row.getCell(index + 1);
      let value = cell?.text;

      if (!value && value !== "0") {
        value = cell?.result ?? cell?.value ?? "";
      }

      cells.push(csvEscapeCell(value));
    }

    lines.push(cells.join(","));
  });

  return lines.join("\n");
}

function isMeaningfulSnippet(text) {
  const normalized = sanitizeSnippet(text);
  if (normalized.length < 4) return false;

  const signalChars = (normalized.match(/[A-Za-z0-9\u00C0-\uFFFF]/g) || []).length;
  return signalChars >= Math.max(4, Math.floor(normalized.length * 0.3));
}

function extractPrintableStrings(buffer) {
  const results = [];
  const seen = new Set();

  const utf16Matches =
    buffer
      .toString("utf16le")
      .match(/[\p{L}\p{N}][\p{L}\p{N}\p{P}\p{Zs}\t\n\r]{3,}/gu) || [];
  const asciiMatches =
    buffer.toString("latin1").match(/[A-Za-z0-9][ -~\t\n\r]{3,}/g) || [];

  for (const candidate of [...utf16Matches, ...asciiMatches]) {
    const snippet = sanitizeSnippet(candidate);
    if (!isMeaningfulSnippet(snippet) || seen.has(snippet)) continue;
    seen.add(snippet);
    results.push(snippet);
  }

  return normalizeExtractedText(results.join("\n"));
}

function extractLegacyPptText(buffer) {
  try {
    const cfb = CFB.parse(buffer, { type: "buffer" });
    const chunks = [];

    for (const entry of cfb.FileIndex || []) {
      if (entry.type !== 2 || !entry.content) continue;
      const entryName = String(entry.name || "");
      if (
        /PowerPoint Document|Current User|SummaryInformation|DocumentSummaryInformation/i.test(
          entryName
        )
      ) {
        chunks.push(extractPrintableStrings(Buffer.from(entry.content)));
      }
    }

    const combined = normalizeExtractedText(chunks.join("\n"));
    if (combined) return combined;
  } catch (err) {
    log("Legacy PPT parse fallback", err?.message || String(err));
  }

  return extractPrintableStrings(buffer);
}

async function loadPdfOcrDeps() {
  if (!pdfOcrDepsPromise) {
    pdfOcrDepsPromise = Promise.all([
      import("pdf-to-img"),
      import("tesseract.js"),
    ]).then(([pdfToImgModule, tesseractModule]) => {
      const pdf =
        pdfToImgModule.pdf ||
        pdfToImgModule.default?.pdf ||
        pdfToImgModule.default;
      const createWorker =
        tesseractModule.createWorker ||
        tesseractModule.default?.createWorker;

      if (typeof pdf !== "function" || typeof createWorker !== "function") {
        throw new Error("OCR dependencies loaded but did not expose the expected API.");
      }

      return { pdf, createWorker };
    });
  }

  return pdfOcrDepsPromise;
}

async function loadMsgReader() {
  if (!msgReaderPromise) {
    msgReaderPromise = import("@kenjiuno/msgreader").then((module) => {
      const MsgReader =
        module.default?.default ||
        module.default ||
        module.MsgReader ||
        module;

      if (typeof MsgReader !== "function") {
        throw new Error(".msg parser dependency loaded but MsgReader export is missing.");
      }

      return MsgReader;
    });
  }

  return msgReaderPromise;
}

async function loadRarDeps() {
  if (!rarDepsPromise) {
    rarDepsPromise = import("node-unrar-js").then((module) => {
      const createExtractorFromData =
        module.createExtractorFromData ||
        module.default?.createExtractorFromData;

      if (typeof createExtractorFromData !== "function") {
        throw new Error("RAR dependency loaded but createExtractorFromData export is missing.");
      }

      return { createExtractorFromData };
    });
  }

  return rarDepsPromise;
}

async function loadSevenZipDeps() {
  if (!sevenZipDepsPromise) {
    sevenZipDepsPromise = Promise.all([
      import("node-7z"),
      import("7zip-bin"),
    ]).then(([sevenModule, binModule]) => {
      const extractFull =
        sevenModule.extractFull || sevenModule.default?.extractFull;
      const path7za = binModule.path7za || binModule.default?.path7za;

      if (typeof extractFull !== "function" || !path7za) {
        throw new Error("7z dependencies loaded but did not expose the expected API.");
      }

      return { extractFull, path7za };
    });
  }

  return sevenZipDepsPromise;
}

async function runPdfOcr(raw, context) {
  const { pdf, createWorker } = await loadPdfOcrDeps();
  const document = await pdf(
    `data:application/pdf;base64,${raw.toString("base64")}`,
    { scale: 2 }
  );

  const worker = await createWorker("eng");
  const pageTexts = [];
  let pageCount = 0;
  let attachedPages = 0;

  try {
    for await (const pageImage of document) {
      pageCount += 1;
      if (pageCount > MAX_PDF_OCR_PAGES) break;

      const imageBuffer = Buffer.from(pageImage);

      if (attachedPages < MAX_PDF_VISUAL_PAGES) {
        const attachmentResult = await addImageBlock(context, imageBuffer, "image/png");
        if (attachmentResult.attached) {
          attachedPages += 1;
        }
      }

      const result = await worker.recognize(imageBuffer);
      const pageText = normalizeExtractedText(result?.data?.text || "");
      if (pageText) {
        pageTexts.push(`### OCR Page ${pageCount}\n${pageText}`);
      }
    }
  } finally {
    await worker.terminate().catch(() => {});
  }

  return {
    text: normalizeExtractedText(pageTexts.join("\n\n")),
    pageCount: Math.min(pageCount, MAX_PDF_OCR_PAGES),
    attachedPages,
  };
}

async function parsePptx(raw, context) {
  const zip = new AdmZip(raw);
  const slideEntries = toNumberAwareSort(
    zip.getEntries().filter((entry) => /^ppt\/slides\/slide\d+\.xml$/i.test(entry.entryName)),
    /slide(\d+)\.xml/i
  );

  const noteEntries = new Map(
    toNumberAwareSort(
      zip.getEntries().filter((entry) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(entry.entryName)),
      /notesSlide(\d+)\.xml/i
    ).map((entry) => [
      Number(entry.entryName.match(/notesSlide(\d+)\.xml/i)?.[1] || 0),
      extractTextFromXml(entry.getData().toString("utf8")),
    ])
  );

  const parts = [];

  for (const entry of slideEntries) {
    const slideNumber = Number(entry.entryName.match(/slide(\d+)\.xml/i)?.[1] || 0);
    const slideText = extractTextFromXml(entry.getData().toString("utf8"));
    appendSection(parts, `Slide ${slideNumber || parts.length + 1}`, slideText);

    const noteText = noteEntries.get(slideNumber);
    if (noteText) {
      appendSection(parts, `Slide ${slideNumber || parts.length + 1} Notes`, noteText);
    }
  }

  const addedImages = await collectZipImages(zip, "ppt/media/", context, MAX_DOC_IMAGES);
  let extractedText = normalizeExtractedText(parts.join("\n\n"));

  if (!extractedText) {
    extractedText = "(No readable slide text was detected in this presentation.)";
  }

  if (addedImages > 0) {
    extractedText += `\n\n[System: MCP extracted ${addedImages} embedded presentation image(s) and sent them as visual blocks.]`;
  }

  return extractedText;
}

async function parseMsg(raw, filename, context, depth) {
  const MsgReader = await loadMsgReader();
  const msg = new MsgReader(raw);
  const info = msg.getFileData();
  const parts = [];

  appendSection(
    parts,
    "Message Summary",
    formatMetadataBlock([
      info?.subject ? `Subject: ${info.subject}` : "",
      info?.senderName ? `Sender Name: ${info.senderName}` : "",
      info?.senderEmail ? `Sender Email: ${info.senderEmail}` : "",
      info?.messageDeliveryTime ? `Delivered: ${info.messageDeliveryTime}` : "",
      Array.isArray(info?.recipients) && info.recipients.length
        ? `Recipients: ${info.recipients
            .map((item) => item.email || item.name)
            .filter(Boolean)
            .join(", ")}`
        : "",
    ])
  );

  const body = normalizeExtractedText(info?.body || info?.bodyHTML || "");
  appendSection(parts, "Message Body", body || "(No readable message body was detected.)");

  const attachments = Array.isArray(info?.attachments) ? info.attachments : [];
  if (attachments.length) {
    appendSection(
      parts,
      "Embedded Attachments",
      attachments
        .slice(0, MAX_MSG_ATTACHMENTS)
        .map(
          (attachment, index) =>
            `${index + 1}. ${attachment.fileName || attachment.fileNameShort || "unnamed attachment"}`
        )
        .join("\n")
    );

    for (const attachment of attachments.slice(0, MAX_MSG_ATTACHMENTS)) {
      const nested = msg.getAttachment(attachment);
      const nestedName =
        nested?.fileName ||
        attachment.fileName ||
        attachment.fileNameShort ||
        "embedded-attachment.bin";
      const nestedRaw = Buffer.from(nested?.content || []);

      if (!nestedRaw.length) {
        appendSection(parts, `Embedded Attachment: ${nestedName}`, "(Attachment payload was empty.)");
        continue;
      }

      if (depth >= MAX_NESTED_DEPTH) {
        appendSection(
          parts,
          `Embedded Attachment: ${nestedName}`,
          "(Nested attachment parsing skipped because the recursion limit was reached.)"
        );
        continue;
      }

      const nestedParsed = await parseAttachmentPayload(
        {
          raw: nestedRaw,
          filename: nestedName,
          contentType: attachment.mimeType || "",
          kind: "attachment",
        },
        context,
        depth + 1
      );

      appendSection(parts, `Embedded Attachment: ${nestedName}`, nestedParsed.text);
    }

    if (attachments.length > MAX_MSG_ATTACHMENTS) {
      appendSection(
        parts,
        "Embedded Attachments Limit",
        `Only the first ${MAX_MSG_ATTACHMENTS} embedded attachments were parsed out of ${attachments.length}.`
      );
    }
  }

  const text = normalizeExtractedText(parts.join("\n\n"));
  return {
    parser: "outlook_msg",
    text:
      text ||
      `[Outlook .msg attachment]\n\nFilename: ${filename}\n\n(No readable message content was extracted.)`,
    details: {
      embeddedAttachmentCount: attachments.length,
    },
  };
}

async function parseArchiveEntries(entries, archiveKind, context, depth) {
  const visibleEntries = entries.filter(
    (entry) => entry && entry.name && !isIgnoredArchiveEntry(entry.name)
  );
  const selectedEntries = visibleEntries.slice(0, MAX_ARCHIVE_ITEMS);
  const parts = [];

  appendSection(
    parts,
    `${archiveKind.toUpperCase()} Archive Summary`,
    formatMetadataBlock([
      `Files in archive: ${visibleEntries.length}`,
      visibleEntries.length > MAX_ARCHIVE_ITEMS
        ? `Parsed entries: ${MAX_ARCHIVE_ITEMS} (remaining entries were summarized only)`
        : `Parsed entries: ${selectedEntries.length}`,
    ])
  );

  for (const entry of selectedEntries) {
    const entrySize = entry.buffer?.length ?? entry.size ?? 0;

    if (!entry.buffer) {
      appendSection(parts, entry.name, "(Archive entry payload was unavailable.)");
      continue;
    }

    if (entrySize > MAX_ARCHIVE_ENTRY_BYTES) {
      appendSection(
        parts,
        entry.name,
        `Skipped because the extracted entry is too large (${entrySize} bytes).`
      );
      continue;
    }

    if (depth >= MAX_NESTED_DEPTH) {
      appendSection(
        parts,
        entry.name,
        "(Nested archive parsing skipped because the recursion limit was reached.)"
      );
      continue;
    }

    const parsed = await parseAttachmentPayload(
      {
        raw: entry.buffer,
        filename: entry.name,
        contentType: entry.contentType || "",
        kind: "attachment",
      },
      context,
      depth + 1
    );

    appendSection(parts, entry.name, parsed.text);
  }

  if (visibleEntries.length > MAX_ARCHIVE_ITEMS) {
    appendSection(
      parts,
      `${archiveKind.toUpperCase()} Archive Overflow`,
      `Only the first ${MAX_ARCHIVE_ITEMS} archive entries were parsed out of ${visibleEntries.length}.`
    );
  }

  return {
    parser: `${archiveKind}_archive`,
    text:
      normalizeExtractedText(parts.join("\n\n")) ||
      `(${archiveKind.toUpperCase()} archive had no readable entries.)`,
    details: {
      archiveEntryCount: visibleEntries.length,
      parsedArchiveEntryCount: selectedEntries.length,
    },
  };
}

async function parseZip(raw, context, depth) {
  const zip = new AdmZip(raw);
  const entries = zip
    .getEntries()
    .filter((entry) => !entry.isDirectory)
    .map((entry) => ({
      name: entry.entryName.replace(/\\/g, "/"),
      buffer: entry.getData(),
      size: entry.header?.size || entry.getData().length,
    }));

  return parseArchiveEntries(entries, "zip", context, depth);
}

async function parseRar(raw, context, depth) {
  const { createExtractorFromData } = await loadRarDeps();
  const extractor = await createExtractorFromData({ data: bufferToArrayBuffer(raw) });
  const list = extractor.getFileList();
  const fileHeaders = [...list.fileHeaders];
  const targetFiles = fileHeaders
    .map((header) => header?.name)
    .filter(Boolean)
    .slice(0, MAX_ARCHIVE_ITEMS);
  const extracted = extractor.extract({ files: targetFiles });
  const entries = [...extracted.files].map((file) => ({
    name: file?.fileHeader?.name || "unnamed-file",
    buffer: file?.extraction ? Buffer.from(file.extraction) : null,
    size: file?.fileHeader?.unpSize || file?.extraction?.length || 0,
  }));

  return parseArchiveEntries(entries, "rar", context, depth);
}

function listFilesRecursive(dir) {
  const out = [];

  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      out.push(...listFilesRecursive(fullPath));
    } else if (entry.isFile()) {
      out.push(fullPath);
    }
  }

  return out;
}

async function parse7z(raw, filename, context, depth) {
  const { extractFull, path7za } = await loadSevenZipDeps();
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "mcp-local-7z-"));
  const archivePath = path.join(tempDir, safeFilename(filename || "archive.7z"));
  const extractDir = path.join(tempDir, "out");

  fs.mkdirSync(extractDir, { recursive: true });
  fs.writeFileSync(archivePath, raw);

  try {
    await new Promise((resolve, reject) => {
      const stream = extractFull(archivePath, extractDir, {
        $bin: path7za,
        $progress: false,
        overwrite: "a",
      });

      stream.on("error", reject);
      stream.on("end", resolve);
    });

    const entries = listFilesRecursive(extractDir).map((fullPath) => ({
      name: path.relative(extractDir, fullPath).replace(/\\/g, "/"),
      buffer: fs.readFileSync(fullPath),
      size: fs.statSync(fullPath).size,
    }));

    return parseArchiveEntries(entries, "7z", context, depth);
  } finally {
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
}

async function parseAttachmentPayload(payload, context, depth = 0) {
  const { raw, filename, contentType = "", kind = "attachment" } = payload;
  const ext = getExtension(filename);

  if (isImageAttachment(ext, contentType)) {
    const mimeType = inferImageMimeType(ext, contentType);
    const attached = await addImageBlock(context, raw, mimeType);

    return {
      parser: "image_passthrough",
      text: normalizeExtractedText(
        [
          `[Image attachment] ${filename}`,
          attached.attached
            ? attached.transformed
              ? `The local backend downscaled this image from ${Math.round((attached.originalBytes || 0) / 1024)} KB to ${Math.round((attached.finalBytes || 0) / 1024)} KB and sent it as a visual block so the model can inspect it directly.`
              : "The local backend sent this image as a visual block so the model can inspect it directly."
            : attached.reason === "single_image_too_large"
              ? "The backend detected an image, but even after downscaling it remained too large for Claude Desktop's MCP tool-result payload budget."
              : attached.reason === "total_image_budget_reached"
                ? "The backend detected an image, but the total inline image budget for this tool result was already reached."
                : "The backend detected an image, but it could not be attached inline.",
        ].join("\n\n")
      ),
      details: {
        attachedImage: attached.attached,
        imageResult: attached,
      },
    };
  }

  if (ext === "pdf") {
    let extractedText = normalizeExtractedText((await parsePdf(raw))?.text || "");
    const notes = [];
    let ocrUsed = false;
    let ocrPageCount = 0;

    if (!extractedText) {
      notes.push(
        "No embedded PDF text layer was detected. The backend treated this as a scanned/image PDF and attempted OCR."
      );

      try {
        const ocr = await runPdfOcr(raw, context);
        ocrPageCount = ocr.pageCount;
        if (ocr.text) {
          extractedText = ocr.text;
          ocrUsed = true;
          notes.push(`OCR extracted readable text from ${ocr.pageCount} page(s).`);
        } else {
          notes.push("OCR completed but did not recover readable text from the rendered pages.");
        }
      } catch (ocrErr) {
        log("PDF OCR error", ocrErr);
        notes.push(`OCR could not run successfully: ${ocrErr?.message || String(ocrErr)}`);
      }
    }

    if (!extractedText) {
      extractedText =
        "(This PDF appears to be image-based or scanned. No readable text was extracted.)";
    }

    return {
      parser: ocrUsed ? "pdf_ocr" : "pdf_text",
      text: normalizeExtractedText([...notes, extractedText].join("\n\n")),
      details: {
        ocrUsed,
        ocrPageCount,
      },
    };
  }

  if (ext === "docx" || ext === "doc") {
    const result = await mammoth.extractRawText({ buffer: raw });
    let extractedText =
      normalizeExtractedText(result.value) ||
      "(No readable plain text detected in this document.)";

    if (ext === "docx") {
      try {
        const zip = new AdmZip(raw);
        const imageCount = await collectZipImages(zip, "word/media/", context, MAX_DOC_IMAGES);
        if (imageCount > 0) {
          extractedText += `\n\n[System: MCP extracted ${imageCount} embedded document image(s) and sent them as visual blocks.]`;
        }
      } catch (zipErr) {
        log("DOCX image extraction error", zipErr);
      }
    }

    return {
      parser: ext === "docx" ? "docx" : "doc",
      text: extractedText,
      details: {},
    };
  }

  if (isPowerPointZipExt(ext)) {
    return {
      parser: "pptx",
      text: await parsePptx(raw, context),
      details: {},
    };
  }

  if (ext === "ppt") {
    const extractedText = extractLegacyPptText(raw);
    return {
      parser: "ppt_legacy",
      text: normalizeExtractedText(
        [
          "Legacy .ppt parsing uses best-effort binary text extraction.",
          extractedText || "(No readable text was extracted from this legacy PowerPoint file.)",
        ].join("\n\n")
      ),
      details: {},
    };
  }

  if (ext === "csv") {
    return {
      parser: "csv",
      text: raw.toString("utf8"),
      details: {},
    };
  }

  if (ext === "xlsx" || ext === "xlsm") {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(raw);
    const sheetTexts = [];

    for (const worksheet of workbook.worksheets) {
      const csv = worksheetToCsv(worksheet);
      sheetTexts.push(`### Sheet: ${worksheet.name}\n\`\`\`csv\n${csv}\n\`\`\``);
    }

    return {
      parser: ext,
      text: normalizeExtractedText(sheetTexts.join("\n\n")),
      details: {
        sheetCount: workbook.worksheets.length,
      },
    };
  }

  if (ext === "xls") {
    const extractedText = extractPrintableStrings(raw);

    return {
      parser: "xls_legacy",
      text: normalizeExtractedText(
        [
          "Legacy .xls parsing uses best-effort binary text extraction.",
          extractedText || "(No readable text was extracted from this legacy Excel file.)",
        ].join("\n\n")
      ),
      details: {},
    };
  }

  if (ext === "msg" || String(contentType).toLowerCase() === "application/vnd.ms-outlook") {
    return parseMsg(raw, filename, context, depth);
  }

  if (ext === "zip") {
    return parseZip(raw, context, depth);
  }

  if (ext === "rar") {
    return parseRar(raw, context, depth);
  }

  if (ext === "7z") {
    return parse7z(raw, filename, context, depth);
  }

  if (isProbablyTextLike(ext, contentType, kind)) {
    return {
      parser: "text",
      text: raw.toString("utf8"),
      details: {},
    };
  }

  return {
    parser: "unsupported",
    text: `(Note: This file format is not supported for direct parsing.)\nFilename: ${filename}\nType: ${contentType || "application/octet-stream"}`,
    details: {},
  };
}

function htmlToPlainText(html) {
  return normalizeExtractedText(
    decodeXmlEntities(
      String(html || "")
        .replace(/<(script|style)[^>]*>[\s\S]*?<\/\1>/gi, "")
        .replace(/<\/(p|div|li|tr|h[1-6]|blockquote)>/gi, "\n")
        .replace(/<br\s*\/?>/gi, "\n")
        .replace(/<li[^>]*>/gi, "- ")
        .replace(/<[^>]+>/g, "")
    )
  );
}

function normalizeEmailList(value) {
  if (value === null || value === undefined) return [];
  const list = Array.isArray(value) ? value : String(value).split(/[;,]/);
  return list.map((item) => String(item).trim()).filter(Boolean);
}

function buildRecipients(value) {
  return normalizeEmailList(value).map((address) => ({
    emailAddress: { address },
  }));
}

function getMimeForExt(ext) {
  return MIME_BY_EXT[String(ext || "").toLowerCase()] || "application/octet-stream";
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function loadLocalAttachment(filePath) {
  if (typeof filePath !== "string" || !filePath.trim()) {
    throw new Error("Each attachment must be a non-empty file path string.");
  }
  if (filePath.includes("\u0000")) {
    throw new Error("Attachment path contains an invalid NUL byte.");
  }

  const resolved = path.resolve(filePath.trim());

  let stat;
  try {
    stat = fs.statSync(resolved);
  } catch {
    throw new Error(`File not found or not readable: ${resolved}`);
  }
  if (!stat.isFile()) {
    throw new Error(`Not a regular file: ${resolved}`);
  }

  if (ALLOWED_SEND_DIRS.length > 0) {
    const allowed = ALLOWED_SEND_DIRS.some(
      (dir) => resolved === dir || resolved.startsWith(dir + path.sep)
    );
    if (!allowed) {
      throw new Error(
        `File is outside the allowed send directories (M365_SEND_ALLOWED_DIRS): ${resolved}`
      );
    }
  }

  const raw = fs.readFileSync(resolved);
  const name = path.basename(resolved);
  const ext = getExtension(name);

  return {
    name,
    ext,
    bytes: raw.length,
    path: resolved,
    attachment: {
      "@odata.type": "#microsoft.graph.fileAttachment",
      name,
      contentType: getMimeForExt(ext),
      contentBytes: raw.toString("base64"),
    },
  };
}

function prepareOutboundAttachments(attachments) {
  const loaded = (attachments || []).map(loadLocalAttachment);
  const totalBytes = loaded.reduce((sum, item) => sum + item.bytes, 0);
  const oversizeFile = loaded.find((item) => item.bytes >= MAX_SEND_FILE_BYTES);

  if (oversizeFile) {
    throw new Error(
      `Attachment "${oversizeFile.name}" is ${formatBytes(oversizeFile.bytes)}, which exceeds the ` +
        `${formatBytes(MAX_SEND_FILE_BYTES)} per-file limit for inline sending.\n\n` +
        "This build only supports small attachments via a single Graph request. Files >= 3 MB " +
        "require the chunked upload-session flow, which is not enabled in this version."
    );
  }

  if (totalBytes > MAX_SEND_TOTAL_BYTES) {
    throw new Error(
      `Total attachment size is ${formatBytes(totalBytes)}, which exceeds the inline send budget of ` +
        `${formatBytes(MAX_SEND_TOTAL_BYTES)}. Send fewer or smaller files, or raise M365_SEND_MAX_BYTES ` +
        "(it must still stay under Microsoft Graph's ~4 MB request limit)."
    );
  }

  return { loaded, totalBytes };
}

function escapeHtml(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function textToHtmlFragment(text) {
  return escapeHtml(text).replace(/\n/g, "<br>\n");
}

function commentForOriginalBody(comment, commentType, originalContentType) {
  const originalIsHtml = String(originalContentType || "").toLowerCase() === "html";
  const commentIsHtml = commentType === "HTML";
  if (!comment) return "";
  if (originalIsHtml && !commentIsHtml) return textToHtmlFragment(comment);
  if (!originalIsHtml && commentIsHtml) return htmlToPlainText(comment);
  return comment;
}

function prependToDraftBody(draftBody, comment, commentType) {
  const contentType = String(draftBody?.contentType || "Text");
  const existing = draftBody?.content || "";
  const prepared = commentForOriginalBody(comment, commentType, contentType);
  if (!prepared) {
    return { contentType, content: existing };
  }

  if (String(contentType).toLowerCase() === "html") {
    const bodyTag = existing.match(/<body[^>]*>/i);
    if (bodyTag) {
      const idx = existing.indexOf(bodyTag[0]) + bodyTag[0].length;
      return {
        contentType: "HTML",
        content: `${existing.slice(0, idx)}${prepared}<br><br>${existing.slice(idx)}`,
      };
    }
    return { contentType: "HTML", content: `${prepared}<br><br>${existing}` };
  }

  return { contentType: "Text", content: `${prepared}\n\n${existing}` };
}

function formatRecipientAddresses(recipients) {
  return recipientAddresses(recipients).join(", ");
}

function summarizeLoadedFiles(loaded) {
  if (!loaded.length) return "(none)";
  return loaded
    .map((item) => `- ${item.name} (${formatBytes(item.bytes)}, ${item.attachment.contentType})`)
    .join("\n");
}

function createServer() {
  const server = new McpServer({
    name: APP_NAME,
    version: "0.2.0",
  });

  const healthHandler = async () => ({
    structuredContent: {
      ok: true,
      server: APP_NAME,
      transport: "stdio",
      time: new Date().toISOString(),
      auth: getPublicAuthState(),
    },
    content: [
      {
        type: "text",
        text: `${APP_NAME} is running locally over stdio.`,
      },
    ],
  });

  registerAliases(
    server,
    "health_check",
    {
      title: "Health Check",
      description:
        "Verify that the local Outlook attachment reader MCP server is running and report auth state.",
      inputSchema: z.object({}),
    },
    healthHandler
  );

  const beginAuthHandler = async () => {
    try {
      ensureConfigured();

      if (getCachedToken()) {
        return {
          structuredContent: { auth: getPublicAuthState() },
          content: [
            {
              type: "text",
              text: "Microsoft 365 is already logged in for this local MCP process.",
            },
          ],
        };
      }

      if (!authPromise && authState.status !== "pending") {
        startDeviceCodeAuth().catch((err) => {
          log("startDeviceCodeAuth background error", err?.message || String(err));
        });
      }

      await new Promise((resolve) => setTimeout(resolve, 1500));

      return {
        structuredContent: { auth: getPublicAuthState() },
        content: [
          {
            type: "text",
            text:
              authState.status === "pending"
                ? buildBeginAuthText()
                : `Current auth status: ${authState.status}${authState.error ? `\nError: ${authState.error}` : ""}`,
          },
        ],
      };
    } catch (err) {
      return {
        isError: true,
        content: [{ type: "text", text: String(err?.message || err) }],
      };
    }
  };

  registerAliases(
    server,
    "begin_auth",
    {
      title: "Begin Microsoft 365 Auth",
      description:
        "Start Microsoft 365 device-code login for the local Claude Desktop MCP process.",
      inputSchema: z.object({}),
    },
    beginAuthHandler
  );

  const authStatusHandler = async () => ({
    structuredContent: { auth: getPublicAuthState() },
    content: [
      {
        type: "text",
        text: JSON.stringify(getPublicAuthState(), null, 2),
      },
    ],
  });

  registerAliases(
    server,
    "auth_status",
    {
      title: "Microsoft 365 Auth Status",
      description:
        "Check whether Microsoft 365 login for this local MCP process has completed.",
      inputSchema: z.object({}),
    },
    authStatusHandler
  );

  const listMessagesHandler = async ({
    mailbox,
    folder,
    top,
    onlyWithAttachments,
    subjectContains,
    fromContains,
  }) => {
    try {
      const requestedTop = top;
      const perPage = 100;
      const maxPages = Math.min(Math.max(Math.ceil(requestedTop / 20), 5), 100);
      const collectionUrl = messageCollectionUrl(mailbox, folder);
      const subjectNeedle = normalizeNeedle(subjectContains);
      const fromNeedle = normalizeNeedle(fromContains);

      let nextUrl =
        `${collectionUrl}` +
        `?$select=id,subject,receivedDateTime,hasAttachments,from,toRecipients,ccRecipients,bodyPreview,isRead,conversationId` +
        `&$orderby=receivedDateTime desc` +
        `&$top=${perPage}`;

      let messages = [];
      let scannedCount = 0;
      let pages = 0;

      while (nextUrl && messages.length < requestedTop && pages < maxPages) {
        const data = await graphGetJson(nextUrl);
        const batch = data.value || [];
        scannedCount += batch.length;
        pages += 1;

        let mapped = batch.map(mapListedMessage);

        if (onlyWithAttachments) {
          mapped = mapped.filter((message) => message.hasAttachments);
        }

        mapped = mapped.filter(
          (message) =>
            matchesNeedle(message.subject, subjectNeedle) &&
            (matchesNeedle(message.fromName, fromNeedle) ||
              matchesNeedle(message.fromAddress, fromNeedle))
        );

        messages = messages.concat(mapped);
        nextUrl = data["@odata.nextLink"] || null;
      }

      messages = messages.slice(0, requestedTop);

      const text = formatMessageListText(
        messages,
        [
          "No recent messages matched the current filter.",
          `Folder searched: ${folder || "inbox"}`,
          `Messages scanned: ${scannedCount}`,
          subjectNeedle ? `Subject filter: ${subjectContains}` : "",
          fromNeedle ? `Sender filter: ${fromContains}` : "",
          onlyWithAttachments ? "Attachment filter: hasAttachments=true" : "",
        ]
          .filter(Boolean)
          .join("\n")
      );

      return {
        structuredContent: {
          mailbox,
          folder: folder || "inbox",
          requestedTop,
          scannedCount,
          pagesFetched: pages,
          returnedCount: messages.length,
          onlyWithAttachments,
          subjectContains: subjectContains || "",
          fromContains: fromContains || "",
          messages,
          nextTools: MESSAGE_CHAIN_NEXT_TOOLS,
        },
        content: [{ type: "text", text }],
      };
    } catch (err) {
      return {
        isError: true,
        content: [{ type: "text", text: String(err?.message || err) }],
      };
    }
  };

  registerAliases(
    server,
    "list_recent_messages",
    {
      title: "List Recent Outlook Messages",
      description:
        "List recent Outlook emails and return a messageId for each. Typical chain: list_recent_messages → read_email(messageId) → reply_outlook_email(messageId) to reply in the same thread. By default this searches the Inbox and prefers emails with attachments; set onlyWithAttachments=false when looking up mail to read or reply to. Can filter by subject or sender name/address.",
      inputSchema: z.object({
        mailbox: z.string().default("me"),
        folder: z.string().default("inbox"),
        top: z.number().int().min(1).max(2000).default(10),
        onlyWithAttachments: z
          .boolean()
          .default(true)
          .describe(
            "true (default) only returns emails with attachments. Set false when listing mail to read or reply to."
          ),
        subjectContains: z.string().optional().default(""),
        fromContains: z.string().optional().default(""),
      }),
    },
    listMessagesHandler
  );

  const searchMessagesHandler = async ({ mailbox, folder, query, top, onlyWithAttachments }) => {
    try {
      const collectionUrl = messageCollectionUrl(mailbox, folder);
      const searchValue = `"${String(query).replace(/"/g, '\\"')}"`;
      const perPage = Math.min(Math.max(top, 1), 100);
      const maxPages = Math.min(Math.max(Math.ceil(top / 20), 5), 100);

      let nextUrl =
        `${collectionUrl}` +
        `?$search=${encodeURIComponent(searchValue)}` +
        `&$select=id,subject,receivedDateTime,hasAttachments,from,toRecipients,ccRecipients,bodyPreview,isRead,conversationId` +
        `&$top=${perPage}`;

      let messages = [];
      let scannedCount = 0;
      let pages = 0;

      while (nextUrl && messages.length < top && pages < maxPages) {
        const data = await graphGetJson(nextUrl, { ConsistencyLevel: "eventual" });
        const batch = data.value || [];
        scannedCount += batch.length;
        pages += 1;

        let mapped = batch.map(mapListedMessage);

        if (onlyWithAttachments) {
          mapped = mapped.filter((message) => message.hasAttachments);
        }

        messages = messages.concat(mapped);
        nextUrl = data["@odata.nextLink"] || null;
      }

      messages = messages.slice(0, top);

      const text = formatMessageListText(
        messages,
        `No messages matched search query: ${query}`
      );

      return {
        structuredContent: {
          mailbox,
          folder: folder || "all",
          query,
          scannedCount,
          pagesFetched: pages,
          returnedCount: messages.length,
          onlyWithAttachments,
          messages,
          nextTools: MESSAGE_CHAIN_NEXT_TOOLS,
        },
        content: [{ type: "text", text }],
      };
    } catch (err) {
      return {
        isError: true,
        content: [{ type: "text", text: String(err?.message || err) }],
      };
    }
  };

  registerAliases(
    server,
    "search_messages",
    {
      title: "Search Outlook Messages",
      description:
        "Full-text search Outlook mail (subject, body, sender, attachments) using Microsoft Graph's native $search, unbounded by recency. Returns a messageId for each hit. Typical chain: search_messages → read_email(messageId) → reply_outlook_email(messageId) to reply in the same thread. Searches the entire mailbox by default (folder='all'); pass a specific folder (e.g. 'inbox', 'archive') to narrow it.",
      inputSchema: z.object({
        mailbox: z.string().default("me"),
        folder: z.string().default("all"),
        query: z.string(),
        top: z.number().int().min(1).max(2000).default(25),
        onlyWithAttachments: z.boolean().default(false),
      }),
    },
    searchMessagesHandler
  );

  const listFlaggedMessagesHandler = async ({ mailbox, folder, top }) => {
    try {
      const collectionUrl = messageCollectionUrl(mailbox, folder);
      const perPage = Math.min(Math.max(top, 1), 100);
      const maxPages = Math.min(Math.max(Math.ceil(top / 20), 5), 100);

      // Graph rejects $orderby combined with a flag/flagStatus $filter as an
      // "InefficientFilter"; results come back newest-first-ish per page but
      // are not guaranteed globally sorted, so pull enough pages to satisfy top.
      let nextUrl =
        `${collectionUrl}` +
        `?$filter=${encodeURIComponent("flag/flagStatus eq 'flagged'")}` +
        `&$select=id,subject,receivedDateTime,hasAttachments,from,toRecipients,ccRecipients,bodyPreview,conversationId` +
        `&$top=${perPage}`;

      let messages = [];
      let scannedCount = 0;
      let pages = 0;

      while (nextUrl && messages.length < top && pages < maxPages) {
        const data = await graphGetJson(nextUrl, { ConsistencyLevel: "eventual" });
        const batch = data.value || [];
        scannedCount += batch.length;
        pages += 1;

        const mapped = batch.map(mapListedMessage);

        messages = messages.concat(mapped);
        nextUrl = data["@odata.nextLink"] || null;
      }

      messages.sort((a, b) => (a.receivedDateTime < b.receivedDateTime ? 1 : -1));
      messages = messages.slice(0, top);

      const text = formatMessageListText(
        messages,
        `No flagged messages found in folder: ${folder || "inbox"}`
      );

      return {
        structuredContent: {
          mailbox,
          folder: folder || "inbox",
          scannedCount,
          pagesFetched: pages,
          returnedCount: messages.length,
          messages,
          nextTools: MESSAGE_CHAIN_NEXT_TOOLS,
        },
        content: [{ type: "text", text }],
      };
    } catch (err) {
      return {
        isError: true,
        content: [{ type: "text", text: String(err?.message || err) }],
      };
    }
  };

  registerAliases(
    server,
    "list_flagged_messages",
    {
      title: "List Flagged Outlook Messages",
      description:
        "List Outlook messages flagged for follow-up (Graph flag/flagStatus = 'flagged') and return a messageId for each. Typical chain: list_flagged_messages → read_email(messageId) → reply_outlook_email(messageId) to reply in the same thread. Defaults to the inbox; pass folder='all' to search the whole mailbox. Newest-flagged-first.",
      inputSchema: z.object({
        mailbox: z.string().default("me"),
        folder: z.string().default("inbox"),
        top: z.number().int().min(1).max(500).default(50),
      }),
    },
    listFlaggedMessagesHandler
  );

  const listAttachmentsHandler = async ({ messageId, mailbox }) => {
    try {
      const base = mailboxPrefix(mailbox);
      const url = `${base}/messages/${encodeURIComponent(messageId)}/attachments`;
      const data = await graphGetJson(url);

      const attachments = (data.value || []).map((attachment) => ({
        id: attachment.id,
        name: attachment.name,
        contentType: attachment.contentType,
        size: attachment.size,
        isInline: attachment.isInline,
        kind: (attachment["@odata.type"] || "").split(".").pop() || "attachment",
      }));

      const text =
        attachments.length > 0
          ? [
              "These attachments are fetched directly from Microsoft Graph by attachment ID, so they are not limited by normal chat file-upload size caps. Large inline image previews may be downscaled before being returned to Claude.",
              attachments
                .map(
                  (attachment, index) =>
                    `${index + 1}. ${attachment.name || "(unnamed attachment)"}\nAttachment ID: ${attachment.id}\nType: ${attachment.contentType || "application/octet-stream"}\nSize: ${attachment.size || 0} bytes\nInline: ${attachment.isInline ? "yes" : "no"}\nKind: ${attachment.kind}`
                )
                .join("\n\n"),
            ].join("\n\n")
          : "This email has no downloadable attachments.";

      return {
        structuredContent: {
          messageId,
          attachments,
          nextTools: MESSAGE_CHAIN_NEXT_TOOLS,
        },
        content: [{ type: "text", text }],
      };
    } catch (err) {
      return {
        isError: true,
        content: [{ type: "text", text: String(err?.message || err) }],
      };
    }
  };

  registerAliases(
    server,
    "list_email_attachments",
    {
      title: "List Email Attachments",
      description:
        "List attachments for a specific Outlook email. Pass the same messageId from list_recent_messages, search_messages, or read_email. After reading, reply in-thread with reply_outlook_email using that messageId.",
      inputSchema: z.object({
        messageId: z.string().describe(MESSAGE_ID_DESCRIPTION),
        mailbox: z.string().default("me"),
      }),
    },
    listAttachmentsHandler
  );

  const readAttachmentHandler = async ({ messageId, attachmentId, mailbox }) => {
    try {
      const base = mailboxPrefix(mailbox);
      const metaUrl =
        `${base}/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`;
      const meta = await graphGetJson(metaUrl);
      const kind = (meta["@odata.type"] || "").split(".").pop() || "attachment";

      if (kind === "referenceAttachment") {
        const output = {
          id: meta.id,
          name: meta.name,
          contentType: meta.contentType,
          size: meta.size,
          isInline: meta.isInline,
          kind,
        };

        return {
          structuredContent: { attachment: output },
          content: [
            {
              type: "text",
              text:
                "This is a referenceAttachment. Only metadata is available; content cannot be downloaded directly.\n\n" +
                JSON.stringify(output, null, 2),
            },
          ],
        };
      }

      const rawUrl =
        `${base}/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}/$value`;
      const raw = await graphGetBytes(rawUrl);
      const filename = safeFilename(meta.name || attachmentId);
      const ext = getExtension(filename);
      const tempPath = path.join(os.tmpdir(), filename);

      fs.writeFileSync(tempPath, raw);

      const contentType = (meta.contentType || "application/octet-stream").toLowerCase();
      const parseContext = createParseContext();
      let parsed;

      try {
        parsed = await parseAttachmentPayload(
          {
            raw,
            filename,
            contentType,
            kind,
          },
          parseContext
        );
      } catch (parseError) {
        log("File parsing error", parseError);
        parsed = {
          parser: "parse_error",
          text: `(Warning: An error occurred while parsing this file: ${parseError.message})`,
          details: {},
        };
      }

      const truncated = truncateExtractedText(parsed.text);
      const previewParts = [
        "[Local MCP backend parsed the attachment]",
        `Filename: ${filename}`,
        `Parser: ${parsed.parser}`,
      ];

      if (parseContext.imageBlocks.length > 0) {
        previewParts.push(`Visual blocks attached: ${parseContext.imageBlocks.length}`);
      }

      if (parseContext.imageNotes.length > 0) {
        previewParts.push(
          `Image handling notes:\n${parseContext.imageNotes.map((note) => `- ${note}`).join("\n")}`
        );
      }

      if (truncated.truncated) {
        previewParts.push(
          `Notice: returned ${truncated.returnedLength.toLocaleString()} of ${truncated.totalLength.toLocaleString()} characters; omitted ${truncated.omittedLength.toLocaleString()} characters.`
        );
      }

      previewParts.push(truncated.text);

      return {
        structuredContent: {
          attachment: {
            id: meta.id,
            name: meta.name,
            contentType: meta.contentType,
            size: meta.size,
            isInline: meta.isInline,
            kind,
            hostTempPath: tempPath,
            ext,
            parser: parsed.parser,
            imageBlockCount: parseContext.imageBlocks.length,
            extractedTextLength: truncated.totalLength,
            returnedTextLength: truncated.returnedLength,
            truncated: truncated.truncated,
            omittedCharacters: truncated.omittedLength,
            ...(parsed.details || {}),
          },
        },
        content: [{ type: "text", text: previewParts.join("\n\n") }, ...parseContext.imageBlocks],
      };
    } catch (err) {
      return {
        isError: true,
        content: [{ type: "text", text: String(err?.message || err) }],
      };
    }
  };

  registerAliases(
    server,
    "read_email_attachment",
    {
      title: "Read Email Attachment",
      description:
        "Download an Outlook attachment directly from Microsoft Graph and parse it locally. Pass messageId from list_recent_messages/search_messages and attachmentId from list_email_attachments. Supports PDF, OCR-scanned PDF, Word, PowerPoint, Excel, images, archives, MSG, and plain text. Large image previews are automatically downscaled to fit MCP payload limits.",
      inputSchema: z.object({
        messageId: z.string().describe(MESSAGE_ID_DESCRIPTION),
        attachmentId: z.string(),
        mailbox: z.string().default("me"),
      }),
    },
    readAttachmentHandler
  );

  async function fetchEmailBody(mailbox, messageId) {
    const base = mailboxPrefix(mailbox);
    const url =
      `${base}/messages/${encodeURIComponent(messageId)}` +
      `?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,hasAttachments,bodyPreview,body,webLink,conversationId`;
    const m = await graphGetJson(url, {
      Prefer: 'outlook.body-content-type="text"',
    });

    const isHtml = String(m.body?.contentType || "").toLowerCase() === "html";
    const bodyText = isHtml
      ? htmlToPlainText(m.body?.content)
      : normalizeEmailBodyText(m.body?.content);

    return { message: m, bodyText };
  }

  const readEmailHandler = async ({ messageId, mailbox }) => {
    try {
      const { message: m, bodyText } = await fetchEmailBody(mailbox, messageId);

      const fromStr = formatSender(
        m.from?.emailAddress?.name,
        m.from?.emailAddress?.address
      );
      const toStr = (m.toRecipients || [])
        .map((r) => r.emailAddress?.address)
        .filter(Boolean)
        .join(", ");
      const ccStr = (m.ccRecipients || [])
        .map((r) => r.emailAddress?.address)
        .filter(Boolean)
        .join(", ");

      const truncated = truncateExtractedText(bodyText);

      const headerLines = [
        `Subject: ${m.subject || "(no subject)"}`,
        `From: ${fromStr}`,
        toStr ? `To: ${toStr}` : "",
        ccStr ? `Cc: ${ccStr}` : "",
        `Received: ${m.receivedDateTime || ""}`,
        `Has attachments: ${m.hasAttachments ? "yes" : "no"}`,
        `messageId: ${m.id}`,
      ].filter(Boolean);

      if (truncated.truncated) {
        headerLines.push(
          `Notice: body truncated to ${truncated.returnedLength.toLocaleString()} of ${truncated.totalLength.toLocaleString()} characters. Call read_email_body_chunk with this messageId and offset=${truncated.returnedLength} to continue reading the rest.`
        );
      }

      headerLines.push(
        "To reply in this same Outlook thread, call reply_outlook_email with this messageId. Do not use send_outlook_email."
      );

      const textOut = `${headerLines.join("\n")}\n\n${truncated.text || "(empty body)"}`;

      return {
        structuredContent: {
          messageId: m.id,
          message: {
            id: m.id,
            messageId: m.id,
            subject: m.subject,
            from: fromStr,
            to: toStr,
            cc: ccStr,
            receivedDateTime: m.receivedDateTime,
            sentDateTime: m.sentDateTime,
            hasAttachments: !!m.hasAttachments,
            webLink: m.webLink,
            conversationId: m.conversationId || null,
            bodyPreview: m.bodyPreview || "",
            bodyTextLength: bodyText.length,
            truncated: truncated.truncated,
            nextOffset: truncated.truncated ? truncated.returnedLength : null,
          },
          nextTools: {
            replyInThread: "reply_outlook_email",
            readMore: truncated.truncated ? "read_email_body_chunk" : null,
            listAttachments: m.hasAttachments ? "list_email_attachments" : null,
            passField: "messageId",
          },
        },
        content: [{ type: "text", text: textOut }],
      };
    } catch (err) {
      return {
        isError: true,
        content: [{ type: "text", text: String(err?.message || err) }],
      };
    }
  };

  registerAliases(
    server,
    "read_email",
    {
      title: "Read Outlook Email Body",
      description:
        "Read one Outlook email by messageId (from list_recent_messages, search_messages, or list_flagged_messages): subject, sender, recipients, date, and body (HTML converted to plain text). After reading, reply in the same thread with reply_outlook_email using this same messageId — do not compose a new mail with send_outlook_email. For long bodies, continue with read_email_body_chunk. For files, use list_email_attachments / read_email_attachment.",
      inputSchema: z.object({
        messageId: z.string().describe(MESSAGE_ID_DESCRIPTION),
        mailbox: z.string().default("me"),
      }),
    },
    readEmailHandler
  );

  const readEmailBodyChunkHandler = async ({ messageId, mailbox, offset, maxChars }) => {
    try {
      const { message: m, bodyText } = await fetchEmailBody(mailbox, messageId);

      const totalLength = bodyText.length;
      let start = Math.min(Math.max(offset || 0, 0), totalLength);
      if (
        start > 0 &&
        /[\uDC00-\uDFFF]/.test(bodyText[start]) &&
        /[\uD800-\uDBFF]/.test(bodyText[start - 1])
      ) {
        start -= 1;
      }

      const chunkSize = Math.min(
        Math.max(maxChars || MAX_EMAIL_BODY_CHUNK_CHARS, 1),
        MAX_EMAIL_BODY_CHUNK_CHARS
      );
      let end = Math.min(start + chunkSize, totalLength);

      // Do not split a Unicode surrogate pair between consecutive chunks.
      if (
        end < totalLength &&
        end > start &&
        /[\uD800-\uDBFF]/.test(bodyText[end - 1]) &&
        /[\uDC00-\uDFFF]/.test(bodyText[end])
      ) {
        end += 1;
      }

      const chunkText = bodyText.slice(start, end);
      const hasMore = end < totalLength;

      const headerLines = [
        `Subject: ${m.subject || "(no subject)"}`,
        `Body length: ${totalLength.toLocaleString()} characters total`,
        chunkText
          ? `Returning characters ${start.toLocaleString()}-${(end - 1).toLocaleString()}`
          : "Returning an empty body",
        hasMore
          ? `More text remains. Call this tool again with offset=${end} to continue reading.`
          : "This is the end of the message body.",
      ];

      const textOut = `${headerLines.join("\n")}\n\n${chunkText || "(empty body)"}`;

      return {
        structuredContent: {
          messageId: m.id,
          subject: m.subject,
          offset: start,
          nextOffset: hasMore ? end : null,
          returnedLength: chunkText.length,
          totalLength,
          hasMore,
          bodyText: chunkText,
        },
        content: [{ type: "text", text: textOut }],
      };
    } catch (err) {
      return {
        isError: true,
        content: [{ type: "text", text: String(err?.message || err) }],
      };
    }
  };

  registerAliases(
    server,
    "read_email_body_chunk",
    {
      title: "Read Outlook Email Body (Full, Chunked)",
      description:
        "Fetch a slice of the full plain-text body of a specific Outlook email, for cases where read_email's body was cut off by its character limit. Pass the same messageId used with read_email. Call with offset=0 first, then re-call with the returned nextOffset until hasMore is false. After reading, reply in-thread with reply_outlook_email using this messageId.",
      inputSchema: z.object({
        messageId: z.string().describe(MESSAGE_ID_DESCRIPTION),
        mailbox: z.string().default("me"),
        offset: z.number().int().min(0).default(0),
        maxChars: z
          .number()
          .int()
          .min(1)
          .max(MAX_EMAIL_BODY_CHUNK_CHARS)
          .default(MAX_EMAIL_BODY_CHUNK_CHARS),
      }),
    },
    readEmailBodyChunkHandler
  );

  const sendEmailHandler = async ({
    to,
    cc,
    bcc,
    subject,
    body,
    bodyType,
    attachments,
    send_now,
    save_to_sent,
    mailbox,
  }) => {
    try {
      ensureConfigured();

      const toRecipients = buildRecipients(to);
      if (toRecipients.length === 0) {
        return {
          isError: true,
          content: [
            { type: "text", text: "At least one recipient is required in 'to'." },
          ],
        };
      }

      let loaded;
      let totalBytes;
      try {
        ({ loaded, totalBytes } = prepareOutboundAttachments(attachments));
      } catch (fileErr) {
        return {
          isError: true,
          content: [{ type: "text", text: String(fileErr?.message || fileErr) }],
        };
      }

      const message = {
        subject: subject || "",
        body: {
          contentType: bodyType === "HTML" ? "HTML" : "Text",
          content: body || "",
        },
        toRecipients,
      };

      const ccRecipients = buildRecipients(cc);
      const bccRecipients = buildRecipients(bcc);
      if (ccRecipients.length) message.ccRecipients = ccRecipients;
      if (bccRecipients.length) message.bccRecipients = bccRecipients;
      if (loaded.length) {
        message.attachments = loaded.map((item) => item.attachment);
      }

      const base = mailboxPrefix(mailbox);
      const fileSummary = summarizeLoadedFiles(loaded);
      const recipientSummary = [
        `To: ${normalizeEmailList(to).join(", ")}`,
        ccRecipients.length ? `Cc: ${normalizeEmailList(cc).join(", ")}` : "",
        bccRecipients.length ? `Bcc: ${normalizeEmailList(bcc).join(", ")}` : "",
      ]
        .filter(Boolean)
        .join("\n");
      const recipientCount =
        toRecipients.length + ccRecipients.length + bccRecipients.length;

      if (send_now) {
        await graphSend("POST", `${base}/sendMail`, {
          message,
          saveToSentItems: save_to_sent !== false,
        });

        return {
          structuredContent: {
            action: "sent",
            mailbox: mailbox || "me",
            subject: message.subject,
            recipientCount,
            attachmentCount: loaded.length,
            totalAttachmentBytes: totalBytes,
            savedToSentItems: save_to_sent !== false,
          },
          content: [
            {
              type: "text",
              text: [
                "Email sent via Microsoft Graph.",
                recipientSummary,
                `Subject: ${message.subject || "(no subject)"}`,
                `Attachments:\n${fileSummary}`,
                `Saved to Sent Items: ${save_to_sent !== false ? "yes" : "no"}`,
              ].join("\n\n"),
            },
          ],
        };
      }

      const draft = await graphSend("POST", `${base}/messages`, message);

      return {
        structuredContent: {
          action: "draft_created",
          mailbox: mailbox || "me",
          draftId: draft?.id || null,
          webLink: draft?.webLink || null,
          subject: message.subject,
          recipientCount,
          attachmentCount: loaded.length,
          totalAttachmentBytes: totalBytes,
          nextTools: DRAFT_NEXT_TOOLS,
        },
        content: [
          {
            type: "text",
            text: [
              "Draft created in your Outlook Drafts folder (nothing was sent).",
              recipientSummary,
              `Subject: ${message.subject || "(no subject)"}`,
              `Attachments:\n${fileSummary}`,
              `draftId: ${draft?.id || ""}`,
              draft?.webLink ? `Open to review and send:\n${draft.webLink}` : "",
              "To change this draft, call edit_outlook_draft with this draftId. To discard it, call delete_outlook_draft. To send without editing, call this tool again with send_now=true.",
            ]
              .filter(Boolean)
              .join("\n\n"),
          },
        ],
      };
    } catch (err) {
      return {
        isError: true,
        content: [{ type: "text", text: String(err?.message || err) }],
      };
    }
  };

  registerAliases(
    server,
    "send_outlook_email",
    {
      title: "Send Outlook Email with Local Attachments",
      description:
        "Compose a NEW Outlook email with local files attached (e.g. a daily dashboard .pptx) and either save it as a draft for review (default) or send it immediately. This starts a new conversation. Do not use it to reply to an existing message — use reply_outlook_email so the reply stays in the original Outlook thread. Attachments are read from the local filesystem and inlined via Microsoft Graph; each file must be under 3 MB. Set send_now=true to send without review.",
      inputSchema: z.object({
        to: z
          .union([z.string(), z.array(z.string())])
          .describe("Recipient email address(es); a string (comma/semicolon separated) or an array."),
        cc: z.union([z.string(), z.array(z.string())]).optional(),
        bcc: z.union([z.string(), z.array(z.string())]).optional(),
        subject: z.string().optional().default(""),
        body: z.string().optional().default(""),
        bodyType: z.enum(["Text", "HTML"]).optional().default("Text"),
        attachments: z
          .array(z.string())
          .optional()
          .default([])
          .describe("Absolute or relative local file paths to attach."),
        send_now: z
          .boolean()
          .optional()
          .default(false)
          .describe("false (default) saves a draft; true sends immediately."),
        save_to_sent: z.boolean().optional().default(true),
        mailbox: z.string().default("me"),
      }),
    },
    sendEmailHandler
  );

  const replyEmailHandler = async ({
    messageId,
    body,
    bodyType,
    replyAll,
    to,
    cc,
    bcc,
    attachments,
    send_now,
    mailbox,
  }) => {
    try {
      ensureConfigured();

      if (!messageId || !String(messageId).trim()) {
        return {
          isError: true,
          content: [
            {
              type: "text",
              text: "messageId is required. Use list_recent_messages, search_messages, or read_email to get it.",
            },
          ],
        };
      }

      let loaded;
      let totalBytes;
      try {
        ({ loaded, totalBytes } = prepareOutboundAttachments(attachments));
      } catch (fileErr) {
        return {
          isError: true,
          content: [{ type: "text", text: String(fileErr?.message || fileErr) }],
        };
      }

      const base = mailboxPrefix(mailbox);
      const encodedId = encodeURIComponent(messageId);
      const toRecipients = buildRecipients(to);
      const ccRecipients = buildRecipients(cc);
      const bccRecipients = buildRecipients(bcc);
      const replyMessage = {};
      if (toRecipients.length) replyMessage.toRecipients = toRecipients;
      if (ccRecipients.length) replyMessage.ccRecipients = ccRecipients;
      if (bccRecipients.length) replyMessage.bccRecipients = bccRecipients;
      if (loaded.length) {
        replyMessage.attachments = loaded.map((item) => item.attachment);
      }

      const fileSummary = summarizeLoadedFiles(loaded);
      const replyKind = replyAll ? "reply-all" : "reply";

      if (send_now) {
        const original = await graphGetJson(
          `${base}/messages/${encodedId}?$select=id,subject,conversationId,body`
        );
        const payload = {
          comment: commentForOriginalBody(
            body || "",
            bodyType,
            original.body?.contentType
          ),
        };
        if (Object.keys(replyMessage).length) {
          payload.message = replyMessage;
        }

        await graphSend(
          "POST",
          `${base}/messages/${encodedId}/${replyAll ? "replyAll" : "reply"}`,
          payload
        );

        return {
          structuredContent: {
            action: "replied",
            replyKind,
            mailbox: mailbox || "me",
            originalMessageId: original.id || messageId,
            conversationId: original.conversationId || null,
            subject: original.subject || "",
            attachmentCount: loaded.length,
            totalAttachmentBytes: totalBytes,
          },
          content: [
            {
              type: "text",
              text: [
                `Reply sent via Microsoft Graph (${replyKind}). It is in the original Outlook conversation, not a new thread.`,
                `Subject: ${original.subject || "(no subject)"}`,
                original.conversationId
                  ? `Conversation ID: ${original.conversationId}`
                  : "",
                `Attachments:\n${fileSummary}`,
              ]
                .filter(Boolean)
                .join("\n\n"),
            },
          ],
        };
      }

      const draft = await graphSend(
        "POST",
        `${base}/messages/${encodedId}/${replyAll ? "createReplyAll" : "createReply"}`,
        {}
      );
      if (!draft?.id) {
        throw new Error(
          "Microsoft Graph did not return a reply draft. The original message may no longer be available."
        );
      }

      const patch = {};
      if (toRecipients.length) patch.toRecipients = toRecipients;
      if (ccRecipients.length) patch.ccRecipients = ccRecipients;
      if (bccRecipients.length) patch.bccRecipients = bccRecipients;
      if (body) {
        patch.body = prependToDraftBody(draft.body, body, bodyType);
      }

      let updated = draft;
      if (Object.keys(patch).length) {
        updated =
          (await graphSend(
            "PATCH",
            `${base}/messages/${encodeURIComponent(draft.id)}`,
            patch
          )) || draft;
      }

      for (const item of loaded) {
        await graphSend(
          "POST",
          `${base}/messages/${encodeURIComponent(draft.id)}/attachments`,
          item.attachment
        );
      }

      const toStr =
        formatRecipientAddresses(updated.toRecipients) ||
        formatRecipientAddresses(draft.toRecipients);
      const ccStr =
        formatRecipientAddresses(updated.ccRecipients) ||
        formatRecipientAddresses(draft.ccRecipients);
      const recipientSummary = [
        toStr ? `To: ${toStr}` : "",
        ccStr ? `Cc: ${ccStr}` : "",
        bccRecipients.length ? `Bcc: ${normalizeEmailList(bcc).join(", ")}` : "",
      ]
        .filter(Boolean)
        .join("\n");

      return {
        structuredContent: {
          action: "reply_draft_created",
          replyKind,
          mailbox: mailbox || "me",
          originalMessageId: messageId,
          draftId: draft.id,
          webLink: updated.webLink || draft.webLink || null,
          conversationId: draft.conversationId || null,
          subject: updated.subject || draft.subject || "",
          attachmentCount: loaded.length,
          totalAttachmentBytes: totalBytes,
          nextTools: DRAFT_NEXT_TOOLS,
        },
        content: [
          {
            type: "text",
            text: [
              `Reply draft created in Outlook (${replyKind}). It is in the original conversation — nothing was sent.`,
              recipientSummary,
              `Subject: ${updated.subject || draft.subject || "(no subject)"}`,
              `draftId: ${draft.id}`,
              draft.conversationId
                ? `Conversation ID: ${draft.conversationId}`
                : "",
              `Attachments:\n${fileSummary}`,
              updated.webLink || draft.webLink
                ? `Open to review and send:\n${updated.webLink || draft.webLink}`
                : "",
              "To change this draft, call edit_outlook_draft with this draftId. To discard it, call delete_outlook_draft. To send this threaded reply without editing, call this tool again with send_now=true.",
            ]
              .filter(Boolean)
              .join("\n\n"),
          },
        ],
      };
    } catch (err) {
      return {
        isError: true,
        content: [{ type: "text", text: String(err?.message || err) }],
      };
    }
  };

  registerAliases(
    server,
    "reply_outlook_email",
    {
      title: "Reply to Outlook Email (Same Thread)",
      description:
        "Reply in the same Outlook conversation as an existing message. Typical chain: list_recent_messages or search_messages → read_email(messageId) → reply_outlook_email(messageId, body). Pass the messageId of the mail you just listed or read; Graph createReply/reply keeps conversationId and In-Reply-To. Defaults to a draft; set send_now=true to send. Use replyAll=true when the original had multiple recipients. Do not use send_outlook_email for replies — that starts a new conversation.",
      inputSchema: z.object({
        messageId: z.string().describe(MESSAGE_ID_DESCRIPTION),
        body: z
          .string()
          .optional()
          .default("")
          .describe("Reply text. Graph prepends this to the quoted original message."),
        bodyType: z.enum(["Text", "HTML"]).optional().default("Text"),
        replyAll: z
          .boolean()
          .optional()
          .default(false)
          .describe("true replies to the sender and all original recipients."),
        to: z
          .union([z.string(), z.array(z.string())])
          .optional()
          .describe(
            "Optional To override. Omit to use Graph's default (original sender, or all recipients when replyAll is true)."
          ),
        cc: z.union([z.string(), z.array(z.string())]).optional(),
        bcc: z.union([z.string(), z.array(z.string())]).optional(),
        attachments: z
          .array(z.string())
          .optional()
          .default([])
          .describe("Absolute or relative local file paths to attach."),
        send_now: z
          .boolean()
          .optional()
          .default(false)
          .describe("false (default) saves a threaded draft; true sends immediately."),
        mailbox: z.string().default("me"),
      }),
    },
    replyEmailHandler
  );

  async function fetchDraft(mailbox, draftId) {
    const base = mailboxPrefix(mailbox);
    const draft = await graphGetJson(
      `${base}/messages/${encodeURIComponent(draftId)}` +
        `?$select=id,subject,isDraft,body,toRecipients,ccRecipients,bccRecipients,webLink,conversationId,hasAttachments,lastModifiedDateTime,from`
    );
    if (!draft?.isDraft) {
      throw new Error(
        "That message is not an unsent draft. edit_outlook_draft and delete_outlook_draft only work on drafts (draftId from send_outlook_email, reply_outlook_email, or list_outlook_drafts)."
      );
    }
    return draft;
  }

  const listDraftsHandler = async ({ mailbox, top }) => {
    try {
      const collectionUrl = messageCollectionUrl(mailbox, "drafts");
      const perPage = Math.min(Math.max(top, 1), 100);
      let nextUrl =
        `${collectionUrl}` +
        `?$select=id,subject,lastModifiedDateTime,hasAttachments,toRecipients,ccRecipients,bodyPreview,conversationId,isDraft` +
        `&$orderby=lastModifiedDateTime desc` +
        `&$top=${perPage}`;

      let drafts = [];
      let pages = 0;
      const maxPages = Math.min(Math.max(Math.ceil(top / 20), 5), 20);

      while (nextUrl && drafts.length < top && pages < maxPages) {
        const data = await graphGetJson(nextUrl);
        const batch = data.value || [];
        pages += 1;
        drafts = drafts.concat(
          batch.map((message) => ({
            ...mapListedMessage(message),
            draftId: message.id,
            lastModifiedDateTime: message.lastModifiedDateTime || "",
          }))
        );
        nextUrl = data["@odata.nextLink"] || null;
      }

      drafts = drafts.slice(0, top);
      const text = drafts.length
        ? `${drafts
            .map((draft, index) => {
              const lines = [
                `${index + 1}. ${draft.subject || "(no subject)"}`,
                draft.to?.length ? `To: ${draft.to.join(", ")}` : "",
                draft.lastModifiedDateTime
                  ? `Modified: ${draft.lastModifiedDateTime}`
                  : "",
                `Has Attachments: ${draft.hasAttachments}`,
                `draftId: ${draft.draftId}`,
              ];
              return lines.filter(Boolean).join("\n");
            })
            .join("\n\n")}\n\n${DRAFT_CHAIN_FOOTER}`
        : "No drafts found in the Drafts folder.";

      return {
        structuredContent: {
          mailbox: mailbox || "me",
          returnedCount: drafts.length,
          drafts,
          nextTools: DRAFT_NEXT_TOOLS,
        },
        content: [{ type: "text", text }],
      };
    } catch (err) {
      return {
        isError: true,
        content: [{ type: "text", text: String(err?.message || err) }],
      };
    }
  };

  registerAliases(
    server,
    "list_outlook_drafts",
    {
      title: "List Outlook Drafts",
      description:
        "List unsent Outlook drafts, newest-modified first. Each result includes a draftId. Typical chain: list_outlook_drafts → read_email(draftId) → edit_outlook_draft(draftId) or delete_outlook_draft(draftId). send_outlook_email and reply_outlook_email also return a draftId when they save a draft.",
      inputSchema: z.object({
        mailbox: z.string().default("me"),
        top: z.number().int().min(1).max(200).default(25),
      }),
    },
    listDraftsHandler
  );

  const editDraftHandler = async ({
    draftId,
    to,
    cc,
    bcc,
    subject,
    body,
    bodyType,
    attachments,
    mailbox,
  }) => {
    try {
      ensureConfigured();

      if (!draftId || !String(draftId).trim()) {
        return {
          isError: true,
          content: [
            {
              type: "text",
              text: "draftId is required. Use list_outlook_drafts, or the draftId returned by send_outlook_email / reply_outlook_email.",
            },
          ],
        };
      }

      const draft = await fetchDraft(mailbox, draftId);
      const patch = {};
      const toRecipients = to === undefined ? [] : buildRecipients(to);
      const ccRecipients = cc === undefined ? [] : buildRecipients(cc);
      const bccRecipients = bcc === undefined ? [] : buildRecipients(bcc);

      if (to !== undefined) {
        if (toRecipients.length === 0) {
          return {
            isError: true,
            content: [
              { type: "text", text: "If you set 'to', include at least one recipient." },
            ],
          };
        }
        patch.toRecipients = toRecipients;
      }
      if (cc !== undefined) patch.ccRecipients = ccRecipients;
      if (bcc !== undefined) patch.bccRecipients = bccRecipients;
      if (subject !== undefined) patch.subject = subject;
      if (body !== undefined) {
        const existingType = String(draft.body?.contentType || "Text");
        const requestedType = bodyType || existingType;
        const contentType =
          String(requestedType).toLowerCase() === "html" ? "HTML" : "Text";
        patch.body = {
          contentType,
          content: body,
        };
      }

      let loaded = [];
      let totalBytes = 0;
      if (attachments && attachments.length) {
        try {
          ({ loaded, totalBytes } = prepareOutboundAttachments(attachments));
        } catch (fileErr) {
          return {
            isError: true,
            content: [{ type: "text", text: String(fileErr?.message || fileErr) }],
          };
        }
      }

      if (!Object.keys(patch).length && !loaded.length) {
        return {
          isError: true,
          content: [
            {
              type: "text",
              text: "Nothing to change. Pass at least one of: subject, body, to, cc, bcc, attachments.",
            },
          ],
        };
      }

      const base = mailboxPrefix(mailbox);
      const encodedId = encodeURIComponent(draft.id);
      let updated = draft;
      if (Object.keys(patch).length) {
        updated = (await graphSend("PATCH", `${base}/messages/${encodedId}`, patch)) || draft;
      }

      for (const item of loaded) {
        await graphSend(
          "POST",
          `${base}/messages/${encodedId}/attachments`,
          item.attachment
        );
      }

      const toStr =
        formatRecipientAddresses(updated.toRecipients) ||
        formatRecipientAddresses(draft.toRecipients);
      const ccStr =
        formatRecipientAddresses(updated.ccRecipients) ||
        formatRecipientAddresses(draft.ccRecipients);
      const fileSummary = summarizeLoadedFiles(loaded);

      return {
        structuredContent: {
          action: "draft_updated",
          mailbox: mailbox || "me",
          draftId: draft.id,
          webLink: updated.webLink || draft.webLink || null,
          subject: updated.subject || draft.subject || "",
          addedAttachmentCount: loaded.length,
          totalAddedAttachmentBytes: totalBytes,
          nextTools: DRAFT_NEXT_TOOLS,
        },
        content: [
          {
            type: "text",
            text: [
              "Draft updated (nothing was sent).",
              toStr ? `To: ${toStr}` : "",
              ccStr ? `Cc: ${ccStr}` : "",
              `Subject: ${updated.subject || draft.subject || "(no subject)"}`,
              `draftId: ${draft.id}`,
              loaded.length ? `Added attachments:\n${fileSummary}` : "",
              updated.webLink || draft.webLink
                ? `Open to review and send:\n${updated.webLink || draft.webLink}`
                : "",
              DRAFT_CHAIN_FOOTER,
            ]
              .filter(Boolean)
              .join("\n\n"),
          },
        ],
      };
    } catch (err) {
      return {
        isError: true,
        content: [{ type: "text", text: String(err?.message || err) }],
      };
    }
  };

  registerAliases(
    server,
    "edit_outlook_draft",
    {
      title: "Edit Outlook Draft",
      description:
        "Update an unsent Outlook draft. Pass draftId from send_outlook_email, reply_outlook_email, or list_outlook_drafts. Only fields you set are changed; omitted fields stay as they are. Providing body replaces the entire draft body (including any quoted original on a reply draft). New attachments are added; existing ones are left in place. Does not send. To discard instead, use delete_outlook_draft.",
      inputSchema: z.object({
        draftId: z.string().describe(DRAFT_ID_DESCRIPTION),
        to: z
          .union([z.string(), z.array(z.string())])
          .optional()
          .describe("Replace To recipients. Omit to leave unchanged."),
        cc: z
          .union([z.string(), z.array(z.string())])
          .optional()
          .describe("Replace Cc recipients. Omit to leave unchanged."),
        bcc: z
          .union([z.string(), z.array(z.string())])
          .optional()
          .describe("Replace Bcc recipients. Omit to leave unchanged."),
        subject: z.string().optional().describe("Replace subject. Omit to leave unchanged."),
        body: z
          .string()
          .optional()
          .describe("Replace the entire draft body. Omit to leave unchanged."),
        bodyType: z
          .enum(["Text", "HTML"])
          .optional()
          .describe("Used only when body is set. Defaults to the draft's current body type."),
        attachments: z
          .array(z.string())
          .optional()
          .describe("Local file paths to add as attachments. Existing attachments are kept."),
        mailbox: z.string().default("me"),
      }),
    },
    editDraftHandler
  );

  const deleteDraftHandler = async ({ draftId, mailbox }) => {
    try {
      ensureConfigured();

      if (!draftId || !String(draftId).trim()) {
        return {
          isError: true,
          content: [
            {
              type: "text",
              text: "draftId is required. Use list_outlook_drafts, or the draftId returned by send_outlook_email / reply_outlook_email.",
            },
          ],
        };
      }

      const draft = await fetchDraft(mailbox, draftId);
      const base = mailboxPrefix(mailbox);
      await graphDelete(`${base}/messages/${encodeURIComponent(draft.id)}`);

      return {
        structuredContent: {
          action: "draft_deleted",
          mailbox: mailbox || "me",
          draftId: draft.id,
          subject: draft.subject || "",
        },
        content: [
          {
            type: "text",
            text: [
              "Draft deleted. It was not sent.",
              `Subject: ${draft.subject || "(no subject)"}`,
              `draftId: ${draft.id}`,
            ].join("\n"),
          },
        ],
      };
    } catch (err) {
      return {
        isError: true,
        content: [{ type: "text", text: String(err?.message || err) }],
      };
    }
  };

  registerAliases(
    server,
    "delete_outlook_draft",
    {
      title: "Delete Outlook Draft",
      description:
        "Permanently delete an unsent Outlook draft. Pass draftId from send_outlook_email, reply_outlook_email, or list_outlook_drafts. Refuses to delete messages that are not drafts. This cannot be undone.",
      inputSchema: z.object({
        draftId: z.string().describe(DRAFT_ID_DESCRIPTION),
        mailbox: z.string().default("me"),
      }),
    },
    deleteDraftHandler
  );

  return server;
}

let activeServer = null;
let activeTransport = null;

async function shutdown(signal) {
  log(`shutdown requested: ${signal}`);

  try {
    if (activeTransport) {
      await activeTransport.close();
    }
  } catch (err) {
    log("transport close failed", err);
  }

  try {
    if (activeServer) {
      await activeServer.close();
    }
  } catch (err) {
    log("server close failed", err);
  }
}

async function main() {
  activeServer = createServer();
  activeTransport = new StdioServerTransport();

  // Restore a previously-authenticated account from the persistent token cache
  // so the server is ready to silently refresh tokens after a restart, with no
  // interactive begin_auth needed (key for unattended scheduled runs).
  if (CLIENT_ID) {
    await restoreAccountFromCache().catch((err) =>
      log("startup restoreAccountFromCache failed", err?.message || String(err))
    );
  }

  process.on("SIGINT", () => {
    shutdown("SIGINT")
      .catch((err) => log("SIGINT shutdown failed", err))
      .finally(() => process.exit(0));
  });

  process.on("SIGTERM", () => {
    shutdown("SIGTERM")
      .catch((err) => log("SIGTERM shutdown failed", err))
      .finally(() => process.exit(0));
  });

  await activeServer.connect(activeTransport);
  log(`${APP_NAME} started over stdio`);
}

function isDirectExecution() {
  const entry = process.argv[1];
  if (!entry) return false;
  return import.meta.url === pathToFileURL(path.resolve(entry)).href;
}

export { createServer, main };

if (isDirectExecution()) {
  main().catch((err) => {
    log("main fatal error", err);
    console.error(err);
    process.exit(1);
  });
}
