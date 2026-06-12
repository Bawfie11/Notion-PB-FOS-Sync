require("dotenv").config();
const { Client } = require("@notionhq/client");

const NOTION_VERSION = process.env.NOTION_VERSION || "2026-03-11";

const notion = new Client({
  auth: process.env.NOTION_TOKEN,
  notionVersion: NOTION_VERSION
});

const PB_API_ROOT = "https://api.productboard.com";
const PB_API_BASE = "https://api.productboard.com/v2";

const PB_TOKEN = process.env.PRODUCTBOARD_TOKEN;
const INITIATIVES_DATA_SOURCE = process.env.NOTION_INITIATIVES_DATA_SOURCE_ID;

const INITIATIVE_NAME_COL = "Initiative Name";
const INITIATIVE_STATUS_COL = "Initiative Status";
const INITIATIVE_TIMELINE_COL = "Timeline";

const TEAMS_WEBHOOK_URL =
  process.env.TEAMS_WEBHOOK_URL || process.env["TEAMS-WEBHOOK-URL"];

function boolEnv(name, defaultValue = false) {
  const raw = process.env[name];
  if (raw === undefined || raw === null || raw === "") return defaultValue;
  return ["true", "1", "yes", "y"].includes(String(raw).toLowerCase());
}

const DRY_RUN = boolEnv("DRY_RUN", true);
const ALLOW_CREATES = boolEnv("ALLOW_CREATES", false);

const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

function validateRequiredEnv() {
  const required = [
    "NOTION_TOKEN",
    "PRODUCTBOARD_TOKEN",
    "NOTION_INITIATIVES_DATA_SOURCE_ID"
  ];

  const missing = required.filter(name => !process.env[name]);
  if (missing.length > 0) {
    throw new Error(`Missing required .env values: ${missing.join(", ")}`);
  }
}

async function sendTeamsAlert(title, message) {
  if (!TEAMS_WEBHOOK_URL) {
    console.warn("TEAMS_WEBHOOK_URL is not set. Skipping Teams alert.");
    return;
  }

  try {
    await fetch(TEAMS_WEBHOOK_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        type: "message",
        attachments: [{
          contentType: "application/vnd.microsoft.card.adaptive",
          content: {
            type: "AdaptiveCard",
            version: "1.2",
            body: [
              { type: "TextBlock", text: `PB -> Notion Initiatives: ${title}`, weight: "Bolder", size: "Medium", wrap: true },
              { type: "TextBlock", text: message, wrap: true }
            ]
          }
        }]
      })
    });
  } catch (e) {
    console.error("Failed to send Teams alert:", e.message);
  }
}

function notionHeaders() {
  return {
    Authorization: `Bearer ${process.env.NOTION_TOKEN}`,
    "Notion-Version": NOTION_VERSION,
    "Content-Type": "application/json"
  };
}

function pbHeaders() {
  return {
    Authorization: `Bearer ${PB_TOKEN}`,
    Accept: "application/json"
  };
}

function makePBUrl(pathOrUrl) {
  if (pathOrUrl.startsWith("http://") || pathOrUrl.startsWith("https://")) {
    return new URL(pathOrUrl);
  }
  if (pathOrUrl.startsWith("/v2/")) {
    return new URL(`${PB_API_ROOT}${pathOrUrl}`);
  }
  const normalizedPath = pathOrUrl.startsWith("/") ? pathOrUrl : `/${pathOrUrl}`;
  return new URL(`${PB_API_BASE}${normalizedPath}`);
}

function buildPBUrl(pathOrUrl, params = {}) {
  const url = makePBUrl(pathOrUrl);
  for (const [key, value] of Object.entries(params)) {
    if (value === undefined || value === null || value === "") continue;
    if (Array.isArray(value)) {
      for (const item of value) url.searchParams.append(key, item);
    } else {
      url.searchParams.set(key, value);
    }
  }
  return url.toString();
}

async function fetchJsonWithRetry(url, options = {}, label = "request") {
  const maxAttempts = 5;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    let res;

    try {
      res = await fetch(url, options);
    } catch (e) {
      if (attempt === maxAttempts) throw new Error(`${label} failed: ${e.message}`);
      await delay(Math.min(30000, 1000 * 2 ** (attempt - 1)));
      continue;
    }

    if (res.status === 429 || res.status >= 500) {
      const retryAfterSeconds = Number(res.headers.get("retry-after"));
      const waitMs = Number.isFinite(retryAfterSeconds) && retryAfterSeconds > 0
        ? retryAfterSeconds * 1000
        : Math.min(30000, 1000 * 2 ** (attempt - 1));

      if (attempt === maxAttempts) {
        const text = await res.text();
        throw new Error(`${label} failed after retries: ${res.status} - ${text}`);
      }

      console.warn(`${label} got ${res.status}. Retrying in ${Math.round(waitMs / 1000)}s...`);
      await delay(waitMs);
      continue;
    }

    if (!res.ok) {
      const text = await res.text();
      throw new Error(`${label} failed: ${res.status} - ${text}`);
    }

    if (res.status === 204) return null;

    const text = await res.text();
    if (!text) return null;

    try {
      return JSON.parse(text);
    } catch {
      throw new Error(`${label} returned non-JSON response: ${text}`);
    }
  }

  throw new Error(`${label} failed unexpectedly`);
}

async function fetchAllPBV2Get(pathOrUrl, params = {}, label = "Productboard list") {
  const results = [];
  let nextUrl = buildPBUrl(pathOrUrl, params);

  while (nextUrl) {
    const json = await fetchJsonWithRetry(
      nextUrl,
      { method: "GET", headers: pbHeaders() },
      label
    );

    results.push(...(json?.data || []));
    nextUrl = json?.links?.next ? buildPBUrl(json.links.next) : null;
  }

  return results;
}

async function fetchAllStandardNotionDataSource(dataSourceId) {
  const results = [];
  let cursor = undefined;

  while (true) {
    const payload = { page_size: 100 };
    if (cursor) payload.start_cursor = cursor;

    const json = await fetchJsonWithRetry(
      `https://api.notion.com/v1/data_sources/${dataSourceId}/query`,
      {
        method: "POST",
        headers: notionHeaders(),
        body: JSON.stringify(payload)
      },
      `Notion query ${dataSourceId}`
    );

    results.push(...(json.results || []));
    if (!json.has_more) break;
    cursor = json.next_cursor;
    await delay(350);
  }

  return results;
}

function sanitizeTitle(value) {
  return String(value || "")
    .replace(/[\u200B-\u200D\uFEFF\u200E\u200F]/g, "")
    .trim();
}

function selectDisplayValue(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  if (Array.isArray(value)) return value.length ? selectDisplayValue(value[0]) : null;

  if (typeof value === "object") {
    if (typeof value.name === "string") return value.name;
    if (typeof value.label === "string") return value.label;
    if (typeof value.displayName === "string") return value.displayName;
    if ("value" in value) return selectDisplayValue(value.value);
  }

  return null;
}

async function validateInitiativesNotionSchema() {
  const json = await fetchJsonWithRetry(
    `https://api.notion.com/v1/data_sources/${INITIATIVES_DATA_SOURCE}`,
    { method: "GET", headers: notionHeaders() },
    "Notion initiatives schema validation"
  );

  const existingColumns = Object.keys(json.properties || {});

  const requiredColumns = [
    INITIATIVE_NAME_COL,
    INITIATIVE_STATUS_COL,
    INITIATIVE_TIMELINE_COL,
    "PBID"
  ];

  const missingColumns = requiredColumns.filter(col => !existingColumns.includes(col));

  if (missingColumns.length > 0) {
    const msg = `Initiatives Notion DB is missing columns:\n\n${missingColumns.join("\n")}\n\nSync aborted.`;
    console.error("Initiatives schema validation failed:", missingColumns);
    await sendTeamsAlert("Schema Mismatch - Sync Aborted", msg);
    throw new Error("Initiatives schema validation failed.");
  }

  const expectedTypes = {
    [INITIATIVE_NAME_COL]: "title",
    [INITIATIVE_STATUS_COL]: "select",
    [INITIATIVE_TIMELINE_COL]: "date",
    "PBID": "rich_text"
  };

  const typeMismatches = Object.entries(expectedTypes)
    .filter(([col, expectedType]) => json.properties[col]?.type !== expectedType)
    .map(([col, expectedType]) => `${col}: expected ${expectedType}, found ${json.properties[col]?.type || "missing"}`);

  if (typeMismatches.length > 0) {
    const msg = `Initiatives Notion DB column types are wrong:\n\n${typeMismatches.join("\n")}\n\nSync aborted.`;
    console.error("Initiatives schema type validation failed:", typeMismatches);
    await sendTeamsAlert("Schema Type Mismatch - Sync Aborted", msg);
    throw new Error("Initiatives schema type validation failed.");
  }

  console.log("Initiatives Notion schema validation passed.");
}

function buildInitiativeNotionProperties(entity) {
  const fields = entity.fields || {};
  const pbId = entity.id;
  const name = sanitizeTitle(selectDisplayValue(fields.name) || "");

  const timeframe = fields.timeframe?.value || fields.timeframe || null;
  let start = timeframe?.startDate || null;
  let end = timeframe?.endDate || null;

  if (start === "none") start = null;
  if (end === "none") end = null;

  let dateObj = null;
  if (start || end) {
    dateObj = { start: start || end };
    if (start && end && start !== end) dateObj.end = end;
  }

  const status = fields.status?.name || selectDisplayValue(fields.status);

  const properties = {
    [INITIATIVE_NAME_COL]: { title: [{ text: { content: name } }] },
    PBID: { rich_text: [{ text: { content: pbId } }] },
    [INITIATIVE_TIMELINE_COL]: { date: dateObj },
    [INITIATIVE_STATUS_COL]: status
      ? { select: { name: status } }
      : { select: null }
  };

  return { pbId, name, properties };
}

function isInitiativeDifferent(newProps, oldProps, initiativeName) {
  const logs = [];

  const newTitle = newProps[INITIATIVE_NAME_COL]?.title?.[0]?.text?.content || "";
  const oldTitle = oldProps[INITIATIVE_NAME_COL]?.title?.[0]?.plain_text || "";
  if (newTitle !== oldTitle) logs.push(`Name: "${oldTitle}" -> "${newTitle}"`);

  const newStart = newProps[INITIATIVE_TIMELINE_COL]?.date?.start || null;
  const newEnd = newProps[INITIATIVE_TIMELINE_COL]?.date?.end || null;
  const oldStart = oldProps[INITIATIVE_TIMELINE_COL]?.date?.start || null;
  const oldEnd = oldProps[INITIATIVE_TIMELINE_COL]?.date?.end || null;
  if (newStart !== oldStart || newEnd !== oldEnd) {
    logs.push(`Timeline: [${oldStart} to ${oldEnd}] -> [${newStart} to ${newEnd}]`);
  }

  const newStatus = newProps[INITIATIVE_STATUS_COL]?.select?.name || null;
  const oldStatus = oldProps[INITIATIVE_STATUS_COL]?.select?.name || null;
  if (newStatus !== oldStatus) logs.push(`Status: "${oldStatus}" -> "${newStatus}"`);

  if (logs.length > 0) {
    console.log(`\nINITIATIVE DIFF FOR "${initiativeName}":`);
    logs.forEach(l => console.log(`   - ${l}`));
    return true;
  }

  return false;
}

async function main() {
  console.log("Starting Productboard -> Notion Initiatives sync...");

  validateRequiredEnv();

  console.log(`DRY_RUN=${DRY_RUN}`);
  console.log(`ALLOW_CREATES=${ALLOW_CREATES}`);

  await validateInitiativesNotionSchema();

  const pbInitiatives = await fetchAllPBV2Get(
    "/entities",
    {
      "type[]": ["initiative"],
      "fields[]": ["name", "status", "timeframe"]
    },
    "PB initiatives"
  );

  console.log(`Fetched ${pbInitiatives.length} PB initiatives.`);

  if (pbInitiatives.length === 0) {
    console.log("--- SYNC COMPLETE ---");
    console.log("No initiatives returned from Productboard.");
    return;
  }

  const existingNotionInitiatives = await fetchAllStandardNotionDataSource(INITIATIVES_DATA_SOURCE);
  const existingByPbId = {};

  existingNotionInitiatives.forEach(page => {
    const pbId = page.properties.PBID?.rich_text?.[0]?.plain_text;
    if (pbId) existingByPbId[pbId] = page;
  });

  let created = 0;
  let updated = 0;
  let skipped = 0;
  let errors = 0;
  let wouldCreate = 0;
  let wouldUpdate = 0;

  for (const entity of pbInitiatives) {
    const { pbId, name, properties } = buildInitiativeNotionProperties(entity);
    const existingPage = existingByPbId[pbId];

    try {
      if (existingPage) {
        if (isInitiativeDifferent(properties, existingPage.properties, name)) {
          if (DRY_RUN) {
            console.log(`DRY RUN - would update: ${name}`);
            wouldUpdate++;
          } else {
            await delay(350);
            await notion.pages.update({
              page_id: existingPage.id,
              properties
            });
            console.log(`Updated: ${name}`);
            updated++;
          }
        } else {
          skipped++;
        }
      } else {
        if (!ALLOW_CREATES) {
          console.warn(`Create blocked by ALLOW_CREATES=false: ${name} (${pbId})`);
          wouldCreate++;
          continue;
        }

        if (DRY_RUN) {
          console.log(`DRY RUN - would create: ${name} (${pbId})`);
          wouldCreate++;
          continue;
        }

        await delay(350);
        await notion.pages.create({
          parent: { data_source_id: INITIATIVES_DATA_SOURCE },
          properties
        });
        console.log(`Created: ${name}`);
        created++;
      }
    } catch (err) {
      console.error(`\nINITIATIVE SYNC FAILED: "${name}" (${pbId})`);
      console.error("Payload Notion rejected:");
      console.error(JSON.stringify(properties, null, 2));
      console.error(`Error: ${err.message}\n`);
      errors++;

      if (errors <= 3) {
        await sendTeamsAlert(
          "Initiative Sync Failed",
          `Initiative: "${name}"\nPBID: ${pbId}\nError: ${err.message}`
        );
      }
    }
  }

  console.log("--- INITIATIVES SYNC COMPLETE ---");
  console.log(`Created: ${created} | Updated: ${updated} | Skipped: ${skipped} | Errors: ${errors}`);

  if (DRY_RUN || !ALLOW_CREATES) {
    console.log(`Would create: ${wouldCreate} | Would update: ${wouldUpdate}`);
  }

  if (errors > 0) {
    await sendTeamsAlert(
      "Initiatives Sync Completed With Errors",
      `${errors} initiative(s) failed.\nCreated: ${created} | Updated: ${updated} | Skipped: ${skipped}`
    );
  }
}

main().catch(async (err) => {
  console.error("Fatal initiatives sync error:", err);
  await sendTeamsAlert(
    "Initiatives Sync Crashed",
    `The initiatives sync crashed.\n\nError: ${err.message}`
  );
  process.exit(1);
});
