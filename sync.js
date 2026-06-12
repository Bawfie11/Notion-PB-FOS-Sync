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
const ROADMAP_DATA_SOURCE = process.env.NOTION_ROADMAP_DATA_SOURCE_ID;
const PRODUCTS_DATA_SOURCE = process.env.NOTION_PRODUCTS_DATA_SOURCE_ID;
const INITIATIVES_DATA_SOURCE = process.env.NOTION_INITIATIVES_DATA_SOURCE_ID;

const ROADMAP_COMPONENT_TEXT_COL = "Component";
const ROADMAP_STATUS_COL = "PB Status";
const ROADMAP_HEALTH_DATE_COL = "Health Updated At";
const TSHIRT_SIZE_COL = "T-Shirt Size";
const DESIGN_STATUS_COL = "Design Status";
const DESIGNER_COL = "Designer";
const PI_FEATURE_COL = "P&I Feature";

const INITIATIVE_NAME_COL = "Initiative Name";
const INITIATIVE_STATUS_COL = "Initiative Status";
const INITIATIVE_TIMELINE_COL = "Timeline";

const TEAMS_WEBHOOK_URL =
  process.env.TEAMS_WEBHOOK_URL || process.env["TEAMS-WEBHOOK-URL"];

const APP_APPROVAL_MAP = {
  "None": "No App Approval Required",
  "Companion": "Companion App Approval Req",
  "Kiosk": "Kiosk App Approval Req"
};

const FEATURE_FLAG_MAP = {
  "Yes": "Has Feature Flag",
  "No": "No Feature Flag"
};

const HEALTH_MAP = {
  notSet: "Not started",
  onTrack: "On track",
  atRisk: "At risk",
  offTrack: "Off track"
};

const CUSTOM_FIELD_NAMES = {
  tier: "Tier",
  appApproval: "App approval",
  featureFlag: "Feature Flag",
  aiFeature: "AI Feature",
  tshirtSize: TSHIRT_SIZE_COL,
  designStatus: DESIGN_STATUS_COL,
  designer: DESIGNER_COL,
  piFeature: PI_FEATURE_COL
};

function boolEnv(name, defaultValue = false) {
  const raw = process.env[name];

  if (raw === undefined || raw === null || raw === "") {
    return defaultValue;
  }

  return ["true", "1", "yes", "y"].includes(String(raw).toLowerCase());
}

const DRY_RUN = boolEnv("DRY_RUN", true);
const ALLOW_CREATES = boolEnv("ALLOW_CREATES", false);

// If true, the script crashes when Designer cannot be mapped.
// If false, the script logs a warning and skips Designer for that feature.
const DESIGNER_SYNC_STRICT = boolEnv("DESIGNER_SYNC_STRICT", false);

// Manual mode:
// true  = fetch all PB feature entities. Best for a full migration validation.
// false = fetch only features updated since your configured lookback window.
const PB_FETCH_ALL_FEATURES = boolEnv("PB_FETCH_ALL_FEATURES", false);

const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

function validateRequiredEnv() {
  const required = [
    "NOTION_TOKEN",
    "PRODUCTBOARD_TOKEN",
    "NOTION_ROADMAP_DATA_SOURCE_ID",
    "NOTION_PRODUCTS_DATA_SOURCE_ID"
  ];

  const missing = required.filter(name => !process.env[name]);

  if (missing.length > 0) {
    throw new Error(`Missing required .env values: ${missing.join(", ")}`);
  }
}

function calculateManualCutoffISO() {
  const exactSince = (process.env.PB_UPDATED_SINCE_ISO || "").trim();

  if (exactSince) {
    const parsed = new Date(exactSince);

    if (Number.isNaN(parsed.getTime())) {
      throw new Error(`PB_UPDATED_SINCE_ISO is not a valid date/time: ${exactSince}`);
    }

    return parsed.toISOString();
  }

  const lookbackHoursRaw = (process.env.PB_UPDATED_LOOKBACK_HOURS || "").trim();

  if (lookbackHoursRaw) {
    const hours = Number(lookbackHoursRaw);

    if (!Number.isFinite(hours) || hours <= 0) {
      throw new Error(`PB_UPDATED_LOOKBACK_HOURS must be a positive number. Received: ${lookbackHoursRaw}`);
    }

    return new Date(Date.now() - hours * 60 * 60 * 1000).toISOString();
  }

  const lookbackDays = Number(process.env.PB_UPDATED_LOOKBACK_DAYS || 30);

  if (!Number.isFinite(lookbackDays) || lookbackDays <= 0) {
    throw new Error(`PB_UPDATED_LOOKBACK_DAYS must be a positive number. Received: ${process.env.PB_UPDATED_LOOKBACK_DAYS}`);
  }

  return new Date(Date.now() - lookbackDays * 24 * 60 * 60 * 1000).toISOString();
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
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: {
              type: "AdaptiveCard",
              version: "1.2",
              body: [
                {
                  type: "TextBlock",
                  text: `PB -> Notion Sync: ${title}`,
                  weight: "Bolder",
                  size: "Medium",
                  wrap: true
                },
                {
                  type: "TextBlock",
                  text: message,
                  wrap: true
                }
              ]
            }
          }
        ]
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

function pbHeaders(method = "GET") {
  const headers = {
    Authorization: `Bearer ${PB_TOKEN}`,
    Accept: "application/json"
  };

  if (method !== "GET") {
    headers["Content-Type"] = "application/json";
  }

  return headers;
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
      for (const item of value) {
        url.searchParams.append(key, item);
      }
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
      if (attempt === maxAttempts) {
        throw new Error(`${label} failed: ${e.message}`);
      }

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

async function pbGet(pathOrUrl, params = {}, label = "Productboard GET") {
  const url = buildPBUrl(pathOrUrl, params);

  return fetchJsonWithRetry(
    url,
    { method: "GET", headers: pbHeaders("GET") },
    label
  );
}

async function fetchAllPBV2Get(pathOrUrl, params = {}, label = "Productboard list") {
  const results = [];
  let nextUrl = buildPBUrl(pathOrUrl, params);

  while (nextUrl) {
    const json = await fetchJsonWithRetry(
      nextUrl,
      { method: "GET", headers: pbHeaders("GET") },
      label
    );

    results.push(...(json?.data || []));

    nextUrl = json?.links?.next
      ? buildPBUrl(json.links.next)
      : null;
  }

  return results;
}

async function fetchAllPBV2Search(body, label = "Productboard search") {
  const results = [];
  let nextUrl = buildPBUrl("/entities/search");

  while (nextUrl) {
    const json = await fetchJsonWithRetry(
      nextUrl,
      {
        method: "POST",
        headers: pbHeaders("POST"),
        body: JSON.stringify(body)
      },
      label
    );

    results.push(...(json?.data || []));

    nextUrl = json?.links?.next
      ? buildPBUrl(json.links.next)
      : null;
  }

  return results;
}

async function validateNotionSchema() {
  const json = await fetchJsonWithRetry(
    `https://api.notion.com/v1/data_sources/${ROADMAP_DATA_SOURCE}`,
    { method: "GET", headers: notionHeaders() },
    "Notion schema validation"
  );

  const existingColumns = Object.keys(json.properties || {});

  const requiredColumns = [
    "Feature",
    "PBID",
    "Development Timeline",
    "Module",
    ROADMAP_COMPONENT_TEXT_COL,
    "GTM Tier",
    "App approval",
    "Feature Flag",
    "AI Feature",
    TSHIRT_SIZE_COL,
    DESIGN_STATUS_COL,
    DESIGNER_COL,
    PI_FEATURE_COL,
    "PB_Health",
    "Latest Health Comment",
    ROADMAP_HEALTH_DATE_COL,
    ROADMAP_STATUS_COL
  ];

  const missingColumns = requiredColumns.filter(col => !existingColumns.includes(col));

  if (missingColumns.length > 0) {
    const msg = `The following Notion columns are missing or renamed:\n\n${missingColumns.join("\n")}\n\nSync aborted.`;
    console.error("Schema validation failed:", missingColumns);
    await sendTeamsAlert("Schema Mismatch - Sync Aborted", msg);
    throw new Error("Schema validation failed. Aborting sync.");
  }

  const expectedTypes = {
    [TSHIRT_SIZE_COL]: "select",
    [DESIGN_STATUS_COL]: "select",
    [DESIGNER_COL]: "people"
  };

  const typeMismatches = Object.entries(expectedTypes)
    .filter(([col, expectedType]) => json.properties[col]?.type !== expectedType)
    .map(([col, expectedType]) => {
      const actualType = json.properties[col]?.type || "missing";
      return `${col}: expected ${expectedType}, found ${actualType}`;
    });

  if (typeMismatches.length > 0) {
    const msg = `The following Notion columns have the wrong property type:\n\n${typeMismatches.join("\n")}\n\nSync aborted.`;
    console.error("Schema type validation failed:", typeMismatches);
    await sendTeamsAlert("Schema Type Mismatch - Sync Aborted", msg);
    throw new Error("Schema type validation failed. Aborting sync.");
  }

  console.log("Notion schema validation passed.");
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

function htmlToPlainText(value) {
  return String(value || "")
    .replace(/<[^>]+>/g, "")
    .trim();
}

function selectDisplayValue(value) {
  if (value === null || value === undefined || value === "") return null;

  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);

  if (Array.isArray(value)) {
    return value.length ? selectDisplayValue(value[0]) : null;
  }

  if (typeof value === "object") {
    if (typeof value.name === "string") return value.name;
    if (typeof value.label === "string") return value.label;
    if (typeof value.displayName === "string") return value.displayName;
    if (typeof value.email === "string") return value.email;
    if (typeof value.fields?.name === "string") return value.fields.name;
    if (typeof value.fields?.email === "string") return value.fields.email;
    if ("value" in value) return selectDisplayValue(value.value);
  }

  return null;
}

function normalizeEmail(value) {
  if (typeof value !== "string") return null;

  const email = value.trim().toLowerCase();

  if (!email) return null;
  if (email === "[redacted]") return null;
  if (!email.includes("@")) return null;

  return email;
}

function isLikelyUuid(value) {
  return typeof value === "string" &&
    /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value);
}

async function fetchAllNotionUsers() {
  const results = [];
  let cursor = undefined;

  while (true) {
    const url = new URL("https://api.notion.com/v1/users");
    url.searchParams.set("page_size", "100");

    if (cursor) {
      url.searchParams.set("start_cursor", cursor);
    }

    const json = await fetchJsonWithRetry(
      url.toString(),
      {
        method: "GET",
        headers: notionHeaders()
      },
      "Notion users"
    );

    results.push(...(json.results || []));

    if (!json.has_more) break;

    cursor = json.next_cursor;
    await delay(350);
  }

  return results;
}

function buildNotionUserMapByEmail(users) {
  const map = {};

  for (const user of users) {
    const email = normalizeEmail(user?.person?.email);

    if (email && user.id) {
      map[email] = user.id;
    }
  }

  return map;
}

function isEmptyPBFieldValue(value) {
  if (value === null || value === undefined || value === "") return true;

  if (Array.isArray(value)) {
    return value.length === 0;
  }

  if (typeof value === "object" && Object.prototype.hasOwnProperty.call(value, "value")) {
    return isEmptyPBFieldValue(value.value);
  }

  return false;
}

function extractEmailsFromPBValue(value) {
  const emails = new Set();

  function walk(v) {
    if (v === null || v === undefined) return;

    const directEmail = normalizeEmail(v);

    if (directEmail) {
      emails.add(directEmail);
      return;
    }

    if (Array.isArray(v)) {
      v.forEach(walk);
      return;
    }

    if (typeof v === "object") {
      for (const child of Object.values(v)) {
        walk(child);
      }
    }
  }

  walk(value);

  return [...emails];
}

function extractPBUserIdsFromDesignerValue(value) {
  const ids = new Set();

  function walk(v) {
    if (v === null || v === undefined) return;

    if (Array.isArray(v)) {
      v.forEach(walk);
      return;
    }

    if (typeof v !== "object") return;

    if (isLikelyUuid(v.id)) {
      ids.add(v.id);
    }

    if (isLikelyUuid(v.userId)) {
      ids.add(v.userId);
    }

    if (isLikelyUuid(v.entityId)) {
      ids.add(v.entityId);
    }

    for (const child of Object.values(v)) {
      walk(child);
    }
  }

  walk(value);

  return [...ids];
}

async function resolvePBDesignerEmails(pbDesignerValue, pbUserEmailCache) {
  const emails = new Set(extractEmailsFromPBValue(pbDesignerValue));
  const userIds = extractPBUserIdsFromDesignerValue(pbDesignerValue);

  for (const userId of userIds) {
    if (!Object.prototype.hasOwnProperty.call(pbUserEmailCache, userId)) {
      try {
        const userEntity = await fetchEntityById(userId, ["name", "email"]);
        const userEmails = extractEmailsFromPBValue(userEntity);

        pbUserEmailCache[userId] = userEmails[0] || null;
      } catch (e) {
        console.warn(`Could not fetch Productboard user ${userId} for Designer mapping: ${e.message}`);
        pbUserEmailCache[userId] = null;
      }
    }

    if (pbUserEmailCache[userId]) {
      emails.add(pbUserEmailCache[userId]);
    }
  }

  return [...emails];
}

async function buildDesignerPeopleProperty(pbDesignerValue, context, featureName) {
  const { notionUserByEmail, pbUserEmailCache } = context;

  // If PB Designer is truly empty, clear Notion Designer.
  if (isEmptyPBFieldValue(pbDesignerValue)) {
    return { people: [] };
  }

  const pbEmails = await resolvePBDesignerEmails(pbDesignerValue, pbUserEmailCache);

  if (pbEmails.length === 0) {
    const msg =
      `Could not resolve Productboard Designer email for "${featureName}". ` +
      "Designer will be skipped for this feature. " +
      "This usually means the Productboard token is missing users:pii:read, " +
      "or the returned Designer value shape needs inspection.";

    if (DESIGNER_SYNC_STRICT) {
      throw new Error(msg);
    }

    console.warn(`WARNING: ${msg}`);
    return null;
  }

  const missingEmails = pbEmails.filter(email => !notionUserByEmail[email]);

  if (missingEmails.length > 0) {
    const msg =
      `Could not find matching Notion user(s) for Designer on "${featureName}": ${missingEmails.join(", ")}. ` +
      "Designer will be skipped for this feature.";

    if (DESIGNER_SYNC_STRICT) {
      throw new Error(msg);
    }

    console.warn(`WARNING: ${msg}`);
    return null;
  }

  return {
    people: pbEmails.map(email => ({
      object: "user",
      id: notionUserByEmail[email]
    }))
  };
}

function peopleIdsString(peopleProperty) {
  return (peopleProperty?.people || [])
    .map(person => person.id)
    .filter(Boolean)
    .sort()
    .join(",");
}

function getRelationships(entity) {
  const rels = entity?.relationships;

  if (Array.isArray(rels)) return rels;
  if (Array.isArray(rels?.data)) return rels.data;

  return [];
}

function getParentRef(entity) {
  return getRelationships(entity).find(r => r.type === "parent")?.target || null;
}

function getEntityName(entity) {
  return sanitizeTitle(selectDisplayValue(entity?.fields?.name) || "");
}

function normalizeFieldsList(config) {
  if (!config) return [];

  if (Array.isArray(config.fields)) {
    return config.fields.map((field, index) => [field.id || field.apiName || String(index), field]);
  }

  if (config.fields && typeof config.fields === "object") {
    return Object.entries(config.fields);
  }

  if (Array.isArray(config.data?.fields)) {
    return config.data.fields.map((field, index) => [field.id || field.apiName || String(index), field]);
  }

  return [];
}

function findFieldIdByDisplayName(config, displayName) {
  const fields = normalizeFieldsList(config);

  for (const [key, field] of fields) {
    const names = [
      field?.name,
      field?.displayName,
      field?.label,
      field?.apiName
    ].filter(Boolean);

    if (names.includes(displayName)) {
      return field.id || field.apiName || key;
    }
  }

  return null;
}

async function getFeatureConfigAndCustomFieldIds() {
  const envIds = {
    tier: (process.env.PB_CUSTOM_FIELD_TIER_ID || "").trim(),
    appApproval: (process.env.PB_CUSTOM_FIELD_APP_APPROVAL_ID || "").trim(),
    featureFlag: (process.env.PB_CUSTOM_FIELD_FEATURE_FLAG_ID || "").trim(),
    aiFeature: (process.env.PB_CUSTOM_FIELD_AI_FEATURE_ID || "").trim(),
    tshirtSize: (
      process.env.PB_CUSTOM_FIELD_TSHIRT_SIZE_ID ||
      process.env.PB_CUSTOM_FIELD_T_SHIRT_SIZE_ID ||
      ""
    ).trim(),
    designStatus: (process.env.PB_CUSTOM_FIELD_DESIGN_STATUS_ID || "").trim(),
    designer: (process.env.PB_CUSTOM_FIELD_DESIGNER_ID || "").trim(),
    piFeature: (process.env.PB_CUSTOM_FIELD_PI_FEATURE_ID || "").trim()
  };

  const allEnvIdsPresent = Object.values(envIds).every(Boolean);

  if (allEnvIdsPresent) {
    console.log("Using Productboard custom field IDs from .env.");
    return { config: null, fieldIds: envIds };
  }

  const json = await pbGet(
    "/entities/configurations/feature",
    {},
    "PB feature configuration"
  );

  const config = json.data || json;

  const fieldIds = {
    tier: envIds.tier || findFieldIdByDisplayName(config, CUSTOM_FIELD_NAMES.tier),
    appApproval: envIds.appApproval || findFieldIdByDisplayName(config, CUSTOM_FIELD_NAMES.appApproval),
    featureFlag: envIds.featureFlag || findFieldIdByDisplayName(config, CUSTOM_FIELD_NAMES.featureFlag),
    aiFeature: envIds.aiFeature || findFieldIdByDisplayName(config, CUSTOM_FIELD_NAMES.aiFeature),
    tshirtSize: envIds.tshirtSize || findFieldIdByDisplayName(config, CUSTOM_FIELD_NAMES.tshirtSize),
    designStatus: envIds.designStatus || findFieldIdByDisplayName(config, CUSTOM_FIELD_NAMES.designStatus),
    designer: envIds.designer || findFieldIdByDisplayName(config, CUSTOM_FIELD_NAMES.designer),
    piFeature: envIds.piFeature || findFieldIdByDisplayName(config, CUSTOM_FIELD_NAMES.piFeature)
  };

  const missing = Object.entries(fieldIds)
    .filter(([, id]) => !id)
    .map(([key]) => `${key} (${CUSTOM_FIELD_NAMES[key]})`);

  if (missing.length > 0) {
    throw new Error(`Could not find these Productboard custom fields on feature configuration: ${missing.join(", ")}`);
  }

  console.log("Discovered Productboard custom field IDs from v2 configuration:");
  console.log(fieldIds);

  return { config, fieldIds };
}

async function fetchParentRef(entityId) {
  const relationships = await fetchAllPBV2Get(
    `/entities/${entityId}/relationships`,
    { type: "parent" },
    `PB parent relationship ${entityId}`
  );

  return relationships.find(r => r.type === "parent")?.target || null;
}

async function fetchEntityById(id, fields = ["name"]) {
  const json = await pbGet(
    `/entities/${id}`,
    { "fields[]": fields },
    `PB entity ${id}`
  );

  return json.data;
}

async function getProductComponentMaps() {
  const entities = await fetchAllPBV2Get(
    "/entities",
    {
      "type[]": ["product", "component"],
      "fields[]": ["name"]
    },
    "PB products/components"
  );

  const productMap = {};
  const componentMap = {};

  for (const entity of entities) {
    const name = getEntityName(entity);
    const parent = getParentRef(entity);

    if (entity.type === "product") {
      productMap[entity.id] = name;
    }

    if (entity.type === "component") {
      componentMap[entity.id] = {
        name,
        parentProduct: parent?.type === "product" ? parent.id : null,
        parentComponent: parent?.type === "component" ? parent.id : null
      };
    }
  }

  return { productMap, componentMap };
}

async function ensureProductName(productId, productMap) {
  if (!productId || productMap[productId]) return;

  const productEntity = await fetchEntityById(productId, ["name"]);
  productMap[productId] = getEntityName(productEntity);
}

async function ensureComponentChain(componentId, productMap, componentMap) {
  let currentComponentId = componentId;
  let safetyCounter = 0;

  while (currentComponentId && safetyCounter < 10) {
    let component = componentMap[currentComponentId];

    if (!component) {
      const componentEntity = await fetchEntityById(currentComponentId, ["name"]);
      const parent = getParentRef(componentEntity) || await fetchParentRef(currentComponentId);

      component = {
        name: getEntityName(componentEntity),
        parentProduct: parent?.type === "product" ? parent.id : null,
        parentComponent: parent?.type === "component" ? parent.id : null
      };

      componentMap[currentComponentId] = component;
    } else if (!component.parentProduct && !component.parentComponent) {
      const parent = await fetchParentRef(currentComponentId);

      if (parent?.type === "product") {
        component.parentProduct = parent.id;
      }

      if (parent?.type === "component") {
        component.parentComponent = parent.id;
      }
    }

    if (component.parentProduct) {
      await ensureProductName(component.parentProduct, productMap);
      break;
    }

    currentComponentId = component.parentComponent;
    safetyCounter++;
  }
}

async function fetchPBFeatures(returnFields, cutoffISO) {
  if (PB_FETCH_ALL_FEATURES) {
    return fetchAllPBV2Get(
      "/entities",
      {
        "type[]": ["feature"],
        "fields[]": returnFields
      },
      "PB all feature entities"
    );
  }

  const body = {
    data: {
      filter: {
        type: ["feature"],
        updatedAt: {
          from: cutoffISO
        }
      },
      return: {
        fields: returnFields
      }
    }
  };

  return fetchAllPBV2Search(body, "PB updated feature search");
}

function toIsoMinute(value) {
  if (!value) return null;

  const date = new Date(value);

  if (Number.isNaN(date.getTime())) {
    return null;
  }

  return date.toISOString().substring(0, 16);
}

function isDifferent(newProps, oldProps, featureName) {
  const logs = [];

  const newTitle = newProps["Feature"]?.title?.[0]?.text?.content || "";
  const oldTitle = oldProps["Feature"]?.title?.[0]?.plain_text || "";
  if (newTitle !== oldTitle) logs.push(`Title: "${oldTitle}" -> "${newTitle}"`);

  const newStart = newProps["Development Timeline"]?.date?.start || null;
  const newEnd = newProps["Development Timeline"]?.date?.end || null;
  const oldStart = oldProps["Development Timeline"]?.date?.start || null;
  const oldEnd = oldProps["Development Timeline"]?.date?.end || null;

  if (newStart !== oldStart || newEnd !== oldEnd) {
    logs.push(`Timeframe: [${oldStart} to ${oldEnd}] -> [${newStart} to ${newEnd}]`);
  }

  const newProd = newProps["Module"]?.relation?.[0]?.id || null;
  const oldProd = oldProps["Module"]?.relation?.[0]?.id || null;
  if (newProd !== oldProd) logs.push("Module Relation ID mismatch");

  const newComp = newProps[ROADMAP_COMPONENT_TEXT_COL]?.rich_text?.[0]?.text?.content || "";
  const oldComp = oldProps[ROADMAP_COMPONENT_TEXT_COL]?.rich_text?.[0]?.plain_text || "";
  if (newComp !== oldComp) logs.push(`Component Text: "${oldComp}" -> "${newComp}"`);

  const newTier = newProps["GTM Tier"]?.select?.name || null;
  const oldTier = oldProps["GTM Tier"]?.select?.name || null;
  if (newTier !== oldTier) logs.push(`GTM Tier: "${oldTier}" -> "${newTier}"`);

  const newAppApproval = newProps["App approval"]?.select?.name || null;
  const oldAppApproval = oldProps["App approval"]?.select?.name || null;
  if (newAppApproval !== oldAppApproval) logs.push(`App approval: "${oldAppApproval}" -> "${newAppApproval}"`);

  const newFeatureFlag = newProps["Feature Flag"]?.select?.name || null;
  const oldFeatureFlag = oldProps["Feature Flag"]?.select?.name || null;
  if (newFeatureFlag !== oldFeatureFlag) logs.push(`Feature Flag: "${oldFeatureFlag}" -> "${newFeatureFlag}"`);

  const newAI = newProps["AI Feature"]?.select?.name || null;
  const oldAI = oldProps["AI Feature"]?.select?.name || null;
  if (newAI !== oldAI) logs.push(`AI Feature: "${oldAI}" -> "${newAI}"`);

  const newTshirtSize = newProps[TSHIRT_SIZE_COL]?.select?.name || null;
  const oldTshirtSize = oldProps[TSHIRT_SIZE_COL]?.select?.name || null;
  if (newTshirtSize !== oldTshirtSize) {
    logs.push(`${TSHIRT_SIZE_COL}: "${oldTshirtSize}" -> "${newTshirtSize}"`);
  }

  const newDesignStatus = newProps[DESIGN_STATUS_COL]?.select?.name || null;
  const oldDesignStatus = oldProps[DESIGN_STATUS_COL]?.select?.name || null;
  if (newDesignStatus !== oldDesignStatus) {
    logs.push(`${DESIGN_STATUS_COL}: "${oldDesignStatus}" -> "${newDesignStatus}"`);
  }

  const newPiFeature = newProps[PI_FEATURE_COL]?.select?.name || null;
  const oldPiFeature = oldProps[PI_FEATURE_COL]?.select?.name || null;
  if (newPiFeature !== oldPiFeature) {
    logs.push(`${PI_FEATURE_COL}: "${oldPiFeature}" -> "${newPiFeature}"`);
  }

  if (Object.prototype.hasOwnProperty.call(newProps, DESIGNER_COL)) {
    const newDesignerIds = peopleIdsString(newProps[DESIGNER_COL]);
    const oldDesignerIds = peopleIdsString(oldProps[DESIGNER_COL]);

    if (newDesignerIds !== oldDesignerIds) {
      logs.push(`${DESIGNER_COL}: Notion user IDs changed`);
    }
  }

  const newHealth = newProps["PB_Health"]?.status?.name || null;
  const oldHealth = oldProps["PB_Health"]?.status?.name || null;
  if (newHealth !== oldHealth) logs.push(`PB_Health: "${oldHealth}" -> "${newHealth}"`);

  const newComment = newProps["Latest Health Comment"]?.rich_text?.[0]?.text?.content || "";
  const oldComment = oldProps["Latest Health Comment"]?.rich_text?.[0]?.plain_text || "";
  if (newComment !== oldComment) logs.push("Comment mismatch");

  const newStatus = newProps[ROADMAP_STATUS_COL]?.status?.name || null;
  const oldStatus = oldProps[ROADMAP_STATUS_COL]?.status?.name || null;
  if (newStatus !== oldStatus) logs.push(`Status: "${oldStatus}" -> "${newStatus}"`);

  const newHealthDate = newProps[ROADMAP_HEALTH_DATE_COL]?.date?.start || null;
  const oldHealthDate = oldProps[ROADMAP_HEALTH_DATE_COL]?.date?.start || null;

  const normNewDate = toIsoMinute(newHealthDate);
  const normOldDate = toIsoMinute(oldHealthDate);

  if (normNewDate !== normOldDate) {
    logs.push(`Health Date: "${normOldDate}" -> "${normNewDate}"`);
  }

  if (logs.length > 0) {
    console.log(`\nDIFF CAUGHT FOR "${featureName}":`);
    logs.forEach(l => console.log(`   - ${l}`));
    return true;
  }

  return false;
}

async function buildNotionPropertiesFromPBFeature(entity, context) {
  const { notionProductMap, productMap, componentMap, fieldIds } = context;

  const fields = entity.fields || {};
  const pbId = entity.id;
  const name = sanitizeTitle(selectDisplayValue(fields.name) || "");

  let parent = getParentRef(entity);

  if (!parent) {
    parent = await fetchParentRef(entity.id);
  }

  const immediateProductId = parent?.type === "product" ? parent.id : null;
  const immediateComponentId = parent?.type === "component" ? parent.id : null;

  if (immediateComponentId) {
    await ensureComponentChain(immediateComponentId, productMap, componentMap);
  }

  const pbComponentName = immediateComponentId
    ? (componentMap[immediateComponentId]?.name || "")
    : "";

  let resolvedProductId = immediateProductId;

  if (!resolvedProductId && immediateComponentId) {
    let currentCompId = immediateComponentId;
    let safetyCounter = 0;

    while (currentCompId && componentMap[currentCompId] && safetyCounter < 10) {
      const compData = componentMap[currentCompId];

      if (compData.parentProduct) {
        resolvedProductId = compData.parentProduct;
        break;
      }

      currentCompId = compData.parentComponent;
      safetyCounter++;
    }
  }

  if (resolvedProductId) {
    await ensureProductName(resolvedProductId, productMap);
  }

  const pbProductName = resolvedProductId ? (productMap[resolvedProductId] || "") : "";

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

  const properties = {
    Feature: { title: [{ text: { content: name } }] },
    PBID: { rich_text: [{ text: { content: pbId } }] },
    "Development Timeline": { date: dateObj }
  };

  if (pbProductName && notionProductMap[pbProductName]) {
    properties["Module"] = { relation: [{ id: notionProductMap[pbProductName] }] };
  } else {
    properties["Module"] = { relation: [] };
  }

  if (pbComponentName) {
    properties[ROADMAP_COMPONENT_TEXT_COL] = {
      rich_text: [{ text: { content: pbComponentName } }]
    };
  } else {
    properties[ROADMAP_COMPONENT_TEXT_COL] = { rich_text: [] };
  }

  const tierVal = selectDisplayValue(fields[fieldIds.tier]);
  properties["GTM Tier"] = tierVal
    ? { select: { name: tierVal } }
    : { select: null };

  const rawAppApprovalVal = selectDisplayValue(fields[fieldIds.appApproval]);
  const appApprovalVal = rawAppApprovalVal
    ? (APP_APPROVAL_MAP[rawAppApprovalVal] || rawAppApprovalVal)
    : null;

  properties["App approval"] = appApprovalVal
    ? { select: { name: appApprovalVal } }
    : { select: null };

  const rawFeatureFlagVal = selectDisplayValue(fields[fieldIds.featureFlag]);
  const featureFlagVal = rawFeatureFlagVal
    ? (FEATURE_FLAG_MAP[rawFeatureFlagVal] || rawFeatureFlagVal)
    : null;

  properties["Feature Flag"] = featureFlagVal
    ? { select: { name: featureFlagVal } }
    : { select: null };

  const aiFeatureVal = selectDisplayValue(fields[fieldIds.aiFeature]);
  properties["AI Feature"] = aiFeatureVal
    ? { select: { name: aiFeatureVal } }
    : { select: null };

  const tshirtSizeVal = selectDisplayValue(fields[fieldIds.tshirtSize]);
  properties[TSHIRT_SIZE_COL] = tshirtSizeVal
    ? { select: { name: tshirtSizeVal } }
    : { select: null };

  const designStatusVal = selectDisplayValue(fields[fieldIds.designStatus]);
  properties[DESIGN_STATUS_COL] = designStatusVal
    ? { select: { name: designStatusVal } }
    : { select: null };

  const piFeatureVal = selectDisplayValue(fields[fieldIds.piFeature]);
  properties[PI_FEATURE_COL] = piFeatureVal
    ? { select: { name: piFeatureVal } }
    : { select: null };

  const designerPeopleProperty = await buildDesignerPeopleProperty(
    fields[fieldIds.designer],
    context,
    name
  );

  if (designerPeopleProperty) {
    properties[DESIGNER_COL] = designerPeopleProperty;
  }

  const health = fields.health?.value || fields.health || null;
  const translatedHealth = HEALTH_MAP[health?.status] || "Not started";

  properties["PB_Health"] = {
    status: { name: translatedHealth }
  };

  const cleanComment = htmlToPlainText(health?.comment || "").substring(0, 2000);

  properties["Latest Health Comment"] = cleanComment
    ? { rich_text: [{ text: { content: cleanComment } }] }
    : { rich_text: [] };

  const healthTimestamp = health?.lastUpdatedAt || null;

  properties[ROADMAP_HEALTH_DATE_COL] = {
    date: healthTimestamp ? { start: healthTimestamp } : null
  };

  const pbStatus = fields.status?.name || selectDisplayValue(fields.status);

  properties[ROADMAP_STATUS_COL] = pbStatus
    ? { status: { name: pbStatus } }
    : { status: null };

  return { pbId, name, properties };
}

async function validateInitiativesNotionSchema() {
  if (!INITIATIVES_DATA_SOURCE) {
    console.log("NOTION_INITIATIVES_DATA_SOURCE_ID not set. Skipping initiatives sync.");
    return false;
  }

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
    const msg = `The Initiatives Notion DB is missing columns:\n\n${missingColumns.join("\n")}\n\nInitiatives sync skipped.`;
    console.error("Initiatives schema validation failed:", missingColumns);
    await sendTeamsAlert("Initiatives Schema Mismatch - Skipped", msg);
    return false;
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
    const msg = `Initiatives Notion DB column types are wrong:\n\n${typeMismatches.join("\n")}\n\nInitiatives sync skipped.`;
    console.error("Initiatives schema type validation failed:", typeMismatches);
    await sendTeamsAlert("Initiatives Schema Type Mismatch - Skipped", msg);
    return false;
  }

  console.log("Initiatives Notion schema validation passed.");
  return true;
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

async function syncInitiatives() {
  console.log("\n--- Starting Initiatives sync ---");

  const schemaOk = await validateInitiativesNotionSchema();
  if (!schemaOk) return;

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
    console.log("No initiatives returned from Productboard. Skipping.");
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
            console.log(`DRY RUN - would update initiative: ${name}`);
            wouldUpdate++;
          } else {
            await delay(350);
            await notion.pages.update({
              page_id: existingPage.id,
              properties
            });
            console.log(`Updated initiative: ${name}`);
            updated++;
          }
        } else {
          skipped++;
        }
      } else {
        if (!ALLOW_CREATES) {
          console.warn(`Create blocked by ALLOW_CREATES=false for initiative: ${name} (${pbId})`);
          wouldCreate++;
          continue;
        }

        if (DRY_RUN) {
          console.log(`DRY RUN - would create initiative: ${name} (${pbId})`);
          wouldCreate++;
          continue;
        }

        await delay(350);
        await notion.pages.create({
          parent: { data_source_id: INITIATIVES_DATA_SOURCE },
          properties
        });
        console.log(`Created initiative: ${name}`);
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

async function main() {
  console.log("Starting manual FacilityOS Productboard v2 -> Notion sync...");

  validateRequiredEnv();

  const cutoffISO = calculateManualCutoffISO();

  console.log(`DRY_RUN=${DRY_RUN}`);
  console.log(`ALLOW_CREATES=${ALLOW_CREATES}`);
  console.log(`DESIGNER_SYNC_STRICT=${DESIGNER_SYNC_STRICT}`);
  console.log(`PB_FETCH_ALL_FEATURES=${PB_FETCH_ALL_FEATURES}`);

  if (!PB_FETCH_ALL_FEATURES) {
    console.log(`Searching Productboard features updated since ${cutoffISO}`);
  } else {
    console.log("Fetching all Productboard feature entities for manual validation.");
  }

  await validateNotionSchema();

  const { fieldIds } = await getFeatureConfigAndCustomFieldIds();

  const returnFields = [
    "name",
    "status",
    "timeframe",
    "health",
    fieldIds.tier,
    fieldIds.appApproval,
    fieldIds.featureFlag,
    fieldIds.aiFeature,
    fieldIds.tshirtSize,
    fieldIds.designStatus,
    fieldIds.designer,
    fieldIds.piFeature
  ].filter(Boolean);

  const pbFeatures = await fetchPBFeatures(returnFields, cutoffISO);

  console.log(`Fetched ${pbFeatures.length} PB v2 feature entities.`);

  if (pbFeatures.length === 0) {
    console.log("--- SYNC COMPLETE ---");
    console.log("No Productboard features matched this run.");
    await syncInitiatives();
    return;
  }

  const { productMap, componentMap } = await getProductComponentMaps();

  const notionProductsRes = await fetchAllStandardNotionDataSource(PRODUCTS_DATA_SOURCE);
  const notionProductMap = {};

  notionProductsRes.forEach(page => {
    const nameProp = page.properties["Product Name"] || page.properties["Name"];
    const name = nameProp?.title?.[0]?.plain_text;

    if (name) {
      notionProductMap[name] = page.id;
    }
  });

  const existingRoadmapRes = await fetchAllStandardNotionDataSource(ROADMAP_DATA_SOURCE);
  const existingFeatures = {};

  existingRoadmapRes.forEach(page => {
    const pbId = page.properties.PBID?.rich_text?.[0]?.plain_text;

    if (pbId) {
      existingFeatures[pbId] = page;
    }
  });

  let notionUserByEmail = {};
  const pbUserEmailCache = {};

  try {
    const notionUsers = await fetchAllNotionUsers();
    notionUserByEmail = buildNotionUserMapByEmail(notionUsers);

    console.log(
      `Loaded ${Object.keys(notionUserByEmail).length} Notion users with email addresses for Designer mapping.`
    );
  } catch (e) {
    throw new Error(
      "Designer sync requires the Notion connection to have user information access with email addresses. " +
      `Could not load Notion users: ${e.message}`
    );
  }

  let updatedCount = 0;
  let createdCount = 0;
  let skippedCount = 0;
  let errorCount = 0;
  let wouldUpdateCount = 0;
  let wouldCreateCount = 0;

  for (const entity of pbFeatures) {
    const { pbId, name, properties } = await buildNotionPropertiesFromPBFeature(entity, {
      notionProductMap,
      productMap,
      componentMap,
      fieldIds,
      notionUserByEmail,
      pbUserEmailCache
    });

    const existingPage = existingFeatures[pbId];

    try {
      if (existingPage) {
        if (isDifferent(properties, existingPage.properties, name)) {
          if (DRY_RUN) {
            console.log(`DRY RUN - would update: ${name}`);
            wouldUpdateCount++;
          } else {
            await delay(350);

            await notion.pages.update({
              page_id: existingPage.id,
              properties
            });

            console.log(`Updated: ${name}`);
            updatedCount++;
          }
        } else {
          skippedCount++;
        }
      } else {
        if (!ALLOW_CREATES) {
          console.warn(`Create blocked by ALLOW_CREATES=false: ${name} (${pbId})`);
          wouldCreateCount++;
          continue;
        }

        if (DRY_RUN) {
          console.log(`DRY RUN - would create: ${name} (${pbId})`);
          wouldCreateCount++;
          continue;
        }

        await delay(350);

        await notion.pages.create({
          parent: { data_source_id: ROADMAP_DATA_SOURCE },
          properties
        });

        console.log(`Created: ${name}`);
        createdCount++;
      }
    } catch (err) {
      console.error(`\nSYNC FAILED ON FEATURE: "${name}" (${pbId})`);
      console.error("Payload Notion rejected:");
      console.error(JSON.stringify(properties, null, 2));
      console.error(`Error: ${err.message}\n`);

      errorCount++;

      if (errorCount <= 3) {
        await sendTeamsAlert(
          "Feature Sync Failed",
          `Feature: "${name}"\nPBID: ${pbId}\nError: ${err.message}`
        );
      }
    }
  }

  console.log("--- SYNC COMPLETE ---");
  console.log(`Created: ${createdCount} | Updated: ${updatedCount} | Skipped: ${skippedCount} | Errors: ${errorCount}`);

  if (DRY_RUN || !ALLOW_CREATES) {
    console.log(`Would create: ${wouldCreateCount} | Would update: ${wouldUpdateCount}`);
  }

  if (errorCount > 0) {
    await sendTeamsAlert(
      "Sync Completed With Errors",
      `${errorCount} feature(s) failed to sync.\n\nCreated: ${createdCount} | Updated: ${updatedCount} | Skipped: ${skippedCount}`
    );
  }

  await syncInitiatives();
}

main().catch(async (err) => {
  console.error("Fatal sync error:", err);

  await sendTeamsAlert(
    "Sync Crashed Completely",
    `The sync script crashed before completing.\n\nError: ${err.message}`
  );

  process.exit(1);
});
