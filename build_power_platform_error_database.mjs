import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = path.join(process.cwd(), "outputs", "power_platform_error_database");
await fs.mkdir(outputDir, { recursive: true });

const verifiedDate = "2026-06-09";

const sources = [
  ["SRC-001", "Power Automate cloud flow error code reference", "Microsoft Learn", "https://learn.microsoft.com/en-us/power-automate/error-reference", "Cloud flow design-time, runtime, connection, API, timeout, throttling, and licensing errors."],
  ["SRC-002", "Troubleshoot Dataverse Web API client errors", "Microsoft Learn", "https://learn.microsoft.com/en-us/troubleshoot/power-platform/dataverse/dataverse-web-api-and-sdk/web-api-client-errors", "Common Dataverse Web API request, method, payload, header, and lookup issues."],
  ["SRC-003", "Compose HTTP requests and handle errors", "Microsoft Learn", "https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/compose-http-requests-handle-errors", "Dataverse Web API URL structure, limits, status codes, headers, and error handling."],
  ["SRC-004", "Service protection API limits", "Microsoft Learn", "https://learn.microsoft.com/en-us/power-apps/developer/data-platform/api-limits", "Dataverse service protection limits and 429 handling guidance."],
  ["SRC-005", "Session creation error codes in unattended desktop flow runs", "Microsoft Learn", "https://learn.microsoft.com/en-us/troubleshoot/power-platform/power-automate/desktop-flows/troubleshoot-session-creation-errrors", "Power Automate desktop unattended-session creation failures."],
  ["SRC-006", "Power Apps wrap: Something went wrong error codes", "Microsoft Learn", "https://learn.microsoft.com/en-us/troubleshoot/power-platform/power-apps/manage-apps/something-went-wrong-error-codes", "Power Apps wrap authentication and build/login error codes."],
  ["SRC-007", "Best practices when updating a flow", "Microsoft Learn", "https://learn.microsoft.com/en-us/troubleshoot/power-platform/power-apps/connections/best-practices-when-updating-a-flow", "ConnectionAuthorizationFailed during Power Apps-triggered flow runs."],
  ["SRC-008", "Troubleshoot desktop flow run queue-based errors", "Microsoft Learn", "https://learn.microsoft.com/en-us/troubleshoot/power-platform/power-automate/desktop-flows/troubleshoot-desktop-flow-run-queue-errors", "Desktop flow queue and attended-user session errors."],
];

const records = [
  ["PP-001", "Power Automate", "Cloud flow design-time", "InvalidTemplate", "", "Flow cannot save or validate because an expression or action definition is malformed.", "Unmatched brackets or quotes; wrong function arguments; invalid type conversion; copied hidden characters; expression references wrong action name.", "Open the highlighted action, inspect the expression, correct action names and syntax, retype suspicious copied text, then save and retest.", "Use expression snippets, rename actions before writing formulas, and keep test data for expression validation.", "High", "Flow maker", "SRC-001", verifiedDate, "template expression save validation"],
  ["PP-002", "Power Automate", "Cloud flow design-time", "FlowCheckerError", "", "Flow checker blocks save or publish with one or more validation issues.", "Required fields are blank; connection not selected; parallel branch dynamic content is referenced unsafely; trigger input incomplete.", "Open the checker banner, visit each flagged action, fill required values, repair expressions, select valid connections, and save again.", "Run flow checker after major edits and avoid cross-branch references unless guarded.", "High", "Flow maker", "SRC-001", verifiedDate, "checker validation required fields"],
  ["PP-003", "Power Automate", "Cloud flow design-time", "DuplicateActionName", "", "Two or more actions share the same internal name.", "Copy and paste created a duplicate action; manually edited flow definition contains duplicate keys; renamed action conflicts with existing action.", "Find duplicate action names, rename one action, then update expressions that reference the old action name.", "Use clear action names immediately after adding or copying steps.", "Medium", "Flow maker", "SRC-001", verifiedDate, "duplicate action rename"],
  ["PP-004", "Power Automate", "Cloud flow design-time", "MissingRequiredProperty", "", "An action or trigger is missing a required input.", "Connector action not fully configured; dynamic value resolves to blank; solution environment variable lacks a value.", "Open the flagged action, fill red-asterisk fields, set environment variable values, and save.", "Document required environment variables in solution deployment notes.", "High", "Flow maker", "SRC-001", verifiedDate, "required property environment variable"],
  ["PP-005", "Power Automate", "Runtime expression", "ExpressionEvaluationFailed", "", "An expression fails at runtime because live data does not match the expected structure or type.", "Null object access; nonnumeric text passed to int; date format mismatch; division by zero from dynamic data.", "Inspect failed run inputs, add null checks, use coalesce defaults, validate values before conversion, and retest with known bad data.", "Add defensive expressions around optional fields and external inputs.", "High", "Flow maker", "SRC-001", verifiedDate, "runtime expression null conversion"],
  ["PP-006", "Power Automate", "Runtime expression", "ContentConversionFailed", "", "A value cannot be converted between expected and actual data types.", "String sent where integer or boolean expected; array sent where object expected; date string in unexpected format; binary content sent to text field.", "Compare action inputs to connector schema, use explicit conversion functions, normalize dates, and extract one item from arrays when needed.", "Validate schema at connector boundaries and keep sample payloads.", "Medium", "Flow maker", "SRC-001", verifiedDate, "conversion data type schema"],
  ["PP-007", "Power Automate", "Connection", "InvalidConnection", "", "A flow references a connection that is broken, expired, deleted, or unavailable in the environment.", "Password or MFA reset; connection deleted; admin removed connection; solution imported without matching connection.", "Open the flow, change or add the connection on warned actions, sign in again, save, and test.", "Use service principal connections for production flows where possible.", "High", "Admin / flow owner", "SRC-001", verifiedDate, "connection broken expired"],
  ["PP-008", "Power Automate", "Connection", "ConnectionNotConfigured", "", "An action requires a connection but none is selected.", "Solution import did not map connection references; new action was added but connection step skipped; environment variable missing.", "Select or create a connection for the action and map solution connection references.", "Include connection-reference checks in deployment runbooks.", "High", "Admin / flow owner", "SRC-001", verifiedDate, "connection reference import"],
  ["PP-009", "Power Automate", "Authentication", "Unauthorized", "401", "The downstream API rejects the request because authentication is invalid or expired.", "OAuth token expired; account disabled; password changed; service principal secret expired; Conditional Access blocked sign-in.", "Fix the connection, reauthenticate, rotate service principal credentials, and review Entra sign-in logs for policy blocks.", "Monitor credential expiry and prefer managed service accounts or service principals for critical flows.", "High", "Admin / identity owner", "SRC-001", verifiedDate, "401 auth token entra"],
  ["PP-010", "Power Automate", "Authorization", "Forbidden", "403", "The authenticated user or app cannot perform the requested operation.", "DLP policy blocks connector; user lacks resource permissions; connector restricted by tenant setting; premium license missing.", "Check DLP policies, verify target-resource permissions, confirm premium licensing, and contact admin if policy changes are needed.", "Review DLP and license impact before moving flows between environments.", "High", "Power Platform admin", "SRC-001", verifiedDate, "403 DLP permission license"],
  ["PP-011", "Power Automate", "Connection", "ConnectionAuthorizationFailed", "", "The connection exists but stored credentials are no longer valid.", "Password or MFA reset; OAuth refresh token expired; admin revoked consent; shared connection was unshared.", "Open Connections, select the affected connection, choose Fix connection, reauthenticate, and ask owners to reshare shared connections.", "Review stale connections and rotate service account secrets before expiry.", "High", "Flow owner", "SRC-001; SRC-007", verifiedDate, "connection authorization failed credentials"],
  ["PP-012", "Power Automate", "Connector/API", "ActionFailed", "", "A generic wrapper indicating an action failed; the detailed API error is in the action outputs.", "Downstream API returned 4xx or 5xx; child flow failed; custom connector response unexpected; run-after settings followed a prior failure.", "Open the failed run, expand the failed action outputs, find the actual status and message, and fix the underlying error.", "Add scoped error handling and log child-flow outputs.", "Medium", "Flow maker", "SRC-001", verifiedDate, "action failed outputs child flow"],
  ["PP-013", "Power Automate", "Connector/API", "BadRequest", "400", "Connector API rejects malformed or invalid input.", "Wrong data type; missing required field; invalid characters; field length exceeded.", "Review action inputs, compare them to connector documentation, sanitize invalid characters, and truncate values to field limits.", "Validate external and user-entered data before connector calls.", "Medium", "Flow maker", "SRC-001", verifiedDate, "400 bad request input"],
  ["PP-014", "Power Automate", "Connector/API", "NotFound", "404", "The resource referenced by the action does not exist or cannot be found.", "SharePoint list, file, Outlook folder, Teams channel, Dataverse row, or hardcoded ID changed or was deleted.", "Confirm the resource exists, update the action to the new name or ID, replace hardcoded IDs with lookups, and add graceful 404 handling.", "Avoid hardcoded resource IDs when a lookup is practical.", "Medium", "Flow maker", "SRC-001", verifiedDate, "404 resource missing"],
  ["PP-015", "Power Automate", "Trigger", "TriggerConditionNotMet", "", "Trigger evaluated its condition and intentionally did not start a run.", "Condition always false; condition references missing field; event does not match trigger filter.", "Review trigger settings, test the condition against known payloads, temporarily remove condition to inspect raw trigger output, then repair it.", "Keep trigger-condition examples beside the flow design notes.", "Low", "Flow maker", "SRC-001", verifiedDate, "trigger condition filter"],
  ["PP-016", "Power Automate", "Timeout", "ActionTimedOut", "", "A single action exceeds its configured timeout and is canceled.", "Slow HTTP API; approval expired; large file transfer; broad Dataverse query without pagination or filtering.", "Increase timeout where appropriate, use polling for long-running APIs, reduce Dataverse result sets, or add timeout branches.", "Set realistic timeouts and query filters during design.", "Medium", "Flow maker", "SRC-001", verifiedDate, "timeout HTTP Dataverse approval"],
  ["PP-017", "Power Automate", "Timeout", "OperationTimedOut", "", "A long-running operation exceeds the maximum wait time.", "Webhook wait never returns; approval has no expiration; Delay until exceeds run-duration limits; external service unavailable.", "Set explicit timeouts, configure run-after timeout handling, break long waits into smaller checks, or redesign as a relay pattern.", "Avoid flows that rely on unbounded waits.", "Medium", "Solution architect", "SRC-001", verifiedDate, "operation timed out relay"],
  ["PP-018", "Power Automate", "Throttling", "WorkflowRunActionRepetitionQuotaExceeded", "", "An Apply to each loop exceeds its allowed iteration count.", "Large list or table processed without filtering; nested loops multiply iterations; Get items returns all rows.", "Filter at the data source, use OData filters and top limits, batch work across runs, or replace loops with Select/Filter array where possible.", "Design flows to reduce rows before looping.", "Medium", "Flow maker", "SRC-001", verifiedDate, "loop quota apply to each"],
  ["PP-019", "Power Automate", "Throttling", "FlowRunQuotaExceeded", "", "Flow or tenant daily action execution limits are exceeded.", "High-volume flow, inefficient loop, repeated retries, or licensing/action quota reached.", "Optimize action count, reduce unnecessary polling, batch operations, and review licensing or capacity needs.", "Monitor action volume and set concurrency/retry policies intentionally.", "High", "Power Platform admin", "SRC-001", verifiedDate, "daily quota action count"],
  ["PP-020", "Power Automate", "Licensing", "DirectApiAuthorizationRequired", "", "Flow uses a premium connector but the caller or owner does not have the required premium entitlement.", "Premium connector used by seeded-license user; owner lost premium license; in-context flow became out-of-context.", "Identify the premium connector, assign appropriate Power Automate Premium or Process licensing, and confirm owner/caller entitlement.", "Review license requirements before sharing or embedding flows.", "High", "Licensing admin", "SRC-001", verifiedDate, "premium connector license"],
  ["PP-021", "Dataverse", "Web API", "Resource not found for the segment", "404", "Web API cannot resolve the entity set, action, or function segment.", "Wrong resource name; wrong case; inactive custom process action.", "Query the service document for entity set names, verify functions/actions in CSDL metadata, and confirm custom actions are active.", "Use metadata-driven constants for entity set names.", "Medium", "Developer", "SRC-002", verifiedDate, "Dataverse resource segment entity set"],
  ["PP-022", "Dataverse", "Web API", "Could not find a property named", "400", "Web API cannot find the requested property on the specified type.", "Property name uses wrong logical name or casing.", "Verify the property exists in CSDL metadata and use the exact case-sensitive name.", "Generate request models from metadata where possible.", "Medium", "Developer", "SRC-002", verifiedDate, "Dataverse property name case"],
  ["PP-023", "Dataverse", "Web API", "No HTTP resource was found that matches the request URI", "404", "The request URI does not match the operation because the wrong HTTP method was used.", "Function called with POST instead of GET, or action/function semantics mixed up.", "Confirm whether the endpoint is an OData function or action and use the correct HTTP method.", "Document method requirements beside custom API calls.", "Medium", "Developer", "SRC-002", verifiedDate, "Dataverse HTTP method function action"],
  ["PP-024", "Dataverse", "Web API", "Invalid property was found in entity", "400", "Create or update payload contains a property that Dataverse does not recognize.", "Incorrect property casing or display name used instead of logical name.", "Use the property logical name from metadata and correct the request payload.", "Use typed SDK or generated schemas for payloads.", "Medium", "Developer", "SRC-002", verifiedDate, "Dataverse payload property"],
  ["PP-025", "Dataverse", "Web API", "Does not support untyped value in non-open type", "400", "OData payload validation fails because an unexpected or incorrectly typed value is supplied.", "Wrong property name or type; sending data that cannot be mapped to a strongly typed Dataverse field.", "Correct property names and data types, then retry with a minimal payload.", "Validate JSON payloads against metadata before sending.", "Medium", "Developer", "SRC-002", verifiedDate, "OData untyped value payload"],
  ["PP-026", "Dataverse", "Web API", "Invalid URL", "400", "OData request segment exceeds the supported segment length.", "Individual OData segment is longer than 260 characters, often from long inline parameters or query strings.", "Shorten the segment, move long FetchXML queries into a batch body where supported, or simplify the query.", "Keep long queries in request bodies when possible.", "Medium", "Developer", "SRC-003", verifiedDate, "OData segment URL length"],
  ["PP-027", "Dataverse", "Web API", "Too Many Requests", "429", "Dataverse service protection limits are triggered by excessive API demand.", "Too many requests, too much execution time, or excessive concurrent requests from a client.", "Honor Retry-After, reduce concurrency, use batching carefully, cache reads, and spread work over time.", "Implement retry with backoff and request-volume monitoring.", "High", "Developer / admin", "SRC-004", verifiedDate, "429 throttling service protection"],
  ["PP-028", "Dataverse", "Web API", "Precondition Failed", "412", "Optimistic concurrency or conditional request precondition fails.", "ETag is stale; If-Match or If-None-Match condition no longer matches the row state.", "Retrieve the current row and ETag, reconcile changes, then retry the update with the correct condition.", "Use optimistic concurrency only where conflict handling is designed.", "Medium", "Developer", "SRC-003", verifiedDate, "412 ETag concurrency"],
  ["PP-029", "Dataverse", "Web API", "Not Modified", "304", "Conditional retrieval finds that the row has not changed since the provided condition.", "If-None-Match or modified-since condition indicates cached data is current.", "Treat as a non-error cache response and continue using cached data.", "Handle 304 separately from failure paths.", "Low", "Developer", "SRC-003", verifiedDate, "304 cache conditional"],
  ["PP-030", "Dataverse", "Web API", "Forbidden", "403", "Dataverse rejects the request due to missing privileges or field-level permissions.", "User lacks table/row/attribute privileges; security role missing; field security profile blocks access.", "Check security roles, team membership, business unit access, and field security permissions.", "Test APIs with least-privilege service principals before production.", "High", "Dataverse admin", "SRC-003", verifiedDate, "Dataverse security privilege 403"],
  ["PP-031", "Power Automate Desktop", "Desktop flow", "SessionCreationErrorWithThirdPartyCredentialProvider", "", "Unattended desktop flow cannot create the session because a third-party credential provider interferes.", "Unsupported or conflicting credential provider installed on the target machine.", "Remove, disable, or reconfigure the third-party credential provider and retest unattended run session creation.", "Keep unattended runner machines dedicated and avoid unsupported sign-in extensions.", "High", "Desktop flow admin", "SRC-005", verifiedDate, "PAD unattended session credential provider"],
  ["PP-032", "Power Automate Desktop", "Desktop flow", "SessionCreationUserPromptedForCredentialsAfterConnection", "", "Desktop flow session creation fails because credentials are requested after connection is attempted.", "Machine or policy prompts for credentials unexpectedly during unattended RDP/session setup.", "Review machine sign-in policies and credentials, confirm unattended connection configuration, and remove prompts that require user interaction.", "Validate unattended runner login after policy or credential changes.", "High", "Desktop flow admin", "SRC-005", verifiedDate, "PAD unattended credentials prompt"],
  ["PP-033", "Power Automate Desktop", "Desktop flow", "SessionCreationError", "", "Generic unattended desktop flow session creation failure.", "RDP/session setup failure, machine configuration issue, or environment policy conflict.", "Check gateway/machine registration, credentials, RDP permissions, machine availability, and desktop flow service status.", "Monitor machine health and keep runner configuration documented.", "High", "Desktop flow admin", "SRC-005", verifiedDate, "PAD session creation generic"],
  ["PP-034", "Power Automate Desktop", "Desktop flow queue", "AttendedUserSessionNotActive", "", "Attended desktop flow cannot run because the user session is not active.", "Target user is signed out, locked, disconnected, or lacks active desktop session.", "Have the attended user sign in and keep the session active, or convert the process to unattended where appropriate.", "Match attended vs unattended design to operational reality.", "Medium", "Desktop flow owner", "SRC-008", verifiedDate, "PAD attended user session"],
  ["PP-035", "Power Automate Desktop", "Desktop flow queue", "AttendedUserNotLoggedIn", "", "Attended desktop flow cannot run because the user is not logged in.", "No active logged-in Windows session for the attended automation user.", "Log in as the target user before running attended automation and confirm UIFlowService permissions.", "Use unattended flows for server-style execution.", "Medium", "Desktop flow owner", "SRC-008", verifiedDate, "PAD attended login"],
  ["PP-036", "Power Apps", "Wrap", "Something went wrong [5objp]", "", "Wrapped app login fails even if build succeeds.", "APK signature hash key mismatch or redirect URI mismatch in Microsoft Entra app registration.", "Generate the correct hash key from the signing keystore, update Android platform authentication settings, and verify redirect URI casing and encoding.", "Treat keystore, hash, bundle ID, and redirect URI as release-controlled values.", "High", "Power Apps maker / identity admin", "SRC-006", verifiedDate, "Power Apps wrap hash redirect"],
  ["PP-037", "Power Apps", "Wrap", "Something went wrong [9n155]", "", "Wrapped app authentication fails due to app registration tenant configuration.", "Microsoft Entra app was registered as single tenant instead of the expected supported account type.", "Update app registration supported account types as required by the wrap configuration, then rebuild and retest login.", "Use a wrap checklist for Entra registration settings.", "High", "Identity admin", "SRC-006", verifiedDate, "Power Apps wrap single tenant"],
  ["PP-038", "Power Apps", "Canvas app / connector", "DataSourceInfo or connector permission error", "403", "Canvas app action or connector call fails because the user lacks permission or the connector is blocked.", "Resource permissions missing, connector blocked by DLP, or environment policy prevents the call.", "Check app sharing, data-source permissions, DLP policy, and connector licensing for the user.", "Validate app access with a test user before release.", "High", "App owner / admin", "SRC-001; SRC-003", verifiedDate, "Power Apps connector permissions"],
];

const workbook = Workbook.create();
const dashboard = workbook.worksheets.add("Dashboard");
const db = workbook.worksheets.add("Error Database");
const lookup = workbook.worksheets.add("Lookup Lists");
const sourceSheet = workbook.worksheets.add("Sources");

for (const sheet of [dashboard, db, lookup, sourceSheet]) {
  sheet.showGridLines = false;
}

const headers = [
  "ID",
  "Product Area",
  "Layer",
  "Error Code / Message",
  "HTTP Status",
  "Symptoms / Meaning",
  "Common Causes",
  "Resolution Steps",
  "Prevention",
  "Severity",
  "Owner",
  "Source ID",
  "Last Verified",
  "Search Tags",
];

db.getRange("A1:N1").values = [headers];
db.getRangeByIndexes(1, 0, records.length, headers.length).values = records;
const dbRange = `A1:N${records.length + 1}`;
const dbTable = db.tables.add(dbRange, true, "ErrorDatabase");
dbTable.style = "TableStyleMedium2";
db.freezePanes.freezeRows(1);
db.getRange("A1:N1").format = {
  fill: "#1F4E5F",
  font: { bold: true, color: "#FFFFFF" },
  wrapText: true,
};
db.getRange(`A2:N${records.length + 1}`).format = {
  wrapText: true,
  verticalAlignment: "top",
};
db.getRange(`M2:M${records.length + 1}`).setNumberFormat("yyyy-mm-dd");
const widths = [70, 130, 150, 230, 85, 310, 330, 360, 300, 90, 150, 110, 105, 220];
for (let i = 0; i < widths.length; i++) {
  db.getRangeByIndexes(0, i, records.length + 1, 1).format.columnWidthPx = widths[i];
}
db.getRange(`A1:N${records.length + 1}`).format.borders = { preset: "all", style: "thin", color: "#D9E2E7" };
db.getRange(`J2:J${records.length + 1}`).dataValidation = { rule: { type: "list", values: ["High", "Medium", "Low"] } };

const productAreas = [...new Set(records.map((r) => r[1]))];
const severities = ["High", "Medium", "Low"];
const layers = [...new Set(records.map((r) => r[2]))];
const owners = [...new Set(records.map((r) => r[10]))];

lookup.getRange("A1").values = [["Product Areas"]];
lookup.getRangeByIndexes(1, 0, productAreas.length, 1).values = productAreas.map((x) => [x]);
lookup.getRange("C1").values = [["Severity"]];
lookup.getRangeByIndexes(1, 2, severities.length, 1).values = severities.map((x) => [x]);
lookup.getRange("E1").values = [["Layers"]];
lookup.getRangeByIndexes(1, 4, layers.length, 1).values = layers.map((x) => [x]);
lookup.getRange("G1").values = [["Owners"]];
lookup.getRangeByIndexes(1, 6, owners.length, 1).values = owners.map((x) => [x]);
lookup.getRange("A1:G1").format = {
  fill: "#1F4E5F",
  font: { bold: true, color: "#FFFFFF" },
};
lookup.getRange("A:G").format.columnWidthPx = 190;

sourceSheet.getRange("A1:E1").values = [["Source ID", "Title", "Publisher", "URL", "Coverage"]];
sourceSheet.getRangeByIndexes(1, 0, sources.length, 5).values = sources;
sourceSheet.tables.add(`A1:E${sources.length + 1}`, true, "SourceTable").style = "TableStyleMedium4";
sourceSheet.freezePanes.freezeRows(1);
sourceSheet.getRange("A1:E1").format = {
  fill: "#1F4E5F",
  font: { bold: true, color: "#FFFFFF" },
  wrapText: true,
};
sourceSheet.getRange(`A2:E${sources.length + 1}`).format = { wrapText: true, verticalAlignment: "top" };
for (const [col, width] of [0, 85, 1, 310, 2, 130, 3, 520, 4, 390].reduce((acc, val, idx, arr) => idx % 2 === 0 ? [...acc, [val, arr[idx + 1]]] : acc, [])) {
  sourceSheet.getRangeByIndexes(0, col, sources.length + 1, 1).format.columnWidthPx = width;
}

dashboard.getRange("A1:H1").merge();
dashboard.getRange("A1").values = [["Power Platform Error Codes and Resolution Database"]];
dashboard.getRange("A2:H2").merge();
dashboard.getRange("A2").values = [[`Starter knowledge base verified against Microsoft Learn on ${verifiedDate}. Use the Error Database tab for filtering and handoff notes.`]];
dashboard.getRange("A1:H1").format.rowHeightPx = 30;
dashboard.getRange("A2:H2").format.rowHeightPx = 22;
dashboard.getRange("A1:H2").format = {
  fill: "#163B4B",
  font: { color: "#FFFFFF", bold: true },
};
dashboard.getRange("A1").format.font = { color: "#FFFFFF", bold: true, size: 18 };
dashboard.getRange("A2").format.font = { color: "#DDECF2", size: 10 };
dashboard.getRange("A1:H2").format.borders = { preset: "outside", style: "medium", color: "#163B4B" };

dashboard.getRange("A4:B8").values = [
  ["Metric", "Value"],
  ["Total records", records.length],
  ["High severity", records.filter((r) => r[9] === "High").length],
  ["Products covered", productAreas.length],
  ["Source documents", sources.length],
];
dashboard.getRange("A4:B4").format = { fill: "#1F4E5F", font: { bold: true, color: "#FFFFFF" } };
dashboard.getRange("A5:A8").format = { fill: "#EAF3F6", font: { bold: true } };
dashboard.getRange("B5:B8").format = { fill: "#F7FBFC", font: { bold: true, color: "#163B4B" } };
dashboard.getRange("A4:B8").format.borders = { preset: "all", style: "thin", color: "#D9E2E7" };

dashboard.getRange("D4:E4").values = [["Severity", "Records"]];
dashboard.getRangeByIndexes(4, 3, severities.length, 1).values = severities.map((s) => [s]);
dashboard.getRange("E5").formulas = [[`=COUNTIF('Error Database'!J2:J${records.length + 1},D5)`]];
dashboard.getRange(`E5:E${4 + severities.length}`).fillDown();
dashboard.getRange(`D4:E${4 + severities.length}`).format.borders = { preset: "all", style: "thin", color: "#D9E2E7" };
dashboard.getRange("D4:E4").format = { fill: "#1F4E5F", font: { bold: true, color: "#FFFFFF" } };

dashboard.getRange("G4:H4").values = [["Product Area", "Records"]];
dashboard.getRangeByIndexes(4, 6, productAreas.length, 1).values = productAreas.map((p) => [p]);
dashboard.getRange("H5").formulas = [[`=COUNTIF('Error Database'!B2:B${records.length + 1},G5)`]];
dashboard.getRange(`H5:H${4 + productAreas.length}`).fillDown();
dashboard.getRange(`G4:H${4 + productAreas.length}`).format.borders = { preset: "all", style: "thin", color: "#D9E2E7" };
dashboard.getRange("G4:H4").format = { fill: "#1F4E5F", font: { bold: true, color: "#FFFFFF" } };

dashboard.getRange("A11:C11").merge();
dashboard.getRange("A11").values = [["Operating Notes"]];
const notes = [
  "1. Filter by Product Area, Severity, Owner, or Search Tags in the Error Database table.",
  "2. Use the Resolution Steps column as the immediate fix path and Prevention as the durable control.",
  "3. Add organization-specific incidents below the current rows and keep Source ID or ticket references.",
  "4. Refresh sources periodically because Microsoft updates Power Platform limits, licensing, and error guidance.",
];
for (let i = 0; i < notes.length; i++) {
  const row = 12 + i;
  dashboard.getRange(`A${row}:C${row}`).merge();
  dashboard.getRange(`A${row}`).values = [[notes[i]]];
  dashboard.getRange(`A${row}:C${row}`).format.rowHeightPx = 58;
}
dashboard.getRange("A11:C11").format = { fill: "#1F4E5F", font: { bold: true, color: "#FFFFFF" } };
dashboard.getRange("A12:C15").format = { fill: "#F7FBFC", wrapText: true, verticalAlignment: "top" };
dashboard.getRange("A11:C15").format.borders = { preset: "all", style: "thin", color: "#D9E2E7" };

const sevChart = dashboard.charts.add("bar", dashboard.getRange(`D4:E${4 + severities.length}`));
sevChart.title = "Records by Severity";
sevChart.hasLegend = false;
sevChart.xAxis = { axisType: "textAxis" };
sevChart.setPosition("D10", "F23");

const productChart = dashboard.charts.add("bar", dashboard.getRange(`G4:H${4 + productAreas.length}`));
productChart.title = "Records by Product Area";
productChart.hasLegend = false;
productChart.xAxis = { axisType: "textAxis" };
productChart.setPosition("G10", "K24");

for (const [col, width] of [["A", 160], ["B", 120], ["C", 30], ["D", 130], ["E", 95], ["F", 30], ["G", 180], ["H", 95]]) {
  dashboard.getRange(`${col}:${col}`).format.columnWidthPx = width;
}
dashboard.getRange("A:H").format.verticalAlignment = "top";

const keyPreview = await workbook.inspect({
  kind: "table",
  range: "Dashboard!A1:H15",
  include: "values,formulas",
  tableMaxRows: 15,
  tableMaxCols: 8,
});
console.log(keyPreview.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

for (const sheetName of ["Dashboard", "Error Database", "Sources"]) {
  const preview = await workbook.render({
    sheetName,
    autoCrop: "all",
    scale: 1,
    format: "png",
  });
  await fs.writeFile(path.join(outputDir, `${sheetName.replaceAll(" ", "_").toLowerCase()}.png`), new Uint8Array(await preview.arrayBuffer()));
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(path.join(outputDir, "power_platform_error_codes_resolution_database.xlsx"));
