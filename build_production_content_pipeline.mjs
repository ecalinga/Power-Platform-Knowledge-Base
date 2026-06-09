import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = process.cwd();
const dataDir = path.join(root, "data");
const outputDir = path.join(root, "outputs", "power_platform_error_database");
const errorDir = path.join(root, "errors");
const indexPath = path.join(root, "index.html");
const siteBaseUrl = "https://example.com";
const verifiedDate = "2026-06-09";
const targetCount = Number.parseInt(process.env.ARTICLE_TARGET_COUNT || "1000", 10);

if (!Number.isInteger(targetCount) || targetCount < 1) {
  throw new Error("ARTICLE_TARGET_COUNT must be a positive integer.");
}

const sources = {
  "SRC-001": {
    title: "Power Automate cloud flow error code reference",
    url: "https://learn.microsoft.com/en-us/power-automate/error-reference",
  },
  "SRC-002": {
    title: "Troubleshoot Dataverse Web API client errors",
    url: "https://learn.microsoft.com/en-us/troubleshoot/power-platform/dataverse/dataverse-web-api-and-sdk/web-api-client-errors",
  },
  "SRC-003": {
    title: "Compose HTTP requests and handle errors",
    url: "https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/compose-http-requests-handle-errors",
  },
  "SRC-004": {
    title: "Service protection API limits",
    url: "https://learn.microsoft.com/en-us/power-apps/developer/data-platform/api-limits",
  },
  "SRC-005": {
    title: "Session creation error codes in unattended desktop flow runs",
    url: "https://learn.microsoft.com/en-us/troubleshoot/power-platform/power-automate/desktop-flows/troubleshoot-session-creation-errrors",
  },
  "SRC-006": {
    title: "Power Apps wrap: Something went wrong error codes",
    url: "https://learn.microsoft.com/en-us/troubleshoot/power-platform/power-apps/manage-apps/something-went-wrong-error-codes",
  },
  "SRC-007": {
    title: "Best practices when updating a flow",
    url: "https://learn.microsoft.com/en-us/troubleshoot/power-platform/power-apps/connections/best-practices-when-updating-a-flow",
  },
  "SRC-008": {
    title: "Troubleshoot desktop flow run queue-based errors",
    url: "https://learn.microsoft.com/en-us/troubleshoot/power-platform/power-automate/desktop-flows/troubleshoot-desktop-flow-run-queue-errors",
  },
};

const contexts = [
  {
    key: "sharepoint-list-item",
    title: "SharePoint List Item Flow",
    appliesTo: "Power Automate cloud flow using SharePoint list items",
    scenario: "The flow reads or updates SharePoint list data during an approval, notification, or sync process.",
    causeHint: "Confirm the site, list, item ID, and connection user permissions before changing expressions.",
    owner: "Flow maker",
  },
  {
    key: "outlook-email-automation",
    title: "Outlook Email Automation",
    appliesTo: "Power Automate cloud flow using Outlook mail actions",
    scenario: "The flow sends, moves, or reads email and then passes mailbox data to later actions.",
    causeHint: "Check mailbox permissions, shared mailbox access, and whether the connection account still has the expected mailbox role.",
    owner: "Flow maker",
  },
  {
    key: "teams-notification-flow",
    title: "Teams Notification Flow",
    appliesTo: "Power Automate cloud flow posting to Microsoft Teams",
    scenario: "The automation posts adaptive cards or channel messages after a business event.",
    causeHint: "Confirm the team, channel, connector permissions, message payload, and DLP policy for the environment.",
    owner: "Flow maker",
  },
  {
    key: "dataverse-row-operation",
    title: "Dataverse Row Operation",
    appliesTo: "Power Automate or app logic creating, reading, updating, or deleting Dataverse rows",
    scenario: "The process calls Dataverse tables directly from a flow, app, plugin, custom connector, or Web API client.",
    causeHint: "Check table permissions, row ownership, logical names, column types, and service protection limits.",
    owner: "Dataverse admin",
  },
  {
    key: "canvas-app-form",
    title: "Canvas App Form",
    appliesTo: "Power Apps canvas app form or formula",
    scenario: "The user saves a form, loads a record, calls a flow, or uses connector data inside a canvas app.",
    causeHint: "Check app sharing, data source permissions, formula null handling, and connector policy restrictions.",
    owner: "App owner",
  },
  {
    key: "model-driven-app-command",
    title: "Model-driven App Command",
    appliesTo: "Power Apps model-driven app command, form, or view",
    scenario: "A user opens a table form, triggers a command, or runs business logic from a model-driven app.",
    causeHint: "Check security roles, command visibility rules, table privileges, and field security profiles.",
    owner: "Model-driven app maker",
  },
  {
    key: "solution-import",
    title: "Solution Import",
    appliesTo: "Power Platform solution import or deployment pipeline",
    scenario: "A managed or unmanaged solution is imported into another environment as part of release work.",
    causeHint: "Review missing dependencies, connection references, environment variables, publisher prefixes, and target environment settings.",
    owner: "Release owner",
  },
  {
    key: "desktop-flow-runner",
    title: "Desktop Flow Runner",
    appliesTo: "Power Automate Desktop attended or unattended run",
    scenario: "A desktop flow starts on a machine or machine group and attempts to create or reuse a Windows session.",
    causeHint: "Check machine availability, credentials, RDP policy, user session state, gateway registration, and installed credential providers.",
    owner: "Desktop flow admin",
  },
  {
    key: "custom-connector-http",
    title: "Custom Connector HTTP Call",
    appliesTo: "Power Platform custom connector or HTTP action",
    scenario: "A cloud flow or app calls an external REST API through an HTTP action or custom connector.",
    causeHint: "Review the request body, headers, authentication scheme, response schema, throttling behavior, and retry policy.",
    owner: "Integration owner",
  },
  {
    key: "gateway-data-source",
    title: "On-premises Gateway Data Source",
    appliesTo: "Power Automate, Power Apps, or connector using an on-premises data gateway",
    scenario: "The automation or app reaches SQL Server, file shares, or internal systems through a gateway connection.",
    causeHint: "Check gateway online status, data source mapping, credentials, network reachability, and connector permissions.",
    owner: "Gateway admin",
  },
  {
    key: "sql-server-connector",
    title: "SQL Server Connector",
    appliesTo: "Power Automate or Power Apps using SQL Server through a connector or gateway",
    scenario: "The app or flow queries, inserts, updates, or deletes SQL Server records as part of a business process.",
    causeHint: "Check SQL permissions, gateway mapping, stored procedure parameters, table schema changes, and query timeout behavior.",
    owner: "Database owner",
  },
  {
    key: "approval-workflow",
    title: "Approval Workflow",
    appliesTo: "Power Automate approval process",
    scenario: "The flow creates, waits for, updates, or reacts to an approval request.",
    causeHint: "Check approver identity, approval timeout, response schema, action run-after settings, and notification delivery.",
    owner: "Flow owner",
  },
  {
    key: "child-flow-call",
    title: "Child Flow Call",
    appliesTo: "Power Automate parent flow calling a child flow",
    scenario: "A parent flow invokes reusable child-flow logic and waits for a result.",
    causeHint: "Check solution membership, trigger type, response action, input schema, connection references, and child-flow ownership.",
    owner: "Solution owner",
  },
  {
    key: "adaptive-card-post",
    title: "Adaptive Card Post",
    appliesTo: "Power Automate or Teams workflow posting adaptive cards",
    scenario: "The flow sends an adaptive card to Teams or Outlook and later processes the response.",
    causeHint: "Check adaptive card JSON, unsupported schema version, user identity, channel permissions, and response parsing.",
    owner: "Collaboration app owner",
  },
  {
    key: "excel-online-table",
    title: "Excel Online Table",
    appliesTo: "Power Automate flow using Excel Online table actions",
    scenario: "The automation reads, lists, inserts, or updates rows in an Excel table stored in OneDrive or SharePoint.",
    causeHint: "Check workbook location, table name, file lock state, column headers, row keys, and connector throttling.",
    owner: "Flow maker",
  },
  {
    key: "power-pages-form",
    title: "Power Pages Form",
    appliesTo: "Power Pages form, list, or Dataverse-backed page",
    scenario: "A portal user submits a form, views a list, signs in, or interacts with Dataverse data through Power Pages.",
    causeHint: "Check table permissions, web roles, site settings, authentication provider configuration, and Dataverse row access.",
    owner: "Power Pages admin",
  },
  {
    key: "copilot-studio-topic",
    title: "Copilot Studio Topic",
    appliesTo: "Copilot Studio topic or action calling Power Platform services",
    scenario: "A copilot topic calls a flow, connector, or Dataverse operation and returns the result to the user.",
    causeHint: "Check topic variables, action input schema, authentication, environment connection references, and user permissions.",
    owner: "Copilot owner",
  },
  {
    key: "managed-solution-upgrade",
    title: "Managed Solution Upgrade",
    appliesTo: "Power Platform managed solution upgrade or patch",
    scenario: "A managed solution update, patch, or stage-for-upgrade operation is applied to a target environment.",
    causeHint: "Check solution layering, missing dependencies, active customizations, publisher identity, and upgrade path order.",
    owner: "Release manager",
  },
  {
    key: "environment-move-copy",
    title: "Environment Move or Copy",
    appliesTo: "Power Platform environment copy, restore, backup, or tenant move",
    scenario: "An environment is copied, restored, moved, or reconfigured and existing apps or flows are tested afterward.",
    causeHint: "Check environment URLs, connection references, callback URLs, app registrations, data source IDs, and tenant policies.",
    owner: "Power Platform admin",
  },
  {
    key: "service-principal-connection",
    title: "Service Principal Connection",
    appliesTo: "Power Platform connection using app registration or service principal authentication",
    scenario: "A flow, app, or integration authenticates with an application identity instead of a personal user account.",
    causeHint: "Check app registration permissions, consent, secret expiry, certificate validity, security roles, and connection ownership.",
    owner: "Identity admin",
  },
];

const templates = [
  {
    code: "InvalidTemplate",
    status: "",
    product: "Power Automate",
    layer: "Cloud flow design-time",
    severity: "High",
    source: "SRC-001",
    meaning: "The flow definition cannot be saved because an expression, action configuration, or template structure is invalid.",
    causes: "Broken expression syntax, wrong action reference, unsupported function arguments, invalid constant conversion, or copied hidden characters.",
    resolution: "Open the highlighted action, inspect the expression, correct syntax and action references, remove hidden characters, save, and retest the flow.",
    prevention: "Use named actions, test expressions with sample payloads, and keep risky expressions small enough to review.",
    tags: "invalid template expression save validation",
  },
  {
    code: "FlowCheckerError",
    status: "",
    product: "Power Automate",
    layer: "Cloud flow design-time",
    severity: "High",
    source: "SRC-001",
    meaning: "The flow checker found one or more validation issues that block saving or publishing.",
    causes: "Required fields are missing, a connector connection is not selected, trigger inputs are incomplete, or dynamic content is referenced unsafely.",
    resolution: "Open the checker banner, visit each flagged action, fill required fields, select connections, repair expressions, save, and test.",
    prevention: "Run flow checker after large edits and keep connection references mapped in every deployment environment.",
    tags: "flow checker validation required fields",
  },
  {
    code: "DuplicateActionName",
    status: "",
    product: "Power Automate",
    layer: "Cloud flow design-time",
    severity: "Medium",
    source: "SRC-001",
    meaning: "Two actions use the same internal action name inside the flow definition.",
    causes: "Copied actions, manually edited JSON, or renamed steps that collide with an existing action name.",
    resolution: "Rename one duplicate action and update expressions that reference the old action name.",
    prevention: "Rename copied actions immediately and avoid manual edits that duplicate action keys.",
    tags: "duplicate action rename flow definition",
  },
  {
    code: "MissingRequiredProperty",
    status: "",
    product: "Power Automate",
    layer: "Cloud flow design-time",
    severity: "High",
    source: "SRC-001",
    meaning: "A trigger or action is missing an input value that is required before the flow can run correctly.",
    causes: "Connector setup is incomplete, an environment variable is blank, or a required dynamic value resolves to empty.",
    resolution: "Open the flagged action, fill required fields, map environment variables, save the flow, and run a controlled test.",
    prevention: "Add deployment checks for required connection references and environment variables.",
    tags: "missing property required field environment variable",
  },
  {
    code: "ExpressionEvaluationFailed",
    status: "",
    product: "Power Automate",
    layer: "Runtime expression",
    severity: "High",
    source: "SRC-001",
    meaning: "A runtime expression failed because live data did not match the expected type, shape, or value.",
    causes: "Null object access, nonnumeric text passed to a number function, date format mismatch, or division by zero.",
    resolution: "Inspect the failed run inputs, add null checks, use coalesce defaults, validate types before conversion, and retest.",
    prevention: "Guard optional fields and external inputs with defensive expressions before using them downstream.",
    tags: "runtime expression null conversion type",
  },
  {
    code: "ContentConversionFailed",
    status: "",
    product: "Power Automate",
    layer: "Runtime expression",
    severity: "Medium",
    source: "SRC-001",
    meaning: "The flow could not convert a value into the data type expected by the next action.",
    causes: "String, number, boolean, object, array, date, or binary content is passed to an incompatible input.",
    resolution: "Compare the action input to the connector schema, normalize values with explicit conversion functions, and retry.",
    prevention: "Validate connector-bound payloads and keep example payloads for regression testing.",
    tags: "content conversion failed data type schema",
  },
  {
    code: "InvalidConnection",
    status: "",
    product: "Power Automate",
    layer: "Connection",
    severity: "High",
    source: "SRC-001",
    meaning: "The flow references a connection that is broken, deleted, expired, or unavailable.",
    causes: "Password reset, MFA reset, deleted connection, revoked admin access, or imported solution without matching connection references.",
    resolution: "Open the flow, choose a valid connection for warned actions, authenticate again, save, and test.",
    prevention: "Use service principals or governed service accounts for production integrations where supported.",
    tags: "invalid connection broken expired deleted",
  },
  {
    code: "ConnectionNotConfigured",
    status: "",
    product: "Power Automate",
    layer: "Connection",
    severity: "High",
    source: "SRC-001",
    meaning: "An action requires a connector connection but no connection is currently selected.",
    causes: "Connection references were not mapped during import, a new action was not configured, or an environment variable lacks a value.",
    resolution: "Select or create the required connection, map connection references in the solution, save, and test.",
    prevention: "Include connection-reference mapping in release checklists.",
    tags: "connection not configured reference import",
  },
  {
    code: "ConnectionAuthorizationFailed",
    status: "",
    product: "Power Automate",
    layer: "Connection",
    severity: "High",
    source: "SRC-001; SRC-007",
    meaning: "The connection exists, but its stored credentials or authorization are no longer valid.",
    causes: "Password change, MFA reset, refresh token expiry, revoked consent, or shared connection access removed.",
    resolution: "Open Connections, select the affected connection, fix or reauthenticate it, then rerun the failed flow.",
    prevention: "Review stale connections and rotate service account credentials before expiry.",
    tags: "connection authorization failed credentials oauth",
  },
  {
    code: "Unauthorized",
    status: "401",
    product: "Power Automate",
    layer: "Authentication",
    severity: "High",
    source: "SRC-001",
    meaning: "The service rejected the request because the authentication token or credential is invalid.",
    causes: "Expired OAuth token, disabled account, changed password, expired service principal secret, or Conditional Access block.",
    resolution: "Fix the connection, reauthenticate, rotate service principal credentials, and review Microsoft Entra sign-in logs.",
    prevention: "Monitor credential expiry and avoid personal-user connections for critical production automations.",
    tags: "401 unauthorized authentication token entra",
  },
  {
    code: "Forbidden",
    status: "403",
    product: "Power Automate",
    layer: "Authorization",
    severity: "High",
    source: "SRC-001",
    meaning: "The authenticated identity does not have permission to perform the requested operation.",
    causes: "DLP policy block, missing resource permission, tenant connector restriction, missing license, or security role gap.",
    resolution: "Check DLP policy, target-resource permissions, connector restrictions, and licensing for the user or service account.",
    prevention: "Test flows with least-privilege accounts and document DLP and license requirements.",
    tags: "403 forbidden permission DLP license",
  },
  {
    code: "ActionFailed",
    status: "",
    product: "Power Automate",
    layer: "Connector/API",
    severity: "Medium",
    source: "SRC-001",
    meaning: "A generic action failure occurred; the detailed error is inside the failed action output.",
    causes: "Downstream API returned an error, child flow failed, custom connector response was unexpected, or run-after settings followed a prior failure.",
    resolution: "Open the failed run, expand the action output, identify the actual status and message, then fix the underlying issue.",
    prevention: "Add scoped error handling and log the original action output for production failures.",
    tags: "action failed connector output child flow",
  },
  {
    code: "BadRequest",
    status: "400",
    product: "Power Automate",
    layer: "Connector/API",
    severity: "Medium",
    source: "SRC-001",
    meaning: "The connector or API rejected the request because the submitted input is malformed or invalid.",
    causes: "Wrong data type, missing required field, invalid characters, unsupported value, or field length exceeded.",
    resolution: "Inspect action inputs, compare them with the connector schema, sanitize values, and retry with a minimal payload.",
    prevention: "Validate user and external data before sending it to connector actions.",
    tags: "400 bad request malformed input",
  },
  {
    code: "NotFound",
    status: "404",
    product: "Power Automate",
    layer: "Connector/API",
    severity: "Medium",
    source: "SRC-001",
    meaning: "The referenced resource cannot be found by the connector or service.",
    causes: "Deleted or renamed list, folder, mailbox, channel, file, Dataverse row, or hardcoded ID.",
    resolution: "Confirm the resource still exists, update IDs or names, replace hardcoded references with lookups, and add graceful handling.",
    prevention: "Avoid static resource IDs when a reliable lookup is available.",
    tags: "404 not found missing resource",
  },
  {
    code: "TriggerConditionNotMet",
    status: "",
    product: "Power Automate",
    layer: "Trigger",
    severity: "Low",
    source: "SRC-001",
    meaning: "The trigger condition evaluated to false, so the flow did not start.",
    causes: "Incorrect trigger expression, missing payload field, or event data that does not match the configured filter.",
    resolution: "Review trigger settings, inspect a real trigger payload, temporarily remove the condition if needed, and correct the expression.",
    prevention: "Keep sample payloads and trigger-condition examples with the flow documentation.",
    tags: "trigger condition not met filter",
  },
  {
    code: "ActionTimedOut",
    status: "",
    product: "Power Automate",
    layer: "Timeout",
    severity: "Medium",
    source: "SRC-001",
    meaning: "A single action exceeded its configured timeout and was canceled.",
    causes: "Slow external API, large file transfer, approval wait, or broad Dataverse query without filtering.",
    resolution: "Increase the timeout where appropriate, reduce query size, use long-running operation polling, or add timeout handling.",
    prevention: "Set realistic timeouts and constrain expensive connector operations.",
    tags: "timeout action timed out HTTP approval",
  },
  {
    code: "OperationTimedOut",
    status: "",
    product: "Power Automate",
    layer: "Timeout",
    severity: "Medium",
    source: "SRC-001",
    meaning: "A long-running operation exceeded the maximum wait time for the flow pattern.",
    causes: "Webhook callback never arrived, approval has no expiration, delay exceeds run duration, or external service stayed unavailable.",
    resolution: "Set explicit timeouts, configure run-after timeout branches, and redesign long waits into relay-style runs.",
    prevention: "Avoid unbounded waits in production flows.",
    tags: "operation timed out webhook approval relay",
  },
  {
    code: "WorkflowRunActionRepetitionQuotaExceeded",
    status: "",
    product: "Power Automate",
    layer: "Throttling",
    severity: "Medium",
    source: "SRC-001",
    meaning: "A loop exceeded the allowed number of repetitions for a single flow run.",
    causes: "Large unfiltered dataset, nested loops multiplying iterations, or Get items returning all rows.",
    resolution: "Filter at the source, use OData filters and top limits, batch work across runs, or use Select and Filter array where possible.",
    prevention: "Reduce rows before loops and monitor run volume.",
    tags: "loop quota apply to each repetitions",
  },
  {
    code: "FlowRunQuotaExceeded",
    status: "",
    product: "Power Automate",
    layer: "Throttling",
    severity: "High",
    source: "SRC-001",
    meaning: "The flow, user, or tenant exceeded its allowed action execution capacity.",
    causes: "High-volume runs, inefficient loops, repeated retries, or insufficient license/capacity for the workload.",
    resolution: "Optimize action count, reduce polling, batch operations, and review licensing or capacity requirements.",
    prevention: "Track action volume and design high-volume processes with capacity limits in mind.",
    tags: "flow run quota exceeded action limit",
  },
  {
    code: "DirectApiAuthorizationRequired",
    status: "",
    product: "Power Automate",
    layer: "Licensing",
    severity: "High",
    source: "SRC-001",
    meaning: "The flow uses a premium connector but the caller or owner lacks the required premium entitlement.",
    causes: "Premium connector used by a seeded-license user, owner lost premium license, or in-context flow became out-of-context.",
    resolution: "Identify the premium connector and assign the appropriate Power Automate Premium or Process license.",
    prevention: "Review premium connector licensing before sharing or embedding flows.",
    tags: "premium connector license direct API authorization",
  },
  {
    code: "Resource not found for the segment",
    status: "404",
    product: "Dataverse",
    layer: "Web API",
    severity: "Medium",
    source: "SRC-002",
    meaning: "Dataverse Web API cannot resolve the entity set, action, or function segment in the request URL.",
    causes: "Wrong entity set name, incorrect casing, wrong custom action name, or inactive custom process action.",
    resolution: "Check the service document and CSDL metadata, use the exact entity set or action name, and confirm custom actions are active.",
    prevention: "Use metadata-driven constants instead of manually typed endpoint names.",
    tags: "dataverse resource segment entity set 404",
  },
  {
    code: "Could not find a property named",
    status: "400",
    product: "Dataverse",
    layer: "Web API",
    severity: "Medium",
    source: "SRC-002",
    meaning: "Dataverse Web API cannot find a requested property on the specified type.",
    causes: "Wrong logical name, display name used instead of schema name, or incorrect property casing.",
    resolution: "Verify the column in CSDL metadata and update the request to use the exact case-sensitive property name.",
    prevention: "Generate request models from metadata where possible.",
    tags: "dataverse property name metadata 400",
  },
  {
    code: "No HTTP resource was found that matches the request URI",
    status: "404",
    product: "Dataverse",
    layer: "Web API",
    severity: "Medium",
    source: "SRC-002",
    meaning: "The request URI and HTTP method do not match a valid Dataverse Web API operation.",
    causes: "Calling a function with POST, calling an action with GET, or mixing action and function semantics.",
    resolution: "Confirm whether the operation is a function or action and send the request with the correct HTTP method.",
    prevention: "Document method requirements for every custom API call.",
    tags: "dataverse HTTP method function action 404",
  },
  {
    code: "Invalid property was found in entity",
    status: "400",
    product: "Dataverse",
    layer: "Web API",
    severity: "Medium",
    source: "SRC-002",
    meaning: "A create or update payload contains a property that Dataverse does not recognize for the target table.",
    causes: "Wrong column logical name, incorrect casing, or display name used in JSON payload.",
    resolution: "Correct the JSON payload to use the logical property name from metadata and retry with a minimal request.",
    prevention: "Use typed SDK classes or generated schemas for write payloads.",
    tags: "dataverse invalid property payload entity",
  },
  {
    code: "Does not support untyped value in non-open type",
    status: "400",
    product: "Dataverse",
    layer: "Web API",
    severity: "Medium",
    source: "SRC-002",
    meaning: "OData payload validation failed because an unexpected property or incompatible value was supplied.",
    causes: "Wrong property name, wrong value type, or data that cannot map to a strongly typed Dataverse field.",
    resolution: "Fix property names and data types, then retry with a minimal payload that contains only known fields.",
    prevention: "Validate JSON payloads against metadata before sending them.",
    tags: "odata untyped value non open type payload",
  },
  {
    code: "Invalid URL",
    status: "400",
    product: "Dataverse",
    layer: "Web API",
    severity: "Medium",
    source: "SRC-003",
    meaning: "The Dataverse Web API request URL is malformed or exceeds a supported URL segment limit.",
    causes: "An OData segment is too long, inline parameter text is excessive, or FetchXML query text is placed in a long GET URL.",
    resolution: "Shorten the URL segment, simplify the query, or move long request content into a supported batch/body pattern.",
    prevention: "Avoid very long inline query strings for complex FetchXML.",
    tags: "invalid URL odata segment length",
  },
  {
    code: "Too Many Requests",
    status: "429",
    product: "Dataverse",
    layer: "Web API",
    severity: "High",
    source: "SRC-004",
    meaning: "Dataverse service protection limits are being triggered by excessive API demand.",
    causes: "Too many requests, too much execution time, high concurrency, aggressive retries, or unbatched high-volume workload.",
    resolution: "Honor Retry-After, reduce concurrency, cache reads, batch carefully, and spread work over time.",
    prevention: "Implement backoff and monitor request volume for integrations.",
    tags: "429 too many requests service protection throttling",
  },
  {
    code: "Precondition Failed",
    status: "412",
    product: "Dataverse",
    layer: "Web API",
    severity: "Medium",
    source: "SRC-003",
    meaning: "A conditional or optimistic-concurrency request failed because the row state no longer matches the provided condition.",
    causes: "Stale ETag, mismatched If-Match header, or row updated by another process.",
    resolution: "Retrieve the latest row and ETag, reconcile changes, and retry with the correct condition.",
    prevention: "Use optimistic concurrency only with clear conflict handling.",
    tags: "412 precondition failed etag concurrency",
  },
  {
    code: "Not Modified",
    status: "304",
    product: "Dataverse",
    layer: "Web API",
    severity: "Low",
    source: "SRC-003",
    meaning: "A conditional retrieval indicates the row has not changed since the supplied condition.",
    causes: "If-None-Match or modified-since condition matched the cached state.",
    resolution: "Treat the response as a cache signal and continue using the existing data.",
    prevention: "Handle 304 separately from actual failure paths.",
    tags: "304 not modified cache conditional",
  },
  {
    code: "Dataverse Forbidden",
    status: "403",
    product: "Dataverse",
    layer: "Security",
    severity: "High",
    source: "SRC-003",
    meaning: "Dataverse rejected the operation because the user or app does not have the required privilege.",
    causes: "Missing table privilege, row access gap, field security restriction, team membership issue, or business unit scope mismatch.",
    resolution: "Review security roles, team membership, field security profiles, and row ownership for the calling identity.",
    prevention: "Test APIs with least-privilege service principals before production release.",
    tags: "dataverse forbidden privilege security role",
  },
  {
    code: "SessionCreationErrorWithThirdPartyCredentialProvider",
    status: "",
    product: "Power Automate Desktop",
    layer: "Desktop flow",
    severity: "High",
    source: "SRC-005",
    meaning: "An unattended desktop flow cannot create a Windows session because a third-party credential provider interferes.",
    causes: "Unsupported credential provider, sign-in extension conflict, or security software that changes Windows credential flow.",
    resolution: "Remove, disable, or reconfigure the third-party credential provider and retest unattended session creation.",
    prevention: "Keep unattended runner machines dedicated and avoid unsupported sign-in extensions.",
    tags: "desktop flow session credential provider",
  },
  {
    code: "SessionCreationUserPromptedForCredentialsAfterConnection",
    status: "",
    product: "Power Automate Desktop",
    layer: "Desktop flow",
    severity: "High",
    source: "SRC-005",
    meaning: "Desktop flow session creation fails because Windows prompts for credentials after connection starts.",
    causes: "Machine policy, credential configuration, RDP setting, or account state causes an interactive credential prompt.",
    resolution: "Review sign-in policies, credentials, unattended connection configuration, and remove prompts that require user action.",
    prevention: "Validate unattended runner login after policy or credential changes.",
    tags: "desktop flow session credential prompt",
  },
  {
    code: "SessionCreationError",
    status: "",
    product: "Power Automate Desktop",
    layer: "Desktop flow",
    severity: "High",
    source: "SRC-005",
    meaning: "A generic session creation failure prevented a desktop flow from starting.",
    causes: "Machine unavailable, bad credentials, RDP restriction, gateway issue, Windows session conflict, or service health problem.",
    resolution: "Check machine registration, credentials, RDP permissions, machine availability, and desktop flow service status.",
    prevention: "Monitor runner health and document machine configuration.",
    tags: "desktop flow session creation generic",
  },
  {
    code: "AttendedUserSessionNotActive",
    status: "",
    product: "Power Automate Desktop",
    layer: "Desktop flow queue",
    severity: "Medium",
    source: "SRC-008",
    meaning: "An attended desktop flow cannot run because the target user session is not active.",
    causes: "User is signed out, locked, disconnected, or not available in the expected interactive session.",
    resolution: "Have the attended user sign in and keep the session active, or redesign the process as unattended.",
    prevention: "Use unattended flows for server-like automation and attended flows only for user-present tasks.",
    tags: "attended user session not active",
  },
  {
    code: "AttendedUserNotLoggedIn",
    status: "",
    product: "Power Automate Desktop",
    layer: "Desktop flow queue",
    severity: "Medium",
    source: "SRC-008",
    meaning: "An attended desktop flow cannot run because the user is not logged in.",
    causes: "No active logged-in Windows session exists for the attended automation user.",
    resolution: "Log in as the target user and confirm the Power Automate service account has the needed machine permissions.",
    prevention: "Schedule attended flows only when the user session is reliably available.",
    tags: "attended user not logged in desktop flow",
  },
  {
    code: "Something went wrong [5objp]",
    status: "",
    product: "Power Apps",
    layer: "Wrap",
    severity: "High",
    source: "SRC-006",
    meaning: "A wrapped Power Apps mobile app fails during authentication even though the build may have succeeded.",
    causes: "APK signature hash mismatch or redirect URI mismatch in Microsoft Entra app registration.",
    resolution: "Generate the correct hash from the signing keystore, update Android authentication settings, and verify redirect URI casing and encoding.",
    prevention: "Control keystore, hash, bundle ID, and redirect URI values in release documentation.",
    tags: "power apps wrap hash redirect 5objp",
  },
  {
    code: "Something went wrong [9n155]",
    status: "",
    product: "Power Apps",
    layer: "Wrap",
    severity: "High",
    source: "SRC-006",
    meaning: "A wrapped Power Apps mobile app fails authentication because tenant registration settings are incorrect.",
    causes: "Microsoft Entra app registration was configured with an unsupported account type such as single tenant when another setting is required.",
    resolution: "Update supported account types in the app registration, rebuild the wrapped app, and retest sign-in.",
    prevention: "Use a wrap checklist for Microsoft Entra registration settings.",
    tags: "power apps wrap tenant 9n155",
  },
  {
    code: "Connector permission error",
    status: "403",
    product: "Power Apps",
    layer: "Canvas app / connector",
    severity: "High",
    source: "SRC-001; SRC-003",
    meaning: "A canvas app connector call fails because the user lacks access or policy blocks the connector.",
    causes: "Missing data-source permission, DLP block, premium connector license gap, or environment policy restriction.",
    resolution: "Check app sharing, data-source permission, DLP policy, connector licensing, and user access.",
    prevention: "Validate app access with representative test users before release.",
    tags: "power apps connector permission canvas app",
  },
  {
    code: "GatewayUnavailable",
    status: "502",
    product: "Power Platform",
    layer: "Gateway",
    severity: "High",
    source: "SRC-001",
    meaning: "The on-premises data gateway or downstream service is unavailable when the platform tries to reach it.",
    causes: "Gateway offline, network path unavailable, service stopped, firewall rule changed, or backend data source unavailable.",
    resolution: "Confirm gateway status, restart gateway services if needed, test network reachability, and validate backend data-source availability.",
    prevention: "Monitor gateway uptime and alert on offline gateway clusters.",
    tags: "gateway unavailable 502 on premises",
  },
  {
    code: "GatewayTimeout",
    status: "504",
    product: "Power Platform",
    layer: "Gateway",
    severity: "Medium",
    source: "SRC-001",
    meaning: "A gateway-backed operation did not complete before the service timeout.",
    causes: "Slow query, overloaded gateway machine, network latency, large payload, or backend system contention.",
    resolution: "Optimize the query, reduce payload size, check gateway CPU and memory, and retry during normal backend performance.",
    prevention: "Set performance baselines for gateway-backed sources.",
    tags: "gateway timeout 504 slow query",
  },
  {
    code: "DependencyMissing",
    status: "",
    product: "Power Platform",
    layer: "Solutions",
    severity: "High",
    source: "SRC-001",
    meaning: "A solution component cannot be imported or activated because another required component is missing.",
    causes: "Missing table, app, flow, plugin step, connection reference, environment variable, or publisher dependency in the target environment.",
    resolution: "Review the import error details, install or include the missing dependency, and rerun the import.",
    prevention: "Export complete solution layers and validate dependencies before deployment.",
    tags: "solution import dependency missing",
  },
  {
    code: "EnvironmentVariableMissing",
    status: "",
    product: "Power Platform",
    layer: "Solutions",
    severity: "High",
    source: "SRC-001",
    meaning: "A flow, app, or connector reference expects an environment variable value that is not set.",
    causes: "Solution imported without current values or deployment pipeline did not populate environment-specific settings.",
    resolution: "Set the environment variable current value in the target environment and rerun the app or flow.",
    prevention: "Maintain deployment parameter files for every environment.",
    tags: "environment variable missing solution deployment",
  },
  {
    code: "DlpPolicyViolation",
    status: "403",
    product: "Power Platform",
    layer: "DLP policy",
    severity: "High",
    source: "SRC-001",
    meaning: "A Data Loss Prevention policy blocks the connector or connector combination used by the app or flow.",
    causes: "Connector assigned to blocked group, business/nonbusiness group conflict, or environment added to a restrictive policy.",
    resolution: "Review the environment DLP policies, identify the blocked connector, and request an approved policy change if justified.",
    prevention: "Check DLP impact before adding connectors to production solutions.",
    tags: "DLP policy violation connector blocked",
  },
  {
    code: "PremiumLicenseRequired",
    status: "",
    product: "Power Platform",
    layer: "Licensing",
    severity: "High",
    source: "SRC-001",
    meaning: "The app, flow, or connector requires a premium license that the user, owner, or process does not have.",
    causes: "Premium connector, premium Dataverse usage, custom connector, HTTP action, or process licensing mismatch.",
    resolution: "Identify the premium feature and assign the correct Power Apps, Power Automate, or Process license.",
    prevention: "Review licensing impact before promoting solutions to shared users.",
    tags: "premium license required connector",
  },
  {
    code: "RateLimitExceeded",
    status: "429",
    product: "Power Platform",
    layer: "Connector/API",
    severity: "Medium",
    source: "SRC-001; SRC-004",
    meaning: "A connector, API, or downstream service rate limit was exceeded.",
    causes: "High request volume, looped connector calls, aggressive retries, concurrency spikes, or shared connector throttling.",
    resolution: "Reduce concurrency, add backoff, batch requests, cache repeated reads, and distribute work over time.",
    prevention: "Design integrations around documented limits and monitor retry counts.",
    tags: "rate limit exceeded 429 connector throttling",
  },
  {
    code: "RequestEntityTooLarge",
    status: "413",
    product: "Power Platform",
    layer: "Connector/API",
    severity: "Medium",
    source: "SRC-001; SRC-003",
    meaning: "The submitted payload is larger than the service or connector accepts.",
    causes: "Large attachment, oversized JSON body, too many records in one call, or base64 content passed through an action.",
    resolution: "Reduce file size, chunk the payload, upload content through a supported file API, or process records in batches.",
    prevention: "Set payload-size checks before connector calls.",
    tags: "413 payload too large request entity",
  },
  {
    code: "InternalServerError",
    status: "500",
    product: "Power Platform",
    layer: "Connector/API",
    severity: "Medium",
    source: "SRC-001; SRC-003",
    meaning: "The downstream service or connector returned a generic server error.",
    causes: "Temporary service fault, backend exception, connector bug, malformed edge-case payload, or downstream outage.",
    resolution: "Check service health, retry once with correlation details, simplify the payload, and escalate with request IDs if repeatable.",
    prevention: "Log correlation IDs and use retry policies only where the operation is idempotent.",
    tags: "500 internal server error connector",
  },
  {
    code: "ServiceUnavailable",
    status: "503",
    product: "Power Platform",
    layer: "Connector/API",
    severity: "Medium",
    source: "SRC-001; SRC-003",
    meaning: "The target service is temporarily unavailable or overloaded.",
    causes: "Service outage, planned maintenance, connector transient failure, or downstream capacity issue.",
    resolution: "Check service health, honor retry guidance, rerun after service recovery, and avoid repeated manual retries.",
    prevention: "Add transient-failure handling and alerting for repeated 503 responses.",
    tags: "503 service unavailable transient",
  },
  {
    code: "Conflict",
    status: "409",
    product: "Power Platform",
    layer: "Connector/API",
    severity: "Medium",
    source: "SRC-003",
    meaning: "The request conflicts with the current state of the target resource.",
    causes: "Duplicate key, concurrent update, existing resource with same identifier, or state transition not allowed.",
    resolution: "Refresh the resource state, resolve the duplicate or conflict, and retry with the expected identifier or state.",
    prevention: "Use idempotency and lookup-before-create patterns.",
    tags: "409 conflict duplicate concurrent update",
  },
  {
    code: "MethodNotAllowed",
    status: "405",
    product: "Power Platform",
    layer: "Connector/API",
    severity: "Medium",
    source: "SRC-002; SRC-003",
    meaning: "The endpoint exists but does not support the HTTP method used by the request.",
    causes: "GET used where POST is required, PATCH used where PUT is expected, or custom API operation type misunderstood.",
    resolution: "Check the operation documentation and send the request with the correct method and headers.",
    prevention: "Keep API operation definitions and connector actions aligned.",
    tags: "405 method not allowed HTTP",
  },
  {
    code: "UnsupportedMediaType",
    status: "415",
    product: "Power Platform",
    layer: "Connector/API",
    severity: "Medium",
    source: "SRC-003",
    meaning: "The service rejected the payload content type.",
    causes: "Missing or wrong Content-Type header, binary sent where JSON is expected, or unsupported file format.",
    resolution: "Set the expected Content-Type, convert the payload to the supported format, and retry.",
    prevention: "Validate headers in custom connectors and HTTP actions.",
    tags: "415 unsupported media type content type",
  },
  {
    code: "ParseJsonSchemaValidationFailed",
    status: "",
    product: "Power Automate",
    layer: "Runtime expression",
    severity: "Medium",
    source: "SRC-001",
    meaning: "A Parse JSON action could not validate the incoming payload against its schema.",
    causes: "Missing property, null where object expected, array/object mismatch, or schema generated from incomplete sample data.",
    resolution: "Inspect the actual payload, update the schema to allow optional or nullable values, and retest.",
    prevention: "Generate schemas from multiple representative payloads.",
    tags: "parse json schema validation failed",
  },
  {
    code: "ApplyToEachConcurrencyConflict",
    status: "",
    product: "Power Automate",
    layer: "Runtime expression",
    severity: "Medium",
    source: "SRC-001",
    meaning: "Parallel loop execution causes conflicting updates or unexpected data order.",
    causes: "Concurrency enabled while updating shared variables, same row, same file, or order-dependent state.",
    resolution: "Turn off concurrency or redesign the loop so each iteration writes to an isolated resource.",
    prevention: "Use concurrency only for independent operations.",
    tags: "apply to each concurrency conflict",
  },
  {
    code: "NullReference",
    status: "",
    product: "Power Platform",
    layer: "Runtime expression",
    severity: "Medium",
    source: "SRC-001",
    meaning: "A formula, expression, or connector output references a value that is null or missing.",
    causes: "Optional field missing, lookup returned no rows, connector output changed, or schema did not include the property.",
    resolution: "Add null checks, default values, and existence checks before reading nested properties.",
    prevention: "Treat external data as optional unless the source guarantees it.",
    tags: "null reference missing value",
  },
  {
    code: "InvalidAuthenticationToken",
    status: "401",
    product: "Power Platform",
    layer: "Authentication",
    severity: "High",
    source: "SRC-001; SRC-003",
    meaning: "The request includes an authentication token that the target service cannot validate.",
    causes: "Expired token, wrong audience, missing consent, incorrect tenant, or stale connector authentication.",
    resolution: "Refresh the connection, confirm the token audience and tenant, grant required consent, and retry with a valid identity.",
    prevention: "Use monitored service identities and review app registration permissions before release.",
    tags: "invalid authentication token 401 audience tenant consent",
  },
  {
    code: "InsufficientPrivileges",
    status: "403",
    product: "Dataverse",
    layer: "Security",
    severity: "High",
    source: "SRC-003",
    meaning: "Dataverse blocks the operation because the caller lacks one or more required privileges.",
    causes: "Missing security role privilege, inadequate business unit scope, row ownership mismatch, or field security restriction.",
    resolution: "Review the caller security roles, team membership, business unit scope, and field security profile assignments.",
    prevention: "Test Dataverse actions with least-privilege users before promoting the solution.",
    tags: "dataverse insufficient privileges security role",
  },
  {
    code: "PrincipalUserNotFound",
    status: "404",
    product: "Dataverse",
    layer: "Security",
    severity: "Medium",
    source: "SRC-003",
    meaning: "The platform cannot resolve the user, owner, or principal referenced by the operation.",
    causes: "Deleted user, disabled account, missing application user, bad owner ID, or environment copy with stale user references.",
    resolution: "Confirm the user or application user exists in the environment and update owner, caller, or connection references.",
    prevention: "Validate application users and owner mappings after environment copy or restore.",
    tags: "principal user not found dataverse owner",
  },
  {
    code: "RecordIsUnavailable",
    status: "404",
    product: "Dataverse",
    layer: "Data access",
    severity: "Medium",
    source: "SRC-003",
    meaning: "A referenced Dataverse record is unavailable to the caller or no longer exists.",
    causes: "Deleted row, archived record, missing permission, wrong environment, or stale lookup reference.",
    resolution: "Verify the row exists in the current environment and confirm the caller has read access to it.",
    prevention: "Avoid storing stale row IDs and add lookup validation before updates.",
    tags: "record unavailable dataverse row lookup",
  },
  {
    code: "DuplicateRecord",
    status: "409",
    product: "Dataverse",
    layer: "Data validation",
    severity: "Medium",
    source: "SRC-003",
    meaning: "The operation would create a duplicate record or violates a uniqueness rule.",
    causes: "Duplicate alternate key, duplicate detection rule, repeated create request, or missing idempotency check.",
    resolution: "Search for the existing row, update it if appropriate, or change the key value before retrying.",
    prevention: "Use lookup-before-create logic and stable alternate keys.",
    tags: "duplicate record dataverse alternate key",
  },
  {
    code: "DuplicateKey",
    status: "409",
    product: "Dataverse",
    layer: "Data validation",
    severity: "Medium",
    source: "SRC-003",
    meaning: "A Dataverse write conflicts with an existing alternate key value.",
    causes: "Attempted insert with a key already used by another row or concurrent create using the same key.",
    resolution: "Find the row with the existing key, decide whether to update or skip, and retry only if the key is unique.",
    prevention: "Implement idempotent upsert patterns for integration loads.",
    tags: "duplicate key alternate key dataverse upsert",
  },
  {
    code: "CannotDeleteDueToAssociation",
    status: "400",
    product: "Dataverse",
    layer: "Data relationship",
    severity: "Medium",
    source: "SRC-003",
    meaning: "A row cannot be deleted because related records or relationship rules block deletion.",
    causes: "Restrict delete relationship behavior, child records still exist, or cascade rules prevent the operation.",
    resolution: "Review related records and relationship behavior, remove or reassign child rows, then retry deletion.",
    prevention: "Document cascade behavior for tables used by automation.",
    tags: "cannot delete association dataverse relationship",
  },
  {
    code: "BusinessProcessError",
    status: "400",
    product: "Dataverse",
    layer: "Business rules",
    severity: "Medium",
    source: "SRC-003",
    meaning: "A business rule, plugin, workflow, or process validation blocked the operation.",
    causes: "Custom validation failed, required business data missing, plugin exception, or business process flow stage rule.",
    resolution: "Review the error details, plugin trace logs, process rules, and required fields for the failing operation.",
    prevention: "Add integration-safe validation messages and maintain plugin trace logging in test environments.",
    tags: "business process error plugin validation",
  },
  {
    code: "PluginExecutionException",
    status: "500",
    product: "Dataverse",
    layer: "Plugin",
    severity: "High",
    source: "SRC-003",
    meaning: "A Dataverse plugin or custom server-side extension failed during execution.",
    causes: "Unhandled plugin exception, bad input data, missing configuration, recursion, or external dependency failure.",
    resolution: "Check plugin trace logs, correlation IDs, secure/unsecure configuration, and recent plugin deployment changes.",
    prevention: "Use defensive plugin code and automated tests for registered steps.",
    tags: "plugin execution exception dataverse trace",
  },
  {
    code: "SandboxWorkerNotAvailable",
    status: "503",
    product: "Dataverse",
    layer: "Plugin",
    severity: "Medium",
    source: "SRC-003",
    meaning: "The sandbox worker needed to run isolated Dataverse code is temporarily unavailable.",
    causes: "Transient platform issue, overloaded sandbox process, plugin crash, or service availability problem.",
    resolution: "Retry after a short delay, check service health, and inspect plugin trace logs for repeat failures.",
    prevention: "Keep plugin execution efficient and monitor repeated sandbox failures.",
    tags: "sandbox worker not available plugin",
  },
  {
    code: "OrganizationServiceFault",
    status: "500",
    product: "Dataverse",
    layer: "SDK",
    severity: "Medium",
    source: "SRC-003",
    meaning: "A Dataverse SDK call returned a service fault.",
    causes: "Server-side validation, security issue, malformed request, plugin exception, or transient Dataverse fault.",
    resolution: "Capture the fault code, message, trace text, and correlation ID, then map it to the underlying cause.",
    prevention: "Log SDK fault details and handle retryable faults separately from validation failures.",
    tags: "organization service fault dataverse sdk",
  },
  {
    code: "CrmServiceClientLoginFailed",
    status: "401",
    product: "Dataverse",
    layer: "SDK authentication",
    severity: "High",
    source: "SRC-003",
    meaning: "A Dataverse SDK client cannot authenticate to the environment.",
    causes: "Invalid client secret, expired certificate, wrong authority, disabled app user, or missing API permissions.",
    resolution: "Validate the connection string, app registration credentials, tenant, URL, and application user security roles.",
    prevention: "Track app credential expiry and rotate secrets before production failure.",
    tags: "crm service client login failed sdk",
  },
  {
    code: "InvalidPluginRegistration",
    status: "400",
    product: "Dataverse",
    layer: "Plugin",
    severity: "High",
    source: "SRC-003",
    meaning: "A plugin assembly, step, image, or message registration is invalid.",
    causes: "Wrong message, missing entity name, invalid image attributes, assembly version mismatch, or deployment packaging issue.",
    resolution: "Review the plugin registration step, entity, message, stage, execution mode, and assembly version.",
    prevention: "Automate plugin registration deployment and validate registrations after import.",
    tags: "invalid plugin registration dataverse",
  },
  {
    code: "SolutionImportFailed",
    status: "",
    product: "Power Platform",
    layer: "Solutions",
    severity: "High",
    source: "SRC-001",
    meaning: "A solution import failed before all components were applied successfully.",
    causes: "Missing dependency, incompatible managed layer, connection reference issue, environment variable issue, or component conflict.",
    resolution: "Open the solution import log, identify the first blocking component, resolve dependencies, and rerun the import.",
    prevention: "Run solution checker and dependency validation before production deployment.",
    tags: "solution import failed deployment",
  },
  {
    code: "SolutionUpgradeFailed",
    status: "",
    product: "Power Platform",
    layer: "Solutions",
    severity: "High",
    source: "SRC-001",
    meaning: "A managed solution upgrade or stage-for-upgrade operation failed.",
    causes: "Layer conflict, missing base solution, deleted component, dependency gap, or unsupported upgrade path.",
    resolution: "Review solution history, import logs, component layers, and apply the correct base/patch/upgrade order.",
    prevention: "Test managed upgrades in a staging environment that mirrors production layers.",
    tags: "solution upgrade failed managed layer",
  },
  {
    code: "PublisherPrefixMismatch",
    status: "",
    product: "Power Platform",
    layer: "Solutions",
    severity: "Medium",
    source: "SRC-001",
    meaning: "A component name or customization conflicts with expected publisher prefix rules.",
    causes: "Manual component creation under the wrong publisher, imported unmanaged component, or naming mismatch.",
    resolution: "Confirm publisher settings and recreate or rename components under the correct solution publisher where possible.",
    prevention: "Create solution components only inside the intended publisher solution.",
    tags: "publisher prefix mismatch solution",
  },
  {
    code: "ConnectionReferenceMissing",
    status: "",
    product: "Power Platform",
    layer: "Solutions",
    severity: "High",
    source: "SRC-001",
    meaning: "A solution component expects a connection reference that is missing or unmapped.",
    causes: "Connection reference excluded from solution, import mapping skipped, or target environment lacks the connector connection.",
    resolution: "Add or map the connection reference, create the target connection, and rerun the import or flow activation.",
    prevention: "Include all connection references in deployment checklists.",
    tags: "connection reference missing solution",
  },
  {
    code: "FlowActivationFailed",
    status: "",
    product: "Power Automate",
    layer: "Flow lifecycle",
    severity: "High",
    source: "SRC-001",
    meaning: "A flow could not be turned on after edit, import, or deployment.",
    causes: "Invalid connection reference, missing required field, premium license issue, invalid trigger, or DLP policy block.",
    resolution: "Open the flow details, review activation errors, fix connections, validate trigger configuration, and save again.",
    prevention: "Activate and test every imported flow as part of release validation.",
    tags: "flow activation failed turn on",
  },
  {
    code: "FlowSuspended",
    status: "",
    product: "Power Automate",
    layer: "Flow lifecycle",
    severity: "High",
    source: "SRC-001",
    meaning: "The flow is suspended and cannot run until the owner or admin resolves the issue.",
    causes: "Repeated failures, expired connection, license issue, policy violation, or owner account problem.",
    resolution: "Review suspension reason, fix the underlying connection, policy, or license issue, and turn the flow back on.",
    prevention: "Monitor critical flows for suspension and repeated failure trends.",
    tags: "flow suspended disabled owner",
  },
  {
    code: "TriggerDisabled",
    status: "",
    product: "Power Automate",
    layer: "Trigger",
    severity: "Medium",
    source: "SRC-001",
    meaning: "The trigger is disabled or unavailable, so new runs are not starting.",
    causes: "Flow turned off, trigger connection invalid, admin policy block, or trigger configuration invalid after import.",
    resolution: "Turn the flow on, fix trigger connection and configuration, and confirm a new event starts a run.",
    prevention: "Include trigger activation checks in deployment validation.",
    tags: "trigger disabled flow not running",
  },
  {
    code: "WebhookRegistrationFailed",
    status: "400",
    product: "Power Automate",
    layer: "Trigger",
    severity: "High",
    source: "SRC-001",
    meaning: "The platform could not register a webhook subscription for a trigger or callback.",
    causes: "Invalid callback URL, connector subscription limit, missing permission, expired connection, or service-specific webhook failure.",
    resolution: "Fix the connection, recreate the trigger subscription by saving the flow, and check connector-specific limits.",
    prevention: "Monitor trigger registration errors after deployments and connection changes.",
    tags: "webhook registration failed trigger",
  },
  {
    code: "InvalidPaginationPolicy",
    status: "",
    product: "Power Automate",
    layer: "Connector/API",
    severity: "Medium",
    source: "SRC-001",
    meaning: "A pagination setting is invalid or cannot be applied to the connector action.",
    causes: "Pagination threshold too high, action does not support pagination, or connector response lacks expected paging tokens.",
    resolution: "Review action settings, lower the threshold, and confirm the connector supports pagination for that operation.",
    prevention: "Test high-volume list actions with realistic production data sizes.",
    tags: "invalid pagination policy flow",
  },
  {
    code: "PaginationThresholdExceeded",
    status: "",
    product: "Power Automate",
    layer: "Connector/API",
    severity: "Medium",
    source: "SRC-001",
    meaning: "The action cannot return more records because it reached the configured or service pagination threshold.",
    causes: "Large list query, insufficient filtering, connector cap, or pagination disabled.",
    resolution: "Add source filters, enable supported pagination, process records in batches, or split the workload.",
    prevention: "Design list queries to retrieve only required records.",
    tags: "pagination threshold exceeded flow",
  },
  {
    code: "ConcurrencyLimitExceeded",
    status: "429",
    product: "Power Platform",
    layer: "Throttling",
    severity: "Medium",
    source: "SRC-001; SRC-004",
    meaning: "Too many operations are running concurrently for the connector, environment, user, or service.",
    causes: "High parallelism, multiple flows sharing the same connection, nested loops, or aggressive retry settings.",
    resolution: "Reduce degree of parallelism, stagger schedules, use queues, and tune retry policies.",
    prevention: "Set concurrency limits deliberately for high-volume automations.",
    tags: "concurrency limit exceeded throttling",
  },
  {
    code: "RetryPolicyExceeded",
    status: "",
    product: "Power Automate",
    layer: "Connector/API",
    severity: "Medium",
    source: "SRC-001",
    meaning: "An action exhausted its retry policy without receiving a successful response.",
    causes: "Persistent downstream failure, throttling, authentication issue, or non-transient validation error retried repeatedly.",
    resolution: "Inspect the final and first failure messages, fix the root cause, and adjust retry policy only for transient errors.",
    prevention: "Use retry policies with clear logging and avoid retrying validation failures.",
    tags: "retry policy exceeded flow",
  },
  {
    code: "InvalidODataFilter",
    status: "400",
    product: "Power Platform",
    layer: "Connector/API",
    severity: "Medium",
    source: "SRC-001; SRC-003",
    meaning: "An OData filter expression could not be parsed or executed by the target service.",
    causes: "Wrong column name, incorrect operator, bad quote escaping, unsupported function, or data type mismatch.",
    resolution: "Validate column names and data types, simplify the filter, escape quotes correctly, and test the query in isolation.",
    prevention: "Keep OData examples with each flow and avoid hand-built filter strings where possible.",
    tags: "invalid odata filter query",
  },
  {
    code: "InvalidFetchXml",
    status: "400",
    product: "Dataverse",
    layer: "Query",
    severity: "Medium",
    source: "SRC-003",
    meaning: "Dataverse could not parse or execute the FetchXML query.",
    causes: "Malformed XML, wrong table or column logical name, invalid link-entity, unsupported operator, or too-long query URL.",
    resolution: "Validate the FetchXML, confirm logical names, test the query in a Dataverse tool, and move long queries into supported request bodies.",
    prevention: "Store tested FetchXML snippets and validate them before embedding in flows.",
    tags: "invalid fetchxml dataverse query",
  },
  {
    code: "InvalidColumnName",
    status: "400",
    product: "Dataverse",
    layer: "Query",
    severity: "Medium",
    source: "SRC-002; SRC-003",
    meaning: "A query, payload, or expression references a column name that is not valid for the table.",
    causes: "Display name used instead of logical name, renamed column, typo, or wrong table context.",
    resolution: "Check table metadata and update the request to use the exact logical column name.",
    prevention: "Use metadata lookup or generated constants for column names.",
    tags: "invalid column name dataverse",
  },
  {
    code: "InvalidEntityName",
    status: "400",
    product: "Dataverse",
    layer: "Query",
    severity: "Medium",
    source: "SRC-002; SRC-003",
    meaning: "The request references a Dataverse table or entity name that is not valid in the current environment.",
    causes: "Wrong logical name, wrong entity set name, table not installed, or environment mismatch.",
    resolution: "Confirm the table exists and use the correct logical name or entity set name from metadata.",
    prevention: "Avoid copying endpoint names between environments without metadata validation.",
    tags: "invalid entity name dataverse table",
  },
  {
    code: "InvalidLookupReference",
    status: "400",
    product: "Dataverse",
    layer: "Data relationship",
    severity: "Medium",
    source: "SRC-003",
    meaning: "A lookup value does not point to a valid related row or is formatted incorrectly.",
    causes: "Wrong entity set binding, missing related row, bad GUID, or lookup target type mismatch.",
    resolution: "Validate the related row ID, target table, and @odata.bind syntax before retrying.",
    prevention: "Resolve lookup rows dynamically instead of hardcoding GUIDs.",
    tags: "invalid lookup reference odata bind",
  },
  {
    code: "RequiredFieldMissing",
    status: "400",
    product: "Power Platform",
    layer: "Data validation",
    severity: "Medium",
    source: "SRC-001; SRC-003",
    meaning: "A required field or required connector input was not supplied.",
    causes: "Missing form value, blank dynamic content, unmapped field, or environment variable missing in a solution.",
    resolution: "Identify the required input, supply a valid value, and add a guard for blank dynamic content.",
    prevention: "Validate required fields before submit or connector calls.",
    tags: "required field missing validation",
  },
  {
    code: "InvalidChoiceValue",
    status: "400",
    product: "Dataverse",
    layer: "Data validation",
    severity: "Medium",
    source: "SRC-003",
    meaning: "A choice, option set, status, or state value is not valid for the target column.",
    causes: "Label used instead of integer value, option not present in environment, status transition not allowed, or metadata mismatch.",
    resolution: "Check choice metadata and submit the supported numeric value or valid transition.",
    prevention: "Map choices from metadata rather than hardcoding labels.",
    tags: "invalid choice value option set",
  },
  {
    code: "InvalidDateTimeFormat",
    status: "400",
    product: "Power Platform",
    layer: "Data validation",
    severity: "Medium",
    source: "SRC-001",
    meaning: "A date or time value is not in a format accepted by the target connector or service.",
    causes: "Locale mismatch, ambiguous date string, missing timezone, or text value passed where ISO date is expected.",
    resolution: "Normalize the value with explicit date parsing and output an ISO-style format expected by the connector.",
    prevention: "Store dates in consistent formats and include timezone handling in flows.",
    tags: "invalid datetime format date parse",
  },
  {
    code: "InvalidNumberFormat",
    status: "400",
    product: "Power Platform",
    layer: "Data validation",
    severity: "Medium",
    source: "SRC-001",
    meaning: "A numeric field receives text or a number formatted in a way the service cannot parse.",
    causes: "Currency symbols, comma separators, locale decimal markers, blank strings, or nonnumeric characters.",
    resolution: "Strip formatting characters, validate the value, and convert it explicitly before submission.",
    prevention: "Validate numeric inputs at the app or flow boundary.",
    tags: "invalid number format conversion",
  },
  {
    code: "FileLocked",
    status: "423",
    product: "Power Platform",
    layer: "Connector/API",
    severity: "Medium",
    source: "SRC-001",
    meaning: "A connector cannot update or read a file because it is locked by another user or process.",
    causes: "Open workbook, active coauthoring session, sync lock, previous connector action still processing, or retention hold.",
    resolution: "Close the file, wait for the lock to clear, retry later, or redesign the process to avoid simultaneous writes.",
    prevention: "Use queues or Dataverse for concurrent writes instead of shared files.",
    tags: "file locked connector excel sharepoint",
  },
  {
    code: "FileNotFound",
    status: "404",
    product: "Power Platform",
    layer: "Connector/API",
    severity: "Medium",
    source: "SRC-001",
    meaning: "A connector cannot find the referenced file.",
    causes: "Moved file, renamed file, deleted file, wrong site or drive, stale file identifier, or permission gap.",
    resolution: "Locate the file, update the connector reference or file ID, and confirm the connection user has access.",
    prevention: "Use dynamic file lookups and avoid brittle path-only references.",
    tags: "file not found connector sharepoint onedrive",
  },
  {
    code: "InvalidFileIdentifier",
    status: "400",
    product: "Power Platform",
    layer: "Connector/API",
    severity: "Medium",
    source: "SRC-001",
    meaning: "The file identifier passed to a connector action is not valid for that connector operation.",
    causes: "Path used where ID is expected, ID from another connector, stale dynamic content, or file moved between libraries.",
    resolution: "Use the identifier output from the matching connector action and refresh stale dynamic content.",
    prevention: "Keep file ID and file path values clearly separated in flow design.",
    tags: "invalid file identifier connector",
  },
  {
    code: "InvalidAdaptiveCardJson",
    status: "400",
    product: "Power Automate",
    layer: "Teams / adaptive cards",
    severity: "Medium",
    source: "SRC-001",
    meaning: "The adaptive card JSON is malformed or unsupported by the target host.",
    causes: "Invalid JSON, unsupported schema version, missing required card fields, or dynamic content breaking JSON escaping.",
    resolution: "Validate the card JSON, lower the schema version if needed, and escape dynamic content correctly.",
    prevention: "Test cards with realistic dynamic values before deployment.",
    tags: "invalid adaptive card json teams",
  },
  {
    code: "AdaptiveCardResponseTimedOut",
    status: "",
    product: "Power Automate",
    layer: "Teams / adaptive cards",
    severity: "Low",
    source: "SRC-001",
    meaning: "The flow did not receive an adaptive card response before the wait expired.",
    causes: "Recipient did not respond, card was not delivered, notification was missed, or timeout was too short.",
    resolution: "Check delivery, extend timeout if appropriate, and add a timeout branch for non-response.",
    prevention: "Design approval and card workflows with explicit non-response handling.",
    tags: "adaptive card response timed out",
  },
  {
    code: "ApprovalNotFound",
    status: "404",
    product: "Power Automate",
    layer: "Approvals",
    severity: "Medium",
    source: "SRC-001",
    meaning: "The approval action cannot find the approval record it expects to update or wait on.",
    causes: "Approval deleted, wrong environment, stale approval ID, or action sequence changed after a run started.",
    resolution: "Confirm the approval exists, use the correct approval ID, and rerun the flow from a clean trigger event.",
    prevention: "Avoid manually deleting active approvals used by production flows.",
    tags: "approval not found flow",
  },
  {
    code: "ApprovalResponseInvalid",
    status: "400",
    product: "Power Automate",
    layer: "Approvals",
    severity: "Medium",
    source: "SRC-001",
    meaning: "The approval response does not match what later flow actions expect.",
    causes: "Unexpected outcome value, reassigned approval, missing comments, or schema mismatch after approval action changes.",
    resolution: "Inspect the approval output, update expressions to match the actual response schema, and handle all outcomes.",
    prevention: "Use explicit branches for approve, reject, cancel, and timeout outcomes.",
    tags: "approval response invalid schema",
  },
  {
    code: "MailboxNotEnabledForRESTAPI",
    status: "403",
    product: "Power Platform",
    layer: "Outlook connector",
    severity: "Medium",
    source: "SRC-001",
    meaning: "The Outlook connector cannot access the mailbox through the expected API.",
    causes: "Mailbox not fully provisioned, shared mailbox permissions missing, account disabled, or Exchange policy restriction.",
    resolution: "Confirm mailbox status, API access, shared mailbox permission, and connector authentication.",
    prevention: "Validate mailbox access before using it in production automations.",
    tags: "mailbox not enabled rest api outlook",
  },
  {
    code: "InvalidSharedMailbox",
    status: "400",
    product: "Power Platform",
    layer: "Outlook connector",
    severity: "Medium",
    source: "SRC-001",
    meaning: "The shared mailbox value or access configuration is invalid for the connector action.",
    causes: "Wrong mailbox address, missing delegate permission, unlicensed account restrictions, or stale connection.",
    resolution: "Verify shared mailbox address and delegate permissions, then refresh the Outlook connection.",
    prevention: "Document shared mailbox access requirements for flow owners.",
    tags: "invalid shared mailbox outlook connector",
  },
  {
    code: "SharePointListNotFound",
    status: "404",
    product: "Power Platform",
    layer: "SharePoint connector",
    severity: "Medium",
    source: "SRC-001",
    meaning: "The SharePoint connector cannot find the configured list or library.",
    causes: "List renamed, deleted, moved to another site, connection user lacks access, or environment variable points to the wrong site.",
    resolution: "Confirm the site and list exist, update the connector action, and verify permissions for the connection user.",
    prevention: "Use environment variables for site/list references and validate them after deployment.",
    tags: "sharepoint list not found connector",
  },
  {
    code: "SharePointItemNotFound",
    status: "404",
    product: "Power Platform",
    layer: "SharePoint connector",
    severity: "Medium",
    source: "SRC-001",
    meaning: "The SharePoint connector cannot find the list item or file item referenced by the flow.",
    causes: "Item deleted, wrong item ID, item moved, filter returned no rows, or permissions changed.",
    resolution: "Confirm the item exists, refresh the ID lookup, and add handling for empty query results.",
    prevention: "Avoid hardcoded item IDs unless the item is governed as configuration data.",
    tags: "sharepoint item not found connector",
  },
  {
    code: "ExcelTableNotFound",
    status: "404",
    product: "Power Platform",
    layer: "Excel connector",
    severity: "Medium",
    source: "SRC-001",
    meaning: "The Excel connector cannot find the named table in the workbook.",
    causes: "Workbook has no formatted table, table renamed, file copied without expected structure, or connector cached stale metadata.",
    resolution: "Open the workbook, confirm the table name, refresh the connector action, and select the correct table.",
    prevention: "Use stable table names and avoid manual workbook structure changes.",
    tags: "excel table not found connector",
  },
  {
    code: "ExcelRowNotFound",
    status: "404",
    product: "Power Platform",
    layer: "Excel connector",
    severity: "Medium",
    source: "SRC-001",
    meaning: "The Excel connector cannot find the row matching the supplied key value.",
    causes: "Wrong key column, blank key value, duplicate keys, row deleted, or workbook not refreshed.",
    resolution: "Check the key column and key value, confirm the row exists, and handle no-match results.",
    prevention: "Use unique key columns and validate inputs before Excel row actions.",
    tags: "excel row not found connector",
  },
  {
    code: "SqlConnectionFailed",
    status: "",
    product: "Power Platform",
    layer: "SQL connector",
    severity: "High",
    source: "SRC-001",
    meaning: "The SQL connector cannot establish a connection to the configured SQL data source.",
    causes: "Bad credentials, gateway issue, firewall rule, server unavailable, database renamed, or TLS/network configuration problem.",
    resolution: "Test the SQL connection, gateway, firewall, credentials, server name, and database availability.",
    prevention: "Monitor gateway and SQL connectivity used by production apps and flows.",
    tags: "sql connection failed connector gateway",
  },
  {
    code: "SqlQueryTimeout",
    status: "",
    product: "Power Platform",
    layer: "SQL connector",
    severity: "Medium",
    source: "SRC-001",
    meaning: "A SQL query or stored procedure did not complete before the connector timeout.",
    causes: "Slow query plan, blocking, missing indexes, large result set, overloaded gateway, or long-running stored procedure.",
    resolution: "Optimize the query, add filters or indexes, reduce returned rows, and check gateway and SQL performance.",
    prevention: "Performance-test SQL connector operations with production-like data volumes.",
    tags: "sql query timeout connector",
  },
  {
    code: "SqlStoredProcedureFailed",
    status: "500",
    product: "Power Platform",
    layer: "SQL connector",
    severity: "Medium",
    source: "SRC-001",
    meaning: "A SQL stored procedure called by Power Platform returned an error.",
    causes: "Procedure exception, missing parameter, permission issue, transaction conflict, or unexpected input value.",
    resolution: "Run the stored procedure with the same inputs in SQL tools and correct parameters, permissions, or procedure logic.",
    prevention: "Validate stored procedure inputs and return predictable error messages.",
    tags: "sql stored procedure failed connector",
  },
  {
    code: "PowerPagesTablePermissionDenied",
    status: "403",
    product: "Power Pages",
    layer: "Security",
    severity: "High",
    source: "SRC-003",
    meaning: "A Power Pages user cannot access Dataverse data because table permissions do not allow it.",
    causes: "Missing table permission, missing web role, wrong contact relationship, or site cache not refreshed.",
    resolution: "Review table permissions, web roles, contact mapping, and clear site cache after changes.",
    prevention: "Test portal access with representative external users.",
    tags: "power pages table permission denied",
  },
  {
    code: "PowerPagesAuthenticationFailed",
    status: "401",
    product: "Power Pages",
    layer: "Authentication",
    severity: "High",
    source: "SRC-003",
    meaning: "A Power Pages sign-in or identity provider flow failed.",
    causes: "Identity provider configuration issue, redirect URI mismatch, client secret expiry, user not mapped, or site setting mismatch.",
    resolution: "Validate identity provider settings, redirect URIs, secrets, site settings, and contact mapping.",
    prevention: "Track identity provider secret expiry and test sign-in after site configuration changes.",
    tags: "power pages authentication failed identity provider",
  },
  {
    code: "PowerPagesSiteSettingMissing",
    status: "",
    product: "Power Pages",
    layer: "Configuration",
    severity: "Medium",
    source: "SRC-003",
    meaning: "A required Power Pages site setting is missing or has an invalid value.",
    causes: "Deleted site setting, typo in setting name, solution deployment gap, or environment-specific value missing.",
    resolution: "Add or correct the site setting and clear site cache before retesting.",
    prevention: "Include required site settings in deployment documentation.",
    tags: "power pages site setting missing",
  },
  {
    code: "CanvasAppFormulaError",
    status: "",
    product: "Power Apps",
    layer: "Canvas app formula",
    severity: "Medium",
    source: "SRC-001",
    meaning: "A Power Fx formula fails or returns an error during app use.",
    causes: "Blank value, type mismatch, delegation issue, invalid control reference, or connector output shape change.",
    resolution: "Use formula error indicators, inspect variables and connector outputs, add IsBlank or IfError handling, and retest.",
    prevention: "Add defensive Power Fx around optional data and connector calls.",
    tags: "canvas app formula error power fx",
  },
  {
    code: "DelegationWarning",
    status: "",
    product: "Power Apps",
    layer: "Canvas app formula",
    severity: "Low",
    source: "SRC-001",
    meaning: "A Power Apps formula cannot be fully delegated to the data source and may return incomplete results.",
    causes: "Unsupported function, nondelegable operator, complex expression, or data source delegation limit.",
    resolution: "Rewrite the formula using delegable functions or move filtering logic to the data source.",
    prevention: "Check delegation warnings during app build and test with large datasets.",
    tags: "delegation warning canvas app power fx",
  },
  {
    code: "PatchFunctionFailed",
    status: "",
    product: "Power Apps",
    layer: "Canvas app data write",
    severity: "Medium",
    source: "SRC-001",
    meaning: "A Power Apps Patch formula failed to create or update a record.",
    causes: "Required field missing, wrong data type, permission issue, invalid lookup, or connector validation failure.",
    resolution: "Inspect Errors() output, validate required fields and data types, and confirm write permissions.",
    prevention: "Show user-friendly validation before Patch is called.",
    tags: "patch function failed power apps",
  },
  {
    code: "SubmitFormFailed",
    status: "",
    product: "Power Apps",
    layer: "Canvas app form",
    severity: "Medium",
    source: "SRC-001",
    meaning: "A form submission in a canvas app failed.",
    causes: "Invalid form mode, required field missing, data card update formula issue, permission problem, or data source validation error.",
    resolution: "Check Form.Error and Errors() output, validate data cards, and confirm the user can write to the data source.",
    prevention: "Use form validation and clear error messages before submit.",
    tags: "submit form failed canvas app",
  },
  {
    code: "ModelDrivenCommandFailed",
    status: "",
    product: "Power Apps",
    layer: "Model-driven app command",
    severity: "Medium",
    source: "SRC-003",
    meaning: "A model-driven app command failed when the user selected it.",
    causes: "JavaScript error, command visibility rule issue, missing privilege, invalid selected record, or flow/action failure.",
    resolution: "Check browser console, command bar rules, user privileges, selected row context, and called action logs.",
    prevention: "Test commands with users across security roles and app modules.",
    tags: "model driven command failed",
  },
  {
    code: "BusinessRuleFailed",
    status: "",
    product: "Power Apps",
    layer: "Model-driven app",
    severity: "Medium",
    source: "SRC-003",
    meaning: "A business rule blocks form save or changes field behavior unexpectedly.",
    causes: "Condition mismatch, required field rule, hidden locked field, conflicting JavaScript, or solution layer difference.",
    resolution: "Review active business rules, rule scope, form context, and recent solution changes.",
    prevention: "Document business rules and test forms after managed solution imports.",
    tags: "business rule failed model driven app",
  },
  {
    code: "FormScriptError",
    status: "",
    product: "Power Apps",
    layer: "Model-driven app",
    severity: "Medium",
    source: "SRC-003",
    meaning: "A model-driven form JavaScript library throws an error.",
    causes: "Invalid control reference, deprecated client API, missing web resource, wrong execution context, or unexpected null value.",
    resolution: "Check browser console, validate form libraries, update deprecated APIs, and add null-safe checks.",
    prevention: "Regression-test form scripts after form or library changes.",
    tags: "form script error model driven javascript",
  },
  {
    code: "CustomConnectorSchemaMismatch",
    status: "400",
    product: "Power Platform",
    layer: "Custom connector",
    severity: "Medium",
    source: "SRC-001",
    meaning: "A custom connector response or request does not match the schema expected by Power Platform.",
    causes: "OpenAPI definition mismatch, changed API response, missing required property, or wrong data type.",
    resolution: "Update the connector definition, retest the operation, and refresh dependent flows or apps.",
    prevention: "Version external APIs and connector definitions together.",
    tags: "custom connector schema mismatch",
  },
  {
    code: "CustomConnectorPolicyFailed",
    status: "",
    product: "Power Platform",
    layer: "Custom connector",
    severity: "Medium",
    source: "SRC-001",
    meaning: "A custom connector policy or transformation failed during execution.",
    causes: "Invalid policy expression, missing header, unexpected body shape, or operation definition mismatch.",
    resolution: "Review connector policies, test the raw API call, and simplify transformations until the operation succeeds.",
    prevention: "Keep connector policies small and covered by operation tests.",
    tags: "custom connector policy failed",
  },
  {
    code: "ConnectorNotAllowedByPolicy",
    status: "403",
    product: "Power Platform",
    layer: "DLP policy",
    severity: "High",
    source: "SRC-001",
    meaning: "A connector cannot run because an environment data policy does not allow it.",
    causes: "Connector blocked, connector group conflict, custom connector not approved, or new policy applied to the environment.",
    resolution: "Review data policies in the admin center and request an approved policy change or connector alternative.",
    prevention: "Check connector policy classification before solution design.",
    tags: "connector not allowed by policy DLP",
  },
  {
    code: "TenantRestrictionViolation",
    status: "403",
    product: "Power Platform",
    layer: "Tenant policy",
    severity: "High",
    source: "SRC-001",
    meaning: "A tenant-level policy prevents the requested app, flow, connector, or data operation.",
    causes: "Admin restriction, cross-tenant access block, connector disabled tenant-wide, or conditional access policy.",
    resolution: "Review tenant settings, environment policy, and identity logs with a Power Platform or identity admin.",
    prevention: "Validate tenant policies before choosing connectors or external identities.",
    tags: "tenant restriction violation policy",
  },
  {
    code: "EnvironmentNotFound",
    status: "404",
    product: "Power Platform",
    layer: "Environment",
    severity: "High",
    source: "SRC-001",
    meaning: "The referenced Power Platform environment cannot be found or accessed.",
    causes: "Wrong environment ID, deleted environment, moved environment, tenant mismatch, or missing admin/user access.",
    resolution: "Confirm the environment exists in the correct tenant and update URLs, IDs, and deployment configuration.",
    prevention: "Use environment variables and deployment settings rather than hardcoded environment IDs.",
    tags: "environment not found power platform",
  },
  {
    code: "EnvironmentCapacityExceeded",
    status: "",
    product: "Power Platform",
    layer: "Environment",
    severity: "High",
    source: "SRC-001",
    meaning: "The environment does not have enough capacity for the requested operation or storage growth.",
    causes: "Database, file, log, process, or request capacity reached by apps, flows, or Dataverse usage.",
    resolution: "Review capacity in the admin center, clean unnecessary data, purchase capacity, or move workloads.",
    prevention: "Monitor environment capacity and set alerts for growth trends.",
    tags: "environment capacity exceeded storage",
  },
];

function slugify(value) {
  return String(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 120);
}

function csvEscape(value) {
  const text = String(value ?? "");
  return /[",\n]/.test(text) ? `"${text.replaceAll("\"", "\"\"")}"` : text;
}

function escapeHtml(value) {
  return String(value || "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    "\"": "&quot;",
    "'": "&#039;",
  })[char]);
}

function makeMetaDescription(record) {
  return `${record.code} in ${record.contextTitle}: what it means, common causes, and how to resolve it in ${record.product}.`.slice(0, 155);
}

function makeArticle(template, context, index) {
  const codeSlug = slugify(template.code.replace(/\[[^\]]+\]/g, "").trim() || template.code);
  const slug = `${codeSlug}-${context.key}`;
  const id = `PP-${String(index).padStart(4, "0")}`;
  const contextMeaning = `${template.meaning} This article focuses on the ${context.title.toLowerCase()} scenario.`;
  const contextCauses = `${template.causes} In this scenario, ${context.causeHint.charAt(0).toLowerCase()}${context.causeHint.slice(1)}`;
  const contextResolution = `${template.resolution} For ${context.title.toLowerCase()}, also verify: ${context.causeHint}`;
  const seoTitle = `${template.code} in ${context.title} | Power Platform Error Fix`;

  return {
    id,
    product: template.product,
    layer: template.layer,
    code: template.code,
    status: template.status,
    contextKey: context.key,
    contextTitle: context.title,
    appliesTo: context.appliesTo,
    scenario: context.scenario,
    meaning: contextMeaning,
    causes: contextCauses,
    resolution: contextResolution,
    prevention: `${template.prevention} For ${context.title.toLowerCase()}, document the expected permissions, connector settings, and test data used to validate the fix.`,
    severity: template.severity,
    owner: context.owner || template.owner || "Power Platform owner",
    source: template.source,
    lastVerified: verifiedDate,
    tags: `${template.tags} ${context.key} ${context.title}`,
    seoTitle,
    seoSlug: slug,
    metaDescription: makeMetaDescription({ ...template, contextTitle: context.title }),
    canonicalPath: `/errors/${slug}.html`,
    confidence: "Medium",
  };
}

function buildRecords() {
  const rows = [];
  const seenSlugs = new Set();
  const seenKeys = new Set();
  const duplicateDrops = [];
  let candidateIndex = 1;

  for (const template of templates) {
    for (const context of contexts) {
      const record = makeArticle(template, context, candidateIndex);
      const slugKey = record.seoSlug;
      const contentKey = [
        record.product,
        record.layer,
        record.code,
        record.status,
        record.contextKey,
      ].map((part) => String(part || "").toLowerCase().trim()).join("|");

      if (seenSlugs.has(slugKey) || seenKeys.has(contentKey)) {
        duplicateDrops.push({ slug: slugKey, key: contentKey });
        continue;
      }

      seenSlugs.add(slugKey);
      seenKeys.add(contentKey);
      rows.push({ ...record, id: `PP-${String(rows.length + 1).padStart(4, "0")}` });
      candidateIndex++;

      if (rows.length === targetCount) return { rows, duplicateDrops };
    }
  }

  return { rows, duplicateDrops };
}

function sourceLinks(sourceIds) {
  return sourceIds
    .split(";")
    .map((id) => id.trim())
    .filter(Boolean)
    .map((id) => ({ id, ...sources[id] }))
    .map((source) => `<li><a href="${escapeHtml(source.url)}" target="_blank" rel="noreferrer">${escapeHtml(source.id)}: ${escapeHtml(source.title)}</a></li>`)
    .join("\n");
}

function articleHtml(item) {
  const canonical = `${siteBaseUrl}${item.canonicalPath}`;
  const jsonLd = {
    "@context": "https://schema.org",
    "@type": "TechArticle",
    headline: item.seoTitle,
    description: item.metaDescription,
    dateModified: item.lastVerified,
    about: ["Microsoft Power Platform", item.product, item.layer, item.code, item.contextTitle],
    mainEntityOfPage: canonical,
    author: {
      "@type": "Organization",
      name: "Power Platform Error Resolution Database",
    },
  };

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeHtml(item.seoTitle)}</title>
  <meta name="description" content="${escapeHtml(item.metaDescription)}">
  <link rel="canonical" href="${escapeHtml(canonical)}">
  <meta property="og:title" content="${escapeHtml(item.seoTitle)}">
  <meta property="og:description" content="${escapeHtml(item.metaDescription)}">
  <meta property="og:type" content="article">
  <meta property="og:url" content="${escapeHtml(canonical)}">
  <script type="application/ld+json">${JSON.stringify(jsonLd)}</script>
  <style>
    :root { --ink:#17212b; --muted:#5b6773; --line:#d8e0e6; --surface:#fff; --brand:#15506a; --danger:#b42318; --warn:#a84618; --shadow:0 16px 42px rgba(20,38,50,.12); }
    * { box-sizing: border-box; }
    body { margin:0; font-family:Arial, Helvetica, sans-serif; color:var(--ink); background:#eef4f7; }
    header { background:var(--brand); color:#fff; padding:26px clamp(18px,4vw,56px) 82px; }
    .wrap, main, footer { max-width:1060px; margin:0 auto; }
    .crumb { color:#dcecf2; text-decoration:none; font-weight:700; }
    h1 { margin:24px 0 12px; font-size:clamp(2rem,5vw,3.4rem); line-height:1.06; letter-spacing:0; }
    .intro { max-width:850px; margin:0; color:#dcecf2; line-height:1.55; }
    main { margin-top:-52px; padding:0 clamp(16px,4vw,56px) 42px; }
    .panel, .card { background:var(--surface); border:1px solid var(--line); border-radius:8px; padding:20px; }
    .panel { box-shadow:var(--shadow); }
    .grid { display:grid; grid-template-columns:1fr 1fr; gap:16px; margin-top:18px; }
    .meta { display:flex; flex-wrap:wrap; gap:8px; margin-bottom:16px; }
    .pill { display:inline-flex; align-items:center; min-height:28px; border-radius:999px; padding:4px 10px; background:#edf4f6; color:#234654; font-size:.82rem; font-weight:700; }
    .pill.high { background:#fde8e4; color:var(--danger); }
    .pill.medium { background:#fff1df; color:var(--warn); }
    .pill.low { background:#e7f5ec; color:#14633c; }
    h2 { margin:0 0 10px; font-size:1.08rem; }
    p, li { color:var(--muted); line-height:1.55; }
    p { margin:0; }
    ul { margin:0; padding-left:20px; }
    a { color:var(--brand); font-weight:700; }
    .back { display:inline-flex; margin-top:18px; }
    footer { padding:0 clamp(16px,4vw,56px) 34px; color:var(--muted); font-size:.88rem; }
    @media (max-width:780px) { .grid { grid-template-columns:1fr; } }
  </style>
</head>
<body>
  <header>
    <div class="wrap">
      <a class="crumb" href="../index.html">Power Platform Error Database</a>
      <h1>${escapeHtml(item.code)} in ${escapeHtml(item.contextTitle)}</h1>
      <p class="intro">${escapeHtml(item.meaning)}</p>
    </div>
  </header>
  <main>
    <article class="panel">
      <div class="meta">
        <span class="pill ${escapeHtml(item.severity.toLowerCase())}">${escapeHtml(item.severity)}</span>
        <span class="pill">${escapeHtml(item.id)}</span>
        <span class="pill">${escapeHtml(item.product)}</span>
        <span class="pill">${escapeHtml(item.layer)}</span>
        ${item.status ? `<span class="pill">HTTP ${escapeHtml(item.status)}</span>` : ""}
      </div>
      <div class="grid">
        <section class="card"><h2>Applies To</h2><p>${escapeHtml(item.appliesTo)}</p></section>
        <section class="card"><h2>Typical Scenario</h2><p>${escapeHtml(item.scenario)}</p></section>
        <section class="card"><h2>What This Error Means</h2><p>${escapeHtml(item.meaning)}</p></section>
        <section class="card"><h2>Common Causes</h2><p>${escapeHtml(item.causes)}</p></section>
        <section class="card"><h2>How To Resolve It</h2><p>${escapeHtml(item.resolution)}</p></section>
        <section class="card"><h2>How To Prevent It</h2><p>${escapeHtml(item.prevention)}</p></section>
        <section class="card"><h2>Ownership And Search Tags</h2><p><strong>Owner:</strong> ${escapeHtml(item.owner)}</p><p><strong>Tags:</strong> ${escapeHtml(item.tags)}</p></section>
        <section class="card"><h2>Source References</h2><ul>${sourceLinks(item.source)}</ul></section>
      </div>
      <a class="back" href="../index.html">Back to searchable database</a>
    </article>
  </main>
  <footer>Power Platform Error Codes and Resolution Database. Last verified ${escapeHtml(item.lastVerified)}.</footer>
</body>
</html>`;
}

function indexErrorObject(item) {
  return {
    id: item.id,
    product: item.product,
    layer: item.layer,
    code: item.code,
    status: item.status,
    meaning: item.meaning,
    causes: item.causes,
    resolution: item.resolution,
    prevention: item.prevention,
    severity: item.severity,
    owner: item.owner,
    source: item.source,
    tags: item.tags,
    seoSlug: item.seoSlug,
  };
}

function updateIndexHtml(indexHtml, rows) {
  const sourceLiteral = `const sources = ${JSON.stringify(sources, null, 6)};`;
  const errorLiteral = `const errors = ${JSON.stringify(rows.map(indexErrorObject), null, 6)};`;
  const links = rows
    .map((item) => `                <a href="errors/${item.seoSlug}.html">${escapeHtml(item.id)}: ${escapeHtml(item.code)} in ${escapeHtml(item.contextTitle)}</a>`)
    .join("\n");

  let next = indexHtml
    .replace(/const sources = \{[\s\S]*?\n    \};/, sourceLiteral)
    .replace(/const errors = \[[\s\S]*?\n\n    const searchInput/, `${errorLiteral}\n\n    const searchInput`)
    .replace(/href="errors\/\$\{item\.id\.toLowerCase\(\)\}\.html"/, 'href="errors/${item.seoSlug}.html"')
    .replace(/Current production build: [\d,]+ SEO article pages\./, `Current production build: ${rows.length.toLocaleString("en-US")} SEO article pages.`)
    .replace(
      /                <!-- STATIC_ERROR_LINKS_START -->[\s\S]*?                <!-- STATIC_ERROR_LINKS_END -->/,
      `                <!-- STATIC_ERROR_LINKS_START -->\n${links}\n                <!-- STATIC_ERROR_LINKS_END -->`,
    );

  next = next.replace(
    /<footer>[\s\S]*?<\/footer>/,
    `<footer>\n      Source-backed starter database last verified ${verifiedDate}. Production build contains ${rows.length} SEO article pages with duplicate slug and content-key checks.\n    </footer>`,
  );

  return next;
}

async function writeSpreadsheet(rows, duplicateDrops) {
  const workbook = Workbook.create();
  const articles = workbook.worksheets.add("Articles");
  const sourceSheet = workbook.worksheets.add("Sources");
  const report = workbook.worksheets.add("Build Report");
  for (const sheet of [articles, sourceSheet, report]) sheet.showGridLines = false;

  const headers = [
    "ID",
    "Product Area",
    "Layer",
    "Error Code / Message",
    "HTTP Status",
    "Context",
    "Applies To",
    "Scenario",
    "Symptoms / Meaning",
    "Common Causes",
    "Resolution Steps",
    "Prevention",
    "Severity",
    "Owner",
    "Source ID",
    "Last Verified",
    "Search Tags",
    "SEO Title",
    "SEO Slug",
    "Meta Description",
    "Canonical Path",
    "Confidence",
  ];

  const articleRows = rows.map((row) => [
    row.id,
    row.product,
    row.layer,
    row.code,
    row.status,
    row.contextTitle,
    row.appliesTo,
    row.scenario,
    row.meaning,
    row.causes,
    row.resolution,
    row.prevention,
    row.severity,
    row.owner,
    row.source,
    row.lastVerified,
    row.tags,
    row.seoTitle,
    row.seoSlug,
    row.metaDescription,
    row.canonicalPath,
    row.confidence,
  ]);

  articles.getRange("A1:V1").values = [headers];
  articles.getRangeByIndexes(1, 0, articleRows.length, headers.length).values = articleRows;
  articles.tables.add(`A1:V${articleRows.length + 1}`, true, "ArticlePipeline").style = "TableStyleMedium2";
  articles.freezePanes.freezeRows(1);
  articles.getRange("A1:V1").format = { fill: "#1F4E5F", font: { bold: true, color: "#FFFFFF" }, wrapText: true };
  articles.getRange(`A2:V${articleRows.length + 1}`).format = { wrapText: true, verticalAlignment: "top" };
  const widths = [80, 150, 150, 230, 90, 180, 260, 330, 340, 360, 380, 340, 90, 160, 120, 110, 260, 280, 260, 340, 260, 95];
  widths.forEach((width, col) => articles.getRangeByIndexes(0, col, articleRows.length + 1, 1).format.columnWidthPx = width);

  sourceSheet.getRange("A1:C1").values = [["Source ID", "Title", "URL"]];
  const sourceRows = Object.entries(sources).map(([id, source]) => [id, source.title, source.url]);
  sourceSheet.getRangeByIndexes(1, 0, sourceRows.length, 3).values = sourceRows;
  sourceSheet.tables.add(`A1:C${sourceRows.length + 1}`, true, "Sources").style = "TableStyleMedium4";
  sourceSheet.getRange("A:C").format.columnWidthPx = 260;

  report.getRange("A1:B8").values = [
    ["Metric", "Value"],
    ["Target articles", targetCount],
    ["Generated articles", rows.length],
    ["Duplicate candidates dropped", duplicateDrops.length],
    ["Unique SEO slugs", new Set(rows.map((row) => row.seoSlug)).size],
    ["Unique content keys", new Set(rows.map((row) => `${row.product}|${row.layer}|${row.code}|${row.status}|${row.contextKey}`)).size],
    ["Sources", Object.keys(sources).length],
    ["Last verified", verifiedDate],
  ];
  report.getRange("A1:B1").format = { fill: "#1F4E5F", font: { bold: true, color: "#FFFFFF" } };
  report.getRange("A:B").format.columnWidthPx = 240;

  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(path.join(outputDir, "power_platform_error_content_pipeline.xlsx"));
}

const { rows, duplicateDrops } = buildRecords();
if (rows.length < targetCount) {
  throw new Error(`Only generated ${rows.length} unique records; target is ${targetCount}.`);
}

await fs.mkdir(dataDir, { recursive: true });
await fs.mkdir(outputDir, { recursive: true });
await fs.rm(errorDir, { recursive: true, force: true });
await fs.mkdir(errorDir, { recursive: true });

const csvHeaders = [
  "id",
  "product",
  "layer",
  "code",
  "status",
  "contextKey",
  "contextTitle",
  "appliesTo",
  "scenario",
  "meaning",
  "causes",
  "resolution",
  "prevention",
  "severity",
  "owner",
  "source",
  "lastVerified",
  "tags",
  "seoTitle",
  "seoSlug",
  "metaDescription",
  "canonicalPath",
  "confidence",
];

const csv = [
  csvHeaders.join(","),
  ...rows.map((row) => csvHeaders.map((header) => csvEscape(row[header])).join(",")),
].join("\n");

await fs.writeFile(path.join(dataDir, "power_platform_error_articles.csv"), csv, "utf8");
await fs.writeFile(path.join(dataDir, "power_platform_error_articles.json"), JSON.stringify(rows, null, 2), "utf8");
await fs.writeFile(path.join(dataDir, "sources.json"), JSON.stringify(sources, null, 2), "utf8");
await fs.writeFile(path.join(dataDir, "build_report.json"), JSON.stringify({
  targetCount,
  generatedCount: rows.length,
  duplicateCandidatesDropped: duplicateDrops.length,
  uniqueSlugs: new Set(rows.map((row) => row.seoSlug)).size,
  uniqueContentKeys: new Set(rows.map((row) => `${row.product}|${row.layer}|${row.code}|${row.status}|${row.contextKey}`)).size,
  lastVerified: verifiedDate,
}, null, 2), "utf8");

await writeSpreadsheet(rows, duplicateDrops);

for (const row of rows) {
  await fs.writeFile(path.join(errorDir, `${row.seoSlug}.html`), articleHtml(row), "utf8");
}

const indexHtml = await fs.readFile(indexPath, "utf8");
await fs.writeFile(indexPath, updateIndexHtml(indexHtml, rows), "utf8");

const sitemap = `<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
  <url>
    <loc>${siteBaseUrl}/index.html</loc>
    <lastmod>${verifiedDate}</lastmod>
  </url>
${rows.map((row) => `  <url>
    <loc>${siteBaseUrl}${row.canonicalPath}</loc>
    <lastmod>${verifiedDate}</lastmod>
  </url>`).join("\n")}
</urlset>
`;

const robots = `User-agent: *
Allow: /

Sitemap: ${siteBaseUrl}/sitemap.xml
`;

await fs.writeFile(path.join(root, "sitemap.xml"), sitemap, "utf8");
await fs.writeFile(path.join(root, "robots.txt"), robots, "utf8");

console.log(JSON.stringify({
  generated: rows.length,
  duplicateCandidatesDropped: duplicateDrops.length,
  pagesDir: errorDir,
  spreadsheet: path.join(outputDir, "power_platform_error_content_pipeline.xlsx"),
}, null, 2));
