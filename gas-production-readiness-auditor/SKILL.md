---
name: gas-production-readiness-auditor
description: A skill for comprehensively auditing Google Apps Script (GAS) projects on performance, maintainability, extensibility, reliability, data integrity, operability, developer experience, and security (GAS Production Readiness Auditor). Rather than simple performance optimization, it evaluates production readiness including architectural design, code quality, and failure-recovery design. Use this whenever asked for a GAS code review, help with a bloated Code.gs, trigger design, Spreadsheet optimization, idempotency / duplicate-processing prevention / failure-recovery design, or a maintainability diagnosis.
---

# GAS Production Readiness Auditor

You comprehensively audit a GAS project on eight axes: Performance, Architecture (extensibility), Maintainability, Reliability, Data Integrity, Operations / Observability, Security, Developer Experience.

This is not just speed advice. Two questions drive every judgment:

- Can this still be changed safely a year from now? (maintainability)
- Can this be recovered when something breaks? (operational readiness)

A recoverable failure path matters more than a fast success path.

> This framework is a collection of production experience, operational lessons, and architectural heuristics for GAS. It is not an official Google standard. Use engineering judgment according to project scale, risk, and requirements.

---

# Final Principles

```text
Rule 1: The fastest API is the one that is never called
Rule 2: Not executing a process 100 times beats executing it 100 times fast
Rule 3: Good GAS isn't fast GAS — it's GAS that can still be changed safely a year from now
Rule 4: Design for the failure case, not the success case
Rule 5: LockService prevents concurrent execution. Idempotency prevents re-execution. You need both
Rule 6: Prioritize "processing that can recover when broken" over "processing that succeeds"
Rule 7: Complexity isn't free. Choose the minimal design that fits the scale and risk
```

Rules 1–3 are the performance / maintainability criteria, Rules 4–6 the operational-readiness criteria, Rule 7 the over-engineering check. When about to recommend "turn everything into a Repository / convert everything to TypeScript / put everything through a State Machine," return to Rule 7.

---

# Output Format

The audit result must always be structured as follows.

```text
## Overall Score
## Production Readiness Verdict
## Scores by Category
## Issues Found (by Severity)
## Suggestions
## Prioritized Action List
## Refactoring Proposals (Before/After)
```

## Scoring Criteria

```text
A+  95-100
A   90-94
B   80-89
C   70-79
D   60-69
F    0-59
```

Each category is scored on the same scale.

## Severity Determination

Issues are classified by "is this safe to ship to production," not "how many points to deduct." One Critical issue means "cannot ship" regardless of score, so Severity is an axis independent of the score. F is about the score; Critical is about ship / no-ship — different axes.

```text
Critical  Cannot ship to production. Leads to data destruction, information leakage, or a guaranteed incident
High      Must fix early. Cannot recover during an outage / integrity breaks
Medium    Fix as planned. The blast radius is limited, but leaving it becomes technical debt
Low       Fix if there's room. The actual harm is small
Info      Noted only. May be an acceptable design decision
```

Examples of classification:

```text
Critical : Hardcoded API key / No LockService / Web App set to "Execute as: Me" + "Anyone can access" / Re-send processing with no idempotency
High     : No checkpoint / No retry / Partial failure not handled / Status is only a binary success|fail
Medium   : Direct reference to row[14] / Dependence on getLastRow / Config values reloaded every time
Low      : Overuse of Logger.log / Sloppy naming
```

## Production Readiness Verdict

Separately from the overall score, explicitly state whether this is fit for production.

```text
READY                  Fit for production
READY WITH CONDITIONS   Fit for production if the stated conditions are met
NOT READY              Not fit for production (blocking issues must be resolved)
```

One or more Critical → `NOT READY`. Only High → `READY WITH CONDITIONS`. Neither → `READY`. Example output:

```text
Overall Score:         92 / A
Production Readiness:  NOT READY
Blocking Issues:
  - LockService Missing (Critical)
  - API Token Hardcoded (Critical)
```

---

# Audit Categories

## Performance

### 1. Cell-by-cell API access
Count `getValue`/`setValue`/`getValues`/`setValues`/`flush` calls. Principle: avoid cell-by-cell access; fetch only the range you need, in bulk. Also check `getDataRange` usage and total API call count.

Bad:
```javascript
for (const row of rows) {
  range.getValue();
}
```
Recommended:
```javascript
const data = sheet
  .getRange(startRow, startColumn, rowCount, columnCount)
  .getValues();
```

`getDataRange().getValues()` is convenient, but when data volume is unpredictable it runs into #6 (chunking / memory). Specify an explicit range whenever possible.

### 2. Read All → Process → Write All
Verify the read-all → process → write-all order is followed. Implementations that alternate reads and writes are prohibited.

### 3. Event Object usage
Check whether `onEdit`/`onChange`/`onOpen`/`onFormSubmit` use the Event Object. Warn if `e.value`/`e.oldValue`/`e.range` are ignored and the value is re-fetched via `sheet.getRange(...).getValue()`.

### 4. Cache
Check whether `CacheService` is used, and whether API responses, config values, or reference data are caching candidates.

At the same time, check whether anything that must NOT be cached has been put in the cache. CacheService can vanish **without warning** due to capacity limits, a max 6-hour TTL, and eviction.

```text
OK to cache    : Config values / GET results / master data / expensive computed results
NOT OK to cache: Lock/exclusion state / processed flags / transactions / monetary data
```

Principle: "If losing it would be a problem, don't cache it." If persistence is required, use PropertiesService or a dedicated sheet.

### 5. Quotas & execution limits
Three independent axes — evaluate each; do not conflate them.

**(a) Per-service daily quota** — Gmail/Drive/UrlFetch/Calendar daily usage and the risk of exhaustion.

**(b) Single-execution time** — the ceiling is **6 minutes per execution** for both Consumer and Google Workspace accounts (not 30 minutes). Custom Functions get 30 seconds; Google Workspace Add-ons have their own limit. Identify the execution context first, then evaluate as a percentage of the limit:

```text
Usage <50%   Safe
Usage 50-75% Warning     (risk zone as data grows)
Usage 75-90% High        (split/chunk early)
Usage >90%   Critical    (on the verge of timing out)
```
When the limit is unknown, use 6 minutes as the conservative baseline. Avoid designs that consume nearly the whole budget in one execution; check whether chunking + checkpoints (#6) can split the work.

**(c) Account-wide / project limits**
```text
Utilities.sleep cap            5 min
Concurrent executions          30 per user
Total trigger runtime per day  Consumer: 90 min / Google Workspace: 6 hours
Per-project limits             caps on trigger count, deployment count, etc.
```
"A single execution finishes within 6 minutes, so it's safe" does not follow — a high-frequency time-driven trigger exhausts the daily total trigger runtime first. Estimate by target account type.

### 6. Chunking & checkpoints
Check for chunked processing with checkpoints (resumability via `PropertiesService`). Watch for `getDataRange().getValues()` bulk-loading tens of thousands of rows at once; evaluate switching to chunked retrieval via `getRange(startRow, 1, chunkSize, cols)`. Initial rule of thumb 100–1000 records, tuned by column count, API call count, per-record processing time, and quota consumption — a fixed 1000 is not the answer.

## Architecture

### 7. Responsibility separation
Warn if `main()` is doing loading, validation, processing, and saving all by itself. Recommend splitting by function: `load()` / `validate()` / `process()` / `save()`.

### 8. Layering & dependency direction
Recommend `UI → Service → Repository → Spreadsheet` layering, with dependencies flowing in a single direction.

```text
Ideal: UI → Controller → Service → Repository → Google API
Bad  : Repository holds business logic / Service references UI (reverse flow or circular)
```

Flag a lower layer referencing a higher one, or business logic living in the Repository layer.

## Maintainability

### 9. Code organization & structural debt
Warn about a single file becoming bloated (e.g. a 3000-line Code.gs); recommend splitting into files such as `Config.gs`/`Main.gs`/`Sheets.gs`/`Drive.gs`/`Utils.gs`. Score function length against thresholds, and flag copy-paste blocks (the same ~10+ lines in 3+ places). This is the "can anyone still safely touch this in six months" axis — distinct from #43 (AI-generated smells), which is about deletion candidates.

```text
Function length  >100 lines  Warning  /  >300 lines  High  /  >500 lines  Critical
```

### 10. Naming
Check the validity of function, variable, and file names. Prohibit meaningless names like `test()`; recommend intention-revealing names like `generateInvoicePdf()`.

### 11. Constants
Detect hardcoded sheet names or IDs; recommend defining constants such as `const SHEET_MAIN = "..."`.

### 12. Column access by name
Prohibit direct index references like `row[14]` — they break the moment a column is inserted or moved. Read the header row into a `{columnName: index}` map and access by name (`COL_EMAIL`, `employee.email`), and replace index-dependent data structures with object mappings such as `{ email, status }`.

## Reliability

### 13. Error handling
Check for the presence of `try/catch`. Treat `catch(e){}` (swallowing errors) as a serious warning. In `catch(e)`, log the failed operation, target ID, input values, and `e.stack` — not a bare `Exception: Service error`.

### 14. Recovery & checkpoints
Check whether processing state is persisted (`PropertiesService`, etc.) so a long-running process can resume after an interruption.

## Observability

### 15. Structured logging
`Logger.log()` itself is not prohibited (both `Logger` and `console` are usable in the execution log). The problems are heavy sequential logging, string-concatenated logs, and a production multi-user environment whose only log destination is the execution log. In production, route logs to `console` / Cloud Logging / Error Reporting, and prefer structured logs that can be searched and aggregated.

```javascript
// Bad
console.log('import failed ' + messageId);

// Recommended
console.log({ level: 'ERROR', event: 'IMPORT_FAILED', messageId, parser, error: e.stack });
```

Check whether `INFO` / `WARN` / `ERROR` are distinguished, and whether failure logs include the target ID and input values.

### 16. Debuggability
Check whether debugging relies solely on logs, and whether the setup allows use of an actual debugger.

## Spreadsheet Formulas

### 17. Formulas vs GAS
Check formula-side usage such as `ARRAYFORMULA`/`OFFSET`/`VLOOKUP`/`IMPORTRANGE`, and whether processing implemented in GAS could instead be replaced by formulas (Named Functions, ArrayFormula).

## UI

### 18. UI surface
Warn against overuse of `Ui.prompt()`; recommend a Sidebar/Dialog instead. Detect reliance on SSR via `<?!= include() ?>`.

## Front-End Architecture

### 19. Scale-appropriate front-end
Judge by whether the cost/benefit of adoption is proportionate to the project's scale. Not having an SPA, a router, TypeScript, or a test framework is not, by itself, a reason to dock points — for a small GAS project, plain JS + `clasp` + light testing is entirely sufficient (Rule 7). Check:

- SPA structure / routing / screen-switching in the Web App/Sidebar/Dialog, where scale warrants it
- `google.script.run` usage — recommend abstraction via a Promise wrapper
- TypeScript adoption
- Test setup (Jest/Vitest)

## Developer Experience

### 20. Toolchain & local dev
Check for a local development environment via `clasp`, and adoption of ESLint/Prettier/Vite/TypeScript, proportionate to scale.

### 21. Version control
Check for version control via Git. Treat operating only in the production environment as a serious warning.

### 22. Deployment hygiene
Check for Dev/Stage/Prod separation, whether the latest code is actually reflected in the deployment, whether `/exec` (production) and `/dev` (latest code, requires permission) are being confused, and whether deployment versions/descriptions are managed.

## Security

### 23. Secrets
Detect and prohibit hardcoding API keys/tokens/passwords into the code. Recommend moving them out of the code into a store such as `PropertiesService` — but that alone isn't the whole story. Script Properties is a shared store accessible from the same script, and anyone with edit permission on the script editor can read it. Check the execution identity (the user the script runs as), the editors, and the access scope together. For a script tied to a broadly shared Spreadsheet, this is effectively close to plaintext, so consider isolating the project or restricting access. For highly sensitive credentials, also consider a service such as Google Cloud Secret Manager.

### 24. OAuth scope minimization
Check whether OAuth scopes and execution permissions are excessive. Cross-check `appsscript.json` `oauthScopes` against the permissions the code actually uses (see #33). If `oauthScopes` has grown on its own alongside code changes (scope drift) and the increase can't be explained → Medium–High.

### 25. Web App execution identity
Check the "Execute as" setting for Web App deployments. "Me" + "Anyone can access" is an information-leak risk, since visitors can operate the Spreadsheet/Drive/Gmail with the deployer's privileges. Treat as a serious warning unless intentional.

### 26. Granular OAuth consent
Users can partially decline scopes via granular consent — don't assume every required scope was granted.

- Do installable triggers validate permissions at creation time via `ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, [...])`, avoiding registration if permissions are insufficient?
- For a user-run Web App / Add-on, can a scope shortfall be detected before it crashes mid-execution?
- Do `onOpen`/`onEdit` account for `NONE` / `LIMITED` via `e.authMode` (#29)?

An implementation that swallows a scope shortfall and processes partway through anyway is High.

## GAS-Specific Pitfalls

Check for GAS-specific pitfalls where common JavaScript/Node assumptions don't hold.

### 27. Concurrency / LockService
Identify processes that can run concurrently, such as `onEdit`/`onFormSubmit`/Web App. For processing that **mutates shared state**, where concurrent execution could cause a race, duplicate execution, or row misalignment, the absence of `LockService` (`getScriptLock`/`getDocumentLock`) or an equivalent exclusion design is Critical (it leads directly to duplicate registrations, duplicate email sends, or data corruption). Do not require a lock for read-only processing that doesn't conflict under concurrency, or for processing whose write target is independent per execution.

### 28. appendRow in a loop
Check whether `appendRow()` is being called in a loop hundreds or thousands of times. Recommend buffering data and writing it in bulk via a single `setValues()` call at the end.

### 29. Trigger context & authorization
`Session.getActiveUser().getEmail()` can return an empty string during trigger execution or Web App execution — for processing that depends on user identification, use `Session.getEffectiveUser()` or form-respondent information. Warn if code that only works under manual execution has been registered as a trigger.

Check `e.authMode` (`ScriptApp.AuthMode`): Simple Triggers such as `onOpen`/`onEdit` run under `LIMITED` or `NONE`, where services requiring authorization (`UrlFetchApp`/`GmailApp`/`DriveApp`/other file access, etc.) cannot be called. Recommend migrating to an Installable Trigger if necessary. Identify places where behavior differs between manual execution, time-driven triggers, and Web App execution.

### 30. getLastRow / getLastColumn
Check whether the code assumes `getLastRow()` can pick up traces of old data, leftover formulas, or format-only cells and return unexpected values. Recommend explicitly managing the data range as "number of header rows + actual record count."

### 31. Empty cell is "" not null
An empty cell returns `""` (an empty string), not `null`. Check for branches or `typeof` checks that assume `value === null` or `!value`, and fix them to use `=== ""` or an explicit emptiness check.

### 32. Timezone
Check the assumption that `new Date()` (dependent on the script's timezone), `Session.getScriptTimeZone()`, and the Spreadsheet's own timezone all match. Recommend explicitly specifying the timezone when formatting dates via `Utilities.formatDate(date, tz, format)`. Also check the `timeZone` setting in `appsscript.json`.

### 33. Manifest audit
Check `appsscript.json` independently.

```text
oauthScopes      Do they match the permissions the code actually uses? Any excessive/unused scopes? (see #24)
timeZone         Does it match the date handling in the code and the Spreadsheet? (see #32)
runtimeVersion   Is it V8? (see #34)
exceptionLogging Is it set to STACKDRIVER? (see #15)
dependencies     Any unused Advanced Services / libraries left in?
webapp / addOn   Are the access / executeAs settings as intended? (see #25)
```

### 34. V8 / legacy syntax
Check whether the runtime in `appsscript.json` is V8. Check for leftover `var`-heavy code, function-expression callbacks, and Rhino-era workarounds, and recommend updating to `const`/`let`/arrow functions/template literals/destructuring.

### 35. HtmlService viewport & mobile readiness
The `<meta name="viewport">` tag inside HTML can be ignored by HtmlService. Attach it via `.addMetaTag('viewport', 'width=device-width, initial-scale=1')` in `doGet`, on every return value of both `createHtmlOutputFromFile` and `createTemplateFromFile().evaluate()`. If missing, the page renders as a desktop layout on mobile and media queries won't trigger. Also check other real-device failure points: `input` font sizes under 16px (causing iOS auto-zoom), tap targets under 44px, and `position: fixed` interfering with the on-screen keyboard.

### 36. HtmlService CSP / CDN
Check whether the Web App/Sidebar is heavily dependent on external CDN `<script src>` tags or CSS. Evaluate the risk from iframe sandbox constraints or CDN deprecation causing breakage, and recommend bundling critical libraries locally or providing a fallback.

### 37. UrlFetchApp resilience & retry safety
Assume `UrlFetchApp.fetch()` can normally fail with a 429/500/timeout. Require `muteHttpExceptions: true` + status code checking + exponential backoff retry (`Utilities.sleep`). Flag unprotected single-shot calls as a warning.

```text
429 / 502 / 503 / 504 → Retry (with exponential backoff)
408 / timeout          → Retry
400 / 401 / 403 / 404  → Do not retry (an input/authorization problem; resending won't fix it)
Retry-After header     → If present, honor it and wait that many seconds
```

Flag an implementation that "retries everything 3 times" — it will pointlessly keep hammering an expired authorization or a validation error. Insert backoff waits via `Utilities.sleep()` (5-minute cap), and design the retry count and interval so waiting and retrying together don't eat up the 6-minute execution budget (#5).

**Is it safe to retry?** Resilience is the mechanism to retry; retry-safety is whether retrying is harmless. For a POST, check that the receiving side has an idempotency key or an existence check — otherwise a failed `fetch()` → retry → duplicate registration on the other side. Links to #44.

### 38. PropertiesService capacity
Check whether large JSON blobs are being stored in `PropertiesService` (the limit is roughly 9KB per value / ~500KB total — it is not a database). For growing structured data, consider `CacheService`, a separate sheet, or an external store such as Firestore.

## Data & Operational Dependencies

### 39. Multi-sheet dependencies
Check whether sheet-to-sheet dependencies such as `SheetA → SheetB → QUERY/IMPORTRANGE → GAS` are implicit. Recommend documenting the dependency graph as a diagram or list, and identifying where changes to columns in a source sheet will ripple outward.

### 40. Drive search by name
Warn against heavy use of `DriveApp.getFilesByName()`/`getFoldersByName()`. Because these can mis-hit on files with the same name, large file counts, or shared drives, recommend referencing files/folders directly by ID (`getFileById`) instead.

### 41. Trigger explosion
Check the number and duplication of installed triggers. Detect duplicate time-driven triggers on the same function, orphaned triggers (whose target function no longer exists), and accumulated unnecessary edit triggers, and recommend auditing them via `ScriptApp.getProjectTriggers()`.

### 42. Google service boundaries
Check whether a single function touches Spreadsheet/Gmail/Drive/Calendar all at once. Recommend separating each service into its own wrapper function so that quota, permissions, and failure units can be isolated.

### 43. AI-generated code smells
Detect smells characteristic of AI-generated code: `catch(e){}`, unused `Utils` functions, dead code, comments that only pad line count, and abstractions with only a single implementation. Review these with an eye toward deletion/consolidation.

## Reliability & Operational Readiness

The most common incidents in GAS operations are not "it's slow" but "duplicate processing" and "it broke and nobody noticed." The focus here is whether the failure path and the re-execution path — not just the success path — have been designed for. These checks are mandatory when external integrations, triggers, or a Web App are present.

### 44. Idempotency & duplicate events
Check whether the result stays the same if the same Message ID / Transaction ID / Response ID is resent. Gmail trigger re-execution, webhook resends, and duplicate FormSubmit events happen routinely, and the event source itself can deliver the same event more than once. Check whether processed IDs (`messageId` / `threadId` / `transactionId`) are persisted in PropertiesService or a dedicated sheet and checked against before processing. A write or send operation with no idempotency key is Critical. (Trigger-level duplication is #41.)

### 45. Data integrity across destinations
For processing that spans multiple destinations (Spreadsheet + Notion, Sheet A + Sheet B, etc.), check whether Partial Failure handling exists. A design where "Spreadsheet save succeeds → Notion save fails" leaves inconsistency unaddressed is High. Check for rollback, compensating actions, or an "unsynced" flag with later catch-up sync.

### 46. State machine
For long-running processes, external integrations, and resumable batches, check whether record state is made explicit as `NEW` / `PROCESSING` / `SUCCESS` / `FAILED` / `RETRY`. If state is only a binary success/failure, a record that crashed mid-processing cannot be distinguished from one that was never started, which breaks the premise of resumable processing (#14). Don't demand a five-state machine for a small one-off task (Rule 7).

### 47. Failed-record queue
Check whether there is a mechanism to set aside failed records. Check whether a fallback sheet such as `FAILED_JOBS` / `FAILED_EMAILS` / `FAILED_IMPORTS` records the failed input, target ID, error, and timestamp of occurrence, so the operations team can reprocess it later. If failures simply flow into `console.log` and disappear, that's High. If a fallback exists, it becomes the reprocessing target for #37.

### 48. Trigger ownership
An installable trigger **executes with the permissions of the account of the user who created it.** If the creator is a personal account, the trigger stops when that person leaves, transfers, or their account is deleted — and processing silently drops without anyone noticing.

- Is it known who owns each trigger?
- Is the procedure for changing/recreating the owner documented?
- Is a production recurring process tied to a specific individual?

In enterprise operations, a recurring trigger with an unknown owner is High.

### 49. Shared ownership / bus factor
Check whether the production GAS project exists only in an individual's My Drive. Check whether ownership sits under a shared drive or organizational management, and whether more than one person can edit/deploy the script (i.e. whether the bus factor is greater than 1). Individual ownership with no handoff procedure is High.

---

# How to Conduct the Audit

1. Understand the target project's file structure, line counts, and trigger definitions
2. Scan Performance first (#1–#6) — the largest source of bottlenecks
3. Check Architecture & Maintainability (#7–#12)
4. Check Reliability & Observability (#13–#16)
5. Check the remaining categories (Formulas, UI, Front-End, DevEx, Security, GAS-Specific Pitfalls, Data/Ops Dependencies) as appropriate to the actual project — explicitly mark inapplicable categories as excluded from scoring
6. Check Reliability & Operational Readiness (#44–#49), plus #27 (Lock) and #37 (retry safety). Missing idempotency, failure fallback, state management, scope validation, trigger ownership, ownership concentration in one person, or manifest drift directly lowers Production Readiness
7. Following the output format, compile the score, issues by Severity, Production Readiness verdict, and improvement suggestions. Keep recommendations to the minimum measures proportionate to the project's scale and risk (Rule 7)

## Checks Especially Prone to Being Missed (Check These First)

```text
1. The 6-minute timeout (#5)
2. Repeated getRange/getValue/setValue calls (#1)
3. Trigger permission differences / Simple vs Installable (#29)
4. Forgetting LockService (#27)
5. Re-send processing with no idempotency (#44)
6. No fallback for failed records (#47)
7. Date timezone issues (#32)
8. HtmlService viewport (#35)
9. Mass appendRow execution (#28)
10. The getLastRow pitfall (#30)
11. Web App execution-permission settings (#25)
```
