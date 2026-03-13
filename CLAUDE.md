# Oral Examiner 4.0

An oral defense examination system for student essays, built as a Google Apps Script web app. Students submit essays, then defend them in a voice conversation with an AI examiner (an ElevenLabs voice agent), and are graded by Gemini AI.

## Architecture

- **GitHub Pages** hosts `index.html` as a top-level page (no iframe) — enables microphone access for ElevenLabs voice widget
- **Google Apps Script** backend — API endpoints via `doGet`/`doPost` with `?action=` routing, plus ElevenLabs webhook receiver (optional backup)
- **Google Sheets** as the database, config store, and prompt store (spreadsheet ID stored in Script Properties, not hardcoded)
- **ElevenLabs Conversational AI** widget for voice-based oral defense
- **Gemini API** for automated grading of defense transcripts

## Files

- `code.gs` — All backend logic: web endpoints (`doGet`/`doPost` with `?action=` routing), API handlers for frontend (`getConfig`, `submitEssay`, `fetchTranscript`), ElevenLabs webhook receiver (optional backup), submission processing, question selection, prompt building, Gemini grading, setup wizard
- `index.html` — Single-page frontend with 4 screens (welcome → submit → defense → complete). Hosted on GitHub Pages. Inline CSS and JS. Calls Apps Script backend via `fetch()` using `callBackend()` helper
- `appsscript.json` — Apps Script manifest (V8 runtime, Sheets + external request scopes)
- `Prompts` — Local reference file (tab-separated: prompt_name → prompt_text). Canonical copy lives in the Google Sheet "Prompts" tab; this file is the local mirror

## Data Flow

1. Student opens GitHub Pages URL → `loadConfig()` calls `GET ?action=getConfig` → frontend receives agent ID, max chars, UI config
2. Student submits essay → `POST ?action=submitEssay` → `processSubmission()` generates UUID session_id, selects random questions, stores in Sheets, returns prompt + questions
3. Frontend configures 11Labs widget with `override-prompt` (essay + questions baked in) and `dynamic-variables` (session_id)
4. Student has voice defense with the examiner
5. Call ends → frontend calls `POST ?action=fetchTranscript` with sessionId and conversationId, with retry logic; also, 11Labs sends transcript webhook → `doPost()` → `handleTranscriptWebhook()` (whichever arrives first wins)
6. **Automatic recovery**: A time-driven trigger runs `autoRecoverTranscripts()` every 5 minutes to catch any missed transcripts
7. `gradeDefense()` sends essay + transcript to Gemini with rubric from Prompts sheet → stores multiplier and structured comments

## Google Sheets Structure

Tabs: **Database** (submissions), **Config** (key-value pairs), **Prompts** (prompt name → text), **Questions** (category + question), **Logs** (debug log)

Key config values: `elevenlabs_agent_id`, `elevenlabs_api_key`, `gemini_api_key`, `gemini_model`, `webhook_secret`, `max_paper_length`, `app_title`, `avatar_url`

## Development Notes

- This is NOT a Node.js project — it's Google Apps Script (V8 runtime). No package manager, no build step
- Code is pushed to Apps Script via `clasp` or copy-paste to the script editor
- All backend functions are global (Apps Script requirement) — no module system
- Frontend is hosted on **GitHub Pages** (top-level page, not in Apps Script iframe) to enable microphone access for the ElevenLabs voice widget
- Frontend calls the Apps Script backend via `fetch()` using `callBackend(action, data)`. POST requests use `Content-Type: text/plain` to avoid CORS preflight (Apps Script doesn't handle OPTIONS). The `APPS_SCRIPT_URL` constant in `index.html` must contain the deployed Apps Script `/exec` URL
- The ElevenLabs SDK is loaded from `cdn.jsdelivr.net/npm/@elevenlabs/client@0.15.0` with `livekit-client@2.11.4`
- **Spreadsheet ID**: `getSpreadsheetId()` reads from Script Properties (set by Setup Wizard). NOT a hardcoded const — this allows each teacher's copy to auto-configure
- **Setup Wizard**: `showSetupWizard()` → HTML dialog → `runSetupWizard(config)` stores keys in PropertiesService, auto-captures spreadsheet ID, installs time-driven trigger
- **`isSetupComplete()`**: Checks that `spreadsheet_id`, `elevenlabs_agent_id`, `elevenlabs_api_key`, and `gemini_api_key` are all set in Script Properties
- Secrets (API keys, agent ID) are stored in `PropertiesService.getScriptProperties()`, not in the spreadsheet. `getConfig()` checks PropertiesService first for ALL keys, then Config sheet, then DEFAULTS
- Prompts are fetched from the Prompts sheet via `getPrompt()` with hardcoded fallbacks in `buildDefensePrompt()` and `getFirstMessage()`
- **Transcript fetch flow**: `fetchAndStoreTranscript(sessionId, conversationId)` is called from the frontend (via `POST ?action=fetchTranscript`) after a defense ends. It queries the ElevenLabs API, stores the transcript, and returns `{success, retryable}` for the frontend to handle retries
- **Auto-recovery**: `autoRecoverTranscripts()` runs via a time-driven trigger (every 5 min) and silently recovers stuck transcripts. Installed by the Setup Wizard
- **Recovery (manual)**: `recoverStuckDefenses()` (Oral Defense menu → "Recover Stuck Defenses") does the same thing but shows a UI dialog with results

## Grading System

- **Rubric**: 4 elements — Paper Knowledge (1-3) and Writing Process (1-3) are capped at 3; Text Knowledge (1-5) and Content Understanding (1-5) can go higher. 3 = meets expectations
- **Adjustment formula**: `(average - 3) × 5`, clamped to [-10, +5], rounded to 1 decimal place
- **Flags for instructor**: Any element scoring below 3, or flag text in the FLAGS FOR INSTRUCTOR section, triggers a flag. Comments are prefixed with "⚠ FLAG FOR INSTRUCTOR ⚠"
- **Parser** (`parseGradingResponse`): extracts `Adjustment: +/-X.X` line from Gemini's structured output; falls back to computing from individual scores if that line is missing
- Prompts: `grading_system_prompt` (role/persona) and `grading_rubric` (rubric + scoring formula + output format)

## Style Conventions

- Use JSDoc comments on functions
- Constants use UPPER_SNAKE_CASE; column indices are 1-based (matching Sheets)
- Status values are defined in the `STATUS` object: Submitted → Defense Started → Defense Complete → Graded → Reviewed (also: Excluded)
- **Exclusion**: Calls shorter than `min_call_length` config (default 60s) are auto-set to "Excluded" status and skipped by grading. To manually exclude, change the status cell to "Excluded" in the spreadsheet; to re-include, change it back to "Defense Complete"
- Log to the Logs sheet via `sheetLog(source, message, data)` for debugging — visible in the spreadsheet
