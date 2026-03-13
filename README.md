# Oral Examiner 4.0

**A voice-based oral defense system for student essays.** Students submit an essay, have a live voice conversation with an AI examiner, and receive a grade adjustment based on how well they defended their work.

Built on Google Sheets + Google Apps Script. The student-facing frontend is hosted on **GitHub Pages** (not served from Apps Script) to enable reliable microphone access for the voice examiner. Free to run (you provide your own API keys).

---

## What You Need Before Starting

You'll need two free API keys (takes ~5 minutes to get both):

1. **ElevenLabs account** (for the voice examiner) — sign up at [elevenlabs.io](https://elevenlabs.io)
   - You'll need an **Agent ID** (from creating a Conversational AI agent) and an **API Key** (from your profile)
   - ElevenLabs has a free tier, but oral exams use voice minutes — a paid plan is recommended for a full class

2. **Google Gemini API key** (for grading) — get one free at [aistudio.google.com](https://aistudio.google.com)

---

## Setup Guide (5 minutes)

### Step 1: Copy the Spreadsheet

1. Click this link: **[Make a copy of Oral Examiner 4.0](https://docs.google.com/spreadsheets/d/1G6Brx32ctj1-VUTcCnVwTcbSkY_9jlZJiLakuN71OMQ/copy)**
2. Click **Make a copy** when prompted

You now have your own spreadsheet with 5 tabs: Database, Config, Prompts, Questions, and Logs.

### Step 2: Add the Code

1. In your spreadsheet, go to **Extensions > Apps Script**
2. This opens the script editor in a new tab
3. Delete any code already in the editor
4. Open the file **`code.gs`** from this repository, copy its entire contents, and paste it into the script editor
5. Click the **+** next to "Files" in the left sidebar, choose **HTML**, and name it **`index`** (not `index.html` — Apps Script adds the extension automatically)
6. Open the file **`index.html`** from this repository, copy its entire contents, and paste it into the `index.html` file you just created
7. In the script editor, click the gear icon (**Project Settings**) on the left sidebar
   - Check **"Show 'appsscript.json' manifest file in editor"**
   - Go back to the Editor, open `appsscript.json`, and replace its contents with the `appsscript.json` from this repository
8. Click **Save** (or Ctrl+S)

### Step 3: Run the Setup Wizard

1. Go back to your **spreadsheet** tab (not the script editor)
2. Refresh the page — wait a few seconds for the menu to appear.  You may need to close the spreadsheet and open it again.
3. Click **Oral Defense > Setup Wizard (start here)**
4. Google will ask you to authorize the script — click through the permissions prompts
   - You may see a "This app isn't verified" warning. Click **Advanced > Go to [project name]** to continue. This is your own script running on your own account — it's safe.
5. The Setup Wizard dialog will appear. Enter:
   - **ElevenLabs Agent ID** — from your ElevenLabs Conversational AI agent
   - **ElevenLabs API Key** — from elevenlabs.io > Profile + API Keys (the icon in the lower left)
   - **Gemini API Key** — from aistudio.google.com
   - **App Title** (optional) — whatever you want students to see in the header
6. Click **Save & Complete Setup**

### Step 4: Deploy the Apps Script Backend

1. Go back to the **Apps Script editor** tab
2. Click **Deploy > New deployment** (blue button, top right)
3. Click the gear icon next to "Select type" and choose **Web app**
4. Set:
   - **Description:** anything (e.g., "v1")
   - **Execute as:** Me
   - **Who has access:** Anyone
5. Click **Deploy**
6. Copy the **Web app URL** that appears — you'll need this in the next step

This URL is the **backend API only**. Students will not access this URL directly.

### Step 5: Set Up the Frontend (GitHub Pages)

The student-facing portal (`index.html`) is hosted on GitHub Pages, not served from Apps Script. This is required because Apps Script serves pages inside an iframe, which blocks microphone access needed for the voice examiner.

1. Open `index.html` in this repository
2. Find the line near the top of the `<script>` block:
   ```javascript
   const APPS_SCRIPT_URL = "https://script.google.com/macros/s/DEPLOYMENT_ID/exec";
   ```
3. Replace `DEPLOYMENT_ID` with your actual Apps Script deployment URL from Step 4
4. Commit and push the change
5. Enable GitHub Pages in your repository settings (Settings > Pages > Source: main branch)
6. Your portal URL will be something like: `https://yourusername.github.io/your-repo-name/`

**That GitHub Pages URL is your exam portal.** Share it with students, put it on your Canvas assignments, whatever.

---

## Customizing for Your Course

The template comes with sample prompts and questions. You'll want to customize these for your subject:

### Questions Tab

This is your question bank. Each row has two columns:
- **category**: either `content` or `process`
- **question**: the question text

**Content questions** test whether students know their essay and source material. **Process questions** ask about their writing process.

The system randomly selects questions for each student (default: 2 content + 1 process). Change the counts in the Config tab (`content_questions_count`, `process_questions_count`).

### Prompts Tab

This controls the examiner's personality and behavior:
- **agent_personality** — How the examiner talks and acts
- **agent_examination_flow** — The structure of the exam (greeting, questions, wrap-up)
- **first_message** — What the examiner says first (use `{student_name}` as a placeholder)
- **grading_system_prompt** — Instructions for the AI grader
- **grading_rubric** — The rubric and scoring formula

### Config Tab

Non-secret settings you can tweak:
- `app_title` — Portal header text
- `app_subtitle` — Smaller text under the title (leave empty to hide)
- `avatar_url` — URL of the bot's profile image (use any public image URL)
- `min_call_length` — Calls shorter than this (seconds) are automatically excluded
- `max_paper_length` — Maximum essay length in characters
- `gemini_model` — Which Gemini model to use for grading

---

## How Students Use It

1. Student opens your GitHub Pages URL
2. Clicks **Enter the Portal**
3. Types their name and pastes their essay, clicks **Submit**
4. Clicks **Begin Oral Defense** — the voice examiner starts talking
5. Has a ~15-minute conversation defending their essay
6. Clicks **Finish Defense** when done
7. The transcript is saved automatically

---

## How You Grade

1. Open your spreadsheet
2. Click **Oral Defense > Grade All Pending**
3. Gemini reads each essay + transcript and produces a grade adjustment (percentage points) and detailed comments
4. Review the grades in the **Database** tab — the AI Adjustment and AI Comment columns are filled in
5. Add your own notes in the Instructor Notes column if needed

---

## Troubleshooting

**"Setup Wizard" doesn't appear in the menu**
- Refresh the spreadsheet page and wait 5-10 seconds. The menu loads when the sheet opens.

**Students can't connect to the examiner**
- Check that your ElevenLabs Agent ID is correct
- Check that your ElevenLabs account has available voice minutes

**Transcripts aren't appearing**
- Click **Oral Defense > Recover Stuck Defenses** — this manually fetches any missing transcripts
- A background process also checks every 5 minutes automatically

**"Excluded" status on a submission**
- The call was too short (under 60 seconds by default). This usually means a mic failure or accidental disconnect.
- To re-include it: change the Status cell to `Defense Complete`, then run Grade All Pending.

**Grading produces unexpected results**
- Review and customize the `grading_rubric` prompt in the Prompts tab
- The default rubric is designed for literary analysis essays — adjust the rubric elements and scoring for your subject

---

## How It Works (Technical Summary)

- The student portal (`index.html`) is hosted on **GitHub Pages** as a top-level page (not in an iframe), which enables reliable microphone access for the ElevenLabs voice SDK
- The frontend communicates with the **Apps Script backend** via `fetch()` requests using `?action=` routing
- The spreadsheet is the database — everything lives in Google Sheets
- When a student finishes their defense, the frontend calls the backend, which queries the ElevenLabs API to retrieve the conversation transcript
- A background trigger checks every 5 minutes for any missed transcripts
- Grading sends the essay + transcript to Google's Gemini AI with your rubric
- No webhook setup required — everything is fetched via API

---

## Updating the Code

**Frontend changes** (index.html): Push to GitHub and GitHub Pages will update automatically. Students use the same URL.

**Backend changes** (code.gs):
1. Go to **Extensions > Apps Script**
2. Click **Deploy > Manage deployments**
3. Click the pencil icon on your deployment
4. Change the version to **New version**
5. Click **Deploy**

The Apps Script URL stays the same, so the frontend continues to work without changes.
