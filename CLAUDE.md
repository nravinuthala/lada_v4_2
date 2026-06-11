# CLAUDE.md — LADA v4.2 (Learning Asset Development Agent)
# HCLTech Career Shaper™ — Multi-Agent Agentic AI Platform
# Project Code: CS-LF | Features: AI_AGN_MCP
# Last Updated: 2026-04-27 | Version: 4.2-fix

---

## 1. PROJECT OVERVIEW

**What this is:**
LADA (Learning Asset Development Agent) is a Streamlit-based multi-agent application that orchestrates
5 AI agents sequentially to generate complete corporate training asset packages from a single user input.

**Business context:**
Built for HCLTech Career Shaper™ to automate training content authoring at scale. Designed for
L&D managers, instructional designers, and training leads who need to produce full learning packages
quickly without manual slide authoring.

**Source file (current):**
- Input: `agentcode.txt` (the working codebase — rename to `.py` before running)
- Fixed output: `lada_v4_1_fixed.py` (model name corrected — use this to run)
- Final target filename: `lada_v4_2.py`

---

## 2. ARCHITECTURE — 5-AGENT SEQUENTIAL PIPELINE

```
User Input (Landing Page Form)
        │
        ▼
┌──────────────────┐
│  Agent 1 (📋)    │  Guide-Sheet Generator
│  Excel .xlsx     │  8-section structured study guide
└────────┬─────────┘
         │ output → session_state.guidesheet_txt
         ▼
┌──────────────────┐
│  Agent 2 (📊)    │  PowerPoint Creator
│  PPTX 35 slides  │  60 min deck + slide-by-slide voice-over scripts
└────────┬─────────┘
         │ output → session_state.pptx_script + session_state.pptx_files
         ▼
┌──────────────────┐
│  Agent 3 (🖼️)    │  Image Spec Generator
│  JSON prompts    │  Per-slide AI image generation prompts
└────────┬─────────┘
         │ output → session_state.image_specs
         ▼
┌──────────────────┐
│  Agent 4 (🔊)    │  Audio Voice-Over Agent
│  SSML scripts    │  TTS-ready narration scripts per slide
└────────┬─────────┘
         │ output → session_state.audio_script
         ▼
┌──────────────────┐
│  Agent 5 (✅)    │  Deck Validator
│  QA Scorecard    │  7-dimension quality report, score/100, verdict
└──────────────────┘
```

**Key agentic properties:**
- Each agent has its own system prompt, role, and task
- Each agent reads all prior agent outputs from session_state (chained context)
- Human review gate between every agent (reviewer name + comments saved to state)
- Each agent can be re-run independently without resetting the whole pipeline
- Per-agent token tracking + cumulative total in sidebar dashboard

---

## 3. PAGES & NAVIGATION

| Page | Route Key | Function | Purpose |
|---|---|---|---|
| Landing Page | `"landing"` | `render_landing()` | Input form — topic, audience, modules |
| Admin Module | `"admin"` | `render_admin()` | API key config, provider switch, token log |
| Agent Workspace | `"workspace"` | `render_workspace()` | 5 tabbed agents, run sequentially |

---

## 4. DUAL AI PROVIDER SUPPORT

Supports **Anthropic Claude** and **OpenAI GPT-4o** — switchable at runtime from Admin.

```python
CLAUDE_MODEL = "claude-sonnet-4-6"     # ← CORRECT model ID (fixed from v4.1)
OPENAI_MODEL = "gpt-4o"
```

**Critical: The original file had `"claude-sonnet-4-20250514"` — this does NOT exist.**
**Always use `"claude-sonnet-4-6"` for Claude Sonnet in the current API.**

All 5 agents call `call_claude()` which is an alias for `call_ai()`.
`call_ai()` reads `st.session_state.ai_provider` to route to the correct SDK.

Admin password (for demo): `EdTech@123`

---

## 5. SESSION STATE SCHEMA

All mutable state lives in `st.session_state`. The full schema:

```
admin_unlocked  bool        Admin panel gate
ai_provider     str         "anthropic" | "openai"
api_key         str         Active provider key (in-memory only, never to disk)
openai_key      str         OpenAI key store
anthropic_key   str         Anthropic key store
api_key_valid   bool|None   None=untested, True=ok, False=invalid
token_log       list[dict]  Per-call usage records
current_page    str         "landing" | "admin" | "workspace"
current_agent   int         0–4, active agent index
agent_status    list[str]   "idle"|"running"|"done"|"error" per agent
agent_progress  list[int]   0–100 per agent
agent_tokens    list[int]   Cumulative tokens per agent
total_tokens    int         Grand total across all agents
form_data       dict        Landing page inputs
guidesheet_txt  str         Agent 1 raw text output
guidesheet_xl   bytes|None  Excel workbook bytes
pptx_files      list[dict]  [{"name": str, "data": bytes}]
pptx_script     str         Slide spec text from Agent 2
image_specs     str         Image prompt JSON from Agent 3
audio_script    str         Voice-over scripts from Agent 4
audio_doc       bytes|None  TTS script package bytes
validation_rpt  str         QA scorecard from Agent 5
review_comments dict        {agent_idx: str}
reviewer_names  dict        {agent_idx: str}
job_name        str         Asset name shorthand
color_palette   str         Active colour palette string
uploaded_doc    bytes|None  Optional uploaded curriculum PDF
```

---

## 6. KEY FUNCTIONS REFERENCE

| Function | Section | Purpose |
|---|---|---|
| `init_session_state()` | 2 | Initialise all state keys idempotently |
| `inject_global_css()` | 3 | HCLTech brand CSS injection |
| `sanitise_filename(raw)` | 4 | Make user input safe for filenames |
| `validate_anthropic_key(key)` | 4 | Live 10-token API ping to validate |
| `validate_openai_key(key)` | 4 | Live 10-token API ping to validate |
| `call_ai(prompt, system, agent_idx, job_name)` | 5 | Unified LLM gateway — routes by provider |
| `call_claude` | 5 | Alias for `call_ai()` used by all agents |
| `build_context_string()` | 5 | Assembles form_data into prompt context block |
| `render_sidebar()` | 6 | Nav + agent tracker + token dashboard |
| `render_review_gate(agent_idx)` | 6 | Human review panel after each agent |
| `render_next_agent_button(idx)` | 6 | Proceed button between agents |
| `render_rerun_button(agent_idx)` | 6 | Re-run single agent without full reset |
| `render_admin()` | 7 | Admin page — key config, token log |
| `render_landing()` | 8 | Input form page |
| `run_agent1()` | 9 | Guide-Sheet agent UI + prompt + Excel gen |
| `generate_guidesheet_excel(content)` | 9 | openpyxl workbook builder |
| `run_agent2()` | 10 | PowerPoint agent UI + prompt |
| `generate_pptx(slide_script)` | 10 | python-pptx deck builder (35 slides) |
| `run_agent3()` | 11 | Image spec agent UI + prompt |
| `_build_image_spec_doc(specs_json, prompts)` | 11 | Image spec document builder |
| `run_agent4()` | 12 | Audio script agent UI + prompt |
| `_build_audio_script_doc(...)` | 12 | TTS script package builder |
| `run_agent5()` | 13 | Validator agent UI + prompt + scorecard |
| `render_workspace()` | 14 | 5-tab agent workspace page |
| `main()` | 15 | Entry point + page router |

---

## 7. BRAND COLOURS (HCLTech)

```python
BRAND = {
    "purple":   "#5F1EBE",   # Primary — CTAs, gradients
    "blue":     "#3C91FF",   # Secondary — links
    "teal":     "#00A4A6",   # Accents, success
    "midnight": "#00112B",   # Header banner ONLY (dark)
    "white":    "#FFFFFF",   # Card backgrounds
    "page_bg":  "#F5F7FF",   # Page canvas
    "text_dark":"#0C1B3A",   # Primary text (WCAG AAA)
    "success":  "#0A8F5E",
    "danger":   "#D41C1C",
    "warning":  "#B86200",
}
```

Excel uses openpyxl hex constants (no `#`): `XL_DARK_BLUE = "001F5B"` etc.
PPTX uses `RGBColor(r, g, b)` tuples: `C_PURPLE = RGBColor(0x5F, 0x1E, 0xBE)` etc.

---

## 8. DEPENDENCIES

```bash
pip install streamlit anthropic openai openpyxl python-pptx Pillow requests pandas
```

All imports are guarded with try/except — missing packages show a friendly `st.error()` then `st.stop()`.

Python: 3.8+
Framework: Streamlit ≥ 1.30

---

## 9. HOW TO RUN

### Step 1 — Prepare the file
The source is `agentcode.txt`. Rename or copy it:
```bash
cp agentcode.txt lada_v4_2.py
```

### Step 2 — Fix the model name (if using original agentcode.txt)
On line 201, change:
```python
# WRONG
CLAUDE_MODEL: str = "claude-sonnet-4-20250514"

# CORRECT
CLAUDE_MODEL: str = "claude-sonnet-4-6"
```
(Already fixed in `lada_v4_1_fixed.py`)

### Step 3 — Install dependencies
```bash
pip install streamlit anthropic openai openpyxl python-pptx Pillow requests pandas
```

### Step 4 — Run
```bash
streamlit run lada_v4_2.py
```

### Step 5 — Demo flow
1. Go to **Admin Module** → Password: `EdTech@123`
2. Select provider: **Anthropic Claude**
3. Enter your API key → click **Validate & Save**
4. Go to **Landing Page** → fill in: Asset Name, Audience, Duration, Topics
5. Click **START** → goes to Agent Workspace
6. Run agents one by one: Agent 1 → review → Agent 2 → review → ... → Agent 5

---

## 10. WHAT CLAUDE CODE MUST DO

When picking up this project, Claude Code should:

1. **Read `agentcode.txt`** (the full source — 3,635 lines, 15 sections)
2. **Fix model name** on line 201: `"claude-sonnet-4-20250514"` → `"claude-sonnet-4-6"`
3. **Save as `lada_v4_2.py`** in the project directory
4. **Install all dependencies** via pip
5. **Run** with `streamlit run lada_v4_2.py`
6. **Verify** the app loads on `localhost:8501` with no import errors

### If errors occur during run:

| Error | Likely cause | Fix |
|---|---|---|
| `AuthenticationError` | Wrong/missing API key | Enter key in Admin Module |
| `model not found` | Model string wrong | Check `CLAUDE_MODEL` constant on line 201 |
| `ModuleNotFoundError` | Missing pip package | `pip install <package>` |
| `AttributeError: module 'openai'` | Old openai SDK | Already handled with explicit imports — check openai SDK is v1.x |
| Streamlit widget key conflicts | State mismatch on rerun | `st.rerun()` clears — expected behaviour |

### Do NOT:
- Change agent prompt text without being asked
- Move session_state keys — the schema is load-bearing across all 15 sections
- Add new pages without updating `main()` router
- Store API keys to disk — keys are in-memory by design (security requirement)
- Use `st.experimental_rerun()` — use `st.rerun()` (Streamlit ≥ 1.30)

---

## 11. COMPLETION TASKS (if file is partially complete)

The file `agentcode.txt` is **structurally complete** — all 15 sections and all functions are present.
There are **no TODO stubs or NotImplemented blocks**.

The only task required before running is:
- [ ] Fix `CLAUDE_MODEL` string (line 201)
- [ ] Rename to `.py`
- [ ] Install dependencies
- [ ] Run

If Claude Code is asked to **extend** the project, priority additions are:
1. **PDF curriculum upload support** — parse `uploaded_doc` bytes in `build_context_string()`
2. **Auto-chain agents** — run all 5 agents sequentially with a single "Run All" button
3. **Export ZIP** — bundle all 5 outputs (Excel, PPTX, image specs, audio, report) into one download
4. **Persistent token log** — write token log to `token_log.csv` on each call

---

## 12. FILE NAMING CONVENTION

All project documents follow:
```
{DOCTYPE}_{PROJECT}_{FEATURES}_{DATE}_{VERSION}.{ext}
```

Example:
```
SPEC_CS-LF_AI_AGN_MCP_20260427_v4.2.md
CODE_CS-LF_AI_AGN_MCP_20260427_v4.2.py
```

Project shortcode: `CS-LF`
Feature tags: `AI_AGN_MCP`

---

## 13. SECURITY NOTES

- API keys stored ONLY in `st.session_state` — never written to disk, never logged
- Admin password validated server-side: `ADMIN_PASSWORD = "EdTech@123"`
- For production: replace hardcoded password with `os.environ.get("LADA_ADMIN_PASSWORD", "")`
- All user inputs sanitised via `sanitise_filename()` before embedding in filenames
- API key displayed masked only: first 12 + last 6 chars

---

*CLAUDE.md generated for LADA v4.2 | Career Shaper™ | HCLTech Edutech*
*Use this file as the canonical development contract for Claude Code sessions.*
