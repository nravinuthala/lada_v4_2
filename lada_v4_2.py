"""
================================================================================
  lada_v4_1.py
  ------------
  Learning Asset Development Agent (LADA)  v4.2
  HCLTech Career Shaper™ — Multi-Agent Agentic AI Platform

  v4.1 Changes: Dual-AI provider support — Anthropic Claude AND OpenAI GPT-4o.
  Admin Module lets you choose provider, enter key, validate, and test.
  call_ai() routes to the active provider transparently across all 5 agents.
  Bright professional theme, dark banner, high-contrast content areas.

  Description:
    A Streamlit-based multi-agent application that orchestrates 5 AI agents
    sequentially to generate complete corporate training assets:
      Agent 1 → Excel Guide-Sheet (8 sections, fully formatted)
      Agent 2 → PowerPoint Deck (35 slides/60 min, voice-over scripts)
      Agent 3 → Slide-wise Image Specifications (per-slide AI prompts)
      Agent 4 → Audio Voice-Over Scripts (SSML-enhanced, TTS-ready)
      Agent 5 → Validation Scorecard (7-dimension QA report)

  Author      : Career Shaper EdTech Platform
  Version     : 4.2
  Python      : 3.8+
  Framework   : Streamlit ≥ 1.30

  Dependencies:
    pip install streamlit anthropic openai openpyxl python-pptx Pillow requests pandas

  Run:
    streamlit run learning_asset_agent.py

  Security Notes:
    - API keys are stored only in st.session_state (in-memory, not on disk)
    - Admin password is validated server-side before any privileged action
    - All user inputs are sanitised before being embedded in file names
    - No credentials are ever logged, printed, or written to files
    - API key is masked in the UI (first 12 + last 6 characters only)

  Architecture:
    ┌────────────────────────────────────────────────────────────┐
    │  CONSTANTS & CONFIG                                        │
    │  SESSION STATE INITIALISATION                              │
    │  GLOBAL CSS (HCLTech Brand)                                │
    │  SECURITY HELPERS (validate_api_key, sanitise_filename)    │
    │  API LAYER (call_claude — centralised, token-logged)       │
    │  UI COMPONENTS (sidebar, badges, review panels)            │
    │  PAGE RENDERERS (landing, admin, workspace)                │
    │  AGENT RUNNERS (run_agent1 … run_agent5)                   │
    │  FILE GENERATORS (Excel, PPTX, audio doc)                  │
    │  MAIN ROUTER                                               │
    └────────────────────────────────────────────────────────────┘
================================================================================
"""

# ── Standard library imports ──────────────────────────────────────────────────
import datetime          # Timestamps for token log and file metadata
import io                # In-memory byte streams for file downloads
import logging           # Application-level logging (non-sensitive)
import os                # OS path utilities
import re                # Regular expressions for content parsing
import time              # Sleep for UX pacing during progress updates
from typing import Dict, List, Optional, Tuple  # Type hints throughout

# ── Third-party imports ───────────────────────────────────────────────────────
import streamlit as st   # Web application framework

try:
    import anthropic     # Official Anthropic Claude SDK
except ImportError:
    st.error("Missing dependency: run  pip install anthropic")
    st.stop()

try:
    import openai                                    # Official OpenAI SDK v1.x
    # ── FIX: In openai v1.x, exception classes are NOT accessible as
    # openai.AuthenticationError on the module object — they must be imported
    # explicitly. This fixes "module 'openai' has no attribute 'AuthenticationError'"
    from openai import (
        AuthenticationError as _OAIAuthErr,
        RateLimitError       as _OAIRateErr,
        APIStatusError       as _OAIStatusErr,
        APIConnectionError   as _OAIConnErr,
    )
    _OPENAI_OK = True
except ImportError:
    # openai not installed — define harmless dummy classes so except clauses
    # never raise NameError even if openai is absent
    openai       = None
    _OPENAI_OK   = False
    class _OAIAuthErr(Exception):   pass   # noqa
    class _OAIRateErr(Exception):   pass   # noqa
    class _OAIStatusErr(Exception): pass   # noqa
    class _OAIConnErr(Exception):   pass   # noqa

try:
    import openpyxl                           # Excel workbook generation (Agent 1)
    from openpyxl.styles import (
        PatternFill, Font, Alignment,
        Border, Side,
    )
except ImportError:
    st.error("Missing dependency: run  pip install openpyxl")
    st.stop()

try:
    from pptx import Presentation             # PowerPoint generation (Agent 2)
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
except ImportError:
    st.error("Missing dependency: run  pip install python-pptx")
    st.stop()

# ── Logging configuration ─────────────────────────────────────────────────────
# INFO level: captures agent lifecycle events without exposing secrets.
# WARNING+ is also written; DEBUG is suppressed in production.
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(funcName)s — %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

# =============================================================================
# SECTION 1 — CONSTANTS & CONFIGURATION
# =============================================================================

# ── Streamlit page configuration ──────────────────────────────────────────────
# Must be the very first Streamlit call in the script.
st.set_page_config(
    page_title="LADA v4.2 — HCLTech Career Shaper™",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── HCLTech / Career Shaper brand colours ─────────────────────────────────────
# Source: HCLTech Brand Book 2024 — digitalfirst palette.
# Used consistently across the Streamlit UI, Excel output, and PPTX deck.
BRAND: Dict[str, str] = {
    # ── Primary HCLTech brand colours ─────────────────────────────────────────
    "purple":    "#5F1EBE",   # Purple Heart  — primary CTA, gradients
    "blue":      "#3C91FF",   # Dodger Blue   — secondary, links
    "teal":      "#00A4A6",   # Persian Green — accents, success states
    "midnight":  "#00112B",   # Midnight      — header banner ONLY
    # ── Content / page colours (HIGH-CONTRAST white theme) ────────────────────
    "white":     "#FFFFFF",   # Pure white    — card + input backgrounds
    "page_bg":   "#F5F7FF",   # Page canvas   — soft blue-tinted white
    "light_bg":  "#EEF2FF",   # Highlight bg  — alternating rows, expanders
    "text_dark": "#0C1B3A",   # Primary text  — near-black navy (WCAG AAA)
    "text_mid":  "#334E7A",   # Secondary     — labels, sub-headings
    "text_soft": "#607090",   # Tertiary      — captions, hints
    "border":    "#C8D8F0",   # Borders       — inputs, cards
    # ── State colours ──────────────────────────────────────────────────────────
    "grey":      "#6B7280",
    "success":   "#0A8F5E",   # Deep green    — validated / approved
    "danger":    "#D41C1C",   # Deep red      — error / inactive
    "warning":   "#B86200",   # Deep amber    — in-progress / conditional
}

# ── Excel colour constants (openpyxl uses hex without '#') ────────────────────
XL_DARK_BLUE  = "001F5B"   # Section header background
XL_LIGHT_BLUE = "E8F0FE"   # Content cell background
XL_GRID_BLUE  = "3C91FF"   # Grid line colour
XL_ACCENT     = "5F1EBE"   # Title bar accent
XL_WHITE      = "FFFFFF"
XL_MID_BLUE   = "C5D5F5"   # Label cell background

# ── PPTX colour constants (RGBColor tuples) ───────────────────────────────────
# Each is constructed once and reused across all slide builder functions.
C_PURPLE = RGBColor(0x5F, 0x1E, 0xBE)
C_BLUE   = RGBColor(0x3C, 0x91, 0xFF)
C_TEAL   = RGBColor(0x00, 0xA4, 0xA6)
C_DARK   = RGBColor(0x00, 0x11, 0x2B)
C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
C_GREY   = RGBColor(0x6B, 0x72, 0x80)
C_LPUR   = RGBColor(0xEE, 0xE5, 0xFF)   # Light purple image placeholder fill

# ── Agent registry ────────────────────────────────────────────────────────────
# Index order is fixed and referenced throughout the codebase.
# Changing order here will automatically propagate to all loop-based references.
AGENT_NAMES: List[str] = [
    "Guide-Sheet Generator",   # 0
    "PowerPoint Creator",      # 1
    "Image Generator",         # 2
    "Audio Voice-Over Agent",  # 3
    "Deck Validator",          # 4
]
AGENT_ICONS: List[str] = ["📋", "📊", "🖼️", "🔊", "✅"]
NUM_AGENTS: int = len(AGENT_NAMES)

# ── Security constants ────────────────────────────────────────────────────────
# IMPORTANT: In a multi-user deployment replace with an environment variable:
#   ADMIN_PASSWORD = os.environ.get("LADA_ADMIN_PASSWORD", "")
# Never hard-code production passwords in source control.
ADMIN_PASSWORD: str = "EdTech@123"

# ── AI Provider constants ─────────────────────────────────────────────────────
# Two providers supported: "anthropic" (Claude) and "openai" (GPT-4o).
# Centralised so a model upgrade requires only one change here.
CLAUDE_MODEL:      str = "claude-sonnet-4-6"           # Anthropic Claude model
OPENAI_MODEL:      str = "gpt-4o"                      # OpenAI GPT-4o model
CLAUDE_MAX_TOKENS: int = 4096    # Anthropic output token ceiling
OPENAI_MAX_TOKENS: int = 4096    # OpenAI output token ceiling

# Provider display metadata — used in UI labels and token log
AI_PROVIDERS: Dict[str, Dict] = {
    "anthropic": {
        "label":       "Anthropic Claude",
        "icon":        "🤖",
        "model":       CLAUDE_MODEL,
        "placeholder": "sk-ant-api03-…",
        "help":        "Get your key at https://console.anthropic.com/api-keys",
        "prefix":      "sk-ant-",
    },
    "openai": {
        "label":       "OpenAI GPT-4o",
        "icon":        "⚡",
        "model":       OPENAI_MODEL,
        "placeholder": "sk-proj-… or sk-…",
        "help":        "Get your key at https://platform.openai.com/api-keys",
        "prefix":      "sk-",
    },
}

# ── Filename-safe characters whitelist pattern ────────────────────────────────
# Used by sanitise_filename() to strip dangerous characters from user input
# before embedding in download file names.
_UNSAFE_FILENAME_RE = re.compile(r'[^\w\s\-]')


# =============================================================================
# SECTION 2 — SESSION STATE INITIALISATION
# =============================================================================

def init_session_state() -> None:
    """
    Initialise all Streamlit session-state keys with safe defaults.

    This function is idempotent — it only sets keys that do not already exist,
    preserving any values set during the current session.  Call it once at
    module load time (before any widget rendering).

    Session-state schema
    --------------------
    admin_unlocked  bool        Whether the admin panel is currently unlocked.
    api_key         str         Active Anthropic API key (never written to disk).
    api_key_valid   bool|None   None=untested, True=valid, False=invalid.
    token_log       list[dict]  Per-call token usage records for Admin display.
    current_page    str         Active page: "landing" | "admin" | "workspace".
    current_agent   int         Index (0–4) of the agent currently shown/active.
    agent_status    list[str]   Per-agent lifecycle: "idle"|"running"|"done"|"error".
    agent_progress  list[int]   Per-agent progress percentage (0–100).
    agent_tokens    list[int]   Cumulative tokens consumed per agent.
    total_tokens    int         Grand-total tokens across all agents.
    form_data       dict        Captured landing-page inputs.
    guidesheet_txt  str         Raw text output from Agent 1.
    guidesheet_xl   bytes|None  Serialised Excel workbook bytes from Agent 1.
    pptx_files      list[dict]  List of {"name": str, "data": bytes} from Agent 2.
    pptx_script     str         Slide-by-slide specification text from Agent 2.
    image_specs     str         Image specification JSON text from Agent 3.
    audio_script    str         Enhanced voice-over scripts from Agent 4.
    audio_doc       bytes|None  Serialised TXT script package from Agent 4.
    validation_rpt  str         Validation scorecard text from Agent 5.
    review_comments dict        {agent_idx: str} reviewer comments.
    reviewer_names  dict        {agent_idx: str} reviewer names.
    job_name        str         Shorthand job identifier (= asset_name).
    color_palette   str         Active colour palette description string.
    uploaded_doc    bytes|None  Raw bytes of any uploaded curriculum document.
    """
    defaults: Dict = {
        "admin_unlocked":  False,
        "ai_provider":     "anthropic",  # "anthropic" | "openai"
        "api_key":         "",           # Active provider key (in-memory only)
        "openai_key":      "",           # OpenAI key store (in-memory only)
        "anthropic_key":   "",           # Anthropic key store (in-memory only)
        "api_key_valid":   None,         # None | True | False
        "openai_valid":    None,         # Separate validation state for OpenAI
        "token_log":       [],
        "current_page":    "landing",
        "current_agent":   0,
        "agent_status":    ["idle"] * NUM_AGENTS,
        "agent_progress":  [0]     * NUM_AGENTS,
        "agent_tokens":    [0]     * NUM_AGENTS,
        "total_tokens":    0,
        "form_data":       {},
        "guidesheet_txt":  "",
        "guidesheet_xl":   None,
        "pptx_files":      [],
        "pptx_script":     "",
        "image_specs":     "",
        "audio_script":    "",
        "audio_doc":       None,
        "validation_rpt":  "",
        "review_comments": {},
        "reviewer_names":  {},
        "job_name":        "",
        "color_palette":   "",
        "uploaded_doc":    None,
    }
    for key, default_val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = default_val


# Initialise session state immediately — must precede all widget calls.
init_session_state()

# ── Auto-load API key from .env (walks up to find the file) ──────────────────
def _load_env_key() -> None:
    """Read ANTHROPIC_API_KEY from the nearest .env file up the directory tree."""
    if st.session_state.get("anthropic_key"):
        return  # already set this session — don't overwrite
    script_dir = os.path.dirname(os.path.abspath(__file__))
    search = script_dir
    for _ in range(4):  # walk up at most 4 levels
        candidate = os.path.join(search, ".env")
        if os.path.isfile(candidate):
            try:
                from dotenv import dotenv_values
                vals = dotenv_values(candidate)
            except ImportError:
                # fallback: manual parse (no python-dotenv)
                vals = {}
                with open(candidate, encoding="utf-8") as fh:
                    for line in fh:
                        line = line.strip()
                        if line and not line.startswith("#") and "=" in line:
                            k, _, v = line.partition("=")
                            vals[k.strip()] = v.strip().strip('"').strip("'")
            key = vals.get("ANTHROPIC_API_KEY", "").strip()
            if key:
                st.session_state.anthropic_key  = key
                st.session_state.api_key        = key
                st.session_state.ai_provider    = "anthropic"
                st.session_state.api_key_valid  = True
            return
        parent = os.path.dirname(search)
        if parent == search:
            break
        search = parent

_load_env_key()


# =============================================================================
# SECTION 3 — GLOBAL CSS (HCLTech Brand)
# =============================================================================

def inject_global_css() -> None:
    """
    Inject global CSS for LADA v4.0.

    Design contract:
    ─────────────────────────────────────────────────────────────────
    TOP BANNER   → dark midnight gradient + Career Shaper logo block
    SIDEBAR      → dark midnight gradient (nav, agent tracker, tokens)
    ALL CONTENT  → white / soft-blue page canvas — high contrast text
    INPUTS       → white fill, dark navy text, purple focus ring
    BUTTONS      → purple-to-blue gradient, 2 px lift on hover
    CARDS        → white with subtle blue border and soft shadow
    ADMIN cards  → white with dark headings — always readable
    ─────────────────────────────────────────────────────────────────
    """
    p  = BRAND["purple"]
    b  = BRAND["blue"]
    t  = BRAND["teal"]
    mn = BRAND["midnight"]
    wh = BRAND["white"]
    pg = BRAND["page_bg"]
    lb = BRAND["light_bg"]
    td = BRAND["text_dark"]
    tm = BRAND["text_mid"]
    ts = BRAND["text_soft"]
    br = BRAND["border"]
    sc = BRAND["success"]
    dc = BRAND["danger"]
    wc = BRAND["warning"]

    css = f"""
<style>
/* ═══════════════════════════════════════════════════════════════
   FONTS
═══════════════════════════════════════════════════════════════ */
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@300;400;600;700;800&family=DM+Sans:wght@300;400;500;600&display=swap');

/* ═══════════════════════════════════════════════════════════════
   GLOBAL — soft white page canvas, dark readable text
   All content areas are bright. Only the top banner and
   sidebar retain the midnight dark background.
═══════════════════════════════════════════════════════════════ */
html, body, [class*="css"] {{
    font-family: 'DM Sans', sans-serif;
    background:  {pg};
    color:       {td};
}}

/* ── Main content block — WHITE canvas ─────────────────────── */
.main .block-container {{
    background:  {pg};
    padding:     1.5rem 2.2rem;
    max-width:   1440px;
}}

/* ═══════════════════════════════════════════════════════════════
   TOP BANNER — dark midnight with logo + title
   Applied to the .top-banner custom div rendered in each page.
═══════════════════════════════════════════════════════════════ */
.top-banner {{
    background:     linear-gradient(135deg, {mn} 0%, #002060 60%, #001840 100%);
    border-radius:  18px;
    padding:        1.6rem 2rem;
    margin-bottom:  1.6rem;
    display:        flex;
    align-items:    center;
    gap:            1.4rem;
    box-shadow:     0 6px 30px rgba(0,17,43,0.35);
    border:         1px solid rgba(63,145,255,0.25);
}}
.top-banner .banner-logo {{
    background:     linear-gradient(135deg, {p}, {b});
    border-radius:  12px;
    padding:        0.75rem 1.1rem;
    text-align:     center;
    min-width:      88px;
    box-shadow:     0 3px 12px rgba(95,30,190,0.35);
}}
.top-banner .banner-logo .logo-hcl {{
    font-family:    'Sora', sans-serif;
    font-size:      1.1rem;
    font-weight:    800;
    color:          #FFFFFF;
    letter-spacing: -0.01em;
    line-height:    1.1;
}}
.top-banner .banner-logo .logo-cs {{
    font-size:      0.58rem;
    color:          rgba(255,255,255,0.78);
    letter-spacing: 0.08em;
    text-transform: uppercase;
    margin-top:     2px;
}}
.top-banner .banner-text {{
    flex: 1;
}}
.top-banner .banner-title {{
    font-family:    'Sora', sans-serif;
    font-size:      1.7rem;
    font-weight:    800;
    background:     linear-gradient(135deg, #FFFFFF 0%, #C5DAFF 50%, #7FFFEF 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    line-height:    1.15;
    letter-spacing: -0.02em;
    margin-bottom:  0.15rem;
}}
.top-banner .banner-sub {{
    font-size:      0.78rem;
    color:          rgba(255,255,255,0.55);
    letter-spacing: 0.1em;
    text-transform: uppercase;
    font-weight:    500;
}}

/* ═══════════════════════════════════════════════════════════════
   PAGE HEADINGS (non-banner) — gradient text on white bg
═══════════════════════════════════════════════════════════════ */
.page-heading {{
    font-family:    'Sora', sans-serif;
    font-size:      1.55rem;
    font-weight:    800;
    background:     linear-gradient(135deg, {p}, {b});
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    margin-bottom:  0.1rem;
    line-height:    1.2;
}}
.page-sub {{
    font-size:      0.82rem;
    color:          {tm};
    letter-spacing: 0.09em;
    text-transform: uppercase;
    font-weight:    500;
    margin-bottom:  1.2rem;
}}

/* ═══════════════════════════════════════════════════════════════
   SIDEBAR — dark midnight (brand signature)
   Sidebar nav, agent tracker, token dashboard stay dark.
═══════════════════════════════════════════════════════════════ */
section[data-testid="stSidebar"] {{
    background: linear-gradient(180deg, {mn} 0%, #001840 65%, #002060 100%);
    border-right: 2px solid rgba(63,145,255,0.22);
}}
section[data-testid="stSidebar"] *,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] div {{ color: #FFFFFF !important; }}
section[data-testid="stSidebar"] hr {{
    border-color: rgba(255,255,255,0.13) !important;
}}
section[data-testid="stSidebar"] .stButton > button {{
    background:   rgba(255,255,255,0.1) !important;
    color:        #FFFFFF !important;
    border:       1px solid rgba(255,255,255,0.2) !important;
    border-radius:10px !important;
    font-weight:  600 !important;
    transition:   all 0.2s !important;
}}
section[data-testid="stSidebar"] .stButton > button:hover {{
    background:   linear-gradient(135deg,{p},{b}) !important;
    border-color: transparent !important;
    transform:    translateX(3px) !important;
}}

/* ═══════════════════════════════════════════════════════════════
   TYPOGRAPHY — white canvas, dark readable text
═══════════════════════════════════════════════════════════════ */
h1, h2, h3, h4 {{
    font-family: 'Sora', sans-serif;
    font-weight: 700;
    color:       {td};
}}

/* ═══════════════════════════════════════════════════════════════
   CARD COMPONENTS
   Standard: white bg + subtle border + soft shadow
   Accent:   left purple border highlight
   Info:     blue-tinted background strip
   Agent:    soft purple gradient strip (description)
═══════════════════════════════════════════════════════════════ */
.card {{
    background:    {wh};
    border:        1.5px solid {br};
    border-radius: 16px;
    padding:       1.4rem 1.75rem;
    margin-bottom: 1.1rem;
    box-shadow:    0 2px 14px rgba(60,145,255,0.07);
    color:         {td};
}}
.card-accent {{
    background:    {wh};
    border:        1.5px solid {br};
    border-left:   4px solid {p};
    border-radius: 16px;
    padding:       1.3rem 1.75rem;
    margin-bottom: 1.1rem;
    box-shadow:    0 2px 14px rgba(95,30,190,0.07);
    color:         {td};
}}
.card-info {{
    background:    #EBF4FF;
    border:        1.5px solid rgba(63,145,255,0.28);
    border-left:   4px solid {b};
    border-radius: 13px;
    padding:       0.85rem 1.25rem;
    margin-bottom: 0.85rem;
    color:         {td};
    font-size:     0.88rem;
    line-height:   1.65;
}}
.card-dark {{
    background:    {wh};
    border:        1.5px solid {br};
    border-left:   4px solid {p};
    border-radius: 16px;
    padding:       1.4rem 1.75rem;
    margin-bottom: 1.1rem;
    box-shadow:    0 2px 14px rgba(95,30,190,0.07);
    color:         {td};
}}
.agent-desc {{
    background:    linear-gradient(135deg,#F0F4FF,#EEF5FF);
    border:        1px solid rgba(95,30,190,0.14);
    border-radius: 13px;
    padding:       0.9rem 1.25rem;
    margin-bottom: 1rem;
    color:         {tm};
    font-size:     0.88rem;
    line-height:   1.65;
}}

/* ═══════════════════════════════════════════════════════════════
   ADMIN MODULE — always white cards, always dark text
═══════════════════════════════════════════════════════════════ */
.admin-card {{
    background:    {wh} !important;
    border:        1.5px solid {br} !important;
    border-radius: 16px !important;
    padding:       1.6rem 1.8rem !important;
    margin-bottom: 1.1rem !important;
    box-shadow:    0 3px 18px rgba(0,17,43,0.07) !important;
    color:         {td} !important;
}}
.admin-card h3, .admin-card h4, .admin-card p,
.admin-card label, .admin-card span {{ color: {td} !important; }}

/* ═══════════════════════════════════════════════════════════════
   KPI METRIC BOXES
═══════════════════════════════════════════════════════════════ */
.kpi-box {{
    background:    {wh};
    border:        1.5px solid {br};
    border-top:    4px solid {p};
    border-radius: 13px;
    padding:       1.1rem 1rem;
    text-align:    center;
    box-shadow:    0 2px 8px rgba(95,30,190,0.05);
}}
.kpi-label {{
    font-size:      0.7rem;
    font-weight:    700;
    color:          {ts};
    text-transform: uppercase;
    letter-spacing: 0.1em;
    margin-bottom:  0.3rem;
}}
.kpi-value {{
    font-family:  'Sora', sans-serif;
    font-size:    1.9rem;
    font-weight:  800;
    background:   linear-gradient(135deg,{p},{b});
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    line-height:  1;
}}

/* ═══════════════════════════════════════════════════════════════
   AGENT STEP CARDS (sidebar)
═══════════════════════════════════════════════════════════════ */
.agent-step {{
    display:       flex;
    align-items:   center;
    gap:           11px;
    padding:       9px 12px;
    border-radius: 11px;
    margin-bottom: 5px;
    border:        1px solid transparent;
    transition:    all 0.25s;
}}
.agent-step.active {{
    background:   linear-gradient(90deg,rgba(95,30,190,0.42),rgba(63,145,255,0.28));
    border-color: rgba(63,145,255,0.6);
}}
.agent-step.done {{
    background:   rgba(0,164,166,0.22);
    border-color: rgba(0,164,166,0.5);
}}
.agent-step.idle {{
    background:   rgba(255,255,255,0.05);
    border-color: rgba(255,255,255,0.1);
    opacity:      0.65;
}}

/* ═══════════════════════════════════════════════════════════════
   STATUS BADGES
═══════════════════════════════════════════════════════════════ */
.badge-active   {{ display:inline-block; background:{sc}; color:#fff; padding:4px 13px; border-radius:20px; font-size:0.77rem; font-weight:700; }}
.badge-inactive {{ display:inline-block; background:{dc}; color:#fff; padding:4px 13px; border-radius:20px; font-size:0.77rem; font-weight:700; }}
.badge-running  {{ display:inline-block; background:{wc}; color:#fff; padding:4px 13px; border-radius:20px; font-size:0.77rem; font-weight:700; }}

/* ═══════════════════════════════════════════════════════════════
   TOKEN COUNTER (sidebar — dark bg context)
   White text correct here — sidebar is intentionally dark
═══════════════════════════════════════════════════════════════ */
.token-box {{
    background:    rgba(255,255,255,0.1);
    border:        1px solid rgba(255,255,255,0.18);
    border-radius: 12px;
    padding:       0.9rem 1rem;
    text-align:    center;
    margin-bottom: 0.4rem;
}}
.token-number {{
    font-family:  'Sora', sans-serif;
    font-size:    1.8rem;
    font-weight:  800;
    color:        #FFFFFF;
}}
.token-label {{
    font-size:      0.68rem;
    font-weight:    700;
    color:          rgba(255,255,255,0.52);
    text-transform: uppercase;
    letter-spacing: 0.09em;
}}

/* ═══════════════════════════════════════════════════════════════
   INPUT FIELDS — white background, dark text, always readable
═══════════════════════════════════════════════════════════════ */
.stTextInput > div > div > input,
.stTextArea   > div > div > textarea,
.stSelectbox  > div > div > select,
.stNumberInput > div > div > input,
input[type="password"] {{
    background:    {wh} !important;
    border:        1.5px solid {br} !important;
    border-radius: 10px !important;
    color:         {td} !important;
    font-family:   'DM Sans', sans-serif !important;
    font-size:     0.9rem !important;
}}
.stTextInput > div > div > input:focus,
.stTextArea  > div > div > textarea:focus {{
    border-color: {p} !important;
    box-shadow:   0 0 0 3px rgba(95,30,190,0.1) !important;
    outline:      none !important;
}}
/* Input labels — always dark, always readable on white bg */
label, .stTextInput label, .stTextArea label,
.stSelectbox label, .stNumberInput label, .stFileUploader label {{
    color:       {td} !important;
    font-weight: 600 !important;
    font-size:   0.87rem !important;
}}

/* ═══════════════════════════════════════════════════════════════
   BUTTONS — gradient + lift on hover
═══════════════════════════════════════════════════════════════ */
.stButton > button {{
    background:     linear-gradient(135deg, {p}, {b}) !important;
    color:          #FFFFFF !important;
    border:         none !important;
    border-radius:  10px !important;
    font-family:    'Sora', sans-serif !important;
    font-weight:    600 !important;
    font-size:      0.87rem !important;
    padding:        0.55rem 1.4rem !important;
    transition:     all 0.22s !important;
    letter-spacing: 0.03em !important;
    box-shadow:     0 2px 8px rgba(95,30,190,0.18) !important;
}}
.stButton > button:hover {{
    transform:  translateY(-2px) !important;
    box-shadow: 0 8px 24px rgba(95,30,190,0.3) !important;
}}
.stButton > button:active {{ transform: translateY(0) !important; }}

/* ═══════════════════════════════════════════════════════════════
   PROGRESS BAR — branded gradient
═══════════════════════════════════════════════════════════════ */
.stProgress > div > div > div > div {{
    background: linear-gradient(90deg, {p}, {b}, {t}) !important;
    border-radius: 99px !important;
}}
.stProgress > div > div {{
    background:    {br} !important;
    border-radius: 99px !important;
}}

/* ═══════════════════════════════════════════════════════════════
   EXPANDERS — white body, readable header
═══════════════════════════════════════════════════════════════ */
.streamlit-expanderHeader {{
    background:    {lb} !important;
    border:        1.5px solid {br} !important;
    border-radius: 10px !important;
    color:         {td} !important;
    font-weight:   600 !important;
    padding:       0.65rem 1rem !important;
}}
.streamlit-expanderContent {{
    background:    {wh} !important;
    border:        1.5px solid {br} !important;
    border-top:    none !important;
    border-radius: 0 0 10px 10px !important;
    padding:       1rem !important;
    color:         {td} !important;
}}

/* ═══════════════════════════════════════════════════════════════
   TABS — brand styled, readable
═══════════════════════════════════════════════════════════════ */
.stTabs [data-baseweb="tab-list"] {{
    background:    {lb};
    border-radius: 12px;
    padding:       4px;
    gap:           4px;
    border:        1px solid {br};
}}
.stTabs [data-baseweb="tab"] {{
    border-radius: 8px !important;
    color:         {tm} !important;
    font-weight:   600 !important;
    font-size:     0.84rem !important;
}}
.stTabs [aria-selected="true"] {{
    background: linear-gradient(135deg,{p},{b}) !important;
    color:      #FFFFFF !important;
}}

/* ═══════════════════════════════════════════════════════════════
   MISC — file uploader, dividers, alerts, data tables
═══════════════════════════════════════════════════════════════ */
.stFileUploader > div {{
    background:    {lb} !important;
    border:        2px dashed rgba(63,145,255,0.38) !important;
    border-radius: 12px !important;
    color:         {tm} !important;
}}
hr {{
    border:     none !important;
    border-top: 1.5px solid {br} !important;
    margin:     1.4rem 0 !important;
}}
.stAlert {{ border-radius: 12px !important; }}
.stDataFrame {{
    border-radius: 12px;
    overflow:      hidden;
    border:        1.5px solid {br};
}}
.stCaption, small {{ color: {ts} !important; font-size: 0.8rem !important; }}
/* Content preview block — soft grey, always readable */
.preview-block {{
    background:    #F3F6FC;
    border:        1.5px solid {br};
    border-radius: 12px;
    padding:       1.2rem 1.5rem;
    font-size:     0.83rem;
    line-height:   1.75;
    white-space:   pre-wrap;
    color:         {td};
    max-height:    480px;
    overflow-y:    auto;
    font-family:   'DM Sans', sans-serif;
}}
/* Review panel */
.review-panel {{
    background:    {lb};
    border:        1.5px solid {br};
    border-radius: 13px;
    padding:       1.1rem 1.5rem;
    margin-top:    1rem;
    color:         {td};
}}
.section-label {{
    font-family:    'Sora', sans-serif;
    font-size:      0.78rem;
    font-weight:    700;
    color:          {p};
    letter-spacing: 0.1em;
    text-transform: uppercase;
    margin-bottom:  0.5rem;
    padding-bottom: 0.35rem;
    border-bottom:  2px solid {br};
}}
.stSlider label {{ color: {td} !important; font-weight: 600 !important; }}
</style>
"""
    st.markdown(css, unsafe_allow_html=True)
inject_global_css()


# =============================================================================
# SECTION 4 — SECURITY & UTILITY HELPERS
# =============================================================================

def sanitise_filename(raw: str, max_len: int = 60) -> str:
    """
    Produce a filesystem-safe filename stem from arbitrary user input.

    Steps:
      1. Strip leading/trailing whitespace.
      2. Replace spaces with underscores.
      3. Remove every character that is not alphanumeric, underscore, or hyphen.
      4. Truncate to `max_len` characters.
      5. Fall back to "Asset" if the result is empty.

    Args:
        raw:     Raw string from the user (e.g. asset name).
        max_len: Maximum number of characters to retain.

    Returns:
        A clean string safe for use as a filename stem.

    Examples:
        >>> sanitise_filename("Gen AI / Banking (2026)")
        'Gen_AI__Banking_2026'
        >>> sanitise_filename("")
        'Asset'
    """
    stem = raw.strip().replace(" ", "_")
    stem = _UNSAFE_FILENAME_RE.sub("", stem)
    stem = stem[:max_len]
    return stem if stem else "Asset"


def mask_api_key(key: str) -> str:
    """
    Return a partially masked representation of an API key for safe display.

    Reveals the first 12 characters (vendor prefix) and the last 6 characters
    (short suffix to confirm the right key is loaded), replacing the secret
    middle portion with '...'.

    Args:
        key: Full API key string.

    Returns:
        Masked key string, e.g. 'sk-ant-api03-...xyzABC'.
        Returns an empty string if `key` is empty.

    Security:
        Never log, print, or store the unmasked return value of this function.
        Only render it via st.caption() or similar non-persistent UI elements.
    """
    if not key:
        return ""
    if len(key) <= 18:
        # Key too short to mask meaningfully — show only a placeholder.
        return "sk-ant-***"
    return f"{key[:12]}...{key[-6:]}"


def validate_anthropic_key(key: str) -> bool:
    """
    Validate an Anthropic API key via a minimal 10-token live call.
    Only exception TYPE is logged — key content never exposed.
    """
    if not key or not key.startswith("sk-ant-"):
        logger.warning("Anthropic key: format check failed.")
        return False
    try:
        client = anthropic.Anthropic(api_key=key)
        client.messages.create(
            model=CLAUDE_MODEL, max_tokens=10,
            messages=[{"role": "user", "content": "Hi"}],
        )
        logger.info("Anthropic key: validated OK.")
        return True
    except anthropic.AuthenticationError:
        logger.warning("Anthropic key: AuthenticationError.")
        return False
    except anthropic.PermissionDeniedError:
        logger.warning("Anthropic key: PermissionDeniedError.")
        return False
    except Exception as exc:   # noqa: BLE001
        logger.warning("Anthropic key: unexpected %s.", type(exc).__name__)
        return False


def validate_openai_key(key: str) -> bool:
    """
    Validate an OpenAI API key via a minimal 10-token live call.
    Only exception TYPE is logged — key content never exposed.
    """
    if openai is None:
        logger.warning("OpenAI key: openai package not installed.")
        return False
    if not key or not key.startswith("sk-"):
        logger.warning("OpenAI key: format check failed.")
        return False
    try:
        client = openai.OpenAI(api_key=key)
        client.chat.completions.create(
            model=OPENAI_MODEL,
            max_tokens=10,
            messages=[{"role": "user", "content": "Hi"}],
        )
        logger.info("OpenAI key: validated OK.")
        return True
    except _OAIAuthErr:
        # Key rejected by OpenAI — wrong, expired, or revoked
        logger.warning("OpenAI key: AuthenticationError.")
        return False
    except _OAIConnErr:
        # Network issue — key may be fine, report without invalidating
        logger.warning("OpenAI key: ConnectionError.")
        return False
    except Exception as exc:   # noqa: BLE001
        logger.warning("OpenAI key: unexpected %s.", type(exc).__name__)
        return False


def validate_api_key(key: str) -> bool:
    """
    Backwards-compatible wrapper — routes to the active provider's validator.
    Used by the admin UI's Validate & Save button.
    """
    provider = st.session_state.get("ai_provider", "anthropic")
    if provider == "openai":
        return validate_openai_key(key)
    return validate_anthropic_key(key)


def get_key_status_html() -> str:
    """
    Return an HTML badge string reflecting the current API key status.

    Reads st.session_state.api_key_valid:
      None  → 🔴 In-Active  (no key entered)
      True  → 🟢 Active     (key validated successfully)
      False → 🔴 Use another key

    Returns:
        An HTML <span> element with appropriate CSS class and label.
    """
    state = st.session_state.api_key_valid
    if state is True:
        return '<span class="badge-active">🟢 Active</span>'
    if state is False:
        return '<span class="badge-inactive">🔴 Use another key</span>'
    return '<span class="badge-inactive">🔴 In-Active</span>'


# =============================================================================
# SECTION 5 — CENTRALISED API LAYER
# =============================================================================

def call_ai(
    prompt:    str,
    system:    str = "",
    agent_idx: int = 0,
    job_name:  str = "",
) -> str:
    """
    Unified AI gateway — routes to Anthropic Claude OR OpenAI GPT-4o based on
    the active provider selected in the Admin Module.

    Responsibilities:
      - Read active provider + API key from session state (never as parameters)
      - Instantiate a fresh, stateless client per call
      - Call the appropriate SDK with normalised message format
      - Account for tokens per agent and append to the token log
      - Handle provider-specific exceptions with friendly messages
      - Always return a string — never raises an exception

    Args:
        prompt    : User-turn message for the AI model.
        system    : Optional system instruction (role / behaviour framing).
        agent_idx : Index 0-4 into AGENT_NAMES for token attribution.
        job_name  : Label for the token log row.

    Returns:
        AI text response on success, or "❌ …" error string on failure.

    Security:
        API keys are read from session_state at call-time; never accepted as
        parameters or logged in any form.
    """
    api_key  = st.session_state.api_key
    provider = st.session_state.get("ai_provider", "anthropic")

    if not api_key:
        return "❌ No API key configured. Please set one in the Admin Module."

    eff_job = job_name or st.session_state.get("job_name", "Unnamed Job")

    # ── Route to Anthropic Claude ─────────────────────────────────────────────
    if provider == "anthropic":
        try:
            client = anthropic.Anthropic(api_key=api_key)
            kwargs = {
                "model":      CLAUDE_MODEL,
                "max_tokens": CLAUDE_MAX_TOKENS,
                "messages":   [{"role": "user", "content": prompt}],
            }
            if system:
                kwargs["system"] = system
            response      = client.messages.create(**kwargs)
            call_tokens   = response.usage.input_tokens + response.usage.output_tokens
            response_text = response.content[0].text

        except anthropic.AuthenticationError:
            st.session_state.api_key_valid = False
            logger.error("call_ai[anthropic]: AuthenticationError.")
            return "❌ Anthropic authentication failed. Please enter a new Claude API key in Admin."
        except anthropic.RateLimitError:
            logger.warning("call_ai[anthropic]: RateLimitError.")
            return "❌ Anthropic rate limit reached. Please wait 60 seconds and retry."
        except anthropic.APIStatusError as exc:
            logger.error("call_ai[anthropic]: APIStatusError %s", exc.status_code)
            return f"❌ Anthropic API error (HTTP {exc.status_code}). Please retry."
        except Exception as exc:   # noqa: BLE001
            logger.exception("call_ai[anthropic]: %s", type(exc).__name__)
            return f"❌ Unexpected error ({type(exc).__name__}). Please retry."

    # ── Route to OpenAI GPT-4o ────────────────────────────────────────────────
    elif provider == "openai":
        if openai is None:
            return "❌ OpenAI package not installed. Run: pip install openai"
        try:
            client   = openai.OpenAI(api_key=api_key)
            messages = []
            if system:
                messages.append({"role": "system", "content": system})
            messages.append({"role": "user", "content": prompt})

            response = client.chat.completions.create(
                model=OPENAI_MODEL,
                max_tokens=OPENAI_MAX_TOKENS,
                messages=messages,
            )
            call_tokens   = (response.usage.prompt_tokens +
                             response.usage.completion_tokens)
            response_text = response.choices[0].message.content or ""

        except _OAIAuthErr:
            # Invalid key — mark immediately so the badge updates
            st.session_state.api_key_valid = False
            logger.error("call_ai[openai]: AuthenticationError.")
            return "❌ OpenAI authentication failed. Please enter a valid OpenAI API key in Admin."
        except _OAIRateErr:
            logger.warning("call_ai[openai]: RateLimitError.")
            return "❌ OpenAI rate limit reached. Please wait 60 seconds and retry."
        except _OAIStatusErr as exc:
            logger.error("call_ai[openai]: APIStatusError %s", exc.status_code)
            return f"❌ OpenAI API error (HTTP {exc.status_code}). Please retry."
        except _OAIConnErr:
            logger.warning("call_ai[openai]: ConnectionError.")
            return "❌ Cannot reach OpenAI. Check your internet connection and retry."
        except Exception as exc:   # noqa: BLE001
            logger.exception("call_ai[openai]: %s", type(exc).__name__)
            return f"❌ Unexpected error ({type(exc).__name__}). Please retry."

    else:
        return f"❌ Unknown AI provider: {provider}. Please select Anthropic or OpenAI in Admin."

    # ── Token accounting (shared for both providers) ──────────────────────────
    st.session_state.agent_tokens[agent_idx] += call_tokens
    st.session_state.total_tokens            += call_tokens
    provider_info = AI_PROVIDERS.get(provider, {})
    st.session_state.token_log.append({
        "Sr. No":      len(st.session_state.token_log) + 1,
        "Job Name":    eff_job,
        "Agent Name":  AGENT_NAMES[agent_idx],
        "Provider":    provider_info.get("label", provider),
        "Model":       provider_info.get("model", ""),
        "Tokens Used": call_tokens,
        "Timestamp":   datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })
    logger.info("call_ai OK | provider=%s | agent=%s | tokens=%d",
                provider, AGENT_NAMES[agent_idx], call_tokens)
    return response_text


# ── Backwards-compatible alias ────────────────────────────────────────────────
# All agent runner functions call call_claude(); this alias ensures they route
# through the new dual-provider gateway without any per-agent code changes.
call_claude = call_ai


def build_context_string() -> str:
    """
    Assemble a compact, structured context block from the Landing Page form data.

    This string is prepended to every agent prompt so Claude always has the full
    programme context regardless of which agent is running.

    Returns:
        Multi-line string with labelled fields.  Empty if form_data is unset.
    """
    fd = st.session_state.form_data
    if not fd:
        return "[No programme context available — please complete the Landing Page form.]"

    lines = [
        f"Asset Name      : {fd.get('asset_name', '')}",
        f"Entity          : {fd.get('entity_name', '')}",
        f"Duration        : {fd.get('duration', '')} hours",
        f"Target Audience : {fd.get('target_audience', '')}",
        f"Content Coverage: {fd.get('content_desc', '')}",
        f"Topics/Modules  : {fd.get('topics_subtopics', '')}",
        f"Colour Palette  : {st.session_state.color_palette}",
    ]
    return "\n".join(lines)


# =============================================================================
# SECTION 6 — SHARED UI COMPONENTS
# =============================================================================

def render_sidebar() -> None:
    """
    Render the persistent left sidebar with branding, navigation, agent
    progress tracker, and live token dashboard.

    Navigation buttons update st.session_state.current_page and rerun the app.
    Agent progress cards show status badges, per-agent progress bars, and token counts.
    """
    with st.sidebar:
        # ── Brand logo area ────────────────────────────────────────────────────
        st.markdown("""
        <div style="text-align:center; padding:1rem 0 0.5rem;">
            <div style="font-family:'Sora',sans-serif; font-size:1.5rem; font-weight:800;
                        background:linear-gradient(135deg,#3C91FF,#5F1EBE);
                        -webkit-background-clip:text; -webkit-text-fill-color:transparent;
                        background-clip:text;">HCLTech</div>
            <div style="font-size:0.75rem; color:rgba(255,255,255,0.5);
                        letter-spacing:0.1em;">Career Shaper™</div>
        </div><hr/>
        """, unsafe_allow_html=True)

        # ── Navigation ────────────────────────────────────────────────────────
        st.markdown("### 🧭 Navigation")
        nav_items = [
            ("🏠 Landing Page",    "landing"),
            ("🔐 Admin Module",    "admin"),
            ("🚀 Agent Workspace", "workspace"),
        ]
        for label, page_key in nav_items:
            if st.button(label, key=f"nav_{page_key}", use_container_width=True):
                st.session_state.current_page = page_key
                st.rerun()

        st.markdown("---")

        # ── Agent progress tracker ────────────────────────────────────────────
        st.markdown("### 🤖 Agent Progress")
        for idx, (name, icon) in enumerate(zip(AGENT_NAMES, AGENT_ICONS)):
            status   = st.session_state.agent_status[idx]
            progress = st.session_state.agent_progress[idx]

            # Determine CSS class for the step card.
            is_current = (idx == st.session_state.current_agent and status == "running")
            css_class  = "active" if is_current else ("done" if status == "done" else "idle")

            # Badge emoji for quick visual scan.
            status_icon = "✅" if status == "done" else "⏳" if status == "running" else "⬜"

            # Label colour per state.
            label_color = (
                "#3C91FF" if css_class == "active" else
                "#00A4A6" if css_class == "done"   else
                "rgba(255,255,255,0.5)"
            )

            st.markdown(f"""
            <div class="agent-step {css_class}">
                <span style="font-size:1.2rem">{icon}</span>
                <div style="flex:1">
                    <div style="font-size:0.8rem; font-weight:600; color:{label_color}">
                        {status_icon} Agent {idx + 1}
                    </div>
                    <div style="font-size:0.7rem; color:rgba(255,255,255,0.4)">{name}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # Show progress bar and token count only after the agent has started.
            if status in ("running", "done"):
                st.progress(progress / 100)
                st.caption(f"  {st.session_state.agent_tokens[idx]:,} tokens")

        st.markdown("---")

        # ── Token dashboard ───────────────────────────────────────────────────
        provider = st.session_state.get("ai_provider", "anthropic")
        pinfo    = AI_PROVIDERS[provider]
        st.markdown(f"### 📊 Token Dashboard")
        st.markdown(
            f'<div style="font-size:0.68rem;color:rgba(255,255,255,0.45);margin-bottom:6px;">' +
            f'{pinfo["icon"]} Active: {pinfo["label"]}</div>',
            unsafe_allow_html=True)
        st.markdown(f"""
        <div class="token-box">
            <div style="font-size:0.75rem; color:rgba(255,255,255,0.5); margin-bottom:4px;">
                TOTAL TOKENS USED
            </div>
            <div class="token-number">{st.session_state.total_tokens:,}</div>
        </div>
        """, unsafe_allow_html=True)

        for idx, (name, icon) in enumerate(zip(AGENT_NAMES, AGENT_ICONS)):
            if st.session_state.agent_tokens[idx] > 0:
                st.caption(f"{icon} {name}: {st.session_state.agent_tokens[idx]:,}")


def render_review_gate(agent_idx: int) -> None:
    """
    Render the review panel that appears after each agent completes.

    Captures:
      - Reviewer's full name
      - Free-text review comments / change requests
    And persists both to session state under agent_idx keys.

    Args:
        agent_idx: Index (0–4) of the agent whose output is being reviewed.
    """
    st.markdown("---")
    st.markdown("#### 📝 Review & Feedback")

    col_name, col_comment = st.columns([2, 3])

    with col_name:
        reviewer_name = st.text_input(
            "Reviewer Name",
            placeholder="Enter reviewer's full name",
            key=f"reviewer_name_{agent_idx}",
            value=st.session_state.reviewer_names.get(agent_idx, ""),
        )

    with col_comment:
        review_text = st.text_area(
            "Review Comments",
            placeholder="Enter review comments, suggestions, or change requests…",
            key=f"review_comment_{agent_idx}",
            height=90,
            value=st.session_state.review_comments.get(agent_idx, ""),
        )

    if st.button("💾 Save Review", key=f"save_review_{agent_idx}"):
        # Persist to session state — only update if non-empty.
        if reviewer_name:
            st.session_state.reviewer_names[agent_idx] = reviewer_name
        if review_text:
            st.session_state.review_comments[agent_idx] = review_text
        st.success(
            f"✅ Review saved — Agent {agent_idx + 1}: {AGENT_NAMES[agent_idx]}"
        )
        logger.info(
            "Review saved | agent=%s | reviewer=%s",
            AGENT_NAMES[agent_idx], reviewer_name or "(unnamed)",
        )


def render_next_agent_button(current_idx: int) -> bool:
    """
    Render a primary "Proceed to next agent" button.

    Args:
        current_idx: Index of the agent that just completed (0–3).

    Returns:
        True if the button was clicked (caller should update state + rerun).
        False if current_idx is the last agent (4) or button not clicked.
    """
    if current_idx >= NUM_AGENTS - 1:
        return False
    next_name = AGENT_NAMES[current_idx + 1]
    return st.button(
        f"➡️ Proceed to Agent {current_idx + 2}: {next_name}",
        key=f"next_agent_{current_idx}",
        type="primary",
    )


def render_rerun_button(agent_idx: int, label: str = "") -> bool:
    """
    Render a secondary "Re-run this agent" button and reset its state on click.

    Resets agent_status[agent_idx] to "idle" and agent_progress[agent_idx] to 0.

    Args:
        agent_idx: Agent to reset.
        label:     Optional custom button label.

    Returns:
        True if the button was clicked (caller should st.rerun()).
    """
    btn_label = label or f"🔄 Re-run Agent {agent_idx + 1}"
    if st.button(btn_label, key=f"rerun_agent_{agent_idx}"):
        st.session_state.agent_status[agent_idx]   = "idle"
        st.session_state.agent_progress[agent_idx] = 0
        logger.info("Agent %s reset to idle.", AGENT_NAMES[agent_idx])
        return True
    return False


# =============================================================================
# SECTION 7 — PAGE: ADMIN MODULE
# =============================================================================

def render_admin() -> None:
    """
    Render the password-protected Admin Module page.

    Features:
      - Password authentication gate (ADMIN_PASSWORD)
      - API key input with live validation (validate_api_key)
      - Masked key display and status badge
      - Optional API test prompt for ad-hoc key verification
      - Token usage log table (Sr. No, Job Name, Agent, Tokens, Timestamp)
      - Log clear button
      - Lock Admin & Exit button
    """
    # ── Dark top banner ───────────────────────────────────────────────────────
    st.markdown("""
    <div class="top-banner">
        <div class="banner-logo"><div class="logo-hcl">HCLTech</div>
        <div class="logo-cs">Career Shaper™</div></div>
        <div class="banner-text">
            <div class="banner-title">Admin Module</div>
            <div class="banner-sub">System Configuration &amp; Usage Analytics</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Authentication gate ───────────────────────────────────────────────────
    if not st.session_state.admin_unlocked:
        _, col_centre, _ = st.columns([1, 2, 1])
        with col_centre:
            st.markdown('<div class="admin-card">', unsafe_allow_html=True)
            st.markdown("### 🔑 Admin Authentication")
            st.markdown("Enter the admin password to access configuration settings.")

            pwd_input = st.text_input(
                "Password",
                type="password",
                key="admin_pwd_field",
                placeholder="Enter admin password",
            )
            if st.button("🔓 Unlock Admin", use_container_width=True):
                if pwd_input == ADMIN_PASSWORD:
                    st.session_state.admin_unlocked = True
                    logger.info("Admin panel unlocked.")
                    st.success("✅ Admin access granted!")
                    time.sleep(0.4)
                    st.rerun()
                else:
                    logger.warning("Admin unlock failed — incorrect password attempt.")
                    st.error("❌ Incorrect password.")
            st.markdown("</div>", unsafe_allow_html=True)
        return  # Do not render the rest of the admin panel until authenticated.

    # ── AI Provider & API Key Configuration ─────────────────────────────────
    st.markdown('<div class="admin-card">', unsafe_allow_html=True)
    col_title, col_badge = st.columns([3, 1])
    with col_title:
        st.markdown('### 🗝️ AI Provider & API Key Configuration')
    with col_badge:
        st.markdown(get_key_status_html(), unsafe_allow_html=True)

    st.markdown(
        '<p style="color:#334E7A;font-size:0.87rem;margin:0.3rem 0 1rem;">' +
        "Keys are stored in-session only — never written to disk. " +
        "Switch providers freely; each key is remembered independently.</p>",
        unsafe_allow_html=True,
    )

    # ── Step 1: Choose AI Provider ────────────────────────────────────────────
    st.markdown("#### Step 1 — Choose AI Provider")
    col_p1, col_p2 = st.columns(2)

    with col_p1:
        if st.button(
            "🤖  Anthropic Claude  (claude-sonnet-4)",
            key="pick_anthropic",
            use_container_width=True,
        ):
            st.session_state.ai_provider     = "anthropic"
            st.session_state.api_key         = st.session_state.anthropic_key
            st.session_state.api_key_valid   = (
                True if st.session_state.anthropic_key and
                        st.session_state.get("anthropic_validated") else None
            )
            st.rerun()

    with col_p2:
        if st.button(
            "⚡  OpenAI GPT-4o  (gpt-4o)",
            key="pick_openai",
            use_container_width=True,
        ):
            st.session_state.ai_provider     = "openai"
            st.session_state.api_key         = st.session_state.openai_key
            st.session_state.api_key_valid   = (
                True if st.session_state.openai_key and
                        st.session_state.get("openai_validated") else None
            )
            st.rerun()

    # Highlight active provider
    provider = st.session_state.get("ai_provider", "anthropic")
    pinfo    = AI_PROVIDERS[provider]
    st.markdown(
        f'<div style="background:#EEF2FF;border:2px solid #5F1EBE;border-radius:12px;' +
        f'padding:0.7rem 1.2rem;margin:0.6rem 0 1.2rem;font-weight:700;' +
        f'color:#0C1B3A;font-size:0.9rem;">' +
        f'{pinfo["icon"]}  Active provider: <strong>{pinfo["label"]}</strong> &nbsp;·&nbsp;' +
        f'Model: <code>{pinfo["model"]}</code></div>',
        unsafe_allow_html=True,
    )

    # ── Step 2: Enter & Validate API Key ─────────────────────────────────────
    st.markdown("#### Step 2 — Enter & Validate API Key")

    # Retrieve stored key for active provider (so it pre-fills on switch)
    stored_key = (
        st.session_state.anthropic_key if provider == "anthropic"
        else st.session_state.openai_key
    )

    col_key, col_btn = st.columns([4, 1])
    with col_key:
        new_key_input = st.text_input(
            f"{pinfo['icon']}  {pinfo['label']} API Key",
            type="password",
            value=stored_key,
            placeholder=pinfo["placeholder"],
            help=pinfo["help"] + "  ·  Key stored in-session only, never on disk.",
        )
    with col_btn:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔍  Validate & Save", use_container_width=True, key="validate_btn"):
            if new_key_input:
                spinner_label = f"Validating with {pinfo['label']}…"
                with st.spinner(spinner_label):
                    is_valid = validate_api_key(new_key_input)

                # Persist to provider-specific store AND active key slot
                st.session_state.api_key       = new_key_input
                st.session_state.api_key_valid = is_valid

                if provider == "anthropic":
                    st.session_state.anthropic_key        = new_key_input
                    st.session_state["anthropic_validated"] = is_valid
                else:
                    st.session_state.openai_key           = new_key_input
                    st.session_state["openai_validated"]   = is_valid

                if is_valid:
                    st.success(f"🟢 {pinfo['label']} key validated — **Active** and ready.")
                else:
                    st.error(f"🔴 Validation failed. Please check your {pinfo['label']} key.")
                st.rerun()
            else:
                st.warning("Please enter an API key before validating.")

    # Show masked key and status
    if st.session_state.api_key:
        masked = mask_api_key(st.session_state.api_key)
        st.markdown(
            f'<p style="color:#607090;font-size:0.82rem;margin-top:0.5rem;">' +
            f'Current {pinfo["label"]} key: ' +
            f'<code style="background:#F0F4FF;padding:2px 7px;border-radius:5px;' +
            f'color:#0C1B3A;">{masked}</code>' +
            f'&nbsp;&nbsp;{get_key_status_html()}</p>',
            unsafe_allow_html=True,
        )

    # Lock button
    if st.button("🔒 Lock Admin & Exit"):
        st.session_state.admin_unlocked = False
        st.session_state.current_page   = "landing"
        logger.info("Admin panel locked.")
        st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    # ── API Test Prompt ───────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown('<div class="admin-card">', unsafe_allow_html=True)
    provider = st.session_state.get("ai_provider", "anthropic")
    pinfo    = AI_PROVIDERS[provider]
    st.markdown(f"### 🧪 API Test — {pinfo['icon']} {pinfo['label']}")
    st.markdown(
        f'<p style="color:#334E7A;font-size:0.87rem;margin:0 0 0.9rem;">' +
        f"Send a test message to <strong>{pinfo['label']}</strong> to verify your key " +
        f"is working and preview response quality before running agents.</p>",
        unsafe_allow_html=True,
    )
    test_col1, test_col2 = st.columns([4, 1])
    with test_col1:
        test_prompt = st.text_input(
            "Test Prompt",
            placeholder="e.g.  Explain Generative AI in two sentences.",
            key="admin_test_prompt",
            label_visibility="collapsed",
        )
    with test_col2:
        run_test = st.button("▶️ Run Test", use_container_width=True, key="admin_test_btn")

    if run_test:
        if not test_prompt:
            st.warning("Enter a prompt to test.")
        elif not st.session_state.api_key or st.session_state.api_key_valid is not True:
            st.error(f"Please validate your {pinfo['label']} key first (Step 2 above).")
        else:
            with st.spinner(f"Calling {pinfo['label']}…"):
                test_response = call_ai(
                    prompt="[Admin Test] " + test_prompt,
                    system="You are a helpful assistant. Respond concisely.",
                    agent_idx=0,
                    job_name="Admin Test",
                )
            st.markdown(
                f'<div class="preview-block">{test_response}</div>',
                unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # ── Token Usage Log ───────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 📈 Token Usage Log")

    if st.session_state.token_log:
        import pandas as pd

        # Summary KPI row.
        total_calls = len(st.session_state.token_log)
        st.markdown(f"""
        <div style="display:flex; gap:1rem; margin-bottom:1rem; flex-wrap:wrap;">
            <div class="token-box" style="flex:1; min-width:140px;">
                <div style="font-size:0.7rem;color:rgba(255,255,255,0.5)">TOTAL TOKENS</div>
                <div class="token-number">{st.session_state.total_tokens:,}</div>
            </div>
            <div class="token-box" style="flex:1; min-width:140px;">
                <div style="font-size:0.7rem;color:rgba(255,255,255,0.5)">API CALLS</div>
                <div class="token-number">{total_calls}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        df_log = pd.DataFrame(st.session_state.token_log)
        # Show Provider/Model columns only if they exist (added in v4.1+)
        base_cols = ["Sr. No", "Job Name", "Agent Name", "Tokens Used", "Timestamp"]
        extra_cols = [c for c in ["Provider", "Model"] if c in df_log.columns]
        display_cols = ["Sr. No", "Job Name", "Agent Name"] + extra_cols + ["Tokens Used", "Timestamp"]
        st.dataframe(
            df_log[[c for c in display_cols if c in df_log.columns]],
            use_container_width=True,
            hide_index=True,
        )

        if st.button("🗑️ Clear Token Log"):
            st.session_state.token_log    = []
            st.session_state.agent_tokens = [0] * NUM_AGENTS
            st.session_state.total_tokens = 0
            logger.info("Token log cleared by admin.")
            st.rerun()
    else:
        st.info("No token usage recorded yet. Run an agent to see logs here.")


# =============================================================================
# SECTION 8 — PAGE: LANDING / INPUT FORM
# =============================================================================

def render_landing() -> None:
    """
    Render the Landing Page — the primary input capture form.

    All inputs are persisted to st.session_state.form_data on submission.
    Clicking START validates required fields and API key status, then
    navigates to the Agent Workspace.

    Required fields  : asset_name, entity_name, content_desc (validated before start)
    Optional fields  : target_audience, color_palette, uploaded_doc, topics_subtopics
    """
    # ── Header with logo ──────────────────────────────────────────────────────
    col_logo, col_title = st.columns([1, 5])
    with col_logo:
        # Attempt to display an uploaded logo; fall back to a branded text badge.
        logo_path = "/mnt/user-data/uploads/Career-Shaper_Logo_1.jpg"
        if os.path.isfile(logo_path):
            st.image(logo_path, width=110)
        else:
            st.markdown("""
            <div style="background:linear-gradient(135deg,#5F1EBE,#3C91FF);
                        border-radius:12px;padding:1rem;text-align:center;
                        font-family:'Sora',sans-serif;font-weight:800;
                        font-size:0.9rem;color:white;width:110px;">
                HCLTech<br>
                <span style="font-size:0.65rem;font-weight:400">Career Shaper™</span>
            </div>
            """, unsafe_allow_html=True)

    with col_title:
        st.markdown(
            '<div class="hero-title">LEARNING ASSET DEVELOPMENT AGENT</div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            '<div class="hero-sub">AI-Powered Multi-Agent Content Engineering Platform</div>',
            unsafe_allow_html=True,
        )

    # ── API key warning banner ─────────────────────────────────────────────────
    if not st.session_state.api_key or st.session_state.api_key_valid is not True:
        provider = st.session_state.get("ai_provider", "anthropic")
        pinfo    = AI_PROVIDERS[provider]
        st.warning(
            f"⚠️ No active API key detected. "
            f"Please configure your **{pinfo['label']}** key in **🔐 Admin Module** → "
            f"Step 1 (choose provider) → Step 2 (enter & validate key)."
        )

    # ── Sample input guide (collapsible) ──────────────────────────────────────
    with st.expander("💡 What can this agent process?  —  Sample Input Guide", expanded=False):
        st.markdown("""
        <div class="card" style="margin:0;">
        <b>📚 Subjects & Curriculum:</b>
        AI &amp; Machine Learning, Data Engineering, Cloud Computing, Cybersecurity,
        Python Programming, DevOps, Agile &amp; Scrum, Data Governance, Blockchain<br><br>

        <b>📋 Coverage Examples:</b><br>
        &nbsp;&nbsp;• "Generative AI covering LLMs, Prompt Engineering, RAG,
           Fine-tuning and AI Ethics — 16 hours"<br>
        &nbsp;&nbsp;• "Python for Data Science: NumPy, Pandas, Matplotlib,
           Scikit-learn — 24 hours"<br>
        &nbsp;&nbsp;• "Azure Data Engineering: ADF, Databricks, Synapse Analytics — 32 hours"<br>
        &nbsp;&nbsp;• "Agile &amp; Scrum for Banking Teams — 8 hours"<br><br>

        <b>🏦 Target Entities:</b>
        Banks, Universities, Enterprise IT teams, Government bodies, EdTech platforms<br><br>

        <b>⏱️ Duration:</b>
        4 hrs (half-day) · 8 hrs (1 day) · 16 hrs · 24 hrs · 32 hrs · 40 hrs (1 week) · 80 hrs
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 📝 Learning Asset Configuration")

    # ── Row 1: Metadata ───────────────────────────────────────────────────────
    col1, col2 = st.columns(2)
    with col1:
        asset_name = st.text_input(
            "📌 Name of the Learning Asset *",
            placeholder="e.g.  Generative AI Foundations for Banking Professionals",
            value=st.session_state.form_data.get("asset_name", ""),
            help="The full title of the training programme to be developed.",
        )
        entity_name = st.text_input(
            "🏢 HEI / Enterprise / Entity Name *",
            placeholder="e.g.  Mashreq Bank · Bank Muscat · Amity University",
            value=st.session_state.form_data.get("entity_name", ""),
            help="The organisation for which the asset is being developed.",
        )
        duration = st.number_input(
            "⏱️ Training Duration (Hours) *",
            min_value=1,
            max_value=200,
            step=1,
            value=int(st.session_state.form_data.get("duration", 16)),
            help="Total instructional hours.  Each 60-min PPTX deck covers one module.",
        )

    with col2:
        color_palette = st.text_area(
            "🎨 Custom Colour Palette (Optional)",
            placeholder=(
                "Leave blank to use HCLTech defaults: "
                "Purple #5F1EBE · Blue #3C91FF · Teal #00A4A6 · Midnight #00112B\n\n"
                "Or paste your brand guidelines, e.g.:\n"
                "Primary: #FF5733  Secondary: #2ECC71  Background: #FFFFFF"
            ),
            height=130,
            value=st.session_state.form_data.get("color_palette", ""),
            help="Custom brand colours override HCLTech defaults in all generated assets.",
        )
        uploaded_doc = st.file_uploader(
            "📄 Upload Curriculum / Programme Document (PDF, DOCX, TXT)",
            type=["pdf", "docx", "txt"],
            help=(
                "Optional: upload a university guidesheet, syllabus, or programme outline. "
                "Content is passed as additional context to Agent 1."
            ),
        )

    # ── Row 2: Content ────────────────────────────────────────────────────────
    st.markdown("#### 📋 Content & Coverage Details")
    col3, col4 = st.columns(2)

    with col3:
        content_desc = st.text_area(
            "🗂️ Content Description & Subject Coverage *",
            placeholder=(
                "Describe subject matter, domains, technologies, frameworks, and concepts.\n\n"
                "Example:\n"
                "This programme covers Generative AI fundamentals including Large Language Models "
                "(LLMs), Transformer Architecture, Prompt Engineering (Zero-shot, Few-shot, "
                "Chain-of-Thought), RAG, AI Ethics & Governance, and practical implementation "
                "using Python and the Anthropic Claude API."
            ),
            height=185,
            value=st.session_state.form_data.get("content_desc", ""),
        )

    with col4:
        topics_subtopics = st.text_area(
            "📂 Modules, Topics & Sub-Topics *",
            placeholder=(
                "List modules, topics, and sub-topics with indicative durations.\n\n"
                "Example:\n"
                "Module 1: AI Foundations (3 hrs)\n"
                "  1.1 History and evolution of AI\n"
                "  1.2 Types of ML models\n"
                "Module 2: LLM Architecture (4 hrs)\n"
                "  2.1 Transformer & Attention\n"
                "  2.2 GPT / Claude families\n"
                "Module 3: Prompt Engineering (4 hrs)\n"
                "  3.1 Zero/Few-shot prompting\n"
                "  3.2 Chain-of-Thought"
            ),
            height=185,
            value=st.session_state.form_data.get("topics_subtopics", ""),
        )

    target_audience = st.text_area(
        "👥 Target Audience & Pre-requisites",
        placeholder=(
            "e.g.  Mid-level banking professionals with basic IT literacy. "
            "No prior AI experience required. Basic Python knowledge preferred."
        ),
        height=80,
        value=st.session_state.form_data.get("target_audience", ""),
    )

    # ── Start button ──────────────────────────────────────────────────────────
    st.markdown("---")
    _, col_start, _ = st.columns([1, 2, 1])
    with col_start:
        start_clicked = st.button(
            "🚀  START AGENTIC ORCHESTRATION",
            use_container_width=True,
            type="primary",
        )

    if start_clicked:
        # ── Input validation ─────────────────────────────────────────────────
        errors: List[str] = []
        if not asset_name.strip():
            errors.append("Asset Name is required.")
        if not entity_name.strip():
            errors.append("Entity / HEI Name is required.")
        if not content_desc.strip():
            errors.append("Content Description is required.")
        provider = st.session_state.get("ai_provider", "anthropic")
        pinfo    = AI_PROVIDERS[provider]
        if not st.session_state.api_key or st.session_state.api_key_valid is not True:
            errors.append(
                f"A validated {pinfo['label']} API key is required. "
                f"Go to Admin Module → Step 1 (choose provider) → Step 2 (validate key)."
            )

        if errors:
            for err in errors:
                st.error(f"❌ {err}")
        else:
            # ── Persist form data to session state ───────────────────────────
            st.session_state.form_data = {
                "asset_name":       asset_name.strip(),
                "entity_name":      entity_name.strip(),
                "duration":         int(duration),
                "color_palette":    color_palette.strip(),
                "content_desc":     content_desc.strip(),
                "topics_subtopics": topics_subtopics.strip(),
                "target_audience":  target_audience.strip(),
            }
            st.session_state.job_name = asset_name.strip()
            st.session_state.color_palette = (
                color_palette.strip()
                or "HCLTech: Purple #5F1EBE, Blue #3C91FF, Teal #00A4A6, Midnight #00112B"
            )

            # Store uploaded document bytes (read once; stored in session).
            if uploaded_doc is not None:
                st.session_state.uploaded_doc = uploaded_doc.read()
                logger.info(
                    "Document uploaded: %s (%d bytes)",
                    uploaded_doc.name, len(st.session_state.uploaded_doc),
                )

            # Reset agent pipeline state for a fresh run.
            st.session_state.current_agent  = 0
            st.session_state.agent_status   = ["idle"] * NUM_AGENTS
            st.session_state.agent_progress = [0] * NUM_AGENTS
            st.session_state.current_page   = "workspace"

            logger.info(
                "New job started: '%s' for '%s' (%d hrs)",
                asset_name, entity_name, duration,
            )
            st.rerun()


# =============================================================================
# SECTION 9 — AGENT 1: GUIDE-SHEET GENERATOR
# =============================================================================

def run_agent1() -> None:
    """
    Agent 1 — Guide-Sheet Generator.

    Orchestration flow:
      1. Send a structured 8-section instructional design prompt to Claude.
      2. Parse the response to extract each [SECTION N: …] block.
      3. Call generate_guidesheet_excel() to produce a formatted .xlsx workbook.
      4. Persist raw text and Excel bytes to session state.
      5. Render preview, download button, review gate, and next-agent control.

    Output artefacts:
      st.session_state.guidesheet_txt  — raw Claude text response
      st.session_state.guidesheet_xl   — serialised Excel workbook bytes
    """
    st.markdown("### 📋 Agent 1: Guide-Sheet Generator")
    st.markdown("""
    <div class="card">
    Generates a comprehensive, professionally formatted training guide-sheet covering all
    8 sections: Programme Overview, Learning Objectives, Outcomes, Pre-requisites,
    Competency Framework, Detailed Coverage, Assessment Rubric, and SW Licences.
    Output is a fully formatted Excel workbook (.xlsx).
    </div>
    """, unsafe_allow_html=True)

    ctx = build_context_string()
    progress_bar = st.progress(
        100 if st.session_state.agent_status[0] == "done" else 0
    )
    status_area = st.empty()

    # ── Run button (shown only when not yet completed) ─────────────────────────
    if st.session_state.agent_status[0] != "done":
        if st.button("▶️ Run Guide-Sheet Generator", key="run_a1"):
            st.session_state.agent_status[0]  = "running"
            st.session_state.current_agent    = 0
            progress_bar.progress(10)
            status_area.info("🔄 Initialising guide-sheet generation…")

            # ── System prompt: role and output format contract ─────────────────
            system_prompt = (
                "You are a senior instructional designer and curriculum architect "
                "specialising in enterprise technology training programmes. "
                "Generate a comprehensive, detailed training guide-sheet. "
                "Structure your output with EXACTLY these markers at the start of each section: "
                "[SECTION 1: …], [SECTION 2: …], … [SECTION 8: …]. "
                "Be verbose, specific, and actionable. Use hierarchical numbering. "
                "Include realistic durations, measurable objectives (Bloom's verbs), "
                "meaningful competencies, and practical assessment blueprints."
            )

            # ── User prompt: detailed 8-section specification ──────────────────
            duration_val = st.session_state.form_data.get("duration", 16)
            full_prompt = f"""Create a complete, professional training guide-sheet for the following programme:

{ctx}

Generate ALL 8 sections in full detail with rich, specific content:

[SECTION 1: PROGRAMME OVERVIEW]
Write a comprehensive 200–300 word overview covering: purpose, scope, strategic alignment
to the entity's goals, target industry context, and expected business impact.

[SECTION 2: LEARNING OBJECTIVES]
Write exactly 5 numbered, measurable learning objectives using Bloom's taxonomy verbs
(analyse, design, implement, evaluate, create). Each objective must be specific,
measurable, achievable, relevant, and time-bound (SMART format).

[SECTION 3: LEARNING OUTCOMES]
A) UNDERSTAND — numbered list of 5 cognitive outcomes (what learners will know)
B) ABLE TO DO  — numbered list of 5 skill outcomes (what learners will be able to perform)

[SECTION 4: PRE-REQUISITE KNOWLEDGE & SKILLS]
For each prerequisite item:
- Statement of the knowledge or skill required
- Proficiency level: Basic / Intermediate / Advanced
- Weightage % for pre-qualifier assessment blueprint
Ensure all weightages sum to exactly 100%.

[SECTION 5: COMPETENCY FRAMEWORK]
5a) DOCUMENTATION COMPETENCIES
    List critical documents, standards, and references learners must be able to work with.
5b) TECHNOLOGY, TOOLS & PROCESS COMPETENCIES
    List tools, platforms, frameworks, and methodologies covered.
5c) POLICIES, STANDARDS & BEST PRACTICES
    List relevant industry standards, compliance requirements, and governance frameworks.
5d) EVENTS & CEREMONIES
    List relevant agile ceremonies, workshops, or collaborative events (where applicable).

[SECTION 6: DETAILED TRAINING COVERAGE]
For each module provide:
  Sr. No | Module Name | Module Objective
  Hierarchical topics with exact durations (15 min / 30 min / 45 min / 60 min each)
  Sub-topics with brief descriptions
  Competency alignment (reference Section 5)
  Recommended pedagogy (lecture, lab, case study, simulation, discussion, assessment)
  Module total duration
Grand total must equal exactly {duration_val} hours.

[SECTION 7: POST-TRAINING ASSESSMENT RUBRIC]
Provide a complete assessment blueprint:
- Assessment types: MCQ knowledge check, practical exercise, case study analysis
- Coverage mapping: each assessment item mapped to a specific learning outcome
- Weightage per competency domain (sum to 100%)
- Passing threshold and remediation pathway
- Effectiveness KPIs (knowledge gain %, skill demonstration %, confidence score)

[SECTION 8: SOFTWARE LICENCES & SUBSCRIPTIONS]
Provide a numbered list of every tool, platform, or software licence required:
- Licence/Tool name
- Purpose within the training programme
- Cost tier: Free / Freemium / Paid (with indicative cost if known)
- Minimum version or tier required"""

            progress_bar.progress(25)
            status_area.info("🤖 Claude is generating the 8-section guide-sheet…")

            result = call_claude(full_prompt, system_prompt, agent_idx=0)

            progress_bar.progress(75)
            status_area.info("📊 Generating formatted Excel workbook…")

            # Persist raw text.
            st.session_state.guidesheet_txt = result

            # Generate Excel workbook.
            try:
                excel_bytes = generate_guidesheet_excel(result)
                st.session_state.guidesheet_xl = excel_bytes
                logger.info("Guide-sheet Excel generated (%d bytes).", len(excel_bytes))
            except Exception as exc:
                st.session_state.guidesheet_xl = None
                logger.exception("Excel generation failed: %s", exc)
                st.warning(f"⚠️ Excel generation encountered an issue: {exc}")

            # Mark complete.
            st.session_state.agent_status[0]   = "done"
            st.session_state.agent_progress[0] = 100
            progress_bar.progress(100)
            status_area.success("✅ Guide-Sheet generated successfully!")
            st.rerun()

    # ── Results panel (shown when complete) ────────────────────────────────────
    if st.session_state.agent_status[0] == "done":
        progress_bar.progress(100)
        status_area.success("✅ Guide-Sheet generation complete!")

        # Preview first 5,000 characters.
        preview_text = st.session_state.guidesheet_txt
        truncated    = len(preview_text) > 5_000
        preview_suffix = "\n\n… [Full content in Excel download]" if truncated else ""
        with st.expander("📖 View Guide-Sheet Content Preview", expanded=True):
            st.markdown(f"""
            <div class="preview-block">
{preview_text[:5_000]}{preview_suffix}
            </div>
            """, unsafe_allow_html=True)

        # Excel download button.
        if st.session_state.guidesheet_xl:
            safe_name = sanitise_filename(
                st.session_state.form_data.get("asset_name", "Asset")
            )
            st.download_button(
                "⬇️ Download Guide-Sheet (Excel .xlsx)",
                data=st.session_state.guidesheet_xl,
                file_name=f"GuideSheet_{safe_name}.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument"
                    ".spreadsheetml.sheet"
                ),
            )

        render_review_gate(0)
        col_rerun, col_next = st.columns(2)
        with col_rerun:
            if render_rerun_button(0):
                st.rerun()
        with col_next:
            if render_next_agent_button(0):
                st.session_state.current_agent = 1
                st.rerun()


def generate_guidesheet_excel(content: str) -> bytes:
    """
    Convert Claude's guide-sheet text into a professionally formatted Excel workbook.

    Layout rules:
      - Column A : Always blank — no gridlines, no content (visual left margin).
      - Column B : Section title rows — dark navy fill (#001F5B), white bold font.
                   Content label rows — mid-blue fill (#C5D5F5), dark bold font.
      - Column C : Content — light blue fill (#E8F0FE), word-wrapped, left-aligned.
      - Section boundaries : dark navy outer border.
      - Content grid lines  : blue inner borders (#3C91FF).
      - Spacer rows        : blank unformatted rows between sections.

    Parsing strategy:
      Sections are split on the regex marker [SECTION N: …].
      Within each section, double-newline blocks are written as label/content pairs
      where the first line (< 60 chars, not a list item) becomes the label.

    Args:
        content: Raw text response from Claude containing [SECTION N: …] markers.

    Returns:
        Serialised bytes of the Excel workbook (.xlsx).

    Raises:
        Any openpyxl exception is allowed to propagate; the caller catches it.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Training Guide-Sheet"

    # ── Pre-built style objects (created once, applied many times) ─────────────
    dark_fill   = PatternFill("solid", fgColor=XL_DARK_BLUE)
    light_fill  = PatternFill("solid", fgColor=XL_LIGHT_BLUE)
    mid_fill    = PatternFill("solid", fgColor=XL_MID_BLUE)
    accent_fill = PatternFill("solid", fgColor=XL_ACCENT)
    white_fill  = PatternFill("solid", fgColor=XL_WHITE)

    thin_blue = Border(
        left   = Side(style="thin",   color=XL_GRID_BLUE),
        right  = Side(style="thin",   color=XL_GRID_BLUE),
        top    = Side(style="thin",   color=XL_GRID_BLUE),
        bottom = Side(style="thin",   color=XL_GRID_BLUE),
    )
    dark_border = Border(
        left   = Side(style="medium", color=XL_DARK_BLUE),
        right  = Side(style="medium", color=XL_DARK_BLUE),
        top    = Side(style="medium", color=XL_DARK_BLUE),
        bottom = Side(style="medium", color=XL_DARK_BLUE),
    )

    # ── Column widths ─────────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 3    # Blank margin
    ws.column_dimensions["B"].width = 30   # Labels / section titles
    ws.column_dimensions["C"].width = 92   # Primary content
    for col_letter in ("D", "E", "F"):
        ws.column_dimensions[col_letter].width = 24

    # ── Mutable row counter (closure variable) ─────────────────────────────────
    current_row = [1]  # Use a list so nested functions can mutate it.

    def _next_row() -> int:
        """Increment and return the current row index."""
        r = current_row[0]
        current_row[0] += 1
        return r

    def _spacer(count: int = 1) -> None:
        """Insert `count` blank spacer rows with minimal height."""
        for _ in range(count):
            ws.row_dimensions[current_row[0]].height = 7
            _next_row()

    def _write_section_header(title: str) -> None:
        """Write a full-width dark navy section header row spanning B:F."""
        _spacer()
        r = current_row[0]
        ws.merge_cells(f"B{r}:F{r}")
        cell = ws[f"B{r}"]
        cell.value     = title.upper()
        cell.fill      = dark_fill
        cell.font      = Font(name="Calibri", bold=True, size=12, color=XL_WHITE)
        cell.alignment = Alignment(
            horizontal="left", vertical="center", wrap_text=True, indent=1
        )
        cell.border    = dark_border
        ws.row_dimensions[r].height = 30
        _next_row()

    def _write_content_row(label: str, value: str) -> None:
        """
        Write a single content row:
          Col A — blank with white fill
          Col B — label with mid-blue fill
          Col C — content with light-blue fill
        Row height is auto-sized based on content length (capped at 120 pt).
        """
        r = _next_row()
        # Column A — always blank.
        ws[f"A{r}"].fill = white_fill

        # Column B — label.
        b = ws[f"B{r}"]
        b.value     = label
        b.font      = Font(name="Calibri", bold=True, size=10, color=XL_DARK_BLUE)
        b.fill      = mid_fill
        b.alignment = Alignment(
            horizontal="left", vertical="top", wrap_text=True, indent=1
        )
        b.border = thin_blue

        # Column C — content.
        c = ws[f"C{r}"]
        c.value     = str(value)
        c.font      = Font(name="Calibri", size=10, color="1A1A2E")
        c.fill      = light_fill
        c.alignment = Alignment(
            horizontal="left", vertical="top", wrap_text=True, indent=1
        )
        c.border = thin_blue

        # Auto row height: proportional to content length, capped.
        ws.row_dimensions[r].height = max(
            28, min(120, len(str(value)) // 3)
        )

    # ── Master title row ──────────────────────────────────────────────────────
    r = current_row[0]
    ws.merge_cells(f"B{r}:F{r}")
    title_cell = ws[f"B{r}"]
    asset_name = st.session_state.form_data.get("asset_name", "Training Programme")
    title_cell.value     = f"TRAINING GUIDE-SHEET  —  {asset_name.upper()}"
    title_cell.fill      = accent_fill
    title_cell.font      = Font(name="Calibri", bold=True, size=15, color=XL_WHITE)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.border    = dark_border
    ws.row_dimensions[r].height = 40
    _next_row()

    # ── Programme metadata rows ────────────────────────────────────────────────
    _write_content_row(
        "Organisation", st.session_state.form_data.get("entity_name", "")
    )
    _write_content_row(
        "Duration", f"{st.session_state.form_data.get('duration', '')} Hours"
    )
    _write_content_row(
        "Target Audience", st.session_state.form_data.get("target_audience", "")
    )
    _write_content_row(
        "Generated On",
        datetime.datetime.now().strftime("%d %B %Y  %H:%M")
    )

    # ── Parse and write each section ──────────────────────────────────────────
    # Split the Claude response on SECTION markers.
    raw_sections = re.split(r'\[SECTION\s+\d+\s*:', content, flags=re.IGNORECASE)

    for sec_idx, raw_sec in enumerate(raw_sections[1:], start=1):
        # Extract title (everything before the closing ']').
        sec_title_match = re.match(r'([^\]]+)\]', raw_sec)
        sec_title = sec_title_match.group(1).strip() if sec_title_match else f"Section {sec_idx}"
        body      = raw_sec[sec_title_match.end():].strip() if sec_title_match else raw_sec.strip()

        _write_section_header(f"Section {sec_idx}: {sec_title}")

        # Split body into logical paragraphs (double newline separator).
        paragraphs = [p.strip() for p in body.split("\n\n") if p.strip()]

        for para in paragraphs:
            lines = para.split("\n")
            first_line = lines[0].strip()

            # Heuristic: if first line is short and not a list item → label.
            is_list_item = first_line.startswith(
                ("A)", "B)", "C)", "D)", "E)", "-", "•", "*", "1.", "2.",
                 "3.", "4.", "5.", "6.", "7.", "8.", "9.", "10.")
            )
            if not is_list_item and len(first_line) < 70 and len(lines) > 1:
                _write_content_row(first_line, "\n".join(lines[1:]).strip())
            else:
                _write_content_row("", para)

        _spacer(2)

    # ── Serialise to bytes ────────────────────────────────────────────────────
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.read()


# =============================================================================
# SECTION 10 — AGENT 2: POWERPOINT CREATOR
# =============================================================================

def run_agent2() -> None:
    """
    Agent 2 — PowerPoint Presentation Creator.

    Orchestration flow:
      1. Send a structured slide-spec prompt to Claude → 20–25 slide scripts.
      2. Parse SLIDE [N] | TYPE / TITLE / CONTENT / VOICEOVER / IMAGE_PROMPT / ACTIVITY blocks.
      3. Call generate_pptx() to build a fully branded .pptx deck.
      4. Persist pptx_script (text) and pptx_files (bytes) to session state.
      5. Render slide script preview, download button, review gate, re-run, next-agent.

    Pre-condition: Agent 1 must be complete (agent_status[0] == "done").

    Output artefacts:
      st.session_state.pptx_script  — structured slide specification text
      st.session_state.pptx_files   — list[{"name": str, "data": bytes}]
    """
    st.markdown("### 📊 Agent 2: PowerPoint Presentation Creator")
    st.markdown("""
    <div class="card">
    Creates a fully branded, business-professional PowerPoint deck aligned to the guide-sheet.
    Includes title slide, learning objectives (SmartArt), agenda, content slides, knowledge
    activities (MCQ + fill-in-blank), recap, next-module teaser, and per-slide voice-over scripts
    in the Presenter Notes section.
    </div>
    """, unsafe_allow_html=True)

    # ── Dependency check ──────────────────────────────────────────────────────
    if st.session_state.agent_status[0] != "done":
        st.warning("⚠️ Please complete **Agent 1 — Guide-Sheet Generator** first.")
        return

    ctx = build_context_string()
    progress_bar = st.progress(
        100 if st.session_state.agent_status[1] == "done" else 0
    )
    status_area = st.empty()

    # ── Run button ────────────────────────────────────────────────────────────
    if st.session_state.agent_status[1] != "done":
        if st.button("▶️ Run PowerPoint Creator Agent", key="run_a2"):
            st.session_state.agent_status[1] = "running"
            st.session_state.current_agent   = 1
            progress_bar.progress(8)
            status_area.info("🔄 Analysing guide-sheet and planning slide structure…")

            # ── System prompt ──────────────────────────────────────────────────
            system_prompt = (
                "You are a senior instructional designer and presentation architect. "
                "Create a detailed, structured PowerPoint slide specification. "
                "Each slide MUST follow this exact format on consecutive lines:\n"
                "SLIDE [N] | TYPE: [Title|Objective|Agenda|Content|Example|Activity|ActivityExplain|Recap|NextModule]\n"
                "TITLE: [slide title — max 10 words]\n"
                "CONTENT: [bullet points, one per line, max 5 bullets, each ≤ 15 words]\n"
                "IMAGE_PROMPT: [one-sentence image description for AI generation]\n"
                "VOICEOVER: [natural presenter script — max 35 words]\n"
                "ACTIVITY: [only for Activity slides: 2 MCQ questions (5 options each, "
                "mark correct with *) and 2 fill-in-blank sentences with [BLANK]]\n"
                "Leave a blank line between slides. Do not add any other text outside this format."
            )

            # ── User prompt ────────────────────────────────────────────────────
            guide_summary = st.session_state.guidesheet_txt[:3_500]
            content_focus = st.session_state.form_data.get("content_desc", "")
            audience      = st.session_state.form_data.get("target_audience", "")

            full_prompt = f"""Create a complete PowerPoint deck specification (20–25 slides) for:

{ctx}

Guide-sheet summary (use for content accuracy):
{guide_summary}

REQUIRED slide sequence:
1. Title slide (TYPE: Title) — programme title, entity, date tagline
2. Learning Objectives slide (TYPE: Objective) — 5 impactful SmartArt-style objectives
3. Agenda slide (TYPE: Agenda) — module-by-module coverage for 60 minutes
4. For EACH major topic/module:
   a. Explanation/Key Terms slide (TYPE: Content) — definitions, concepts, SmartArt
   b. Example/Scenario slide (TYPE: Example) — realistic scenario using FICTITIOUS
      business entity names (e.g. "AlphaCorp Bank", "NovaTech Solutions")
   c. Activity slide (TYPE: Activity) — 2 MCQ (5 options, mark correct with *)
      + 2 fill-in-blank sentences with [BLANK] placeholder
5. Activity Explanation slide (TYPE: ActivityExplain) — answers with slide references
6. Recap slide (TYPE: Recap) — 5 key takeaways with impact statements
7. Next Module teaser (TYPE: NextModule) — preview of coming content

Content focus: {content_focus}
Target audience: {audience}

Image prompts must describe professional, brand-appropriate visuals using
HCLTech colours (Purple #5F1EBE, Blue #3C91FF, Teal #00A4A6).

Voice-over scripts must be natural, warm, and suitable for a professional
Indian female TTS voice. Navigation phrasing for transitions."""

            progress_bar.progress(25)
            status_area.info("🤖 Claude is designing the complete slide deck…")

            slide_script = call_claude(full_prompt, system_prompt, agent_idx=1)

            progress_bar.progress(65)
            status_area.info("📊 Building branded PPTX file…")

            st.session_state.pptx_script = slide_script

            # Generate the PPTX workbook.
            try:
                pptx_bytes = generate_pptx(slide_script)
                safe_name  = sanitise_filename(
                    st.session_state.form_data.get("asset_name", "Deck")
                )
                st.session_state.pptx_files = [
                    {"name": f"Deck_1_{safe_name}.pptx", "data": pptx_bytes}
                ]
                logger.info("PPTX generated (%d bytes).", len(pptx_bytes))
            except Exception as exc:
                st.session_state.pptx_files = []
                logger.exception("PPTX generation failed: %s", exc)
                st.warning(f"⚠️ PPTX generation encountered an issue: {exc}")

            st.session_state.agent_status[1]   = "done"
            st.session_state.agent_progress[1] = 100
            progress_bar.progress(100)
            status_area.success("✅ PowerPoint deck generated!")
            st.rerun()

    # ── Results panel ─────────────────────────────────────────────────────────
    if st.session_state.agent_status[1] == "done":
        progress_bar.progress(100)
        status_area.success("✅ PowerPoint creation complete!")

        with st.expander("📖 View Slide Script (structured specification)", expanded=False):
            st.markdown(f"""
            <div class="preview-block">
{st.session_state.pptx_script[:4_500]}
            </div>""", unsafe_allow_html=True)

        for deck in st.session_state.pptx_files:
            st.download_button(
                f"⬇️ Download {deck['name']}",
                data=deck["data"],
                file_name=deck["name"],
                mime=(
                    "application/vnd.openxmlformats-officedocument"
                    ".presentationml.presentation"
                ),
            )

        render_review_gate(1)
        col_rerun, col_next = st.columns(2)
        with col_rerun:
            if render_rerun_button(1, "🔄 Re-run Agent 2 (update deck)"):
                st.rerun()
        with col_next:
            if render_next_agent_button(1):
                st.session_state.current_agent = 2
                st.rerun()


def generate_pptx(slide_script: str) -> bytes:
    """
    Build a fully branded PowerPoint presentation from a structured slide script.

    Slide types and their visual treatments:
      Title       — Dark midnight background, purple left bar, purple footer strip.
      Content     — White body, dark navy header bar, branded bullets, image placeholder.
      Activity    — Teal header, activity content box, "KNOWLEDGE CHECK" label.
      (All others treated as Content type.)

    Every slide includes:
      - HCLTech Career Shaper™ logo text (top-left header)
      - Slide number (top-right header, teal)
      - Confidentiality clause (dark navy footer)
      - Image placeholder box (right panel, labelled with IMAGE_PROMPT text)
      - Voice-over script in PowerPoint Presenter Notes

    Args:
        slide_script: Structured text from Agent 2 with SLIDE [N] | TYPE: … blocks.

    Returns:
        Serialised bytes of the .pptx file.
    """
    prs = Presentation()
    prs.slide_width  = Inches(13.33)   # Widescreen 16:9
    prs.slide_height = Inches(7.5)

    blank_layout = prs.slide_layouts[6]  # Layout 6 = completely blank

    # ─────────────────────────────────────────────────────────────────────────
    # INNER HELPER: add_rect
    # ─────────────────────────────────────────────────────────────────────────
    def add_rect(
        slide,
        left: float, top: float,
        width: float, height: float,
        fill: Optional[RGBColor] = None,
        line_color: Optional[RGBColor] = None,
        line_pt: Optional[float] = None,
    ):
        """
        Add a rectangle shape to a slide.

        Args:
            slide      : pptx Slide object.
            left/top   : Position in EMUs (use Inches() to convert).
            width/height: Dimensions in EMUs.
            fill       : Optional RGBColor fill.  None → transparent.
            line_color : Optional RGBColor border.  None → no border.
            line_pt    : Border width in points (only used if line_color is set).
        """
        shape = slide.shapes.add_shape(1, left, top, width, height)
        shape.line.fill.background()   # Start with no line.
        if fill:
            shape.fill.solid()
            shape.fill.fore_color.rgb = fill
        else:
            shape.fill.background()
        if line_color:
            shape.line.color.rgb = line_color
            if line_pt:
                shape.line.width = Pt(line_pt)
        return shape

    # ─────────────────────────────────────────────────────────────────────────
    # INNER HELPER: add_textbox
    # ─────────────────────────────────────────────────────────────────────────
    def add_textbox(
        slide,
        text: str,
        left: float, top: float,
        width: float, height: float,
        font_size: float = 14,
        bold: bool = False,
        color: Optional[RGBColor] = None,
        align: int = PP_ALIGN.LEFT,
        font_name: str = "Calibri",
    ):
        """
        Add a text box to a slide with a single styled text run.

        Args:
            slide      : pptx Slide object.
            text       : Content string (multi-line supported via \\n in PPTX).
            left/top   : Position in EMUs.
            width/height: Dimensions in EMUs.
            font_size  : Point size.
            bold       : Bold flag.
            color      : RGBColor for the text.
            align      : PP_ALIGN constant.
            font_name  : Font family name.
        """
        txb = slide.shapes.add_textbox(left, top, width, height)
        tf  = txb.text_frame
        tf.word_wrap = True
        p   = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text           = text
        run.font.size      = Pt(font_size)
        run.font.bold      = bold
        run.font.name      = font_name
        if color:
            run.font.color.rgb = color
        return txb

    # ─────────────────────────────────────────────────────────────────────────
    # SLIDE BUILDER: Title slide
    # ─────────────────────────────────────────────────────────────────────────
    def make_title_slide(title: str, subtitle: str, entity: str) -> None:
        """
        Build the opening title slide with a dark midnight background.

        Layout:
          - Midnight background fill
          - Purple left accent bar (0.5 in wide)
          - Blue logo/brand box (top-left)
          - Large white title text (36pt)
          - Teal subtitle text (18pt)
          - Entity + date label (12pt, grey)
          - Purple footer strip with confidential text
        """
        slide = prs.slides.add_slide(blank_layout)

        # Background + accents.
        add_rect(slide, 0, 0, prs.slide_width, prs.slide_height, fill=C_DARK)
        add_rect(slide, 0, 0, Inches(0.5), prs.slide_height, fill=C_PURPLE)
        add_rect(slide, 0, Inches(6.8), prs.slide_width, Inches(0.7), fill=C_PURPLE)

        # Brand box.
        add_rect(
            slide, Inches(0.8), Inches(0.4), Inches(3.2), Inches(0.7), fill=C_BLUE
        )
        add_textbox(
            slide, "HCLTech  |  Career Shaper™",
            Inches(0.82), Inches(0.4), Inches(3.1), Inches(0.7),
            font_size=11, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER,
        )

        # Title, subtitle, entity.
        add_textbox(
            slide, title,
            Inches(0.8), Inches(1.8), Inches(10), Inches(2.2),
            font_size=34, bold=True, color=C_WHITE,
        )
        add_textbox(
            slide, subtitle,
            Inches(0.8), Inches(4.05), Inches(10), Inches(0.8),
            font_size=18, color=C_TEAL,
        )
        add_textbox(
            slide,
            f"{entity}   |   {datetime.datetime.now().strftime('%B %Y')}",
            Inches(0.8), Inches(4.95), Inches(10), Inches(0.5),
            font_size=12, color=C_GREY,
        )

        # Footer.
        add_textbox(
            slide, "CONFIDENTIAL  —  For Internal Use Only",
            Inches(5.0), Inches(6.82), Inches(7.0), Inches(0.4),
            font_size=9, color=C_WHITE, align=PP_ALIGN.RIGHT,
        )

        # Presenter notes.
        notes_tf = slide.notes_slide.notes_text_frame
        notes_tf.text = (
            f"Welcome to {title}. "
            f"This training programme has been specially designed for {entity}. "
            "Let us begin by looking at what we will cover today."
        )

    # ─────────────────────────────────────────────────────────────────────────
    # SLIDE BUILDER: Content slide
    # ─────────────────────────────────────────────────────────────────────────
    def make_content_slide(
        slide_num: int,
        title: str,
        bullets: str,
        voiceover: str = "",
        img_prompt: str = "",
    ) -> None:
        """
        Build a standard content slide with a white body and branded header.

        Layout:
          - White body background
          - Dark navy top bar (1.15 in)
          - Purple accent line below top bar
          - HCLTech brand text (top-left, white)
          - Slide number (top-right, teal)
          - Slide title in header bar (white bold)
          - Bullet content (left column, dark, ▸ prefix)
          - Image placeholder box (right column, purple border)
          - Dark navy footer with confidential text
          - Voiceover in Presenter Notes
        """
        slide = prs.slides.add_slide(blank_layout)

        # Background layers.
        add_rect(slide, 0, 0, prs.slide_width, prs.slide_height, fill=C_WHITE)
        add_rect(slide, 0, 0, prs.slide_width, Inches(1.15), fill=C_DARK)
        add_rect(slide, 0, Inches(1.15), prs.slide_width, Inches(0.07), fill=C_PURPLE)
        add_rect(slide, 0, Inches(7.1), prs.slide_width, Inches(0.4), fill=C_DARK)

        # Header content.
        add_textbox(
            slide, "HCLTech  Career Shaper™",
            Inches(0.2), Inches(0.16), Inches(3.5), Inches(0.5),
            font_size=11, bold=True, color=C_WHITE,
        )
        add_textbox(
            slide, f"{slide_num:02d}",
            Inches(12.0), Inches(0.18), Inches(1.1), Inches(0.5),
            font_size=22, bold=True, color=C_TEAL, align=PP_ALIGN.RIGHT,
        )
        add_textbox(
            slide, title,
            Inches(0.35), Inches(0.22), Inches(11.3), Inches(0.72),
            font_size=22, bold=True, color=C_WHITE,
        )

        # Determine content column width based on whether an image is present.
        has_image    = bool(img_prompt.strip())
        content_width = Inches(7.6) if has_image else Inches(12.5)

        # Bullet content.
        if bullets.strip():
            body_box = slide.shapes.add_textbox(
                Inches(0.4), Inches(1.38), content_width, Inches(5.4)
            )
            tf = body_box.text_frame
            tf.word_wrap = True

            bullet_lines = [
                ln.strip() for ln in bullets.split("\n") if ln.strip()
            ]
            for i, line in enumerate(bullet_lines[:7]):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.alignment  = PP_ALIGN.LEFT
                p.space_before = Pt(4)
                run = p.add_run()
                # Strip leading list characters before adding the ▸ prefix.
                clean = re.sub(r'^[•\-\*▸]\s*', '', line)
                run.text           = f"▸   {clean}"
                run.font.size      = Pt(14)
                run.font.name      = "Calibri"
                run.font.color.rgb = C_DARK

        # Image placeholder box.
        if has_image:
            add_rect(
                slide,
                Inches(8.35), Inches(1.32), Inches(4.6), Inches(5.45),
                fill=C_LPUR,
                line_color=C_PURPLE, line_pt=1.5,
            )
            add_textbox(
                slide,
                f"🖼️  IMAGE PLACEHOLDER\n\n{img_prompt[:90]}",
                Inches(8.45), Inches(1.5), Inches(4.35), Inches(5.2),
                font_size=10, color=C_PURPLE, align=PP_ALIGN.CENTER,
            )

        # Footer.
        add_textbox(
            slide, "CONFIDENTIAL  —  For Internal Use Only",
            Inches(4.0), Inches(7.12), Inches(5.0), Inches(0.3),
            font_size=8, color=C_WHITE, align=PP_ALIGN.CENTER,
        )

        # Presenter notes (voice-over script).
        if voiceover:
            slide.notes_slide.notes_text_frame.text = voiceover

    # ─────────────────────────────────────────────────────────────────────────
    # SLIDE BUILDER: Activity slide
    # ─────────────────────────────────────────────────────────────────────────
    def make_activity_slide(
        slide_num: int,
        title: str,
        activity_content: str,
    ) -> None:
        """
        Build an interactive activity slide with a teal header.

        The activity content is expected to contain:
          - 2 MCQ questions (5 options, correct marked with *)
          - 2 fill-in-blank sentences (blanks shown as [BLANK])

        Presenter Notes include instructions for the facilitator.
        """
        slide = prs.slides.add_slide(blank_layout)

        # Background layers.
        add_rect(slide, 0, 0, prs.slide_width, prs.slide_height, fill=C_WHITE)
        add_rect(slide, 0, 0, prs.slide_width, Inches(1.15), fill=C_TEAL)
        add_rect(slide, 0, Inches(1.15), prs.slide_width, Inches(0.06), fill=C_BLUE)
        add_rect(slide, 0, Inches(7.1), prs.slide_width, Inches(0.4), fill=C_DARK)

        # Header.
        add_textbox(
            slide, "HCLTech  Career Shaper™",
            Inches(0.2), Inches(0.16), Inches(3.5), Inches(0.5),
            font_size=11, bold=True, color=C_WHITE,
        )
        add_textbox(
            slide, f"{slide_num:02d}",
            Inches(12.0), Inches(0.18), Inches(1.1), Inches(0.5),
            font_size=22, bold=True, color=C_WHITE, align=PP_ALIGN.RIGHT,
        )
        add_textbox(
            slide, f"⚡  {title}",
            Inches(0.35), Inches(0.22), Inches(11.3), Inches(0.72),
            font_size=22, bold=True, color=C_WHITE,
        )

        # "KNOWLEDGE CHECK" label.
        add_textbox(
            slide, "⬤  KNOWLEDGE CHECK  —  Test your understanding",
            Inches(0.4), Inches(1.3), Inches(8.0), Inches(0.38),
            font_size=10, bold=True, color=C_TEAL,
        )

        # Activity content (MCQ + fill-blank).
        add_textbox(
            slide, activity_content[:700],
            Inches(0.4), Inches(1.72), Inches(12.5), Inches(5.1),
            font_size=12, color=C_DARK,
        )

        # Footer.
        add_textbox(
            slide, "CONFIDENTIAL  —  For Internal Use Only",
            Inches(4.0), Inches(7.12), Inches(5.0), Inches(0.3),
            font_size=8, color=C_WHITE, align=PP_ALIGN.CENTER,
        )

        # Facilitator notes.
        slide.notes_slide.notes_text_frame.text = (
            "FACILITATOR NOTE: Pause here and allow learners 4–5 minutes to attempt "
            "all activities independently. For MCQ: only one option will give "
            "'Well Done' — others show 'Try Again'. "
            "For fill-in-blank: learners drag the correct word into the blank space. "
            "Review answers as a group before advancing."
        )

    # ─────────────────────────────────────────────────────────────────────────
    # PARSE THE SLIDE SCRIPT AND BUILD SLIDES
    # ─────────────────────────────────────────────────────────────────────────
    asset_name = st.session_state.form_data.get("asset_name", "Training Programme")
    entity     = st.session_state.form_data.get("entity_name", "")

    # Slide 1: title slide.
    make_title_slide(
        title    = asset_name,
        subtitle = f"A Comprehensive Training Programme for {entity}",
        entity   = entity,
    )

    # Parse remaining slides from the Claude script.
    # Split on the SLIDE [N] | marker pattern.
    slide_blocks = re.split(
        r'SLIDE\s*\[?\s*\d+\s*\]?\s*\|',
        slide_script,
        flags=re.IGNORECASE,
    )

    slide_num = 2
    for block in slide_blocks[1:]:   # skip text before first SLIDE marker
        if slide_num > 38:           # hard cap to prevent runaway decks
            break

        # Parse field tags from the block.
        def _extract(tag: str, text: str) -> str:
            """Extract a single-line field value by tag name."""
            m = re.search(
                rf'^{tag}:\s*(.+)',
                text,
                flags=re.MULTILINE | re.IGNORECASE,
            )
            return m.group(1).strip() if m else ""

        def _extract_multiline(tag: str, text: str) -> str:
            """
            Extract a multi-line field: starts at TAG: and ends at next uppercase TAG:.
            """
            pattern = rf'^{tag}:\s*(.*?)(?=\n[A-Z_]+:|$)'
            m = re.search(pattern, text, flags=re.MULTILINE | re.DOTALL | re.IGNORECASE)
            return m.group(1).strip() if m else ""

        s_type     = _extract("TYPE",         block)
        title      = _extract("TITLE",        block)
        voiceover  = _extract("VOICEOVER",    block)
        img_prompt = _extract("IMAGE_PROMPT", block)
        content    = _extract_multiline("CONTENT",  block)
        activity   = _extract_multiline("ACTIVITY", block)

        if not title:
            title = f"Slide {slide_num}"

        # Route to the correct slide builder.
        if "activity" in s_type.lower() or (activity and len(activity) > 20):
            make_activity_slide(
                slide_num,
                title,
                activity or content,
            )
        else:
            make_content_slide(
                slide_num,
                title,
                content,
                voiceover,
                img_prompt,
            )

        slide_num += 1

    # Always add a recap slide as the penultimate slide.
    make_content_slide(
        slide_num,
        "📌 Module Recap — Key Takeaways",
        (
            "▸   Core concepts and frameworks introduced this module\n"
            "▸   Learning objectives achieved and skills developed\n"
            "▸   Tools, technologies, and platforms practised\n"
            "▸   Real-world application scenarios explored\n"
            "▸   Assessment alignment to stated learning outcomes"
        ),
        voiceover=(
            "Excellent work reaching the end of this module. "
            "Let us quickly recap the key takeaways before your knowledge check."
        ),
        img_prompt="Professional team celebrating a milestone, HCLTech brand colours",
    )

    # ── Serialise to bytes ────────────────────────────────────────────────────
    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer.read()


# =============================================================================
# SECTION 11 — AGENT 3: IMAGE GENERATOR
# =============================================================================

def run_agent3() -> None:
    """
    Agent 3 — Image Generator.

    Orchestration flow:
      1. Extract all IMAGE_PROMPT tags from the stored pptx_script.
      2. Send prompts to Claude for enhancement into detailed visual briefs.
      3. Return structured JSON specifications per slide.
      4. Provide a downloadable spec sheet (.txt) that can be used with
         any image generation API (DALL-E 3, Stable Diffusion, Firefly).

    Note on actual image rendering:
      Streamlit cannot embed binary images into PPTX in-session without
      a third-party image generation API.  The connector code stub is
      provided in the output spec for developers to plug in their API key.

    Pre-condition: Agent 2 must be complete (agent_status[1] == "done").

    Output artefact:
      st.session_state.image_specs  — structured image spec text/JSON
    """
    st.markdown("### 🖼️ Agent 3: Image Generator")
    st.markdown("""
    <div class="card">
    Parses the PowerPoint deck slide-by-slide, extracts every IMAGE_PROMPT tag,
    and generates detailed visual briefs — brand-aligned, contextually relevant,
    and ready to submit to DALL-E 3, Stable Diffusion, or Adobe Firefly.
    Download the image specification sheet to complete the image embedding step.
    </div>
    """, unsafe_allow_html=True)

    # ── Dependency check ──────────────────────────────────────────────────────
    if st.session_state.agent_status[1] != "done":
        st.warning("⚠️ Please complete **Agent 2 — PowerPoint Creator** first.")
        return
    if not st.session_state.pptx_script:
        st.warning("⚠️ No slide script found. Please re-run Agent 2.")
        return

    progress_bar = st.progress(
        100 if st.session_state.agent_status[2] == "done" else 0
    )
    status_area = st.empty()

    # ── Run button ────────────────────────────────────────────────────────────
    if st.session_state.agent_status[2] != "done":
        if st.button("▶️ Run Image Generator Agent", key="run_a3"):
            st.session_state.agent_status[2] = "running"
            st.session_state.current_agent   = 2
            progress_bar.progress(15)
            status_area.info("🔄 Extracting image prompts from slide script…")

            # Extract all IMAGE_PROMPT values.
            raw_prompts: List[str] = re.findall(
                r'IMAGE_PROMPT:\s*(.+)', st.session_state.pptx_script
            )

            progress_bar.progress(30)
            status_area.info("🤖 Generating enhanced image specifications…")

            system_prompt = (
                "You are a professional visual content designer for corporate training materials. "
                "For each image prompt, produce an enhanced, detailed visual brief. "
                "Output ONLY valid JSON — an array of objects, no preamble, no markdown. "
                "Each object must have exactly these keys: "
                "slide_number (int), original_prompt (str), enhanced_description (str), "
                "style (str: 'photorealistic'|'flat-illustration'|'infographic'|'3d-render'), "
                "composition (str: brief layout description), "
                "brand_colors (str: hex codes to use), "
                "negative_prompt (str: elements to avoid), "
                "suggested_api_params (object: width, height, quality)."
            )

            numbered_prompts = "\n".join(
                f"{i + 1}. {p}" for i, p in enumerate(raw_prompts[:20])
            )
            asset_name = st.session_state.form_data.get("asset_name", "")
            audience   = st.session_state.form_data.get("target_audience", "")

            full_prompt = f"""Training asset: {asset_name}
Target audience: {audience}
Brand palette: Purple #5F1EBE · Blue #3C91FF · Teal #00A4A6 · Midnight #00112B

Enhance each image prompt below into a detailed visual brief for AI image generation:

{numbered_prompts}

Return a JSON array.  Ensure no placeholder text, no brand logos, no copyrighted elements."""

            image_specs_raw = call_claude(full_prompt, system_prompt, agent_idx=2)

            # Build the downloadable specification sheet.
            image_spec_doc = _build_image_spec_doc(image_specs_raw, raw_prompts)

            st.session_state.image_specs     = image_specs_raw
            st.session_state.image_spec_doc  = image_spec_doc

            st.session_state.agent_status[2]   = "done"
            st.session_state.agent_progress[2] = 100
            progress_bar.progress(100)
            status_area.success(
                "✅ Image specifications generated! "
                "Download the spec sheet and submit prompts to your image generation API."
            )
            st.rerun()

    # ── Results panel ─────────────────────────────────────────────────────────
    if st.session_state.agent_status[2] == "done":
        progress_bar.progress(100)
        status_area.success("✅ Image generation specifications complete!")

        st.info(
            "ℹ️ Connect to **DALL-E 3** (`openai` library), **Stable Diffusion** "
            "(`diffusers`), or **Adobe Firefly** to render actual images and embed "
            "them into the PPTX deck automatically."
        )

        with st.expander("🖼️ View Image Specifications", expanded=True):
            st.markdown(f"""
            <div class="preview-block">
{st.session_state.image_specs[:3_500]}
            </div>""", unsafe_allow_html=True)

        # Download spec sheet.
        if st.session_state.get("image_spec_doc"):
            safe_name = sanitise_filename(
                st.session_state.form_data.get("asset_name", "Asset")
            )
            st.download_button(
                "⬇️ Download Image Specification Sheet (.txt)",
                data=st.session_state.image_spec_doc,
                file_name=f"ImageSpecs_{safe_name}.txt",
                mime="text/plain",
            )

        render_review_gate(2)
        col_rerun, col_next = st.columns(2)
        with col_rerun:
            if render_rerun_button(2, "🔄 Re-run Agent 3"):
                st.rerun()
        with col_next:
            if render_next_agent_button(2):
                st.session_state.current_agent = 3
                st.rerun()


def _build_image_spec_doc(specs_json: str, raw_prompts: List[str]) -> bytes:
    """
    Assemble a human-readable image specification text document for download.

    Args:
        specs_json  : JSON string from Agent 3's Claude response.
        raw_prompts : Original image prompt strings extracted from the slide script.

    Returns:
        UTF-8 encoded bytes of the specification document.
    """
    asset  = st.session_state.form_data.get("asset_name", "")
    entity = st.session_state.form_data.get("entity_name", "")
    ts     = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    lines = [
        "=" * 72,
        "  IMAGE SPECIFICATION SHEET",
        f"  Asset  : {asset}",
        f"  Entity : {entity}",
        f"  Date   : {ts}",
        "=" * 72,
        "",
        "USAGE INSTRUCTIONS",
        "-" * 40,
        "1. Review each enhanced description below.",
        "2. Submit to your image API (DALL-E 3 / Stable Diffusion / Firefly).",
        "3. Save returned images as slide_01.png, slide_02.png, …",
        "4. Insert into the matching PPTX slide placeholder.",
        "",
        "API CODE STUB (Python / DALL-E 3):",
        "    import openai",
        "    response = openai.images.generate(",
        '        model="dall-e-3",',
        "        prompt=enhanced_description,",
        '        size="1792x1024",',
        '        quality="hd",',
        "        n=1,",
        "    )",
        "    image_url = response.data[0].url",
        "",
        "=" * 72,
        "  ENHANCED IMAGE SPECIFICATIONS (JSON)",
        "=" * 72,
        "",
        specs_json,
        "",
        "=" * 72,
        "  ORIGINAL PROMPTS (REFERENCE)",
        "=" * 72,
    ]
    for i, p in enumerate(raw_prompts, 1):
        lines.append(f"  Slide {i:02d}: {p}")

    return "\n".join(lines).encode("utf-8")


# =============================================================================
# SECTION 12 — AGENT 4: AUDIO VOICE-OVER AGENT
# =============================================================================

def run_agent4() -> None:
    """
    Agent 4 — Audio Voice-Over Agent.

    Orchestration flow:
      1. User selects voice configuration (speed, style, pause preference).
      2. Extract all VOICEOVER lines from pptx_script.
      3. Send to Claude for SSML enhancement (pause markers, emphasis, pronunciation).
      4. Generate a structured, downloadable TTS script package (.txt).
      5. Provide an audio preview simulation (text display with timing estimates).

    The output .txt file is ready for submission to:
      - Azure Cognitive Services TTS (en-IN, female voice)
      - Amazon Polly (Aditi or Kajal — Indian English)
      - ElevenLabs (custom voice clone)
      - Google TTS (en-IN wavenet female)

    Pre-condition: Agent 2 must be complete (agent_status[1] == "done").

    Output artefacts:
      st.session_state.audio_script  — enhanced script text
      st.session_state.audio_doc     — serialised TXT package bytes
    """
    st.markdown("### 🔊 Agent 4: Audio Voice-Over Agent")
    st.markdown("""
    <div class="card">
    Extracts presenter notes from every slide, enhances them with SSML markers,
    and generates a professional TTS-ready script package.
    Voice profile: warm, professional Indian female.
    Compatible with Azure TTS · Amazon Polly · ElevenLabs · Google TTS.
    </div>
    """, unsafe_allow_html=True)

    # ── Dependency check ──────────────────────────────────────────────────────
    if st.session_state.agent_status[1] != "done":
        st.warning("⚠️ Please complete **Agent 2 — PowerPoint Creator** first.")
        return

    progress_bar = st.progress(
        100 if st.session_state.agent_status[3] == "done" else 0
    )
    status_area = st.empty()

    # ── Voice configuration controls (shown before the run button) ─────────────
    if st.session_state.agent_status[3] != "done":
        st.markdown("#### 🎚️ Voice Configuration")
        col_speed, col_style, col_pause = st.columns(3)

        with col_speed:
            voice_speed = st.select_slider(
                "Speaking Speed",
                options=["Slow (0.75×)", "Normal (1.0×)", "Fast (1.25×)", "Brisk (1.5×)"],
                value="Normal (1.0×)",
                help="Controls SSML prosody rate attribute.",
            )
        with col_style:
            voice_style = st.selectbox(
                "Delivery Style",
                ["Professional & Warm", "Academic & Clear", "Engaging & Dynamic"],
                help="Influences tone and phrasing choices in the enhanced scripts.",
            )
        with col_pause:
            pause_style = st.selectbox(
                "Pause Style",
                [
                    "Natural breathing pauses (0.5s–1.0s)",
                    "Section-break pauses (1.5s–2.0s)",
                    "Minimal pauses (0.2s)",
                ],
                help="Controls the SSML <break> tag durations inserted between sentences.",
            )

        if st.button("▶️ Run Audio Voice-Over Agent", key="run_a4"):
            st.session_state.agent_status[3] = "running"
            st.session_state.current_agent   = 3
            progress_bar.progress(12)
            status_area.info("🔄 Extracting voice-over scripts from slide notes…")

            # Extract all VOICEOVER lines from the slide script.
            raw_voiceovers: List[str] = re.findall(
                r'VOICEOVER:\s*(.+)', st.session_state.pptx_script
            )

            system_prompt = (
                "You are a professional TTS script writer and voice-over director "
                "specialising in corporate e-learning content. "
                "Enhance each voice-over script to be natural, warm, and engaging. "
                "Add SSML-compatible markers using this notation: "
                "<break time='Xs'/> for pauses, <emphasis> for stressed words, "
                "[pronunciation: word=phonetic] for technical terms. "
                "Keep each enhanced script to 30–40 words. "
                "Always write in second person (you/your) for direct learner engagement."
            )

            numbered_voiceovers = "\n".join(
                f"Slide {i + 1}: {v}" for i, v in enumerate(raw_voiceovers[:25])
            )
            asset_name = st.session_state.form_data.get("asset_name", "")
            entity     = st.session_state.form_data.get("entity_name", "")

            full_prompt = f"""Enhance these voice-over scripts for professional TTS synthesis:

Voice profile  : Professional Indian female
Delivery style : {voice_style}
Speaking speed : {voice_speed}
Pause style    : {pause_style}
Asset          : {asset_name}
Entity         : {entity}

For EACH slide, respond with:
SLIDE [N]:
ENHANCED_SCRIPT: [30–40 words with SSML markers]
SSML_HINT      : [key pronunciation notes or emphasis guidance]
DURATION_EST   : [estimated playback seconds at chosen speed]

After all slides, add:
INTRO_SCRIPT   : [20-word deck opening / navigation script]
OUTRO_SCRIPT   : [20-word closing encouragement script]
TOTAL_DURATION : [estimated total audio duration in minutes]

Original scripts:
{numbered_voiceovers}"""

            progress_bar.progress(40)
            status_area.info("🤖 Generating SSML-enhanced voice-over scripts…")

            audio_script = call_claude(full_prompt, system_prompt, agent_idx=3)
            st.session_state.audio_script = audio_script

            # Build the downloadable TTS script package.
            audio_doc = _build_audio_script_doc(
                audio_script, raw_voiceovers, voice_style, voice_speed, pause_style
            )
            st.session_state.audio_doc = audio_doc

            st.session_state.agent_status[3]   = "done"
            st.session_state.agent_progress[3] = 100
            progress_bar.progress(100)
            status_area.success("✅ Voice-over scripts generated and packaged!")
            st.rerun()

    # ── Results panel ─────────────────────────────────────────────────────────
    if st.session_state.agent_status[3] == "done":
        progress_bar.progress(100)
        status_area.success("✅ Audio Agent complete!")

        st.info(
            "ℹ️ Submit the downloaded .txt script package to your TTS service. "
            "Recommended: **Azure TTS** (voice: `hi-IN-SwaraNeural`) or "
            "**ElevenLabs** (Rachel or custom Indian English voice clone)."
        )

        with st.expander("🔊 View Enhanced Voice-Over Scripts", expanded=True):
            st.markdown(f"""
            <div class="preview-block">
{st.session_state.audio_script[:3_500]}
            </div>""", unsafe_allow_html=True)

        # Download button.
        if st.session_state.audio_doc:
            safe_name = sanitise_filename(
                st.session_state.form_data.get("asset_name", "Asset")
            )
            st.download_button(
                "⬇️ Download Voice-Over Script Package (.txt)",
                data=st.session_state.audio_doc,
                file_name=f"VoiceOver_{safe_name}.txt",
                mime="text/plain",
            )

        render_review_gate(3)
        col_rerun, col_next = st.columns(2)
        with col_rerun:
            if render_rerun_button(3, "🔄 Re-run Agent 4 (new voice config)"):
                st.rerun()
        with col_next:
            if render_next_agent_button(3):
                st.session_state.current_agent = 4
                st.rerun()


def _build_audio_script_doc(
    enhanced: str,
    originals: List[str],
    style: str,
    speed: str,
    pause: str,
) -> bytes:
    """
    Assemble the final TTS script package as a downloadable .txt file.

    Args:
        enhanced  : SSML-enhanced scripts from Claude.
        originals : Original voiceover strings from the slide script.
        style     : Voice delivery style string.
        speed     : Speaking speed string.
        pause     : Pause style string.

    Returns:
        UTF-8 encoded bytes of the script package.
    """
    asset  = st.session_state.form_data.get("asset_name", "")
    entity = st.session_state.form_data.get("entity_name", "")
    ts     = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    lines = [
        "=" * 72,
        "  VOICE-OVER SCRIPT PACKAGE — TTS READY",
        f"  Asset         : {asset}",
        f"  Entity        : {entity}",
        f"  Generated     : {ts}",
        "=" * 72,
        "",
        "VOICE PROFILE SETTINGS",
        "-" * 40,
        "  Voice     : Professional Indian Female",
        f"  Style     : {style}",
        f"  Speed     : {speed}",
        f"  Pauses    : {pause}",
        "",
        "SSML NOTATION GUIDE",
        "-" * 40,
        "  <break time='1s'/>           — 1-second pause",
        "  <emphasis level='strong'>    — stress this word",
        "  [pronunciation: X=Y]         — pronounce X as phonetic Y",
        "",
        "RECOMMENDED TTS SERVICES",
        "-" * 40,
        "  Azure TTS    : hi-IN-SwaraNeural  (Hindi-English natural voice)",
        "  Amazon Polly : Aditi / Kajal      (Indian English NTTS)",
        "  ElevenLabs   : Custom voice clone  (highest naturalness)",
        "  Google TTS   : en-IN-Wavenet-C    (Indian English female)",
        "",
        "=" * 72,
        "  ENHANCED SCRIPTS (SSML-MARKED)",
        "=" * 72,
        "",
        enhanced,
        "",
        "=" * 72,
        "  ORIGINAL SCRIPTS (REFERENCE)",
        "=" * 72,
    ]
    for i, v in enumerate(originals, 1):
        lines.append(f"  Slide {i:02d}: {v}")

    return "\n".join(lines).encode("utf-8")


# =============================================================================
# SECTION 13 — AGENT 5: DECK VALIDATOR
# =============================================================================

def run_agent5() -> None:
    """
    Agent 5 — Deck Validator.

    Orchestration flow:
      1. Aggregate context: programme details, guide-sheet, slide script, image specs, audio.
      2. Send to Claude with a 7-dimension evaluation rubric.
      3. Parse the response for scores and overall verdict.
      4. Display scorecard with KPI metrics and colour-coded verdict.
      5. Provide a downloadable validation report.

    Validation dimensions (each scored 0–100):
      1. Content Authenticity    — originality, plagiarism risk
      2. Content Accuracy        — factual correctness, standard alignment
      3. Activity Validation     — MCQ logic, fill-blank coherence
      4. Learning Effectiveness  — objective–outcome–assessment alignment
      5. Visual Appropriateness  — image prompt relevance, brand compliance
      6. Voice-Over Quality      — script fluency, timing, navigation
      7. Overall Feasibility     — cognitive load, time-on-task, implementation

    Pre-condition: Agent 2 must be complete (agent_status[1] == "done").

    Output artefact:
      st.session_state.validation_rpt  — full report text
    """
    st.markdown("### ✅ Agent 5: Deck Validator")
    st.markdown("""
    <div class="card">
    Performs a rigorous 7-dimension quality assurance review of the entire generated
    learning asset.  Scores each dimension 0–100 with TESTED OK / NEEDS REVIEW verdicts,
    and issues an overall APPROVED FOR USE / CONDITIONAL APPROVAL / NEEDS REVISION verdict.
    </div>
    """, unsafe_allow_html=True)

    # ── Dependency check ──────────────────────────────────────────────────────
    if st.session_state.agent_status[1] != "done":
        st.warning("⚠️ Please complete at least **Agent 2 — PowerPoint Creator** first.")
        return

    progress_bar = st.progress(
        100 if st.session_state.agent_status[4] == "done" else 0
    )
    status_area = st.empty()

    # ── Run button ────────────────────────────────────────────────────────────
    if st.session_state.agent_status[4] != "done":
        if st.button("▶️ Run Deck Validator Agent", key="run_a5"):
            st.session_state.agent_status[4] = "running"
            st.session_state.current_agent   = 4
            progress_bar.progress(8)
            status_area.info("🔄 Aggregating all agent outputs for validation…")

            system_prompt = (
                "You are a rigorous quality assurance specialist and learning experience designer. "
                "Validate training content across multiple professional dimensions. "
                "Be thorough, objective, specific, and constructive. "
                "Every score must be supported by a concrete finding. "
                "Use the EXACT response format specified below."
            )

            ctx         = build_context_string()
            guide_sum   = st.session_state.guidesheet_txt[:2_000]
            slide_sum   = st.session_state.pptx_script[:2_000]
            image_sum   = st.session_state.image_specs[:800]
            audio_sum   = st.session_state.audio_script[:800]

            full_prompt = f"""Perform a comprehensive validation of this training learning asset.

PROGRAMME CONTEXT:
{ctx}

GUIDE-SHEET SUMMARY (first 2000 chars):
{guide_sum}

SLIDE SCRIPT SUMMARY (first 2000 chars):
{slide_sum}

IMAGE SPECS SUMMARY:
{image_sum}

AUDIO SCRIPT SUMMARY:
{audio_sum}

VALIDATION RUBRIC — score each dimension 0–100:

1. CONTENT AUTHENTICITY
   • Is the content original and industry-accurate?
   • Plagiarism risk: Low / Medium / High
   • Does content reflect current (2024–2025) standards?
   Verdict: TESTED OK / NEEDS REVIEW

2. CONTENT ACCURACY
   • Factual correctness of technical content
   • Accuracy of examples, scenarios, and statistics
   • Alignment with current industry standards and best practices
   Verdict: TESTED OK / NEEDS REVIEW

3. ACTIVITY VALIDATION
   • Each MCQ must have exactly ONE correct option (marked *)
   • Fill-in-blank sentences must have clear, unambiguous answers
   • Difficulty level is appropriate for the target audience
   • "Well Done" / "Try Again" feedback logic is coherent
   Verdict: TESTED OK / NEEDS REVIEW

4. LEARNING EFFECTIVENESS
   • All slides align to stated learning objectives (Section 2 of guide-sheet)
   • Coverage is complete vs the guide-sheet module structure
   • Bloom's taxonomy progression is evident (remember → evaluate → create)
   • Assessment items map to specific learning outcomes
   Verdict: TESTED OK / NEEDS REVIEW

5. VISUAL & IMAGE APPROPRIATENESS
   • Image prompts are relevant, brand-compliant, and professional
   • No inappropriate, ambiguous, or culturally insensitive imagery
   • Accessibility: adequate contrast and clear visual communication
   • Image placement enhances rather than distracts from content
   Verdict: TESTED OK / NEEDS REVIEW

6. VOICE-OVER QUALITY
   • Scripts are natural, warm, and suitable for a professional female voice
   • Each script is ≤ 35 words (optimal for 30–45 second audio segments)
   • Navigation and transition language is clear and engaging
   • SSML markers are correctly applied
   Verdict: TESTED OK / NEEDS REVIEW

7. OVERALL FEASIBILITY
   • Total content fits within the stated {st.session_state.form_data.get("duration", 16)}-hour duration
   • Cognitive load per slide is appropriate for the target audience
   • Technical implementation (PPTX, audio, images) is achievable
   • Learner engagement level is maintained throughout
   Verdict: TESTED OK / NEEDS REVIEW

REQUIRED RESPONSE FORMAT:

## VALIDATION SCORECARD

| Dimension | Score /100 | Status | Key Finding |
|-----------|-----------|--------|-------------|
| Content Authenticity | XX | TESTED OK/NEEDS REVIEW | one-line finding |
| Content Accuracy | XX | … | … |
| Activity Validation | XX | … | … |
| Learning Effectiveness | XX | … | … |
| Visual Appropriateness | XX | … | … |
| Voice-Over Quality | XX | … | … |
| Overall Feasibility | XX | … | … |

## DETAILED FINDINGS

[For each dimension: 3–5 specific findings, with evidence from the content]

## OVERALL VERDICT
Overall Score: XX/100
Status: APPROVED FOR USE / CONDITIONAL APPROVAL / NEEDS REVISION
Conditions (if Conditional): [list any conditions]
Recommended for use by: [target entity and audience]

## RECOMMENDATIONS FOR IMPROVEMENT
[Numbered list of 5–8 specific, actionable, prioritised improvements]"""

            progress_bar.progress(25)
            status_area.info("🤖 Running 7-dimension quality assurance review…")

            # Staged progress updates for UX responsiveness.
            for pct in (45, 62, 78):
                time.sleep(0.25)
                progress_bar.progress(pct)

            validation_report = call_claude(full_prompt, system_prompt, agent_idx=4)
            st.session_state.validation_rpt = validation_report

            st.session_state.agent_status[4]   = "done"
            st.session_state.agent_progress[4] = 100
            progress_bar.progress(100)
            status_area.success("✅ Deck validation complete!")
            st.rerun()

    # ── Results panel ─────────────────────────────────────────────────────────
    if st.session_state.agent_status[4] == "done":
        progress_bar.progress(100)
        status_area.success("✅ Validation complete!")

        report = st.session_state.validation_rpt

        # ── Parse score and verdict ───────────────────────────────────────────
        score_match   = re.search(r'Overall Score\s*[:\-]\s*(\d+)', report, re.I)
        verdict_match = re.search(
            r'Status\s*[:\-]\s*(APPROVED FOR USE|CONDITIONAL APPROVAL|NEEDS REVISION)',
            report, re.I,
        )
        overall_score   = score_match.group(1)   if score_match   else "—"
        overall_verdict = verdict_match.group(1) if verdict_match else "REVIEWED"

        # Colour-code the verdict.
        verdict_color = (
            BRAND["success"] if "APPROVED"     in overall_verdict.upper() else
            BRAND["warning"] if "CONDITIONAL"  in overall_verdict.upper() else
            BRAND["danger"]
        )

        # ── KPI summary row ───────────────────────────────────────────────────
        col_score, col_verdict, col_tokens = st.columns(3)
        with col_score:
            st.markdown(f"""
            <div class="token-box">
                <div style="font-size:0.7rem;color:rgba(255,255,255,0.5)">OVERALL SCORE</div>
                <div class="token-number">{overall_score}/100</div>
            </div>""", unsafe_allow_html=True)
        with col_verdict:
            st.markdown(f"""
            <div style="background:rgba(0,0,0,0.3);border:1px solid {verdict_color};
                        border-radius:12px;padding:1rem;text-align:center;">
                <div style="font-size:0.7rem;color:rgba(255,255,255,0.5)">VERDICT</div>
                <div style="font-family:'Sora',sans-serif;font-size:0.95rem;
                            font-weight:700;color:{verdict_color}">
                    {overall_verdict}
                </div>
            </div>""", unsafe_allow_html=True)
        with col_tokens:
            st.markdown(f"""
            <div class="token-box">
                <div style="font-size:0.7rem;color:rgba(255,255,255,0.5)">AGENT TOKENS</div>
                <div class="token-number">
                    {st.session_state.agent_tokens[4]:,}
                </div>
            </div>""", unsafe_allow_html=True)

        # ── Full report ───────────────────────────────────────────────────────
        st.markdown('<div class="section-label">📋 Full Validation Report</div>',
                    unsafe_allow_html=True)
        st.markdown(
            f'<div class="preview-block" style="max-height:520px;">{report}</div>',
            unsafe_allow_html=True)

        # Download report.
        safe_name = sanitise_filename(
            st.session_state.form_data.get("asset_name", "Asset")
        )
        st.download_button(
            "⬇️ Download Validation Report (.txt)",
            data=report.encode("utf-8"),
            file_name=f"ValidationReport_{safe_name}.txt",
            mime="text/plain",
        )

        render_review_gate(4)

        # Re-run button (no next-agent for the final agent).
        if render_rerun_button(4, "🔄 Re-run Validator (after deck updates)"):
            st.rerun()


# =============================================================================
# SECTION 14 — PAGE: AGENT WORKSPACE
# =============================================================================

def render_workspace() -> None:
    """
    Render the Agent Workspace page.

    Displays a tabbed interface with one tab per agent.  Each tab renders
    its corresponding run_agentN() function independently.

    Pre-condition: form_data must be populated (i.e. user has submitted the
    Landing Page form).  If not, redirects back to Landing Page.
    """
    # ── Dark top banner ───────────────────────────────────────────────────────
    job    = st.session_state.get("job_name", "Unnamed Job")
    entity = st.session_state.form_data.get("entity_name", "")
    st.markdown(f"""
    <div class="top-banner">
        <div class="banner-logo"><div class="logo-hcl">HCLTech</div>
        <div class="logo-cs">Career Shaper™</div></div>
        <div class="banner-text">
            <div class="banner-title">Agent Workspace</div>
            <div class="banner-sub">{job} &nbsp;·&nbsp; {entity}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Guard: redirect if no form data ───────────────────────────────────────
    if not st.session_state.form_data:
        st.warning(
            "⚠️ No job configured. "
            "Please complete the Landing Page form and click START."
        )
        if st.button("← Return to Landing Page"):
            st.session_state.current_page = "landing"
            st.rerun()
        return

    # ── Agent tabs ────────────────────────────────────────────────────────────
    tab_labels = [
        f"{icon}  Agent {i + 1}"
        for i, icon in enumerate(AGENT_ICONS)
    ]
    tabs = st.tabs(tab_labels)

    with tabs[0]:
        run_agent1()
    with tabs[1]:
        run_agent2()
    with tabs[2]:
        run_agent3()
    with tabs[3]:
        run_agent4()
    with tabs[4]:
        run_agent5()


# =============================================================================
# SECTION 15 — MAIN ROUTER
# =============================================================================

def main() -> None:
    """
    Application entry point and page router.

    Renders the sidebar (always present) then dispatches to the appropriate
    page renderer based on st.session_state.current_page.

    Pages:
      "landing"   → render_landing()   — input form
      "admin"     → render_admin()     — API key management + token log
      "workspace" → render_workspace() — 5-agent pipeline tabs
      (default)   → render_landing()
    """
    render_sidebar()

    page = st.session_state.current_page

    if page == "landing":
        render_landing()
    elif page == "admin":
        render_admin()
    elif page == "workspace":
        render_workspace()
    else:
        logger.warning("Unknown page '%s' — defaulting to landing.", page)
        render_landing()


# ── Script entry point ────────────────────────────────────────────────────────
# Streamlit re-executes the entire script on each user interaction.
# All top-level code must be idempotent; mutable state lives in session_state.
if __name__ == "__main__":
    main()
else:
    # When imported by Streamlit's runner (not __main__), still call main().
    main()