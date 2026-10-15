#!/usr/bin/env python3
"""
Generates PowerPoint (PPTX) and PDF presentation slides for the College Mini-Project:
Break-Even Point Predictor & Feasibility Engine.
Features:
- Unified Full-Stack Architecture slide displaying both Frontend and Backend together on the SAME slide.
- Embedded live screenshots.
- Exactly 7 slides covering the full academic project lifecycle.
"""

import sys
import os

# Ensure user site-packages are accessible
user_site = os.path.expanduser('~/Library/Python/3.9/lib/python/site-packages')
if user_site not in sys.path:
    sys.path.insert(0, user_site)

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as RLImage

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PPTX_PATH = os.path.join(BASE_DIR, "Break-Even-Project-Presentation.pptx")
PDF_PATH = os.path.join(BASE_DIR, "Break-Even-Project-Presentation.pdf")

SCREENSHOT_DASH = os.path.join(BASE_DIR, "screenshot_dashboard.png")
SCREENSHOT_LOGIN = os.path.join(BASE_DIR, "screenshot_login.png")
SCREENSHOT_CHART = os.path.join(BASE_DIR, "screenshot_chart.png")

# Color Palette
C_PRIMARY = RGBColor(37, 99, 235)      # #2563eb Blue
C_DARK = RGBColor(15, 23, 42)          # #0f172a Navy
C_MUTED = RGBColor(100, 116, 139)      # #64748b Slate
C_LIGHT_BG = RGBColor(248, 250, 252)   # #f8fafc Light Gray
C_CARD_BORDER = RGBColor(226, 232, 240)# #e2e8f0 Border
C_SUCCESS = RGBColor(5, 150, 105)      # #059669 Green
C_ACCENT = RGBColor(124, 58, 237)      # #7c3aed Purple
C_WHITE = RGBColor(255, 255, 255)

# ==============================================================================
# 1. PPTX GENERATION (7 SLIDES WITH FRONTEND & BACKEND ON SAME SLIDE)
# ==============================================================================
def create_pptx():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    def add_header(slide, title_text, category_text="COLLEGE MINI PROJECT"):
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.35))
        tf_c = cat_box.text_frame
        tf_c.word_wrap = True
        p_c = tf_c.paragraphs[0]
        p_c.text = category_text.upper()
        p_c.font.size = Pt(11)
        p_c.font.bold = True
        p_c.font.color.rgb = C_PRIMARY

        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(11.7), Inches(0.65))
        tf_t = title_box.text_frame
        tf_t.word_wrap = True
        p_t = tf_t.paragraphs[0]
        p_t.text = title_text
        p_t.font.size = Pt(24)
        p_t.font.bold = True
        p_t.font.color.rgb = C_DARK

    def add_card(slide, left, top, width, height, bg_color=C_LIGHT_BG, border_color=C_CARD_BORDER):
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = bg_color
        shape.line.color.rgb = border_color
        shape.line.width = Pt(1.5)
        return shape

    # --------------------------------------------------------------------------
    # SLIDE 1: Title & Project Overview (With Live Preview)
    # --------------------------------------------------------------------------
    slide1 = prs.slides.add_slide(blank_layout)
    add_card(slide1, Inches(0.8), Inches(0.8), Inches(11.733), Inches(5.9), C_LIGHT_BG, C_PRIMARY)

    title_box = slide1.shapes.add_textbox(Inches(1.2), Inches(1.2), Inches(6.5), Inches(2.2))
    tf1 = title_box.text_frame
    tf1.word_wrap = True

    p0 = tf1.paragraphs[0]
    p0.text = "ACADEMIC MINI PROJECT PRESENTATION"
    p0.font.size = Pt(12)
    p0.font.bold = True
    p0.font.color.rgb = C_PRIMARY

    p1 = tf1.add_paragraph()
    p1.text = "Break-Even Point Predictor & Feasibility Engine"
    p1.font.size = Pt(28)
    p1.font.bold = True
    p1.font.color.rgb = C_DARK
    p1.space_before = Pt(6)

    p2 = tf1.add_paragraph()
    p2.text = "Full-Stack Financial Modeling Web App with Real-Time SVG Charts & Multi-Runtime Backend"
    p2.font.size = Pt(13)
    p2.font.color.rgb = C_MUTED
    p2.space_before = Pt(4)

    m_card = add_card(slide1, Inches(1.2), Inches(3.6), Inches(5.8), Inches(2.7), C_WHITE)
    tf_m = m_card.text_frame
    tf_m.word_wrap = True
    tf_m.margin_left = Inches(0.25)
    tf_m.margin_top = Inches(0.2)

    p_h = tf_m.paragraphs[0]
    p_h.text = "PROJECT SPECIFICATIONS"
    p_h.font.size = Pt(11)
    p_h.font.bold = True
    p_h.font.color.rgb = C_PRIMARY

    meta = [
        ("Student / Developer:", "arjun2007arush2011-max"),
        ("Live Web Application:", "https://arjun2007arush2011-max.github.io/break_even_point/"),
        ("GitHub Repository:", "https://github.com/arjun2007arush2011-max/break_even_point"),
        ("Domain:", "FinTech & Decision Support Systems"),
        ("Architecture:", "Client-Side Zero-Dependency + REST Backends")
    ]
    for k, v in meta:
        p = tf_m.add_paragraph()
        p.text = f"{k} {v}"
        p.font.size = Pt(10)
        p.font.color.rgb = C_DARK
        p.space_before = Pt(3)

    if os.path.exists(SCREENSHOT_DASH):
        slide1.shapes.add_picture(SCREENSHOT_DASH, Inches(7.3), Inches(1.4), Inches(4.8), Inches(4.8))

    # --------------------------------------------------------------------------
    # SLIDE 2: Problem Statement & Project Objectives
    # --------------------------------------------------------------------------
    slide2 = prs.slides.add_slide(blank_layout)
    add_header(slide2, "Problem Statement & Project Objectives", "BACKGROUND & MOTIVATION")

    c2_l = add_card(slide2, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.1), C_WHITE)
    tf2_l = c2_l.text_frame
    tf2_l.word_wrap = True
    tf2_l.margin_left = Inches(0.35)
    tf2_l.margin_top = Inches(0.3)

    h2_l = tf2_l.paragraphs[0]
    h2_l.text = "THE PROBLEM STATEMENT"
    h2_l.font.size = Pt(12)
    h2_l.font.bold = True
    h2_l.font.color.rgb = RGBColor(220, 38, 38)

    pts1 = [
        "High Startup Mortality: Over 65% of small business failures stem from inaccurate cost estimation and a failure to identify break-even volume early.",
        "Spreadsheet Fragility: Static Excel files lack interactive real-time visual curves and are error-prone on mobile or collaborative settings.",
        "Blind Pricing Strategies: Entrepreneurs frequently set prices without analyzing per-unit variable costs, selling products at an unseen structural loss.",
        "Lack of Sensitivity Analysis: Traditional tools fail to simulate what happens when supplier variable costs rise by 15% or rent increases by 20%."
    ]
    for pt in pts1:
        p = tf2_l.add_paragraph()
        p.text = f"• {pt}"
        p.font.size = Pt(10.5)
        p.font.color.rgb = C_DARK
        p.space_before = Pt(8)

    c2_r = add_card(slide2, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.1), C_WHITE)
    tf2_r = c2_r.text_frame
    tf2_r.word_wrap = True
    tf2_r.margin_left = Inches(0.35)
    tf2_r.margin_top = Inches(0.3)

    h2_r = tf2_r.paragraphs[0]
    h2_r.text = "PROJECT OBJECTIVES"
    h2_r.font.size = Pt(12)
    h2_r.font.bold = True
    h2_r.font.color.rgb = C_SUCCESS

    pts2 = [
        "Zero-Dependency Engineering: Deliver a self-contained, lightweight financial simulator operating without third-party frameworks or CDNs.",
        "Dynamic SVG Visualization: Render real-time cost curves, break-even intersection callouts, and shaded loss/profit zones.",
        "Real-Time What-If Simulations: Provide instant slider testing (-50% to +50%) and sensitivity tables to evaluate cost risks.",
        "Multi-Tier REST Integration: Implement backend APIs across Python, Ruby, and Node.js with SQLite3 database persistence.",
        "Executive Feasibility Reporting: Support one-click printable reports formatted with professional letterheads and KPI matrices."
    ]
    for pt in pts2:
        p = tf2_r.add_paragraph()
        p.text = f"✔ {pt}"
        p.font.size = Pt(10.5)
        p.font.color.rgb = C_DARK
        p.space_before = Pt(8)

    # --------------------------------------------------------------------------
    # SLIDE 3: UNIFIED FULL-STACK ARCHITECTURE (FRONTEND & BACKEND ON SAME SLIDE!)
    # --------------------------------------------------------------------------
    slide3 = prs.slides.add_slide(blank_layout)
    add_header(slide3, "Full-Stack System Architecture: Frontend & Backend", "SYSTEM ARCHITECTURE")

    # Top Left Card: Frontend Client Tier
    c3_l = add_card(slide3, Inches(0.8), Inches(1.5), Inches(5.7), Inches(3.95), C_WHITE)
    tf3_l = c3_l.text_frame
    tf3_l.word_wrap = True
    tf3_l.margin_left = Inches(0.25)
    tf3_l.margin_top = Inches(0.2)

    h3_l = tf3_l.paragraphs[0]
    h3_l.text = "FRONTEND CLIENT TIER (Browser)"
    h3_l.font.size = Pt(12)
    h3_l.font.bold = True
    h3_l.font.color.rgb = C_PRIMARY

    f_points = [
        ("Vanilla ES5 Standards:", "Zero build pipelines, npm, or Babel. Operates natively in 100% of web browsers."),
        ("Responsive CSS Grid:", "Two-column analytical dashboard adapting smoothly to single-column mobile viewports."),
        ("Dark & Light Mode Engine:", "CSS Custom Properties with persistent user theme saved in localStorage."),
        ("Client-Side Cryptography:", "Web Crypto API (crypto.subtle) SHA-256 password hashing with pure JS fallback."),
        ("Modular Namespaces:", "Separation of concerns: AuthModule, ProfileModule, MetricsModule, ChartModule, ApiClient.")
    ]
    for title, desc in f_points:
        p = tf3_l.add_paragraph()
        p.text = f"• {title} "
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = C_DARK
        p.space_before = Pt(3)

        p_desc = tf3_l.paragraphs[-1]
        run_desc = p_desc.add_run()
        run_desc.text = desc
        run_desc.font.bold = False
        run_desc.font.color.rgb = C_MUTED

    # Top Right Card: Backend Server & Database Tier
    c3_r = add_card(slide3, Inches(6.8), Inches(1.5), Inches(5.7), Inches(3.95), C_WHITE)
    tf3_r = c3_r.text_frame
    tf3_r.word_wrap = True
    tf3_r.margin_left = Inches(0.25)
    tf3_r.margin_top = Inches(0.2)

    h3_r = tf3_r.paragraphs[0]
    h3_r.text = "BACKEND SERVER & DATABASE TIER"
    h3_r.font.size = Pt(12)
    h3_r.font.bold = True
    h3_r.font.color.rgb = C_SUCCESS

    b_points = [
        ("Triple-Runtime Implementations:", "Native Python 3 (http.server), Ruby (WEBrick), and Node.js. Zero pip/npm dependencies."),
        ("Universal Launcher (start.sh):", "Intelligently auto-detects the host system's installed runtime and binds on port 8000."),
        ("Relational SQLite3 Database:", "Structured breakeven.db with normalized schemas for users, profiles, and inputs."),
        ("RESTful JSON API Contract:", "Endpoints for /api/auth/signup, /api/auth/login, /api/profile, /api/inputs, /api/calculate."),
        ("Server Mathematical Validation:", "Performs secondary verification of break-even formulas to enforce data integrity.")
    ]
    for title, desc in b_points:
        p = tf3_r.add_paragraph()
        p.text = f"• {title} "
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = C_DARK
        p.space_before = Pt(3)

        p_desc = tf3_r.paragraphs[-1]
        run_desc = p_desc.add_run()
        run_desc.text = desc
        run_desc.font.bold = False
        run_desc.font.color.rgb = C_MUTED

    # Bottom Full-Width Card: Communication Protocol & Dual Sync Bridge
    c3_b = add_card(slide3, Inches(0.8), Inches(5.6), Inches(11.7), Inches(1.4), C_LIGHT_BG, C_ACCENT)
    tf3_b = c3_b.text_frame
    tf3_b.word_wrap = True
    tf3_b.margin_left = Inches(0.3)
    tf3_b.margin_top = Inches(0.15)

    h3_b = tf3_b.paragraphs[0]
    h3_b.text = "REST API PROTOCOL & RESILIENT DUAL-ENGINE SYNCHRONIZATION BRIDGE (Frontend ⟷ Backend)"
    h3_b.font.size = Pt(11)
    h3_b.font.bold = True
    h3_b.font.color.rgb = C_ACCENT

    p_sync1 = tf3_b.add_paragraph()
    p_sync1.text = "• Debounced REST Sync: Client debounces user input (300ms) and dispatches JSON payloads with Bearer auth headers (Authorization: Bearer <token>) to /api/inputs and /api/profile."
    p_sync1.font.size = Pt(9)
    p_sync1.font.color.rgb = C_DARK
    p_sync1.space_before = Pt(2)

    p_sync2 = tf3_b.add_paragraph()
    p_sync2.text = "• Resilient Dual Persistence: ApiClient probes /api/health. If server is active, data saves to SQLite3; if offline or hosted on GitHub Pages (file://), it gracefully falls back to localStorage with zero downtime."
    p_sync2.font.size = Pt(9)
    p_sync2.font.color.rgb = C_DARK
    p_sync2.space_before = Pt(2)

    # --------------------------------------------------------------------------
    # SLIDE 4: Interactive SVG Chart & Visualization Engine
    # --------------------------------------------------------------------------
    slide4 = prs.slides.add_slide(blank_layout)
    add_header(slide4, "Interactive SVG Chart & Visualization Engine", "VISUALIZATION ENGINE")

    # Left: Chart Screenshot
    if os.path.exists(SCREENSHOT_CHART):
        slide4.shapes.add_picture(SCREENSHOT_CHART, Inches(0.8), Inches(1.6), Inches(6.0), Inches(5.1))
    elif os.path.exists(SCREENSHOT_DASH):
        slide4.shapes.add_picture(SCREENSHOT_DASH, Inches(0.8), Inches(1.6), Inches(6.0), Inches(5.1))

    # Right: Visualization Details
    c4_r = add_card(slide4, Inches(7.1), Inches(1.6), Inches(5.4), Inches(5.1), C_WHITE)
    tf4_r = c4_r.text_frame
    tf4_r.word_wrap = True
    tf4_r.margin_left = Inches(0.3)
    tf4_r.margin_top = Inches(0.3)

    h4_r = tf4_r.paragraphs[0]
    h4_r.text = "CUSTOM COORDINATE MAPPING ENGINE"
    h4_r.font.size = Pt(12)
    h4_r.font.bold = True
    h4_r.font.color.rgb = C_ACCENT

    vis_details = [
        ("No Third-Party Charting Libraries:", "All graphics generated dynamically via native inline SVG elements, avoiding heavy dependencies like Chart.js or D3."),
        ("Dynamic Coordinate Projection:", "Translates financial values into SVG viewBox space using getX(q) and getY(val) with dynamic domain bounds."),
        ("Visual Profit & Loss Polygons:", "Shades the operating loss zone in red (0 to Break-Even) and profit zone in green (Break-Even to Max Volume)."),
        ("Break-Even Callout Point:", "Highlights the exact intersection point of Total Revenue and Total Cost with a pulsing halo and coordinate badge."),
        ("Interactive Crosshair & Tooltip:", "Hover/touch tracking computes volume, revenue, costs, and profit with 60fps responsiveness.")
    ]
    for title, desc in vis_details:
        p = tf4_r.add_paragraph()
        p.text = f"• {title}"
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = C_DARK
        p.space_before = Pt(5)

        p_b = tf4_r.add_paragraph()
        p_b.text = f"  {desc}"
        p_b.font.size = Pt(9.5)
        p_b.font.color.rgb = C_MUTED
        p_b.space_before = Pt(1)

    # --------------------------------------------------------------------------
    # SLIDE 5: Mathematical Formulation & Core Financial Algorithms
    # --------------------------------------------------------------------------
    slide5 = prs.slides.add_slide(blank_layout)
    add_header(slide5, "Mathematical Formulation & Core Financial Algorithms", "FINANCIAL MODEL & ALGORITHMS")

    c5_l = add_card(slide5, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.1), C_WHITE)
    tf5_l = c5_l.text_frame
    tf5_l.word_wrap = True
    tf5_l.margin_left = Inches(0.3)
    tf5_l.margin_top = Inches(0.3)

    h5_l = tf5_l.paragraphs[0]
    h5_l.text = "CORE BREAK-EVEN EQUATIONS"
    h5_l.font.size = Pt(12)
    h5_l.font.bold = True
    h5_l.font.color.rgb = C_PRIMARY

    math_eqs = [
        ("Contribution Margin (CM):", "CM = Selling Price (P) - Variable Cost (V)"),
        ("Contribution Margin Ratio:", "CM% = (P - V) / P"),
        ("Break-Even Volume (Units):", "Q_BE = Fixed Costs (F) / (P - V)"),
        ("Break-Even Sales Revenue:", "R_BE = Q_BE × P = F / CM%"),
        ("Net Operating Profit:", "Profit = Q × (P - V) - F"),
        ("Margin of Safety (MOS):", "MOS = (Q - Q_BE) / Q × 100%"),
        ("Degree of Operating Leverage:", "DOL = (Q × CM) / Profit"),
        ("Target Profit Units:", "Q_Target = (F + Target) / CM")
    ]
    for title, formula in math_eqs:
        p = tf5_l.add_paragraph()
        p.text = f"{title} {formula}"
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_DARK
        p.space_before = Pt(4)

    c5_r = add_card(slide5, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.1), C_WHITE)
    tf5_r = c5_r.text_frame
    tf5_r.word_wrap = True
    tf5_r.margin_left = Inches(0.3)
    tf5_r.margin_top = Inches(0.3)

    h5_r = tf5_r.paragraphs[0]
    h5_r.text = "ALGORITHMIC GUARDS & INTEGRITY CHECKS"
    h5_r.font.size = Pt(12)
    h5_r.font.bold = True
    h5_r.font.color.rgb = C_SUCCESS

    guards = [
        ("Structural Loss Detection (P ≤ V):", "Instantly flags when price does not exceed variable costs. Prevents division-by-zero exceptions and alerts that every unit sold widens net deficit."),
        ("Sub-Break-Even Operating Alert:", "Dynamically signals danger when planned sales volume Q < Q_BE, detailing the precise unit deficit required to achieve solvency."),
        ("Precision Floating-Point Normalization:", "Sanitizes user currency entries and eliminates IEEE 754 floating-point inaccuracies through explicit rounding rules."),
        ("Client-Server Calculation Parity:", "Validates client results against backend /api/calculate endpoints to guarantee consistent cross-platform mathematical output.")
    ]
    for title, desc in guards:
        p = tf5_r.add_paragraph()
        p.text = f"✔ {title}"
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = C_DARK
        p.space_before = Pt(6)

        p_b = tf5_r.add_paragraph()
        p_b.text = f"  {desc}"
        p_b.font.size = Pt(9.5)
        p_b.font.color.rgb = C_MUTED
        p_b.space_before = Pt(1)

    # --------------------------------------------------------------------------
    # SLIDE 6: What-If Sensitivity Simulations & Scenario Planning
    # --------------------------------------------------------------------------
    slide6 = prs.slides.add_slide(blank_layout)
    add_header(slide6, "What-If Sensitivity Simulations & Scenario Planning", "SIMULATION & SENSITIVITY")

    c6_l = add_card(slide6, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.1), C_WHITE)
    tf6_l = c6_l.text_frame
    tf6_l.word_wrap = True
    tf6_l.margin_left = Inches(0.3)
    tf6_l.margin_top = Inches(0.3)

    h6_l = tf6_l.paragraphs[0]
    h6_l.text = "REAL-TIME WHAT-IF SLIDERS"
    h6_l.font.size = Pt(12)
    h6_l.font.bold = True
    h6_l.font.color.rgb = C_PRIMARY

    slider_pts = [
        ("Bidirectional Sliders (-50% to +50%):", "Instant continuous adjustment of Selling Price, Variable Cost, and Fixed Costs."),
        ("Live Delta & Variance Badges:", "Computes real-time percentage shifts and absolute monetary changes against base parameters."),
        ("Immediate Visual Curve Morphing:", "Recalculates the SVG graph at 60fps, visually shifting the intersection point along the volume axis."),
        ("One-Click Baseline Restoration:", "Resets all simulation parameters back to saved user baselines without page reloads.")
    ]
    for title, desc in slider_pts:
        p = tf6_l.add_paragraph()
        p.text = title
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = C_DARK
        p.space_before = Pt(6)

        p_b = tf6_l.add_paragraph()
        p_b.text = f"  {desc}"
        p_b.font.size = Pt(9.5)
        p_b.font.color.rgb = C_MUTED
        p_b.space_before = Pt(1)

    c6_r = add_card(slide6, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.1), C_WHITE)
    tf6_r = c6_r.text_frame
    tf6_r.word_wrap = True
    tf6_r.margin_left = Inches(0.3)
    tf6_r.margin_top = Inches(0.3)

    h6_r = tf6_r.paragraphs[0]
    h6_r.text = "STRESS TESTING & REPORT GENERATION"
    h6_r.font.size = Pt(12)
    h6_r.font.bold = True
    h6_r.font.color.rgb = C_SUCCESS

    sims = [
        ("5-Tier Sensitivity Matrix Table:", "Evaluates Required BE Units when each driver shifts by -20%, -10%, 0%, +10%, and +20%."),
        ("Volume Scenario Forecasting:", "Projects outcomes under Pessimistic (-25% volume), Expected Base, and Optimistic (+25% volume) market conditions."),
        ("Executive Printable Report:", "Dedicated @media print layout produces clean, investor-ready PDF summaries with corporate letterheads."),
        ("Decision Support Warnings:", "Alerts decision-makers when fixed costs exceed safe coverage thresholds.")
    ]
    for title, desc in sims:
        p = tf6_r.add_paragraph()
        p.text = title
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = C_DARK
        p.space_before = Pt(6)

        p_b = tf6_r.add_paragraph()
        p_b.text = f"  {desc}"
        p_b.font.size = Pt(9.5)
        p_b.font.color.rgb = C_MUTED
        p_b.space_before = Pt(1)

    # --------------------------------------------------------------------------
    # SLIDE 7: Results, Deployment & Academic Conclusion
    # --------------------------------------------------------------------------
    slide7 = prs.slides.add_slide(blank_layout)
    add_header(slide7, "Project Results, Live Deployment & Conclusion", "CONCLUSION & IMPACT")

    c7_l = add_card(slide7, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.1), C_WHITE)
    tf7_l = c7_l.text_frame
    tf7_l.word_wrap = True
    tf7_l.margin_left = Inches(0.3)
    tf7_l.margin_top = Inches(0.3)

    h7_l = tf7_l.paragraphs[0]
    h7_l.text = "KEY DELIVERABLES & OUTCOMES"
    h7_l.font.size = Pt(12)
    h7_l.font.bold = True
    h7_l.font.color.rgb = C_PRIMARY

    outcomes = [
        "Live Production Deployment: Successfully deployed and hosted globally on GitHub Pages at https://arjun2007arush2011-max.github.io/break_even_point/",
        "Zero Third-Party Dependencies: Operates entirely standalone in any browser with zero node_modules or CDN downloads.",
        "High Performance: Sub-millisecond mathematical recalculation with smooth 60fps SVG rendering.",
        "Clean Print Reports: Built-in @media print executive feasibility report for stakeholders and investors.",
        "Search-Engine Ready: Integrated with robots.txt, sitemap.xml, and OpenGraph metadata."
    ]
    for o in outcomes:
        p = tf7_l.add_paragraph()
        p.text = f"✔ {o}"
        p.font.size = Pt(10.5)
        p.font.color.rgb = C_DARK
        p.space_before = Pt(7)

    c7_r = add_card(slide7, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.1), C_WHITE)
    tf7_r = c7_r.text_frame
    tf7_r.word_wrap = True
    tf7_r.margin_left = Inches(0.3)
    tf7_r.margin_top = Inches(0.3)

    h7_r = tf7_r.paragraphs[0]
    h7_r.text = "ACADEMIC IMPACT & FUTURE SCOPE"
    h7_r.font.size = Pt(12)
    h7_r.font.bold = True
    h7_r.font.color.rgb = C_SUCCESS

    academic = [
        ("Cross-Disciplinary Learning:", "Combines financial accounting, web engineering, database architecture, and computer graphics algorithms."),
        ("Multi-Product Inventory Mix:", "Future scope to support multi-item catalog weighting for complex retail operations."),
        ("Monte Carlo Stochastic Modeling:", "Adding probability distribution curves to model demand uncertainty."),
        ("Live Demonstration URL:", "https://arjun2007arush2011-max.github.io/break_even_point/")
    ]
    for title, desc in academic:
        p = tf7_r.add_paragraph()
        p.text = title
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = C_DARK
        p.space_before = Pt(6)

        p_b = tf7_r.add_paragraph()
        p_b.text = f"  {desc}"
        p_b.font.size = Pt(9.5)
        p_b.font.color.rgb = C_MUTED
        p_b.space_before = Pt(1)

    prs.save(PPTX_PATH)
    print(f"[OK] Enhanced PowerPoint (7 slides with unified Frontend & Backend) generated at: {PPTX_PATH}")


# ==============================================================================
# 2. PDF PRESENTATION GENERATION (MATCHING 7-SLIDE DECK)
# ==============================================================================
def create_pdf():
    doc = SimpleDocTemplate(
        PDF_PATH,
        pagesize=landscape(letter),
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36
    )

    styles = getSampleStyleSheet()

    style_cat = ParagraphStyle(
        'CatStyle', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=10, leading=12,
        textColor=colors.HexColor('#2563eb'), spaceAfter=3
    )
    style_title = ParagraphStyle(
        'TitleStyle', parent=styles['Heading1'],
        fontName='Helvetica-Bold', fontSize=18, leading=22,
        textColor=colors.HexColor('#0f172a'), spaceAfter=8
    )
    style_card_title = ParagraphStyle(
        'CardTStyle', parent=styles['Heading2'],
        fontName='Helvetica-Bold', fontSize=11, leading=14,
        textColor=colors.HexColor('#2563eb'), spaceAfter=4
    )
    style_card_title_green = ParagraphStyle('CGStyle', parent=style_card_title, textColor=colors.HexColor('#059669'))
    style_card_title_purple = ParagraphStyle('CPStyle', parent=style_card_title, textColor=colors.HexColor('#7c3aed'))
    style_card_title_red = ParagraphStyle('CRStyle', parent=style_card_title, textColor=colors.HexColor('#dc2626'))

    style_body = ParagraphStyle(
        'BodyDStyle', parent=styles['Normal'],
        fontName='Helvetica', fontSize=8.5, leading=12,
        textColor=colors.HexColor('#0f172a'), spaceAfter=3
    )
    style_body_muted = ParagraphStyle('BodyMStyle', parent=style_body, textColor=colors.HexColor('#64748b'))

    def make_card(flowables, width=340, bg_color='#f8fafc', border_color='#cbd5e1'):
        t = Table([[flowables]], colWidths=[width])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor(bg_color)),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor(border_color)),
            ('LEFTPADDING', (0,0), (-1,-1), 10),
            ('RIGHTPADDING', (0,0), (-1,-1), 10),
            ('TOPPADDING', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ]))
        return t

    story = []

    # Slide 1: Title & Overview (with Dashboard Screenshot)
    story.append(Paragraph("COLLEGE MINI PROJECT PRESENTATION", style_cat))
    story.append(Paragraph("Break-Even Point Predictor & Feasibility Engine", style_title))
    
    s1_left = [
        Paragraph("PROJECT METADATA", style_card_title),
        Paragraph("<b>Domain:</b> FinTech / Financial Engineering & Web Systems", style_body),
        Paragraph("<b>Developer:</b> arjun2007arush2011-max", style_body),
        Paragraph("<b>Architecture:</b> Zero-Dependency Vanilla ES5 + Multi-Runtime REST", style_body),
        Paragraph("<b>Database:</b> SQLite3 Relational DB + LocalStorage Fallback", style_body),
        Paragraph("<b>Live URL:</b> <font color='#2563eb'><u>https://arjun2007arush2011-max.github.io/break_even_point/</u></font>", style_body),
        Paragraph("<b>GitHub:</b> <font color='#2563eb'><u>https://github.com/arjun2007arush2011-max/break_even_point</u></font>", style_body)
    ]
    img1 = RLImage(SCREENSHOT_DASH, width=350, height=270) if os.path.exists(SCREENSHOT_DASH) else Paragraph("Screenshot", style_body)
    story.append(Table([[make_card(s1_left, width=340), img1]], colWidths=[360, 360], style=[('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
    story.append(PageBreak())

    # Slide 2: Problem Statement & Objectives
    story.append(Paragraph("BACKGROUND & MOTIVATION", style_cat))
    story.append(Paragraph("Problem Statement & Project Objectives", style_title))
    s2_l = [
        Paragraph("THE PROBLEM STATEMENT", style_card_title_red),
        Paragraph("• <b>High Failure Rate:</b> Over 65% of startup failures stem from premature cost miscalculations and hidden overhead.", style_body),
        Paragraph("• <b>Static Spreadsheet Errors:</b> Traditional Excel files lack interactive dynamic curves and mobile responsiveness.", style_body),
        Paragraph("• <b>Blind Pricing Decisions:</b> Entrepreneurs often underprice without analyzing variable cost compression on margins.", style_body),
        Paragraph("• <b>No Sensitivity Simulation:</b> Business owners struggle to forecast the impact of 15% cost inflation or rent spikes.", style_body)
    ]
    s2_r = [
        Paragraph("KEY PROJECT OBJECTIVES", style_card_title_green),
        Paragraph("✔ <b>Zero-Dependency Web Platform:</b> Completely self-contained vanilla ES5 implementation running without CDNs or frameworks.", style_body),
        Paragraph("✔ <b>Dynamic SVG Graphics:</b> Visual curves for Total Revenue, Total Cost, Fixed Costs, and profit/loss zones.", style_body),
        Paragraph("✔ <b>Real-Time What-If Engine:</b> Sliders (-50% to +50%) and sensitivity tables to test cost and price variations.", style_body),
        Paragraph("✔ <b>Multi-Runtime REST Integration:</b> Idiomatic backends in Python, Ruby, and Node.js with SQLite3 persistence.", style_body),
        Paragraph("✔ <b>Executive Feasibility Reports:</b> Formatted printable reports for stakeholders with company letterhead.", style_body)
    ]
    story.append(Table([[make_card(s2_l), make_card(s2_r)]], colWidths=[360, 360], style=[('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(PageBreak())

    # Slide 3: Full-Stack Architecture (Frontend and Backend on the SAME slide!)
    story.append(Paragraph("SYSTEM ARCHITECTURE", style_cat))
    story.append(Paragraph("Full-Stack System Architecture: Frontend & Backend", style_title))
    
    s3_front = [
        Paragraph("FRONTEND CLIENT TIER (Browser)", style_card_title),
        Paragraph("• <b>Vanilla ES5 Standard:</b> 100% universal compatibility across all browser engines without build steps or Babel.", style_body),
        Paragraph("• <b>Responsive Two-Column Grid:</b> Adapts smoothly from desktop to mobile screens.", style_body),
        Paragraph("• <b>Dark / Light Theming:</b> Driven by CSS Custom Properties with localStorage persistence.", style_body),
        Paragraph("• <b>Client-Side Crypto:</b> Passwords hashed via crypto.subtle (SHA-256) prior to storage.", style_body),
        Paragraph("• <b>Modular Namespaces:</b> AuthModule, ProfileModule, MetricsModule, ChartModule, ApiClient.", style_body)
    ]
    s3_back = [
        Paragraph("BACKEND SERVER & DATABASE TIER", style_card_title_green),
        Paragraph("• <b>Triple-Runtime Servers:</b> Python 3 (http.server), Ruby (WEBrick), and Node.js. Zero pip/npm dependencies.", style_body),
        Paragraph("• <b>Universal Launcher (start.sh):</b> Auto-detects installed runtime and binds port 8000.", style_body),
        Paragraph("• <b>SQLite3 Database (breakeven.db):</b> Tables for Users, Profiles, and Cost Inputs.", style_body),
        Paragraph("• <b>REST API Contracts:</b> Endpoints for /api/auth/*, /api/profile, /api/inputs, /api/calculate.", style_body),
        Paragraph("• <b>Server Mathematical Verification:</b> Confirms calculation parity with client engine.", style_body)
    ]
    s3_bridge = [
        Paragraph("COMMUNICATION PROTOCOL & RESILIENT DUAL-SYNC BRIDGE (Frontend ⟷ Backend)", style_card_title_purple),
        Paragraph("• <b>Debounced REST API Sync:</b> Client debounces user inputs (300ms) and sends JSON payloads with Bearer token authentication to /api/inputs and /api/profile.", style_body),
        Paragraph("• <b>Dual Persistence Resilience:</b> ApiClient pings /api/health. When server is active, records persist to SQLite3; if offline or on GitHub Pages (file://), it gracefully falls back to localStorage with zero downtime.", style_body)
    ]
    
    top_table = Table([[make_card(s3_front, width=340), make_card(s3_back, width=340)]], colWidths=[360, 360], style=[('VALIGN', (0,0), (-1,-1), 'TOP')])
    bot_table = Table([[make_card(s3_bridge, width=700, bg_color='#f8fafc', border_color='#7c3aed')]], colWidths=[720], style=[('VALIGN', (0,0), (-1,-1), 'TOP')])
    
    story.append(top_table)
    story.append(Spacer(1, 6))
    story.append(bot_table)
    story.append(PageBreak())

    # Slide 4: Interactive SVG Chart
    story.append(Paragraph("VISUALIZATION ENGINE", style_cat))
    story.append(Paragraph("Interactive SVG Chart & Visual Capabilities", style_title))
    img_chart = RLImage(SCREENSHOT_CHART, width=350, height=270) if os.path.exists(SCREENSHOT_CHART) else Paragraph("Chart UI", style_body)
    s4_r = [
        Paragraph("CUSTOM COORDINATE MAPPING ENGINE", style_card_title_purple),
        Paragraph("• <b>Zero External Chart Libraries:</b> Custom SVG mathematical projection avoiding heavy Chart.js or D3.js libraries.", style_body),
        Paragraph("• <b>Loss & Profit Polygons:</b> Fills red loss zone (0 to BE) and green profit zone (BE to Max Volume) dynamically.", style_body),
        Paragraph("• <b>Break-Even Point Marker:</b> Displays intersection point with a pulsing halo and coordinate callout badge.", style_body),
        Paragraph("• <b>Expected Plan Line:</b> Vertical dashed indicator showing planned sales volume versus break-even.", style_body),
        Paragraph("• <b>Touch & Hover Crosshairs:</b> Real-time cursor reading volume, revenue, costs, and profit at 60fps.", style_body)
    ]
    story.append(Table([[img_chart, make_card(s4_r)]], colWidths=[360, 360], style=[('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
    story.append(PageBreak())

    # Slide 5: Mathematical Formulation & Financial Algorithms
    story.append(Paragraph("FINANCIAL MODEL & ALGORITHMS", style_cat))
    story.append(Paragraph("Mathematical Formulation & Financial Algorithms", style_title))
    s5_l = [
        Paragraph("CORE ACCOUNTING FORMULAS", style_card_title),
        Paragraph("• <b>Contribution Margin (CM):</b> CM = Selling Price (P) - Variable Cost (V)", style_body),
        Paragraph("• <b>Contribution Margin Ratio:</b> CM% = (P - V) / P", style_body),
        Paragraph("• <b>Break-Even Volume:</b> Q_BE = Fixed Costs (F) / (P - V)", style_body),
        Paragraph("• <b>Break-Even Revenue:</b> R_BE = Q_BE × P = F / CM%", style_body),
        Paragraph("• <b>Net Operating Profit:</b> Profit = Q × (P - V) - F", style_body),
        Paragraph("• <b>Margin of Safety (MOS):</b> MOS = (Q - Q_BE) / Q × 100%", style_body),
        Paragraph("• <b>Operating Leverage (DOL):</b> DOL = (Q × CM) / Profit", style_body),
        Paragraph("• <b>Target Profit Units:</b> Q_Target = (F + Target) / CM", style_body)
    ]
    s5_r = [
        Paragraph("ALGORITHMIC GUARDS & INTEGRITY", style_card_title_green),
        Paragraph("✔ <b>Structural Loss Detection (P ≤ V):</b> Instantly alerts if price is less than or equal to variable cost. Prevents division-by-zero exceptions.", style_body),
        Paragraph("✔ <b>Sub-Break-Even Operating Alert:</b> Dynamically signals financial danger when planned sales volume Q < Q_BE.", style_body),
        Paragraph("✔ <b>Floating-Point Normalization:</b> Sanitizes user numeric inputs and eliminates IEEE 754 floating-point inaccuracies.", style_body),
        Paragraph("✔ <b>Server Verification Parity:</b> Backend /api/calculate endpoint independently validates mathematical formulas.", style_body)
    ]
    story.append(Table([[make_card(s5_l), make_card(s5_r)]], colWidths=[360, 360], style=[('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(PageBreak())

    # Slide 6: What-If Sensitivity Simulations & Scenario Planning
    story.append(Paragraph("SIMULATION & SENSITIVITY", style_cat))
    story.append(Paragraph("What-If Sensitivity Simulations & Scenario Planning", style_title))
    s6_l = [
        Paragraph("REAL-TIME WHAT-IF SLIDERS", style_card_title),
        Paragraph("• <b>Bidirectional Sliders (-50% to +50%):</b> Instant adjustment of Price, Variable Cost, and Fixed Costs.", style_body),
        Paragraph("• <b>Live Variance Badges:</b> Computes real-time percentage shifts and monetary changes against base parameters.", style_body),
        Paragraph("• <b>Real-Time SVG Morphing:</b> Updates coordinate curves at 60fps, shifting the break-even intersection point dynamically.", style_body),
        Paragraph("• <b>Single-Click Baseline Reset:</b> Restores saved base parameters without page reloads.", style_body)
    ]
    s6_r = [
        Paragraph("STRESS TESTING & REPORT GENERATION", style_card_title_green),
        Paragraph("• <b>5-Tier Sensitivity Matrix:</b> Evaluates BE Volume shifts across -20%, -10%, 0%, +10%, and +20% cost driver changes.", style_body),
        Paragraph("• <b>Volume Scenario Analysis:</b> Models Pessimistic (-25%), Base, and Optimistic (+25%) market outcomes.", style_body),
        Paragraph("• <b>Executive Printable Report:</b> Dedicated @media print stylesheet for clean, publication-ready PDF exports.", style_body),
        Paragraph("• <b>Decision Support Warnings:</b> Highlights when fixed costs exceed safe coverage thresholds.", style_body)
    ]
    story.append(Table([[make_card(s6_l), make_card(s6_r)]], colWidths=[360, 360], style=[('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(PageBreak())

    # Slide 7: Results & Academic Conclusion
    story.append(Paragraph("CONCLUSION & IMPACT", style_cat))
    story.append(Paragraph("Project Results, Live Deployment & Academic Scope", style_title))
    s7_l = [
        Paragraph("KEY DELIVERABLES & OUTCOMES", style_card_title),
        Paragraph("✔ <b>Live Production Deployment:</b> Hosted globally on GitHub Pages at <font color='#2563eb'><u>https://arjun2007arush2011-max.github.io/break_even_point/</u></font>", style_body),
        Paragraph("✔ <b>Zero Dependencies:</b> Operates without npm packages, external build tools, or CDN dependencies.", style_body),
        Paragraph("✔ <b>High Performance:</b> Sub-millisecond mathematical calculations and smooth 60fps SVG chart recalculations.", style_body),
        Paragraph("✔ <b>Print Feasibility Report:</b> Clean @media print executive feasibility report for stakeholders.", style_body),
        Paragraph("✔ <b>SEO Optimization:</b> Integrated robots.txt, sitemap.xml, and OpenGraph social metadata.", style_body)
    ]
    s7_r = [
        Paragraph("ACADEMIC IMPACT & FUTURE SCOPE", style_card_title_green),
        Paragraph("• <b>Cross-Disciplinary Integration:</b> Merges software engineering, managerial cost accounting, database design, and interactive computer graphics.", style_body),
        Paragraph("• <b>Multi-Product Sales Mix:</b> Future expansion to calculate weighted average contribution margins for multi-item inventory catalogs.", style_body),
        Paragraph("• <b>Monte Carlo Risk Modeling:</b> Simulating demand curve fluctuations and price elasticity for probabilistic risk assessment.", style_body),
        Paragraph("• <b>GitHub Repository:</b> <font color='#2563eb'><u>https://github.com/arjun2007arush2011-max/break_even_point</u></font>", style_body)
    ]
    story.append(Table([[make_card(s7_l), make_card(s7_r)]], colWidths=[360, 360], style=[('VALIGN', (0,0), (-1,-1), 'TOP')]))

    doc.build(story)
    print(f"[OK] Enhanced PDF (7 slides with unified Frontend & Backend) generated at: {PDF_PATH}")


if __name__ == "__main__":
    create_pptx()
    create_pdf()
