#!/usr/bin/env python3
"""
Generates PowerPoint (PPTX) and PDF presentation slides with embedded screenshots
and dedicated Frontend & Backend architecture slides for the college mini-project.
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
# 1. PPTX GENERATION WITH SCREENSHOTS & SPECIALIZED SLIDES
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
    # SLIDE 1: Title Slide (With UI Preview)
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

    # Left metadata card
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

    # Right side: Screenshot of the live app
    if os.path.exists(SCREENSHOT_DASH):
        slide1.shapes.add_picture(SCREENSHOT_DASH, Inches(7.3), Inches(1.4), Inches(4.8), Inches(4.8))

    # --------------------------------------------------------------------------
    # SLIDE 2: Problem Statement & Objectives
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
    # SLIDE 3: Frontend Architecture & User Interface (With Screenshot)
    # --------------------------------------------------------------------------
    slide3 = prs.slides.add_slide(blank_layout)
    add_header(slide3, "Frontend Architecture & User Interface", "FRONTEND TIER")

    # Left: Explanation
    c3_l = add_card(slide3, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.1), C_WHITE)
    tf3_l = c3_l.text_frame
    tf3_l.word_wrap = True
    tf3_l.margin_left = Inches(0.3)
    tf3_l.margin_top = Inches(0.3)

    h3_l = tf3_l.paragraphs[0]
    h3_l.text = "FRONTEND DESIGN & USER EXPERIENCE"
    h3_l.font.size = Pt(12)
    h3_l.font.bold = True
    h3_l.font.color.rgb = C_PRIMARY

    f_points = [
        ("Vanilla ES5 Architecture:", "Strict ES5 standards ensuring 100% universal browser compatibility without Babel or bundling overhead."),
        ("Responsive Two-Column Layout:", "Adapts automatically from a two-column desktop grid to an optimized single-column mobile view."),
        ("Dark & Light Mode:", "Managed via CSS custom properties with persistent preference saved in localStorage."),
        ("Client-Side SHA-256 Hashing:", "Passwords hashed via browser Web Crypto API (crypto.subtle) before storage with pure JS fallback."),
        ("Modular Namespaces:", "Clear separation of concerns across AuthModule, ProfileModule, MetricsModule, ChartModule, and ApiClient.")
    ]
    for title, desc in f_points:
        p = tf3_l.add_paragraph()
        p.text = title
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = C_DARK
        p.space_before = Pt(6)

        p_b = tf3_l.add_paragraph()
        p_b.text = f"  {desc}"
        p_b.font.size = Pt(9.5)
        p_b.font.color.rgb = C_MUTED
        p_b.space_before = Pt(1)

    # Right: Login & UI Screenshot
    if os.path.exists(SCREENSHOT_LOGIN):
        slide3.shapes.add_picture(SCREENSHOT_LOGIN, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.1))

    # --------------------------------------------------------------------------
    # SLIDE 4: Interactive SVG Chart & Visualization Engine (With Chart Screenshot)
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
    # SLIDE 5: Backend Architecture & REST API (Dedicated Backend Slide)
    # --------------------------------------------------------------------------
    slide5 = prs.slides.add_slide(blank_layout)
    add_header(slide5, "Backend Architecture, REST APIs & SQLite3", "BACKEND & DATABASE")

    c5_l = add_card(slide5, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.1), C_WHITE)
    tf5_l = c5_l.text_frame
    tf5_l.word_wrap = True
    tf5_l.margin_left = Inches(0.3)
    tf5_l.margin_top = Inches(0.3)

    h5_l = tf5_l.paragraphs[0]
    h5_l.text = "TRIPLE-RUNTIME BACKEND SERVERS"
    h5_l.font.size = Pt(12)
    h5_l.font.bold = True
    h5_l.font.color.rgb = C_PRIMARY

    b_pts = [
        ("Python 3 Server (server.py):", "Built with native http.server and sqlite3. Zero pip packages required."),
        ("Ruby Server (server.rb):", "Built with WEBrick and native sqlite3 gem. Zero bundler/gem installs needed."),
        ("Node.js Server (server.js):", "Built with standard http, crypto, and fs modules."),
        ("Universal Launcher (start.sh):", "Auto-detects which runtime is installed on the host and starts on port 8000."),
        ("Relational SQLite Database (breakeven.db):", "Structured schemas for Users (with SHA-256 tokens), Business Profiles, and Cost Inputs.")
    ]
    for title, desc in b_pts:
        p = tf5_l.add_paragraph()
        p.text = title
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = C_DARK
        p.space_before = Pt(6)

        p_b = tf5_l.add_paragraph()
        p_b.text = f"  {desc}"
        p_b.font.size = Pt(9.5)
        p_b.font.color.rgb = C_MUTED
        p_b.space_before = Pt(1)

    c5_r = add_card(slide5, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.1), C_WHITE)
    tf5_r = c5_r.text_frame
    tf5_r.word_wrap = True
    tf5_r.margin_left = Inches(0.3)
    tf5_r.margin_top = Inches(0.3)

    h5_r = tf5_r.paragraphs[0]
    h5_r.text = "REST API CONTRACT & PERSISTENCE FLOW"
    h5_r.font.size = Pt(12)
    h5_r.font.bold = True
    h5_r.font.color.rgb = C_SUCCESS

    apis = [
        ("POST /api/auth/signup", "Registers user, hashes password, seeds default profile & inputs."),
        ("POST /api/auth/login", "Authenticates credentials and returns a 24-byte bearer token."),
        ("GET /api/auth/me", "Validates session token and returns user profile & inputs."),
        ("GET & POST /api/profile", "Retrieves and updates business metadata in SQLite3."),
        ("GET & POST /api/inputs", "Persists cost drivers with debounced client auto-saving."),
        ("POST /api/calculate", "Server-side verification of mathematical break-even model.")
    ]
    for ep, desc in apis:
        p = tf5_r.add_paragraph()
        p.text = f"• {ep}: {desc}"
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_DARK
        p.space_before = Pt(5)

    p_d = tf5_r.add_paragraph()
    p_d.text = "Dual-Engine Resilience Logic:"
    p_d.font.size = Pt(10.5)
    p_d.font.bold = True
    p_d.font.color.rgb = C_ACCENT
    p_d.space_before = Pt(8)

    p_dd = tf5_r.add_paragraph()
    p_dd.text = "The ApiClient probes /api/health. If connected, data writes to SQLite3; if the server is offline or opened as file://, it gracefully falls back to localStorage."
    p_dd.font.size = Pt(9)
    p_dd.font.color.rgb = C_MUTED
    p_dd.space_before = Pt(1)

    # --------------------------------------------------------------------------
    # SLIDE 6: Mathematical Model & What-If Simulations
    # --------------------------------------------------------------------------
    slide6 = prs.slides.add_slide(blank_layout)
    add_header(slide6, "Mathematical Model & What-If Simulations", "ANALYTICS ENGINE")

    c6_l = add_card(slide6, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.1), C_WHITE)
    tf6_l = c6_l.text_frame
    tf6_l.word_wrap = True
    tf6_l.margin_left = Inches(0.3)
    tf6_l.margin_top = Inches(0.3)

    h6_l = tf6_l.paragraphs[0]
    h6_l.text = "FINANCIAL MODEL EQUATIONS"
    h6_l.font.size = Pt(12)
    h6_l.font.bold = True
    h6_l.font.color.rgb = C_PRIMARY

    math_eqs = [
        ("Contribution Margin (CM):", "CM = Selling Price (P) - Variable Cost (V)"),
        ("Contribution Margin Ratio:", "CM% = (P - V) / P"),
        ("Break-Even Volume (Units):", "Q_BE = Fixed Costs (F) / (P - V)"),
        ("Break-Even Revenue:", "R_BE = Q_BE × P = F / CM%"),
        ("Net Operating Profit:", "Profit = Q × (P - V) - F"),
        ("Margin of Safety (MOS):", "MOS = (Q - Q_BE) / Q × 100%"),
        ("Degree of Operating Leverage:", "DOL = (Q × CM) / Profit"),
        ("Target Profit Units:", "Q_Target = (F + Target) / CM")
    ]
    for title, formula in math_eqs:
        p = tf6_l.add_paragraph()
        p.text = f"{title} {formula}"
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_DARK
        p.space_before = Pt(4)

    c6_r = add_card(slide6, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.1), C_WHITE)
    tf6_r = c6_r.text_frame
    tf6_r.word_wrap = True
    tf6_r.margin_left = Inches(0.3)
    tf6_r.margin_top = Inches(0.3)

    h6_r = tf6_r.paragraphs[0]
    h6_r.text = "WHAT-IF SIMULATIONS & STRESS MATRICES"
    h6_r.font.size = Pt(12)
    h6_r.font.bold = True
    h6_r.font.color.rgb = C_SUCCESS

    sims = [
        ("Real-Time What-If Sliders:", "Dual-direction sliders (-50% to +50%) for Price, Var Cost, and Fixed Costs compute live variances versus baseline."),
        ("Sensitivity Matrix Table:", "Evaluates Required BE Units when each driver shifts by -20%, -10%, 0%, +10%, and +20%."),
        ("Volume Scenario Table:", "Projects outcomes under Pessimistic (-25% volume), Expected Base, and Optimistic (+25% volume) market conditions."),
        ("Automated Feasibility Guards:", "Immediately detects and warns if Price ≤ Variable Cost (structural loss condition) or if planned volume Q < Q_BE.")
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
    print(f"[OK] Enhanced PowerPoint with screenshots generated at: {PPTX_PATH}")


# ==============================================================================
# 2. PDF PRESENTATION GENERATION WITH EMBEDDED SCREENSHOTS
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
        fontName='Helvetica-Bold', fontSize=19, leading=23,
        textColor=colors.HexColor('#0f172a'), spaceAfter=10
    )
    style_card_title = ParagraphStyle(
        'CardTStyle', parent=styles['Heading2'],
        fontName='Helvetica-Bold', fontSize=11, leading=14,
        textColor=colors.HexColor('#2563eb'), spaceAfter=5
    )
    style_card_title_green = ParagraphStyle('CGStyle', parent=style_card_title, textColor=colors.HexColor('#059669'))
    style_card_title_purple = ParagraphStyle('CPStyle', parent=style_card_title, textColor=colors.HexColor('#7c3aed'))
    style_card_title_red = ParagraphStyle('CRStyle', parent=style_card_title, textColor=colors.HexColor('#dc2626'))

    style_body = ParagraphStyle(
        'BodyDStyle', parent=styles['Normal'],
        fontName='Helvetica', fontSize=9, leading=13,
        textColor=colors.HexColor('#0f172a'), spaceAfter=4
    )
    style_body_muted = ParagraphStyle('BodyMStyle', parent=style_body, textColor=colors.HexColor('#64748b'))

    def make_card(flowables, width=340, bg_color='#f8fafc', border_color='#cbd5e1'):
        t = Table([[flowables]], colWidths=[width])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor(bg_color)),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor(border_color)),
            ('LEFTPADDING', (0,0), (-1,-1), 12),
            ('RIGHTPADDING', (0,0), (-1,-1), 12),
            ('TOPPADDING', (0,0), (-1,-1), 10),
            ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ]))
        return t

    story = []

    # Slide 1: Title (with Dashboard Screenshot)
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

    # Slide 3: Frontend Architecture & UI Screenshot
    story.append(Paragraph("FRONTEND TIER", style_cat))
    story.append(Paragraph("Frontend Architecture & User Interface Design", style_title))
    s3_l = [
        Paragraph("UI DESIGN & ARCHITECTURE", style_card_title),
        Paragraph("• <b>Vanilla ES5 Standard:</b> Ensures 100% universal browser compatibility without build pipelines.", style_body),
        Paragraph("• <b>Responsive Two-Column Layout:</b> Smoothly adapts between desktop and mobile screens.", style_body),
        Paragraph("• <b>Dark / Light Themes:</b> Powered by CSS Custom Properties with persistent storage.", style_body),
        Paragraph("• <b>Client-Side SHA-256 Hashing:</b> Passwords hashed via crypto.subtle before storage or transmission.", style_body),
        Paragraph("• <b>Clean Modular Objects:</b> AuthModule, ProfileModule, MetricsModule, ChartModule, and ApiClient.", style_body)
    ]
    img_login = RLImage(SCREENSHOT_LOGIN, width=340, height=270) if os.path.exists(SCREENSHOT_LOGIN) else Paragraph("Login UI", style_body)
    story.append(Table([[make_card(s3_l), img_login]], colWidths=[360, 360], style=[('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
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

    # Slide 5: Backend Architecture (Dedicated Backend Slide)
    story.append(Paragraph("BACKEND & DATABASE", style_cat))
    story.append(Paragraph("Backend Architecture, REST APIs & SQLite3 Storage", style_title))
    s5_l = [
        Paragraph("TRIPLE RUNTIME BACKEND IMPLEMENTATION", style_card_title),
        Paragraph("• <b>Python 3 Backend (server.py):</b> Built with native http.server and sqlite3. Zero external pip packages.", style_body),
        Paragraph("• <b>Ruby Backend (server.rb):</b> Built with WEBrick and native sqlite3 gem. Zero external dependencies.", style_body),
        Paragraph("• <b>Node.js Backend (server.js):</b> Native http and crypto standard library.", style_body),
        Paragraph("• <b>Auto-Launcher (start.sh):</b> Automatically detects available runtime and launches on port 8000.", style_body),
        Paragraph("• <b>SQLite Database (breakeven.db):</b> Tables for Users, Profiles, and Cost Inputs.", style_body)
    ]
    s5_r = [
        Paragraph("REST API CONTRACT & DUAL SYNC", style_card_title_green),
        Paragraph("• <b>POST /api/auth/signup & login:</b> SHA-256 auth returning 24-byte bearer tokens.", style_body),
        Paragraph("• <b>GET /api/auth/me:</b> Validates bearer session and returns profile & inputs.", style_body),
        Paragraph("• <b>GET & POST /api/profile:</b> SQLite3 persistence for business letterhead details.", style_body),
        Paragraph("• <b>GET & POST /api/inputs:</b> Persists cost drivers with debounced auto-saving.", style_body),
        Paragraph("• <b>POST /api/calculate:</b> Server-side model calculation verification.", style_body),
        Paragraph("• <b>Dual-Engine Synchronization:</b> Probes /api/health; if online, syncs with SQLite3; if offline, operates via localStorage seamlessly.", style_body)
    ]
    story.append(Table([[make_card(s5_l), make_card(s5_r)]], colWidths=[360, 360], style=[('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(PageBreak())

    # Slide 6: Mathematical Model & Simulations
    story.append(Paragraph("ANALYTICS ENGINE", style_cat))
    story.append(Paragraph("Mathematical Model & What-If Simulations", style_title))
    s6_l = [
        Paragraph("CORE ACCOUNTING FORMULAS", style_card_title),
        Paragraph("• <b>Contribution Margin (CM):</b> CM = P - V", style_body),
        Paragraph("• <b>Contribution Margin Ratio:</b> CM% = (P - V) / P", style_body),
        Paragraph("• <b>Break-Even Volume:</b> Q_BE = F / (P - V)", style_body),
        Paragraph("• <b>Break-Even Revenue:</b> R_BE = Q_BE × P", style_body),
        Paragraph("• <b>Net Operating Profit:</b> Profit = Q × (P - V) - F", style_body),
        Paragraph("• <b>Margin of Safety:</b> MOS = (Q - Q_BE) / Q × 100%", style_body),
        Paragraph("• <b>Operating Leverage:</b> DOL = (Q × CM) / Profit", style_body),
        Paragraph("• <b>Target Profit Units:</b> Q_Target = (F + Target) / CM", style_body)
    ]
    s6_r = [
        Paragraph("WHAT-IF SLIDERS & SCENARIO MATRICES", style_card_title_green),
        Paragraph("• <b>Real-Time What-If Sliders:</b> Adjust Price, Var Cost, and Fixed Costs (-50% to +50%) with live variance indicators.", style_body),
        Paragraph("• <b>Sensitivity Matrix:</b> Displays BE Volume shifts at -20%, -10%, 0%, +10%, +20% for each cost driver.", style_body),
        Paragraph("• <b>Volume Scenario Analysis:</b> Evaluates Pessimistic (-25%), Base, and Optimistic (+25%) market conditions.", style_body),
        Paragraph("• <b>Structural Loss Guard:</b> Raises instant alert if P ≤ V (selling below variable cost) with full loss shading.", style_body)
    ]
    story.append(Table([[make_card(s6_l), make_card(s6_r)]], colWidths=[360, 360], style=[('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(PageBreak())

    # Slide 7: Results & Conclusion
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
    print(f"[OK] Enhanced PDF with screenshots generated at: {PDF_PATH}")


if __name__ == "__main__":
    create_pptx()
    create_pdf()
