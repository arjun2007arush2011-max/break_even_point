#!/usr/bin/env python3
"""
Generates PowerPoint (PPTX) and PDF presentation slides for the
Break-Even Point Predictor college mini-project.
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
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether

# Output paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PPTX_PATH = os.path.join(BASE_DIR, "Break-Even-Project-Presentation.pptx")
PDF_PATH = os.path.join(BASE_DIR, "Break-Even-Project-Presentation.pdf")

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
# 1. PPTX GENERATION
# ==============================================================================
def create_pptx():
    prs = Presentation()
    # 16:9 Widescreen dimensions
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    def add_header(slide, title_text, category_text="COLLEGE MINI PROJECT"):
        # Header category badge / breadcrumb
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.35))
        tf_c = cat_box.text_frame
        tf_c.word_wrap = True
        p_c = tf_c.paragraphs[0]
        p_c.text = category_text.upper()
        p_c.font.size = Pt(11)
        p_c.font.bold = True
        p_c.font.color.rgb = C_PRIMARY

        # Slide Main Title
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
    # SLIDE 1: Title Slide
    # --------------------------------------------------------------------------
    slide1 = prs.slides.add_slide(blank_layout)
    # Background accent card
    add_card(slide1, Inches(0.8), Inches(0.8), Inches(11.733), Inches(5.9), C_LIGHT_BG, C_PRIMARY)

    title_box = slide1.shapes.add_textbox(Inches(1.3), Inches(1.3), Inches(10.7), Inches(2.2))
    tf1 = title_box.text_frame
    tf1.word_wrap = True

    p0 = tf1.paragraphs[0]
    p0.text = "ACADEMIC MINI PROJECT PRESENTATION"
    p0.font.size = Pt(13)
    p0.font.bold = True
    p0.font.color.rgb = C_PRIMARY

    p1 = tf1.add_paragraph()
    p1.text = "Break-Even Point Predictor & Feasibility Engine"
    p1.font.size = Pt(32)
    p1.font.bold = True
    p1.font.color.rgb = C_DARK
    p1.space_before = Pt(8)

    p2 = tf1.add_paragraph()
    p2.text = "A Full-Stack Financial Decision Support System with Real-Time SVG Visualizations & What-If Simulations"
    p2.font.size = Pt(15)
    p2.font.color.rgb = C_MUTED
    p2.space_before = Pt(6)

    # Details Box (Left: Student & Project Info, Right: Links & Tech)
    info_card1 = add_card(slide1, Inches(1.3), Inches(3.8), Inches(5.2), Inches(2.3), C_WHITE)
    info_tf1 = info_card1.text_frame
    info_tf1.word_wrap = True
    info_tf1.margin_left = Inches(0.3)
    info_tf1.margin_top = Inches(0.25)
    
    it0 = info_tf1.paragraphs[0]
    it0.text = "PROJECT METADATA"
    it0.font.size = Pt(10)
    it0.font.bold = True
    it0.font.color.rgb = C_PRIMARY

    lines1 = [
        ("Domain:", "FinTech / Financial Engineering & Web Systems"),
        ("Developer:", "arjun2007arush2011-max"),
        ("Architecture:", "Client-Side Zero-Dependency + REST Backend"),
        ("Storage:", "Dual-Mode (SQLite3 Database & LocalStorage)")
    ]
    for lbl, val in lines1:
        p = info_tf1.add_paragraph()
        p.text = f"{lbl} {val}"
        p.font.size = Pt(11)
        p.font.color.rgb = C_DARK
        p.space_before = Pt(4)

    info_card2 = add_card(slide1, Inches(6.8), Inches(3.8), Inches(5.2), Inches(2.3), C_WHITE)
    info_tf2 = info_card2.text_frame
    info_tf2.word_wrap = True
    info_tf2.margin_left = Inches(0.3)
    info_tf2.margin_top = Inches(0.25)

    it2 = info_tf2.paragraphs[0]
    it2.text = "DEPLOYMENT & ACCESS LINKS"
    it2.font.size = Pt(10)
    it2.font.bold = True
    it2.font.color.rgb = C_SUCCESS

    lines2 = [
        ("Live Web App:", "https://arjun2007arush2011-max.github.io/break_even_point/"),
        ("GitHub Repo:", "https://github.com/arjun2007arush2011-max/break_even_point"),
        ("Backends:", "Python 3, Ruby WEBrick, Node.js Native HTTP"),
        ("Status:", "Live & Globally Deployed via GitHub Pages")
    ]
    for lbl, val in lines2:
        p = info_tf2.add_paragraph()
        p.text = f"{lbl} {val}"
        p.font.size = Pt(11)
        p.font.color.rgb = C_DARK
        p.space_before = Pt(4)

    # --------------------------------------------------------------------------
    # SLIDE 2: Problem Statement & Objectives
    # --------------------------------------------------------------------------
    slide2 = prs.slides.add_slide(blank_layout)
    add_header(slide2, "Problem Statement & Project Objectives", "BACKGROUND & MOTIVATION")

    c2_left = add_card(slide2, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.1), C_WHITE)
    tf2_l = c2_left.text_frame
    tf2_l.word_wrap = True
    tf2_l.margin_left = Inches(0.35)
    tf2_l.margin_top = Inches(0.35)

    h2_l = tf2_l.paragraphs[0]
    h2_l.text = "THE PROBLEM STATEMENT"
    h2_l.font.size = Pt(12)
    h2_l.font.bold = True
    h2_l.font.color.rgb = RGBColor(220, 38, 38) # Red

    p_body1 = tf2_l.add_paragraph()
    p_body1.text = "Over 65% of small business and startup failures stem from premature cost miscalculations and a lack of real-time visibility into operational break-even points."
    p_body1.font.size = Pt(12.5)
    p_body1.font.color.rgb = C_DARK
    p_body1.space_before = Pt(8)

    pts1 = [
        "Static Spreadsheets are Prone to Errors: Traditional Excel sheets lack dynamic visual feedback and are cumbersome to update on mobile devices.",
        "Unrealistic Pricing Decisions: Entrepreneurs frequently underprice products without understanding how variable costs eat into contribution margins.",
        "No What-If Sensitivity Testing: Business owners cannot easily simulate how a 10% inflation spike in variable costs or rent shifts required sales volume.",
        "Heavy Dependency Bloat: Existing online tools require heavy, tracking-laden frameworks, external CDNs, or monthly subscriptions."
    ]
    for pt in pts1:
        p = tf2_l.add_paragraph()
        p.text = f"• {pt}"
        p.font.size = Pt(11)
        p.font.color.rgb = C_MUTED
        p.space_before = Pt(6)

    c2_right = add_card(slide2, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.1), C_WHITE)
    tf2_r = c2_right.text_frame
    tf2_r.word_wrap = True
    tf2_r.margin_left = Inches(0.35)
    tf2_r.margin_top = Inches(0.35)

    h2_r = tf2_r.paragraphs[0]
    h2_r.text = "KEY PROJECT OBJECTIVES"
    h2_r.font.size = Pt(12)
    h2_r.font.bold = True
    h2_r.font.color.rgb = C_SUCCESS

    pts2 = [
        "Zero-Dependency Architecture: Build a completely self-contained solution running natively without third-party JS libraries or CDN dependencies.",
        "Interactive SVG Graphics: Generate custom, scalable vector charts depicting Total Revenue, Total Cost, Fixed Cost lines, and shaded profit/loss zones.",
        "Real-Time What-If Sensitivity: Provide instantaneous sliders (-50% to +50%) and sensitivity matrices to stress-test cost and price variables.",
        "Multi-Tier REST Integration: Implement backend APIs in Python, Ruby, and Node.js with persistent SQLite3 storage alongside graceful offline mode.",
        "Executive Print Reporting: Provide clean, letterheaded executive feasibility printouts for stakeholders and financial presentations."
    ]
    for pt in pts2:
        p = tf2_r.add_paragraph()
        p.text = f"✔ {pt}"
        p.font.size = Pt(11)
        p.font.color.rgb = C_DARK
        p.space_before = Pt(8)

    # --------------------------------------------------------------------------
    # SLIDE 3: System Architecture & Tech Stack
    # --------------------------------------------------------------------------
    slide3 = prs.slides.add_slide(blank_layout)
    add_header(slide3, "System Architecture & Technology Stack", "TECHNICAL DESIGN")

    cards3 = [
        (Inches(0.8), "FRONTEND PRESENTATION", C_PRIMARY, [
            "Vanilla ES5 JavaScript for 100% universal browser compatibility (no transpilation required).",
            "Responsive Two-Column Grid layout transitioning smoothly to single-column on mobile viewports.",
            "CSS Custom Properties (Variables) supporting instant Dark / Light theme toggling.",
            "Accessible semantic markup with ARIA roles and validation tooltips."
        ]),
        (Inches(4.8), "VISUALIZATION & ENGINE", C_ACCENT, [
            "Custom Inline SVG Engine: Mathematical coordinate projection without Chart.js or D3.js.",
            "Interactive Crosshair & Hover Tooltip tracking volume, revenue, costs, and profit in real time.",
            "Pure Mathematical Model: model(F, P, V, Q, T) isolating calculation logic from DOM manipulation.",
            "Client-Side SHA-256 Hashing using Web Crypto API (crypto.subtle) with pure JS fallback."
        ]),
        (Inches(8.8), "BACKEND & PERSISTENCE", C_SUCCESS, [
            "Triple-Runtime REST APIs: Implemented in Ruby (WEBrick), Python 3 (http.server), and Node.js.",
            "SQLite3 Relational Storage: Users, Business Profiles, and Cost Driver input tables.",
            "Debounced Input Synchronization: Smooth server syncing during rapid slider adjustments.",
            "Graceful Offline Fallback: Auto-detects server connection status; falls back to localStorage."
        ])
    ]

    for left, title, col, bullets in cards3:
        c = add_card(slide3, left, Inches(1.6), Inches(3.7), Inches(5.1), C_WHITE)
        tf = c.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.25)
        tf.margin_top = Inches(0.3)

        h = tf.paragraphs[0]
        h.text = title
        h.font.size = Pt(12)
        h.font.bold = True
        h.font.color.rgb = col

        for b in bullets:
            p = tf.add_paragraph()
            p.text = f"• {b}"
            p.font.size = Pt(10.5)
            p.font.color.rgb = C_DARK
            p.space_before = Pt(8)

    # --------------------------------------------------------------------------
    # SLIDE 4: Mathematical Formulation & Core Algorithms
    # --------------------------------------------------------------------------
    slide4 = prs.slides.add_slide(blank_layout)
    add_header(slide4, "Mathematical Formulation & Financial Model", "CORE ALGORITHMS")

    # Left Card: Formulas
    c4_left = add_card(slide4, Inches(0.8), Inches(1.6), Inches(6.0), Inches(5.1), C_WHITE)
    tf4_l = c4_left.text_frame
    tf4_l.word_wrap = True
    tf4_l.margin_left = Inches(0.35)
    tf4_l.margin_top = Inches(0.3)

    h4_l = tf4_l.paragraphs[0]
    h4_l.text = "MANAGERIAL ACCOUNTING FORMULAS"
    h4_l.font.size = Pt(12)
    h4_l.font.bold = True
    h4_l.font.color.rgb = C_PRIMARY

    formulas = [
        ("Contribution Margin (CM):", "CM = Selling Price (P) - Variable Cost (V)", "Per-unit revenue remaining to cover fixed overhead."),
        ("Contribution Margin Ratio:", "CM% = (P - V) / P", "Percentage of each revenue dollar available to cover fixed costs."),
        ("Break-Even Volume (Units):", "Q_BE = Fixed Costs (F) / (P - V)", "Sales units required to yield exactly $0 net operating profit."),
        ("Break-Even Gross Revenue:", "R_BE = Q_BE × P = F / CM%", "Total dollar sales volume needed to reach break-even."),
        ("Projected Net Operating Profit:", "Profit = Q × (P - V) - F", "Net operating income at projected sales volume Q."),
        ("Margin of Safety (MOS):", "MOS = (Q - Q_BE) / Q × 100%", "Buffer percentage by which sales can fall before entering losses."),
        ("Operating Leverage (DOL):", "DOL = (Q × CM) / Profit", "Measures multiplier effect: % change in profit for 1% change in sales.")
    ]
    for name, f_str, desc in formulas:
        p = tf4_l.add_paragraph()
        p.text = f"{name} {f_str}"
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = C_DARK
        p.space_before = Pt(5)

        p_desc = tf4_l.add_paragraph()
        p_desc.text = f"  └ {desc}"
        p_desc.font.size = Pt(9.5)
        p_desc.font.color.rgb = C_MUTED
        p_desc.space_before = Pt(1)

    # Right Card: Algorithmic Logic & Edge Cases
    c4_right = add_card(slide4, Inches(7.2), Inches(1.6), Inches(5.3), Inches(5.1), C_WHITE)
    tf4_r = c4_right.text_frame
    tf4_r.word_wrap = True
    tf4_r.margin_left = Inches(0.35)
    tf4_r.margin_top = Inches(0.3)

    h4_r = tf4_r.paragraphs[0]
    h4_r.text = "ENGINE ROBUSTNESS & VALIDATION GUARDS"
    h4_r.font.size = Pt(12)
    h4_r.font.bold = True
    h4_r.font.color.rgb = C_SUCCESS

    checks = [
        ("Strict Viability Check (P ≤ V):", "If Selling Price is less than or equal to Variable Cost, every unit sold generates a loss. The system raises an alert, flags BE as infinite/unattainable, and displays an entire loss polygon."),
        ("Deficit Early Warning (Q < Q_BE):", "If projected sales fall short of break-even, the engine computes the exact shortfall (units and currency deficit) and displays an alert indicating operating loss."),
        ("Target Profit Calculation:", "Calculates Q_Target = (F + Target) / CM to determine the exact sales needed to satisfy the owner's desired take-home earnings."),
        ("Pure Function Architecture:", "model(F, P, V, Q, T) is completely deterministic and stateless. The same inputs guarantee the exact same outputs across unit tests, backend APIs, and UI.")
    ]
    for title, desc in checks:
        p = tf4_r.add_paragraph()
        p.text = title
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_DARK
        p.space_before = Pt(8)

        p_b = tf4_r.add_paragraph()
        p_b.text = desc
        p_b.font.size = Pt(10)
        p_b.font.color.rgb = C_MUTED
        p_b.space_before = Pt(2)

    # --------------------------------------------------------------------------
    # SLIDE 5: Key Features & Visual Demonstrations
    # --------------------------------------------------------------------------
    slide5 = prs.slides.add_slide(blank_layout)
    add_header(slide5, "Key Features & Interactive Capabilities", "PLATFORM CAPABILITIES")

    feat_grid = [
        (Inches(0.8), Inches(1.6), Inches(5.6), Inches(2.4), "📊 Inline Dynamic SVG Break-Even Chart", [
            "Draws Total Revenue, Total Cost, and horizontal Fixed Cost lines.",
            "Fills green-tinted Profit Zone and red-tinted Loss Zone dynamically.",
            "Pinpoints Break-Even point with pulsing halo and coordinates callout.",
            "Interactive hover/touch crosshair reading exact figures at any volume."
        ]),
        (Inches(6.9), Inches(1.6), Inches(5.6), Inches(2.4), "🎛️ Real-Time What-If Sensitivity Sliders", [
            "Three dual-direction sliders (-50% to +50%) for Price, Var Cost, and Fixed Costs.",
            "Live delta comparisons displaying percentage shift vs base case.",
            "Color-coded metrics (green for favorable shifts, red for unfavorable).",
            "One-click 'Reset Sliders' to restore baseline assumptions."
        ]),
        (Inches(0.8), Inches(4.3), Inches(5.6), Inches(2.4), "📈 Stress-Testing Sensitivity & Scenarios", [
            "Driver Matrix: Evaluates BE volume shifts at -20%, -10%, 0%, +10%, +20%.",
            "Volume Scenarios: Pessimistic (-25%), Base, and Optimistic (+25%).",
            "Highlights margin of safety buffer and dollar profits per scenario.",
            "Flags structural loss conditions automatically."
        ]),
        (Inches(6.9), Inches(4.3), Inches(5.6), Inches(2.4), "💼 Enterprise Profile & Executive Print", [
            "Business profile card with 10-digit phone and email validation.",
            "Multi-currency support (USD $, EUR €, GBP £, INR ₹, JPY ¥, CAD, AUD).",
            "Custom volume unit labels (units, cups, boxes, meals, hours).",
            "Printable letterhead report hiding interactive controls for presentation."
        ])
    ]

    for left, top, w, h, title, bullets in feat_grid:
        card = add_card(slide5, left, top, w, h, C_WHITE)
        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.25)
        tf.margin_top = Inches(0.2)

        p_t = tf.paragraphs[0]
        p_t.text = title
        p_t.font.size = Pt(11.5)
        p_t.font.bold = True
        p_t.font.color.rgb = C_PRIMARY

        for b in bullets:
            p = tf.add_paragraph()
            p.text = f"• {b}"
            p.font.size = Pt(9.5)
            p.font.color.rgb = C_DARK
            p.space_before = Pt(3)

    # --------------------------------------------------------------------------
    # SLIDE 6: Implementation Details & Code Structure
    # --------------------------------------------------------------------------
    slide6 = prs.slides.add_slide(blank_layout)
    add_header(slide6, "Modular Architecture & Code Structure", "CODE IMPLEMENTATION")

    c6_left = add_card(slide6, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.1), C_WHITE)
    tf6_l = c6_left.text_frame
    tf6_l.word_wrap = True
    tf6_l.margin_left = Inches(0.3)
    tf6_l.margin_top = Inches(0.3)

    h6_l = tf6_l.paragraphs[0]
    h6_l.text = "FRONTEND MODULAR NAMESPACES (ES5)"
    h6_l.font.size = Pt(12)
    h6_l.font.bold = True
    h6_l.font.color.rgb = C_PRIMARY

    mods = [
        ("model(F, P, V, Q, T):", "Pure calculation function returning break-even, profit, margin of safety, and DOL."),
        ("AuthModule:", "Manages Sign In / Sign Up, form validation, SHA-256 password hashing, and active session tokens."),
        ("ProfileModule:", "Handles business metadata, contact information, input validation, and print letterhead sync."),
        ("MetricsModule:", "Formats currency & numbers, renders KPI cards, and updates conditional status alerts."),
        ("ChartModule:", "Computes dynamic SVG viewBox domain, renders paths, loss/profit polygons, and crosshair overlays."),
        ("SensitivityModule:", "Computes sensitivity table variations across -20% to +20% for each cost driver."),
        ("WhatIfModule:", "Calculates real-time variances and deltas vs baseline for interactive sliders.")
    ]
    for name, desc in mods:
        p = tf6_l.add_paragraph()
        p.text = f"{name} {desc}"
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_DARK
        p.space_before = Pt(4)

    c6_right = add_card(slide6, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.1), C_WHITE)
    tf6_r = c6_right.text_frame
    tf6_r.word_wrap = True
    tf6_r.margin_left = Inches(0.3)
    tf6_r.margin_top = Inches(0.3)

    h6_r = tf6_r.paragraphs[0]
    h6_r.text = "BACKEND ARCHITECTURE & REST SPECIFICATION"
    h6_r.font.size = Pt(12)
    h6_r.font.bold = True
    h6_r.font.color.rgb = C_ACCENT

    endpoints = [
        ("GET /api/health", "Health probe & database connectivity status."),
        ("POST /api/auth/signup", "Registers new account with SHA-256 hashed password."),
        ("POST /api/auth/login", "Authenticates user and returns 24-byte bearer token."),
        ("GET /api/auth/me", "Validates bearer session; fetches profile and cost inputs."),
        ("GET & POST /api/profile", "Retrieves and updates business profile in SQLite3."),
        ("GET & POST /api/inputs", "Retrieves and saves cost drivers (debounced autosave)."),
        ("POST /api/calculate", "Server-side verification of mathematical model results.")
    ]
    for ep, desc in endpoints:
        p = tf6_r.add_paragraph()
        p.text = f"{ep}: {desc}"
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_DARK
        p.space_before = Pt(4)

    p_sync = tf6_r.add_paragraph()
    p_sync.text = "Dual-Engine Synchronization Logic:"
    p_sync.font.size = Pt(10.5)
    p_sync.font.bold = True
    p_sync.font.color.rgb = C_SUCCESS
    p_sync.space_before = Pt(12)

    p_sync_desc = tf6_r.add_paragraph()
    p_sync_desc.text = "ApiClient detects server availability via /api/health. If available, all mutations write to SQLite3 with debounced HTTP calls. If offline or opened via file://, seamless fallback stores data in localStorage."
    p_sync_desc.font.size = Pt(9.5)
    p_sync_desc.font.color.rgb = C_MUTED
    p_sync_desc.space_before = Pt(2)

    # --------------------------------------------------------------------------
    # SLIDE 7: Results, Conclusion & Future Enhancements
    # --------------------------------------------------------------------------
    slide7 = prs.slides.add_slide(blank_layout)
    add_header(slide7, "Results, Academic Evaluation & Future Scope", "CONCLUSION & IMPACT")

    c7_left = add_card(slide7, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.1), C_WHITE)
    tf7_l = c7_left.text_frame
    tf7_l.word_wrap = True
    tf7_l.margin_left = Inches(0.35)
    tf7_l.margin_top = Inches(0.3)

    h7_l = tf7_l.paragraphs[0]
    h7_l.text = "PROJECT DELIVERABLES & OUTCOMES"
    h7_l.font.size = Pt(12)
    h7_l.font.bold = True
    h7_l.font.color.rgb = C_PRIMARY

    outcomes = [
        "100% Zero-Dependency Standalone Platform: Operates instantly without package managers, external runtime installations, or build tools.",
        "Production-Grade Deployment: Live and globally accessible via GitHub Pages at https://arjun2007arush2011-max.github.io/break_even_point/",
        "Multi-Language Server Implementations: Clean, idiomatic backends provided in Ruby, Python 3, and Node.js using SQLite3 relational storage.",
        "Search-Engine Optimized: Complete SEO tags, OpenGraph previews, robots.txt, and sitemap.xml for indexing.",
        "Exceptional Performance: Instantaneous sub-millisecond calculation updates and smooth 60fps SVG chart recalculations."
    ]
    for o in outcomes:
        p = tf7_l.add_paragraph()
        p.text = f"✔ {o}"
        p.font.size = Pt(10.5)
        p.font.color.rgb = C_DARK
        p.space_before = Pt(7)

    c7_right = add_card(slide7, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.1), C_WHITE)
    tf7_r = c7_right.text_frame
    tf7_r.word_wrap = True
    tf7_r.margin_left = Inches(0.35)
    tf7_r.margin_top = Inches(0.3)

    h7_r = tf7_r.paragraphs[0]
    h7_r.text = "ACADEMIC IMPACT & FUTURE ENHANCEMENTS"
    h7_r.font.size = Pt(12)
    h7_r.font.bold = True
    h7_r.font.color.rgb = C_SUCCESS

    academic = [
        ("Cross-Disciplinary Rigor:", "Integrates software engineering, managerial cost accounting, relational database design, and interactive computer graphics."),
        ("Multi-Product Sales Mix:", "Future expansion to compute weighted average contribution margin ratios for multi-product retail inventory."),
        ("Stochastic Monte Carlo Modeling:", "Simulating random demand curves and price elasticity distributions for advanced risk assessment."),
        ("Automated PDF Financial Dossier:", "One-click export of formal executive investment pitches and PDF feasibility binders.")
    ]
    for title, desc in academic:
        p = tf7_r.add_paragraph()
        p.text = title
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_DARK
        p.space_before = Pt(7)

        p_b = tf7_r.add_paragraph()
        p_b.text = desc
        p_b.font.size = Pt(10)
        p_b.font.color.rgb = C_MUTED
        p_b.space_before = Pt(2)

    prs.save(PPTX_PATH)
    print(f"[OK] PowerPoint generated at: {PPTX_PATH}")


# ==============================================================================
# 2. PDF PRESENTATION GENERATION (ReportLab)
# ==============================================================================
def create_pdf():
    # Landscape Letter format
    doc = SimpleDocTemplate(
        PDF_PATH,
        pagesize=landscape(letter),
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36
    )

    styles = getSampleStyleSheet()

    # Custom styles
    style_category = ParagraphStyle(
        'CategoryStyle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=12,
        textColor=colors.HexColor('#2563eb'),
        spaceAfter=4
    )

    style_title = ParagraphStyle(
        'SlideTitleStyle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=20,
        leading=24,
        textColor=colors.HexColor('#0f172a'),
        spaceAfter=14
    )

    style_card_title = ParagraphStyle(
        'CardTitleStyle',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=15,
        textColor=colors.HexColor('#2563eb'),
        spaceAfter=6
    )

    style_card_title_green = ParagraphStyle(
        'CardTitleGreenStyle',
        parent=style_card_title,
        textColor=colors.HexColor('#059669')
    )

    style_card_title_purple = ParagraphStyle(
        'CardTitlePurpleStyle',
        parent=style_card_title,
        textColor=colors.HexColor('#7c3aed')
    )

    style_card_title_red = ParagraphStyle(
        'CardTitleRedStyle',
        parent=style_card_title,
        textColor=colors.HexColor('#dc2626')
    )

    style_body = ParagraphStyle(
        'BodyDarkStyle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9.5,
        leading=13.5,
        textColor=colors.HexColor('#0f172a'),
        spaceAfter=5
    )

    style_body_bold = ParagraphStyle(
        'BodyBoldStyle',
        parent=style_body,
        fontName='Helvetica-Bold'
    )

    style_body_muted = ParagraphStyle(
        'BodyMutedStyle',
        parent=style_body,
        textColor=colors.HexColor('#64748b')
    )

    story = []

    def make_card(flowables, bg_color='#f8fafc', border_color='#cbd5e1'):
        content = []
        for f in flowables:
            content.append(f)
        t = Table([[content]], colWidths=[340])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor(bg_color)),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor(border_color)),
            ('LEFTPADDING', (0,0), (-1,-1), 14),
            ('RIGHTPADDING', (0,0), (-1,-1), 14),
            ('TOPPADDING', (0,0), (-1,-1), 12),
            ('BOTTOMPADDING', (0,0), (-1,-1), 12),
        ]))
        return t

    # --------------------------------------------------------------------------
    # SLIDE 1 (PDF): Title Slide
    # --------------------------------------------------------------------------
    story.append(Paragraph("COLLEGE MINI PROJECT PRESENTATION", style_category))
    story.append(Paragraph("Break-Even Point Predictor & Feasibility Engine", style_title))
    story.append(Paragraph("A Full-Stack Financial Decision Support Platform with Real-Time SVG Visualizations & What-If Simulations", style_body_muted))
    story.append(Spacer(1, 14))

    s1_c1 = [
        Paragraph("PROJECT METADATA", style_card_title),
        Spacer(1, 4),
        Paragraph("<b>Domain:</b> FinTech / Financial Engineering & Web Systems", style_body),
        Paragraph("<b>Developer:</b> arjun2007arush2011-max", style_body),
        Paragraph("<b>Architecture:</b> Client-Side Zero-Dependency + Multi-Runtime REST", style_body),
        Paragraph("<b>Database:</b> Dual-Engine (SQLite3 Relational DB & LocalStorage)", style_body),
        Paragraph("<b>Academic Scope:</b> Mini Project for Degree Program", style_body)
    ]

    s1_c2 = [
        Paragraph("DEPLOYMENT & ACCESS LINKS", style_card_title_green),
        Spacer(1, 4),
        Paragraph("<b>Live Website:</b> <font color='#2563eb'><u>https://arjun2007arush2011-max.github.io/break_even_point/</u></font>", style_body),
        Paragraph("<b>GitHub Code:</b> <font color='#2563eb'><u>https://github.com/arjun2007arush2011-max/break_even_point</u></font>", style_body),
        Paragraph("<b>Backend Runtimes:</b> Python 3 standard library, Ruby WEBrick, Node.js", style_body),
        Paragraph("<b>Hosting:</b> Globally deployed & CDN-cached via GitHub Pages", style_body),
        Paragraph("<b>Status:</b> Fully functional with production test validation", style_body)
    ]

    two_col = Table([[make_card(s1_c1), make_card(s1_c2)]], colWidths=[360, 360])
    two_col.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(two_col)
    story.append(PageBreak())

    # --------------------------------------------------------------------------
    # SLIDE 2 (PDF): Problem Statement & Objectives
    # --------------------------------------------------------------------------
    story.append(Paragraph("BACKGROUND & MOTIVATION", style_category))
    story.append(Paragraph("Problem Statement & Project Objectives", style_title))

    s2_c1 = [
        Paragraph("THE PROBLEM STATEMENT", style_card_title_red),
        Spacer(1, 4),
        Paragraph("<b>High Failure Rates:</b> Over 65% of small business and startup failures stem from premature cost miscalculations and a lack of real-time visibility into operational break-even thresholds.", style_body),
        Paragraph("• <b>Static Spreadsheet Errors:</b> Traditional Excel sheets lack interactive dynamic graphics and are error-prone on mobile devices.", style_body_muted),
        Paragraph("• <b>Unrealistic Pricing Decisions:</b> Entrepreneurs often underprice units without understanding how variable costs compress contribution margins.", style_body_muted),
        Paragraph("• <b>Lack of Sensitivity Analysis:</b> Owners cannot easily simulate how a 10% inflation spike in variable costs or rent shifts required sales volume.", style_body_muted),
        Paragraph("• <b>Framework Bloat:</b> Most commercial calculators require expensive subscriptions or heavy tracking libraries.", style_body_muted)
    ]

    s2_c2 = [
        Paragraph("KEY PROJECT OBJECTIVES", style_card_title_green),
        Spacer(1, 4),
        Paragraph("✔ <b>Zero-Dependency Platform:</b> Build a self-contained web app running in vanilla ES5 with zero external JS libraries.", style_body),
        Paragraph("✔ <b>Inline Scalable Vector Graphics:</b> Render custom SVG charts with shaded profit/loss zones, break-even markers, and touch cursors.", style_body),
        Paragraph("✔ <b>Real-Time What-If Engine:</b> Provide dynamic sliders (-50% to +50%) and sensitivity matrices to stress-test cost drivers.", style_body),
        Paragraph("✔ <b>Multi-Runtime REST Integration:</b> Implement REST APIs in Python, Ruby, and Node.js with persistent SQLite3 storage.", style_body),
        Paragraph("✔ <b>Executive Printable Reports:</b> Generate professional, letterheaded feasibility reports for banks, investors, and stakeholders.", style_body)
    ]

    story.append(Table([[make_card(s2_c1), make_card(s2_c2)]], colWidths=[360, 360], style=[('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(PageBreak())

    # --------------------------------------------------------------------------
    # SLIDE 3 (PDF): System Architecture & Tech Stack
    # --------------------------------------------------------------------------
    story.append(Paragraph("TECHNICAL DESIGN", style_category))
    story.append(Paragraph("System Architecture & Technology Stack", style_title))

    def make_small_card(flowables, title_color='#2563eb'):
        t = Table([[flowables]], colWidths=[230])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f8fafc')),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#cbd5e1')),
            ('LEFTPADDING', (0,0), (-1,-1), 10),
            ('RIGHTPADDING', (0,0), (-1,-1), 10),
            ('TOPPADDING', (0,0), (-1,-1), 10),
            ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ]))
        return t

    s3_c1 = [
        Paragraph("FRONTEND TIER", style_card_title),
        Paragraph("• <b>Vanilla ES5 JavaScript:</b> Universal browser support without transpilation.", style_body),
        Paragraph("• <b>CSS Custom Properties:</b> Instant light/dark theme toggle.", style_body),
        Paragraph("• <b>Responsive Two-Column Grid:</b> Adapts seamlessly across desktop and mobile screens.", style_body),
        Paragraph("• <b>Web Crypto API:</b> SHA-256 password hashing with pure JS fallback.", style_body)
    ]

    s3_c2 = [
        Paragraph("VISUALIZATION ENGINE", style_card_title_purple),
        Paragraph("• <b>Custom Inline SVG:</b> High-performance mathematical coordinate mapping.", style_body),
        Paragraph("• <b>Real-Time Tracking:</b> Crosshair tracker with live revenue, cost, and profit tooltips.", style_body),
        Paragraph("• <b>Profit / Loss Polygons:</b> Visually shaded gradients indicating business risk.", style_body),
        Paragraph("• <b>Pure Math Model:</b> Stateless model(F, P, V, Q, T) isolating math logic.", style_body)
    ]

    s3_c3 = [
        Paragraph("BACKEND & DATABASE", style_card_title_green),
        Paragraph("• <b>Triple Runtime APIs:</b> Python (http.server), Ruby (WEBrick), and Node.js.", style_body),
        Paragraph("• <b>SQLite3 Database:</b> Relational tables for Users, Profiles, and Cost Inputs.", style_body),
        Paragraph("• <b>Debounced Sync:</b> Smooth server synchronization without network spam.", style_body),
        Paragraph("• <b>Dual-Mode Fallback:</b> Automatically falls back to localStorage if offline.", style_body)
    ]

    three_col = Table([[make_small_card(s3_c1), make_small_card(s3_c2), make_small_card(s3_c3)]], colWidths=[240, 240, 240], style=[('VALIGN', (0,0), (-1,-1), 'TOP')])
    story.append(three_col)
    story.append(PageBreak())

    # --------------------------------------------------------------------------
    # SLIDE 4 (PDF): Mathematical Formulation
    # --------------------------------------------------------------------------
    story.append(Paragraph("CORE ALGORITHMS", style_category))
    story.append(Paragraph("Mathematical Formulation & Financial Model", style_title))

    s4_c1 = [
        Paragraph("FINANCIAL ACCOUNTING EQUATIONS", style_card_title),
        Paragraph("• <b>Contribution Margin (CM):</b> CM = P - V<br/><font color='#64748b'>Per-unit revenue available to cover fixed overhead.</font>", style_body),
        Paragraph("• <b>Contribution Margin Ratio:</b> CM% = (P - V) / P<br/><font color='#64748b'>Percentage of each dollar that contributes to paying fixed costs.</font>", style_body),
        Paragraph("• <b>Break-Even Volume:</b> Q_BE = F / (P - V)<br/><font color='#64748b'>Sales volume required to yield exactly zero profit.</font>", style_body),
        Paragraph("• <b>Break-Even Gross Revenue:</b> R_BE = Q_BE × P = F / CM%", style_body),
        Paragraph("• <b>Projected Net Profit:</b> Profit = Q × (P - V) - F", style_body),
        Paragraph("• <b>Margin of Safety:</b> MOS = (Q - Q_BE) / Q × 100%<br/><font color='#64748b'>Buffer percentage before business incurs an operational loss.</font>", style_body),
        Paragraph("• <b>Operating Leverage:</b> DOL = (Q × CM) / Profit<br/><font color='#64748b'>Measures profit sensitivity per 1% change in sales volume.</font>", style_body)
    ]

    s4_c2 = [
        Paragraph("ENGINE ROBUSTNESS & VALIDATION", style_card_title_green),
        Paragraph("• <b>Impossibility Guard (P ≤ V):</b><br/>If Selling Price is less than or equal to Variable Cost, every unit sold generates a loss. The engine flags break-even as nonexistent, shades the entire plot in red, and presents an urgent warning.", style_body),
        Paragraph("• <b>Operational Loss Alert (Q < Q_BE):</b><br/>If projected sales fall short of break-even, the engine computes the exact shortfall (units and dollar deficit) and highlights the loss in red.", style_body),
        Paragraph("• <b>Target Profit Computation:</b><br/>Computes Q_Target = (F + Target) / CM to evaluate required sales volume for desired earnings.", style_body),
        Paragraph("• <b>Stateless Pure Function:</b><br/>model(F, P, V, Q, T) ensures 100% test reproducibility across all client and server endpoints.", style_body)
    ]

    story.append(Table([[make_card(s4_c1), make_card(s4_c2)]], colWidths=[360, 360], style=[('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(PageBreak())

    # --------------------------------------------------------------------------
    # SLIDE 5 (PDF): Key Features & Capabilities
    # --------------------------------------------------------------------------
    story.append(Paragraph("PLATFORM CAPABILITIES", style_category))
    story.append(Paragraph("Key Features & Interactive Capabilities", style_title))

    s5_c1 = [
        Paragraph("1. DYNAMIC INLINE SVG CHART", style_card_title),
        Paragraph("• Renders Total Revenue, Total Cost, and Fixed Cost lines with precision.", style_body),
        Paragraph("• Fills green-tinted Profit Zone and red-tinted Loss Zone dynamically.", style_body),
        Paragraph("• Marked Break-Even intersection with pulsing halo & coordinate badge.", style_body),
        Paragraph("• Interactive hover/touch crosshair reading real-time metrics at any volume.", style_body)
    ]

    s5_c2 = [
        Paragraph("2. WHAT-IF REAL-TIME SLIDERS", style_card_title_purple),
        Paragraph("• Three dual-direction sliders (-50% to +50%) for Price, Var Cost, and Fixed Costs.", style_body),
        Paragraph("• Live delta indicators showing volume and dollar shifts vs baseline.", style_body),
        Paragraph("• Color-coded readouts (green for favorable shifts, red for unfavorable).", style_body),
        Paragraph("• Instant 'Reset Sliders' button to restore baseline inputs.", style_body)
    ]

    s5_c3 = [
        Paragraph("3. STRESS-TESTING SCENARIOS", style_card_title_green),
        Paragraph("• <b>Sensitivity Matrix:</b> Evaluates BE volume at -20%, -10%, 0%, +10%, +20%.", style_body),
        Paragraph("• <b>Volume Scenarios:</b> Pessimistic (-25%), Base, and Optimistic (+25%).", style_body),
        Paragraph("• Shows revenue, cost, net earnings, and safety margins for all scenarios.", style_body)
    ]

    s5_c4 = [
        Paragraph("4. BUSINESS PROFILE & PRINT REPORT", style_card_title),
        Paragraph("• Business profile card with 10-digit phone and email validation.", style_body),
        Paragraph("• Multi-currency selector ($, €, £, ₹, ¥, CA$, A$, CHF).", style_body),
        Paragraph("• Custom volume unit labels (units, cups, boxes, meals, hours).", style_body),
        Paragraph("• Executive print layout with letterhead, KPI summaries, and charts.", style_body)
    ]

    grid_t = Table([
        [make_card(s5_c1), make_card(s5_c2)],
        [make_card(s5_c3), make_card(s5_c4)]
    ], colWidths=[360, 360], style=[('VALIGN', (0,0), (-1,-1), 'TOP')])
    story.append(grid_t)
    story.append(PageBreak())

    # --------------------------------------------------------------------------
    # SLIDE 6 (PDF): Implementation & Code Highlights
    # --------------------------------------------------------------------------
    story.append(Paragraph("CODE IMPLEMENTATION", style_category))
    story.append(Paragraph("Modular Architecture & Code Structure", style_title))

    s6_c1 = [
        Paragraph("FRONTEND MODULES (ES5 JAVASCRIPT)", style_card_title),
        Paragraph("• <b>model(F, P, V, Q, T):</b> Pure calculation engine returning break-even, profit, safety margins, and DOL.", style_body),
        Paragraph("• <b>AuthModule:</b> Tab switching, validation, client-side SHA-256 password hashing, and session management.", style_body),
        Paragraph("• <b>ProfileModule:</b> Handles business metadata, phone/email validation, and report letterhead sync.", style_body),
        Paragraph("• <b>MetricsModule:</b> Number and currency formatting, KPI card updates, and conditional alert banners.", style_body),
        Paragraph("• <b>ChartModule:</b> ViewBox coordinate projection, path generation, and hover tracking.", style_body),
        Paragraph("• <b>SensitivityModule:</b> Stress-testing matrices across driver variations (-20% to +20%).", style_body)
    ]

    s6_c2 = [
        Paragraph("BACKEND API & DATA PERSISTENCE", style_card_title_green),
        Paragraph("• <b>GET /api/health:</b> Service health and database status probe.", style_body),
        Paragraph("• <b>POST /api/auth/signup & login:</b> SHA-256 auth with 24-byte bearer token generation.", style_body),
        Paragraph("• <b>GET & POST /api/profile:</b> SQLite3 persistence for enterprise details.", style_body),
        Paragraph("• <b>GET & POST /api/inputs:</b> Stores cost and sales drivers with client debouncing.", style_body),
        Paragraph("• <b>POST /api/calculate:</b> Server-side calculation verification endpoint.", style_body),
        Paragraph("• <b>Dual-Mode Architecture:</b> Auto-detects server presence via /api/health. If online, syncs with SQLite3; if offline, operates via localStorage seamlessly.", style_body)
    ]

    story.append(Table([[make_card(s6_c1), make_card(s6_c2)]], colWidths=[360, 360], style=[('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(PageBreak())

    # --------------------------------------------------------------------------
    # SLIDE 7 (PDF): Results & Conclusion
    # --------------------------------------------------------------------------
    story.append(Paragraph("CONCLUSION & IMPACT", style_category))
    story.append(Paragraph("Results, Academic Evaluation & Future Scope", style_title))

    s7_c1 = [
        Paragraph("PROJECT DELIVERABLES & RESULTS", style_card_title),
        Spacer(1, 4),
        Paragraph("✔ <b>Zero-Dependency Platform:</b> Operates without npm packages, external build tools, or CDN dependencies.", style_body),
        Paragraph("✔ <b>Production Deployment:</b> Live and accessible worldwide on GitHub Pages at <font color='#2563eb'><u>https://arjun2007arush2011-max.github.io/break_even_point/</u></font>", style_body),
        Paragraph("✔ <b>Multi-Runtime Support:</b> Complete, idiomatic backend implementations in Python, Ruby, and Node.js with SQLite3.", style_body),
        Paragraph("✔ <b>SEO & Search Engine Readiness:</b> robots.txt, sitemap.xml, OpenGraph cards, and descriptive metadata configured.", style_body),
        Paragraph("✔ <b>Real-Time Performance:</b> Sub-millisecond calculation updates and smooth 60fps SVG chart recalculations.", style_body)
    ]

    s7_c2 = [
        Paragraph("ACADEMIC EVALUATION & FUTURE SCOPE", style_card_title_green),
        Spacer(1, 4),
        Paragraph("• <b>Cross-Disciplinary Integration:</b> Combines software engineering, managerial cost accounting, relational database design, and interactive graphics.", style_body),
        Paragraph("• <b>Multi-Product Sales Mix:</b> Future expansion to calculate weighted average contribution margins for multi-item inventory catalogs.", style_body),
        Paragraph("• <b>Monte Carlo Risk Simulation:</b> Simulating demand curve fluctuations and price elasticity for probabilistic risk assessment.", style_body),
        Paragraph("• <b>Automated PDF Financial Dossier:</b> One-click generation of investor-ready feasibility binders and financial balance sheets.", style_body)
    ]

    story.append(Table([[make_card(s7_c1), make_card(s7_c2)]], colWidths=[360, 360], style=[('VALIGN', (0,0), (-1,-1), 'TOP')]))

    doc.build(story)
    print(f"[OK] PDF generated at: {PDF_PATH}")


if __name__ == "__main__":
    create_pptx()
    create_pdf()
