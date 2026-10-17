# Break-Even Point Predictor & Feasibility Engine

A full-stack, standalone financial forecasting platform designed to calculate break-even volume, profitability thresholds, margin of safety, and operational sensitivity with real-time what-if simulations and inline SVG charts.


---

## 📱 Live Web Application & Mobile QR Code

| **Live Website & Presentation** | **Scan with Mobile Camera** |
| :--- | :---: |
| 🌐 **Live Web Application:** [https://arjun2007arush2011-max.github.io/break_even_point/](https://arjun2007arush2011-max.github.io/break_even_point/)<br><br>📊 **Presentation Slides (PPTX):** [Break-Even-Project-Presentation.pptx](Break-Even-Project-Presentation.pptx)<br>📄 **Presentation Document (PDF):** [Break-Even-Project-Presentation.pdf](Break-Even-Project-Presentation.pdf)<br><br>• Hosted live on GitHub Pages with global SSL/HTTPS.<br>• Full financial simulator, interactive SVG charts, and sensitivity sliders.<br>• Unified Full-Stack Architecture & zero external dependencies. | <a href="https://arjun2007arush2011-max.github.io/break_even_point/"><img src="qrcode_live.png" alt="Scan to Launch Live App" width="160" height="160"></a><br><sub><i>Point camera to launch on phone</i></sub> |


## Architecture Overview

```
breakeven-predictor/
├── index.html       # Self-contained frontend (HTML5, responsive CSS variables, ES5 Vanilla JS)
├── start.sh         # Universal launcher script (auto-detects Ruby, Python 3, or Node.js)
├── server.rb        # Ruby backend (WEBrick + SQLite3, zero external gem dependencies)
├── server.py        # Python 3 backend (http.server + SQLite3, standard library only)
├── server.js        # Node.js backend (native http + JSON/SQLite storage, standard modules only)
├── breakeven.db     # SQLite3 database (users, profiles, cost & sales inputs)
└── README.md        # Technical reference & API documentation
```

---

## Quick Start: Running the Application

### Option 1: Universal Launcher (Recommended)
Open a terminal in the project directory and run:
```bash
./start.sh
```
The script detects which runtime is available on your machine (Ruby, Python 3, or Node.js) and starts the server on port 8000. Then open:
👉 **[http://localhost:8000/](http://localhost:8000/)**

### Option 2: Run via Ruby (Native on macOS)
```bash
ruby server.rb
```
- Built with WEBrick and native `sqlite3`.
- Automatically initializes `breakeven.db` and seeds the demo account.

### Option 3: Run via Python 3
```bash
python3 server.py
```
- Uses Python's standard `http.server` and `sqlite3`.

### Option 4: Run via Node.js
```bash
node server.js
```
- Uses Node.js standard `http` and `crypto` modules.

### Option 5: Standalone Client Mode (Zero Backend)
You can also open [index.html](file:///Users/arjunshukla/Downloads/breakevven%20point/index.html) directly in any web browser without running a server. The application will detect offline mode and store all user records, profiles, and inputs safely in `localStorage` with SHA-256 password hashing.

---

## Default Demo Credentials

Click the **"⚡ Quick Fill Demo Account"** button on the login screen, or enter:
- **Email:** `demo@business.com`
- **Password:** `Demo1234`
- **Profile:** *Velvet Bean Roastery* (Elena Rostova, Restaurant / Food, Seattle, WA)

---

## ⚡ Pre-Loaded Industry Templates (1-Click Demo)

The platform includes 6 real-world industry templates accessible via 1-click preset buttons or dropdown:

| Icon | Industry Template | Sample Firm | Unit Price | Variable Cost | Fixed Costs / Mo | Expected Volume |
| :---: | :--- | :--- | :---: | :---: | :---: | :---: |
| ☕ | **Artisan Coffee Shop & Bakery** | Velvet Bean Roastery | $5.50 / cup | $1.40 | $12,500 | 3,600 cups |
| 💻 | **B2B Cloud SaaS Platform** | CloudMetrics Analytics | $149.00 / sub | $18.00 | $28,000 | 320 subs |
| 👕 | **D2C Apparel E-Commerce** | Aura Urban Streetwear | $55.00 / order | $21.00 | $9,500 | 500 orders |
| 🏭 | **Precision CNC Manufacturing** | TitanPrecision Forge | $340.00 / part | $155.00 | $42,000 | 380 parts |
| 🩺 | **Dental & Health Clinic** | Horizon Dental Care | $175.00 / visit | $32.00 | $21,500 | 240 visits |
| 🎨 | **Digital Creative Agency** | Apex Creative Labs | $2,800 / retainer | $750.00 | $16,000 | 12 retainers |

---

## Backend REST API Endpoints

All backend implementations provide identical RESTful API contracts:

| Method | Endpoint | Description | Auth Required |
|---|---|---|:---:|
| `GET` | `/api/health` | Service health status and database engine info | No |
| `POST` | `/api/auth/signup` | Register a new account with hashed password and initial profile | No |
| `POST` | `/api/auth/login` | Authenticate user, verify credentials, return session token | No |
| `GET` | `/api/auth/me` | Fetch currently authenticated user, profile, and cost inputs | Bearer Token |
| `POST` | `/api/auth/logout` | Invalidate active session token | Bearer Token |
| `GET` | `/api/profile` | Retrieve user's business profile | Bearer Token |
| `POST` / `PUT` | `/api/profile` | Update firm name, owner, type, phone, email, and address | Bearer Token |
| `GET` | `/api/inputs` | Retrieve saved cost and sales drivers | Bearer Token |
| `POST` / `PUT` | `/api/inputs` | Save updated fixed costs, prices, variable costs, and volume | Bearer Token |
| `POST` | `/api/calculate` | Server-side calculation of full break-even model | Optional |

---

## Mathematical Model Reference

All financial calculations follow standardized managerial accounting formulas:

$$\text{Contribution Margin (CM)} = P - V$$
$$\text{Contribution Margin Ratio} = \frac{P - V}{P}$$
$$\text{Break-Even Units} = \frac{F}{P - V}$$
$$\text{Break-Even Revenue} = \text{Break-Even Units} \times P = \frac{F}{\text{CM Ratio}}$$
$$\text{Net Profit at Volume } Q = Q \times (P - V) - F$$
$$\text{Margin of Safety (Units)} = Q - \text{Break-Even Units}$$
$$\text{Margin of Safety (\%)} = \frac{Q - \text{Break-Even Units}}{Q} \times 100\%$$
$$\text{Units for Target Profit } T = \frac{F + T}{P - V}$$
$$\text{Degree of Operating Leverage (DOL)} = \frac{Q \times (P - V)}{\text{Net Profit}}$$

Where:
- $F$ = Total Fixed Costs
- $P$ = Selling Price per Unit
- $V$ = Variable Cost per Unit
- $Q$ = Expected Sales Volume
- $T$ = Target Profit

---

## Features

- **Inline SVG Charting**: Interactive crosshair tracking, dynamic scaling, shaded loss and profit zones, break-even point callout with halo, and vertical planned volume line.
- **What-If Adjustment Sliders**: Real-time slider adjustments (-50% to +50%) for Price, Variable Cost, and Fixed Costs with delta comparison versus base case.
- **Sensitivity Matrix**: Impact of $\pm 10\%$ and $\pm 20\%$ variations across each cost driver on break-even volume.
- **Volume Scenarios**: Pessimistic (-25%), Base (Expected), and Optimistic (+25%) sales volume projections.
- **Executive Printable Report**: Dedicated `@media print` layout formatting the firm's letterhead, KPIs, chart, and scenarios for presentation or PDF export.
- **Theme Customization**: Responsive dark and light mode toggle with color persistence.
- **Dual-Engine Persistence**: Automatically operates with backend SQLite server when connected, with instantaneous offline fallback to localStorage.
