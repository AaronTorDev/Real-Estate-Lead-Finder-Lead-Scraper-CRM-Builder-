# Real-Estate-Lead-Finder-Lead-Scraper-CRM-Builder-
Automated Python scraper &amp; lead scorer using Playwright. Generates a prioritized real estate CRM based on digital footprint analysis to pinpoint operational bottlenecks

An automated **Python** script designed to solve cold outreach friction and uncover operational inefficiencies in the real estate sector. The system extracts, filters, scores, and structures a database of active real estate brokers and agencies from the local Chamber of Associates, generating a high-value, action-ready CRM in Excel format.

## 📌 The Business Problem (The "Why")

In the real estate industry, manual prospecting is slow and inefficient. Developers often pitch automations to low-budget agencies or offer social media services that the client doesn't see as a critical, immediate need.

This project shifts the focus: **it targets only agencies with an active digital footprint** (those paying for their own domain, web hosting, and running active channels on Instagram, Facebook, and LinkedIn). These businesses have the budget to invest in tech, but they suffer from **backstage operational bottlenecks** (slow response times, manual lead copy-pasting, and lack of instant customer filtering).

This script acts as the **first step in the value chain**: it generates a structured CRM in seconds, ranking the highest-value prospects using a custom digital footprint *scoring* algorithm.

---

## 🛠️ How It Works

The script automates the extraction and scoring process through four main stages:

1. **Dynamic Web Scraping:** Uses **Playwright** to render the chamber's directory and handle infinite scroll dynamically, ensuring 100% data extraction with no manual limits.
2. **Digital Scoring Algorithm:** Evaluates each agency's digital assets and ranks them based on strategic weights:
   * **Custom Website:** +4 points (indicates ownership of digital assets and dedicated budget).
   * **Facebook / Instagram:** +2 points each.
   * **LinkedIn:** +1 point.
   * **Ecosystem Synergy:** Up to +2 bonus points for fully integrating multiple platforms.
3. **Data Normalization:** Sanitizes string names, parses raw emails from `mailto` links, and cleans up phone number formatting.
4. **Interactive CRM Export (Excel):** Generates a professional `.xlsx` file containing:
   * Automatic sorting by highest digital score.
   * Clickable hyperlinks pointing directly to their website and social media profiles.
   * In-cell data validation (interactive dropdown menus) to track review states and call status.
   * Dynamic conditional formatting for clean, visual pipeline tracking.

---

## 📋 System Requirements

The script is lightweight and runs locally on any environment supporting Python 3.

### Prerequisites
* **Python 3.8** or higher
* **pip** (Python package installer)

---

## 🚀 Installation & Setup

Follow these steps to clone the repository, set up your virtual environment, and run the lead scraper:

### 1. Clone the repository
```bash
git clone [https://github.com/your-username/your-repo-name.git](https://github.com/your-username/your-repo-name.git)
cd your-repo-name
```

### 2. Create and activate a Virtual Environment (Recommended)
*On Linux / WSL / macOS:*
```bash
python3 -m venv venv
source venv/bin/activate
```
*On Windows (PowerShell):*
```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```

### 3. Install dependencies
Install the required libraries using the package manager:
```bash
pip install playwright beautifulsoup4 openpyxl
```

### 4. Install Playwright browser binaries
Playwright needs its integrated headless browser binaries to run in the background:
```bash
playwright install chromium
```

### 5. Run the Script
Execute the main script to start the scraping process and build the CRM:
```bash
python scraper.py
```

Once completed, a success message will print in your terminal, and `crm_prospectos_completo.xlsx` will be generated in your project's root folder.

---

## 📁 CRM Output Structure

The generated Excel file contains the following columns optimized for business development:

| Column | Description |
|---|---|
| **ID** | Unique auto-incrementing identifier. |
| **SCORE** | Assigned digital footprint rating (Max: 11.0). |
| **FULL NAME** | Cleaned and normalized agency/broker name. |
| **WEB / IG / FB / LI** | Direct interactive links to easily audit their channels. |
| **EMAIL / PHONE** | Cleaned contact information. |
| **REVIEW STATUS** | Interactive dropdown: `[ ] Unreviewed`, `[X] Reviewed`. |
| **WORK STATUS** | Pipeline stage: `[ ] To Evaluate`, `REJECTED`, `PENDING (Free sample)`, `DONE (Ready to pitch)`. |

---

## 📈 Future Roadmap & Scalability
* **Automated Outbound Pipeline:** Direct integration with messaging and cold emailing APIs (e.g., SendGrid, Twilio) for instant touchpoints.
* **AI Auditor:** Integrating LLM agents to read the scraped websites and automatically pinpoint slow load speeds or outdated layout practices before making the sales pitch.
```
