# ⚡ LeadExtractor Pro v5.0

A professional AI-powered lead intelligence platform built with Python & Streamlit.

## 🚀 Features

| Feature | Description |
|---|---|
| ⚡ Single URL Extraction | Extract emails, phones, social profiles from any website |
| 🌍 Multi-page Crawling | Automatically crawls contact/about pages |
| 📱 Phone Extraction | Detects international phone numbers |
| 🔗 Social Profiles | LinkedIn, Facebook, Instagram detection |
| 🖼 Logo Detection | Finds company favicon/logo |
| 🤖 AI Summary | Claude-powered company summaries |
| 📥 Bulk CSV Upload | Upload hundreds of URLs at once |
| ⚡ Real-time Progress | Live extraction progress bar |
| 📅 Date Analytics | Daily/range-based lead analytics |
| 📄 Export Excel | Styled Excel export with summary sheet |
| 📑 Export PDF | Professional PDF report |
| ✏️ Edit Leads | Edit company, email, phone |
| 🗑 Delete Leads | Delete individual or all leads |
| 🔍 Search & Filter | Search + date range filtering |

## 📦 Installation

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 📁 Bulk CSV Format

Your CSV must have a column named `url`, `website`, `domain`, or `link`:

```csv
url
https://example.com
https://another.com
startup.io
```

## 🛠 Tech Stack

- Python 3.9+
- Streamlit (UI)
- BeautifulSoup4 (HTML parsing)
- SQLite (Database)
- OpenPyXL (Excel export)
- ReportLab (PDF export)
- Anthropic Claude API (AI summaries)
- Requests (HTTP)

---
Built by **Nakul Bidhuri** 🚀
