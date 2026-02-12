# Medical Invoice System

A professional web application for generating and managing medical invoices with Arabic/English support.

## Features

- 📄 **Excel Upload**: Process bulk invoice data from Excel files
- 🎨 **Professional PDF Generation**: Creates styled invoices matching Andalusia Hospital design
- 🌐 **Bilingual**: Full Arabic & English support with proper RTL text rendering
- 💾 **Local Storage**: Persistent invoice database
- 📥 **Batch Download**: Download all invoices as a single ZIP file
- 🎯 **Accurate Data Extraction**: Automatically extracts metadata and line items

## Quick Start

### Installation

```bash
pip install -r requirements.txt
```

### Running Locally

```bash
streamlit run app.py
```

Visit `http://localhost:8501` in your browser.

## Deployment

### Streamlit Community Cloud (Recommended)

1. Push this repository to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub account and select this repository
4. Deploy!

The app will automatically detect `app.py` and `requirements.txt`.

### Other Platforms

The included `Procfile` supports deployment to:
- Railway
- Heroku
- Any platform supporting Python web apps

## File Structure

```
.
├── app.py                      # Main Streamlit application
├── generate_invoices.py        # PDF generation engine
├── requirements.txt            # Python dependencies
├── Procfile                    # Deployment configuration
├── fonts/                      # Arabic font files
│   └── Arial Unicode.ttf
├── Picture1.png                # Hospital logo
└── invoices/                   # Generated PDFs (auto-created)
```

## Usage

1. **Upload Data**: Go to "Upload New Data" and upload your Excel file
2. **View Dashboard**: See all processed invoices in the Dashboard
3. **Download**: Download individual invoices or all at once as ZIP

## Technologies

- **Streamlit**: Web framework
- **ReportLab**: PDF generation
- **Pandas**: Data processing
- **Arabic-Reshaper & Python-BIDI**: Arabic text support

## License

© 2026 Rockai Dev

---

Powered by Advanced Agentic AI
