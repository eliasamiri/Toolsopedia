# Toolsopedia

Toolsopedia is a collection of lightweight web tools. The root landing page lists available tools, and each tool lives in its own folder under `tools/`.

## Current Tools

- **PDF to Excel** — `tools/pdf-to-excel-app/`

## How It Works

The landing page lives at the project root (`index.html`) and is served by the Flask app. Each tool is mounted under a URL that matches its folder name, for example:

- Landing page: `http://127.0.0.1:5000/`
- PDF to Excel: `http://127.0.0.1:5000/tools/pdf-to-excel-app/`

## Project Structure

```
Toolsopedia/
├── index.html
├── styles.css
├── tools/
│   └── pdf-to-excel-app/
│       ├── src/
│       ├── templates/
│       ├── static/
│       ├── uploads/
│       ├── converted/
│       └── requirements.txt
└── .gitignore
```

## Run Locally

1. Create and activate a virtual environment (optional but recommended).
2. Install the PDF to Excel app dependencies.
3. Start the Flask server.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r tools/pdf-to-excel-app/requirements.txt
python tools/pdf-to-excel-app/src/app.py
```

Then open `http://127.0.0.1:5000/`.

## Add New Tools

Create a new folder under `tools/` and add a card to the root `index.html` linking to `tools/<folder-name>/`.
