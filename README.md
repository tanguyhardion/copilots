# Office & Document Copilots — PowerPoint & CV Suite

A clean, maintainable, and modern desktop application built with **pywebview** (HTML5 / CSS3 / JavaScript frontend backed by native WebView2 and Python automation engines) featuring two essential Office & Document AI Copilots:

1. **PowerPoint Copilot** — Design Consultant & Shape Architect (DSL to PowerPoint shapes, slides, tables, gradients, and icons via COM).
2. **CV Copilot** — Deterministic Data Quality (DQ) validation engine, 1-Slide Executive PowerPoint (.pptx) proposal generator, and standard Europass Word (.docx) generator.

---

## Architecture & Project Structure

```
copilots/
├── app.py                      # Main pywebview application launcher
├── requirements.txt            # Dependencies
├── assets/
│   ├── icons/                  # App branding icons (copilots.png, powerpoint.png, cv.png)
│   └── templates/              # Document templates (cv_template.pptx, photo_placeholder.jpg)
└── copilots_app/
    ├── core/                   # Design system tokens, config, events, prompt manager
    │   ├── theme.py
    │   ├── config.py
    │   ├── events.py
    │   └── prompt_manager.py
    ├── prompts/                # Bundled system prompts
    │   ├── powerpoint_prompt.md
    │   └── cv_prompt.md
    ├── web/                    # pywebview HTML5/CSS/JS frontend
    │   ├── index.html
    │   ├── app.js
    │   ├── css/
    │   └── js/
    └── services/               # Automation engines & business logic
        ├── powerpoint/         # PowerPoint DSL parser, shape engine, COM connector
        └── cv/                 # CV JSON parser, DQ validation engine, docx & pptx generators
```

---

## Getting Started

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Launch the Application
```bash
python app.py
```
