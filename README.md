# Office & Document Copilots — Unified Desktop Suite

A unified, maintainable, and modern desktop application built with **Flet** that brings together four essential Office & Document AI Copilots:

1. **PowerPoint Copilot** — Design Consultant & Shape Architect (DSL to PowerPoint shapes, slides, tables, gradients, and icons via COM).
2. **Word Copilot** — Document generator, live cursor injector, and active document editor with DSL extraction.
3. **Excel Copilot** — Air-gapped semantic analyzer, LLM context generator, and deterministic JSON Action Protocol executor.
4. **CV Copilot** — Deterministic Data Quality (DQ) validation engine and standard Europass Word (.docx) generator.

---

## Architecture & Project Structure

```
copilots/
├── app.py                      # Main Flet application launcher
├── requirements.txt            # Unified dependencies
├── assets/
│   └── icons/                  # App branding icons (word.png, excel.png, powerpoint.png, cv.png)
└── copilots_app/
    ├── core/                   # Design system tokens, config, events
    │   ├── theme.py
    │   ├── config.py
    │   └── events.py
    ├── ui/                     # Reusable UI components and unified Views
    │   ├── components.py       # AppHeader, StatusBar, CodeEditor, MetricCard, ActionButton
    │   ├── sidebar.py          # Navigation rail with active indicators & branding
    │   └── views/              # Dedicated views for each Copilot
    │       ├── powerpoint_view.py
    │       ├── word_view.py
    │       ├── excel_view.py
    │       └── cv_view.py
    └── services/               # Modularized business logic & engines
        ├── powerpoint/         # PowerPoint DSL parser, shape engine, COM connector
        ├── word/               # Word DSL parser, editor, extractor, COM connector
        ├── excel/              # Excel analyzer, protocol models, action executor, backups
        └── cv/                 # CV JSON parser, DQ validation engine, docx generator
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
