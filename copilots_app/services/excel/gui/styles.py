"""Modern Dark Theme QSS Stylesheet and Palette definitions."""

DARK_THEME = """
QMainWindow, QDialog {
    background-color: #121824;
    color: #E2E8F0;
    font-family: 'Segoe UI', 'Roboto', sans-serif;
    font-size: 13px;
}

QWidget {
    color: #E2E8F0;
    font-family: 'Segoe UI', 'Roboto', sans-serif;
}

/* ToolBar & MenuBar */
QToolBar {
    background-color: #1E293B;
    border-bottom: 1px solid #334155;
    spacing: 8px;
    padding: 6px;
}

QToolButton {
    background-color: #334155;
    color: #F8FAFC;
    border: 1px solid #475569;
    border-radius: 6px;
    padding: 6px 12px;
    font-weight: 600;
}

QToolButton:hover {
    background-color: #0EA5E9;
    border-color: #38BDF8;
}

QToolButton:pressed {
    background-color: #0284C7;
}

/* Splitter */
QSplitter::handle {
    background-color: #334155;
    width: 2px;
}

/* Tab Widget */
QTabWidget::pane {
    border: 1px solid #334155;
    background-color: #1E293B;
    border-radius: 8px;
}

QTabBar::tab {
    background-color: #0F172A;
    color: #94A3B8;
    padding: 10px 18px;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    margin-right: 4px;
    font-weight: 600;
}

QTabBar::tab:selected {
    background-color: #1E293B;
    color: #38BDF8;
    border-bottom: 2px solid #38BDF8;
}

QTabBar::tab:hover:!selected {
    background-color: #1E293B;
    color: #F1F5F9;
}

/* GroupBox & Panels */
QGroupBox {
    background-color: #1E293B;
    border: 1px solid #334155;
    border-radius: 8px;
    margin-top: 12px;
    padding-top: 14px;
    font-weight: bold;
    color: #38BDF8;
}

QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 6px;
}

/* PushButtons */
QPushButton {
    background-color: #2563EB;
    color: #FFFFFF;
    border: none;
    border-radius: 6px;
    padding: 8px 16px;
    font-weight: 600;
}

QPushButton:hover {
    background-color: #3B82F6;
}

QPushButton:pressed {
    background-color: #1D4ED8;
}

QPushButton:disabled {
    background-color: #334155;
    color: #64748B;
}

QPushButton#SuccessButton {
    background-color: #059669;
}
QPushButton#SuccessButton:hover {
    background-color: #10B981;
}

QPushButton#WarningButton {
    background-color: #D97706;
}
QPushButton#WarningButton:hover {
    background-color: #F59E0B;
}

/* Text Editors & Inputs */
QPlainTextEdit, QTextEdit, QLineEdit {
    background-color: #0F172A;
    color: #F8FAFC;
    border: 1px solid #334155;
    border-radius: 6px;
    padding: 8px;
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 12px;
}

QPlainTextEdit:focus, QTextEdit:focus, QLineEdit:focus {
    border: 1px solid #38BDF8;
}

/* Tree & List Views */
QTreeWidget, QListWidget, QTableWidget {
    background-color: #0F172A;
    color: #E2E8F0;
    border: 1px solid #334155;
    border-radius: 6px;
    gridline-color: #1E293B;
}

QHeaderView::section {
    background-color: #1E293B;
    color: #38BDF8;
    padding: 6px;
    border: 1px solid #334155;
    font-weight: bold;
}

QTreeWidget::item:selected, QTableWidget::item:selected {
    background-color: #0369A1;
    color: #FFFFFF;
}

/* ScrollBars */
QScrollBar:vertical {
    background: #0F172A;
    width: 10px;
    border-radius: 5px;
}
QScrollBar::handle:vertical {
    background: #334155;
    border-radius: 5px;
}
QScrollBar::handle:vertical:hover {
    background: #475569;
}

/* Statusbar */
QStatusBar {
    background-color: #0F172A;
    color: #94A3B8;
    border-top: 1px solid #334155;
}
"""
