"""
PowerPoint Copilot Service Facade and Sample Templates.
"""

from copilots_app.services.powerpoint.constants import VALID_SHAPE_TYPES, DEFAULT_DSL_COLORS, THEME_COLORS
from copilots_app.services.powerpoint.parser import parse_dsl, parse_dsl_slides, refresh_dsl_theme_colors
from copilots_app.services.powerpoint.connector import PowerPointConnector

PPT_SAMPLES = {
    "Overview Demo": """\
// Header banner
rect left=0 top=0 width=960 height=60 color=a4 | "Executive Summary & Strategy" size=22 bold=true color=#FFFFFF

// Metric Card 1
rounded_rect left=48 top=90 width=260 height=110 color=a6 outline=a1,1.5 border_radius=10
icon name=chart-line style=solid left=68 top=110 width=32 height=32 color=a1
rect left=110 top=105 width=180 height=40 color=a6 | "Total Revenue" size=12 color=t2
rect left=110 top=135 width=180 height=45 color=a6 | "$24.8M" size=22 bold=true color=a4

// Metric Card 2
rounded_rect left=350 top=90 width=260 height=110 color=a6 outline=a2,1.5 border_radius=10
icon name=users style=solid left=370 top=110 width=32 height=32 color=a2
rect left=412 top=105 width=180 height=40 color=a6 | "Active Clients" size=12 color=t2
rect left=412 top=135 width=180 height=45 color=a6 | "1,420" size=22 bold=true color=a4

// Metric Card 3
rounded_rect left=652 top=90 width=260 height=110 color=a6 outline=a3,1.5 border_radius=10
icon name=award style=solid left=672 top=110 width=32 height=32 color=a3
rect left=714 top=105 width=180 height=40 color=a6 | "Customer Satisfaction" size=12 color=t2
rect left=714 top=135 width=180 height=45 color=a6 | "98.4%" size=22 bold=true color=a4

// Deliverables Table
table left=48 top=230 width=864 height=240
cols=200,364,150,150
header="Workstream","Key Objective","Owner","Status"
row="Digital Transformation","Migrate legacy workflows to modern web cloud","Engineering","✓ On Track"
row="Data Quality Audit","Automated validation rules across all documents","Analytics","✓ Completed"
row="Enterprise Rollout","Staff enablement and training workshops","Operations","⏳ In Progress"
row="Security Compliance","Air-gapped deployment verification","Security","✓ Verified"

// Footer divider
line x1=48 y1=495 x2=912 y2=495 color=a3 weight=1.5 dash=solid
rect left=48 top=505 width=400 height=25 color=bg1 | "Confidential — Internal Strategy" size=10 color=t2
""",
    "Process Flow": """\
rect left=0 top=0 width=960 height=60 color=a4 | "4-Step Implementation Roadmap" size=20 bold=true color=#FFFFFF

// Step 1
rounded_rect left=48 top=180 width=180 height=160 color=a6 outline=a1,2 border_radius=8
icon name=magnifying-glass style=solid left=118 top=200 width=40 height=40 color=a1
rect left=58 top=255 width=160 height=60 color=a6 halign=center | "1. Discover" size=14 bold=true color=a4 + "\nRequirements elicitation" size=10 color=t2

// Arrow 1
chevron left=238 top=235 width=30 height=50 color=a1

// Step 2
rounded_rect left=278 top=180 width=180 height=160 color=a6 outline=a2,2 border_radius=8
icon name=code style=solid left=348 top=200 width=40 height=40 color=a2
rect left=288 top=255 width=160 height=60 color=a6 halign=center | "2. Build" size=14 bold=true color=a4 + "\nUnified architecture" size=10 color=t2

// Arrow 2
chevron left=468 top=235 width=30 height=50 color=a2

// Step 3
rounded_rect left=508 top=180 width=180 height=160 color=a6 outline=a3,2 border_radius=8
icon name=vial-circle-check style=solid left=578 top=200 width=40 height=40 color=a3
rect left=518 top=255 width=160 height=60 color=a6 halign=center | "3. Test" size=14 bold=true color=a4 + "\nAutomated validation" size=10 color=t2

// Arrow 3
chevron left=698 top=235 width=30 height=50 color=a3

// Step 4
rounded_rect left=738 top=180 width=180 height=160 color=a4 border_radius=8
icon name=rocket style=solid left=808 top=200 width=40 height=40 color=#FFFFFF
rect left=748 top=255 width=160 height=60 color=a4 halign=center | "4. Deploy" size=14 bold=true color=#FFFFFF + "\nSeamless distribution" size=10 color=a6
""",
}
