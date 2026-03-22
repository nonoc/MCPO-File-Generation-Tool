"""
Shared styles for ReportLab PDF generation.
"""

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, ListFlowable, ListItem, Spacer, Table as ReportLabTable, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# Base styles from reportlab
styles = getSampleStyleSheet()

# Custom Heading 1 - Main document titles
styles.add(ParagraphStyle(
    name="CustomHeading1",
    parent=styles["Heading1"],
    textColor=colors.HexColor("#0A1F44"),
    fontSize=18,
    spaceAfter=16,
    spaceBefore=12,
    alignment=TA_LEFT
))

# Custom Heading 2 - Section headings
styles.add(ParagraphStyle(
    name="CustomHeading2",
    parent=styles["Heading2"],
    textColor=colors.HexColor("#1C3F77"),
    fontSize=14,
    spaceAfter=12,
    spaceBefore=10,
    alignment=TA_LEFT
))

# Custom Heading 3 - Subsection headings
styles.add(ParagraphStyle(
    name="CustomHeading3",
    parent=styles["Heading3"],
    textColor=colors.HexColor("#3A6FB0"), 
    fontSize=12,
    spaceAfter=10,
    spaceBefore=8,
    alignment=TA_LEFT
))

# Custom Normal - Body text
styles.add(ParagraphStyle(
    name="CustomNormal",
    parent=styles["Normal"],
    fontSize=11,
    leading=14,
    alignment=TA_LEFT
))

# Custom List Item
styles.add(ParagraphStyle(
    name="CustomListItem",
    parent=styles["Normal"],
    fontSize=11,
    leading=14,
    alignment=TA_LEFT
))

# Custom Code block
styles.add(ParagraphStyle(
    name="CustomCode",
    parent=styles["Code"],
    fontSize=10,
    leading=12,
    fontName="Courier",
    backColor=colors.HexColor("#F5F5F5"),
    borderColor=colors.HexColor("#CCCCCC"),
    borderWidth=1,
    leftIndent=10,
    rightIndent=10,
    topPadding=5,
    bottomPadding=5
))

# Structured Document Title - Large title for structured content
styles.add(ParagraphStyle(
    name="StructuredDocumentTitle",
    parent=styles["CustomHeading1"],
    fontSize=24,
    leading=28,
    alignment=TA_LEFT,
    spaceBefore=0,
    spaceAfter=18
))

# Structured Paragraph - Body text with proper spacing
styles.add(ParagraphStyle(
    name="StructuredParagraph",
    parent=styles["CustomNormal"],
    leading=16,
    spaceBefore=4,
    spaceAfter=10
))

# Structured List Item - Nested list formatting
styles.add(ParagraphStyle(
    name="StructuredListItem",
    parent=styles["CustomListItem"],
    leftIndent=6,
    spaceBefore=2,
    spaceAfter=2
))

# Structured Sources Heading - Sources section title
styles.add(ParagraphStyle(
    name="StructuredSourcesHeading",
    parent=styles["CustomHeading2"],
    textColor=colors.HexColor("#1F3474"),
    fontSize=13,
    leading=16,
    alignment=TA_LEFT,
    spaceBefore=18,
    spaceAfter=6
))

# Structured Sources Item - Individual source references
styles.add(ParagraphStyle(
    name="StructuredSourcesItem",
    parent=styles["CustomNormal"],
    fontSize=10,
    italic=True,
    leading=14,
    leftIndent=4,
    spaceBefore=2,
    spaceAfter=4
))

# Table styles
def get_table_style():
    """Get the standard table style for PDF tables."""
    return TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.75, colors.HexColor("#9AA0B6")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8EEF9")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#0A1F44")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ])

def get_list_flowable(items, ordered=False, depth=1):
    """Create a ListFlowable from a list of items."""
    bullet_type = "1" if ordered else "bullet"
    return ListFlowable(
        items,
        bulletType=bullet_type,
        leftIndent=8 * mm * depth,
        bulletIndent=4 * mm,
        spaceBefore=8,
        spaceAfter=8,
    )