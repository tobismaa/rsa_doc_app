from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


BASE_DIR = Path(__file__).resolve().parent
OUTPUT_PDF = BASE_DIR / "quotation_mortgage_system_2026-06-17-v3.pdf"


def money(amount: str) -> Paragraph:
    return Paragraph(amount, STYLES["Amount"])


def text(value: str) -> Paragraph:
    return Paragraph(value, STYLES["Body"])


def bullet(value: str) -> Paragraph:
    return Paragraph(f"• {value}", STYLES["Body"])


SAMPLE = getSampleStyleSheet()
STYLES = {
    "Title": ParagraphStyle(
        "Title",
        parent=SAMPLE["Title"],
        fontName="Helvetica-Bold",
        fontSize=18,
        leading=22,
        textColor=colors.HexColor("#0f3b67"),
        spaceAfter=6,
    ),
    "SubTitle": ParagraphStyle(
        "SubTitle",
        parent=SAMPLE["BodyText"],
        fontName="Helvetica",
        fontSize=10.5,
        leading=14,
        textColor=colors.HexColor("#475569"),
        spaceAfter=10,
    ),
    "Section": ParagraphStyle(
        "Section",
        parent=SAMPLE["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=12,
        leading=15,
        textColor=colors.HexColor("#0f3b67"),
        spaceBefore=6,
        spaceAfter=6,
    ),
    "Body": ParagraphStyle(
        "Body",
        parent=SAMPLE["BodyText"],
        fontName="Helvetica",
        fontSize=10,
        leading=14,
        textColor=colors.HexColor("#1f2937"),
        spaceAfter=4,
    ),
    "Amount": ParagraphStyle(
        "Amount",
        parent=SAMPLE["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=10.5,
        leading=14,
        alignment=2,
        textColor=colors.HexColor("#111827"),
    ),
    "Total": ParagraphStyle(
        "Total",
        parent=SAMPLE["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=12,
        leading=16,
        textColor=colors.HexColor("#0f172a"),
    ),
}


def build_pdf() -> None:
    doc = SimpleDocTemplate(
        str(OUTPUT_PDF),
        pagesize=A4,
        leftMargin=22 * mm,
        rightMargin=22 * mm,
        topMargin=20 * mm,
        bottomMargin=20 * mm,
    )

    story = []

    story.append(Paragraph("RSA Portal App", STYLES["Title"]))
    story.append(Paragraph("Quotation", STYLES["SubTitle"]))
    story.append(Paragraph("Date: 17 June 2026", STYLES["Body"]))
    story.append(Spacer(1, 4))

    story.append(Paragraph("Quotation Breakdown", STYLES["Section"]))

    breakdown = Table(
        [
            [
                text("<b>1. Fully developed and already deployed system</b><br/>"
                     "This covers the current live mortgage workflow system and its major modules."),
                money("NGN 3,500,000"),
            ],
            [
                text("<b>2. Additional functionality and integration</b><br/>"
                     "This covers integration of the existing letter application, automatic document merging, and in-app automatic email trigger for sending applications to PFAs."),
                money("NGN 1,750,000"),
            ],
        ],
        colWidths=[128 * mm, 32 * mm],
    )
    breakdown.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#cbd5e1")),
                ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#e2e8f0")),
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f8fafc")),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 10),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
            ]
        )
    )
    story.append(breakdown)
    story.append(Spacer(1, 10))

    total = Table(
        [[Paragraph("Total Quotation", STYLES["Total"]), Paragraph("NGN 5,250,000", STYLES["Total"])]],
        colWidths=[128 * mm, 32 * mm],
    )
    total.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#dbeafe")),
                ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#93c5fd")),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    story.append(total)
    story.append(Spacer(1, 12))

    story.append(Paragraph("Included Existing System Modules", STYLES["Section"]))
    modules = [
        "Uploader dashboard",
        "Reviewer dashboard",
        "RSA dashboard",
        "Payment dashboard",
        "Admin dashboard",
        "Reports and monitoring dashboard",
        "Role-based workflow processing",
        "Document upload and tracking",
        "Status tracking and workflow movement",
        "User management and operational controls",
    ]
    for item in modules:
        story.append(bullet(item))

    story.append(Spacer(1, 8))
    story.append(Paragraph("Notes", STYLES["Section"]))
    notes = [
        "The amount above covers the current deployed system and the listed additional functionality.",
        "Any new features, third-party integrations, or process changes outside the scope listed above will be treated separately.",
        "Support and maintenance after handover can be agreed under a separate support arrangement if required.",
    ]
    for item in notes:
        story.append(bullet(item))

    doc.build(story)


if __name__ == "__main__":
    build_pdf()
