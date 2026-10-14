from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT

out = Path(__file__).resolve().with_name("Muhammed_Tallat_CV_Proofread_Final.docx")

doc = Document()
sec = doc.sections[0]
sec.top_margin = Inches(0.55)
sec.bottom_margin = Inches(0.55)
sec.left_margin = Inches(0.7)
sec.right_margin = Inches(0.7)

styles = doc.styles
styles["Normal"].font.name = "Aptos"
styles["Normal"].font.size = Pt(9)
styles["Normal"].paragraph_format.space_after = Pt(2)

def set_cell_shading(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), fill)
    tcPr.append(shd)

def add_heading(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(7)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text.upper())
    r.bold = True
    r.font.name = "Aptos Display"
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(31, 78, 121)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '1F4E79')
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p

def add_bullet(text):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Inches(0.18)
    p.paragraph_format.first_line_indent = Inches(-0.12)
    p.paragraph_format.space_after = Pt(2)
    p.add_run(text)
    return p

def add_contact_item(paragraph, icon, label, text, url=None):
    icon_run = paragraph.add_run(f"{icon} ")
    icon_run.font.name = "Aptos"
    icon_run.font.size = Pt(8)
    icon_run.font.color.rgb = RGBColor(55, 55, 55)

    label_run = paragraph.add_run(f"{label}: ")
    label_run.bold = True
    label_run.font.size = Pt(8)
    label_run.font.color.rgb = RGBColor(55, 55, 55)

    if url:
        relationship_id = doc.part.relate_to(url, RT.HYPERLINK, is_external=True)
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("r:id"), relationship_id)
        run = OxmlElement("w:r")
        properties = OxmlElement("w:rPr")
        font = OxmlElement("w:rFonts")
        font.set(qn("w:ascii"), "Aptos")
        font.set(qn("w:hAnsi"), "Aptos")
        properties.append(font)
        size = OxmlElement("w:sz")
        size.set(qn("w:val"), "16")
        properties.append(size)
        color = OxmlElement("w:color")
        color.set(qn("w:val"), "373737")
        properties.append(color)
        run.append(properties)
        text_element = OxmlElement("w:t")
        text_element.text = text
        run.append(text_element)
        hyperlink.append(run)
        paragraph._p.append(hyperlink)
    else:
        text_run = paragraph.add_run(text)
        text_run.font.size = Pt(8)
        text_run.font.color.rgb = RGBColor(55, 55, 55)

# Header
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(1)
r = p.add_run("MUHAMMED TALLAT HASSAN AHMED")
r.bold = True
r.font.name = "Aptos Display"
r.font.size = Pt(17)
r.font.color.rgb = RGBColor(31, 78, 121)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(3)
r = p.add_run("FULL-STACK DEVELOPER")
r.bold = True
r.font.size = Pt(10)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(1)
p.paragraph_format.keep_together = True
add_contact_item(p, "✉", "Email", "muhammedtallat4@gmail.com", "mailto:muhammedtallat4@gmail.com")
p.add_run("  |  ")
add_contact_item(p, "☎", "Phone", "+20 102 818 2637", "tel:+201028182637")
p.add_run("  |  ")
add_contact_item(p, "⌖", "Location", "Nasr City, Cairo, Egypt")

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
p.paragraph_format.keep_together = True
add_contact_item(p, "in", "LinkedIn", "linkedin.com/in/muhammed-tallat-a440881b7",
                 "https://linkedin.com/in/muhammed-tallat-a440881b7")
p.add_run("  |  ")
add_contact_item(p, "GH", "GitHub", "github.com/MhmdTalat", "https://github.com/MhmdTalat")

add_heading("Professional Summary")
doc.add_paragraph(
    "Full-Stack Developer with 3+ years of experience building backend systems with ASP.NET Core and Django, "
    "alongside React front ends. Experienced in microservices, secure API development, and payment gateway "
    "integration with Paymob and Stripe. Delivered measurable improvements in performance, security, and "
    "operational efficiency across enterprise applications."
)

add_heading("Technical Skills")
skills = [
    ("Programming Languages", "C#, SQL, PL/SQL, Python, TypeScript, JavaScript"),
    ("Backend Development", "ASP.NET Core Web API, ASP.NET MVC, .NET 8, Entity Framework Core, LINQ, ADO.NET, Django REST Framework"),
    ("Databases", "Microsoft SQL Server, Oracle Database, PostgreSQL, stored procedures, query optimization, indexing, ERD design"),
    ("APIs & Integration", "RESTful API development, Swagger/OpenAPI, Postman, Paymob, Stripe API"),
    ("Architecture & Design", "Clean Architecture, SOLID, object-oriented programming (OOP), Repository Pattern, Unit of Work, CQRS, microservices"),
    ("Security", "JWT authentication, role-based access control (RBAC), input validation, data encryption"),
    ("Frontend Development", "React, HTML5, CSS3, Bootstrap, Tailwind CSS"),
    ("Tools & Practices", "Git, GitHub, Docker, Azure, CI/CD, unit testing, Agile"),
]
for label, value in skills:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(f"{label}: ")
    r.bold = True
    p.add_run(value)

add_heading("Professional Experience")
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(1)
r = p.add_run("Full-Stack Developer (.NET & React) ")
r.bold = True
p.add_run("| April 2025 – Present")
p = doc.add_paragraph("Egyptian Takaful Insurance")
p.paragraph_format.space_after = Pt(2)
p.runs[0].italic = True
for b in [
    "Architected 15+ RESTful web services using ASP.NET Core, improving system scalability and supporting 500+ daily active users across insurance policy and claims modules.",
    "Integrated the Paymob payment gateway through encrypted API channels, processing 200+ transactions per month and maintaining 99.9% uptime for critical financial operations.",
    "Implemented JWT-based authentication and role-based access control, reducing unauthorized access incidents by 95% across endpoints.",
    "Designed normalized entity-relationship diagrams (ERDs) and automated report-generation pipelines, reducing manual reporting effort by 60% (approximately 12 hours per week).",
    "Collaborated with QA and product teams during Agile sprints and conducted 50+ code reviews to support release stability."
]:
    add_bullet(b)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(4)
p.paragraph_format.space_after = Pt(1)
r = p.add_run("Back-End Developer (.NET) ")
r.bold = True
p.add_run("| June 2021 – September 2021")
p = doc.add_paragraph("Egyptian Projects Operation and Maintenance Company (EPROM)")
p.paragraph_format.space_after = Pt(2)
p.runs[0].italic = True
for b in [
    "Developed three internal ASP.NET MVC applications following enterprise development practices and supporting more than 100 operational staff.",
    "Optimized SQL Server queries through indexing and query restructuring, reducing average response time by 40%.",
    "Identified and resolved 30+ production bugs through systematic code reviews, improving deployment stability across releases."
]:
    add_bullet(b)

add_heading("Key Projects")
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(1)
r = p.add_run("Deepfake Media Detection ")
r.bold = True
p.add_run("| Graduation Project | Flutter, Python, CRNN")
add_bullet("Developed a cross-platform application to detect deepfake videos using CRNN algorithms, achieving an AUROC score of 89.35%. The research was published in IEEE Xplore (MIUCC-22).")

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(3)
p.paragraph_format.space_after = Pt(1)
r = p.add_run("E-Commerce Platform ")
r.bold = True
p.add_run("| ASP.NET Core, Entity Framework Core, SQL Server, JWT, Swagger, React, Stripe")
add_bullet("Developed a modular e-commerce platform featuring JWT authentication, a product catalog, cart management, and Stripe checkout, using Clean Architecture with repository and service layers.")

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(3)
p.paragraph_format.space_after = Pt(1)
r = p.add_run("Gold Market Analytics ")
r.bold = True
p.add_run("| React, Python, Django REST Framework, SQLite, Pandas, XGBoost")
add_bullet("Developed a full-stack gold market analytics application using React for the frontend and Django REST Framework for backend API development.")
add_bullet("Built React interfaces for visualizing historical gold prices, market trends, technical indicators, and machine learning-based price forecasts, integrating RESTful APIs to retrieve and display analytical results.")
add_bullet("Implemented an automated data ingestion pipeline with CSV normalization, OHLC validation, and idempotent bulk database upserts.")
add_bullet("Developed analytical endpoints for market returns, SMA, EMA, RSI, MACD, Bollinger Bands, ATR, drawdown analysis, and trading signals, alongside strategy backtesting and forecast evaluation with safeguards against look-ahead bias.")
add_bullet("Evaluated Linear Regression, Random Forest, and Gradient Boosting models using chronological train-validation-test splits, with model freshness checks and empirical prediction intervals for next-day price forecasts.")

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(3)
p.paragraph_format.space_after = Pt(1)
r = p.add_run("Egyptian National ID Field Detector (eKYC) ")
r.bold = True
p.add_run("| Python, Django, OpenCV, NumPy, Computer Vision")
add_bullet("Developed an identity-verification system for Egyptian national IDs using region-based field extraction and signature recognition, achieving 95% detection accuracy across varied image conditions.")

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(3)
p.paragraph_format.space_after = Pt(1)
r = p.add_run("E-Learning Management System ")
r.bold = True
p.add_run("| ASP.NET Core (.NET 8), Entity Framework Core, SQL Server, React, TypeScript")
add_bullet("Built a full-stack platform for managing departments, instructors, courses, and enrollments using the Repository Pattern, EF Core migrations, and a React front end.")

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(2)
p.paragraph_format.space_after = Pt(2)
p.add_run("For more details about my projects, visit my GitHub profile: github.com/MhmdTalat")

add_heading("Education")
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(1)
r = p.add_run("Bachelor of Science in Computing and Information ")
r.bold = True
p.add_run("| 2022")
doc.add_paragraph("Arab Academy for Science, Technology and Maritime Transport")

add_heading("Certifications & Publications")
for b in [
    "Oracle Data Platform 2025 Certified Foundations Associate — Oracle University (2025)",
    "IEEE Conference Paper: “Deepfake Video Detection” — MIUCC-22 (2022)",
    "Front-End Development — Information Technology Institute (ITI) (2021)",
    "CCNA 7: Introduction to Networks (88%) — National Telecommunication Institute (NTI) (2021)"
]:
    add_bullet(b)

add_heading("Key Achievements")
doc.add_paragraph(
    "Certificate of Contribution — MIUCC-22 Conference. Recognized for contributing to the 2nd International Mobile, "
    "Intelligent, and Ubiquitous Computing Conference, organized by Misr International University and sponsored by "
    "the IEEE Egypt Section on May 8–9, 2022. Contributed a research paper titled “Deepfake Video Detection,” "
    "applying artificial intelligence and deep learning techniques to identify manipulated video content."
)

# Footer page number
for section in doc.sections:
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer.add_run("Muhammed Tallat Hassan Ahmed | CV")
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(100, 100, 100)

doc.save(out)
print(out)
