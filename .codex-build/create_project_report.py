from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pathlib import Path

ROOT = Path(r"C:\Users\patil\OneDrive\Desktop\College DA\Sem 3\DBMS\Project")
OUT = ROOT / "deliverables" / "Project_Work_and_Future_Enhancement.docx"
OUT.parent.mkdir(parents=True, exist_ok=True)

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Inches(8.27), Inches(11.69)
sec.top_margin = Inches(0.68)
sec.bottom_margin = Inches(0.68)
sec.left_margin = Inches(0.78)
sec.right_margin = Inches(0.78)
sec.header_distance = Inches(0.3)
sec.footer_distance = Inches(0.3)

NAVY = RGBColor(20, 36, 55)
TEAL = RGBColor(23, 149, 126)
GRAY = RGBColor(83, 99, 113)
BLACK = RGBColor(0, 0, 0)

def set_font(style, name="Aptos", size=10.5, color=BLACK, bold=False):
    style.font.name = name
    style._element.rPr.rFonts.set(qn("w:ascii"), name)
    style._element.rPr.rFonts.set(qn("w:hAnsi"), name)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    style.font.size = Pt(size)
    style.font.color.rgb = color
    style.font.bold = bold

normal = doc.styles["Normal"]
set_font(normal, size=10.5, color=BLACK)
normal.paragraph_format.space_after = Pt(5)
normal.paragraph_format.line_spacing = 1.08

title_style = doc.styles["Title"]
set_font(title_style, size=24, color=BLACK, bold=True)
title_style.paragraph_format.space_after = Pt(5)
title_style.paragraph_format.keep_with_next = True
title_ppr = title_style._element.get_or_add_pPr()
for border in title_ppr.findall(qn("w:pBdr")):
    title_ppr.remove(border)

for name, size in (("Heading 1", 15), ("Heading 2", 12)):
    st = doc.styles[name]
    set_font(st, size=size, color=BLACK, bold=True)
    st.paragraph_format.space_before = Pt(10 if name == "Heading 1" else 6)
    st.paragraph_format.space_after = Pt(3)
    st.paragraph_format.keep_with_next = True

def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shade = OxmlElement("w:shd")
    shade.set(qn("w:fill"), fill)
    tc_pr.append(shade)

def set_run_font(run, size=None, bold=None, color=None, italic=None):
    run.font.name = "Aptos"
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), "Aptos")
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), "Aptos")
    if size is not None: run.font.size = Pt(size)
    if bold is not None: run.bold = bold
    if color is not None: run.font.color.rgb = color
    if italic is not None: run.italic = italic

def add_body(text, bold_lead=None):
    p = doc.add_paragraph()
    p.paragraph_format.widow_control = True
    if bold_lead and text.startswith(bold_lead):
        r = p.add_run(bold_lead)
        set_run_font(r, bold=True, color=NAVY)
        r2 = p.add_run(text[len(bold_lead):])
        set_run_font(r2)
    else:
        r = p.add_run(text)
        set_run_font(r)
    return p

def add_bullet(text):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Inches(0.22)
    p.paragraph_format.first_line_indent = Inches(-0.14)
    p.paragraph_format.space_after = Pt(2.5)
    p.paragraph_format.line_spacing = 1.03
    r = p.add_run(text)
    set_run_font(r, size=10.2)
    return p

# Page 1: project work and model
p = doc.add_paragraph(style="Title")
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
p.add_run("Project Work and Future Enhancement")

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(3)
r = p.add_run("Modern Airline Logistics and Sustainability Management System")
set_run_font(r, size=14, bold=True, color=NAVY)

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(10)
r = p.add_run("Pranali Paygonda Patil, 25BCE2957, B.Tech Computer Science and Engineering, VIT SCOPE, AY 2026–27")
set_run_font(r, size=9.2, color=GRAY)

doc.add_heading("Project purpose", level=1)
add_body("This project develops a database-backed system for organizing airline logistics in one connected model. Airline work depends on accurate links among aircraft, flights, crew, passengers, cargo and operational events. When those records sit in separate lists, staff may face inconsistent information, slower coordination and limited visibility into delays or environmental impact. The project combines a relational database with a browser console so users can review and maintain key records in a shared operational view.")

doc.add_heading("Project objectives", level=1)
add_bullet("Integrate flight, aircraft, crew and passenger records in a relational model.")
add_bullet("Represent cold-chain requirements, flight-level fuel and CO₂, and delay events.")
add_bullet("Use keys and constraints to protect relationships and required data.")
add_bullet("Demonstrate how a shared database can support search, record maintenance and reporting.")

doc.add_heading("Work completed", level=1)
add_body("The Review 0 work defines a normalized relational design with nine entities: AIRCRAFT, FLIGHT, CREW, FLIGHT_CREW, LOYALTY_TIER, PASSENGER, SPEC_CARGO, EMISSIONS and DISRUPTION. FLIGHT acts as the operational anchor. Primary and foreign keys preserve record identity and relationships. UNIQUE, NOT NULL, CHECK and DEFAULT constraints help prevent invalid or duplicate data. The many-to-many relationship between flights and crew is represented through FLIGHT_CREW, which also records shift hours.")
add_body("The database coursework demonstrates table creation, inserts, updates and deletions, along with SELECT statements, single-row and aggregate functions, GROUP BY, HAVING, subqueries and joins. These operations support both routine record management and questions that combine information across related tables.")

doc.add_heading("Relational design in practice", level=1)
add_body("Each flight references one aircraft, and crew members connect to flights through FLIGHT_CREW. A loyalty tier can apply to many passengers, while each passenger references a tier. Cargo, emissions and disruption records point to the flight they describe. These relationships support queries across scheduling, staffing, baggage, cargo handling, environmental logging and delay history without repeating core flight details in every related record.")

doc.add_heading("Sustainability motivation", level=1)
add_body("The sustainability focus comes from making environmental impact visible alongside flight operations. EMISSIONS stores fuel consumed, CO₂ recorded and the time of each log, linked to a flight. This creates a basis for later route or aircraft comparisons. The project records and organizes sustainability data; it does not claim that the prototype has reduced emissions.")

# Page 2: practical implementation and value
doc.add_page_break()
doc.add_heading("Implementation and practical use", level=1)
add_body("The current application serves a single-page HTML, CSS and JavaScript interface through a local FastAPI service. Python’s python-oracledb driver connects to Oracle XE in Thin mode. The browser requests JSON data from the API, and create, update and delete actions are committed to Oracle. The README describes running the service with Uvicorn at 127.0.0.1:8001 after the local Oracle credentials and service name are configured.")

doc.add_heading("Operational features", level=1)
add_bullet("The console provides views for flights, aircraft, crew, passengers, specialized cargo, emissions and disruptions. Loyalty tiers serve as a passenger lookup.")
add_bullet("Users can search records, filter by status, export CSV files, inspect record details and use forms to add or edit records. The interface refreshes Oracle data every five seconds.")
add_bullet("Cargo records include category, weight, required temperature, current temperature and tracking status. Disruption records capture a flight, reason, delay duration, event time and status.")
add_bullet("Emission records link fuel and CO₂ values to a flight. SQL joins and grouped summaries can use these links to compare routes or investigate patterns.")
add_bullet("Account registration and sign-in protect the operational API. Passwords use scrypt hashes, and sessions use a signed HttpOnly cookie with a 12-hour expiry.")

doc.add_heading("Business and operational relevance", level=1)
add_body("A shared flight-linked record can help operations staff review aircraft assignments and disruption status without reconciling separate lists. Cargo coordinators can compare a shipment’s required and current temperature. Sustainability analysts can prepare route-level comparisons from fuel and CO₂ logs. Managers can group disruption events by reason or status to identify recurring causes for further investigation. These are potential uses of the data model, not measured improvements from a live airline deployment.")

doc.add_heading("Current scope and limitations", level=1)
add_body("The Review 0 document lists MySQL 8.0, while the current runnable code and README use Oracle XE. The database design describes FLIGHT_CREW, but the browser application does not provide a dedicated screen for managing crew-to-flight assignments. User accounts currently share the same access level, the README limits use to localhost coursework, and Oracle key generation must be configured for add forms when an ID is omitted. Bundled browser sample records and the overview chart are illustrative; use live Oracle data for operational reporting.")

# Page 3: roadmap
doc.add_page_break()
doc.add_heading("Future enhancements", level=1)
add_body("The next stage should improve data quality and operational usefulness before expanding deployment. Enhancements below are proposals and are not part of the current implementation.")

doc.add_heading("1. Strengthen validation and access control", level=2)
add_body("Add role-based permissions for operations, cargo and sustainability users. Validate date order, numeric ranges, status values and required fields in both the interface and API. Define an explicit identity strategy for every table and retain an audit history of who changed each record and when. These changes would make the system safer to use and easier to review.")

doc.add_heading("2. Add live cold-chain monitoring", level=2)
add_body("Extend cargo records to hold timestamped sensor readings rather than a single current temperature. Let staff configure acceptable ranges by cargo type, raise an alert when a reading crosses a limit, and record acknowledgement and resolution. A pilot could begin with one shipment category and compare sensor readings with the existing manual workflow.")

doc.add_heading("3. Build sustainability analysis", level=2)
add_body("Create dashboards for fuel and CO₂ by flight, route and aircraft. Record the source and method used for each emissions value, and add route distance or payload information before calculating intensity measures. This would make comparisons more meaningful and help analysts distinguish a higher total caused by a longer route from a less efficient operation.")

doc.add_heading("4. Improve disruption and fleet planning", level=2)
add_body("Add summaries for delay duration and disruption reason, then connect those findings with schedule and aircraft data. Later versions could incorporate weather or air-traffic feeds to help staff see developing risks. Any recommendation feature should explain which records support it and leave operational decisions with authorized staff.")

doc.add_heading("5. Prepare for reliable deployment", level=2)
add_body("Before use beyond a local demonstration, add protected secret storage, encrypted network traffic, structured logs, monitoring, backups and recovery procedures. Test database constraints and API behavior with representative cases, document setup for a clean environment, and integrate other systems only through defined and access-controlled interfaces.")

doc.add_heading("Conclusion", level=1)
add_body("The project establishes a practical foundation for connected airline logistics records and introduces sustainability and cold-chain data into the operational model. Its strongest next steps are reliable validation, traceable live cargo readings and clearly sourced emissions analysis. With these improvements, the prototype could provide more dependable information for airline coordination while preserving a clear record of how each operational decision is supported.")

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(7)
p.paragraph_format.space_after = Pt(0)
r = p.add_run("Project basis: Review 0 report, current README, server.py and index.html.")
set_run_font(r, size=8.5, color=GRAY, italic=True)

doc.core_properties.title = "Project Work and Future Enhancement"
doc.core_properties.subject = "Modern Airline Logistics and Sustainability Management System"
doc.core_properties.author = "Pranali Paygonda Patil"
doc.save(OUT)
print(OUT)
