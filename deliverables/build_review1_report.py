from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
from PIL import Image, ImageDraw, ImageFont

ROOT=Path.cwd(); OUT=ROOT/'deliverables'/'Review_1_Project_Implementation_Report.docx'
ASSET=ROOT/'deliverables'/'review1_assets'; ASSET.mkdir(exist_ok=True)
NAVY='10283F'; TEAL='159D87'; LIGHT='F2F6F8'; MUTED='607487'; TEXT='23384A'
def font(sz,bold=False):
    f=Path(r'C:\Windows\Fonts\arialbd.ttf' if bold else r'C:\Windows\Fonts\arial.ttf')
    return ImageFont.truetype(str(f),sz) if f.exists() else ImageFont.load_default()
def make_diagrams():
    # ER diagram: all nine Review 0 entities and their central foreign-key relationships.
    W,H=1700,1000; im=Image.new('RGB',(W,H),'white'); d=ImageDraw.Draw(im)
    boxes={'AIRCRAFT':(45,70,365,235),'FLIGHT':(690,70,1010,235),'CREW':(1335,70,1655,235),
    'LOYALTY_TIER':(45,390,365,555),'FLIGHT_CREW':(690,390,1010,555),'PASSENGER':(1335,390,1655,555),
    'SPEC_CARGO':(45,735,365,900),'EMISSIONS':(690,735,1010,900),'DISRUPTION':(1335,735,1655,900)}
    fields={'AIRCRAFT':'PK Aircraft_ID\nModel · Capacity · Efficiency','FLIGHT':'PK Flight_ID\nRoute · Dep_Time · Arr_Time\nFK Aircraft_ID · Status','CREW':'PK Crew_ID\nName · Role · Base','LOYALTY_TIER':'PK Tier_ID\nUNIQUE Tier_Name\nExtra_Baggage_kg','FLIGHT_CREW':'PK Assignment_ID\nFK Flight_ID · FK Crew_ID\nShift_Hours','PASSENGER':'PK Passenger_ID\nPassport_No · Baggage_kg\nFK Tier_ID','SPEC_CARGO':'PK Cargo_ID · FK Flight_ID\nDescription · Weight · Category\nRequired/Current temperature','EMISSIONS':'PK Log_ID · FK Flight_ID\nFuel_Liters · CO2_Tons · Logged_At','DISRUPTION':'PK Event_ID · FK Flight_ID\nDelay_Min · Reason_Code\nEvent_Date · Event_Status'}
    for a,b in [((365,150),(690,150)),((365,470),(1335,470)),((850,235),(850,390)),((1495,235),(850,390)),((850,235),(205,735)),((850,235),(850,735)),((850,235),(1495,735))]:
        d.line((a,b),fill='#73879A',width=4)
    for name,(x1,y1,x2,y2) in boxes.items():
        d.rounded_rectangle((x1,y1,x2,y2),radius=18,fill='#F6FAFC',outline='#159D87',width=4)
        d.rounded_rectangle((x1,y1,x2,y1+50),radius=18,fill='#10283F');d.rectangle((x1,y1+30,x2,y1+50),fill='#10283F')
        d.text((x1+16,y1+10),name,font=font(25,True),fill='white')
        d.multiline_text((x1+16,y1+67),fields[name],font=font(19),fill='#23384A',spacing=7)
    im.save(ASSET/'er_diagram.png')
    # Architecture diagram.
    W,H=1550,520; im=Image.new('RGB',(W,H),'white');d=ImageDraw.Draw(im)
    for title,body,x in [('BROWSER','HTML · CSS · JavaScript\nOperator workspace',55),('FASTAPI','Python service\nAuth · API · data mapping',585),('ORACLE XE','Relational database\nNine domain entities',1115)]:
        d.rounded_rectangle((x,100,x+375,365),radius=24,fill='#F4F8FA',outline='#159D87',width=5)
        d.rounded_rectangle((x,100,x+375,170),radius=24,fill='#10283F');d.rectangle((x,145,x+375,170),fill='#10283F')
        d.text((x+22,120),title,font=font(29,True),fill='white');d.multiline_text((x+22,220),body,font=font(25),fill='#23384A',spacing=12)
    for x in [440,970]:
        d.line((x,235,x+115,235),fill='#159D87',width=7);d.polygon([(x+115,220),(x+140,235),(x+115,250)],fill='#159D87')
    d.text((465,420),'HTTP / JSON   •   python-oracledb Thin mode   •   bound values',font=font(22),fill='#607487')
    im.save(ASSET/'architecture.png')
    # Workflow diagram.
    W,H=1550,400;im=Image.new('RGB',(W,H),'white');d=ImageDraw.Draw(im)
    steps=['Sign in','Load Oracle\ndata','Choose module\nand record','Create / edit /\ndelete','Commit to\nOracle','Refresh view\nand export']
    for i,s in enumerate(steps):
        x=20+i*255;d.rounded_rectangle((x,105,x+220,295),radius=20,fill='#F4F8FA',outline='#159D87',width=4)
        d.multiline_text((x+20,165),s,font=font(22,True),fill='#10283F',spacing=8)
        if i<5:
            xx=x+225;d.line((xx,200,xx+20,200),fill='#159D87',width=5);d.polygon([(xx+20,190),(xx+35,200),(xx+20,210)],fill='#159D87')
    im.save(ASSET/'workflow.png')
make_diagrams()

doc=Document(); sec=doc.sections[0]
sec.top_margin=Inches(.72);sec.bottom_margin=Inches(.68);sec.left_margin=Inches(.82);sec.right_margin=Inches(.82);sec.header_distance=Inches(.35);sec.footer_distance=Inches(.35);sec.different_first_page_header_footer=True
styles=doc.styles; normal=styles['Normal'];normal.font.name='Times New Roman';normal.font.size=Pt(10.5);normal.font.color.rgb=RGBColor.from_string(TEXT);normal.paragraph_format.space_after=Pt(6);normal.paragraph_format.line_spacing=1.12
for sn,size in [('Title',24),('Heading 1',16),('Heading 2',12),('Heading 3',11)]:
 st=styles[sn];st.font.name='Times New Roman';st.font.size=Pt(size);st.font.bold=True;st.font.color.rgb=RGBColor.from_string('000000');st.paragraph_format.space_before=Pt(10);st.paragraph_format.space_after=Pt(5);st.paragraph_format.keep_with_next=True
code=styles.add_style('Code Listing',WD_STYLE_TYPE.PARAGRAPH);code.font.name='Consolas';code.font.size=Pt(7);code.font.color.rgb=RGBColor.from_string(TEXT);code.paragraph_format.space_after=Pt(0);code.paragraph_format.line_spacing=1
h=sec.header.paragraphs[0];h.alignment=WD_ALIGN_PARAGRAPH.RIGHT;r=h.add_run('BACSE202  |  REVIEW 1  |  AIRLINE LOGISTICS & SUSTAINABILITY');r.font.name='Arial';r.font.size=Pt(8);r.font.bold=True;r.font.color.rgb=RGBColor.from_string(MUTED)
def field(run,name):
    for typ,val in [('begin',None),('instr',name),('separate',None),('text','1'),('end',None)]:
        e=OxmlElement('w:fldChar' if typ!='instr' and typ!='text' else ('w:instrText' if typ=='instr' else 'w:t'))
        if typ in ('begin','separate','end'):e.set(qn('w:fldCharType'),typ)
        elif typ=='instr':e.set(qn('xml:space'),'preserve');e.text=val
        elif typ=='text':e.text=val
        run._r.append(e)
f=sec.footer.paragraphs[0];f.alignment=WD_ALIGN_PARAGRAPH.CENTER;r=f.add_run('25BCE2957  •  ');r.font.size=Pt(8);r.font.color.rgb=RGBColor.from_string(MUTED);field(f.add_run(),'PAGE')
def shade(cell,color):
    sh=OxmlElement('w:shd');sh.set(qn('w:fill'),color);cell._tc.get_or_add_tcPr().append(sh)
def margins(cell):
    mar=OxmlElement('w:tcMar')
    for k,v in [('top',100),('start',100),('bottom',100),('end',100)]:
        e=OxmlElement('w:'+k);e.set(qn('w:w'),str(v));e.set(qn('w:type'),'dxa');mar.append(e)
    cell._tc.get_or_add_tcPr().append(mar)
def table(head,rows,widths=None,fs=8):
    t=doc.add_table(rows=1,cols=len(head));t.alignment=WD_TABLE_ALIGNMENT.CENTER;t.style='Table Grid';t.autofit=False
    for j,v in enumerate(head):
        c=t.rows[0].cells[j];c.text=str(v);shade(c,NAVY);margins(c);c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        for r in c.paragraphs[0].runs:r.font.name='Arial';r.font.size=Pt(fs);r.font.bold=True;r.font.color.rgb=RGBColor(255,255,255)
    for i,row in enumerate(rows):
        cells=t.add_row().cells
        for j,v in enumerate(row):
            c=cells[j];c.text=str(v);margins(c);c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if i%2:shade(c,LIGHT)
            for p in c.paragraphs:
                p.paragraph_format.space_after=Pt(1)
                for r in p.runs:r.font.name='Arial';r.font.size=Pt(fs);r.font.color.rgb=RGBColor.from_string(TEXT)
    if widths:
        for row in t.rows:
            for j,w in enumerate(widths):row.cells[j].width=Inches(w)
    doc.add_paragraph().paragraph_format.space_after=Pt(1);return t
def para(s='',align=None,italic=False,boldlead=None):
    p=doc.add_paragraph()
    if align is not None:p.alignment=align
    if boldlead and s.startswith(boldlead):p.add_run(boldlead).bold=True;p.add_run(s[len(boldlead):])
    else:p.add_run(s).italic=italic
    return p
def head(s,l=1):doc.add_heading(s,level=l)
def page():doc.add_page_break()
def image(path,w=6.5):
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.add_run().add_picture(str(path),width=Inches(w))
def caption(s):
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_after=Pt(8);r=p.add_run(s);r.italic=True;r.font.size=Pt(9);r.font.color.rgb=RGBColor.from_string(MUTED)

# Cover page
p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_before=Pt(12);r=p.add_run('SCHOOL OF COMPUTER SCIENCE AND ENGINEERING (SCOPE)');r.bold=True;r.font.size=Pt(15);r.font.color.rgb=RGBColor.from_string(NAVY)
para('Vellore Institute of Technology',WD_ALIGN_PARAGRAPH.CENTER)
p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_before=Pt(32);r=p.add_run('REVIEW 1  |  ASSESSMENT 9');r.bold=True;r.font.size=Pt(14);r.font.color.rgb=RGBColor.from_string(TEAL)
p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_before=Pt(14);r=p.add_run('MODERN AIRLINE LOGISTICS AND\nSUSTAINABILITY MANAGEMENT SYSTEM');r.bold=True;r.font.size=Pt(22);r.font.color.rgb=RGBColor.from_string(NAVY)
para('A Project Report',WD_ALIGN_PARAGRAPH.CENTER,True);para('Submitted in partial fulfilment of the requirements for the BACSE202 Lab Component Project',WD_ALIGN_PARAGRAPH.CENTER)
p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_before=Pt(20);p.add_run('Submitted by').italic=True
p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;r=p.add_run('25BCE2957  |  PRANALI PAYGONDA PATIL');r.bold=True;r.font.size=Pt(15);r.font.color.rgb=RGBColor.from_string(NAVY)
para('B.Tech Computer Science and Engineering',WD_ALIGN_PARAGRAPH.CENTER);para('Project Guide: Dr. Kavitha G',WD_ALIGN_PARAGRAPH.CENTER);para('Head of School: Dr. Boominathan P',WD_ALIGN_PARAGRAPH.CENTER)
p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_before=Pt(30);p.add_run('Academic Year 2026–27').bold=True
page()

# Front matter with live TOC and figure list.
head('CONTENT',1);p=doc.add_paragraph();r=p.add_run()
for typ,val in [('begin',None),('instr','TOC \\o "1-3" \\h \\z \\u'),('separate',None)]:
    e=OxmlElement('w:fldChar' if typ!='instr' else 'w:instrText')
    if typ in ('begin','separate'):e.set(qn('w:fldCharType'),typ)
    else:e.set(qn('xml:space'),'preserve');e.text=val
    r._r.append(e)
r.add_text('');e=OxmlElement('w:fldChar');e.set(qn('w:fldCharType'),'end');r._r.append(e)
page()
head('LIST OF FIGURES',1)
for i,s in enumerate(['ER diagram for nine domain entities','Overall system architecture','Operational workflow','Sign-in page','Overview dashboard','Account registration page','Flight schedule','Oracle flight rows before insert','Add flight form','Oracle flight rows after insert','Flight detail','Edit flight form','Oracle updated record','Delete confirmation prompt','Fleet module','Crew roster module','Cargo tracking module','Passengers and loyalty module','Disruptions module','Emissions module','Updated overview dashboard'],1):para(f'Figure {i}. {s}')
page()

head('ABSTRACT',1)
para('The Modern Airline Logistics and Sustainability Management System is a database-backed application for viewing and maintaining interconnected airline operations. It brings flight schedules, aircraft, crew, passenger loyalty, specialized cargo, emissions and disruption records into one operational workspace. The project combines the relational design documented in Review 0 with a working browser interface and a local FastAPI service connected to Oracle XE.')
para('The data model contains nine domain entities: AIRCRAFT, FLIGHT, CREW, FLIGHT_CREW, LOYALTY_TIER, PASSENGER, SPEC_CARGO, EMISSIONS and DISRUPTION. Primary and foreign keys and integrity constraints support consistent data. The web interface provides sign-in and registration, an overview dashboard, operational modules, search and status filtering, CSV export, and create, view, update and delete actions for supported records. Passwords are stored as scrypt hashes, and authenticated sessions use signed HttpOnly cookies.')
para('The sustainability focus is practical: fuel and CO2 records can be reviewed by flight, while specialized cargo records retain handling and temperature information. The system does not itself reduce emissions or control aircraft; it provides an information layer to support monitoring and decisions. Screenshots show the implementation with the supplied Oracle dataset, including create and update examples. The delete capture documents the confirmation step. Security, data quality, scalability and sustainability analytics remain future work.')
para('Keywords: airline logistics, Oracle XE, FastAPI, relational database, flight operations, emissions tracking, cold-chain cargo, sustainability.')
page()
head('ACKNOWLEDGEMENT',1)
para('I would like to express my sincere gratitude to Dr. Boominathan P, Head of the School of Computer Science and Engineering, for reviewing the project, and to my project guide, Dr. Kavitha G, for guidance throughout its development. I also thank the faculty members, classmates and others who offered feedback and support during the design, implementation and documentation of this work.')
page()

head('1 INTRODUCTION OF THE PROJECT',1);head('1.1 Project overview',2)
para('Airline operations depend on coordinated decisions about flight schedules, aircraft availability, crew assignments, passengers, cargo and disruption response. Operational facts are related: a flight uses an aircraft, may carry several cargo shipments, can generate multiple emission logs, and may have one or more disruption events. A reliable data model must represent these links and allow staff to retrieve current records without duplicating data across disconnected lists.')
para('This project develops a database-oriented airline operations workspace. The browser interface organizes information into distinct modules, while a local Python service reads and writes records in Oracle. Review 0 established the relational model and SQL coursework requirements; Review 1 adds application implementation and visual evidence.')
head('1.2 Problem statement',2)
para('When flight, fleet, crew, cargo and environmental information are maintained separately, staff may face inconsistent records, delayed updates and limited visibility into network status. Temperature-sensitive cargo requires explicit handling information. Emission reporting requires fuel and CO2 values to be associated with flights. Disruption records need consistent reason, delay and resolution fields. The project addresses these information-management needs through a centralized relational database and integrated interface.')
head('1.3 Objectives and scope',2)
table(['Objective','Project scope'],[('Centralize records','Linked tables and dedicated views for principal airline entities.'),('Support maintenance','Authenticated create, view, update and delete workflows for supported records.'),('Preserve consistency','Relational keys and constraints; parameterized Oracle binds in the service.'),('Improve inspection','Overview metrics, lists, search, status filters, refresh and CSV export.'),('Surface sustainability data','Store/display fuel and CO2 logs and cold-chain temperature fields.'),('Demonstrate DBMS concepts','Document schema design and SQL operations required by the review rubric.')],[1.8,4.8])
head('1.4 Scope boundaries',2)
para('The implementation is a local coursework demonstrator. It is not connected to airline reservation, airport, aircraft telemetry, payment or regulatory systems. Screenshots use demonstration records. The project does not claim real-time aircraft locations, live booking, automated crew legality checks, predictive disruption analysis or verified carbon accounting.')

head('2 BACKGROUND OF THE PROJECT WORK',1);head('2.1 Operational and database context',2)
para('Airline information is naturally relational. Flight records refer to aircraft; crew-to-flight assignments form a many-to-many relationship; passengers refer to loyalty tiers; cargo, emission logs and disruptions reference flights. A normalized schema reduces repeated descriptive information and makes referential integrity enforceable. Review 0 describes the nine-entity design and illustrates DDL, DML, query functions and joins.')
head('2.2 Motivation and sustainability inspiration',2)
para('The sustainability inspiration is to make environmental information part of everyday operations data rather than treating it as an unrelated report. Recording fuel and CO2 at flight level gives analysts a starting point for comparing routes, aircraft and operating periods. Recording required and observed cargo temperatures supports visibility for medical and perishable shipments. Better information can inform decisions, but the application itself does not prove a reduction in fuel burn or emissions.')
head('2.3 Review 0 to Review 1 progression',2)
table(['Review 0 foundation','Review 1 implementation'],[('Problem context, objectives and requirements','Local web app with registration/sign-in and modules'),('Nine-entity ER and relational schema','Oracle XE connectivity through FastAPI and python-oracledb'),('Keys, constraints and SQL coursework coverage','Browser CRUD for supported entities; search, filters, CSV export'),('Illustrative sample rows and query examples','Application screenshots and Oracle create/update evidence')],[2.8,3.8])
head('2.4 Database platform clarification',2)
para('Review 0 names MySQL 8.0 and MySQL Workbench in its software-requirements table. The delivered application code and README instead connect to Oracle XE through python-oracledb in Thin mode. This report describes the running implementation as Oracle XE. SQL syntax and key-generation behavior must follow the Oracle schema actually installed; the MySQL entry is treated as an initial coursework requirement, not the application runtime.')

head('3 HARDWARE AND SOFTWARE REQUIREMENTS',1)
table(['Category','Minimum / current setup','Recommended extension'],[('Processor','Intel Core i3 equivalent','Intel Core i5 or higher'),('Memory','4 GB RAM','8 GB RAM or higher'),('Storage','20 GB free','SSD for database, backups and logs'),('Operating system','Windows 10/11, Linux or macOS','Supported 64-bit OS with updates'),('Database','Oracle XE schema used by app','Managed Oracle with backup/monitoring'),('Backend','Python 3.10+; FastAPI; python-oracledb','Pinned dependencies; production ASGI'),('Frontend','Modern browser; HTML/CSS/JavaScript','Current Chrome, Edge or Firefox'),('Network','Local TCP access to Oracle','TLS-protected private network')],[1.2,2.7,2.7],7.8)
head('3.1 Runtime configuration',2)
para('The README instructs the operator to install requirements.txt dependencies, configure Oracle credentials and service DSN in a private .env file, and start the application on 127.0.0.1:8001. The browser should use the FastAPI-served page rather than opening index.html directly. The configured service must point to the same Oracle container/schema as the airline tables. Credentials and the session secret must remain private and must not be included in a report or shared repository.')

head('4 E-R DIAGRAM',1)
para('This ER model reproduces the nine domain entities documented in Review 0. FLIGHT is the operational hub. FLIGHT_CREW resolves the many-to-many relationship between flights and crew. Foreign keys connect related records. The authentication support table is separate from the nine-entity domain model.')
image(ASSET/'er_diagram.png');caption('Figure 1. ER diagram for the nine domain entities; PK/FK labels identify keys.')
head('4.1 Relationship cardinalities',2)
table(['Parent','Relationship','Child','Cardinality / meaning'],[('AIRCRAFT','assigned to','FLIGHT','1:M — each flight uses one aircraft; aircraft can operate flights over time.'),('FLIGHT','assigned through','FLIGHT_CREW','1:M — a flight can have multiple crew assignments.'),('CREW','assigned through','FLIGHT_CREW','1:M — a crew member can have multiple assignment records.'),('LOYALTY_TIER','assigned to','PASSENGER','1:M — each passenger has a tier; a tier can include many passengers.'),('FLIGHT','carries','SPEC_CARGO','1:M — a flight can carry multiple shipments.'),('FLIGHT','generates','EMISSIONS','1:M — a flight can have multiple emission logs.'),('FLIGHT','experiences','DISRUPTION','1:M — a flight can have multiple disruption events.')],[1.1,1.5,1.3,2.8],7.8)

head('5 DESCRIPTION ABOUT DESIGN OF THE DATABASE',1);head('5.1 Design approach',2)
para('The design separates entities by responsibility and uses primary keys to identify records. Foreign keys express operational relationships. Unique constraints are appropriate for tier names and passport numbers. NOT NULL constraints protect required fields, CHECK constraints constrain valid ranges or enumerated statuses, and DEFAULT constraints supply initial values where appropriate. The database remains the final authority for referential integrity during application writes.')
head('5.2 Entity summaries',2)
table(['Entity','Purpose','Important links / integrity'],[('AIRCRAFT','Fleet model, capacity, efficiency, service status','PK Aircraft_ID; referenced by FLIGHT.'),('FLIGHT','Route, departure/arrival, aircraft and status','PK Flight_ID; FK Aircraft_ID.'),('CREW','Crew identity, role and base','PK Crew_ID; linked by FLIGHT_CREW.'),('FLIGHT_CREW','Assignment bridge and shift hours','FK Flight_ID + FK Crew_ID; resolves M:N.'),('LOYALTY_TIER','Tier label and extra baggage allowance','PK Tier_ID; Tier_Name unique.'),('PASSENGER','Passenger, passport, tier and baggage','Passport_No unique; FK Tier_ID.'),('SPEC_CARGO','Shipment category, weight and temperatures','FK Flight_ID; status and temperature fields.'),('EMISSIONS','Fuel and CO2 by timestamp','FK Flight_ID; route/flight analysis.'),('DISRUPTION','Delay, reason, date and status','FK Flight_ID; event tracking.')],[1.3,2.5,2.6],7.6)
head('5.3 Normalization and integrity',2)
para('Separating aircraft, flights, crew, tiers, passengers and flight-linked records avoids storing repeated descriptive attributes in the flight table. FLIGHT_CREW prevents multi-valued crew lists in a single field. Bound values in the API reduce SQL injection risk for values, while database constraints reject invalid relationships and domains.')

head('6 BACKEND DATABASE TABLES',1)
para('This inventory presents the Review 0 attributes and logical Oracle-friendly types. Exact deployed widths should be verified against Oracle metadata before a production migration; undocumented lengths are not inferred here. Key and uniqueness requirements follow the project schema.')
rows=[('AIRCRAFT','AIRCRAFT_ID','NUMBER','MODEL VARCHAR2; CAPACITY_KG NUMBER; FUEL_EFFICIENCY NUMBER; AIRCRAFT_STATUS VARCHAR2','PK; valid capacity/status'),('FLIGHT','FLIGHT_ID','NUMBER','ROUTE_CODE VARCHAR2; DEP_TIME TIMESTAMP; ARR_TIME TIMESTAMP; AIRCRAFT_ID NUMBER; FLIGHT_STATUS VARCHAR2','PK; FK AIRCRAFT_ID; required fields'),('CREW','CREW_ID','NUMBER','CREW_NAME VARCHAR2; ROLE VARCHAR2; BASE_LOCATION VARCHAR2','PK; required identifying fields'),('FLIGHT_CREW','ASSIGNMENT_ID','NUMBER','FLIGHT_ID NUMBER; CREW_ID NUMBER; SHIFT_HOURS NUMBER','PK; two FKs; valid hours'),('LOYALTY_TIER','TIER_ID','NUMBER','TIER_NAME VARCHAR2; EXTRA_BAGGAGE_KG NUMBER','PK; unique name; nonnegative allowance'),('PASSENGER','PASSENGER_ID','NUMBER','PASSENGER_NAME VARCHAR2; PASSPORT_NO VARCHAR2; TIER_ID NUMBER; BAGGAGE_KG NUMBER','PK; unique passport; FK tier'),('SPEC_CARGO','CARGO_ID','NUMBER','FLIGHT_ID NUMBER; DESCRIPTION VARCHAR2; WEIGHT_KG NUMBER; CATEGORY VARCHAR2; REQ_TEMP_C NUMBER; CURRENT_TEMP_C NUMBER; TRACKING_STATUS VARCHAR2','PK; FK flight; nonnegative weight'),('EMISSIONS','LOG_ID','NUMBER','FLIGHT_ID NUMBER; FUEL_LITERS NUMBER; CO2_TONS NUMBER; LOGGED_AT TIMESTAMP','PK; FK flight; nonnegative values'),('DISRUPTION','EVENT_ID','NUMBER','FLIGHT_ID NUMBER; DELAY_MIN NUMBER; REASON_CODE VARCHAR2; EVENT_DATE TIMESTAMP; EVENT_STATUS VARCHAR2','PK; FK flight; nonnegative delay')]
table(['Table','PK','Type','Other attributes (logical types)','Constraints / links'],rows,[.9,.9,.6,2.9,1.2],6.8)
para('Implementation note: the FastAPI service also creates APP_LOGIN_USERS as needed. It contains identity-generated USER_ID, NAME, EMAIL and PASSWORD_HASH for authentication, separate from the nine airline-domain entities.',boldlead='Implementation note:')

head('7 EXECUTION OF SQL QUERIES WITH OUTPUT',1)
para('Review 0 contains the coursework SQL section: table creation with constraints, five-row inserts, delete and update demonstrations for the domain tables, SELECT, single-row functions, aggregate functions, GROUP BY, HAVING, subqueries and joins. The application screenshots add Oracle terminal evidence for flight records and the live create/update path. The application issues parameterized Oracle SELECT, INSERT, UPDATE and DELETE statements through python-oracledb.')
table(['Required operation','Evidence / purpose'],[('CREATE TABLE + constraints','Review 0 DDL; PK/FK and domain integrity establish schema.'),('INSERT up to five sample rows','Review 0 sample rows; app insert shown for flight ID 309.'),('DELETE','Review 0 deletion examples; UI confirmation prompt shown.'),('UPDATE','Review 0 per-table examples; flight 309 change shown in app and query output.'),('SELECT','Review 0 examples and live Oracle query outputs.'),('Single-row functions','Review 0 examples for text, date and numeric transformations.'),('Aggregate functions','Review 0 examples for counts, totals and averages.'),('GROUP BY / HAVING','Review 0 grouped summaries and filtered groups.'),('Subqueries','Review 0 nested query examples.'),('JOIN','Review 0 joins across aircraft/flights, passengers/tiers, crew and flight assignments, cargo and disruptions.')],[1.7,4.9],7.8)
head('7.1 Create and update evidence',2)
para('The supplied captures show eight initial flight rows, an Add Flight form for ID 309, and query output after insertion with nine rows. A detail view shows NS 309. The edit form changes its schedule and status to DELAYED; the subsequent Oracle output reflects that updated record. For this example, the evidence links the UI action to the database result.')
head('7.2 Delete evidence boundary',2)
para('The delete screenshot shows the browser confirmation prompt for NS 301. It evidences the confirmation step and intended delete action. The supplied capture set does not show a post-confirmation list or query proving absence, so this report does not claim the deletion result is independently verified by that screenshot.')
head('7.3 Oracle behavior',2)
para('The live implementation uses Oracle. The service binds values and commits successful writes. Numeric key generation depends on deployed identity/default/trigger behavior; the README flags this as a setup check. Confirm key generation and actual constraints in Oracle metadata before extending the system.')

head('8 FRONT END APPLICATION DESIGN',1)
para('The browser interface in index.html uses a persistent sidebar and content area. Its theme combines dark navigation, a light work surface, summary cards, record tables and modal forms. The supplied screens establish the Northstar Airways / Flight Operations identity for the demonstration.')
head('8.1 Screens and controls',2)
table(['Area','Visible purpose / controls'],[('Authentication','Sign-in and account creation before operational data.'),('Overview','Flight count, active fleet, open disruptions, CO2, departure visualization and upcoming flights.'),('Flight schedule','Route/time/aircraft/status; add, detail, edit, delete, search, filter, CSV export.'),('Fleet','Aircraft, capacity, fuel efficiency, today’s flight and status.'),('Crew roster','Name, role, base, assignment, hours and duty status.'),('Cargo tracking','Shipment, category, flight, weight, temperature and tracking status.'),('Passengers & loyalty','Passenger, tier, baggage and allowance.'),('Disruptions','Flight, route, reason, delay, event timestamp and status.'),('Emissions','Flight fuel/CO2 logs, summary metrics and timestamps.')],[1.65,4.95],7.8)
head('8.2 Usability and data presentation',2)
para('Search and status filters narrow lists; CSV export supports offline review; refresh reloads current records. Overview cards provide a summary, while module tables show record-level details. The departure chart is labelled as sample data and is illustrative unless the underlying aggregation is verified. Counts are screenshots of a changing demo dataset, not a live airline network report.')

head('9 OVERALL SYSTEM DESIGN AND CONNECTIVITY',1)
image(ASSET/'architecture.png');caption('Figure 2. Browser, FastAPI service and Oracle XE architecture.')
head('9.1 Front-end and back-end connectivity procedure',2)
table(['Step','Procedure'],[('1. Configure','Keep Oracle credentials, service DSN and session secret in private environment settings.'),('2. Start service','Launch FastAPI at 127.0.0.1:8001; serve index.html from the same origin.'),('3. Authenticate','Register/sign in; server verifies the hash and issues a signed HttpOnly session cookie.'),('4. Read','Browser requests API data; service uses a pooled python-oracledb connection.'),('5. Write','Create/edit/delete calls map to allowed tables, bind values and commit successful writes.'),('6. Refresh','Browser reloads current Oracle records; README indicates refresh every five seconds.')],[1.25,5.35],7.8)
head('9.2 API and data flow',2)
para('The service provides health and authentication endpoints, consolidated data, reference lookups and entity-level create/update/delete routes. It maps presentation fields to Oracle columns and derives some view values from related tables. Health reports database connectivity and table visibility. Operational reads and writes require an authenticated user.')
image(ASSET/'workflow.png');caption('Figure 3. Typical authenticated data-maintenance workflow.')
head('9.3 Authentication and security',2)
para('The README and source describe scrypt password hashing, signed HttpOnly session cookies with a 12-hour expiry, a private session secret, and APP_LOGIN_USERS rather than the existing subscription-oriented USERS table. All accounts currently share the same permissions; role-based access is not implemented. The app is intended for local coursework and should not be exposed to an untrusted network without production security work.')

head('10 PRACTICAL IMPLEMENTATION',1);head('10.1 Implementation stack',2)
table(['Layer','Technology / role'],[('Browser','HTML, CSS and JavaScript for screens and interaction.'),('Application service','Python FastAPI for authentication, API routing, mapping and validation.'),('Database adapter','python-oracledb Thin mode and connection pooling.'),('Database','Oracle XE tables, relationships and constraints.'),('Transport','Local HTTP in coursework; JSON request/response bodies.')],[1.5,5.1],7.8)
head('10.2 CRUD implementation',2)
para('The API centralizes entity mapping so write logic can build INSERT or UPDATE statements from permitted fields and bind values. Delete routes target record keys and commit the operation. The interface presents forms and detail modals and reloads records after changes. Foreign-key violations, duplicates, invalid types and missing key generation remain database errors that should be clearly communicated.')
head('10.3 Demonstration workflow',2)
para('The screenshots cover sign-in and account creation; overview; flight list and detail; creation of NS 309; the Oracle row count change from eight to nine; an edit changing status to DELAYED; Oracle output reflecting the update; delete confirmation; and the fleet, crew, cargo, passenger, disruptions and emissions modules.')
head('10.4 Current implementation limits',2)
para('The service is local-only, has no separate user roles and uses illustrative sample views in parts of the overview. The evidence does not establish equivalent write testing for every module. FLIGHT_CREW is in the schema, while the visible menu has a crew roster and no dedicated assignment-management screen. These are practical next targets.')

head('11 BUSINESS ASPECT AND PRACTICAL VALUE',1);head('11.1 Stakeholders and value proposition',2)
para('Potential users include operations coordinators, fleet planners, cargo teams and sustainability analysts. A consolidated operational view can reduce time spent switching between lists, make update ownership clearer and make sustainability fields available alongside flights. The demonstrator illustrates this concept; operational savings have not been measured.')
head('11.2 Business model options',2)
table(['Model','Application','Considerations'],[('Internal deployment','Operate as an internal support workspace.','Needs integration, access control, audit, uptime and governance.'),('SaaS subscription','Tier by airline size, users or modules.','Needs tenant isolation, cloud hosting, support and compliance.'),('Modular add-ons','Cargo temperature, disruption analytics or emissions reports.','Needs verified data definitions and reliable integrations.'),('Implementation services','Configuration, migration and staff training.','Requires domain validation and maintenance.')],[1.45,2.5,2.65],7.7)
head('11.3 Costs, value and adoption',2)
para('A commercial estimate needs discovery rather than speculative pricing. Cost drivers include hosting, database licensing, identity management, monitoring, backups, security review, integration, support and training. Potential value drivers include less manual reconciliation, faster exception visibility, better cargo traceability and easier fuel/emissions reporting. A pilot should measure task time, data error rate, disruption response time and emissions-record completeness against a baseline.')
para('A practical adoption path is to validate the model with an operations stakeholder, pilot a limited route or team, connect approved source systems, reconcile in parallel, train users and expand only after data-quality and security gates pass. This student build is a proof of concept, not production-ready software.')

head('12 SUSTAINABILITY IMPACT AND INSPIRATION',1)
para('The project was inspired by the need to make sustainability measurable in operational workflows. Fuel and CO2 logging makes environmental impact visible at flight level, and fuel-efficiency fields create a basis for comparison. Cargo temperature fields support careful movement of medical and perishable goods and can help reduce spoilage when connected to reliable measurements and response processes.')
para('Environmental benefit depends on accurate data and decisions. The application stores and presents records; it does not independently measure emissions, calculate lifecycle impact, optimize routes or prove carbon reduction. Future work should document emission factors, provenance, reporting boundaries and verification before external sustainability claims.')
head('12.1 Potential sustainability KPIs',2)
table(['KPI','Example definition','Data needed'],[('Fuel per flight','Fuel liters per flight, normalized by distance','Fuel log; route distance; aircraft type'),('CO2 intensity','CO2 per passenger-km or tonne-km','Verified emissions; loads; distance'),('Cold-chain excursion rate','Shipments outside allowed range / monitored shipments','Timestamped sensor readings and thresholds'),('Data completeness','Flights with required logs / operated flights','Schedule and emission logs')],[1.55,2.65,2.4],7.8)

head('13 SCREENSHOTS OF THE APPLICATION WITH DATA',1)
para('Figures below reproduce the supplied application evidence. Counts, dates and values are shown as captured; dashboard metrics and sample records are illustrative, not live airline reports.')
para('Demo video (Google Drive): link not provided. Insert the shareable demonstration URL here before submission.',boldlead='Demo video (Google Drive):')
shots=[
('Screenshot 2026-10-13 200937.png','Sign-in screen for the Flight Operations workspace.'),('Screenshot 2026-10-13 213801.png','Local account registration screen.'),('Screenshot 2026-10-13 201226.png','Overview cards and sample departure flow.'),('Screenshot 2026-10-13 213839.png','Flight schedule with eight records before create.'),('Screenshot 2026-10-13 214639.png','Oracle SELECT output showing initial flight rows.'),('Screenshot 2026-10-13 214957.png','Add Flight form populated for flight 309.'),('Screenshot 2026-10-13 215010.png','Oracle output after insert; nine flight rows.'),('Screenshot 2026-10-13 215259.png','Detail view for NS 309 after creation.'),('Screenshot 2026-10-13 215327.png','Edit form with schedule change and DELAYED status.'),('Screenshot 2026-10-13 215355.png','Oracle output reflecting updated flight 309.'),('Screenshot 2026-10-13 220246.png','Delete confirmation prompt for NS 301; prompt only.'),('Screenshot 2026-10-13 220350.png','Fleet module with aircraft and service details.'),('Screenshot 2026-10-13 220356.png','Crew roster with roles, bases and assignments.'),('Screenshot 2026-10-13 220402.png','Cargo tracking with category and temperature fields.'),('Screenshot 2026-10-13 220409.png','Passenger and loyalty module.'),('Screenshot 2026-10-13 220417.png','Disruptions module with causes, delay and status.'),('Screenshot 2026-10-13 220424.png','Emissions module with fuel and CO2 records.'),('Screenshot 2026-10-13 220449.png','Overview after additional flight is present.')]
imgdir=Path(r'C:\Users\patil\Pictures\Screenshots')
for i,(fn,desc) in enumerate(shots):
    head(f'Figure {i+4}. {desc}',2);fp=imgdir/fn
    if fp.exists():
        with Image.open(fp) as im:
            w,h=im.size;factor=min(6.35*120/w,7.0*120/h);image(fp,min(6.35,w*factor/120))
    else:para('[Screenshot unavailable at report-generation time.]')
    caption(f'Figure {i+4}. {desc}')
    if i<len(shots)-1:page()

head('14 CONCLUSION AND FUTURE SCOPE',1);head('14.1 Conclusion',2)
para('The project has progressed from a relational database design in Review 0 to a working local airline-operations application for Review 1. The implementation brings key records into a browser workspace and connects supported CRUD workflows to Oracle XE through FastAPI. Screenshots document authentication, flight schedule operations, database persistence for a create and update example, and fleet, crew, cargo, passenger, disruption and emissions views.')
para('The design provides a foundation for structured operational data and sustainability visibility. It demonstrates database relationships, constraints, SQL concepts and an application layer. It remains a coursework proof of concept: data is illustrative, the service is local-only, permissions are uniform and sustainability outcomes have not been measured.')
head('14.2 Future enhancements',2)
table(['Priority','Enhancement','Expected benefit'],[('1','Role access, session hardening, CSRF protection, TLS and audit logs','Safer deployment; accountability.'),('2','Validation, allowlists, pagination, transactions and clear errors','Reliable maintenance at higher record volumes.'),('3','Crew assignment UI and duty/rest rules','Complete staffing workflow; fewer invalid assignments.'),('4','Sensor ingestion and cold-chain threshold alerts','Timely response and shipment traceability.'),('5','Verified emissions factors and normalized reporting','Comparable, auditable sustainability analytics.'),('6','Disruption analytics, alerts and recovery suggestions','Faster coordination and learning from recurring causes.'),('7','Automated checks, CI, backups, migrations and monitoring','Safer releases and recoverability.'),('8','Booking, airport and maintenance-system integration','Less duplicate entry; fuller network status.')],[.65,3.2,2.75],7.5)
head('14.3 Measurable next-step plan',2)
para('First validate deployed columns and constraints against Oracle metadata. Add API and browser checks for create, update and delete, including unauthorized access and database constraint failures. Add roles and audit records before network deployment. Pilot emissions and cargo measurements with timestamped, attributed sources, then evaluate completeness, exception response time and normalized fuel/emissions indicators.')

head('REFERENCES AND PROJECT MATERIALS',1)
for s in ['Review 0 report: Modern Airline Logistics and Sustainability Management System, BACSE202 Lab Component Project, AY 2026–27.','Review 1 guideline: Database Systems QP 9 (2).pdf, School of Computer Science and Engineering (SCOPE), AY 2026–27.','Project source files: server.py, index.html, requirements.txt and README.txt.','Implementation evidence: application and Oracle screenshots supplied by the author on 13 October 2026.']:para(s)

# Full source appendices, excluding credentials and local secrets.
for title,filename in [('APPENDIX A FULL IMPLEMENTATION CODE SERVER','server.py'),('APPENDIX B FULL IMPLEMENTATION CODE FRONT END','index.html')]:
    page();head(title,1);para(f'Listing reproduced from {filename}. Private .env and session-secret files are intentionally excluded.',italic=True)
    for n,line in enumerate((ROOT/filename).read_text(encoding='utf-8').splitlines(),1):
        p=doc.add_paragraph(style='Code Listing');r=p.add_run(f'{n:04d}  {line}');r.font.name='Consolas';r.font.size=Pt(6.7)

props=doc.core_properties;props.title='Review 1 Project Implementation Report';props.subject='Modern Airline Logistics and Sustainability Management System';props.author='Pranali Paygonda Patil';props.keywords='BACSE202, Oracle XE, airline logistics, sustainability'
upd=OxmlElement('w:updateFields');upd.set(qn('w:val'),'true');doc.settings.element.append(upd)
doc.save(OUT)
print(OUT,'paragraphs',len(doc.paragraphs),'tables',len(doc.tables),'figures',len(shots))
