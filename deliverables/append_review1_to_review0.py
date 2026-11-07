from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from PIL import Image, ImageDraw, ImageFont

ROOT=Path.cwd()
SRC=ROOT/'Review 0.docx'
OUT=ROOT/'deliverables'/'Review_0_with_Review_1_Additions.docx'
ASSET=ROOT/'deliverables'/'review0_extension_assets';ASSET.mkdir(exist_ok=True)

def font(sz,bold=False):
    p=Path(r'C:\Windows\Fonts\arialbd.ttf' if bold else r'C:\Windows\Fonts\arial.ttf')
    return ImageFont.truetype(str(p),sz) if p.exists() else ImageFont.load_default()

def draw_architecture():
    im=Image.new('RGB',(1500,510),'white');d=ImageDraw.Draw(im)
    for title,body,x in [('BROWSER','HTML · CSS · JavaScript\nOperations workspace',45),('FASTAPI','Python service\nAuthentication · API · mapping',560),('ORACLE XE','Relational database\nNine domain entities',1075)]:
        d.rounded_rectangle((x,105,x+365,360),radius=16,fill='#F5F7F8',outline='#52606B',width=3)
        d.text((x+20,130),title,font=font(28,True),fill='#111111')
        d.multiline_text((x+20,220),body,font=font(24),fill='#222222',spacing=12)
    for x in [425,940]:
        d.line((x,235,x+110,235),fill='#222222',width=5);d.polygon([(x+110,224),(x+130,235),(x+110,246)],fill='#222222')
    d.text((400,420),'HTTP / JSON    |    python-oracledb Thin mode    |    parameter-bound SQL',font=font(20),fill='#333333')
    im.save(ASSET/'architecture.png')

def draw_workflow():
    im=Image.new('RGB',(1500,360),'white');d=ImageDraw.Draw(im)
    labels=['Sign in','Load current\nOracle records','Select module\nand record','Create / edit /\ndelete','Commit to\nOracle','Refresh the view\nand export']
    for i,label in enumerate(labels):
        x=15+i*248;d.rounded_rectangle((x,95,x+210,270),radius=14,fill='#F5F7F8',outline='#52606B',width=3)
        d.multiline_text((x+16,150),label,font=font(22,True),fill='#111111',spacing=7)
        if i<5:
            xx=x+213;d.line((xx,180,xx+20,180),fill='#222222',width=4);d.polygon([(xx+20,172),(xx+34,180),(xx+20,188)],fill='#222222')
    im.save(ASSET/'workflow.png')
draw_architecture();draw_workflow()

doc=Document(SRC)
paras=doc.paragraphs
reference={
 'body':doc.paragraphs[80], 'center_title':doc.paragraphs[78],
 'left_title':doc.paragraphs[306], 'sub_title':doc.paragraphs[340]
}

def insert_text(text,ref='body',before=None,align=None,keep=False):
    p=doc.add_paragraph()
    source=reference[ref]
    if source._p.pPr is not None:p._p.insert(0,deepcopy(source._p.pPr))
    if align is not None:p.alignment=align
    if keep:p.paragraph_format.keep_with_next=True
    r=p.add_run(text)
    if source.runs and source.runs[0]._r.rPr is not None:r._r.insert(0,deepcopy(source.runs[0]._r.rPr))
    if before is not None:before._p.addprevious(p._p)
    return p

def add_para(text,lead=None):
    p=doc.add_paragraph()
    source=reference['body']
    if source._p.pPr is not None:p._p.insert(0,deepcopy(source._p.pPr))
    if lead and text.startswith(lead):
        r=p.add_run(lead); rr=p.add_run(text[len(lead):])
        if source.runs and source.runs[0]._r.rPr is not None:
            r._r.insert(0,deepcopy(source.runs[0]._r.rPr));rr._r.insert(0,deepcopy(source.runs[0]._r.rPr))
        r.bold=True
    else:
        r=p.add_run(text)
        if source.runs and source.runs[0]._r.rPr is not None:r._r.insert(0,deepcopy(source.runs[0]._r.rPr))
    return p

def add_heading(text,level=1):
    p=doc.add_paragraph()
    ref=reference['center_title'] if level==1 else (reference['left_title'] if level==2 else reference['sub_title'])
    if ref._p.pPr is not None:p._p.insert(0,deepcopy(ref._p.pPr))
    # These source headings carry a page-break-before property. New sections
    # already use explicit breaks, so remove it here to avoid blank pages.
    if p._p.pPr is not None:
        for node in p._p.pPr.findall(qn('w:pageBreakBefore')):
            p._p.pPr.remove(node)
    r=p.add_run(text)
    if ref.runs and ref.runs[0]._r.rPr is not None:r._r.insert(0,deepcopy(ref.runs[0]._r.rPr))
    p.paragraph_format.keep_with_next=True
    return p

def add_page_break():doc.add_page_break()

def add_picture(path,width=6.2,max_height=8.8):
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(path),width=Inches(width))
    return p

def add_caption(text):
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(text)
    if reference['body'].runs and reference['body'].runs[0]._r.rPr is not None:r._r.insert(0,deepcopy(reference['body'].runs[0]._r.rPr))
    r.italic=True
    return p

def add_table(headers,rows,source_table=3,widths=None,fontsize=9):
    t=doc.add_table(rows=1,cols=len(headers));t.style=doc.tables[source_table].style
    t.autofit=False
    # Mirror existing table header and cell typography rather than applying a new design.
    src=doc.tables[source_table]
    for j,h in enumerate(headers):
        c=t.rows[0].cells[j];c.text=h;c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        if src.rows[0].cells[min(j,len(src.columns)-1)].paragraphs[0].runs:
            old=src.rows[0].cells[min(j,len(src.columns)-1)].paragraphs[0].runs[0]
            for r in c.paragraphs[0].runs:
                if old._r.rPr is not None:r._r.insert(0,deepcopy(old._r.rPr))
                r.bold=True
    for i,row in enumerate(rows):
        cells=t.add_row().cells
        for j,value in enumerate(row):
            c=cells[j];c.text=str(value);c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            oldcell=src.rows[min(i+1,len(src.rows)-1)].cells[min(j,len(src.columns)-1)]
            oldruns=oldcell.paragraphs[0].runs
            for p in c.paragraphs:
                for r in p.runs:
                    if oldruns and oldruns[0]._r.rPr is not None:r._r.insert(0,deepcopy(oldruns[0]._r.rPr))
                    r.font.size=Pt(fontsize)
    if widths:
        for row in t.rows:
            for j,w in enumerate(widths):row.cells[j].width=Inches(w)
    doc.add_paragraph()
    return t

def fill_toc_and_figures():
    # Preserve the original contents table and adjust only its existing page references.
    toc=doc.tables[0]
    pages={1:'6',2:'7',3:'8',4:'9',5:'10',6:'12',7:'13',8:'15',9:'16',10:'19'}
    for rownum,value in pages.items():
        cell=toc.rows[rownum].cells[2]
        p=cell.paragraphs[0]
        if p.runs:
            p.runs[0].text=value
            for extra_run in list(p.runs[1:]):extra_run._r.getparent().remove(extra_run._r)
        else:p.add_run(value)
        for extra_p in list(cell.paragraphs[1:]):extra_p._p.getparent().remove(extra_p._p)
    return toc

toc=fill_toc_and_figures()

# Insert a Review 1 figure list page between the existing contents and abstract.
abstract=next(p for p in doc.paragraphs if p.text.strip()=='ABSTRACT')
abstract.paragraph_format.page_break_before=True
# Remove only empty spacer paragraphs between the contents table and abstract.
body_children=list(doc._body._element);start=body_children.index(toc._tbl);end=body_children.index(abstract._p)
for node in body_children[start+1:end]:
    if node.tag==qn('w:p') and not Paragraph(node,doc._body).text.strip():
        node.getparent().remove(node)
fighead=insert_text('LIST OF FIGURES','center_title',before=abstract,align=WD_ALIGN_PARAGRAPH.CENTER,keep=True)
fighead.paragraph_format.page_break_before=True
figures=['ER Diagram','Overall System Design and Architecture','Workflow of the Project',
'Sign-in page','Overview dashboard','Account registration page','Flight schedule',
'Oracle flight rows before insertion','Create flight form','Oracle flight rows after insertion',
'Flight details','Update flight form','Oracle updated flight record','Delete confirmation prompt',
'Fleet module','Crew roster module','Cargo tracking module','Passengers and loyalty module',
'Disruptions module','Emissions module','Updated overview dashboard']
for i,title in enumerate(figures,1):
    pp=insert_text(f'Figure {i}. {title}','body',before=abstract)
    pp.paragraph_format.left_indent=Inches(.35)

# Add a contents continuation page in the same three-column table style so the
# original merged SQL entry rows and all other template contents remain intact.
continued_title=doc.add_paragraph()
ref=reference['center_title']
if ref._p.pPr is not None:continued_title._p.insert(0,deepcopy(ref._p.pPr))
rr=continued_title.add_run('CONTENT (CONTINUED)')
if ref.runs and ref.runs[0]._r.rPr is not None:rr._r.insert(0,deepcopy(ref.runs[0]._r.rPr))
continued_title.paragraph_format.page_break_before=True
continued=doc.add_table(rows=0,cols=3);continued.style=toc.style;continued.autofit=False
continued_entries=[('S.NO','TITLE','PAGE NO'),('','LIST OF FIGURES','5'),('10.','FRONT END DESIGN','41'),('11.','BACK END TABLES & CONNECTIVITY','42'),('12.','PRACTICAL / BUSINESS / SUSTAINABILITY','44'),('13.','CODE AND APPLICATION SCREENSHOTS','46–78'),('14.','CONCLUSION / FUTURE SCOPE','79')]
for ri,entry in enumerate(continued_entries):
    row=continued.add_row()
    source_row=toc.rows[0] if ri==0 else toc.rows[1]
    for j,value in enumerate(entry):
        cell=row.cells[j];src_cell=source_row.cells[j]
        cell.width=src_cell.width
        if src_cell._tc.tcPr is not None:
            current=cell._tc.tcPr
            if current is not None:cell._tc.remove(current)
            cell._tc.insert(0,deepcopy(src_cell._tc.tcPr))
        p=cell.paragraphs[0]
        if src_cell.paragraphs[0]._p.pPr is not None:p._p.insert(0,deepcopy(src_cell.paragraphs[0]._p.pPr))
        r=p.add_run(value)
        if src_cell.paragraphs[0].runs and src_cell.paragraphs[0].runs[0]._r.rPr is not None:
            r._r.insert(0,deepcopy(src_cell.paragraphs[0].runs[0]._r.rPr))
        if ri==0:r.bold=True
        r.font.name='Times New Roman';r.font.size=Pt(12)
    for j,w in enumerate(toc.rows[1].cells):row.cells[j].width=w.width
body_children=list(doc._body._element)
fig_node=next(p._p for p in doc.paragraphs if p.text=='LIST OF FIGURES')
fig_node.addprevious(continued._tbl)
fig_node.addprevious(continued_title._p)
continued._tbl.addprevious(continued_title._p)

# Append Review 1 sections using the Review 0 direct formatting and page system.
add_page_break();add_heading('FRONT END APPLICATION DESIGN')
add_para('The front end is a browser-based airline operations workspace implemented in index.html. It uses a persistent navigation sidebar and a content area for module-specific tables and dashboards. The screens use a Northstar Airways identity, with clear status labels, search controls, refresh actions, modal forms and CSV export.')
add_para('The overview presents flight counts, fleet availability, open disruptions and CO2 totals, with a departure-flow visualization and upcoming departures. The chart is labelled as sample data and should be treated as illustrative unless the underlying aggregation is verified. The displayed values are snapshots of the supplied demonstration database.')
add_heading('Implemented modules',2)
add_table(['Module','Main information and actions'],[
('Overview','Summary cards, departure visualization and upcoming flights.'),
('Flight schedule','Flight ID, route, departure/arrival, aircraft, status; search, filters, add, details, edit, delete and CSV.'),
('Fleet','Aircraft ID/model, capacity, fuel efficiency, today’s flight and status.'),
('Crew roster','Crew member, role, base, assignment, shift hours and duty status.'),
('Cargo tracking','Shipment category, flight, weight, required/current temperature and tracking state.'),
('Passengers & loyalty','Passenger ID, loyalty tier and baggage allowance.'),
('Disruptions','Flight, route, cause, delay, event time and response status.'),
('Emissions','Flight-level fuel and CO2 records and summary metrics.')],source_table=3,widths=[1.6,4.7],fontsize=9)
add_para('Sign-in and account registration gate access to the workspace. The module screens provide a unified view over the nine database entities. A dedicated FLIGHT_CREW assignment-management page is not visible in the supplied interface; adding one is a future enhancement.')

add_page_break();add_heading('BACK END DATABASE TABLES')
add_para('This data dictionary supplements the Review 0 relational schema with logical Oracle-friendly types and integrity rules. Exact deployed column widths and constraints should be checked against Oracle metadata before migration. Types below are representative and are not asserted as exact column declarations where the source schema does not record lengths.')
schema=[
('AIRCRAFT','AIRCRAFT_ID NUMBER (PK); MODEL VARCHAR2; CAPACITY_KG NUMBER; FUEL_EFFICIENCY NUMBER; AIRCRAFT_STATUS VARCHAR2','Primary key; valid capacity/status domains.'),
('FLIGHT','FLIGHT_ID NUMBER (PK); ROUTE_CODE VARCHAR2; DEP_TIME TIMESTAMP; ARR_TIME TIMESTAMP; AIRCRAFT_ID NUMBER; FLIGHT_STATUS VARCHAR2','FK AIRCRAFT_ID → AIRCRAFT; required route/time/status.'),
('CREW','CREW_ID NUMBER (PK); CREW_NAME VARCHAR2; ROLE VARCHAR2; BASE_LOCATION VARCHAR2','Primary key; required crew identity fields.'),
('FLIGHT_CREW','ASSIGNMENT_ID NUMBER (PK); FLIGHT_ID NUMBER; CREW_ID NUMBER; SHIFT_HOURS NUMBER','FKs to FLIGHT and CREW; valid assignment hours.'),
('LOYALTY_TIER','TIER_ID NUMBER (PK); TIER_NAME VARCHAR2; EXTRA_BAGGAGE_KG NUMBER','Unique tier name; nonnegative allowance.'),
('PASSENGER','PASSENGER_ID NUMBER (PK); PASSENGER_NAME VARCHAR2; PASSPORT_NO VARCHAR2; TIER_ID NUMBER; BAGGAGE_KG NUMBER','Unique passport; FK TIER_ID → LOYALTY_TIER.'),
('SPEC_CARGO','CARGO_ID NUMBER (PK); FLIGHT_ID NUMBER; CARGO_DESCRIPTION VARCHAR2; WEIGHT_KG NUMBER; CATEGORY VARCHAR2; REQ_TEMP_C NUMBER; CURRENT_TEMP_C NUMBER; TRACKING_STATUS VARCHAR2','FK FLIGHT_ID; nonnegative weight; temperature fields.'),
('EMISSIONS','LOG_ID NUMBER (PK); FLIGHT_ID NUMBER; FUEL_LITERS NUMBER; CO2_TONS NUMBER; LOGGED_AT TIMESTAMP','FK FLIGHT_ID; nonnegative logged quantities.'),
('DISRUPTION','EVENT_ID NUMBER (PK); FLIGHT_ID NUMBER; DELAY_MIN NUMBER; REASON_CODE VARCHAR2; EVENT_DATE TIMESTAMP; EVENT_STATUS VARCHAR2','FK FLIGHT_ID; nonnegative delay.')]
add_table(['Table','Attributes and logical data types','Key and constraints'],schema,source_table=3,widths=[1.1,3.7,1.5],fontsize=8)
add_para('The application creates APP_LOGIN_USERS separately when needed. Its identity key and password hash support local authentication; it is an application-support table rather than a tenth airline-domain entity.')
add_heading('Database platform used by the application',2)
add_para('Review 0 records MySQL 8.0 in the initial software-requirement table. The delivered server.py and README use Oracle XE through python-oracledb Thin mode. This Review 1 implementation therefore documents Oracle XE as the actual runtime. SQL dialect, identity generation and constraints must match the Oracle schema used by the application.')

add_page_break();add_heading('FRONT END AND BACK END CONNECTIVITY PROCEDURE')
add_para('The application runs locally. A browser loads index.html from FastAPI; JavaScript sends JSON requests to the API; the service maps allowed presentation fields to Oracle columns and uses python-oracledb connections to retrieve or write records.')
add_picture(ASSET/'architecture.png',width=6.1);add_caption('Figure 2. Overall system design and database connectivity.')
add_heading('Connectivity sequence',2)
add_table(['Step','Procedure'],[
('1. Configure','Set Oracle username, password, service DSN and session secret in private environment settings.'),
('2. Start','Run the FastAPI application locally at 127.0.0.1:8001.'),
('3. Authenticate','Register or sign in; the service checks a password hash and issues a signed HttpOnly cookie.'),
('4. Read','The browser requests current operational data; the API queries Oracle using a connection pool.'),
('5. Write','Create/update/delete requests use mapped fields and bound values; successful writes are committed.'),
('6. Refresh','The browser reloads Oracle data; the project README describes automatic refresh every five seconds.')],source_table=3,widths=[1.1,5.2],fontsize=9)
add_picture(ASSET/'workflow.png',width=6.1);add_caption('Figure 3. Authenticated operational data workflow.')
add_heading('Authentication and security notes',2)
add_para('Passwords use scrypt hashes and signed HttpOnly cookies expire after 12 hours. All accounts share the same permissions; roles are not implemented. The app is for local coursework and needs production security before network exposure.')

add_page_break();add_heading('PRACTICAL IMPLEMENTATION, BUSINESS AND SUSTAINABILITY')
add_heading('Practical implementation',2)
add_para('The stack comprises a browser interface (HTML, CSS and JavaScript), a Python FastAPI service, python-oracledb in Thin mode and Oracle XE. The API provides health and authentication endpoints, consolidated data reads, reference lookups and entity-level create, update and delete routes. Parameter binding is used for values, and database constraints remain authoritative for relationships and valid domains.')
add_para('The supplied evidence shows a flight create flow from the Add Flight form through Oracle insertion, then a flight update reflected in a SQL query. Delete is represented by the browser confirmation prompt; the screenshot set does not show the subsequent deletion result. Key generation depends on the deployed identity/default/trigger configuration and should be confirmed against Oracle metadata.')
add_heading('Business aspect',2)
add_para('Potential users include flight operations coordinators, fleet planners, cargo teams and sustainability analysts. A shared view can reduce manual switching among separate lists and make linked flight, fleet, cargo and emissions records easier to review. These are potential benefits; the coursework demonstration has not measured time savings or operational outcomes.')
add_table(['Business path','Application','Requirements before adoption'],[
('Internal operations tool','Deploy within an airline operations team.','Integration, access control, audit, uptime and governance.'),
('Subscription service','Offer modules by airline size or user count.','Tenant isolation, cloud security, support and compliance.'),
('Modular services','Add cargo monitoring, disruption analytics or reporting.','Verified data definitions and reliable source integrations.'),
('Implementation support','Provide setup, data migration and training.','Domain validation and ongoing maintenance.')],source_table=3,widths=[1.35,2.25,2.7],fontsize=8.5)
add_para('Cost drivers include database and service hosting, security review, integration work, identity management, monitoring, backup, support and training. A pilot should measure task time, data error rates, disruption response time and the completeness of fuel and emissions records before making financial claims.')
add_heading('Sustainability inspiration and practical scope',2)
add_para('The project was inspired by making sustainability information part of day-to-day operational records. Fuel and CO2 logging can give analysts a starting point for comparing flights, routes and aircraft. Required and current cargo temperatures can support visibility for medical and perishable shipments. Better information can guide decisions, but the application does not itself reduce emissions, optimize routes or prove avoided spoilage.')
add_para('Future sustainability work should define verified emission factors, data provenance, reporting boundaries, route distance and passenger/cargo load. Potential indicators include fuel per flight-distance, CO2 per passenger-kilometre or tonne-kilometre, cold-chain excursion rate and completeness of emission logs. External sustainability claims should be based on traceable and validated data.')

add_page_break();add_heading('FULL IMPLEMENTATION CODE')
add_para('The following appendices reproduce the supplied implementation source. Private .env credentials, session secrets and local probes are excluded.')
for appendix_index,(title,filename) in enumerate([('APPENDIX A SERVER SIDE IMPLEMENTATION','server.py'),('APPENDIX B BROWSER SIDE IMPLEMENTATION','index.html')]):
    if appendix_index>0:add_page_break()
    add_heading(title,2)
    for n,line in enumerate((ROOT/filename).read_text(encoding='utf-8').splitlines(),1):
        p=doc.add_paragraph();p.paragraph_format.space_after=Pt(0);p.paragraph_format.line_spacing=1.0
        r=p.add_run(f'{n:04d}  {line}');r.font.name='Courier New';r.font.size=Pt(7.5)

add_page_break();add_heading('SCREENSHOTS OF THE APPLICATION WITH DATA')
add_para('Application and Oracle captures supplied for Review 1. Displayed counts and rows are sample/demo data at the time of capture. The deletion image shows the confirmation dialogue only.')
add_para('Demo video (Google Drive): shareable URL not supplied. Add the project demonstration link here before submission.',lead='Demo video (Google Drive):')
shots=[
('Screenshot 2026-10-13 200937.png','Sign-in page.'),
('Screenshot 2026-10-13 201226.png','Overview dashboard and sample departure flow.'),
('Screenshot 2026-10-13 213801.png','Local account registration page.'),
('Screenshot 2026-10-13 213839.png','Flight schedule table before insertion.'),
('Screenshot 2026-10-13 214639.png','Oracle flight query before insertion.'),
('Screenshot 2026-10-13 214957.png','Create flight form for NS 309.'),
('Screenshot 2026-10-13 215010.png','Oracle rows after insertion; nine rows shown.'),
('Screenshot 2026-10-13 215259.png','Flight detail view for NS 309.'),
('Screenshot 2026-10-13 215327.png','Update form for NS 309, with DELAYED status.'),
('Screenshot 2026-10-13 215355.png','Oracle record after update.'),
('Screenshot 2026-10-13 220246.png','Delete confirmation dialogue for NS 301.'),
('Screenshot 2026-10-13 220350.png','Fleet module.'),
('Screenshot 2026-10-13 220356.png','Crew roster module.'),
('Screenshot 2026-10-13 220402.png','Cargo tracking module.'),
('Screenshot 2026-10-13 220409.png','Passengers and loyalty module.'),
('Screenshot 2026-10-13 220417.png','Disruptions module.'),
('Screenshot 2026-10-13 220424.png','Emissions module.'),
('Screenshot 2026-10-13 220449.png','Overview dashboard after the additional flight.')]
imgdir=Path(r'C:\Users\patil\Pictures\Screenshots')
for i,(filename,desc) in enumerate(shots,4):
    if i>4:add_page_break()
    fp=imgdir/filename
    if fp.exists():
        with Image.open(fp) as im:
            ratio=min(6.2/im.width,8.4/im.height)
            add_picture(fp,width=im.width*ratio)
    else:add_para('[Screenshot was not available at report generation.]')
    add_caption(f'Figure {i}. {desc}')

add_page_break();add_heading('CONCLUSION AND FUTURE SCOPE')
add_para('The project has progressed from the Review 0 relational design to a local working application connected to Oracle XE. The browser workspace brings key airline records together and demonstrates authentication, flight schedule management, CRUD operations for the demonstrated flight, and module views for fleet, crew, cargo, passengers, disruptions and emissions. The data model uses relational keys and supports sustainability-related records.')
add_para('The project remains a coursework proof of concept. The supplied dataset is illustrative, the backend is intended for localhost, all registered accounts share the same permissions, and sustainability outcomes have not been quantified. The supplied delete screenshot records the confirmation prompt, not the resulting database state.')
add_heading('Future enhancements',2)
add_table(['Priority','Enhancement','Expected outcome'],[
('1','Role-based access, stronger session controls, TLS, audit logs and CSRF protection','Safer deployment and accountability.'),
('2','Validation, pagination, transaction handling and clearer database errors','More dependable workflows at larger scale.'),
('3','Crew assignment screen with duty/rest validation','Complete staffing workflow.'),
('4','Sensor ingestion and temperature threshold alerts','Faster response to cold-chain excursions.'),
('5','Verified emissions factors and normalized route/load reports','Traceable sustainability analysis.'),
('6','Disruption alerts and historical cause analysis','Improved coordination and operational learning.'),
('7','Automated API/browser checks, backups, migration and monitoring','Safer releases and recovery.'),
('8','Approved airport, booking and maintenance integrations','Fewer duplicate records and broader network view.')],source_table=3,widths=[.65,3.25,2.4],fontsize=8)
add_para('A practical next step is to validate every deployed table and key against Oracle metadata, then verify create, update, delete and unauthorized-access cases. Before any shared-network deployment, add roles and audit records. A limited pilot can then measure data completeness, exception-response time and normalized fuel/emissions indicators against an agreed baseline.')

# Keep the package's existing pagination fields and footer untouched. Ask Word to refresh fields when opened.
upd=OxmlElement('w:updateFields');upd.set(qn('w:val'),'true');doc.settings.element.append(upd)
doc.core_properties.title='Review 0 with Review 1 implementation additions'
doc.core_properties.subject='Modern Airline Logistics and Sustainability Management System'
doc.save(OUT)
print('OUTPUT',OUT)
