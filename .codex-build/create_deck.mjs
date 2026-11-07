import fs from 'node:fs/promises';
import path from 'node:path';
import { Presentation, PresentationFile } from '@oai/artifact-tool';

const root = 'C:/Users/patil/OneDrive/Desktop/College DA/Sem 3/DBMS/Project';
const build = path.join(root, '.codex-build');
const outDir = path.join(root, 'deliverables');
const draftDir = path.join(build, 'draft');
await fs.mkdir(draftDir, { recursive: true });
await fs.mkdir(outDir, { recursive: true });

const W=1280,H=720;
const C={navy:'#0B1421',navy2:'#142438',ink:'#142437',muted:'#607181',teal:'#17B394',mint:'#DDF5EF',pale:'#F4F7F8',white:'#FFFFFF',line:'#DDE6E8',amber:'#D99438',rose:'#D36C62',blue:'#5875BA',soft:'#90A0AB'};
const FONT='Aptos';
const ppt=Presentation.create({slideSize:{width:W,height:H}});

function shape(slide,geometry,x,y,w,h,fill='none',lineFill='none',lineWidth=0,name=''){
  return slide.shapes.add({geometry,name,position:{left:x,top:y,width:w,height:h},fill,line:{style:'solid',fill:lineFill,width:lineWidth}});
}
function txt(slide,text,x,y,w,h,size=24,color=C.ink,bold=false,opts={}){
  const s=shape(slide,'textbox',x,y,w,h,'none','none',0,opts.name||''); s.text=text;
  s.text.style={fontFamily:FONT,fontSize:size,color,bold,alignment:opts.align||'left',verticalAlignment:opts.valign||'middle',wrap:true};
  return s;
}
function rect(slide,x,y,w,h,fill,r=14,stroke='none',sw=0){return shape(slide,r?'roundRect':'rect',x,y,w,h,fill,stroke,sw);}
function line(slide,x1,y1,x2,y2,color=C.line,weight=2){shape(slide,'line',x1,y1,x2-x1||0.01,y2-y1||0.01,'none',color,weight);}
function base(title,n,opts={}){
  const s=ppt.slides.add(); s.background.fill=opts.dark?C.navy:C.pale;
  if(!opts.cover){
    txt(s,opts.kicker||'MODERN AIRLINE LOGISTICS',72,34,700,22,14,opts.dark?'#8BA3B4':C.teal,true);
    txt(s,title,72,66,1120,66,44,opts.dark?C.white:C.ink,true);
    if(opts.sub)txt(s,opts.sub,72,130,1120,36,21,opts.dark?'#A9BAC5':C.muted,false);
    line(s,72,680,1208,680,opts.dark?'#2B3A49':C.line,1);
    txt(s,String(n).padStart(2,'0'),1152,686,56,20,13,opts.dark?'#7890A0':C.soft,true,{align:'right'});
  }
  return s;
}
function bulletList(s,items,x,y,w,{size=24,gap=74,color=C.ink,accent=C.teal}={}){
  items.forEach((item,i)=>{const yy=y+i*gap; rect(s,x,yy+8,9,9,accent,1); txt(s,item,x+26,yy,w-26,gap-10,size,color,false);});
}
function note(s,text){s.speakerNotes.text=text;}
function chip(s,label,x,y,w,fill=C.mint,color=C.teal){rect(s,x,y,w,32,fill,16);txt(s,label,x+9,y,w-18,32,14,color,true,{align:'center'});}
function label(s,title,body,x,y,w,{accent=C.teal,dark=false}={}){rect(s,x,y,4,74,accent,2);txt(s,title,x+18,y-2,w-18,28,21,dark?C.white:C.ink,true);txt(s,body,x+18,y+29,w-18,50,17,dark?'#B6C4CD':C.muted,false);}
function notes(s,files='Review 0.docx; 25BCE2957.pdf; README.txt; server.py; index.html'){note(s,`Sources: ${files}. Statements about future capabilities are proposals unless labeled implemented.`);}

// 01 cover
{
 const s=ppt.slides.add();s.background.fill=C.navy;
 rect(s,850,0,430,720,C.navy2,0);rect(s,850,0,8,720,C.teal,0);
 txt(s,'VIT SCHOOL OF COMPUTER SCIENCE AND ENGINEERING',72,54,660,30,16,'#94B4BD',true);
 txt(s,'MODERN AIRLINE\nLOGISTICS &\nSUSTAINABILITY\nMANAGEMENT SYSTEM',72,147,730,245,46,C.white,true);
 txt(s,'Database project for BACSE202 Lab Component',74,415,680,38,22,'#C0CED6',false);
 line(s,74,484,750,484,'#344759',1);
 txt(s,'Pranali Paygonda Patil',74,508,600,32,24,C.white,true);
 txt(s,'25BCE2957, B.Tech Computer Science and Engineering',74,546,690,26,17,'#B2C3CC',false);
 txt(s,'School of Computer Science and Engineering, AY 2026–27',74,580,720,25,16,'#B2C3CC',false);
 txt(s,'REVIEW 0 PROJECT PRESENTATION',890,602,310,24,14,C.teal,true);
 txt(s,'AIRLINE OPERATIONS\nCONNECTED THROUGH DATA',890,266,320,92,28,C.white,true);
 txt(s,'Flight, fleet, crew, passenger, cargo and emissions',890,378,320,68,18,'#B2C3CC',false);
 notes(s);
}
// 02 abstract
{
 const s=base('Abstract',2,{sub:'A centralized relational system for connected airline operations'});
 txt(s,'The project brings essential airline logistics records into one structured database. It links aircraft and flights with crew assignments, passenger loyalty, specialized cargo, emissions, and disruption events.',72,194,650,218,28,C.ink,false);
 line(s,760,190,760,560,C.line,2);
 label(s,'Operational control','Search and maintain linked records through a browser console.',800,198,390);
 label(s,'Sustainability visibility','Capture fuel and CO₂ at flight level for later analysis.',800,316,390,{accent:C.blue});
 label(s,'Decision support','Use joins, aggregates, grouping, and subqueries to answer operational questions.',800,434,390,{accent:C.amber});
 chip(s,'9 relational entities',72,548,220);chip(s,'Oracle + FastAPI web app',310,548,280,C.navy2,C.white);chip(s,'Cold-chain + emissions focus',608,548,286);
 notes(s);
}
// 03 problem
{
 const s=base('The operating problem',3,{sub:'Airline logistics crosses teams, systems, and time-sensitive decisions'});
 const items=[['Schedules','Flight and aircraft status can drift across separate records.'],['People','Crew assignments must stay connected to flights and duty details.'],['Cargo','Temperature-sensitive goods need handling requirements and checkpoints.'],['Disruptions','Causes and delay duration need a consistent event history.']];
 items.forEach((v,i)=>{let x=72+i*286;txt(s,String(i+1),x,206,50,48,34,C.teal,true);txt(s,v[0],x,276,240,34,24,C.ink,true);txt(s,v[1],x,322,238,110,18,C.muted,false);if(i<3)line(s,x+252,205,x+252,460,C.line,1);});
 rect(s,72,496,1136,102,C.navy,14);txt(s,'Project response',98,514,210,30,19,'#83DCC9',true);txt(s,'A shared relational model links operational events to the flight that anchors them.',98,548,1050,36,25,C.white,true);
 notes(s);
}
// 04 inspiration
{
 const s=base('Why sustainability belongs in logistics',4,{sub:'The project was inspired by the need to make environmental impact visible alongside operations'});
 txt(s,'Every flight already creates operational records. Adding fuel and CO₂ fields makes sustainability measurable at the same unit of work: the flight.',72,205,490,190,28,C.ink,false);
 const steps=[['Record','Fuel liters + CO₂ tons'],['Relate','Tie each log to a flight'],['Analyse','Compare routes and aircraft']];
 steps.forEach((v,i)=>{let x=650+i*178;rect(s,x,240,150,150,i===1?C.teal:C.navy2,75);txt(s,String(i+1),x,250,150,46,28,C.white,true,{align:'center'});txt(s,v[0],x-10,420,170,30,22,C.ink,true,{align:'center'});txt(s,v[1],x-20,457,190,56,16,C.muted,false,{align:'center'});if(i<2)line(s,x+150,315,x+178,315,C.teal,3);});
 rect(s,72,548,490,50,C.mint,12);txt(s,'Inspiration: sustainability through measurement and operational awareness',88,548,460,50,16,'#167C69',true);
 notes(s);
}
// 05 goals
{
 const s=base('Project objectives',5,{sub:'What the database and interface are designed to support'});
 const goals=[['Integrate','Relate operational records through keys and constraints.'],['Protect quality','Reduce inconsistent references and invalid values.'],['Monitor','Keep cold-chain readings, emissions, and disruptions queryable.'],['Inform','Use SQL summaries and joins to support operational decisions.']];
 goals.forEach((g,i)=>{let y=205+i*91;txt(s,String(i+1).padStart(2,'0'),74,y,58,44,30,C.teal,true);txt(s,g[0],154,y,232,36,25,C.ink,true);txt(s,g[1],396,y,730,40,21,C.muted,false);line(s,154,y+59,1185,y+59,C.line,1);});notes(s);
}
// 06 stakeholders
{
 const s=base('Users and business stakeholders',6,{sub:'A coursework prototype for roles that depend on a reliable shared view'});
 const rows=[['Operations controller','Review departures, statuses, disruptions, and linked resources.'],['Cargo coordinator','Check category, shipment weight, required temperature, and current reading.'],['Sustainability analyst','Summarize fuel and CO₂ logs by flight or route.'],['Database administrator','Maintain schema, accounts, constraints, and database availability.']];
 rows.forEach((r,i)=>{let y=196+i*90;rect(s,72,y,1136,72,i%2?C.white:'#EAF1F2',10);txt(s,r[0],94,y+13,290,28,21,C.ink,true);txt(s,r[1],395,y+12,790,48,18,C.muted,false);});
 txt(s,'Current prototype: all registered users have the same access. Separate job roles and permissions remain future work.',72,584,1125,40,18,C.rose,true);notes(s);
}
// 07 scope
{
 const s=base('Functional scope',7,{sub:'The report defines nine entities; the current browser interface exposes operational views'});
 const col1=['Aircraft and flight scheduling','Crew and flight assignments','Passenger loyalty and baggage'];
 const col2=['Specialized cargo and temperature checks','Fuel and CO₂ logs','Disruption events and response status'];
 txt(s,'DATABASE DOMAIN',72,196,420,24,15,C.teal,true);bulletList(s,col1,72,238,500,{size:22,gap:85});
 txt(s,'OPERATIONAL USE',650,196,420,24,15,C.blue,true);bulletList(s,col2,650,238,530,{size:22,gap:85,accent:C.blue});
 rect(s,72,514,1136,74,C.navy,12);txt(s,'The FLIGHT_CREW relationship exists in the schema; the current UI does not provide a dedicated assignment screen.',94,527,1088,48,19,C.white,true);notes(s);
}
// 08 architecture
{
 const s=base('Solution architecture',8,{dark:true,sub:'Browser requests flow through a local API to Oracle XE'});
 const xs=[72,367,662,957],labels=[['BROWSER UI','HTML · CSS · JavaScript'],['FASTAPI','Routes · validation · sessions'],['ORACLE DRIVER','python-oracledb · Thin mode'],['ORACLE XE','Relational tables · constraints']];
 xs.forEach((x,i)=>{rect(s,x,250,240,158,i===3?'#174538':'#17293A',12,i===3?'#218A70':'#314657',1);txt(s,labels[i][0],x+18,272,204,32,19,i===3?'#74E0C8':C.white,true,{align:'center'});txt(s,labels[i][1],x+22,319,196,58,17,'#C1CDD5',false,{align:'center'});if(i<3){line(s,x+240,328,x+290,328,C.teal,3);txt(s,'JSON',x+245,294,48,20,12,'#8CA5B3',true,{align:'center'});}});
 txt(s,'Oracle changes appear in the UI after refresh; the page polls every five seconds.',72,478,1120,42,22,C.white,true);
 chip(s,'localhost:8001',72,552,190,'#1E3242','#AFC0CA');chip(s,'Signed-in routes',278,552,205,'#1E3242','#AFC0CA');chip(s,'Oracle commit on writes',499,552,260,'#1E3242','#AFC0CA');notes(s,'README.txt; server.py; index.html');
}
// 09 stack & implementation
{
 const s=base('Technology and practical implementation',9,{sub:'The runnable codebase reflects the current implementation'});
 const rows=[['Frontend','Single-page HTML/CSS/JavaScript console served by FastAPI'],['API','Python FastAPI with JSON endpoints and CRUD routes'],['Database','Oracle XE via python-oracledb Thin mode'],['Configuration','Local .env variables for Oracle user, password, and DSN'],['Run path','Start Uvicorn, then open http://127.0.0.1:8001']];
 rows.forEach((r,i)=>{let y=188+i*74;txt(s,r[0].toUpperCase(),72,y,188,26,14,C.teal,true);txt(s,r[1],278,y-4,890,38,20,C.ink,i===0);line(s,72,y+49,1206,y+49,C.line,1);});
 rect(s,72,577,1136,45,'#FBF0DE',9);txt(s,'Review 0 lists MySQL Workbench/MySQL 8.0; the current app code and README use Oracle XE.',88,578,1098,42,16,'#95631E',true);notes(s,'README.txt; requirements.txt; server.py; index.html; Review 0.docx');
}
// 10 ER
{
 const s=base('Entity relationship model',10,{sub:'Nine entities organize the airline domain around FLIGHT'});
 const nodes=[['AIRCRAFT',80,210],['FLIGHT',430,210],['CREW',80,410],['FLIGHT_CREW',430,410],['SPEC_CARGO',795,210],['EMISSIONS',795,310],['DISRUPTION',795,410],['LOYALTY_TIER',80,310],['PASSENGER',430,310]];
 nodes.forEach(([name,x,y])=>{rect(s,x,y,190,54,name==='FLIGHT'?C.teal:C.white,10,name==='FLIGHT'?C.teal:C.line,1);txt(s,name,x+8,y,174,54,17,name==='FLIGHT'?C.white:C.ink,true,{align:'center'});});
 [[270,237,430,237],[270,437,430,437],[270,337,430,337],[620,237,742,237],[742,237,742,437],[742,237,795,237],[742,337,795,337],[742,437,795,437]].forEach(a=>line(s,...a,C.teal,2));
 chip(s,'PK = primary key',837,510,205,C.navy2,C.white);chip(s,'FK = foreign key',1059,510,149,C.mint,C.teal);
 txt(s,'FLIGHT_CREW resolves the many-to-many flight and crew relationship.',72,590,850,32,18,C.muted,false);notes(s);
}
// 11 relational schema
{
 const s=base('Relational schema at a glance',11,{sub:'Primary keys identify records; foreign keys preserve cross-table links'});
 const groups=[['CORE','AIRCRAFT (Aircraft_ID PK)','FLIGHT (Flight_ID PK, Aircraft_ID FK)'],['PEOPLE','CREW (Crew_ID PK)','FLIGHT_CREW (Assignment_ID PK, Flight_ID FK, Crew_ID FK)','LOYALTY_TIER (Tier_ID PK)','PASSENGER (Passenger_ID PK, Tier_ID FK)'],['OPERATIONS','SPEC_CARGO (Cargo_ID PK, Flight_ID FK)','EMISSIONS (Log_ID PK, Flight_ID FK)','DISRUPTION (Event_ID PK, Flight_ID FK)']];
 groups.forEach((g,i)=>{let x=72+i*380;txt(s,g[0],x,191,340,24,14,[C.teal,C.blue,C.amber][i],true);line(s,x,222,x+340,222,[C.teal,C.blue,C.amber][i],3);g.slice(1).forEach((v,j)=>txt(s,v,x,242+j*54,340,42,17,C.ink,j===0));});
 rect(s,72,504,1136,92,C.navy,12);txt(s,'Foreign-key paths',96,517,210,25,18,'#77DEC7',true);txt(s,'FLIGHT references AIRCRAFT. FLIGHT_CREW links flights with crew. PASSENGER references LOYALTY_TIER.\nSPEC_CARGO, EMISSIONS and DISRUPTION records reference FLIGHT.',96,544,1080,45,15,C.white,false);notes(s);
}
// 12 relationships
{
 const s=base('Relationships and cardinality',12,{sub:'The model links each operational record to its parent entity'});
 const data=[['AIRCRAFT','1 : many','FLIGHT'],['FLIGHT','many : many via FLIGHT_CREW','CREW'],['FLIGHT','1 : many','SPEC_CARGO / EMISSIONS / DISRUPTION'],['LOYALTY_TIER','1 : many','PASSENGER']];
 data.forEach((r,i)=>{let y=192+i*88;txt(s,r[0],82,y,315,35,21,C.ink,true);line(s,420,y+18,540,y+18,C.teal,3);txt(s,r[1],550,y,280,35,18,C.teal,true,{align:'center'});line(s,850,y+18,945,y+18,C.teal,3);txt(s,r[2],960,y,235,50,18,C.ink,true);});
 rect(s,72,570,1136,52,C.mint,10);txt(s,'The associative table stores one row per crew-to-flight assignment, including shift hours.',90,572,1100,48,18,'#167C69',true);notes(s);
}
// 13 integrity
{
 const s=base('Data integrity and normalization',13,{sub:'Constraints make the database enforce rules at the point of change'});
 const items=[['PRIMARY KEY','Unique identity for each entity row'],['FOREIGN KEY','Valid links between related records'],['UNIQUE','No duplicate tier names or passport numbers'],['NOT NULL / CHECK','Required fields and allowed-value rules'],['DEFAULT','A defined value when an insert omits one']];
 items.forEach((r,i)=>{let y=187+i*73;chip(s,r[0],72,y,212,i===0?C.navy2:C.mint,i===0?C.white:C.teal);txt(s,r[1],320,y-1,820,38,20,C.ink,false);line(s,72,y+51,1208,y+51,C.line,1);});
 txt(s,'Logical separation reduces repeated facts while referential integrity keeps records connected.',72,576,1100,32,19,C.muted,true);notes(s);
}
// 14 workflow
{
 const s=base('Flight operations workflow',14,{sub:'A flight record anchors connected planning and monitoring data'});
 const stages=[['01','Select aircraft','FLIGHT.AIRCRAFT_ID'],['02','Schedule route','Route, departure and arrival'],['03','Assign crew','FLIGHT_CREW links people'],['04','Load cargo','Shipment, category and temperature'],['05','Track outcomes','Disruption and emissions logs']];
 stages.forEach((a,i)=>{let x=68+i*238;txt(s,a[0],x,226,90,35,24,C.teal,true);if(i<4)line(s,x+57,243,x+221,243,C.line,3);txt(s,a[1],x,294,210,50,21,C.ink,true);txt(s,a[2],x,356,210,86,16,C.muted,false);});
 rect(s,72,510,1136,78,C.white,12,C.line,1);txt(s,'Operational rule',92,523,180,28,17,C.teal,true);txt(s,'A flight must reference a real aircraft; child events must reference a valid flight.',282,519,890,48,20,C.ink,true);notes(s);
}
// 15 cold chain
{
 const s=base('Cold-chain cargo management',15,{sub:'Temperature-sensitive shipments carry their own monitoring fields'});
 const facts=[['CATEGORY','Medical, perishable, hazardous and standard'],['REQUIRED','Shipment target temperature (°C)'],['CURRENT','Latest recorded temperature (°C)'],['TRACKING','Status label for shipment follow-up']];
 facts.forEach((r,i)=>{let x=72+(i%2)*570,y=200+Math.floor(i/2)*146;rect(s,x,y,536,116,C.white,12,C.line,1);txt(s,r[0],x+20,y+17,156,22,14,C.teal,true);txt(s,r[1],x+20,y+50,495,50,20,C.ink,true);});
 rect(s,72,522,1136,72,'#FBF0DE',11);txt(s,'Current limitation',92,532,190,24,16,'#95631E',true);txt(s,'The prototype stores readings and status; automated sensor feeds and alert thresholds are future work.',282,528,894,47,18,'#715727',false);notes(s);
}
// 16 passengers
{
 const s=base('Passenger loyalty and baggage',16,{sub:'A tier reference keeps allowance rules separate from passenger records'});
 const tiers=[['SUMMIT','+15 kg'],['AURORA','+5 kg'],['HORIZON','+0 kg']];
 tiers.forEach((r,i)=>{let y=216+i*98;txt(s,r[0],94,y,350,40,23,C.ink,true);line(s,452,y+20,822,y+20,C.line,2);txt(s,r[1],865,y,220,42,28,[C.teal,C.blue,C.amber][i],true);});
 rect(s,72,546,1136,62,C.navy,10);txt(s,'PASSENGER references LOYALTY_TIER',98,554,500,44,20,C.white,true);txt(s,'Tier name and extra baggage allowance remain in the reference table.',590,554,590,44,17,'#B9C7CF',false);notes(s,'Review 0.docx; index.html; server.py');
}
// 17 disruption
{
 const s=base('Disruption management',17,{sub:'Each disruption record captures the cause, delay, time, and response state'});
 const why=['Weather','Technical issue','Air traffic control','Crew availability','Security / ground handling'];
 why.forEach((v,i)=>{let y=194+i*62;rect(s,72,y+8,14,14,[C.rose,C.amber,C.blue,C.teal,C.navy2][i],7);txt(s,v,104,y,440,38,20,C.ink,true);});
 line(s,640,198,640,492,C.line,2);txt(s,'EVENT RECORD FIELDS',694,201,430,24,14,C.teal,true);bulletList(s,['Delay duration in minutes','Reason code','Event date and time','Response status'],694,244,450,{size:18,gap:57,accent:C.blue});
 txt(s,'Management can group events by reason or status to see recurring pressure points.',72,568,1100,38,20,C.teal,true);notes(s);
}
// 18 emissions
{
 const s=base('Flight-level sustainability records',18,{sub:'The EMISSIONS table stores fuel and CO₂ measurements against each flight'});
 const nums=[['FUEL_LITERS','Fuel consumed'],['CO2_TONS','Carbon dioxide recorded'],['LOGGED_AT','Time of measurement'],['FLIGHT_ID','Operational link']];
 nums.forEach((r,i)=>{let x=72+(i%2)*568,y=205+Math.floor(i/2)*126;txt(s,r[0],x,y,490,26,15,C.teal,true);txt(s,r[1],x,y+35,490,42,24,C.ink,true);line(s,x,y+89,x+500,y+89,C.line,2);});
 rect(s,72,499,1136,92,C.mint,12);txt(s,'Illustrative front-end sample logs total 84,000 L of fuel and 223.2 t of CO₂.',94,511,1082,36,20,'#167C69',true);txt(s,'Use live Oracle data for operational reporting; these bundled values are demonstrative.',94,548,1082,26,16,'#715727',false);notes(s,'server.py; index.html; Review 0.docx');
}
// 19 SQL portfolio
{
 const s=base('SQL portfolio in Review 0',19,{sub:'The coursework demonstrates both record maintenance and analytical queries'});
 const left=['DDL: create tables with constraints','DML: insert, update, and delete','SELECT and single-row functions','Aggregate functions and GROUP BY'];
 const right=['HAVING for group-level filters','Subqueries for nested conditions','JOINs across operational entities','Five initial rows per table in the report'];
 txt(s,'QUERY TOOLKIT',72,192,470,24,15,C.teal,true);bulletList(s,left,72,234,510,{size:20,gap:73});
 txt(s,'ANALYTICS PATH',654,192,470,24,15,C.blue,true);bulletList(s,right,654,234,530,{size:20,gap:73,accent:C.blue});
 rect(s,72,545,1136,56,C.navy,10);txt(s,'Examples in the report cover aircraft and flight, passenger and tier, crew and flight, cargo and flight, and disruption and flight.',92,548,1090,50,17,C.white,true);notes(s);
}
// 20 decision queries
{
 const s=base('SQL and business decisions',20,{sub:'The relational model supports practical questions without claiming automated optimization'});
 const qs=[['Fleet','Which aircraft operate which routes?','Review aircraft assignment and reported fuel metrics.'],['Cold chain','Which shipment readings differ from targets?','Prioritize manual inspection or follow-up.'],['Disruptions','Which reasons recur with longer delays?','Focus investigation on repeated causes.'],['Sustainability','How do fuel and CO₂ vary by route?','Compare operational records and identify candidates for review.']];
 qs.forEach((q,i)=>{let y=182+i*94;txt(s,q[0].toUpperCase(),72,y,180,24,14,C.teal,true);txt(s,q[1],252,y-3,540,34,20,C.ink,true);txt(s,q[2],800,y-3,380,48,16,C.muted,false);line(s,72,y+59,1208,y+59,C.line,1);});
 notes(s,'Review 0.docx. Business actions are proposed interpretations of the supported query types.');
}
// 21 console
{
 const s=base('Browser operations console',21,{sub:'A single interface brings the database entities into browsable views'});
 rect(s,72,185,780,397,C.white,14,C.line,1);rect(s,72,185,780,44,C.navy2,14);rect(s,72,215,780,14,C.navy2,0);
 txt(s,'NORTHSTAR  ·  NETWORK CONTROL',94,192,490,30,15,C.white,true);
 txt(s,'Flights today     Fleet in service     Open disruptions     CO₂ logged',98,254,700,26,14,C.muted,true);
 line(s,98,297,822,297,C.line,1);txt(s,'NS 204   DEL–DXB     Scheduled',98,312,680,30,17,C.ink,true);line(s,98,351,822,351,C.line,1);txt(s,'NS 118   BOM–SIN     Boarding',98,366,680,30,17,C.ink,true);line(s,98,405,822,405,C.line,1);txt(s,'NS 356   BLR–LHR     Monitoring',98,420,680,30,17,C.ink,true);
 txt(s,'Interface capabilities',900,203,280,30,19,C.teal,true);bulletList(s,['Entity navigation','Search and status filters','CSV export','Record detail + CRUD forms'],900,250,300,{size:18,gap:71});
 txt(s,'Sample values shown here come from the bundled front-end seed data.',72,600,1100,25,15,C.muted,false);notes(s,'index.html; server.py');
}
// 22 CRUD
{
 const s=base('API and data lifecycle',22,{sub:'The current API reads Oracle data and commits create, update, and delete actions'});
 const actions=[['READ','GET /api/data','Page reads Oracle rows'],['CREATE','POST /api/{entity}','Form values create a record'],['UPDATE','PUT /api/{entity}/{id}','Changed fields update a record'],['DELETE','DELETE /api/{entity}/{id}','Selected record is deleted']];
 actions.forEach((a,i)=>{let y=188+i*87;chip(s,a[0],72,y,150,i===0?C.navy2:C.mint,i===0?C.white:C.teal);txt(s,a[1],254,y-2,430,35,19,C.ink,true);txt(s,a[2],720,y-2,390,35,18,C.muted,false);line(s,72,y+53,1208,y+53,C.line,1);});
 rect(s,72,555,1136,55,'#FBF0DE',10);txt(s,'Oracle constraints remain authoritative; key generation must be configured for add forms where IDs are omitted.',90,557,1080,50,16,'#95631E',true);notes(s,'README.txt; server.py; index.html');
}
// 23 security
{
 const s=base('Authentication and security design',23,{sub:'The local prototype adds account sign-in around data routes'});
 const sec=[['Password storage','scrypt hash with random salt'],['Session','Signed HttpOnly cookie, 12-hour expiry'],['Database config','Credentials live in local .env'],['Route access','Operational data and CRUD require sign-in']];
 sec.forEach((r,i)=>{let x=72+(i%2)*568,y=196+Math.floor(i/2)*148;txt(s,r[0].toUpperCase(),x,y,500,25,14,C.teal,true);txt(s,r[1],x,y+36,500,60,21,C.ink,true);line(s,x,y+109,x+500,y+109,C.line,2);});
 rect(s,72,514,1136,77,C.navy,12);txt(s,'Deployment boundary',94,527,220,26,17,'#7BDCC7',true);txt(s,'The README intends localhost coursework use. Keep the API off untrusted networks. The app does not implement role-based permissions.',308,520,870,55,17,C.white,false);notes(s,'README.txt; server.py');
}
// 24 live demo sequence
{
 const s=base('Live demonstration walkthrough',24,{sub:'A short review path that shows the application end to end'});
 const demo=[['1','Start','Launch the local FastAPI service and confirm Oracle health.'],['2','Sign in','Create or use a local account.'],['3','Explore','Open flights, fleet, cargo, disruptions, emissions.'],['4','Change','Add or edit a record and observe the refreshed view.'],['5','Analyse','Show a SQL join or grouped summary in the database client.']];
 demo.forEach((d,i)=>{let y=183+i*82;rect(s,72,y,52,52,i===4?C.blue:C.teal,26);txt(s,d[0],72,y,52,52,20,C.white,true,{align:'center'});txt(s,d[1],152,y-2,210,32,22,C.ink,true);txt(s,d[2],370,y-1,800,48,18,C.muted,false);});
 notes(s,'README.txt; server.py; index.html');
}
// 25 video link
{
 const s=base('Project demo video',25,{dark:true,sub:'Add the shareable Google Drive link before submitting the presentation'});
 rect(s,122,246,1036,190,'#14293A',18,'#315166',1);
 txt(s,'GOOGLE DRIVE DEMO LINK',162,274,940,28,16,'#83DCC9',true,{align:'center'});
 txt(s,'[PASTE YOUR SHAREABLE DRIVE LINK HERE]',162,322,940,58,27,C.white,true,{align:'center'});
 txt(s,'Set Drive access to “Anyone with the link can view” if your course allows it.',162,389,940,30,17,'#B8C7D0',false,{align:'center'});
 txt(s,'Suggested clip: sign in, inspect a flight, review cargo temperature, open emissions, and demonstrate one database update.',122,496,1036,72,20,C.white,false,{align:'center'});
 notes(s,'Placeholder retained because no Google Drive link was included with the project files.');
}
// 26 deployment
{
 const s=base('Practical setup and deployment',26,{sub:'The README documents a local development workflow'});
 const steps=['Install Python 3.10+ and requirements','Copy .env.example to .env and set Oracle credentials + DSN','Start Uvicorn on 127.0.0.1:8001','Open the local URL in a browser','Register, then use the console'];
 steps.forEach((v,i)=>{let y=190+i*72;txt(s,String(i+1).padStart(2,'0'),72,y,62,36,23,C.teal,true);txt(s,v,160,y,980,40,20,C.ink,i===2);line(s,160,y+51,1208,y+51,C.line,1);});
 rect(s,72,566,1136,48,C.mint,10);txt(s,'Oracle tables and service must match the configured schema. XEPDB1 differs from the README’s CDB$ROOT setup.',88,568,1090,42,15,'#167C69',true);notes(s,'README.txt; requirements.txt');
}
// 27 validation
{
 const s=base('Verification and review evidence',27,{sub:'The report documents SQL exercises; the code documents app behavior and setup'});
 const columns=[['DATABASE REVIEW','Create tables with constraints\nInsert sample rows\nUpdate and delete demonstrations\nJoin, aggregate, group, HAVING, subquery'],['APPLICATION REVIEW','Health endpoint for Oracle connectivity\nSigned-in data routes\nCRUD requests to mapped tables\nFive-second refresh from Oracle']];
 columns.forEach((c,i)=>{let x=72+i*568;txt(s,c[0],x,197,500,24,15,i?C.blue:C.teal,true);line(s,x,230,x+510,230,i?C.blue:C.teal,3);txt(s,c[1],x,252,510,205,19,C.ink,false);});
 rect(s,72,502,1136,88,'#FBF0DE',12);txt(s,'Important evidence limit',92,514,250,28,17,'#95631E',true);txt(s,'The browser’s chart and bundled records are illustrative. Use live Oracle rows in the demo for operational results.',344,507,836,57,18,'#715727',false);notes(s,'Review 0.docx; README.txt; server.py; index.html');
}
// 28 business
{
 const s=base('Business value and operating model',28,{sub:'Potential value comes from clearer coordination and more usable operational records'});
 const values=[['Fewer handoffs','Shared flight-linked records reduce lookup across separate lists.'],['Cargo accountability','Temperature requirements and readings can be reviewed together.'],['Environmental reporting','Fuel and CO₂ logs create a basis for route-level comparison.'],['Operational learning','Disruption history can reveal recurring causes and delay patterns.']];
 values.forEach((v,i)=>{let y=188+i*88;txt(s,v[0],72,y,250,32,22,C.ink,true);txt(s,v[1],358,y-2,790,52,19,C.muted,false);line(s,72,y+58,1208,y+58,C.line,1);});
 txt(s,'Adoption path',72,560,200,28,16,C.teal,true);txt(s,'Pilot with operations and cargo teams, validate data quality, then add reporting and access controls.',278,553,900,42,18,C.ink,true);
 notes(s,'Review 0.docx; README.txt. Business value statements are intended outcomes, not measured results.');
}
// 29 future scope
{
 const s=base('Future scope',29,{sub:'A staged roadmap can grow the coursework prototype into a more operational tool'});
 const roadmap=[['NEAR TERM','Role-based access, stronger validation, explicit ID strategy'],['NEXT','Live temperature sensor feeds, threshold alerts, audit trail'],['LATER','Emissions dashboards, route benchmarks, exportable reports'],['SCALE','Cloud deployment with monitoring, backup and recovery, and integrations']];
 roadmap.forEach((r,i)=>{let y=185+i*91;chip(s,r[0],72,y,176,i===3?C.navy2:C.mint,i===3?C.white:C.teal);txt(s,r[1],280,y-1,870,47,19,C.ink,i===0);line(s,72,y+61,1208,y+61,C.line,1);});
 txt(s,'All roadmap items are proposals; they are not implemented in the current code.',72,584,1100,28,17,C.rose,true);notes(s,'README.txt; server.py; index.html');
}
// 30 close
{
 const s=ppt.slides.add();s.background.fill=C.navy;rect(s,0,0,14,720,C.teal,0);
 txt(s,'THANK YOU',86,212,1100,84,66,C.white,true,{align:'center'});
 line(s,440,326,840,326,C.teal,3);
 txt(s,'Questions and discussion',86,364,1100,42,24,'#B9C8D1',false,{align:'center'});
 txt(s,'MODERN AIRLINE LOGISTICS & SUSTAINABILITY MANAGEMENT SYSTEM',130,546,1020,30,16,'#87A0AF',true,{align:'center'});
 notes(s);
}

if(ppt.slides.items.length!==30) throw new Error(`Expected 30 slides, built ${ppt.slides.items.length}`);
for(const [i,s] of ppt.slides.items.entries()){
 const png=await ppt.export({slide:s,format:'png',scale:1});
 await fs.writeFile(path.join(draftDir,`slide-${String(i+1).padStart(2,'0')}.png`),new Uint8Array(await png.arrayBuffer()));
}
const montage=await ppt.export({format:'png',montage:true});
await fs.writeFile(path.join(draftDir,'montage.png'),new Uint8Array(await montage.arrayBuffer()));
const candidate=path.join(draftDir,'candidate.pptx');
await (await PresentationFile.exportPptx(ppt)).save(candidate);
console.log(JSON.stringify({slides:ppt.slides.items.length,candidate,montage}));
