#!/usr/bin/env python3
"""Build PDF and ZIP submission files from GIU Germany Connect repository sources."""
from pathlib import Path
from xml.etree import ElementTree as ET
from html import unescape, escape as hescape
from zipfile import ZipFile, ZIP_DEFLATED
from textwrap import wrap
import re, math
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, black, white
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Preformatted, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from docx import Document

ROOT=Path(__file__).resolve().parent.parent
OUT=ROOT/'submission'
OUT.mkdir(exist_ok=True)
ERD=ROOT/'MASTER_EERD.drawio'
SCHEMAS=[
 'RELATIONAL_SCHEMA_01_PEOPLE_ACADEMIC.docx',
 'RELATIONAL_SCHEMA_02_UNIVERSITY_PARTNERSHIP.docx',
 'RELATIONAL_SCHEMA_03_OPPORTUNITIES_APPLICATIONS.docx',
 'RELATIONAL_SCHEMA_04_PARTICIPATION_RESEARCH.docx']
TITLES=['People and Academic','University and Partnership',
        'Opportunities and Applications','Participation and Research']

def fmt_style(s):
 return dict(p.split('=',1) for p in (s or '').split(';') if '=' in p)
def geom(c):
 g=c.find('mxGeometry')
 if g is None: return None
 return tuple(float(g.get(k,0)) for k in ('x','y','width','height'))
def label(v):
 v=unescape(v or '').replace('<br/>',' ').replace('<br>',' ').replace('</div>',' ')
 return re.sub(r'\s+',' ',re.sub(r'<[^>]*>',' ',v)).strip()

def make_erd():
 tree=ET.parse(ERD)
 nodes={x.get('id'):x for x in tree.findall('.//mxCell') if x.get('vertex')=='1'}
 edges=[x for x in tree.findall('.//mxCell') if x.get('edge')=='1']
 gs=[geom(n) for n in nodes.values() if geom(n)]
 xmin=min(g[0] for g in gs)-55
 ymin=min(g[1] for g in gs)-55
 xmax=max(g[0]+g[2] for g in gs)+55
 ymax=max(g[1]+g[3] for g in gs)+55
 w=xmax-xmin;h=ymax-ymin
 p=OUT/'GIU_Connect_ERD.pdf'
 c=canvas.Canvas(str(p), pagesize=(w,h),pageCompression=1)
 c.setTitle('GIU Germany Connect - Master EERD')
 def xy(x,y):return (x-xmin, ymax-y)
 def attach(key,fx,fy):
  q=geom(nodes[key]);return (q[0]+q[2]*fx, q[1]+q[3]*fy)
 for e in edges:
  st=fmt_style(e.get('style'))
  g=e.find('mxGeometry')
  path=[]
  sid=e.get('source');tid=e.get('target')
  if sid in nodes:path.append(attach(sid,float(st.get('exitX',.5)),float(st.get('exitY',.5))))
  elif g is not None:
   pt=g.find("mxPoint[@as='sourcePoint']")
   if pt is not None:path.append((float(pt.get('x',0)),float(pt.get('y',0))))
  if g is not None:
   pts=g.find("Array[@as='points']")
   if pts is not None:
    for pt in pts.findall('mxPoint'):path.append((float(pt.get('x',0)),float(pt.get('y',0))))
  if tid in nodes:path.append(attach(tid,float(st.get('entryX',.5)),float(st.get('entryY',.5))))
  elif g is not None:
   pt=g.find("mxPoint[@as='targetPoint']")
   if pt is not None:path.append((float(pt.get('x',0)),float(pt.get('y',0))))
  if len(path)<2:continue
  try: c.setStrokeColor(HexColor(st.get('strokeColor','#666666')))
  except: c.setStrokeColorRGB(.4,.4,.4)
  c.setLineWidth(.75)
  pth=c.beginPath()
  pth.moveTo(*xy(*path[0]))
  for pt in path[1:]:pth.lineTo(*xy(*pt))
  c.drawPath(pth,stroke=1)
 for cell in nodes.values():
  g=geom(cell)
  if not g:continue
  x,y,ww,hh=g
  if ww<=0 or hh<=0:continue
  st=fmt_style(cell.get('style'));s=cell.get('style','')
  try:c.setStrokeColor(HexColor(st.get('strokeColor','#333333')))
  except:c.setStrokeColor(black)
  try:c.setFillColor(HexColor(st.get('fillColor','#ffffff')))
  except:c.setFillColor(white)
  xx,yy=xy(x,y+hh)
  c.setLineWidth(1)
  if 'ellipse' in s:
   c.ellipse(xx,yy,xx+ww,yy+hh,fill=1,stroke=1)
   if 'double=1' in s:c.ellipse(xx+4,yy+4,xx+ww-4,yy+hh-4,fill=0,stroke=1)
  elif 'rhombus' in s:
   pth=c.beginPath();pth.moveTo(xx+ww/2,yy+hh)
   for a,b in [(xx+ww,yy+hh/2),(xx+ww/2,yy),(xx,yy+hh/2)]:pth.lineTo(a,b)
   pth.close();c.drawPath(pth,fill=1,stroke=1)
  elif s.startswith('text;'):
   pass
  else:c.rect(xx,yy,ww,hh,fill=1,stroke=1)
  val=label(cell.get('value'))
  if not val:continue
  fs=float(st.get('fontSize','10'))
  if len(val)>35 and ww<130:fs=8
  pieces=wrap(val,width=max(8,int(ww/(fs*.57))),break_long_words=True)
  if not pieces:continue
  c.setFillColor(black);c.setFont('Helvetica',fs)
  lineh=fs*1.15
  for i,line in enumerate(pieces[:7]):
   ypos=yy+hh/2+(len(pieces[:7])-1)*lineh/2-i*lineh-fs*.33
   c.drawCentredString(xx+ww/2,ypos,line)
 c.showPage();c.save()
 print('ERD PDF:',p,'nodes',len(nodes),'edges',len(edges))

def make_schema():
 p=OUT/'GIU_Connect_Relational_Schema.pdf'
 styles=getSampleStyleSheet()
 title=ParagraphStyle('TitleGIU',parent=styles['Title'],fontName='Helvetica-Bold',fontSize=16,spaceAfter=12)
 section=ParagraphStyle('SectionGIU',parent=styles['Heading2'],fontName='Helvetica-Bold',fontSize=11,spaceBefore=14,spaceAfter=7)
 normal=ParagraphStyle('SchemaText',fontName='Courier',fontSize=8,leading=11,spaceAfter=5)
 story=[Paragraph('GIU Germany Connect — Combined Relational Schema',title),
        Paragraph('PK = Primary Key &nbsp;&nbsp;&nbsp; FK = Foreign Key',styles['Normal']),Spacer(1,10)]
 for idx,(fn,head) in enumerate(zip(SCHEMAS,TITLES),1):
  story.append(Paragraph(f'{idx}. {head}',section))
  doc=Document(ROOT/fn)
  for para in doc.paragraphs:
   t=para.text.strip()
   if not t:continue
   if t.lower() in ('relational schema','opportunities & applications relational schema'):continue
   lines=wrap(t,100,break_long_words=False,break_on_hyphens=False) or ['']
   for line in lines:story.append(Preformatted(line,normal))
 SimpleDocTemplate(str(p),pagesize=A4,rightMargin=35,leftMargin=35,topMargin=34,bottomMargin=34).build(story)
 print('Schema PDF:',p)

def make_zip():
 txt=OUT/'SUBMISSION_README.txt'
 txt.write_text('''GIU Germany Connect — Team 44 submission files
GIU Notification System mandatory fields:
1. ERD (PDF): GIU_Connect_ERD.pdf
2. Relational Schema (PDF): GIU_Connect_Relational_Schema.pdf
3. ERD Shared Link: Obtain the authentic link from app.diagrams.net.
   GitHub's source-file URL is not confirmed as a substitute.
ALL THREE FIELDS ARE REQUIRED. Previously submitted forms must be resubmitted.
Master EERD in GitHub: https://github.com/kniveexe/giu-germany-connect/blob/main/MASTER_EERD.drawio
The diagram PDF is programmatically vector-rendered and may differ in connector
routing from the native diagrams.net export. Verify visually before submitting.
''',encoding='utf-8')
 dest=OUT/'TEAM_44_GIU_Germany_Connect_Submission.zip'
 files=[OUT/'GIU_Connect_ERD.pdf',OUT/'GIU_Connect_Relational_Schema.pdf',ERD,
        OUT/'SUBMISSION_README.txt']+[ROOT/f for f in SCHEMAS]
 with ZipFile(dest,'w',ZIP_DEFLATED,compresslevel=9) as z:
  for f in files:z.write(f,f.name)
 with ZipFile(dest) as z:assert z.testzip() is None
 print('ZIP:',dest, 'bytes',dest.stat().st_size)

if __name__=='__main__':
 make_erd()
 make_schema()
 make_zip()
