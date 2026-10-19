import os,re,json,io
from pathlib import Path
from fastapi import FastAPI,UploadFile,File,Form,HTTPException
import fitz,docx,gspread
from pptx import Presentation
from striprtf.striprtf import rtf_to_text
from google import genai
from google.genai import types
from google.oauth2.service_account import Credentials

app=FastAPI(title="Evalora AI")
SPREADSHEET_ID=os.getenv("SPREADSHEET_ID","1L8FuY_CFURGrRkWbEFAG_QsHTyF0oUJYfpCfNAl-sOw")
WORKSHEET=os.getenv("GOOGLE_WORKSHEET","Evaluations")
MODELS=["gemini-3.5-flash-lite"]  # Evalora uses only this model; no fallback models.
RATING_SCORES={"Outstanding":100,"Very Good":85,"Good":70,"Needs Improvement":40,"Incomplete":0}
BASE_HEADERS=["Name","Project ID","Project Name","Phase Type","Questions","AI Confidence Score","Feedback","Suggestions","Overall Score"]
EXTRA_HEADERS=["Model Name","Answer Accuracy (%)"]
HEADERS=BASE_HEADERS+EXTRA_HEADERS
PHASES={
"Ideation":["How well did the team articulate the problem statement and support it with relevant statistics, data or evidence?","Is the proposed solution addressing a meaningful need for the identified users/stakeholders?","How well has the team researched existing solutions and competitors?","Is the proposed solution feasible and unique?","How would you rate the overall quality of the Ideation submission?","Additional feedback to candidate"],
"Design":["How effectively does the design demonstrate the proposed solution?","How well does the design address the identified problem?","How effectively has the team demonstrated the core features?","How feasible is the design from a technical perspective?","How would you rate the overall quality of the Design submission?","Additional feedback to candidate"],
"Build":["How effectively has the team implemented the proposed solution?","How well does the implementation address the identified problem?","How effectively has the team demonstrated the core features?","How feasible is the implementation from a technical perspective?","How would you rate the overall quality of the Build submission?","Additional feedback to candidate"],
"Demo":["How effectively does the final solution address the stated problem?","How well does the solution demonstrate its key features?","How effectively has the team demonstrated testing and validation?","How feasible and sustainable is the proposed solution?","How would you rate the overall quality of the final submission?","Additional feedback to candidate"]}

def extract_document(filename,data):
 ext=Path(filename or '').suffix.lower()
 if ext in ['.txt','.md','.csv']: text=data.decode('utf-8-sig',errors='replace')
 elif ext=='.pdf':
  with fitz.open(stream=data,filetype='pdf') as pdf: text='\n'.join(p.get_text() for p in pdf)
 elif ext=='.docx':
  d=docx.Document(io.BytesIO(data)); parts=[p.text for p in d.paragraphs]
  for t in d.tables:
   for row in t.rows: parts.append(' | '.join(c.text for c in row.cells))
  text='\n'.join(parts)
 elif ext=='.pptx':
  prs=Presentation(io.BytesIO(data)); slides=[]
  for n,s in enumerate(prs.slides,1):
   parts=[f'Slide {n}']
   for sh in s.shapes:
    if getattr(sh,'has_text_frame',False) and sh.text.strip(): parts.append(sh.text.strip())
    if getattr(sh,'has_table',False):
     for row in sh.table.rows: parts.append(' | '.join(c.text for c in row.cells))
   try:
    if s.has_notes_slide and s.notes_slide.notes_text_frame.text.strip(): parts.append('Speaker notes: '+s.notes_slide.notes_text_frame.text)
   except Exception: pass
   slides.append('\n'.join(parts))
  text='\n\n'.join(slides)
 elif ext=='.rtf': text=rtf_to_text(data.decode('utf-8',errors='replace'))
 else: raise ValueError('Supported files: TXT, PDF, DOCX, PPTX, RTF, MD, CSV. Convert legacy DOC/PPT files first.')
 if not text.strip(): raise ValueError('No readable text found. Scanned PDFs or image-only slides require OCR.')
 return text.strip()[:50000]

def parse_json(s):
 s=re.sub(r'^```(?:json)?\s*|\s*```$','',(s or '').strip(),flags=re.I)
 try: v=json.loads(s)
 except json.JSONDecodeError:
  m=re.search(r'\{.*\}',s,re.S)
  if not m: raise ValueError('Gemini returned invalid JSON.')
  v=json.loads(m.group())
 if not isinstance(v,dict): raise ValueError('Gemini response must be a JSON object.')
 return v

def normalize_rating(v):
 v=str(v or '').strip().lower()
 for phrase,rating in [('needs improvement','Needs Improvement'),('outstanding','Outstanding'),('very good','Very Good'),('incomplete','Incomplete'),('good','Good')]:
  if re.search(r'\b'+re.escape(phrase)+r'\b',v): return rating
 return 'Incomplete'

def normalize_answers(raw,questions):
 lookup={}
 for x in raw if isinstance(raw,list) else []:
  if isinstance(x,dict) and str(x.get('question','')).strip(): lookup[str(x['question']).strip()]=x.get('answer',x.get('rating',''))
 out=[]
 for i,q in enumerate(questions):
  a=lookup.get(q,'')
  a=normalize_rating(a) if i<5 else (str(a or '').strip() or 'No additional feedback provided.')
  item={'question':q,'answer':a}
  if i < 5:
   item['rating_score']=RATING_SCORES[a]
  out.append(item)
 return out

def add_rating_scores_to_report(report):
 """Ensure report question ratings include their numeric rating scores."""
 if not report:
  return report
 pattern=r'(\*\*Rating:\*\*\s*)(Outstanding|Very Good|Needs Improvement|Incomplete|Good)(?!\s*\n\s*\*\*Rating Score:\*\*)'
 def replace(match):
  rating=match.group(2)
  return f'{match.group(1)}{rating}\n\n**Rating Score:** {RATING_SCORES[rating]}/100'
 return re.sub(pattern,replace,report)

def get_phase_ratings(report):
 """Extract display-ready question ratings from the saved evaluation report."""
 report=add_rating_scores_to_report(report or '')
 pattern=r'### Q(\d+):\s*(.*?)\n\n\*\*Rating:\*\*\s*(Outstanding|Very Good|Needs Improvement|Incomplete|Good)'
 items=[]
 for number,question,rating in re.findall(pattern,report):
  items.append({'question_number':int(number),'question':question.strip(),'rating':rating,'rating_score':RATING_SCORES[rating]})
 return items

def normalize_compliance(raw):
 if not isinstance(raw,list): return []
 valid={'satisfied':'Satisfied','partially satisfied':'Partially Satisfied','not satisfied':'Not Satisfied','not applicable':'Not Applicable'}; out=[]
 for x in raw:
  if not isinstance(x,dict): continue
  req=str(x.get('requirement','')).strip()
  if req: out.append({'requirement':req,'status':valid.get(str(x.get('status','')).strip().lower(),'Not Satisfied'),'evidence':str(x.get('evidence','')).strip() or 'No supporting evidence identified.','gap':str(x.get('gap','')).strip() or 'No additional gap explanation provided.'})
 return out

def create_prompt(name, project_id, project_name, phase, guidelines="", submission="", sample_reference=""):
 qs = PHASES[phase]
 qtext = "\n".join(f"Q{i}: {q}" for i, q in enumerate(qs, 1))
 if phase == "Ideation":
  guideline_section = guidelines[:20000] if guidelines.strip() else "No readable guideline text was supplied."
  sample_section = sample_reference[:15000] if sample_reference.strip() else "No separate sample project PPT was supplied. Do not infer mandatory requirements from a missing sample."
  ideation_instructions = '''
PHASE SCOPE — IDEATION ONLY:
The guideline document and sample project PPT apply ONLY to the Ideation phase.
The OFFICIAL GUIDELINES define explicit mandatory requirements and evaluation criteria. The SAMPLE PROJECT PPT is only an example to help students understand structure, presentation, and expected level of detail. Its topic, domain, project title, technology, features, and solution are NOT mandatory just because they appear in the sample.
Evaluate the student's own project idea. Do not penalize a different topic from the sample. Extract requirements for Part B only from explicit official guideline requirements, not illustrative sample content. If formal requirements and examples appear mixed together and cannot be distinguished, do not assume the example topic is mandatory.
'''
  compliance_instructions = '''PART B — IDEATION GUIDELINE COMPLIANCE:
Extract every distinct actionable requirement explicitly stated in the official Ideation guidelines. Do not convert the sample PPT's topic/title/features into requirements. Search the student submission for semantic evidence; synonyms and equivalent wording count. For each requirement return status exactly one of Satisfied, Partially Satisfied, Not Satisfied, Not Applicable. Missing evidence is Not Satisfied, not Not Applicable. Include requirement, evidence (quote briefly or faithfully paraphrase; say when none is found), and gap/improvement. A feature being mentioned is not proof it was implemented or tested. Do not invent statistics, tests, page numbers, or citations. If no explicit requirements can be extracted, return an empty list and mention the limitation in feedback. Keep compliance separate from phase ratings.'''
 else:
  guideline_section = "NOT APPLICABLE: This guideline document is used only for the Ideation phase. Do not use or infer any requirements from it for this evaluation."
  sample_section = "NOT APPLICABLE: The sample project PPT is used only for the Ideation phase. Do not use its topic, domain, content, format, or features to evaluate this phase."
  ideation_instructions = '''
PHASE SCOPE — NON-IDEATION:
This is a Design, Build, or Demo evaluation. Evaluate the student's actual submission using only the selected phase's questions and the phase-appropriate evidence present in the student submission. The Ideation guideline document and sample project PPT do not apply to this phase and must not affect ratings, feedback, suggestions, or requirements. Do not include any claim that the student failed to follow the Ideation guideline/sample PPT.
'''
  compliance_instructions = '''PART B — GUIDELINE COMPLIANCE:
Guideline compliance is only performed during the Ideation phase. For this Design, Build, or Demo evaluation, return an empty guideline_compliance array: []. Do not extract or assess requirements from the Ideation guideline document or sample project PPT.'''
 return f'''You are Evalora AI, an evidence-based project evaluator. Perform (A) phase-specific evaluation and (B) guideline compliance only where applicable. Compare meaning and evidence, not literal wording. Never invent evidence.

STUDENT DETAILS: Name={name}; Project ID={project_id}; Project={project_name}; Phase={phase}
PHASE QUESTIONS:
{qtext}
{ideation_instructions}
OFFICIAL GUIDELINES INPUT:
{guideline_section}

SAMPLE PROJECT PPT / REFERENCE INPUT:
{sample_section}

STUDENT SUBMISSION:
{submission[:30000]}

PART A — PHASE-SPECIFIC RATINGS:
CONSISTENCY RULES (MANDATORY): Apply the same rubric consistently on every evaluation. Judge only evidence present in the student submission; do not reward or penalize writing style, project topic, or similarity to the sample unless the selected question explicitly assesses it. For each Q1-Q5, internally identify the exact supporting evidence and the most important missing evidence before choosing a rating. Use these fixed thresholds: Outstanding = all key elements are explicitly evidenced with specific details and no material gap; Very Good = most key elements are evidenced with only minor gaps; Good = core element is present but one or more meaningful details/evidence are missing; Needs Improvement = major evidence or key elements are missing; Incomplete = the answer/section is absent, unreadable, or not assessable. Do not upgrade a rating based on generic claims without evidence. If evidence is ambiguous, choose the lower rating that the evidence clearly supports. Q1-Q4 must each be judged against their own question, and Q5 must summarize the quality of the whole submission using the same thresholds. Never change a rating merely to create score variation.
Evaluate Q1-Q5 independently using the student's submission and the selected phase's criteria. For Ideation, use explicit official guidelines as relevant but never treat sample topic/content as mandatory. For Design, Build, and Demo, use the student submission and the selected phase's questions only; do not use the Ideation guideline or sample PPT. Return only these exact ratings: Outstanding, Very Good, Good, Needs Improvement, Incomplete. Outstanding=strong evidence/no major gaps; Very Good=minor gaps; Good=some meaningful details missing; Needs Improvement=major gaps; Incomplete=absent/not assessable. Q5 is an independent overall quality rating. Q6 is constructive feedback, not a rating. Give specific actionable suggestions grounded in the student's own project. ai_confidence_score is 0-100 confidence in the readability and completeness of the student evidence, not confidence that the project matches the sample. answer_accuracy is a separate AI-estimated 0-100 assessment of how accurately the ratings, feedback, and suggestions reflect evidence in the student's submission. Be conservative: deduct for unsupported claims, overlooked contradictory evidence, or conclusions not justified by the submission. This is an AI estimate, not independently verified ground-truth accuracy. Do not return overall score; app calculates it. Mapping: Outstanding 100, Very Good 85, Good 70, Needs Improvement 40, Incomplete 0.

{compliance_instructions}

Return valid JSON only, no markdown fences, with exactly this structure and actual results:
{{"phase_answers":[{{"question":"{qs[0]}","answer":"Good"}},{{"question":"{qs[1]}","answer":"Very Good"}},{{"question":"{qs[2]}","answer":"Needs Improvement"}},{{"question":"{qs[3]}","answer":"Good"}},{{"question":"{qs[4]}","answer":"Needs Improvement"}},{{"question":"{qs[5]}","answer":"Constructive feedback."}}],"ai_confidence_score":80,"answer_accuracy":90,"feedback":"Evidence-based overall feedback.","suggestions":"Specific actionable suggestions.","guideline_compliance":[]}}
Example values are structural only; generate actual answers from the documents. For Ideation, populate guideline_compliance with actual requirement-level results; for Design, Build, or Demo, return an empty array.'''

def call_gemini(prompt):
 # Use Gemini 3.5 Flash-Lite only. Limit output size to reduce latency and
 # keep the request comfortably within the configured Vercel function duration.
 key = os.getenv('GEMINI_API_KEY')
 if not key:
  raise ValueError('GEMINI_API_KEY is missing from Vercel Environment Variables.')
 model = 'gemini-3.5-flash-lite'
 try:
  client = genai.Client(api_key=key, http_options=types.HttpOptions(timeout=180000))
  res = client.models.generate_content(
   model=model,
   contents=prompt,
   config=types.GenerateContentConfig(
    response_mime_type='application/json',
    temperature=0,
    max_output_tokens=4096,
   ),
  )
  if not res.text:
   raise ValueError('Gemini returned an empty response. Please retry once.')
  return parse_json(res.text), model
 except Exception as e:
  # Return a readable backend error rather than silently switching models.
  raise RuntimeError(f'{model} evaluation failed: {str(e)[:700]}') from e

def connect_sheets():
 """Connect to the visible evaluations sheet and keep it to columns A:K.

 Each request is evaluated anew; this function does not create or use a cache.
 If an older version added Input Fingerprint and Cached Report in columns L:M,
 those trailing columns are removed while the existing A:K evaluation records
 are preserved.
 """
 raw=os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON')
 if not raw: raise ValueError('GOOGLE_SERVICE_ACCOUNT_JSON is missing from Vercel Environment Variables.')
 creds=Credentials.from_service_account_info(json.loads(raw),scopes=['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive'])
 book=gspread.authorize(creds).open_by_key(SPREADSHEET_ID)
 try:
  ws=book.worksheet(WORKSHEET)
 except gspread.WorksheetNotFound:
  ws=book.add_worksheet(title=WORKSHEET,rows=1000,cols=len(HEADERS))

 old_values=ws.get_all_values()
 old_headers=old_values[0] if old_values else []
 if old_headers and old_headers[:len(BASE_HEADERS)] != BASE_HEADERS:
  raise ValueError('Worksheet row 1 must begin with these nine headers: '+' | '.join(BASE_HEADERS))

 # The two requested cache columns are not needed. Keep only A:K and preserve
 # all existing evaluation rows in those columns.
 if ws.col_count != len(HEADERS):
  ws.resize(cols=len(HEADERS))
 ws.update(range_name='A1:K1',values=[HEADERS],value_input_option='RAW')
 return ws

def append_to_google_sheet(row):
 if len(row)!=len(HEADERS): raise ValueError(f'Google Sheets expects exactly {len(HEADERS)} columns (A:K).')
 ws=connect_sheets()
 ws.append_row(row,value_input_option='RAW')

def compliance_report(items, phase="Ideation"):
 lines=['### Guideline vs Student Submission','','Assessment compares meaning and evidence, not literal text matching.','']
 if phase != "Ideation":
  return '\n'.join(lines+['**Not evaluated for this phase. The guideline document and sample project PPT are used only for Ideation.**'])
 if not items: return '\n'.join(lines+['**Compliance assessment unavailable.** No usable requirement-level results were returned.'])
 labels={'Satisfied':'SATISFIED','Partially Satisfied':'PARTIALLY SATISFIED','Not Satisfied':'NOT SATISFIED','Not Applicable':'NOT APPLICABLE'}; counts={k:0 for k in labels}
 for i,x in enumerate(items,1):
  s=x['status']; counts[s]+=1; lines += [f"#### Requirement {i}: {x['requirement']}",'',f"**Status:** {labels[s]}",'',f"**Student Evidence:** {x['evidence']}",'',f"**Gap / Improvement:** {x['gap']}",'','---','']
 applicable=len(items)-counts['Not Applicable']; lines += ['### Compliance Summary','',f'- Total requirements identified: {len(items)}',f"- Satisfied: {counts['Satisfied']}",f"- Partially Satisfied: {counts['Partially Satisfied']}",f"- Not Satisfied: {counts['Not Satisfied']}",f"- Not Applicable: {counts['Not Applicable']}"]
 lines += [f"- Fully satisfied applicable requirements: {round(counts['Satisfied']/applicable*100,2)}%" if applicable else '- No applicable requirements available for percentage calculation.','', 'This percentage is not the overall project score.']
 return '\n'.join(lines)

@app.get('/api/health')
def health(): return {'ok':True,'service':'Evalora AI','models':MODELS}
@app.get('/api/phases')
def phases(): return {'phases':list(PHASES),'questions':PHASES}
@app.post('/api/evaluate')
async def evaluate(
 name: str = Form(...),
 project_id: str = Form(...),
 project_name: str = Form(...),
 phase: str = Form(...),
 guideline_file: UploadFile | None = File(None),
 submission_file: UploadFile = File(...),
 sample_file: UploadFile | None = File(None),
):
 if phase not in PHASES:
  raise HTTPException(400, 'Select a valid phase.')
 try:
  sd = await submission_file.read()
  if len(sd) > 8 * 1024 * 1024:
   raise HTTPException(413, 'The student submission must be 8 MB or smaller.')
  submission = extract_document(submission_file.filename, sd)

  # Guidelines and sample reference are used in Ideation only.
  guidelines = ""
  sample_text = ""
  if phase == "Ideation":
   if guideline_file is None or not guideline_file.filename:
    raise HTTPException(400, 'Upload the official evaluation guideline for the Ideation phase.')
   gd = await guideline_file.read()
   if len(gd) > 8 * 1024 * 1024:
    raise HTTPException(413, 'The Ideation guideline file must be 8 MB or smaller.')
   guidelines = extract_document(guideline_file.filename, gd)
   if sample_file is not None and sample_file.filename:
    smd = await sample_file.read()
    if len(smd) > 8 * 1024 * 1024:
     raise HTTPException(413, 'The sample project PPT must be 8 MB or smaller.')
    sample_text = extract_document(sample_file.filename, smd)
  else:
   # Intentionally do not read, extract, or send the Ideation guideline/sample content
   # to Gemini for Design, Build, or Demo evaluations.
   guidelines = ""
   sample_text = ""

  # IMPORTANT: Always run a fresh Gemini evaluation for every submission,
  # including when the same document was submitted earlier. Each result gets
  # its own new row in Google Sheets; no old score is reused.
  result, model = call_gemini(create_prompt(name, project_id, project_name, phase, guidelines, submission, sample_text))
  answers = normalize_answers(result.get('phase_answers', []), PHASES[phase])
  compliance = normalize_compliance(result.get('guideline_compliance', [])) if phase == 'Ideation' else []
  score = round(sum(RATING_SCORES[x['answer']] for x in answers[:5]) / 5, 2)
  try:
   confidence = float(result.get('ai_confidence_score', 0))
  except (ValueError, TypeError):
   confidence = 0
  confidence = max(0, min(100, confidence))
  try:
   answer_accuracy = round(max(0, min(100, float(result.get('answer_accuracy')))), 2)
  except (ValueError, TypeError):
   answer_accuracy = ''
  feedback = str(result.get('feedback', '')).strip() or 'No overall feedback provided.'
  suggestions = str(result.get('suggestions', '')).strip() or 'No suggestions provided.'
  questions = '\n\n'.join((f"Q{i}: {x['question']}\nRating: {x['answer']} | Rating Score: {x['rating_score']}/100" if i <= 5 else f"Q{i}: {x['question']}\nFeedback: {x['answer']}") for i, x in enumerate(answers, 1))
  lines = ['# Evaluation Results', '', f'**Student:** {name}', f'**Project ID:** {project_id}', f'**Project:** {project_name}', f'**Phase:** {phase}', '', '## Overall Score: ' + str(score) + '/100', '', f'**AI Confidence:** {confidence}/100', '', '## Phase Ratings', '']
  for i, x in enumerate(answers[:5], 1):
   lines += [f"### Q{i}: {x['question']}", '', f"**Rating:** {x['answer']}", '', f"**Rating Score:** {x['rating_score']}/100", '']
  lines += ['## Additional Feedback', '', answers[5]['answer'], '', '## Overall Feedback', '', feedback, '', '## Suggestions', '', suggestions, '', compliance_report(compliance, phase)]
  report='\n'.join(lines)
  phase_ratings=[{'question_number':i,'question':x['question'],'rating':x['answer'],'rating_score':x['rating_score']} for i,x in enumerate(answers[:5],1)]
  sheet_row=[name, project_id, project_name, phase, questions, confidence, feedback, suggestions, score, model, answer_accuracy]
  append_to_google_sheet(sheet_row)
  compliance_status = 'Ideation guideline compliance included.' if phase == 'Ideation' else 'Ideation guideline/sample PPT skipped for this phase.'
  return {'success': True, 'report': report, 'phase_ratings': phase_ratings, 'rating_score_mapping': RATING_SCORES, 'status': f'Fresh evaluation completed using {model}. Saved as a new row in Google Sheets. {compliance_status}', 'overall_score': score, 'confidence': confidence, 'model_used': model, 'sheet_saved': True, 'cache_hit': False}
 except HTTPException:
  raise
 except Exception as e:
  raise HTTPException(500, str(e)[:1200])
