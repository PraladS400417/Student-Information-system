from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pathlib import Path
import json

ROOT=Path(__file__).parent
OUT=ROOT/'Pralad_Raya_S400417_Detailed_Test_Report.docx'
E=ROOT/'pralad_evidence'
performance=json.loads((E/'performance.json').read_text())
d=Document(); s=d.sections[0]
s.page_width=Inches(8.5); s.page_height=Inches(11)
s.left_margin=s.right_margin=Inches(.82); s.top_margin=Inches(.72); s.bottom_margin=Inches(.65)
styles=d.styles
styles['Normal'].font.name='Aptos'; styles['Normal'].font.size=Pt(10.5)
styles['Normal'].paragraph_format.space_after=Pt(6)
for name,size in [('Title',21),('Heading 1',15),('Heading 2',12),('Heading 3',10.5)]:
 st=styles[name]; st.font.name='Aptos Display' if name=='Title' else 'Aptos'; st.font.size=Pt(size); st.font.color.rgb=RGBColor(0,0,0); st.font.bold=True
 st.font.underline=False
 st.paragraph_format.space_before=Pt(12 if name!='Title' else 0); st.paragraph_format.space_after=Pt(5)
 if name=='Title':
  ppr=st.element.get_or_add_pPr()
  for old in ppr.findall(qn('w:pBdr')): ppr.remove(old)
styles['Subtitle'].font.color.rgb=RGBColor(0,0,0)
styles['Subtitle'].font.italic=False

def para(text='',style=None): return d.add_paragraph(text,style)
def heading(text,l=1): d.add_heading(text,l)
def caption(text):
 p=para(text); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
 for r in p.runs: r.italic=True; r.font.size=Pt(9)
 return p
def table(title,headers,rows):
 caption(title); t=d.add_table(rows=1,cols=len(headers)); t.style='Table Grid'; t.alignment=WD_TABLE_ALIGNMENT.CENTER
 for i,h in enumerate(headers): t.rows[0].cells[i].text=h
 for row in rows:
  cs=t.add_row().cells
  for i,v in enumerate(row): cs[i].text=str(v)
 for row in t.rows:
  for c in row.cells:
   c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
   tcPr=c._tc.get_or_add_tcPr(); mar=OxmlElement('w:tcMar')
   for side in ('top','left','bottom','right'):
    z=OxmlElement('w:'+side); z.set(qn('w:w'),'85'); z.set(qn('w:type'),'dxa'); mar.append(z)
   tcPr.append(mar)
   for p in c.paragraphs:
    p.paragraph_format.space_after=Pt(0)
    for r in p.runs: r.font.size=Pt(8.5)
 for c in t.rows[0].cells:
  sh=OxmlElement('w:shd'); sh.set(qn('w:fill'),'DDEAF4'); c._tc.get_or_add_tcPr().append(sh)
  for r in c.paragraphs[0].runs: r.bold=True
 d.add_paragraph('')
 return t
def fig(file,number,description,width=6.4):
 p=d.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
 p.add_run().add_picture(str(E/file),width=Inches(width))
 caption(f'Figure {number}. {description}')
def page(): d.add_page_break()

d.add_paragraph('Student Record System Individual Testing Report','Title')
p=para('PRT582 Software Group Project'); p.style='Subtitle'
para('Pralad Raya  |  Student ID S400417')
para('Teacher workflow  |  Functional and non functional testing')
para('4 October 2026')
para('Group name and number: to be inserted by the group before submission.')
para('This report records tests I performed on a local copy of the Student Record System, including Robot Framework execution, browser screenshots, database verification, a response time baseline, and a controlled security probe. Results are limited to the teacher workflow assigned to me.')
table('Table 1. Report evidence at a glance',['Evidence','Result'],[
('Local environment','PHP 7.4.33, MariaDB 10.4.14, Microsoft Edge 154.0.4258.53'),
('Robot Framework','7 tests passed, 0 failed after local buffering configuration'),
('Browser and database checks','Teacher login and PIN worked; profile update worked with debug output; enrollment failed'),
('Performance baseline','90 sequential local requests, all HTTP 200'),
('Security probe','Teacher login bypass observed with a crafted registration input')])
page()
heading('Contents')
for line in ['Executive Summary','1 Introduction and Background','2 Individual Contribution Statement','3 Requirements and Test Planning','4 Individual Functional Testing','5 Individual Non Functional Testing','6 Defect Register and Analysis','7 Ethical Consideration','8 Conclusion and Reflection','Appendix A Reproduction and Test Data','Appendix B Robot Framework Evidence','Appendix C Figures and Data Files']:
 para(line)
heading('List of Figures',2)
for n,title in [(1,'Teacher sign in form'),(2,'Teacher area after valid login'),(3,'Wrong PIN feedback'),(4,'Language enrollment form'),(5,'Enrollment failure'),(6,'Teacher profile debug output'),(7,'Security probe result')]: para(f'Figure {n}. {title}')
heading('Evidence conventions',2)
para('PASS means the observed result met the stated expectation. FAIL means it did not. A source finding is a code or schema observation, and is not presented as a runtime exploit unless a corresponding test was executed. “Not tested” identifies work outside this run. Screenshots show a disposable local database only.')
page()
heading('Executive Summary')
para('The teacher portal was launched successfully on a local PHP and MariaDB stack. I executed a seven-case Robot Framework browser suite, with all seven passing in the final run. Separate browser checks verified the PIN and profile workflows and exposed a failed language enrollment. A controlled local login probe also bypassed authentication with an incorrect password. These two defects require attention before the system can be treated as reliable or secure. The profile update printed raw submitted fields on the page, which is a further privacy and usability defect.')
para('The result is mixed rather than a clean pass: basic navigation, authentication with valid and invalid ordinary credentials, access redirection, PIN verification, and profile display worked, while enrollment did not save and malicious input bypassed login. Response times were low in a single-user local baseline, but these numbers do not establish production capacity. No participant usability study or concurrent load test was conducted.')
heading('1 Introduction and Background')
para('The Student Record System is a PHP and MySQL web application with separate teacher and administrator portals. My assignment within the group was to test the teacher-side functions: login, PIN verification, language course enrollment, teacher profile, password management, and enrollment history. The project repository includes the PHP source, static assets, and a SQL dump; it has no Composer or npm application dependency.')
para('The assignment brief specifies an independent individual report with a contribution statement, usability, performance, and security testing, functional acceptance scenarios, Robot Framework results, ethics, and personal reflection. The LocalLink final report was used as a reference for report depth, cover matter, tables, figures, and explanatory style. This document follows the PRT582 individual structure rather than copying LocalLink project content.')
heading('1.1 Aim',2)
para('The aim was to establish whether a teacher can complete the core workflow on a locally running copy, identify reproducible defects, and provide evidence that another tester can repeat. The acceptance focus was the observable teacher experience and the corresponding database effects.')
heading('1.2 Scope and limits',2)
para('The tested environment was local and isolated. I did not test the other group members’ admin features, deploy to a public server, conduct interviews with users, or run a concurrent load campaign. The seeded account used for browser tests was 070101, present in the supplied SQL dump. After changing its displayed name during testing, I restored it to Rahul. No permanent application source changes were made to repair defects.')
heading('2 Individual Contribution Statement')
para('I traced teacher pages to their data tables, prepared acceptance cases, wrote and ran the Robot Framework suite, captured browser evidence, checked the database after form submissions, measured local response times, and performed a constrained authentication probe on the disposable copy. I documented both successful and failed outcomes. The group name and number were not provided, so those cover fields remain marked for completion.')
table('Table 2. Contribution map',['Activity','My deliverable','Evidence'],[
('Environment setup','Portable PHP and MariaDB; Acourse imported from sqlfile/acourse.sql','Local HTTP 200 and 12 imported tables'),
('Functional automation','Seven Robot teacher acceptance cases','output.xml, log.html, report.html'),
('Browser walkthrough','Screenshots of login, PIN, enrollment, and profile','Figures 1–6 and observations.json'),
('Non functional work','Usability inspection, sequential timing baseline, local security probe','Figure 7 and performance.json'),
('Defect analysis','Three confirmed defects plus configuration and source risks','Table 8')])

heading('3 Requirements and Test Planning')
heading('3.1 User stories and acceptance criteria',2)
table('Table 3. Teacher requirements',['ID','User story','Acceptance criterion'],[
('F01','As a teacher, I can log in and log out.','Valid credentials open the teacher area; invalid ordinary credentials remain at login.'),
('F02','As a teacher, I can reach protected pages only after login.','Unauthenticated enrollment request redirects to login.'),
('F03','As a teacher, I can verify my PIN.','Wrong PIN is rejected; correct PIN opens enrollment.'),
('F04','As a teacher, I can enroll in a language course.','A valid submission creates one courseenrolls row and shows success.'),
('F05','As a teacher, I can view and update my profile.','Correct registration number and saved name display without debug output.'),
('F06','As a teacher, I can view enrollment history.','History page loads and shows the relevant enrollment type.'),
('NF01','The forms are understandable and accessible.','Labels, errors, role terms, and keyboard interactions are clear.'),
('NF02','The local workflow responds promptly.','Single-user page responses can be measured reproducibly.'),
('NF03','Authentication resists crafted input.','A malformed ID does not bypass password validation.')])
heading('3.2 Environment and setup',2)
para('A portable PHP 7.4.33 runtime and MariaDB 10.4.14 were placed in a local tooling directory. The Acourse database was created and sqlfile/acourse.sql was imported, producing 12 tables and three seeded staff records. PHP served the repository at 127.0.0.1:8000. The final Robot run used Microsoft Edge 154.0.4258.53 and a matching EdgeDriver. PHP mysqli and pdo_mysql extensions were enabled. Output buffering was set to 4096 bytes in the local php.ini because the PIN page otherwise showed a session header warning due to leading output before session_start(). This setting changed the test runtime, not the application file.')
para('The README’s onlinecourse database name and 10806121 / 123 teacher login do not match the supplied source and SQL. The configured database is Acourse, and the test account 070101 / 1234 is present in the SQL dump. These details matter because following the README verbatim would prevent a repeatable run.')
heading('3.3 Test approach and completion criteria',2)
para('I used Robot Framework with SeleniumLibrary for repeatable browser acceptance tests. I then used a separate Playwright walkthrough to capture screenshots and exercise form submissions. I queried MariaDB after the enrollment attempt and profile restoration. A small Python urllib script issued 30 sequential requests to each of three pages. For security, I used a single local crafted-ID login probe and checked the resulting session destination. A functional case passed only when the page outcome and, where relevant, database state matched the requirement.')
table('Table 4. Automation tool comparison',['Tool','Fit for task','Decision'],[
('Robot Framework with SeleniumLibrary','Readable acceptance cases and required HTML/XML reports; named by assignment','Selected for seven functional cases'),
('Playwright','Reliable scripted browser capture and detailed observations','Used for supplementary screenshots and form checks'),
('JMeter or k6','Useful for concurrent load and trend analysis','Not run; single-user local timing is reported instead'),
('OWASP ZAP','Useful for passive and authenticated scans','Not run; targeted local probe and code analysis are clearly labelled')])

heading('4 Individual Functional Testing')
heading('4.1 Robot Framework execution',2)
para('The final Robot run completed with 7 PASS and 0 FAIL. An earlier run had 5 PASS and 2 FAIL because the PIN page printed output before PHP could start its session under the initial local configuration. Setting output_buffering=4096 allowed the PIN tests to execute. The final report and raw log files are preserved in pralad_robot_results. A Selenium Manager update check warned that an online driver lookup failed; the matching official driver was already present locally and did not prevent execution.')
table('Table 5. Final Robot results',['Case','Scenario','Result'],[
('TC01','Valid teacher login','PASS'),('TC02','Invalid teacher password rejected','PASS'),('TC03','Protected enrollment redirects without login','PASS'),('TC04','Incorrect PIN rejected','PASS'),('TC05','Correct PIN opens enrollment','PASS'),('TC06','Profile shows registration number','PASS'),('TC07','Enrollment history page loads','PASS')])
heading('4.2 Manual browser and database results',2)
table('Table 6. Supplementary acceptance results',['ID','Observed test and evidence','Result'],[
('BT01','Valid login redirected to change-password.php','PASS'),
('BT02','Wrong PIN displayed “Wrong Pincode”','PASS'),
('BT03','Correct PIN opened enroll.php','PASS'),
('BT04','Profile showed teacher ID 070101','PASS'),
('BT05','Profile name changed to Rahul Test, then restored to Rahul; database confirmed restoration','PASS with defect: raw POST fields printed'),
('BT06','Completed language enrollment form returned “Error: Not Enroll”; row count for 070101 stayed at 7','FAIL'),
('BT07','Enrollment history page loaded','PASS for load; relation to language enrollment unresolved because BT06 failed'),
('BT08','Crafted teacher ID and incorrect password opened protected teacher page','FAIL security requirement')])
para('BT06 is a real failed submission. The courseenrolls table defines marks as NOT NULL without a default, while enroll.php inserts a row without marks. MariaDB was running STRICT_TRANS_TABLES, so the insert was rejected. The form displayed its generic failure message, and the database row count remained unchanged. This is a strong cause supported by schema and query inspection; adding explicit database error logging would make diagnosis more direct. The application was left unchanged for the report.')
heading('4.3 Visual evidence',2)
fig('02_login_form.png',1,'Teacher sign in form opened from the landing image',5.8)
fig('03_teacher_area.png',2,'Teacher area after valid login',5.8)
fig('04_wrong_pin.png',3,'Invalid PIN produces visible feedback',5.8)
fig('05_enrollment_form.png',4,'Language course enrollment form after valid PIN',5.8)
fig('06_enrollment_result.png',5,'Valid form attempt displays Error Not Enroll',5.8)

heading('5 Individual Non Functional Testing')
heading('5.1 Usability testing',2)
para('Method: I conducted a single expert walkthrough using the same teacher tasks as the acceptance cases and inspected visible labels, error feedback, and page flow. This was not a participant study and does not support claims about average user task time or satisfaction. Positive observations were a clear PIN error and a consistent navigation bar after login. Issues were the landing page’s image-only sign-in invitation, mixed Student and Teacher terminology, a profile page headed “Teacher Registration,” a generic enrollment failure without a cause, and raw POST values printed before the profile page content.')
table('Table 7. Usability observations',['Observation','Impact','Improvement'],[
('Sign in starts from a large image rather than a conventional button','May be missed by keyboard or screen-reader users','Use a semantic button with accessible text'),
('Teacher navigation includes Student terms','Role and task ambiguity','Use Teacher and Student consistently'),
('Enrollment reports only Error Not Enroll','No recovery guidance','Show field or server-side validation details without exposing internals'),
('Profile update prints POST fields','Distracts users and reveals form data','Remove var_dump and show only confirmation')])
fig('09_profile_update.png',6,'Profile update succeeded but exposed submitted fields above the page',5.8)
heading('5.2 Performance testing',2)
para('Method: 30 sequential HTTP GET requests were sent to each of the login, authenticated profile, and authenticated enrollment history pages through a cookie-preserving client. The local PHP development server and database ran on the same Windows host, with no concurrent virtual users. All 90 requests returned HTTP 200. This is a response-time baseline, not a scalability or production load result. The p95 shown below is the 95th percentile sample from the 30 observed times.')
table('Table 8. Sequential local response times',['Page','Requests','Median ms','P95 ms','Max ms'],[
(name.title(),v['requests'],v['median_ms'],v['p95_ms'],v['max_ms']) for name,v in performance.items()])
para('The tested pages were fast in this isolated setting. A later concurrent JMeter or k6 run should test multiple logged-in users, database growth, and resource use before any deployment claim. No such result is invented here.')
heading('5.3 Security testing',2)
para('The distinct security test used a fresh browser context and a controlled malformed teacher ID with an incorrect password against 127.0.0.1 only. The application redirected to the protected change-password page and displayed a teacher session, demonstrating an authentication bypass in this local run. The login query concatenates StaffRegno directly into SQL, so operator injection is possible. The finding is critical because a person without a valid password can enter the teacher area. The payload is stored in the local test script rather than repeated in this report. I also verified that a normal unauthenticated direct request to enroll.php redirected to login; this does not mitigate the bypass.')
para('Additional source risks, not separately exploited here, include MD5 password storage, direct use of posted identity fields in enrollment, and an unrestricted filename path for profile photo uploads. Recommended order: replace raw SQL concatenation with prepared statements everywhere, migrate passwords to password_hash/password_verify, derive account identity from the session, validate uploads, add CSRF tokens, and retest all entry points. The seeded credentials must remain confined to local testing.')
fig('10_security_probe.png',7,'Protected teacher page reached after crafted-ID login with an incorrect password',5.8)

heading('6 Defect Register and Analysis')
table('Table 9. Confirmed defects and risks',['ID','Severity','Evidence','Recommended action'],[
('DEF-01','Critical','Crafted-ID login reached protected page with wrong password','Parameterise login query; regression test malicious input'),
('DEF-02','High','Enrollment showed Error Not Enroll; DB count unchanged; insert omits required marks','Supply a validated marks value or schema default; expose safe error'),
('DEF-03','Medium','Profile POST array displayed above page after update','Remove var_dump from my-profile.php'),
('DEF-04','Medium','PIN page failed under unbuffered PHP because output precedes session_start()','Remove leading output and start session before response body'),
('RISK-05','High source risk','Passwords stored and checked with MD5','Migrate to modern password hashing'),
('RISK-06','Medium source risk','enroll.php writes courseenrolls; enroll-history.php reads student','Clarify intended histories and cover both with tests')])
heading('6.1 Reproduction priority',2)
para('Fix DEF-01 first, because it undermines every teacher permission check. Fix DEF-02 next so the assigned enrollment journey can be completed. DEF-03 and DEF-04 should then be corrected and regression tested. The risks listed after the confirmed defects need separate targeted tests before they are called exploitable or user-facing failures.')
heading('7 Ethical Consideration')
para('All tests ran on a local imported database using supplied sample accounts. The security probe was limited to localhost and used an incorrect password; no real account or remote host was targeted. Profile data changed for the test was restored, and the enrollment attempt did not create a row. Screenshots were kept in the project workspace. The report distinguishes actual observation from source analysis and from unperformed tests. Before submission, the group should ensure screenshots and logs do not disclose usable credentials or personal data beyond the supplied dummy records.')
heading('8 Conclusion and Reflection')
para('The teacher pages can be started and several core interactions work, as shown by the final 7/7 Robot pass and supplementary browser evidence. The complete teacher workflow does not pass: language enrollment fails, and the login query permits a local authentication bypass. Those findings are more consequential than the quick local response times. The work provided a reproducible test suite, screenshots, raw logs, a database-backed failure check, and a prioritised defect register for the group.')
para('The main technical lesson was to confirm a browser message against database state. The enrollment form’s generic error alone did not show why it failed; tracing its INSERT against the courseenrolls schema exposed the missing required marks field. A second lesson was to distinguish a passing ordinary invalid-login test from a security test using malformed input. Both can be true at the same time. The setup challenge was the mismatch between README text, configured database name, seeded credentials, and PHP output buffering. Recording those conditions makes the tests repeatable.')
heading('8.1 Further work',2)
para('After correcting the confirmed defects, rerun the Robot suite, add a database assertion for enrollment, test password change and upload validation with disposable accounts, conduct a participant usability study, run a controlled concurrent load test, and review other raw SQL queries. The group must insert its group name and number, and add any required YouTube presentation link and group-wide findings before final submission.')

heading('Appendix A Reproduction and Test Data')
para('Local runtime: PHP 7.4.33 with mysqli enabled and output_buffering=4096; MariaDB 10.4.14 with STRICT_TRANS_TABLES; Edge 154.0.4258.53. Base URL: http://127.0.0.1:8000. Database: Acourse imported from sqlfile/acourse.sql. Seeded test teacher: 070101. For safety, the password and PIN values are in the local test suite and should not be copied into a public report. The initial language enrollment row count for this teacher was 7; it remained 7 after the failed submission. The profile name was restored to Rahul.')
para('Reproduce DEF-02 by logging in, entering the valid teacher PIN, selecting populated Session, Department, Level, Semester, and Course options, then submitting. Observe Error Not Enroll. Query courseenrolls for the teacher and confirm no new row. Inspect enroll.php INSERT and the table schema. Reproduce DEF-03 by updating the profile name on a disposable account and viewing the raw array printed above the page. DEF-01 should be repeated only in an authorised local test environment; the recorded Robot suite covers the ordinary invalid-login case separately.')
heading('Appendix B Robot Framework Evidence')
para('Suite file: pralad_teacher_tests.robot. Run command used: python -m robot --variable BASE_URL:http://127.0.0.1:8000 --variable BROWSER:edge --outputdir pralad_robot_results pralad_teacher_tests.robot. The local Python environment included Robot Framework 7.5 and SeleniumLibrary 6.9.0. The output directory contains output.xml, log.html, report.html, and failure screenshots from earlier runs. The final output.xml is the source of the 7 PASS / 0 FAIL result. The final suite does not submit an enrollment row, so its green status must not be interpreted as full workflow success.')
heading('Appendix C Figures and Data Files')
para('Supporting evidence in pralad_evidence: 01_login_page.png to 10_security_probe.png, observations.json, enrollment_options.json, and performance.json. The Playwright walkthrough is capture_pralad_evidence.js; the timing script is measure_pralad.py. The test report cites the relevant screenshot numbers above. Source evidence is in index.php, pincode-verification.php, enroll.php, my-profile.php, enroll-history.php, includes/config.php, and sqlfile/acourse.sql. Assignment structure was read from PRT582 Software Group Project.docx. LocalLink_Final_Report.docx was used only as a reference for report depth and presentation.')
para('Video presentation link: to be supplied by the group if required for the combined submission.')

foot=s.footer.paragraphs[0]; foot.alignment=WD_ALIGN_PARAGRAPH.CENTER; foot.text='Pralad Raya  |  S400417  |  PRT582 Individual Testing Report'
d.save(OUT)
print(OUT)
