import openpyxl as xl
import smtplib
wb = xl.load_workbook('Book1.xlsx')
sheet1=wb.get_sheet_by_name('Sheet1')
names=[]
university=[]
emails=[]
for cell in sheet1['A']:
       emails.append(cell.value)
for cell in sheet1['B']:
       names.append(cell.value)
for cell in sheet1['C']:
       university.append(cell.value)

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

email_user = '<your email from which you have to send email>'
email_password = '<your password>'

subject = 'Possible Undergraduate Research Opportunities During Summers'

for i in range(len(emails)):
       msg = MIMEMultipart('alternative')
       msg['From'] = "<Your Name>"
       msg['To'] = emails[i]
       msg['Subject'] = subject

       body =               """\
              <html>
                <body>
                  <p>Dear Professor {},<br><br>
                     I am an undergraduate second-year student in the Department of Mechanical Engineering at <b>Indian Institute of Technology, Delhi</b>(Ranked 1st in India).I am looking for an opportunity to do a <b>Summer Internship</b> under you in {} for the duration of 10 weeks (May- July 2020).<br><br>
                     Impressed by your research interests, I was wondering if there be an opportunity to work with you and assist you in the coming Summer in your prestigious University as I am extremely motivated to work under you. I ensure you of relentless cooperation and utmost dedication if I am given a chance. I am willing to learn anything in advance should you want me to. I am enclosing a Curriculum Vitae for your kind perusal.<br><br>
                     I am really interested in the field of <b>Artificial Intelligence, Machine Learning, and Data Structure</b>. I have maintained a good academic record throughout my school and college. I have been doing courses on <b>Artificial Intelligence, Machine Learning, Data Science, Probability, Statistics and Data Structures</b>. I am also working on a <b>Self-Driving Car project</b> which uses Deep Neural Network. I am <b>Department Rank 1</b> of Production and Industrial Engineering and won the <b>IITD Semester Merit Award</b> for being among Top 7% students of my batch in Semester III.<br><br>
                     Kindly consider my application and convey me of any suitable openings that your Institution may offer. I look forward to receiving a positive response from your side and remain at your disposal for any further information you may require or an eventual interview.<br><br>
                     Thank you for your time and consideration.<br><br>
                     Yours Sincerely,<br>
                     Your Name<br>
                     Email id:<br>
                     Contact no.:
                  </p>
                </body>
              </html>
              """.format(names[i],university[i])

       msg.attach(MIMEText(body,'html'))

       filename='<resume file name with extension'
       attachment  =open(filename,'rb')

       part = MIMEBase('application','octet-stream')
       part.set_payload((attachment).read())
       encoders.encode_base64(part)
       part.add_header('Content-Disposition',"attachment; filename= "+filename)

       msg.attach(part)
       text = msg.as_string()
       server = smtplib.SMTP('smtp.gmail.com',587)
       server.starttls()
       server.login(email_user,email_password)
       server.sendmail(email_user,emails[i],text)
       print("Email sent to",emails[i])
server.quit()
print("All emails sent!")