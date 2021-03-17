import pandas as pd
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

'''total 6 things needs to be changed
    from_emai
    from_password
    database path
    contact_no
    my_name
    my_position'''

from_email = "Enter your email"


# use app password
# to use app
# step 1= manage your google account
# step 2= Go to Security
# step 3= turn on 2-step verification
# now create you app password (jgzpotkwrdjrxzgt)


from_password = "Enter your app pass"

'''
format=
NAME         EMAIL
COMP1        EMAIL1
'''
database = pd.read_excel("Path to the excel")

names = database['NAME']
emails = database['EMAIL']

contact_no = "Enter your contact number"
my_name = "Enter your name"
my_position = "Your position"


for i in range(len(names)):
    message = MIMEMultipart()
    message["Subject"] = "Invitation for DJ ACM's Internship Fair"
    message["From"] = from_email
    message["To"] = emails[i]
    html = """<p>Greetings,</p>
    <br>
    <p>&nbsp; &nbsp; &nbsp; &nbsp; Dwarkadas J. Sanghvi's Association of Computing Machinery (ACM ) is a student chapter that helps students by connecting their curriculum to the industry to facilitate a structured path from education to employment. The committee strongly believes in the 'learning by doing' ethos and aims to give the next generation of engineers an opportunity to explore several domains.</p>
    <br>
    <p>&nbsp; &nbsp; &nbsp; &nbsp; This is to bring to your kind notice that DJ ACM is organizing it's annual Internship Fair on the 9th and 10th of April that will be held virtually this year. We are seeking companies that are currently looking to fill both technical and/or non-technical internship positions with individuals having relevant work experience and would be honoured to have your esteemed company on board. We can discuss the job description with your representatives, learn the specific requirements and compile a resume list to allow for efficient filtering. Executives from {company_name} can conduct the interviews of the shortlisted candidates in the aforementioned timeframe. The stipend, which is a necessity, is to be discussed with the selected candidates. Please note that the candidates should not be required to pay any course/training fees during the duration of their internship. {company_name} would receive extensive publicity on our social media handles and banners, posters and standees of the company would be put up during the event to ensure maximum possible reach within the student community.
    </p>
    <br>
    <p>&nbsp; &nbsp; &nbsp; &nbsp; We look forward to a mutually beneficial collaboration. Our team would be happy to coordinate the details or clarify questions. Please reach out to us at {from_email} or {contact_no}. Please do respond at your earliest convenience. </p>
    <br>
    <div style="padding:0;">Thanking you in anticipation,<br>
    {name}<br>
    {position}</div>
    """.format(company_name=names[i], from_email=from_email, name=my_name, position=my_position, contact_no=contact_no)
    content = MIMEText(html, "html")
    message.attach(content)
    context = ssl.create_default_context()
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(message["From"], from_password)
    server.sendmail(message["From"], message["To"], message.as_string())
    server.quit()

    print("sent", i+1)
