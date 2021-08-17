import openpyxl,os,smtplib


userEmail=input("Enter Email")
userPassword=input("Enter Password")

conn=smtplib.SMTP('smtp.gmail.com',587)
conn.ehlo()
conn.starttls()
conn.login(userEmail,userPassword)

# NOT NECESSARY BUT THIS IS EASIER THAN WRITING FULL PATH OF THE FILE
os.chdir('C:\\Users\\hp\\Desktop\\PsychBot')

event=openpyxl.load_workbook('Event.xlsx')

# REPLACE sheetName WITH NAME OF THE SHEET
responses=event['sheetName']

emails=[]


# USE YOUR EMAIL COLUMN LETTER 
for i in responses['E']:
    emails.append(i.value)

# REPLACE email WITH NAME OF THE COLUMN
emails.remove('Email')

for email in emails:
    print(email)
    conn.sendmail(from_addr=userEmail,to_addrs=email, msg=u"""Subject: The Subject of The Email \n \n 
    
Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse ornare lacus quis orci luctus porttitor. 
In sed neque vestibulum, finibus nulla vitae, tincidunt eros. 
In mi mi, consectetur ut urna et, condimentum dapibus augue. Suspendisse at aliquam justo. 
Integer dignissim metus eget ipsum imperdiet, at lobortis nisl pretium. 
In ultricies orci vitae magna dapibus finibus. Mauris tempor et dui ac ultrices. 
Sed feugiat elementum efficitur. Vivamus in pharetra dolor. Fusce placerat tortor ac aliquet congue. 
Curabitur fermentum volutpat interdum. Morbi elementum tellus sit amet massa tincidunt accumsan. 
Cras id condimentum massa, in laoreet mi. Sed vel malesuada sem, in sodales orci.
""".encode('utf-8'))
conn.quit()
