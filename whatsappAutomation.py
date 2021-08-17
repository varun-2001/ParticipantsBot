import openpyxl
import os
import pywhatkit as kit


os.chdir('C:\\Users\\hp\\Desktop\\PsychBot')

event=openpyxl.load_workbook('Event.xlsx')

responses=event['sheetName']
index=1
number=[]
numbers=[]
message="""
Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse ornare lacus quis orci luctus porttitor. 
In sed neque vestibulum, finibus nulla vitae, tincidunt eros. 
In mi mi, consectetur ut urna et, condimentum dapibus augue. Suspendisse at aliquam justo. 
Integer dignissim metus eget ipsum imperdiet, at lobortis nisl pretium. 
In ultricies orci vitae magna dapibus finibus. Mauris tempor et dui ac ultrices. 
"""

for i in responses['F']:
    numbers.append(i.value)

numbers.remove('Whatsapp Number')

for i in numbers:
    number.append('+91'+str(int(i)))

for i in number:
    try:
        kit.sendwhatmsg_instantly(phone_no=i,message=message,wait_time=10,tab_close=True)
        print(str(index)+'.'+i)
        index+=1
    except:
        print(str(index)+'.'+str(i)+'FAILED')
        index+=1
        continue


