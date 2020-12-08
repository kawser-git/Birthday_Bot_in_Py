import pandas as pd
import datetime
import smtplib
import os

#change Dir for sheduler
#os.chdir()

#enter your email and password 
GMAIL_ID = 'manish6pal1999@gmail.com'
GMAIL_PSWD = 'manish681999'
def sendMail(to,sub,msg,name):

    s = smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(GMAIL_ID,GMAIL_PSWD)

    s.sendmail(GMAIL_ID,to,f"Subject: {sub}\n\n To {name} \n{msg} \n FROM \n MANISH PAL")

    s.quit()

if __name__ == "__main__":
    
    #Read Data Excel
    df = pd.read_excel("Data.xlsx")
    
    today = datetime.datetime.now().strftime('%d-%m')
    yearNow = datetime.datetime.now().strftime('%Y')
    
    writeInd=[]
    for index, item in df.iterrows():
        bday = item['Birthday'].strftime('%d-%m')
        mailid = item['Email']
        dailog = item['Dailog']
        name = item['Name']
       
        if((today == bday) and (yearNow not in str(item['Year']))):
            sendMail(mailid,"Happy Birthday",dailog,name)
            writeInd.append(index)
    
    # update the year column to send the mail single time
        for i in writeInd:
            yr = df.loc[i,'Year']
            df.loc[i,'Year'] = f"{yr},{yearNow}"
            
        
        df.to_excel('data.xlsx',index=False)  #Save the Data excel file