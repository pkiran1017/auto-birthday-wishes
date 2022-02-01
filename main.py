import pandas as pd
import datetime
import smtplib
import os

os.chdir(r"E:\Python Project\birthdaywishes")

GMAIL_ID='mal_should_be_here'
GMAIL_PSWD='password_should_be_here.'

def sendEmail(to, sub, msg):
    s = smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(GMAIL_ID,GMAIL_PSWD)
    s.sendmail(GMAIL_ID,to,f"Subject:- {sub}\n\n{msg}")
    s.quit()

if __name__ == "__main__":
    df = pd.read_excel("data.xlsx")

    yearNow = datetime.datetime.now().strftime("%Y")
    today = datetime.datetime.now().strftime("%d-%m")

    writeInd = []
    for index, item in df.iterrows():
        bday = item["Birthday"].strftime("%d-%m")
        if (today == bday) and yearNow not in str(item["Year"]):
            sendEmail(item["Email"], "Birthday Wishes", item["Dialogue"])
            writeInd.append(index)

    print(writeInd)
    for i in writeInd:
        yr = df.loc[i, 'Year']
        df.loc[i, 'Year'] = str(yr) + ',' + str(yearNow)
    df.to_excel('data.xlsx', index=False)