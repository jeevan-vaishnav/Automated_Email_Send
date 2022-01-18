import pandas as pd  # red for files
import datetime  # for date and time
import smtplib
# for task schedular testing
import os
os.chdir(r"E:\Python Learning\PythonPratice\BIRTHDAY_WISHSES")
# os.mkdir("testing")

# Enter your details
GMAIL_ID = ''
GMAIL_PWD = ''


def sendEmail(to, sub, msg):
    print(f"Email to {to} sent with  Subject : {sub} and message {msg}")
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PWD)
    s.sendmail(GMAIL_ID, to, f"Subject: {sub}\n\n{msg}")
    s.quit()


# main function
if __name__ == "__main__":

    # # for testing
    # sendEmail(GMAIL_ID, "Subject", "Test")  # code1
    # exit()  # code1

    # pip install xlrd and pip install openpyxl
    df = pd.read_excel("data.xlsx", engine='openpyxl')
    # print(df)
    today = datetime.datetime.now().strftime("%d-%m")
    #  it date  is a string form
    # print(today)
    yearNow = datetime.datetime.now().strftime("%Y")
    writeInd = []
    for index, item in df.iterrows():
        # print(index, item['Birthday'])
        bday = item['Birthday'].strftime("%d-%m")
        to = item['Email']
        msg = item['Dialoge']
        # print(bday)
        if(today == bday) and yearNow not in str(item['Year']):
            sendEmail(to, "Happy Birtdhay", msg)
            writeInd.append(index)

    # print(writeInd)
    l = len(writeInd)
    if l > 0:
        # if l greater then 0 then will execute the block
        for i in writeInd:
            yr = df.loc[i, 'Year']
            df.loc[i, 'Year'] = str(yr) + "," + str(yearNow)  # set
            # print(df.loc[i, 'Year'])
        df.to_excel('data.xlsx', index=False)
