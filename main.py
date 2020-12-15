from flask import Flask,render_template,request
import pandas as pd
from datetime import datetime
import smtplib
GMAIL_ID="execute@gmail.com"
GMAIL_PSWD="000000"
app=Flask(__name__)
@app.route("/")
def index():
    return render_template("index.html")
@app.route("/birthdaywishing",methods=['GET','POST'])
def birthdaywishing():
    if request.method == "POST":
        Name = request.form['Name']
        Birthday= request.form['Birthday']
        Dialogue = request.form['Dialogue']
        Year = request.form['Year']
        Email = request.form['Email']
        path = "data.xlsx"
        df1 = pd.read_excel(path)
        SeriesA = df1['Name']
        SeriesB = df1['Birthday']
        SeriesC = df1['Dialogue']
        SeriesD = df1['Year']
        SeriesE = df1['Email']
        A = pd.Series(Name)
        B = pd.Series(Birthday)
        C = pd.Series(Dialogue)
        D = pd.Series(Year)
        E = pd.Series(Email)
        SeriesA = SeriesA.append(A)
        SeriesB = SeriesB.append(B)
        SeriesC = SeriesC.append(C)
        SeriesD = SeriesD.append(D)
        SeriesE = SeriesE.append(E)
        df2 = pd.DataFrame({"Name": SeriesA, "Birthday": SeriesB, "Dialogue": SeriesC, "Year": SeriesD, "Email": SeriesE})
        df2.to_excel(path, index=False)
    return render_template("birthdaywishing.html")
@app.route("/annivesarywishing",methods=['GET','POST'])
def annivesarywishing():
    if request.method == "POST":
        CoupleName = request.form['CoupleName']
        Annivesary= request.form['Annivesary']
        AnnivesaryDialogue = request.form['AnnivesaryDialogue']
        AYear = request.form['AYear']
        AEmail = request.form['AEmail']
        path = "data.xlsx"
        df1 = pd.read_excel(path)
        SeriesA = df1['CoupleName']
        SeriesB = df1['Annivesary']
        SeriesC = df1['AnnivesaryDialogue']
        SeriesD = df1['AYear']
        SeriesE = df1['AEmail']
        A = pd.Series(CoupleName)
        B = pd.Series(Annivesary)
        C = pd.Series(AnnivesaryDialogue)
        D = pd.Series(AYear)
        E = pd.Series(AEmail)
        SeriesA = SeriesA.append(A)
        SeriesB = SeriesB.append(B)
        SeriesC = SeriesC.append(C)
        SeriesD = SeriesD.append(D)
        SeriesE = SeriesE.append(E)
        df2 = pd.DataFrame({"CoupleName": SeriesA, "Annivesary": SeriesB, "AnnivesaryDialogue": SeriesC, "AYear": SeriesD, "AEmail": SeriesE})
        df2.to_excel(path, index=False)
    return render_template("annivesarywishing.html")
@app.route("/addinvitation",methods=['GET','POST'])
def addinvitation():
    if request.method == "POST":
        EventName = request.form['EventName']
        Date= request.form['Date']
        Time = request.form['Time']
        Venue = request.form['Venue']
        Invitation= request.form['Invitation']
        path1 = "data1.xlsx"
        df1 = pd.read_excel(path1)
        SeriesA = df1['EventName']
        SeriesB = df1['Date']
        SeriesC = df1['Time']
        SeriesD = df1['Venue']
        SeriesE = df1['Invitation']
        A = pd.Series(EventName)
        B = pd.Series(Date)
        C = pd.Series(Time)
        D = pd.Series(Venue)
        E = pd.Series(Invitation)
        SeriesA = SeriesA.append(A)
        SeriesB = SeriesB.append(B)
        SeriesC = SeriesC.append(C)
        SeriesD = SeriesD.append(D)
        SeriesE = SeriesE.append(E)
        df2 = pd.DataFrame({"EventName": SeriesA, "Date": SeriesB, "Time": SeriesC, "Venue": SeriesD, "Invitation": SeriesE})
        df2.to_excel(path1, index=False)
    return render_template("addinvitation.html")
def sendEmail(to,sub,name,msg):
    s= smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(GMAIL_ID,GMAIL_PSWD)
    s.sendmail(GMAIL_ID,to,f"Subject:{sub+' '+name}\n\n{msg}")
    s.quit()
def sendmail(GMAIL_ID,dest, sub, name,body):
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PSWD)
    s.sendmail(GMAIL_ID, dest, f"Subject:{sub + ' ' + name}\n\n{body}")
    s.quit()
    return True;


if __name__ == '__main__':
    df=pd.read_excel("data.xlsx")
    df3 = pd.read_excel("data1.xlsx")
    df4 = pd.read_excel("data2.xlsx")
    today=datetime.now().strftime("%d-%m")
    yearNow=datetime.now().strftime("%Y")
    writeInd=[]
    writeIndx = []
    for index,item in df.iterrows():
        bday=item['Birthday'].strftime("%d-%m")
        if(today==bday)and yearNow not in str(item['Year']):
            sendEmail(item['Email'],"Happy Birthday",item['Name'],item['Dialogue'])
            writeInd.append(index)
    for index,item in df4.iterrows():
        aday=item['Annivesary'].strftime("%d-%m")
        if(today==aday)and yearNow not in str(item['AYear']):
            sendEmail(item['AEmail'],"Happy Annivesary",item['CoupleName'],item['AnnivesaryDialogue'])
            writeIndx.append(index)

    Invite = []
    ms=[]
    for index, item in df.iterrows():
        if (item['Email'] not in Invite):
                Invite.append(item['Email'])
    for index, item in df4.iterrows():
        if (item['AEmail'] not in Invite):
                Invite.append(item['AEmail'])
    for index, item in df3.iterrows():
        iday = item['Date'].strftime("%d-%m")
        if (today==iday):
            ms.extend([item['Invitation'],str(item['Date']),item['Time'],item['Venue']])
            body=(ms[0]+' '+"on"+' '+ms[1]+' '+"at"+' '+ms[2]+' '+"Venue"+' '+ms[3])
            for dest in Invite:
                #print(dest)
                sendmail(GMAIL_ID, dest,"Welcome to ",item['EventName'],body)

for i in writeInd:
    yr=df.loc[i,'Year']
    df.loc[i,'Year']=str(yr)+','+str(yearNow)
df.to_excel('data.xlsx',index=False)
for j in writeIndx:
    dm=df4.loc[j,'AYear']
    df4.loc[j,'AYear']=str(dm)+','+str(yearNow)
df4.to_excel('data2.xlsx',index=False)
app.run(debug=True, use_reloader=False)
