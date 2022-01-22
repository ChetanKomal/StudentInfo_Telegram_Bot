import os
import openpyxl as xl
import telebot
from time import sleep
my_secret = "YOUR_KEY"

API_KEY = my_secret
bot = telebot.TeleBot(API_KEY)
#  xxx= datetime.datetime.fromtimestamp(message.date) #gives the date and time on which the command is passed. have to import datetime for it to work
# @bot.message_handler(func=lambda m: True)
# def hello(message):
#     bot.send_message(message.chat.id,message.text)
def find(chetan):
 wb = xl.load_workbook("h1.xlsx")
 ws = wb.active
 a = ["0","1","rollno: ","Student Name: ","Father Name: ","Gender: ","Branch: ","Group: ","specialization: ","Mobile no: ","Email: ","Status: "]
 tmp = []
 x=chetan
 for row in range(3,259):
     if(int(ws[f"B{row}"].value) == int(x)):
         for i in range(3,12):
             y=chr(64+i)
             tmp.append(str(a[i]) + " " + str(ws[f"{y}{row}"].value))
             #print( str(a[i]) + " " + str(ws[f"{y}{row}"].value))
            
 
 return tmp           


@bot.message_handler(commands=["help","start"])
def helper(message):
  info = ''' 
         This Bot gives information about students of B.E CSE Branch Batch 2019 CUHP. Just Pass a valid rollno.
         \n Bot Commands:\n{valid_university_id}(Eg. 191198xxxx) \n /help \n /start \n /moreinfo\n  \nDeveloped by: @CK
   '''
  bot.send_message(message.chat.id,info)
@bot.message_handler(commands=["moreinfo"])
def moreinfo(message):
 ing = '''
       This bot provides the following information about students:\n
       Student Photo (if available)
       ID: X
       Student name: X
       Father name: X
       Gender: X
       Branch:X
       Group: X
       Specializaton: X
       Mobile No: X
       Email: xx@chitkarauniversity.edu.in
       Status: Dayscholar or Hosteller\n
       
       Developed By: @CK

 '''
 
 bot.send_message(message.chat.id,ing)
@bot.message_handler(regexp="191198\d\d\d\d")
def rep(message):
 
  if(find(message.text) != []):
   tmp = find(message.text)
   ff = "\n"+"\n" + "ID: " + message.text + "\n" + tmp[0] + "\n"+tmp[1]+"\n"+tmp[2]+"\n"+tmp[3]+"\n"+tmp[4]+"\n"+tmp[5]+"\n"+tmp[6]+"\n"+tmp[7]+"\n"+tmp[8]+"\n" +"\n" + "\n"+"Developed by @CK"
   pic = f"https://hp.chitkara.edu.in//Storage/Images/Student/{str(message.text)}.jpg"
   try:
    bot.send_photo(message.chat.id,photo=pic,caption=ff)  
   except:
    bot.send_message(message.chat.id,"IMAGE NOT FOUND"+ff)   
  else:
   bot.send_message(message.chat.id,"No Data Found")

@bot.message_handler(func=lambda m: True)
def echo_all(message):
	bot.send_message(message.chat.id,"Incorrect RollNumber")  

print("Bot running")
try:
    bot.polling() 
except:
    sleep(15)
    bot.polling()    
