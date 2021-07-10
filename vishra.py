'''
import statement is used to import the packages 
into the python file. We are going to use all the
imported files in our project.  
'''

import subprocess 
import wolframalpha 
import pyttsx3 
import datetime
import tkinter 
import json 
import random 
import operator 
import speech_recognition as sr 
import datetime 
import wikipedia 
import webbrowser 
import os 
import winshell 
import pyjokes 
import feedparser 
import smtplib 
import ctypes 
import time 
import requests 
import shutil 
from twilio.rest import Client 
from clint.textui import progress 
from ecapture import ecapture as ec 
from bs4 import BeautifulSoup 
import win32com.client as wincl 
from urllib.request import urlopen 
import google
from bs4 import BeautifulSoup
import urllib.request

engine = pyttsx3.init('sapi5') 
voices = engine.getProperty('voices') 
# print(voices[1].id)  #prints the voice info and name of the voice (zira)
engine.setProperty('voice', voices[1].id) 

def speak(audio): 
    print('Assistant: ' +audio)
    engine.say(audio) 
    engine.runAndWait() 

def speaknoprint(audio): 
    engine.say(audio) 
    engine.runAndWait() 

def wishMe(): 
    hour = int(datetime.datetime.now().hour) 
    if hour>= 0 and hour<12: 
        speak("Good Morning Mam!")
    elif hour>= 12 and hour<18: 
        speak("Good Afternoon Mam!")  
    else: 
        speak("Good Evening Mam!")  
    assname =("Vishra") 
    speak("I am Vishra, a virtual assistant. I am at your service!") 
      
def usrname(): 
    speak("What should i call you?") 
    uname = takeCommand() 
    columns = shutil.get_terminal_size().columns 
    print("-----------------------------------------------------------------------".center(columns)) 
    print('Welcome '.center(columns)) 
    print(uname.center(columns))
    print("-----------------------------------------------------------------------".center(columns)) 
    speaknoprint('Welcome '+uname)
    speak("How can i help you?") 
 
def takeCommand(): 
    r = sr.Recognizer() 
    with sr.Microphone() as source: 
        print("Listening...") 
        r.energy_threshold = 600
        r.pause_threshold = 1
        audio = r.listen(source) 
    try: 
        print("Recognizing...")     
        query = r.recognize_google(audio, language ='en-in') 
        print(f"User said: {query}") 
    except sr.UnknownValueError:
        speak('Sorry  didn\'t get that! Try typing the command')
        query = str(input('Command: '))
    return query 

def sendEmail(to, content): 
    server = smtplib.SMTP('smtp.gmail.com', 587) 
    server.ehlo() 
    server.starttls() 
    # Enable low security in gmail 
    server.login('sender_email', 'sender_pw') 
    server.sendmail('reciever_email', to, content) 
    server.close() 

if __name__ == '__main__': 
    clear = lambda: os.system('cls') 
    # This Function will clean any command before execution of this python file 
    clear() 
    wishMe() 
    speak('Do you want me to register your name in my database now?') 
    query = takeCommand().lower()
    if 'yes' in query or 'ok' in query or 'sure' in query:
        usrname()
    while True: 
        query = takeCommand().lower() 
        # All the commands said by user will be stored here in 'query' and will be converted to lower case for easily recognition of command 
        if 'wikipedia' in query: 
            speak('Searching Wikipedia...') 
            query = query.replace("wikipedia", "") 
            results = wikipedia.summary(query, sentences = 3) 
            speak("According to Wikipedia") 
            print(results) 
            speak(results)

        elif 'open youtube' in query: 
            speak("Here you go to Youtube\n") 
            webbrowser.open("youtube.com") 

        elif 'open google' in query: 
            speak("Here you go to Google\n") 
            webbrowser.open("google.com") 

        elif 'Trace your response' in query or 'trace' in query:
             statment = takeCommand()
             for j in search(statment, tld="co.in", num=10, stop=10, pause=2): 
                print(j)  

        elif 'open stack overflow' in query: 
            speak("Here you go to Stack Over flow.Happy coding") 
            webbrowser.open("stackoverflow.com")    
            
        elif 'play music' in query or "play song" in query: 
            speak("Here you go with music") 
            music_dir = "Music_Directory_Path"
            songs = os.listdir(music_dir) 
            print(songs)     
            random = os.startfile(os.path.join(music_dir, songs[1])) 

        elif 'question mark' in query:
            speak('This is a question mark symbol - ?')
            print('This is a question mark symbol - ?')

        elif 'open zoom' in query: 
            codePath = r"Zoom_Path"
            os.startfile(codePath) 

        elif 'email' in query: 
            try: 
                speak("What should I say?") 
                content = takeCommand() 
                to = "Receiver email address"    
                sendEmail(to, content) 
                speak("Email has been sent !") 
            except Exception as e: 
                print(e) 
                speak("I am not able to send this email") 

        elif 'send a mail' in query: 
            try: 
                speak("What should I say?") 
                content = takeCommand() 
                speak("whome should i send") 
                to = input()     
                sendEmail(to, content) 
                speak("Email has been sent !") 
            except Exception as e: 
                print(e) 
                speak("I am not able to send this email") 

        elif 'how are you' in query: 
            speak("I am fine, Thank you") 
            speak("How are you") 
  
        elif 'fine' in query or "good" in query: 
            speak("It's good to know that your fine") 
    
        elif "change my name to" in query: 
            query = query.replace("change my name to", "") 
            assname = query 
  
        elif "change name" in query: 
            speak("What would you like to call me ") 
            assname = takeCommand() 
            speak("Thanks for naming me") 

        elif "what's your name" in query or "What is your name" in query: 
            speak("My friends call me") 
            speak(assname) 
            print("My friends call me", assname) 

        elif 'exit' in query: 
            speak("Thanks for giving me your time") 
            exit() 

        elif "who made you" in query or "who created you" in query:  
            speak("I have been created by Pavithra.") 

        elif 'joke' in query: 
            speak(pyjokes.get_joke()) 

        elif "calculate" in query:  
            app_id = "5VY3V5-WRK55P7PY3" 
            client = wolframalpha.Client(app_id) 
            indx = query.lower().split().index('calculate')  
            query = query.split()[indx + 1:]  
            res = client.query(' '.join(query))  
            answer = next(res.results).text 
            print("The answer is " + answer)  
            speak("The answer is " + answer)  

        elif 'search' in query or 'play' in query: 
            query = query.replace("search", "")  
            query = query.replace("play", "")           
            webbrowser.open(query)  
  
        elif "who am i" in query: 
            speak("If you talk then definately your human.") 

        elif "why did you come to world" in query: 
            speak("Thanks to my creators. i have been built to assist and help people with computation and data assistance") 

        elif 'powerpoint presentation' in query or 'open presentation' in query: 
            speak("opening Power Point presentation") 
            power = r"Ppt_path"#give your own loacation 
            os.startfile(power) 

        elif 'how old are you' in query or 'what is your age' in query:
            speak("I am almost new born! I have completed 2 years.")

        elif 'i like you' in query:
            speak("Thats great! I like you too...")

        elif 'thank' in query:
            speak("No problem! Any time")

        elif "what are you doing" in query or 'what\'s going on' in query:
            stMsgs = ['Just doing my thing!', 'Time pass! What are you doing?','I am nice and full of energy for your querries!']
            speak(random.choice(stMsgs))

        elif "what's up" in query or 'how are you' in query:
            stMsgs = ['Just doing my thing!', 'I am fine! Thanks for asking.','Nice! How are you?', 'I am nice and full of energy!']
            speak(random.choice(stMsgs))

        elif 'you are good' in query or 'your good' in query:
            speak("Oh No! Dont praise me that much")

        elif 'open google' in query:
            speak("Opening Google!")
            webbrowser.open("google.com")

        elif 'open gmail' in query:
            speak("Opening G-mail!")
            webbrowser.open('www.gmail.com')

        elif 'what is love' in query: 
            speak("It is 7th sense that destroy all other senses") 

        elif "who are you" in query: 
            speak("Hi, I'm Vishra, the Virtual Assistant. Speed 1 terahertz, memory 1 zigabyte.")   

        elif 'reason for your existence' in query: 
            speak("I was created as a Minor project by Pavithra ") 

        elif 'change background' in query: 
            ctypes.windll.user32.SystemParametersInfoW(20, 30, 'Wallpapers_Path', 30) 
            speak("Background changed succesfully") 

        elif 'open bluestack' in query: 
            appli = r""#give your own location
            os.startfile(appli) 

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Mam,right now the time is {strTime}") 

        elif 'score' in query:
            score_page = 'XML_FILE'
            page = urllib.request.urlopen(score_page)
            soup = BeautifulSoup(page, 'html.parser')
            result = soup.find_all('description') 
            ls=[]
            for match in result:
                ls.append(match.get_text())
            speak(ls)

        elif 'news' in query: 
            try:  
                jsonObj = urlopen('''API_KEY''') 
                data = json.load(jsonObj) 
                i = 1
                speak('here are top 10 Headlines from BBC News') 
                print('''=============== BBC News ============'''+ '\n') 
                for item in data['articles']: 
                    print(str(i) + '. ' + item['title'] + '\n') 
                    print(item['description'] + '\n') 
                    speaknoprint(str(i) + '. ' + item['title'] + '\n') 
                    i += 1
            except Exception as e: 
                print(str(e)) 

        elif 'lock window' in query: 
                speak("locking the device") 
                ctypes.windll.user32.LockWorkStation() 
 
        elif 'shutdown system' in query or 'shutdown' in query: 
                speak("Hold On a Sec ! Your system is on its way to shut down") 
                subprocess.call('shutdown / p /f') 
                 
        elif 'empty recycle bin' in query: 
            speak("Recycling your bin...") 
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True) 
            speak("Recycle Bin Recycled") 

        elif "don't listen" in query or "stop listening" in query or "terminate" in query: 
            speak("for how much time you want to stop vishra from listening commands") 
            a = int(takeCommand()) 
            time.sleep(a) 
            print(a) 
 
        elif "where is" in query: 
            query = query.replace("where is", "") 
            location = query 
            speak("User asked to Locate") 
            speak(location) 
            webbrowser.open("https://www.google.com/maps/place/"+location+"") 
 
        elif "camera" in query or "take a photo" in query: 
            ec.capture(0, "vishra Camera ", "img.jpg") 

        elif "restart" in query: 
            subprocess.call(["shutdown", "/r"]) 
             
        elif "hibernate" in query or "sleep" in query: 
            speak("Hibernating") 
            subprocess.call("shutdown / h") 

        elif "log off" in query or "sign out" in query: 
            speak("Make sure all the application are closed before sign-out") 
            time.sleep(5) 
            subprocess.call(["shutdown", "/l"]) 
 
        elif "write a note" in query or "make a note" in query: 
            speak("What should i write") 
            note = takeCommand() 
            file = open('vishranote.txt', 'w') 
            speak(" Should i include date and time") 
            snfm = takeCommand() 
            if 'yes' in snfm or 'sure' in snfm or 'ok' in snfm: 
                strTime = datetime.datetime.now().strftime("%H:%M:%S") 
                file.write(strTime) 
                file.write(" :- ") 
                file.write(note) 
            else: 
                file.write(note) 
         
        elif "show note" in query: 
            speak("Showing Notes") 
            file = open("vishranote.txt", "r")  
            print(file.read()) 
            speak(file.read(6)) 
  
        elif "update assistant" in query: 
            speak("After downloading file please replace this file with the downloaded one") 
            url = '# url after uploading file'
            r = requests.get(url, stream = True) 
            with open("Voice.py", "wb") as Pypdf:                 
                total_length = int(r.headers.get('content-length')) 
                for ch in progress.bar(r.iter_content(chunk_size = 2391975), 
                                       expected_size =(total_length / 1024) + 1): 
                    if ch: 
                      Pypdf.write(ch) 
        # NPPR9-FWDCX-D2C8J-H872K-2YT43 

        elif "vishra" in query: 
            wishMe() 
            speak("vishra at your service") 
            speak(assname) 

        elif "weather" in query: 
            # Google Open weather website to get API of Open weather  
            api_key = "a7fc04335e8d8dbc0f9fac1ff2e727dc" 
            base_url = "http://api.openweathermap.org/data/2.5/weather?"
            speak(" City name ") 
            print("City name : ") 
            city_name = takeCommand() 
            complete_url = base_url+"appid="+api_key+"&q="+city_name 
            response = requests.get(complete_url)  
            x = response.json()  
            if x["cod"] != "404":  
                 y = x["main"]  
                 current_temperature = y["temp"]  
                 current_pressure = y["pressure"]  
                 current_humidiy = y["humidity"]  
                 z = x["weather"]  
                 weather_description = z[0]["description"]  
                 speak(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description))  
                 print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description))
            else:  
                  speak(" City Not Found ")
              

        elif " send a message  " in query or 'text' in query: 
                # You need to create an account on Twilio to use this service 
                account_sid = 'AC6332b5454c63272aee510c96248f0639'
                auth_token = 'ab2c6f0c07f3ef314182e71296bd15c7'
                client = Client(account_sid, auth_token) 
                speak("please give the body of the message and enter the recipiant number")
                print("please give the body of the message and enter the recipiant number")
                message = client.messages.create( body = takeCommand(),from_='+12054174830', to =input()) 
                print(message.sid) 

        elif "open wikipedia" in query: 
            webbrowser.open("wikipedia.com") 

        elif "Good Morning" in query: 
            speak("A warm" +query) 
            speak("How are you ") 
            speak(assname) 

        elif 'bye bye' in query or 'get lost' in query or 'stop' in query or 'quit' in query or 'exit' in query:
            hour = int(datetime.datetime.now().hour)
            if hour >= 2 and hour < 18:
                 speak("Have a nice day! Hope  to see you soon!")
            else:
                 speak("Good night . Sweet dreams.")
            speak("Terminating program vishra Signing off!..")
            quit()

        # most asked question from google Assistant 
        elif "will you be my gf" in query or "will you be my bf" in query:    
            speak("I'm not sure if my creators would like it, may be you should give me some time") 

        elif "how are you" in query: 
            speak("I'm fine, glad you ask me that") 

        elif "i love you" in query: 
            speak("It's hard to understand") 

        elif "what is" in query or "who is" in query: 
            client = wolframalpha.Client("5VY3V5-WRK55P7PY3") 
            res = client.query(query) 
            try: 
                print (next(res.results).text) 
                speak (next(res.results).text) 
            except StopIteration: 
                print ("No results")
        else:
            query = query
            try:
                try:
                    res = client.query(query)
                    results = next(res.results).text
                    speak('Searching WOLFRAM-ALPHA...')
                    speak('According to WOLFRAM-ALPHA...')
                    speak(results)
                except:
                    results = wikipedia.summary(query, sentences=2)
                    speak('Searching Wikipedia...')
                    speak('According to Wikipedia...')
                    speak(results)
            except:
                base_url = "https://www.google.com/search?q="
                cmplt_url = base_url+query
                webbrowser.open(cmplt_url)

        anMsgs = ['Is there anything-else i can help you with Mam!?', 'I am glad i could help you! So! Anything-else?.', 'Do you need anything-else?',
                  'Can i help you with anything-else?', 'Shooot at me if u have to know more?', 'If nothing more ! Say bye-bye !', 'You want me to search anything more?']
        speak(random.choice(anMsgs))