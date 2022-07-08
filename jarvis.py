import json
from urllib.request import urlopen
import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import winshell
import operator
import random
import wolframalpha
import subprocess
import ctypes
import os
import smtplib
import pyjokes
import ctypes
import json
import feedparser
import win32com.client as wincl
import time
from bs4 import BeautifulSoup
import win32com.client as wincl
#from ecapture import ecapture as ec

import requests


engine = pyttsx3.init('sapi5') #used for taking voices
voices = engine.getProperty('voices')
#print(voices) #bydefault we have 2, male & female
engine.setProperty('voice',voices[0].id) #seting voice property to 0th voice 

def speak(audio):
    engine.say(audio) #engine speaks the audio string
    engine.runAndWait()

def wishMe():
    speak("what should i call you?")
    uname=takeCommand()

    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<=0:
        speak("good morning !")
    elif hour>=12 and hour<18:
        speak("good afternoon")
    else:
        speak("Good Evening!")
    speak(uname)
    speak(" How may i help you?")

def takeCommand():
    #takes microphone i/p from user & returns string o/p
    r = sr.Recognizer() #recognizes audio
    with sr.Microphone() as source:  
        print("Listening...")
        r.pause_threshold = 1   #seconds of nonspeaking audio before its considered that phrase is complete  
        r.adjust_for_ambient_noise(source)
        r.energy_threshold = 100 #determines how loudly you need to speak for it to recognize
        
        audio = r.listen(source) #Records a single phrase from ``source``
        
    try:
        
        print("Recognizing...")
        query = r.recognize(audio)
        print(f"User said: {query}\n")
        

    except Exception as e:
        print(e) #this prints exception
 
        print("Please say that again..")
        return "None"
    return query

'''def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
     
    # Enable low security in gmail
    server.login('your email id', 'your email password')
    server.sendmail('your email id', to, content)
    server.close()
'''


if __name__ == '__main__':
   speak(" im speaking and i am JARVIS")
   wishMe()
   while True:
    query = takeCommand().lower()
    #logic for executing tasks based on query

    if 'wikipedia' in query :
        speak("searching wikipedia")
        query=query.replace("wikipedia", "")
        results= wikipedia.summary(query,sentences=2)
        speak("According to wikipedia")
        print(results)
        speak(results)

    elif 'open youtube' in query:
        webbrowser.open("www.youtube.com",2)
    


    elif 'open google' in query:
        webbrowser.open("www.google.com",2)

    elif 'open code' in query:
        location="C:\\Users\\DELL\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
        os.startfile(location)

    elif 'play music' in query or "play song" in query:
            speak("Here you go with music")
            # music_dir = "G:\\Song"
            music_dir = "C:\\Users\\DELL\\Music"
            songs = os.listdir(music_dir)
            print(songs)   
            random = os.startfile(os.path.join(music_dir, songs[1]))
    
    #elif 'send a mail' in query:
    #        try:
    #            speak("What should I say?")
    #            content = takeCommand()
    #            speak("whom should i send")
    #            to = input()   
    #            sendEmail(to, content)
    #            speak("Email has been sent !")
    #        except Exception as e:
    #            print(e)
    #           speak("I am not able to send this email")


    elif "write a note" in query:
            speak("What should i write, sir")
            note = takeCommand()
            file = open('jarvis.txt', 'w')
            speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("% H:% M:% S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)
         
    elif "show note" in query:
            speak("Showing Notes")
            file = open("jarvis.txt", "r")
            print(file.read())
            speak(file.read(6))


    elif "weather" in query:
             
            # Google Open weather website
            # to get API of Open weather
            api_key = "Api key"
            base_url = "http://api.openweathermap.org / data / 2.5 / weather?"
            speak(" What is the City name? ")
            print("City name : ")
            city_name = takeCommand()
            complete_url = base_url + "appid =" + api_key + "&q =" + city_name
            response = requests.get(complete_url)
            x = response.json()
             
            if x["cod"] != "404":
                y = x["main"]
                current_temperature = y["temp"]
                current_pressure = y["pressure"]
                current_humidiy = y["humidity"]
                z = x["weather"]
                weather_description = z[0]["description"]
                print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description))
             
            else:
                speak("city not found")


    elif 'time now' in query:
        strTime  = datetime.datetime.now().strftime("%H:%M:%S")
        speak(f"the time is {strTime}\n")

    elif 'joke' in query:
            speak(pyjokes.get_joke())
             
    
 
    #elif "calculate" in query:
      #    app_id = "Wolframalpha api id"
      #      indx = query.lower().split().index('calculate')
      #      query = query.split()[indx + 1:]
      #     res = client.query(' '.join(query))
      #     answer = next(res.results).text
      #     print("The answer is " + answer)
      #     speak("The answer is " + answer)
      #  
     #
     #
     #
     #
     #
     #



    elif 'change background' in query:
            
       ctypes.windll.user32.SystemParametersInfoW(20,0,"Location of wallpaper",0)
       speak("Background changed successfully")
 
    
    
    elif 'exit' in query:
        speak("Thanks for giving me your time")
        exit()
    
    elif 'search' in query or 'play' in query:
             
            query = query.replace("search", "")
            query = query.replace("play", "")         
            webbrowser.open(query)
 
    elif 'news' in query:
             
            try:
                jsonObj = urlopen('''https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =\\times of India Api key\\''')
                data = json.load(jsonObj)
                i = 1
                 
                speak('here are some top news from the times of india')
                print('''=============== TIMES OF INDIA ============'''+ '\n')
                 
                for item in data['articles']:
                     
                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str(i) + '. ' + item['title'] + '\n')
                    i += 1
            except Exception as e:
                 
                print(str(e))
    
    elif 'lock window' in query:
                speak("locking the device")
                ctypes.windll.user32.LockWorkStation()
 
    elif 'shutdown system' in query:
                speak("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call('shutdown / p /f')
                 
    elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
            speak("Recycle Bin Recycled")
 
    elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop jarvis from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)
 
    elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.com / maps / place/" + location + "")
 
    #elif "camera" in query or "take a photo" in query:
    #        ec.capture(0, "Jarvis Camera ", "img.jpg")
 
    elif "restart" in query:
            subprocess.call(["shutdown", "/r"])
             
    elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown / h")
 
    elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])


    elif "who am i" in query:
            speak("If you talk then definitely your human.")
 
    elif "who are you" in query:
            speak("I am a virtual assistant created by kumuda")
    

 