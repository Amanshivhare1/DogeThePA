import pyttsx3 #pip install pyttsx3
import speech_recognition as sr #pip install speechRecognition
import datetime # for date and time
import wikipedia #pip install wikipedia
import webbrowser
import random
import os
import sys
import re
import smtplib
import requests
import subprocess
import youtube_dl
import urllib
#import urllib2
import urllib3
import json
#from urllib3 import urlopen
from pyowm import OWM

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

'''def speak(text):
    engine.say(text)
    engine.runAndWait()'''


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    # It wishes the user according to the time 
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Hey its morning, why you disturbed me. By the way ")

    elif hour>=12 and hour<18:
        speak("Hey its  Afternoon! I am taking rest. By The way ")   

    else:
        speak("I am free now, tell me ") 
         
    speak("How may i help you?")       

def takeCommand():
    #It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1.5
        r.energy_threshold = 200
        audio = r.listen(source)

    try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language='en-in')
        print(f"User: {query}\n")

    except Exception as e:    
        print("Say it again please...")  
        return "None"
    return query

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    #have to make change here
    server.login('youremail@gmail.com', 'your-password')
    server.sendmail('youremail@gmail.com', to, content)
    server.close()

if __name__ == "__main__":
    speak("Initializing Doge..")
    wishMe()
    while True:
        query = takeCommand().lower()

        # Logic for executing tasks based on query
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)
        elif 'open' in query:
            reg_ex = re.search('open (.+)', query)
        if reg_ex:
            domain = reg_ex.group(1)
            print(domain)
            url = 'https://www.' + domain
            webbrowser.open(url)
        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speak(f"Now it's {strTime}")
        elif 'launch' in query:
            reg_ex = re.search('launch (.*)', query)
        if reg_ex:
            appname = reg_ex.group(1)
            appname1 = appname+".app"
            subprocess.Popen(["open", "-n", "/Applications/" + appname1], stdout=subprocess.PIPE)

        elif 'email to harry' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "harryyourEmail@gmail.com"    
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Damn this is not me, send it again. ")    
