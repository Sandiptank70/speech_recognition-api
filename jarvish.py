import pyttsx3 #pip install pyttsx3
import speech_recognition as sr #pip install speechRecognition
import datetime
import wikipedia #pip install wikipedia
import webbrowser
import os
import googlesearch #pip install googlesearch
import random
from time import sleep
import win32com.client as win32
import time


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[1].id)
engine.setProperty('voice', voices[1].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon!")

    else:
        speak("Good Evening!")

    speak("I am Jarvis Sir. Please tell me how may I help you")

def takeCommand():
    #It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        r.energy_threshold=150
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        # print(e)
        print("Say that again please...")
        return "None"
    return query


RANGE = range(3, 8)
def word():

        word = win32.gencache.EnsureDispatch('Word.Application')
        doc = word.Documents.Add()
        word.Visible = True
        sleep(1)

        rng = doc.Range(0, 0)
        rng.InsertAfter('')
        sleep(1)
        for i in RANGE:
            rng.InsertAfter('')
            sleep(1)
        rng.InsertAfter("")


def excel():
    """"""
    xl = win32.gencache.EnsureDispatch('Excel.Application')
    ss = xl.Workbooks.Add()
    sh = ss.ActiveSheet

    xl.Visible = True
    time.sleep(1)

    sh.Cells(1, 1).Value = ''

    time.sleep(1)
    for i in range(2, 8):

        time.sleep(1)

    # ss.Close(False)
    # xl.Application.Quit()

if __name__ == "__main__":
    wishMe()
    while True:
    # if 1:
        query = takeCommand().lower()

        # Logic for executing tasks based on query
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=20)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'in google' in query:
             try:
                 from googlesearch import search
             except:
                 print("No module named 'google' found")
             speak('Searching google...')
             query = query.replace("in google", "")
             for i in search(query, tld="co.in", num=10, stop=1, pause=2):
                 speak("According to google")
                 webbrowser.open(i)

        elif 'youtube' in query:
            webbrowser.open("youtube.com")
        elif 'open google' in query:
            webbrowser.open("google.com")
        elif 'whatsapp' in query:
            webbrowser.open('whatsapp.com')
        elif 'facebook' in query:
            webbrowser.open('facebook.com')
        elif 'instagram' in query:
            webbrowser.open('instagaram.com')
        elif 'music' in query:
            music_dir="D:\\music"
            songs=os.listdir(music_dir)
            a=len(songs)
            b=random.uniform(0,a-1)
            os.startfile(os.path.join(music_dir,songs[int(b)]))
        elif 'editor' in query:
             codepath="C:\\Users\\radhe\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
             os.startfile(codepath)
        elif 'open computer' in query:
             codepath = "C:\\Users\\radhe"
             os.startfile(codepath)

        elif 'open pycharm' in query:
            codepath="C:\\Program Files\\JetBrains\\PyCharm Community Edition 2019.1.3\\bin\\pycharm64.exe"
            os.startfile(codepath)
        elif 'sleep' in query:
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
        elif 'shutdown' in query or 'computer' in query:
            os.system("shutdown /s /t 1")
        elif 'open chrome' in query:
            codepath="C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"
            os.startfile(codepath)
        elif 'close chrome' in query:
            os.system("taskkill /im chrome.exe /f")
        elif 'open word' in query:
            word()
        elif 'open excel' in query:
            excel()








