import speech_recognition as sr
import win32com.client
import webbrowser
import os
import random
import datetime
import sys
import wikipedia
import pywhatkit
import requests
from bs4 import BeautifulSoup
import pyautogui



speaker = win32com.client.Dispatch("SAPI.SpVoice")


def takeCommand():
    r = sr.Recognizer()

    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User Said: {query}")
            return query
        except Exception as e:
            return "some error occurred"


def alarm(query):
    deletetime = open("alarmtext.txt", "r+")
    deletetime.truncate(0)
    deletetime.close()

    timehere=open("alarmtext.txt","a")
    timehere.write(query)
    timehere.close()

    extractedtime=open("alarmtext.txt","rt")
    time=extractedtime.read()
    Time=str(time)
    extractedtime.close()

    deletetime = open("alarmtext.txt","r+")
    deletetime.truncate(0)
    deletetime.close()

    timeset = str(time)
    timenow = timeset.replace("baymax","")
    timenow = timenow.replace("set an alarm", "")
    timenow = timenow.replace(" and ", ":")
    Alarmtime = str(timenow)
    print(Alarmtime)

    alarm_datetime = datetime.datetime.strptime(Alarmtime, "%H:%M:%S")


    alarm_datetime += datetime.timedelta(seconds=30)
    while True:
        currenttime = datetime.datetime.now().strftime("%H:%M:%S")
        current_datetime = datetime.datetime.strptime(currenttime, "%H:%M:%S")

        if currenttime == Alarmtime:
            speaker.Speak("Alarm ringing")
            os.startfile("F:\Others\Suzume.mp3")

        elif current_datetime>=alarm_datetime:
            break

    return
if __name__ == '__main__':

    intro = "Hi, I am Beymax! Your desktop Assistant."
    print(intro)
    speaker.Speak(intro)
    while 1:
        print("Listening....")
        query = takeCommand()

        if "how are you" in query:
            greetings = ['I am good! What about you?', 'I am so so', 'Not in a good mood',
                         'I am pretty good! Thank you', 'I cant complain!Thanks for asking!', 'I am happy to be here']
            k = random.randint(0, 5)
            speaker.Speak(greetings[k])

        if "hey" in query:
            speaker.Speak("Yes?")

        if "thank" in query:
            speaker.Speak("My pleasure")

        if "help me" in query:
            surety = ['Of course!', 'Absolutely! How can I help you?', 'Yes! Please tell me what can I do']
            k = random.randint(0, 2)
            speaker.Speak(surety[k])

        sites = [["youtube", "http://www.youtube.com"], ["google", "http://www.google.com"],
                 ["wikipedia", "http://www.wikipedia.com"]]

        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                speaker.Speak(f"Opening {site[0]}")
                webbrowser.open(site[1])

        if "play music" in query:
            musicdir = "F:\Music"
            songs = os.listdir(musicdir)
            k = random.randint(0, 7)
            print("Now playing...")
            print(songs[k])
            os.startfile(os.path.join(musicdir, songs[k]))

        if "time" in query:
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            speaker.speak(f"The time is {strfTime}")

        if "open code" in query:
            path = "C:\Program Files\CodeBlocks\CodeBlocks.exe"
            os.startfile(path)
        if "open Adobe" in query:
            path = "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
            os.startfile(path)

        if "open Chrome" in query:
            path = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
            os.startfile(path)

        if "open Word" in query:
            path = "C:\Program Files (x86)\Microsoft Office\\root\Office16\WINWORD.exe"
            os.startfile(path)

        if "open Firefox" in query:
            path = "C:\Program Files (x86)\Mozilla Firefox\\firefox.exe"
            os.startfile(path)

        if "open vs" in query:
            path = "C:\\Users\Asus\AppData\Local\Programs\Microsoft VS Code\Code.exe"
            os.startfile(path)

        if "open Microsoft edge" in query:
            path = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
            os.startfile(path)

        if "open" in query:
            query=query.replace("open","")
            pyautogui.press("super")
            pyautogui.typewrite(query)
            pyautogui.sleep(2)
            pyautogui.press("enter")

        if "close" in query:
            query=query.replace("close ","")
            # cap=query.capitalize()
            os.system(f"taskkill /f /im {query}.exe")

        if "search Google".lower() in query.lower():
            import wikipedia as googleScrap
            query = query.replace("beymax", "")
            query = query.replace("search Google", "")
            query = query.replace("about", "")

            speaker.speak("This is what I found on google")

            try:
                pywhatkit.search(query)
                result = googleScrap.summary(query, 1)
                print(result)
                speaker.speak(result)

            except:
                speaker.speak("No speakable output")

        if "search YouTube" in query:
            speaker.speak("This is what I have found!")
            query = query.replace("beymax", "")
            query = query.replace("search Youtube", "")
            query = query.replace("about", "")

            try:
                web = "https://www.youtube.com/results?search_query=" + query;
                webbrowser.open(web)
                # pywhatkit.playonyt(query)
                speaker.speak("Done!")

            except:
                speaker.speak("No speakable output")

        if "search wikipedia" in query:
            speaker.speak("This is what I have found!")
            query = query.replace("beymax", "")
            query = query.replace("search wikipedia", "")
            query = query.replace("about", "")

            try:
                results = wikipedia.summary(query, sentences=2)
                speaker.speak("According to wikipedia..")
                print(results)
                speaker.speak(results)
            except:
                speaker.speak("No speakable output")

        if "temperature" in query:
            search = "temperature in Bangladesh"
            url = f"https://www.google.com/search?q={search}"
            ru = requests.get(url)
            data = BeautifulSoup(ru.text, "html.parser")
            temp = data.find("div", class_="BNeawe").text
            print(temp)
            speaker.speak(f"current{search} is {temp}")


        if "set an alarm" in query:
            print("Input time like this: 10 and 20 and 57")
            speaker.speak("Set the time")
            a = input("Please tell the time: ")
            alarm(a)
            

        if "you can go now" in query:
            speaker.speak("It was my pleasure to help you")
            sys.exit()
