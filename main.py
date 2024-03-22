import speech_recognition as sr
import win32com.client
import webbrowser
import os
import random


speaker=win32com.client.Dispatch("SAPI.SpVoice")
# voices = speaker.GetVoices()
# desired_voice_index=1
# speaker.Voice = voices.Item(desired_voice_index)


def takeCommand():
    r = sr.Recognizer()

    with sr.Microphone() as source:
        r.pause_threshold =1
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User Said: {query}")
            return query
        except Exception as e:
            return "some error occurred"

if __name__=='__main__':
    intro = "Hi, I am Beymax. Your desktop Assistant."
    print(intro)
    speaker.Speak(intro)
    while 1:
        print("Listening....")
        query = takeCommand()

        if "how are you" in query:
            greetings=['I am good! What about you?', 'I am so so', 'Not in a good mood', 'I am pretty good! Thank you', 'I cant complain!Thanks for asking!','I am happy to be here']
            k=random.randint(0,5)
            speaker.Speak(greetings[k])

        if "hey baymax" in query:
            speaker.Speak("Yes?")
