import threading

import speech_recognition as sr
import os
import win32com.client
import webbrowser
import time
from plyer import notification

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def takeCommand():
    r = sr.Recognizer()
    try:
        with sr.Microphone() as source:
            #r.pause_threshold = 0.4
            audio = r.listen(source)
            query = r.recognize_google(audio, language = "En-in" )
            print(f"User said:  {query}")
            return query
    except:
        return "Some Error Occured"

def set_reminder():
    count = 0
    speaker.Speak("What should i remind you about?")
    text3 = takeCommand()
    speaker.Speak("How frequently do you want to be reminded. Hours, Minutes or Seconds")
    ff = takeCommand()

    if "minutes".lower() in ff.lower():
        speaker.Speak("Tell me how many Minutes")
        time1 = takeCommand()
        timeF = int(time1)
        timeF = timeF * 60
        count = count +1

    elif "seconds".lower() in ff.lower():
        speaker.Speak("Tell me how many Seconds")
        time1 = takeCommand()
        index_of_character = time1.index(' ')
        timeF = int(time1[:index_of_character])
        count = count+1


    elif "hours".lower() in ff.lower():
        speaker.Speak("Tell me how many Hours")
        time1 = takeCommand()
        timeF = int(time1)
        timeF = timeF * 60 * 60
        count = count+1

    else:
        speaker.Speak("Couldnt Undderstand")

    if count == 1:
        if __name__ == "__main__":
            while True:
                notification.notify(
                    title="***REMINDER***",
                    message=text3,
                    app_icon="C:\\Users\\sansh\\Downloads\\to-do-list.ico",
                    timeout=5
                )
                time.sleep(timeF)


speaker.Speak("Hello i am Dexter, your virtual assistant")
speaker.Speak("How can i help you Today?")


while True:
    print("Listening...")
    text = takeCommand()
    # todo: Add more sites
    sites = [["youtube","https://www.youtube.com"], ["google", "https://www.google.com"], ["mail","https://mail.google.com/mail/u/0/#inbox"],
             ["GitHub", "https://github.com/"], ["wikipedia", "https://wikipedia.com"]]

    for site in sites:
        if f"Open {site[0]}".lower() in text.lower():
            speaker.Speak(f"Opening {site[0]}")
            webbrowser.open(site[1])

    if "set a reminder".lower() in text.lower():
        threading.Thread(target=set_reminder).start()

    if "open website".lower() in text.lower():
        speaker.Speak(f"Just mention the websites name")
        text2 = takeCommand()
        website = "https://"+text2+".com/"
        webbrowser.open(website)

    if "open local music".lower() in text.lower():
        speaker.Speak("Opening Music")
        music = "C:\\Users\\sansh\\Music"
        os.startfile(music)

    if "open pictures".lower() in text.lower():
        speaker.Speak("Opening Pictures")
        pictures = "C:\\Users\\sansh\\OneDrive\\Pictures"
        os.startfile(pictures)

    if "open onenote".lower() in text.lower():
        speaker.Speak("Opening oneNote")
        note = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\OneNote.lnk"
        os.startfile(note)

    if "open spotify".lower() in text.lower():
        speaker.Speak("Opening Spotify")
        spot = "C:\\Users\\sansh\\OneDrive\\Desktop\\Spotify - Shortcut.lnk"
        os.startfile(spot)

    if "How are you ".lower() in text.lower():
        speaker.Speak("I am fine Sir. Hope you are doing good too")

    if "sleep now".lower() in text.lower():
        speaker.Speak("Going to sleep mode. Have a good day sir")
        exit(1)







