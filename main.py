import speech_recognition as sr
import pyttsx3
import datetime
import webbrowser
from pptx import Presentation
import pyjokes
import smtplib

# initailizing the pyttsx3
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)


# function to take speak that is taking the audio as the input
def speak(audio):
    engine.say(audio)
    engine.runAndWait()


# function that greets you when you first open the terminal
def wishme():
    # this is going to return the present hour and this is in the int  so it will return the whole number
    assname = "Siri"
    speak("I am your Assistant")
    speak(assname)
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good morning Ma'am")
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon Ma'am")
    else:
        speak("Good evening Ma'am")


# function for the username
def username():
    speak('What  do I call you')
    uname = takecommand()
    speak('welcome Miss ')
    speak(uname)


# function to take the command
def takecommand():
    # starting the recognizer
    s = sr.Recognizer()
    # Micorphone take the input from the user
    # and the recognizer recognizes the voices
    with sr.Microphone() as source:
        print('Listening.....')
        s.pause_threshold = 1
        audio = s.listen(source)
    try:
        print("Recognizing...")
        # this will return the string
        query = s.recognize_google(audio, language="en-in")
        print(f"user said {query}\n")
    except Exception as e:
        print(e)
        print("Unable to recognize the voice ")
        return "None"
    return query


# function to send the email to someone with our email address
def sendemail(to ,content):
    server = smtplib.SMTP("smtp.gmail.com", 534)
    server.ehlo()
    server.starttls()
    server.login("krupalitandel190180107072@gmail.com", "KRU@tandel1234")
    server.sendmail("krupalitandel190180107072@gmail.com", to, content)
    server.close()


if __name__ == '__main__':
    wishme()
    # username()
    # taking the input from the user continusly
    while (True):
        query = takecommand().lower()

        if 'open youtube' in query:
            speak("Opening Youtube\n")
            webbrowser.open("youtube.com")
        elif 'open google' in query:
            speak("Opening Google\n")
            webbrowser.open("google.com")
        # elif 'take a photo' or 'camera' in query:
        #     ec.capture(0,"Jarvis camera","img.jpg")
        elif 'take me to stack overflow' in query:
            speak("opening Stack over flow Ma'am . Happy coding")
            webbrowser.open("stackoverflow.com")
        elif 'exit' in query:
            speak("Thanks for giving me your time")
            exit()

        elif 'siri' in query:
            speak("Siri in your service Ma'am")
        # playing the music
        # elif 'play music' or 'play song' in query:
        #     speak("Enjoy the music Ma'am")

        elif 'joke' or 'jokes' in query:
            speak(pyjokes.get_joke())

        # to send the email to someone
        elif "send email" in query:
            try:
                speak("whom to send the mail")
                to = input()
                speak("what should I say ?")
                content = takecommand()
                sendemail(to, content)
                speak('Email has been send successfully')
            except Exception as e:
                print(e)
                speak('I am unable to send this mail')
        elif "powerpoint presentation" in query:
            speak("Opening the power point presentation Ma'am")
            ppt = Presentation()

        elif "who made you" or "who created you" in query:
            speak("I have been created by Krupaali")
        elif "can you hear me" or  "am i audible":
            speak("Yes Ma'am, you are audible")
