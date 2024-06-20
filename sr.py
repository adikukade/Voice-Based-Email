import speech_recognition as sr
import pyttsx3

from preplexity.final import send_email

# Initialize text-to-speech engine
engine = pyttsx3.init()

def speak(text):
    engine.say(text)
    engine.runAndWait()

def listen():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language='en-in')
            print(f"User said: {query}\n")
        except Exception as e:
            print(e)
            print("Unable to Recognize your voice.")
            return "None"
    return query

# Use the listen function to get user input
query = listen().lower()

# Perform actions based on user input, such as sending an email
if 'send email' in query:
    # Get email details from user
    recipient_name = listen()
    recipient_email = listen().lower()
    subject = listen()
    body = listen()

    # Send email
    send_email(recipient_name, recipient_email, subject, body)
    speak("Email sent!")