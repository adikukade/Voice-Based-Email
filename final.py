from flask import Flask, render_template, request
import smtplib
import speech_recognition as sr
import pyttsx3
import pandas as pd
from openpyxl import load_workbook
import openpyxl 
app = Flask(__name__)

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

    # Render the templates

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/send_email', methods=['POST'])
def send_email():
    recipient_name = request.form['recipient_name']
    recipient_email = request.form['recipient_email']
    subject = request.form['subject']
    body = request.form['body']

    # Load existing data if the file exists
    try:
        df = pd.read_excel('receivers.xlsx')
    except FileNotFoundError:
        df = pd.DataFrame()

    # Append the new data to the DataFrame
    new_data = pd.DataFrame({'Name': [recipient_name], 'Email': [recipient_email], 'Subject': [subject]})
    df = pd.concat([df, new_data], ignore_index=True)

    # Save the DataFrame to Excel
    df.to_excel('receivers.xlsx', index=False, engine='openpyxl')

    sender_email = "avidhule7752@gmail.com"
    sender_password = "btxr cdxb vpmi oyol"

    message = f"Subject: {subject}\n\n{body}"

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(sender_email, sender_password)
    server.sendmail(sender_email, recipient_email, message)
    server.quit()
    
    # Return the email sent template
    return render_template('result.html')



if __name__ == '__main__':
    app.run(debug=True)
    