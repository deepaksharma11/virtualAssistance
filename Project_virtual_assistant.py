import random # for random number generation
import datetime # for data and time
# install (speechRecognition) module to convert speech to text in microsoft
# install pyaudio if missing(for windows use:- pip install pipwin and then pipwin install pyaudio)
import speech_recognition as sr
import pynotifier # used to send notification on windows os
import PIL # Python Image Library used for processing jpeg images using tkinter
import os # for operating system commands
import wikipedia # for using wikipedia
import webbrowser # for using web browser
import smtplib # for using mailing functions


# offline voice
def speak(audio):
    ''' This function take string and convert into speech '''
    from win32com.client import Dispatch
    speak = Dispatch("SAPI.SpVoice")    
    # microsoft speech api
    speak.Speak(audio)

def takeCommand():
    '''This function take command as speech and returns string'''
    var1 = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        var1.pause_threshold = 1
        audio = var1.listen(source)
    try:
        print("Recognizing...")
        query = var1.recognize_google(audio, language = 'en-in')
        print(f"User said : {query}\n")

    except Exception as e:
        print(e) #printing error
        print("Say that again please...")
        return "None"

    return query

def wishMe():
    ''' This function wish user as per the time '''
    hour = int(datetime.datetime.now().hour)
    if hour>=4 and hour<12:
        speak("Good Morning, Sir")
    elif hour>=12 and hour<16:
        speak("Good Afternoon, Sir")
    elif hour>=16 and hour<=21:
        speak("Good Evening, Sir")
    else:
        speak("It's night already, anything urgent sir?")
    speak(f"Myself {vname}, How may I help you?")

def tellTime():
    ''' This function wish user as per the time '''
    right = datetime.datetime.now().time()
    return f"it is {right.hour} hours: {right.minute} minutes : {right.second} seconds"

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('youremailid@gmail.com', "password")    
    server.sendmail('youremailid@gmail.com', to, content)
    server.close()

if __name__ == "__main__":    
    vname = 'Shivi'
    wishMe()
    while True:
        command = takeCommand().lower()     
        #tasks are written after it
        if "the time" in command:
            result = tellTime()
            speak(result)

        elif "wikipedia" in command:
            speak("Searching in Wikipedia...")
            command = command.replace('wikipedia', "")
            result = wikipedia.summary(command, sentences = 2)
            print(result)
            speak(result)

        elif "open youtube" in command:
            webbrowser.open_new_tab("youtube.com")

        elif "open google" in command:
            webbrowser.open_new_tab("google.com")
            
        elif "search" in command:
            speak("Tell me, what do you want me to search sir?")
            search_text = takeCommand()
            search_text = search_text.replace(" ", "+")
            webbrowser.open_new_tab(f"https://www.google.com/search?q={search_text}")

        elif "open visual" in command:
            os.startfile("C:\\Users\\Deepak Sharma\\AppData\\Roaming\\Microsoft\\Windows\Start Menu\\Programs\\Visual Studio Code\\Visual Studio Code")
        
        elif "write down" in command:
            speak("Tell me what do you want me to write sir")
            file_text = takeCommand()
            while True:
                speak("Title your file sir")
                file_title = takeCommand()
                if file_title != "None":
                    break                  
            try:
                with open(f"C:\\Users\\Deepak Sharma\\Desktop\\{file_title}.txt", "wt") as f:
                    f.write(file_text)
            except Exception as e:
                print(e)
                speak("Sorry, I failed this time.")
            else:
                speak("Great! we are done saving your thought.")

        elif "play music" in command:
            print("Mix\nGym\nMotivational\nRomantic\nSlow")
            speak("Which genre would you like to listen right now sir. Here we have Mix, Gym, Motivational, Romantic and Slow music")
            # Create these folder in your music folder with seperated music
            while True:
                song = takeCommand().lower()
                try:
                    music_folder = f"D:\\Music\\{song}"
                    songs = os.listdir(music_folder)
                    os.startfile(os.path.join(music_folder, random.choice(songs)))
                except Exception as e:
                    print(e) #printing error
                else:
                    break

        elif "play movie" in command:
            movie_folder = "D:\\Movies"
            movies = os.listdir(movie_folder)
            for i in range(len(movies)):
                print(f"{i+1}. {movies[i]}")
            speak("Tell me which movie you want to watch")
            while True:
                movieno = takeCommand()
                try:
                    os.startfile(os.path.join(movie_folder, movies[int(movieno) - 1]))
                except Exception as e:
                    print(e) #printing error
                else:
                    break         

        elif "email" in command:
            maildic= {"name" : "emailid@gmail.com", "name2": "emailid@gmail.com"}
            while True:
                speak("To whome you want to send the email")
                print("We have [name, name2]")
                too = takeCommand().lower()
                too = too.replace(" ", "")
                if '@' in too and "." in too:
                    to = too
                else:
                    for x in maildic:
                        if too == x:
                            to = maildic[x]
                            break
                        else:
                            to = None
                    if to == None:
                        speak("Sorry Sir, I didn't get it")
                        continue
                print(to)
                speak(f"Is {to} is a correct address")
                correct = takeCommand().lower()
                if "yes" not in correct:
                    continue
                speak(f"What should I say to {too}?")
                message = takeCommand()
                message = message + f"\n\n\n\n\nThanks and Regards,\nDeepak Sharma"
                speak("I hope you are done!")
                sender = "youremailid@gmail.com"
                speak("Please, title your mail")
                subject = takeCommand().capitalize()
                # speak("")
                body = '\r\n'.join(['To: %s' % to, 'From: %s' % sender, 'Subject: %s' % subject,'', message])
                speak("May I send it sir?")
                permission = takeCommand().lower()
                if permission != "yes" or permission != "send":
                    break                  
                try:
                    sendEmail(to, body)
                except Exception as e:
                    print(e)
                    speak("Sorry sir, I am not able to send this email")
                else:
                    speak("Email has been sent successfully!")
                finally:
                    speak("Would you like to send any other e-mail?")
                    ask = takeCommand().lower()
                    if 'yes' in ask:
                        pass
                    else:
                        speak("Returning to main menu")
                        break

        elif "exit" in command:
            hour = int(datetime.datetime.now().hour)
            if hour>=6 and hour<22:
                speak("Good Bye, Sir")
            elif hour<6 or hour>=22:
                speak("Good Night, Sir")
            break

