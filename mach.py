# Important note: all the library imported here are not required by this code, extra libraries have been added for adding further features to code. Please install only those which are required. Run this code in your IDE to see which are required to install. 
import subprocess
import wolframalpha
import pyttsx3
import tkinter as tk # for gui
import json
import random
import operator
import speech_recognition as sr # important for sppech recognition
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes # for jokes
import feedparser
import smtplib # for mail
import ctypes
import time
import requests # for sending request 
import shutil
import threading
import ssl
import io
import sys
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
ssl._create_default_https_context = ssl._create_unverified_context
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id) # 0 - male voice and 1 - female voice
assname='mach'  # name of assistance 
def speak(audio):
	engine.say(audio)
	engine.runAndWait()

def wishMe():
	hour = int(datetime.datetime.now().hour)
	if hour>= 0 and hour<12:
		speak("Good Morning Sir !")

	elif hour>= 12 and hour<18:
		speak("Good Afternoon Sir !") 

	else:
		speak("Good Evening Sir !") 

	
	speak("I am your Assistant")
	speak(assname)
	#speak("here")
	
# # following 8 lines are used to take user name as input and use throughout the session 
#def username():
#	speak("What should i call you sir")
#	uname = takeCommand()
#	speak("Welcome Mister")
#	speak(uname)
#	columns = shutil.get_terminal_size().columns

#	print("#####################".center(columns))
#	print("Welcome Mr.", uname.center(columns))
#	print("#####################".center(columns))
	#speak("here")
	speak("How can i Help you,")

def takeCommand():
	
	r = sr.Recognizer()
	
	with sr.Microphone() as source:
		r.adjust_for_ambient_noise(source)
		print("Listening...")
		r.pause_threshold = .5  # Adjust threshold according to your requirement# threshold is the minimum level of sound that must be present in order for the system to recognize a phrase or command. 
		audio = r.listen(source)

	try:
		print("Recognizing...") 
		query = r.recognize_google(audio, language ='en-in') #language
		print(f"User said: {query}\n")

	except Exception as e:
		print(e) 
		print("Unable to Recognize your voice.") 
		return "None"
	scroll_to_end()
	
	return query


def sendEmail(to_email, subject, message):
    sender_email = "*"  # email address
    sender_password = "*"  # email password / for sequrity use app password in gmail

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, to_email, f"Subject: {subject}\n\n{message}") # if voice assistance is unable to send recognice through command then use default mail to email(it is recipient mail)
        server.quit()
        speak("Email has been sent successfully.")
    except Exception as e:
		   #pass
        print(e)
        speak("Sorry, I am unable to send the email at the moment.")
def start_voice_assistant(stop_event):
	class StdoutRedirector(io.TextIOBase):
		def write(self, string):
			output_text.insert(tk.END, string)
	sys.stdout = StdoutRedirector()		
	print("mach Here,")
	wishMe()
	
	while not stop_event.is_set():
		
		query = takeCommand().lower()
		scroll_to_end()
		if 'wikipedia' in query:
			speak('Searching Wikipedia...')
			query = query.replace("wikipedia", "")
			results = wikipedia.summary(query, sentences = 3)
			speak("According to Wikipedia")
			print(results)
			speak(results)

		elif 'open youtube' in query:
			speak("Here you go to Youtube\n")
			webbrowser.open("youtube.com")

		elif 'open google' in query:
			speak("Here you go to Google\n")
			webbrowser.open("google.com")

		elif 'open stackoverflow' in query:
			speak("Here you go to Stack Over flow.Happy coding")
			webbrowser.open("stackoverflow.com") 

		elif 'play music' in query or "play song" in query:
			speak("Here you go with music")
			music_dir = ""   #add path to music drive
			songs = os.listdir(music_dir)
			print(songs) 
			random = os.startfile(os.path.join(music_dir, songs[1]))

		elif 'the time' in query:
			strTime = datetime.datetime.now().strftime("%H:%M:%S") ;print(strTime)
			speak(f"Sir, the time is {str(strTime)}") 


		elif 'send a mail' in query:
			
			
				speak("whome should i send")
				receipient = takeCommand()
				speak("What should be the subject of the email")
				subject = takeCommand() 
				speak("What should I say")
				message = takeCommand() 
				sendEmail(receipient,subject,message)
                


		elif 'how are you' in query:
			speak("I am fine, Thank you")
			speak("How are you, Sir")



		elif "change my name to" in query:
			query = query.replace("change my name to", "")
			assname = query

		elif "change name" in query:
			speak("What would you like to call me, Sir ")
			assname = takeCommand()
			speak("Thanks for naming me")

		elif "what's your name" in query: #or "What is your name" in query:
			speak("My friends call me")
			speak("mach")
			print("My friends call me mach" )

		elif 'exit' in query:
			speak("Thanks for giving me your time")
			exit()
			#sys.exit()

		elif "who made you" in query or "who created you" in query: 
			speak("I have been created by Rishi Singh.")
			
		elif 'joke' in query:
			speak(pyjokes.get_joke())


		elif 'search' in query or 'play' in query:
			
			query = query.replace("search", "") 
			query = query.replace("play", "")		 
			webbrowser.open("https://www.google.com/search?q=" +query) 

		elif "who i am" in query:
			speak("If you talk then definitely your human.")

		elif "why you came to world" in query:
			speak("Thanks to error404. further It's a secret")


		elif 'is love' in query:
			speak("It is 7th sense that destroy all other senses")

		elif "who are you" in query:
			speak("I am your virtual assistant created by error404")

		elif 'reason for you' in query:
			speak("I was created as a project by error404 ")


		elif 'news' in query:
			
			try: 
				jsonObj = urlopen('*') ##place your api key here
				data = json.load(jsonObj)
				i = 1
				
				speak('here are some top news ')
				print('''=============== Top headlines around World ============'''+ '\n')
				
				for item in data['articles']:
					if i==6:continue
					print(str(i) + '. ' + item['title'] + '\n')
					print(item['description'] + '\n')
					speak(str(i) + '. ' + item['title'] + '\n')
					i += 1
					scroll_to_end()
                    
     
			except Exception as e:
				
				print(str(e))

		
		elif 'lock window' in query:
				speak("locking the device")
				ctypes.windll.user32.LockWorkStation()

		elif 'shutdown system' in query:
				speak("Hold On a Sec ! Your system is on its way to shut down")
				subprocess.call('shutdown / p /f')
				

		elif "don't listen" in query or "stop listening" in query:
			speak("for how much time you want to stop mach from listening commands")
			a = int(input())
			time.sleep(a)
			print(a)

		elif "where is" in query:
			query = query.replace("where is", "")
			location = query
			speak("User asked to Locate")
			speak(location)
			webbrowser.open("https://www.google.nl/maps/place/"+location)

		elif "camera" in query or "take a photo" in query:
			ec.capture(0, "Jarvis Camera ", "img.jpg")

		elif "restart" in query:
			subprocess.call(["shutdown", "/r"])
			
		elif "hibernate" in query or "sleep" in query:
			speak("Hibernating")
			subprocess.call("shutdown / h")

		elif "log off" in query or "sign out" in query:
			speak("Make sure all the application are closed before sign-out")
			time.sleep(5)
			subprocess.call(["shutdown", "/l"])

		elif "write a note" in query:
			speak("What should i write, sir")
			note = takeCommand()
			file = open('mach.txt', 'w')
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
			file = open("mach.txt", "r") 
			print(file.read())
			speak(file.read(6))

		elif "mark" in query:
			
			speak("mach 1 point o in your service Mister")
			print("mach 1 point o in your service Mister")


		elif "weather" in query:
			api_key = "*"  #place weather api here
			base_url = "http://api.openweathermap.org/data/2.5/weather?q={}&appid={}" 

			speak("Which city's weather do you want to know?")
			city_name = takeCommand()  # Get the city name from the user

			complete_url = base_url.format(city_name, api_key)
			response = requests.get(complete_url)

			if response.status_code == 200:
				weather_data = response.json()
				if weather_data["cod"] != "404":
					main_data = weather_data["main"]
					weather_description = weather_data["weather"][0]["description"]

					temperature = main_data["temp"]
					pressure = main_data["pressure"]
					humidity = main_data["humidity"]
					print(f"In {city_name}, the temperature is {temperature} Kelvin.")
					speak(f"In {city_name}, the temperature is {temperature} Kelvin.")
					print(f"The weather is {weather_description}.")
					speak(f"The weather is {weather_description}.")
					print(f"The atmospheric pressure is {pressure} hPa.")
					speak(f"The atmospheric pressure is {pressure} hPa.")
					print(f"The humidity is {humidity}%.")
					speak(f"The humidity is {humidity}%.")

				else:
					speak("City not found. Please try again.")
			else:
				speak("Sorry, I couldn't fetch the weather data at the moment.")
			scroll_to_end()
			

		elif "wikipedia" in query:
			webbrowser.open("wikipedia.com")

		elif "Good Morning" in query:
			speak("A warm" +query)
			speak("How are you Mister")
			speak(assname)

		elif "will you be my gf" in query or "will you be my bf" in query: 
			speak("I'm not sure about, may be you should give me some time")

		elif "how are you" in query:
			speak("I'm fine, glad you me that")

		elif "i love you" in query:
			speak("It's hard to understand")

		elif "what is" in query or "who is" in query:
			try:
				app_id = "*" #place short answer api key here
				client = wolframalpha.Client(app_id)
				res = client.query(query)

				result = next(res.results).text
				print(result)
				speak(result)

			except Exception as e:
				print("Error:", str(e))
				speak("I'm sorry, I couldn't find information about that.")


		elif "very good" in query:
			speak("thank you. glad to hear that")	
		elif 'fine' in query or "good" in query:
			speak("It's good to know that your fine")

############  place multiple statements using this syntax
		# elif '****' in query or "****" in query:
		# 	speak("****************")


		
def stop_voice_assistant():
    global stop_event
    stop_event.set()

#GUI code 

root = tk.Tk()
root.title("mach")
root.geometry("600x500")  # window dimensions

background_color = "#272829" #color
font_style = "Times new roman" #font 
font_size = 16  #font size 

content_frame = tk.Frame(root, bg=background_color)
content_frame.place(relwidth=1, relheight=1)  


output_text = tk.Text(content_frame, wrap="word", background="#EAD7BB")
output_text.place(relx=0.05, rely=0.05, relwidth=0.9, relheight=0.6)


def on_click_start():
    global stop_event
    stop_event = threading.Event()
    assistant_thread = threading.Thread(target=start_voice_assistant, args=(stop_event,))
    assistant_thread.daemon = True
    assistant_thread.start()

#  to stop 
def on_click_stop():
    stop_voice_assistant()

#  start button 
button_start = tk.Button(content_frame, text="Start Assistant", command=on_click_start)
button_start.place(relx=0.2, rely=0.7, relwidth=0.2, relheight=0.1)

# stop button
button_stop = tk.Button(content_frame, text="Stop Assistant", command=on_click_stop)
button_stop.place(relx=0.6, rely=0.7, relwidth=0.2, relheight=0.1)


def scroll_to_end():
    output_text.yview_moveto(1.0)  

scroll_to_end() # will scroll the output in the window when it reaches to the end




#  loop
root.mainloop()
