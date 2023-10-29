import speech_recognition as sr
import os
import openai
import win32com.client
import webbrowser
import datetime
import AppOpener

import pyautogui as pag
import time



apikey="Your openai api key"

speaker=win32com.client.Dispatch('SAPI.SpVoice')

chatstr=""

def chat(query):
    global chatstr  
    openai.api_key=apikey

    print(chatstr)
    chatstr+=f"Mowassir: {query} \n jarvis: "

  
    response = openai.ChatCompletion.create(
      model="gpt-3.5-turbo",
      messages=[
    {
      "role": "user",
      "content": chatstr
    },
    {
      "role": "assistant",
      "content": ""
    },
    
  ],
  temperature=1,
  max_tokens=256,
  top_p=1,
  frequency_penalty=0,
  presence_penalty=0
)

    say(f"{response['choices'][0]['message']['content']}")
    chatstr+=f"{response['choices'][0]['message']['content']}\n"

    return response["choices"][0]["message"]["content"]





def ai(prompt):
    openai.api_key = apikey
    text = f"OpenAI response for Prompt: {prompt} \n *************************\n\n"

    response = openai.ChatCompletion.create(
     model="gpt-3.5-turbo",
      messages=[
    {
      "role": "user",
      "content": prompt
    },
    {
      "role": "assistant",
      "content": ""
    },
    
  ],
  temperature=1,
  max_tokens=256,
  top_p=1,
  frequency_penalty=0,
  presence_penalty=0
)

    # todo: Wrap this inside of a  try catch block
    
    text += response["choices"][0]["message"]["content"]
    if not os.path.exists("Openai"):
        os.mkdir("Openai")

    
    with open(f"Openai/{''.join(prompt.split('a.i')[1:]).strip() }.txt", "w") as f:
        f.write(text)




#Defining the audio output 
def say(text):
    speaker.speak(f"{text}")


#Defining audio input
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        # r.pause_threshold =  0.6
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some Error Occurred. Sorry from Jarvis"



def scroll_down():
    for i in range(10):
        pag.press("down")   


def scroll_up():
    for i in range(10):
        pag.press("up") 


def scroll_left():
    pag.press("left")

def scroll_right():
    pag.press("right")


def screenshot():
    pag.hotkey("win","prtsc")
def enter():
    pag.press("enter")
def spacebar():
    pag.press("space")
def delete_word():
    pag.hotkey("ctrl","backspace")

    
#volume down up and mute  ---------------done
#browser back forward refresh-------------done
#browser search and then write using listener----------half done
#clear button
#decimal
#esc enter space------------------------------done
#sleep shutdown restart-----------------------done
#select all and select word and select letters------------------half done
#copy and paste---------------------------------done




if __name__ == '__main__':
   
   say("Hello Mowassir, How can i assist you ")
   
   while True:
     
     print("listening...")
     query=takeCommand()
     #query=input("Enter the prompt:")
    

     

    
     
     
     apps=["zoom","notepad","brave","chrome","word","vlc","powerpoint","github","settings","clion","pycharm","virtual studio code"]
    # todo: list for songs
    #todo: list for videos
     sites = [["youtube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.com"], ["google", "https://www.google.com"],["facebook","facebook.com"],["instagram","instagram.com"],["gazi website","https://gazi.edu.tr/"]]
     for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                say(f"Opening {site[0]} sir...")
                webbrowser.open(site[1])

     for app in apps:  
         if f"open {app}".lower() in query.lower():
              pag.hotkey("win","m")
              say(f"opening {app}")
              AppOpener.open(app,match_closest=True)
         if f"close {app}".lower() in query.lower():
             say(f"Closing {app}")
             AppOpener.close(app,match_closest=True)

     if "the time".lower() in  query.lower():
         time=datetime.datetime.now().strftime("%H:%M")
         say(f"The time is {time}")



     elif "using a.i".lower() in query.lower():
         ai(prompt=query)

    
    
     

     elif "reset chat".lower() in query.lower():
         chatstr=""

     elif "jarvis".lower() in query.lower():
         print("chatting...")
         chat(query)

     elif "exit".lower() in query.lower():
         say("Bye sir , See you soon")
         exit()
    
     elif "hello world".lower() in query.lower():
         print(query.lower())


#Navigation realted functions
     elif "scroll down".lower() in query.lower():
         scroll_down()

     elif "scroll up".lower() in query.lower():
         scroll_up()
     
     elif "scroll left".lower() in query.lower():
         scroll_left()
     
     elif "scroll right".lower() in query.lower():
         scroll_right()
     
     elif "screenshot".lower() in query.lower():
         screenshot()

#Volume related functions

     elif "increase" and "volume".lower() in query.lower():
         pag.press("volumeup")

     elif "decrease" and "volume".lower() in query.lower():
         pag.press("volumedown0")

     elif "mute".lower() in query.lower():
         pag.press("volumemute")

     elif "unmute".lower() in query.lower():
         pag.press("volumemute")
    #  elif "pause".lower() or "play".lower() in query.lower():
    #      pag.press("pause")
    #      print("paused the video")


     
    
# Webbrowser and search related functions
     elif "search".lower() in query.lower():
         say("searching"+query[len("search"):])
         pag.press("browsersearch")
         pag.write(query[len("search"):])
         pag.press("enter")

     elif "go".lower() and "forward".lower() in query.lower():
         pag.press("browserforward")
         

     elif "go" and "back".lower() in query.lower():
         pag.press("browserback")
         
    
     elif "refresh".lower() in query.lower():
         pag.press("browserrefresh")
         
    #  elif "new tab" or "newtab".lower() in query.lower():
    #      pag.press("ctrl","t")
    #      print("opening new tab")


#Adding special keyboard keys
     elif "press" and "enter".lower() in query.lower():
        enter()
        

     elif "space".lower() in query.lower():
         spacebar()
         
     
     elif "capslock".lower() in query.lower():
         pag.press("capslock")
         

    #  elif "backspace".lower() in query.lower():
    #      print("backspacing a word")
    #      pag.press("backspace")
         
         #delete a word
     elif "delete" and "word".lower() in query.lower():
         pag.hotkey("ctrl","backspace")
         

#Adding shortcut Keys
     elif "select all".lower() in query.lower():
         pag.hotkey("ctrl","a")
         
     elif "deselect".lower() in query.lower(): 
         pag.click()
        

         
    # Adding cut copy and paste command
     elif "cut this".lower() in query.lower():
         pag.hotkey("ctrl","x")

    #  elif "copy".lower() in query.lower():
        #  pag.hotkey("ctrl","c")
        # pag.keyDown("ctrl")
        # pag.press("c")
        # pag.keyUp("ctrl")
        # print("copying")

     elif "paste".lower() in query.lower():
         pag.hotkey("ctrl","v")
         print("pasitong")

     elif "undo".lower() in query.lower():
         pag.hotkey("ctrl","z")
         print("undoing")

     elif "redo".lower() in query.lower():
        pag.hotkey("ctrl","y")
        print("redoing")
    
    
    




#typing related functions
     
     elif "type".lower() in query.lower():
         pag.write(query[len("type "):])











     elif "sleep".lower() in query.lower():
         say("tell me the code sir")
         code=takeCommand()
         if "10" or "ten".lower() in code.lower():
            say("Good night, Sir")
            pag.hotkey("alt","f4")
            pag.press("up")
            pag.press("enter")
         else:
           say("wrong code")
           exit

     elif "restart this laptop".lower() in query.lower():
         say("tell me the code sir")
         code=takeCommand()
         if "10" or "ten".lower() in code:
             say("see you soon, Sir")
             pag.hotkey("alt","f4")
             pag.press("down")
             pag.press("enter")
         else:
             say("wrong code sir")
             exit

     elif "shutdown this laptop".lower() in query.lower():
         say("Tell me the code, Sir")
         code=takeCommand()
         if "10" or "ten".lower() in code:
             say("see you soon, Sir")
             pag.hotkey("alt","f4")
             pag.press("enter")
         else:
             say("Wrong code, Sir")
             exit
###     Another way to write shutdown and other functions functions   
#     elif "shutdown" and "code" and "ten".lower() in query.lower():
#          pag.hotkey("alt","f4")
#          pag.press("enter")###


     
    





     

#-----------------------------------Typewriter (with some issue)-----------------------------------------------------------
#  elif "start" and "typewriter".lower() in query.lower():
    #      while True:
    #          typer=input("Enter what you want to write: ")

    #          time.sleep(5)
    #          if typer.lower() !="exit typewriter".lower():
    #                  if ("newline" or "new line").lower() in typer.lower():
    #                    enter()
                    
    #                  elif "space".lower() in typer.lower():
    #                    spacebar()

    #                  elif "backspace" or "delete".lower() in typer.lower():
    #                        delete_word()
            
    #                  pag.write(typer)
                         
                     
    #          else:
    #                break










         
         




