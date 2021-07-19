"""
Created on Sat Mar 21 22:47:28 2020

@author: CodeSRJ
"""

import os
import wolframalpha
appID="Enter your Wolframalpha API"  # Wolframalpha API
import wikipedia
import win32com.client as wincl
import speech_recognition as sr
import datetime
import time
import webbrowser

c=0 # tracks the no of times it greets the user
repeat=0 # in case of unsuccessful audio capture, asks to repeat
speak = wincl.Dispatch("SAPI.SpVoice") # Microsoft Text-to-Speech Engine


def eight_age():
    build_date = datetime.date(2020,3,18)
    current_date = datetime.date.today()
    try:
        if (abs(current_date.year-build_date.year)!=0 and abs(current_date.month-build_date.month)!=0 and abs(current_date.day-build_date.day)!=0):
            text=("I am",abs(current_date.year-build_date.year),"year",abs(current_date.month-build_date.month),"month and",abs(current_date.day-build_date.day),"day old")
            speak.Speak(text)
        elif (abs(current_date.year-build_date.year)!=0 and abs(current_date.month-build_date.month)!=0 and abs(current_date.day-build_date.day)==0):
            text=("I am",abs(current_date.year-build_date.year),"year and",abs(current_date.year-build_date.year),"month old")
            speak.Speak(text)
        elif (abs(current_date.year-build_date.year)!=0 and abs(current_date.month-build_date.month)==0 and abs(current_date.day-build_date.day)!=0):
            text=("I am",abs(current_date.year-build_date.year)," year and",abs(current_date.year-build_date.year),"day old")
            speak.Speak(text)
        elif (abs(current_date.year-build_date.year)==0 and abs(current_date.month-build_date.month)!=0 and abs(current_date.day-build_date.day)!=0):
            text=("I am",abs(current_date.month-build_date.month),"month and",abs(current_date.day-build_date.day),"day old")
            speak.Speak(text)
        elif (abs(current_date.year-build_date.year)==0 and abs(current_date.month-build_date.month)==0 and abs(current_date.day-build_date.day)!=0):
            text=("I am",abs(current_date.day-build_date.day),"day old")
            speak.Speak(text)
        elif (abs(current_date.year-build_date.year)==0 and abs(current_date.month-build_date.month)!=0 and abs(current_date.day-build_date.day)==0):
            text=("I am",abs(current_date.month-build_date.month),"month old")
            speak.Speak(text)
        elif (abs(current_date.year-build_date.year)!=0 and abs(current_date.month-build_date.month)==0 and abs(current_date.day-build_date.day)==0):
            text=("I am",abs(current_date.year-build_date.year),"year old")
            speak.Speak(text)
    except:
        speak.Speak("Some error occured.")
        
        
        
def openweb(question):
    if 'youtube' in question.lower():
        webbrowser.open('http://youtube.com', new=0)
    elif 'google' in question:
        webbrowser.open('http://google.com', new=0)
    else:
        print(question)
        r=sr.Recognizer()  # initialising the recognizer
        
        with sr.Microphone() as source:
            
            while True:
                try:
                    while(True): # to get either yes or no
                        speak.Speak("wanna spell")
                        print("speak")
                        audio=r.listen(source)
                        spell_question=r.recognize_google(audio)
                        print(spell_question)

                        if 'no' in spell_question:
                            wpage=r.recognize_google(audio)
                            speak.Speak("Can I know the name of webpage you want to open")
                            print('speak')
                            audio=r.listen(source)
                            wpage=r.recognize_google(audio)
                            webpage=("http://"+wpage+".com").lower()
                            webbrowser.open( webpage, new=0)
                            break

                        elif 'yes' in spell_question:
                            speak.Speak("Okay, you can spell the domain name now")
                            print('speak')
                            audio=r.listen(source)
                            spell=r.recognize_google(audio)
                            print(spell)
                            spell_ans=spell.replace(' ','').lower()
                            print(spell_ans)
                            webpage=("http://"+spell_ans+".com").lower()
                            webbrowser.open( webpage, new=0)
                            break

                        else:
                            speak.Speak("Sorry, i could not understand the spell , please try again")
                            break  # if any gets executed within try
                    break # if try succeeds

                except:
                    speak.Speak('Sorry I could not understand, please try again')


while(True):
    r=sr.Recognizer()  # initialising the recognizer
    with sr.Microphone() as source:   # using microphone as the source for input

        # audio/question capture loop
        while(True):
            try:
                 if repeat==0:   
                    if c==0:
                        speak.Speak("How can I help you?")
                        print('speak')
                        audio=r.listen(source)
                        question=r.recognize_google(audio).lower()
                        c=1  # keepd track so that it speaks "How can I help you?" only once
                    else:
                        time.sleep(0.5)
                        speak.Speak("Anything Else")
                        print('speak')
                        audio=r.listen(source)
                        question=r.recognize_google(audio).lower()

                    break # exits audio/question capture on successfil capture

                 else:  # in case some error occured while capturing the voice
                    speak.Speak("can you please repeat!")
                    print('speak')
                    audio=r.listen(source)
                    question=r.recognize_google(audio)
                    repeat=0  # audio capture successful
                    break # exits audio/question capture on successfil capture

            except:
                speak.Speak("Sorry, I could not understand ")
                repeat=1  # audio capture failed

    # question processing begins here
    if ('nothing' in question) or ('no' in question) or ('nope' in question) or ('nah' in question) or ('quit' in question) or ('exit' in question) or ('none' in question):
        print(question)
        speak.Speak("Thank You")
        break # exits outer while / program terminates

    elif ('who are you'== question) or ("what are you")==question or ("your name" in question) or ("your identity" in question) or ("origin" in question) or ('about yourself'== question) :
            speak.Speak("I am EIGHT, The Personal Digital Assistant")
    
    elif ("your age" in question) or ("old are you" in question):
        print(question)
        eight_age()

    elif ('made you' in question) or ('designed you' in question) or ('created you' in question) or ('build you' in question) or ('built you' in question) or ('programmed you' in question) or ('code' in question) or ('invented you' in question) or (('you ' in question) and ('invented' in question)):
        print(question)
        speak.Speak("I have been designed, programmed and created by Code S R J  ")

    elif ('what can you do' in question) or ('what can you perform' in question) or ('task' in question) or ('action' in question) or ('tasks' in question) or ('actions' in question) or ('can you help me' in question) or ('help me' in question):
        print(question)
        speak.Speak("For now, i can open URL or webpages, look into wikipedia and other resources. I can also tell you about myself. I am continuously being updated with new features. You can check it later.")

    elif 'open' in question:
        openweb(question)
        
    else:
        try:
            client=wolframalpha.Client(appID)
            res=client.query(question)
            answer=next(res.results).text # extracting answer from query result
            print(question)
            speak.Speak(answer)
            
        except:
            wiki_text="Hey i looked into wikipedia and found " + wikipedia.summary(question,sentences=2)
            # speak.Speak(wiki_text)
            speak.Speak(wikipedia.summary(question,sentences=2))