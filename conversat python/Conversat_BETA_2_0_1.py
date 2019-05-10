import speech_recognition as sr
import win32com.client as wincl
import time
import os
line_no =1;
line_no2=1;
get_Input =" ";
speak = wincl.Dispatch("SAPI.SpVoice")
print("hello, this is conversat beta 2.0.1 python version\n")
speak.Speak("hello, this is conversat beta 2.0.1 python version\n")
print("Type your queries\n")
speak.Speak("talk to me...")
# obtain audio from the microphone
r = sr.Recognizer()
while True:
    with sr.Microphone() as source:
        print("Please wait. Calibrating microphone...")
        speak.Speak("Please wait. Calibrating microphone...")
        # listen for 5 seconds and create the ambient noise energy level  
        r.adjust_for_ambient_noise(source, duration=5)  
        print("Say something!")
        speak.Speak("Say somthing!")
        audio = r.listen(source)
    #recognize using google
    try:
        os.system("cls")
        get_Input = r.recognize_google(audio)
        if (get_Input == "exit"):
            speak.Speak("Conversat closing, thank you sir!")
            break
        else:
            found = 0;line_no = 1;line_no2 = 1;
            for line in open("QUES.txt","r"):
                if get_Input in line: found=1; break;
                else: line_no += 1;

            if found==1:
                for line in open("ANS.txt","r"):
                    if line_no2==line_no:speak.Speak(line);time.sleep(1);break;
                    else:line_no2+=1;
            else:
                with open("QUES.txt","a+") as f:
                    f.seek(0,0)
                    f.write(get_Input+"\n");
                    with open("ANS.txt","a+") as fo:
                        f.seek(0,0)
                        speak.Speak("I'm Sorry sir, I don't know the answer, please tell me what to say")
                        with sr.Microphone() as source1:
                            print("Please wait. Calibrating microphone...")
                            speak.Speak("Please wait. Calibrating microphone...")
                            # listen for 5 seconds and create the ambient noise energy level  
                            r.adjust_for_ambient_noise(source1, duration=5)  
                            print("Say something!")
                            speak.Speak("Say somthing!")
                            audio1 = r.listen(source1)
                        try:
                            ans = r.recognize_google(audio1)
                            spk = "your answer is "+ans
                            speak.Speak(spk)
                            fo.write(ans+"\n")
                        except sr.UnknownValueError:
                            print("google could not understand audio")
                        except sr.RequestError as e:
                            print("google error; {0}".format(e))
                            

    except sr.UnknownValueError:  
        print("google could not understand audio")  
    except sr.RequestError as e:  
        print("google error; {0}".format(e))  


