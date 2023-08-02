from typing import Any, Union

import gtts
from win32com.client import Dispatch
import speech_recognition as sp
from win32com.client.dynamic import CDispatch
import math

r = sp.Recognizer()


def speak(string):
    spear = Dispatch("SAPI.SpVoice")
    spear.Speak(f"okay.answer={string}")


def listen():
    print("speak")
    with sp.Microphone() as source:
        audio = r.listen(source)
        MyText = r.recognize_google(audio)
        print(MyText)
        return MyText


words = listen()
x = words.split()
if x[1] == '+':
    speak(int(x[0]) + int(x[2]))
    print(int(x[0]) + int(x[2]))
elif x[1] == '-':
    speak(int(x[0]) - int(x[2]))
    print(int(x[0]) - int(x[2]))
elif x[1] == 'into':
    speak(int(x[0]) * int(x[2]))
    print(int(x[0]) * int(x[2]))
elif x[1] == 'by':
    speak(int(x[0]) / int(x[2]))
    print(int(x[0]) / int(x[2]))
elif x[1] == 'modulo':
    speak(int(x[0]) % int(x[2]))
    print(int(x[0]) % int(x[2]))
elif x[1] == 'power':
    a = int(x[0])
    b = int(x[2])
    speak(math.pow(a, b))
    print(math.pow(a, b))
elif x[1] == 'factorial':
    a = int(x[0])
    speak(math.factorial(a))
    print(math.factorial(a))
elif x[1] == 'floor':
    a = float(x[0])
    speak(math.floor(a))
    print(math.floor(a))
elif x[1] == 'degree':
    a = int(x[0])
    speak(math.radians(a))
    print(math.radians(a))
elif x[0] == 'sin':
    a = int(x[1])
    speak(math.sin(a))
    print(math.sin(a))
elif x[0] == 'cos':
    a = int(x[1])
    speak(math.cos(a))
    print(math.cos(a))
elif x[0] == 'Tan':
    a = int(x[1])
    speak(math.tan(a))
    print(math.tan(a))
elif x[0] == 'hyperbolic':
    a = int(x[1])
    speak(math.tanh(a))
    print(math.tanh(a))
else:
    speak("try again")
    print("try again")
