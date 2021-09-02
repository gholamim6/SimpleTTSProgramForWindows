# simple sapi5 text to speech program
# this code is most derived from Charlies answer in this StackOverFellow post with a bit improvement
# https://stackoverflow.com/a/61388230/15433957

import win32com.client as wincl

SAPI = wincl.Dispatch("SAPI.SpVoice")
voices = SAPI.GetVoices()
text = input("enter a text to read: ")
print("Enter the number of speech engine to for reading this text:")
for number, voice in enumerate(voices, start=1):
    print("{}: {}".format(number, voice.GetAttribute("name")))
    """
    voices object is an itterable and can be iterated using loop and other functions. i removed speaker number. because different computer has different sapi five speech engines
    for example speakerNumber 1 caused me error. because i had only microsoft Anna on my Windows 7 and i could choose only voice number 0.
    Later i installed my additional sapi5 speech engines. and In windows 10, We have mor than one sapi5 speech engine.
    """

voiceNumber = input()
try:
    voiceNumber = int(voiceNumber) - 1
    if voiceNumber < 0 or voiceNumber >= len(voices):
        print("Wrong Number")
    else:
        SAPI.Voice = voices.Item(voiceNumber)
        
        """
        set_voice method also doesn't work
        I guess in latest version, this module uses class property for getting and setting value for object attributes.
        so we can use Voice attribute to set and get value. which i used.
        """
        
        print("you successfully selected {} as your speech engine. let's listen to your text.".format(SAPI.Voice.GetAttribute("Name")))
        SAPI.Speak(text)
except ValueError:
    print("you didn't enter a valid number")
    # end of code