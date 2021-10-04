from comtypes.client import CreateObject 
from comtypes.gen import SpeechLib


engine = CreateObject("SAPI.SpVoice")
stream = CreateObject("SAPI.SpFileStream")


infile = input("Enter File Name .... \n")
pages = input("Enter The range Of pages..... \n")
outfile = f"sound-{infile}-{pages}.mp3"


stream.Open(outfile, SpeechLib.SSFMCreateForWrite)
engine.AudioOutputStream = stream
f = open(f"{infile}.txt", 'r')
theText = f.read()
f.close()
engine.speak(theText)
stream.Close()