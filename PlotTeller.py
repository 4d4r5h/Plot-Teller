from bs4 import BeautifulSoup as bs
import urllib.request as req
from win32com.client import Dispatch
import speech_recognition as sr

url = "https://www.imdb.com/find?q="
r = sr.Recognizer()
speaker = Dispatch("SAPI.SpVoice")
sr.Microphone.list_microphone_names()
mic = sr.Microphone(device_index=1)
with mic as source:
    tellName = "Say the name of the movie for which you want plot."
    speaker.Speak(tellName)
    print("Speak now.")
    audio = r.listen(source)
    nameUrl = r.recognize_google(audio)
    name = "You said : " + nameUrl
    print(name)
    speaker.Speak(name)
nameUrl = nameUrl.strip().replace(" ", "+")
url = url + nameUrl + "&s=tt"
page = req.urlopen(url)
soup = bs(page, features="html.parser")
result = soup.find("td", class_="result_text")
if result is None:
    speaker.Speak("No such movie exists.")
    del speaker
    quit()
openTitle = result.find("a")
titleUrl = "https://www.imdb.com" + openTitle.get("href")
tPage = req.urlopen(titleUrl)
tSoup = bs(tPage, features="html.parser")
summary = tSoup.find("div", class_="summary_text")
if summary.find("a"):
    fullSummary = summary.find("a")
    if str(fullSummary.get("href")).startswith("https:"):
        speaker.Speak("No plot available.")
        del speaker
        quit()
    summaryUrl = "https://www.imdb.com" + fullSummary.get("href")
    summaryPage = req.urlopen(summaryUrl)
    newSoup = bs(summaryPage, features="html.parser")
    newSummary = newSoup.find("li", class_="ipl-zebra-list__item")
    plot = newSummary.find("p").string
else:
    plot = summary.string.strip()
print("Plot :")
print(plot)
speaker.Speak(plot)
del speaker
