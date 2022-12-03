from bs4 import BeautifulSoup as bs
from urllib.request import Request, urlopen
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
req = Request(
    url=url,
    headers={'User-Agent': 'Mozilla/5.0'}
)
page = urlopen(req).read()
soup = bs(page, features="html.parser")
result = soup.find("div", class_="ipc-metadata-list-summary-item__tc")
if result is None:
    speaker.Speak("No such movie exists.")
    del speaker
    quit()
openTitle = result.find("a")
titleUrl = "https://www.imdb.com" + openTitle.get("href")
req = Request(
    url=titleUrl,
    headers={'User-Agent': 'Mozilla/5.0'}
)
tPage = urlopen(req)
tSoup = bs(tPage, features="html.parser")
summary = tSoup.find("span", class_="sc-16ede01-1 kgphFu")
if summary is None:
    speaker.Speak("Plot does not exists.")
    del speaker
    quit()
plot = summary.string
print("Plot :")
print(plot)
speaker.Speak(plot)
del speaker
