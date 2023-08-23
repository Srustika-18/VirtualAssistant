from win32com.client import Dispatch
from gtts import gTTS
import re
import json
import requests
import os
import webbrowser
import wikipedia
import time
import speech_recognition as sprec
from playsound import playsound


def speakmay1(talkie):
    spk = Dispatch("SAPI.SpVoice")
    spk.Speak(f'{talkie}')


def speakmay2(talkie):
    mytalks = str(talkie)
    language = 'hi'
    myobj = gTTS(text=mytalks, lang=language, tld='co.in', slow=False)
    date_string = time.strftime("%d%m%Y%H%M%S", time.localtime())
    filename = "voice"+date_string+".mp3"
    myobj.save(filename)
    playsound(filename)
    os.remove(filename)


def speakmay(talkie):
    try:
        speakmay1(talkie)
    except Exception as e:
        speakmay2(talkie)


def wishmemay():
    nowhr = int(time.strftime("%H", time.localtime()))

    if nowhr < 12:
        speakmay("Good Morning!")
    elif nowhr >= 12 and nowhr <= 16:
        speakmay("Good Afternoon!")
    elif nowhr > 16 and nowhr <= 21:
        speakmay("Good Evening!")
    else:
        speakmay("Good Night!")
    speakmay("How can I help you!?")


def listen():
    r = sprec.Recognizer()
    with sprec.Microphone() as source:
        print("\nSay. Press Ctrl + C to exit.")
        r.pause_threshold = 1
        audio = r.listen(source, timeout=7, phrase_time_limit=12)
    try:
        print("Recognising...")
        usertalks = r.recognize_google(audio, language="en-in")
        print(f'\n ~ {usertalks}\n')
        return usertalks
    except Exception as e:
        time.sleep(1)
        print("Say that again.")


def OpenweatherAPIcalltext():
    '''Uses Open Weather API to fetch weather data of Bhubaneswar'''
    import json
    import requests

    weatherjson = requests.get(
        "https://api.openweathermap.org/data/2.5/weather?lat=20.292&lon=85.838&units=metric&appid=145c485b75dd3ef3084c53041a74560f")

    textparsed = json.loads(weatherjson.text)

    weather = textparsed['weather'][0]['description']
    temperature = textparsed['main']['temp']
    humidity = textparsed['main']['humidity']
    feels_like = textparsed['main']['feels_like']
    clouds = textparsed['clouds']['all']
    wind = textparsed['wind']['speed']
    name = textparsed['name']

    openweathercall = f'''The weather is {weather} in {name}. 
	Temperature is {temperature}°C. 
	Due to {humidity}% humidity, it feels like {feels_like}°C.
	Cloud Coverage is {clouds}%.
	And Wind Speed is {wind}m/s.'''

    return openweathercall


def newsreturn(numberofnews, topic):
    newsap = requests.get(
        f"https://newsapi.org/v2/top-headlines?country=in&pageSize={numberofnews+1}&page=1category={topic}&apiKey=ebfb37ba82c542b1979be299547484dd").text
    jsondata = json.loads(newsap)

    newslist = [jsondata['articles'][i]['title'] for i in range(numberofnews)]
    return newslist


def quotereturn():
    url = "https://zenquotes.io/?api=random"
    quotetext = requests.get(url).text
    return [json.loads(quotetext)[0]['q'], json.loads(quotetext)[0]['a']]


if __name__ == '__main__':
    os.system('cls')

    wishmemay()

    queuedict = {}

    while True:
        breakingnum = 0

        while True:
            say = listen()
            if say is not None:
                break

        wants = say.lower()  # type: ignore

        if wants == "exit" or 'bye' in wants:
            speakmay("Bye Bye! Shutting down.")
            break

        elif "wikipedia" in wants:
            try:
                speakmay("Found this from Wikipedia.")
                results = wikipedia.summary(
                    wants.replace("wikipedia", ""), sentences=4)
                print(results)
                speakmay(results)

            except Exception as e:
                print("Couldn't fetch information. Try again.")

        elif "time" in wants:
            timenow = f'Time is {time.strftime("%I:%M%p",time.localtime())}.'
            print(timenow)
            speakmay(timenow)

        elif "weather" in wants:
            try:
                speakmay(OpenweatherAPIcalltext())
            except Exception as e:
                print("Couldn't fetch weather. Try again.")

        elif "news" in wants:
            try:
                # business entertainment general health science sports technology
                if "business" in wants:
                    topic = "business"
                elif "entertainment" in wants:
                    topic = "entertainment"
                elif "health" in wants:
                    topic = "health"
                elif "science" in wants:
                    topic = "science"
                elif "sport" in wants:
                    topic = "sports"
                elif "tech" in wants:
                    topic = "technology"
                else:
                    topic = "general"

                numnews = re.findall(r'\d+', wants)
                if numnews == ["0"] or numnews == []:
                    numberofnews = 5
                else:
                    numberofnews = int(numnews[0])
                print(f"\nTop {topic.title()} Headlines of Today:\n")
                for index, item in enumerate(newsreturn(numberofnews, topic), 1):
                    print(f'{index}. {item}')
                    speakmay(item)

            except Exception as e:
                print("Couldn't get news. Try again.")

        elif "date" in wants or "today" in wants:
            todayday = time.strftime("%A", time.localtime())
            todaydate = time.strftime("%d %B", time.localtime())
            daydatewish = f'Today is {todayday} and Date is {todaydate}.'
            print(daydatewish)
            speakmay(daydatewish)

        elif "open" in wants:

            if "google" in wants:
                webbrowser.open("www.google.com")
            elif "youtube" in wants:
                webbrowser.open("www.youtube.com")
            elif "whatsapp" in wants:
                webbrowser.open("https://web.whatsapp.com")

        elif "search" in wants and "youtube" in wants:
            keyword = wants.replace("search ", "").replace("youtube ", "")
            print(f"Searching for {keyword}")
            url = f'https://www.youtube.com/results?search_query={keyword.replace(" ","+")}'
            webbrowser.open_new_tab(url)

        elif "search" in wants:
            url = f'https://www.google.com/search?q={wants.replace("search","").replace("google","").replace(" ","+")}'
            webbrowser.open_new_tab(url)

        elif "joke" in wants:
            try:
                url = "https://v2.jokeapi.dev/joke/Any"
                joketext = requests.get(url)
                jokejson = json.loads(joketext.text)
                print(jokejson['joke'])
                speakmay(jokejson['joke'])
            except Exception as e:
                print("Couldn't get joke. Try again.")

        elif "quote" in wants:
            try:
                quote = quotereturn()[0]
                author = quotereturn()[1]
                print(quote, "  -", author)
                speakmay(quote)
            except Exception as e:
                print("Couldn't get quote. Try again.")

        elif "say" in wants[0:4] or "speak" in wants[0:6]:
            speakmay(wants.replace("say", "").replace("speak", ""))

        else:
            speakmay("Sorry, I can't help with this.")

        time.sleep(2)