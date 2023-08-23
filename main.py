from gtts import gTTS
import re
import json
import requests
import os
import webbrowser
import wikipedia
import youtube_dl
import time
import speech_recognition as sprec
from playsound import playsound
from win32com.client import Dispatch
from pygame import mixer
mixer.init()


def speak1(talkie):
    spk = Dispatch("SAPI.SpVoice")
    spk.Speak(f'{talkie}')


def speak2(talkie):
    mytalks = str(talkie)
    language = 'hi'
    myobj = gTTS(text=mytalks, lang=language, tld='co.in', slow=False)
    date_string = time.strftime("%d%m%Y%H%M%S", time.localtime())
    filename = "voice"+date_string+".mp3"
    myobj.save(filename)
    playsound(filename)
    os.remove(filename)


def speak(talkie):
    try:
        speak1(talkie)
    except Exception as e:
        speak2(talkie)


def wishme():
    nowhr = int(time.strftime("%H", time.localtime()))

    if nowhr < 12:
        speak("Good Morning!")
    elif nowhr >= 12 and nowhr <= 16:
        speak("Good Afternoon!")
    elif nowhr > 16 and nowhr <= 21:
        speak("Good Evening!")
    else:
        speak("Good Night!")
    speak("How can I help you!?")


def listen():
    r = sprec.Recognizer()
    with sprec.Microphone() as source:
        print("\nSay.<3")
        r.pause_threshold = 1
        audio = r.listen(source, timeout=7, phrase_time_limit=12)

    try:
        print("Recognising...")
        usertalks = r.recognize_google(audio, language="en-in")
        print(f'\n ~ {usertalks}\n')
        return usertalks
    except Exception as e:
        if breakingnum >= 20 and mixer.music.get_busy() == 0:
            print("\nShutting Down.")
            speak("Shutting Down.")
        else:
            print("Say that again.")
            time.sleep(1)
        return None  # not working as of now


def ytfirsturlreturn(query):

    global queuedict
    results = requests.get(
        f'https://www.youtube.com/results?search_query={query.replace(" ","+")}').text

    resultsIndex = results.index('{"videoId":')
    resultsRefined = results[resultsIndex:resultsIndex+3000]
    found = re.search(r'{"videoId":"[-.\d\w]+',
                      resultsRefined)[0].split("\"")[3]
    foundindex = resultsRefined.index(found)
    containingtitle = resultsRefined[foundindex:foundindex+1500]
    title = re.findall(r'{"text":"[^"]+"}', containingtitle)[0].split('"')[3]
    url = f'https://youtu.be/{found}'
    queuedict[url] = title

    return url


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


def ytdlMusicPlay(keyword):
    infoFile = ytdl.extract_info(keyword, download=True)

    nowplaying = infoFile['entries'][0]['title']
    nowid = infoFile['entries'][0]['id']

    if mixer.music.get_busy() == 0:
        print(f"Playing {nowplaying} !")

        mixer.music.load(fr'{nowid}.mp3')
        mixer.music.play()

    else:
        mixer.music.queue(fr'{nowid}.mp3')

        print("\nWait for this song to play to add another song to the queue!")
        time.sleep(2)


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


ytdl_format_options = {
    # 'format': 'bestaudio/best',
    'format': 'worstaudio/worst',
    'outtmpl': '%(id)s.%(ext)s',

    'ffmpeg_location': r"ffmpeg-master-latest-win64-gpl\bin\ffmpeg.exe",

    'postprocessors': [
        {'key': 'FFmpegExtractAudio'},
        {'key': 'FFmpegExtractAudio',
         'preferredcodec': 'mp3'}],

    'restrictfilenames': True,
    'noplaylist': True,
    # 'forceduration': True,
    'nocheckcertificate': True,
    'ignoreerrors': False,
    'logtostderr': False,
    'quiet': True,
    'no_warnings': True,
    'default_search': 'auto',
    # bind to ipv4 since ipv6 addresses cause issues sometimes
    'source_address': '0.0.0.0'
}

ytdl = youtube_dl.YoutubeDL(ytdl_format_options)


if __name__ == '__main__':
    os.system('cls')

    wishme()

    queuedict = {}

    while True:
        breakingnum = 0
        while True:
            os.system('cls')
            try:
                say = listen()
                if say is not None and breakingnum < 20:
                    breakingnum = 0
                    break
                elif mixer.music.get_busy() == 1:
                    breakingnum = 1
                elif breakingnum >= 20 and mixer.music.get_busy() == 0:
                    exit()
                else:
                    breakingnum += 1
            except Exception as e:
                speak("Speak a bit earlier!")

        # wants = input("\n\nEnter: ").lower()
        wants = say.lower()
        # listen()
        if mixer.music.get_busy() == 1:
            if "pause" in wants:
                mixer.music.pause()
                continue

            elif "stop" in wants:
                mixer.music.stop()
                continue

            elif "rewind" in wants:
                mixer.music.rewind()
                continue

            elif "volume" in wants:
                try:
                    vol = re.findall(r'\d+', wants)[0]
                    if "percent" in wants or "%" in wants:
                        mixer.music.set_volume(int(vol)/100)
                    else:
                        mixer.music.set_volume(int(vol)/10)
                except Exception as e:
                    print("Please give a value of volume.")
                    speak("Please give a value of volume.")
                finally:
                    continue

        if wants == "exit" or 'bye' in wants:
            speak("Bye Bye! Take Care.")
            break

        elif "sleep" in wants and [x for x in wants if x.isnumeric()] != []:
            if "min" in wants:
                time.sleep(int(re.findall(r'\d+', wants)[0])*60)
            else:
                time.sleep(int(re.findall(r'\d+', wants)[0]))

        elif "wikipedia" in wants:
            try:
                speak("Found this from Wikipedia.")
                # find = wants.split(" ")[0]
                results = wikipedia.summary(
                    wants.replace("wikipedia", ""), sentences=2)
                print(results)
                speak(results)

            except Exception as e:
                print("Couldn't fetch information. Try again.")

        elif "time" in wants:
            timenow = f'Time is {time.strftime("%I:%M%p",time.localtime())}.'
            print(timenow)
            speak(timenow)

        elif "weather" in wants:
            try:
                speak(OpenweatherAPIcalltext())
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
                    speak(item)

            except Exception as e:
                print("Couldn't get news. Try again.")

        elif "date" in wants or "today" in wants:
            todayday = time.strftime("%A", time.localtime())
            todaydate = time.strftime("%d %B", time.localtime())
            daydatewish = f'Today is {todayday} and Date is {todaydate}.'
            print(daydatewish)
            speak(daydatewish)

        elif "open" in wants:

            if "google" in wants:
                webbrowser.open("www.google.com")
            elif "youtube" in wants:
                webbrowser.open("www.youtube.com")
            elif "whatsapp" in wants:
                webbrowser.open("https://web.whatsapp.com")

        elif "play" in wants and "youtube" in wants:
            url = ytfirsturlreturn(wants.replace(
                "play", "").replace("youtube", "").replace("on", ""))
            print(f"Playing {queuedict[url]} on Youtube")
            speak(f"Playing {queuedict[url]} on Youtube")
            webbrowser.open_new_tab(url)

        elif "play" in wants:
            keyword = wants.replace("play", "")
            if keyword == "":
                continue
            print("Please Wait...")
            ytdlMusicPlay(keyword)

        elif "search" in wants and "google" in wants:
            url = f'https://www.google.com/search?q={wants.replace("search ","").replace("google ","").replace(" ","+")}'
            webbrowser.open_new_tab(url)

        elif "search" in wants and "youtube" in wants:
            keyword = wants.replace("search ", "").replace("youtube ", "")
            print(f"Searching for {keyword}")
            url = f'https://www.youtube.com/results?search_query={keyword.replace(" ","+")}'
            webbrowser.open_new_tab(url)

        elif "search" in wants:
            try:
                keyword = wants.replace("search", "")
                print(f"Searching for {keyword}")
                results = wikipedia.summary(keyword, sentences=2)
                if "==" in results:
                    print(wikipedia.summary(keyword))
                    speak(wikipedia.summary(keyword))
                else:
                    print(results)
                    speak(results)
            except Exception as e:
                print(e)

        elif "joke" in wants:
            try:
                url = "https://v2.jokeapi.dev/joke/Any"
                joketext = requests.get(url)
                jokejson = json.loads(joketext.text)
                print(jokejson['joke'])
                speak(jokejson['joke'])
            except Exception as e:
                print("Couldn't get joke. Try again.")

        elif "insult me" in wants:
            try:
                url = "https://evilinsult.com/generate_insult.php"
                insulttext = requests.get(url).text
                print(insulttext)
                speak(insulttext)
            except Exception as e:
                print("Couldn't get Insult. Try again.")

        elif "quote" in wants:
            try:
                quote = quotereturn()[0]
                author = quotereturn()[1]
                print(quote, "  -", author)
                speak(quote)
            except Exception as e:
                print("Couldn't get quote. Try again.")

        elif "say" in wants[0:4] or "speak" in wants[0:6]:
            speak(wants.replace("say", "").replace("speak", ""))

        else:
            speak("Sorry, I can't help with this.")

        time.sleep(2)
