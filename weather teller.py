import datetime
import requests
import json
import win32com.client

def speak(str):
    # from win32com.client import Dispatch
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(str)

if __name__ == '__main__':

    try:
        requests.get("https://www.youtube.com")

    except Exception as e:
        print("Try Checking Your Internet Connection")
        speak("Try Checking Your Internet Connection")
        exit()
    speak("Please enter Your City Id")
    city_id = input("Please Enter Your City id\n")   #Enter Your city id
    api_key = "7e2e1fe81de762f2f2f0177417064b65"  #Enter Your Api Key
    url = f"http://api.openweathermap.org/data/2.5/weather?id={city_id}&appid={api_key}"
    weather = requests.get(url).text

    weather_dict = json.loads(weather)
    coordinates = weather_dict['coord']
    lat = coordinates['lat']
    lon = coordinates['lon']
    # print(lat, lon)

    weather_report = weather_dict['weather']
    weather_report1 = weather_report[0]
    main_type = weather_report1['main']
    description = weather_report1['description']
    main_report = weather_dict['main']
    feels_like = main_report['feels_like']
    max_temp = main_report['temp_max']
    average_temp = main_report['temp']
    min_temp = main_report['temp_min']
    pressure = main_report['pressure']
    humidity = main_report['humidity']
    visibility = weather_dict['visibility']
    city = weather_dict['name']
    timezone = weather_dict['timezone']
    wind = weather_dict['wind']
    speed = wind['speed']
    deg = wind['deg']
    sys = weather_dict['sys']
    sunrise = sys['sunrise']
    sunset = sys['sunset']
    clouds = weather_dict['clouds']
    clouds_percent = clouds['all']
    f = open("oldWeatherReport.txt", "a")
    f.write(f"DATE IS : - {datetime.datetime.now()}\n")
    f.write(f"Your City name is {city}\n")
    f.write(f"maximum temperature is {round((max_temp-273), 2) } celsius\n")
    f.write(f"minimum temperature is {round((min_temp-273), 2) } degree celsius\n")
    f.write(f"average temperature is {round((average_temp-273), 2) } degree celsius\n")
    f.write(f"it feels like {round((feels_like-273), 2) } degree celsius\n")
    f.write(f"humidity was {humidity} %\n")
    f.write(f"Visibility was {visibility}\n")
    f.write(f"Wind Speed was {speed} kilometers per hour at {deg} degrees\n")
    f.write(f"Sunrise at {sunrise} seconds and Sunset at {sunset} seconds\n")
    f.write(f"Clouds were {clouds_percent} percent\n")
    f.write(f"Weather Condition was {main_type}\n\n\n")
    f.write("_____________________________________________________________________________________________________________")
    f.write("\n\n")
    f.close()
    print(f"Your current geographic location is {lat} degree North lattitude and {lon} degree East longitude")
    speak(f"Your current geographic location is {lat} degree North lattitude and {lon} degree East longitude")
    print(f"Your City name is {city}")
    speak(f"Your City name is {city}")
    print("Weather Report for Today")
    speak("Weather Report for Today")
    print("Lets Begin Now With the weather report")
    speak("Lets Begin Now With the weather report")
    print(f"Todays maximum temperature is {round((max_temp-273), 2) } celsius")
    speak(f"Todays maximum temperature is {round((max_temp-273), 2) } degree celsius")
    print(f"Todays minimum temperature is {round((min_temp-273), 2) } degree celsius")
    speak(f"Todays minimum temperature is {round((min_temp-273), 2) } degree celsius")
    print(f"Todays average temperature is {round((average_temp-273), 2) } degree celsius")
    speak(f"Todays average temperature is {round((average_temp-273), 2) } degree celsius")
    print(f"Today it feels like {round((feels_like-273), 2) } degree celsius")
    speak(f"Today it feels like {round((feels_like-273), 2) } degree celsius")
    print(f"Todays humidity is {humidity} %")
    speak(f"Todays humidity is {humidity} percent")
    print(f"Visibility is {visibility}")
    speak(f"Visibility is {visibility}")
    print(f"Wind Speed is {speed} km/hr at {deg} degrees")
    speak(f"Wind Speed is {speed} kilometers per hour at {deg} degrees")
    print(f"Sunrise at {sunrise} seconds")
    speak(f"Sunrise at {sunrise} seconds")
    print(f"Sunset at {sunset} seconds")
    speak(f"Sunset at {sunset} seconds")
    print(f"Clouds will be {clouds_percent} %")
    speak(f"Clouds will be {clouds_percent} percent")
    print(f"Weather Condition will be {main_type}")
    speak(f"Weather Condition will be {main_type}")
    print("Thanks For listening")
    speak("Thanks For listening")
    print("See You Tomorrow ! GoodBye")
    speak("See You Tomorrow  GoodBye")

