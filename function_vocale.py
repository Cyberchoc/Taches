import pyttsx3
import speech_recognition as sr
 
# Initialise la bibliothèque de synthèse vocale
engine = pyttsx3.init()

# Lit la réponse avec la synthèse vocale
def speak(result):
    engine.say(result)
    engine.runAndWait()

def enregistrer_vocale(texte, vocale=True):
    speak(texte)
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print(texte)
        audio = r.listen(source, phrase_time_limit=15)

    # Reconnaissance de la parole
    try:
        reponse = r.recognize_google(audio, language='fr-FR')
        print("Vous avez dit : {}".format(reponse))
        return reponse

    except:
        reponse_sortie = print("Désolé, je n'ai pas compris.")
        speak(reponse_sortie)
    return reponse_sortie