# pypdf2 for extracting text form pdf, docx for word documents, gtts for text to speech conversion and os for checking if the path is valid and starting the file
from docx import Document
import pyttsx3
from PyPDF2 import PdfFileReader
from os.path import exists
from os import startfile


class AudioBook():
    """
    Converts the given PDF file, word file or text file into audio(mp3) format.
    """

    def __init__(self, stream:str, audio_file_name:str) -> None:
        """Constructor"""
        # manipulating and checking inputs
        
        self.stream = stream
        if self.stream.endswith(".pdf"):
            self.form = "pdf"
        elif self.stream.endswith(".txt"):
            self.form = "txt"
        elif self.stream.endswith(".docx"):
            self.form = "docx"
        else:
            print("Unsupported file type!")
            quit()

        self.audiofile = audio_file_name
        if not self.audiofile.endswith(".mp3"):
            self.audiofile += ".mp3"


    def isValid(self):
        """
        Checks whether the given path is valid or not.
        """
        if not exists(self.stream):
            print("The given path is invalid.")
            quit()

    def listen(self):
        """
        Opens the audiobook when ready.
        """
        start = input("Do you want to listen to the audiobook now?(Y/N) ")
        if start.lower() == "y":
            startfile(self.audiofile)

    def convert_to_audio(self, text:str, voice:str="male"):
        """
        Converts the given string into audio.
        """
        engine = pyttsx3.init()
        if voice == "female":
            engine.setProperty('voice', r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_ZIRA_11.0")
        engine.save_to_file(text, self.audiofile)
        engine.runAndWait()


    def text_to_speech_PDF(self, voice:str="male"):
        """
        Converts the given PDF text to speech.
        """
        try:
            print("Reading the book...")
            # making a book object and extracting text from the pdf
            book = PdfFileReader(self.stream)
            text = ""
            for i in range(book.getNumPages()):
                page_text = book.getPage(i).extractText()
                text += page_text

            print("Converting to audio... This may take a while depending on the size of the book...")
            # converting the text to audio and saving the audio file
            self.convert_to_audio(text, voice)

        except Exception as e:
            print(e)

        else:
            print("Your audiobook is ready!")
            self.listen()

    def text_to_speech_TXT(self, voice:str="male"):
        """
        Converts the text file into speech.
        """
        try:
            print("Reading the book...")
            with open(self.stream) as f:
                fc = f.read()

            print("Converting to audio... This may take a while depending the size of the book...")    
            self.convert_to_audio(fc, voice)
        
        except Exception as e:
            print(e)

        else:
            print("Your audiobook is ready!")
            self.listen()

    def text_to_speech_DOCX(self, voice:str="male"):
        """
        Converts the given word file into audio.
        """
        try:
            print("Reading the book...")
            # making a doc object and extracting text from the docx file
            doc = Document(self.stream)
            text = ""
            for para in doc.paragraphs:
                text += para.text + "\n"

            print("Converting to audio... This may take a while depending on the size of the book...")
            self.convert_to_audio(text, voice)
            
        except Exception as e:
            print(e)

        else:
            self.listen()


    def create_audiobook(self, voice):
        """
        Calls the isValid and text_to_speech functions in order to create the audiobook.
        """
        self.isValid()

        if self.form == "pdf":
            self.text_to_speech_PDF(voice)
        elif self.form == "txt":
            self.text_to_speech_TXT(voice)
        elif self.form == "docx":
            self.text_to_speech_DOCX(voice)


if __name__ == "__main__":
    print("Welcome to the Python Audio Book Maker!")

    source = input("Enter the full path of the PDF/text/docx file you want to turn into audiobook: ")
    destination = input("Enter the name of the audiobook to save: ")
    voice = input("In which audio do you want the audiobook to be?(male/female) :")

    if voice.lower() not in ["male", "female"]:
        print("You entered the wrong audio! Therefore, defaulting to male voice.")
        ab = AudioBook(source, destination)
        ab.create_audiobook("male")

    else:
        ab = AudioBook(source, destination)
        ab.create_audiobook(voice)
