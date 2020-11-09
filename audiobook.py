# pypdf2 for extracting text form pdf, docx for word documents, gtts for text to speech conversion and os for checking if the path is valid and starting the file
from docx import Document
from gtts import gTTS
from playsound import playsound
from PyPDF2 import PdfFileReader
from os import path as ospath
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
        if not ospath.exists(self.stream):
            print("The given path is invalid.")
            quit()

    def listen(self):
        """
        Opens the audiobook when ready.
        """
        start = input("Do you want to listen to the audiobook now?(Y/N) ")
        if start.lower() == "y":
            startfile(self.audiofile)


    def text_to_speech_PDF(self):
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

            print("Converting to audio... This may take a while depending on your internet speed and size of the book...")
            # converting the text to audio and saving the audio file
            speech = gTTS(text, lang='en')
            speech.save(self.audiofile)

        except Exception as e:
            playsound(r"C:\Users\Lenovo\Documents\Python Codes\Awkward Cricket.mp3")
            print(e)

        else:
            playsound(r"C:\Users\Lenovo\Documents\Python Codes\anime wow.mp3")
            print("Your audiobook is ready!")
            self.listen()

    def text_to_speech_TXT(self):
        """
        Converts the text file into speech.
        """
        try:
            print("Reading the book...")
            with open(self.stream) as f:
                fc = f.read()

            print("Converting to audio... This may take a while depending on your internet speed and size of the book...")    
            speech = gTTS(fc, lang='en')
            speech.save(self.audiofile)
        
        except Exception as e:
            playsound(r"C:\Users\Lenovo\Documents\Python Codes\Awkward Cricket.mp3")
            print(e)

        else:
            playsound(r"C:\Users\Lenovo\Documents\Python Codes\anime wow.mp3")
            print("Your audiobook is ready!")
            self.listen()

    def text_to_speech_DOCX(self):
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

            print("Converting to audio... This may take a while depending on your internet speed and size of the book...")
            speech = gTTS(text, lang='en')
            speech.save(self.audiofile)
            
        except Exception as e:
            playsound(r"C:\Users\Lenovo\Documents\Python Codes\Awkward Cricket.mp3")
            print(e)

        else:
            playsound(r"C:\Users\Lenovo\Documents\Python Codes\anime wow.mp3")
            self.listen()


    def create_audiobook(self):
        """
        Calls the isValid and text_to_speech functions in order to create the audiobook.
        """
        self.isValid()

        if self.form == "pdf":
            self.text_to_speech_PDF()
        elif self.form == "txt":
            self.text_to_speech_TXT()
        elif self.form == "docx":
            self.text_to_speech_DOCX()


if __name__ == "__main__":
    print("Welcome to the Python Audio Book Maker!")
    source = input("Enter the full path of the PDF/text/docx file you want to turn into audiobook: ")
    destination = input("Enter the name of the audiobook to save: ")
    ab = AudioBook(source, destination)
    ab.create_audiobook()
