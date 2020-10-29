import PyPDF2, os, docx
from gtts import gTTS

class AudioBook():
    """
    Converts the given PDF file, word file or text file into audio(mp3) format.
    """
    task_executed = False

    def __init__(self, stream:str, audio_file_name:str) -> None:
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
        if self.audiofile.endswith(".mp3"):
            pass
        else:
            self.audiofile = self.audiofile + ".mp3"

    def isValid(self, path:str):
        """
        Checks whether the given path is valid or not.
        """
        if os.path.exists(path):
            pass
        else:
            print("The given path is invalid.")
            quit()

    def text_to_speech_PDF(self):
        """
        Converts the given PDF text to speech.
        """
        try:
            print("Reading the book...")
            book = PyPDF2.PdfFileReader(self.stream)
            text = ""
            pages = book.getNumPages()
            for i in range(pages):
                page_text = book.getPage(i).extractText()
                text += page_text

            print("Converting to audio... This may take a few minutes depending on your internet speed and size of the book...")
            speech = gTTS(text, lang='en')
            speech.save(self.audiofile)

        except Exception as e:
            print(e)

        else:
            print("Your audiobook is ready!")
            AudioBook.task_executed = True

    def text_to_speech_TXT(self):
        """
        Converts the text file into speech.
        """
        try:
            print("Reading the book...")
            with open(self.stream) as f:
                fc = f.read()

            print("Converting to audio... This may take a few minutes depending on your internet speed and size of the book...")    
            speech = gTTS(fc, lang='en')
            speech.save(self.audiofile)
        
        except Exception as e:
            print(e)

        else:
            print("Your audiobook is ready!")
            AudioBook.task_executed = True

    def text_to_speech_DOCX(self):
        """
        Converts the given word file into audio.
        """
        try:
            print("Reading the book...")
            doc = docx.Document(self.stream)
            text = ""
            for para in doc.paragraphs:
                text += para.text + "\n"

            print("Converting to audio... This may take a few minutes depending on your internet speed and size of the book...")
            speech = gTTS(text, lang='en')
            speech.save(self.audiofile)
        
        except Exception as e:
            print(e)

        else:
            print("Your audiobook is ready!")
            AudioBook.task_executed = True

    def listen(self):
        """
        Opens the audiobook when ready!
        """
        if AudioBook.task_executed:
            os.startfile(self.audiofile)
        else:
            quit()

    def main(self):
        """
        Calls the isValid and text_to_speech functions in order to create the audiobook.
        """
        self.isValid(self.stream)

        if self.form == "pdf":
            self.text_to_speech_PDF()
        elif self.form == "txt":
            self.text_to_speech_TXT()
        elif self.form == "docx":
            self.text_to_speech_DOCX()

        start = input("Wanna listen the audiobook now?(Y/N) ")
        if start.lower() == "y":
            self.listen()


if __name__ == "__main__":
    print("Welcome to the Python Audio Book Maker!")
    source = input("Enter the full path of the text or PDF file you want to turn into audiobook: ")
    destination = input("Enter the name of the audiobook to save: ")
    audiobook = AudioBook(source, destination)
    audiobook.main()