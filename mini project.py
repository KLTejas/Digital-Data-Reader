from tkinter import *
import shutil         
import os
import easygui
from tkinter import filedialog
from tkinter import messagebox as mb
from PIL import ImageTk, Image
import speech_recognition as sr
import pyttsx3
from fpdf import FPDF
from pathlib import Path
from sys import argv
from PyPDF4.pdf import PdfFileReader, PdfFileWriter
from win32com import client
from tkPDFViewer import tkPDFViewer as pdf
from PyPDF2 import PdfFileMerger
from pdf2docx import Converter
from docx2pdf import convert
from PyPDF2 import PdfWriter, PdfReader
from pdf2image import convert_from_path
import io    
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter




# open a file box window 
# when we want to select a file
def open_window():
    read=easygui.fileopenbox()
    return read
# open file function
def open_file():
    string = open_window()
    try:
        os.startfile(string)
    except:
        mb.showinfo('confirmation', "File not found!")
# copy file function
def copy_file():
    source1 = open_window()
    destination1=filedialog.askdirectory()
    shutil.copy(source1,destination1)
    mb.showinfo('confirmation', "File Copied !")
# delete file function
def delete_file():
    del_file = open_window()
    if os.path.exists(del_file):
        os.remove(del_file)             
    else:
        mb.showinfo('confirmation', "File not found !")
# rename file function
def rename_file():
    chosenFile = open_window()
    path1 = os.path.dirname(chosenFile)
    extension=os.path.splitext(chosenFile)[1]
    print("Enter new name for the chosen file")
    newName=input()
    path = os.path.join(path1, newName+extension)
    print(path)
    os.rename(chosenFile,path) 
    mb.showinfo('confirmation', "File Renamed !")
# move file function
def move_file():
    source = open_window()
    destination =filedialog.askdirectory()
    if(source==destination):
        mb.showinfo('confirmation', "Source and destination are same")
    else:
        shutil.move(source, destination)  
        mb.showinfo('confirmation', "File Moved !")
# function to make a new folder
def make_folder():
    newFolderPath = filedialog.askdirectory()
    print("Enter name of new folder")
    newFolder=input()
    path = os.path.join(newFolderPath, newFolder)  
    os.mkdir(path)
    mb.showinfo('confirmation', "Folder created !")
# function to remove a folder
def remove_folder():
    delFolder = filedialog.askdirectory()
    os.rmdir(delFolder)
    mb.showinfo('confirmation', "Folder Deleted !")
# function to list all the files in folder
def list_files():
    folderList = filedialog.askdirectory()
    sortlist=sorted(os.listdir(folderList))       
    i=0
    print("Files in ", folderList, "folder are:")
    while(i<len(sortlist)):
        print(sortlist[i]+'\n')
        i+=1
def text_speech():
 
    # Initialize the recognizer
    r = sr.Recognizer()
 
    # Function to convert text to
    # speech
    def SpeakText(command):

        # Initialize the engine
        engine = pyttsx3.init()
        engine.say(command)
        engine.runAndWait()


    # Loop infinitely for user to
    # speak

    while(1):   

        # Exception handling to handle
        # exceptions at the runtime
        try:

            # use the microphone as source for input.
            with sr.Microphone() as source2:

                # wait for a second to let the recognizer
                # adjust the energy threshold based on
                # the surrounding noise level
                r.adjust_for_ambient_noise(source2, duration=0.2)

                #listens for the user's input
                audio2 = r.listen(source2)

                # Using google to recognize audio
                MyText = r.recognize_google(audio2)
                MyText = MyText.lower()

                print("Did you say "+MyText)
#                 SpeakText(MyText)

        except sr.RequestError as e:
            print("Could not request results; {0}".format(e))

        except sr.UnknownValueError:
            print("unknown error occured")

def text_pdf():
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size = 15)
    f = open("C:\\Users\\hrith\\OneDrive\\Desktop\\text_pdf.txt", "r")
    for x in f:
        pdf.cell(200, 10, txt = x, ln = 1, align = 'C')
    pdf.output("myfile.pdf")
    
def pdf_merge():
    #Create and instance of PdfFileMerger() class
    merger = PdfFileMerger()
    #Create a list with PDF file names
    path_to_files = r'pdf_files/'

    #Get the file names in the directory
    for root, dirs, file_names in os.walk(path_to_files):
        #Iterate over the list of file names
        for file_name in file_names:
            #Append PDF files
            merger.append(path_to_files + file_name)

    #Write out the merged PDF
    merger.write("merge.pdf")
    merger.close()
    

# class PdfSplitter:
#     def _init_(self, input_file: str):
#         """
#         Set input PDF file
#         @param input_file: filename of the input PDF file
#         """
#         self.input_file = Path(input_file)
#         self.output_directory: Path or None = None

#         self.input_pdf = self.__open_pdf()

#     def __open_pdf(self) -> PdfFileReader:
#         """
#         Open PDF as bytes
#         @return: 'PdfFileReader' object
#         """
#         return PdfFileReader(open(self.input_file, mode="rb"))

#     def split(self, output_directory: str or None) -> None:
#         """
#         Split PDF into separate PDFs, one file per page
#         @param output_directory: output directory to save files
#         @return: none
#         """
#         if not output_directory:
#             output_directory = "results"
#         self.output_directory = Path(output_directory)
#         self.output_directory.mkdir(exist_ok=True, parents=True)

#         pages = self.input_pdf.getNumPages()
#         for page in range(pages):
#             # FIXME: Library bug? Stream closes:
#             #  https://www.  reddit.com/r/learnpython/comments/58kdfj/pypdf2_pdffilewriter_has_no_attribute_stream/
#             #  Currently fxd with this dirty re-open method, fine for me lol
#             self.input_pdf = self.__open_pdf()
#             output_pdf = PdfFileWriter()
#             output_pdf.addPage(self.input_pdf.getPage(pageNumber=page))
#             with open(
#                 self.output_directory / f"{self.input_file.stem}_{page}.pdf", mode="wb"
#             ) as out_file:
#                 output_pdf.write(out_file)


# if _name_ == "_main_":
#     splitter = PdfSplitter(input_file=argv[1])
#     splitter.split(output_directory=argv[2] if len(argv) >= 3 else None)

def pdfview():
    #from tkinter import *
    #from tkinter import filedialog
    #from tkPDFViewer import tkPDFViewer as pdf
    rooT=Tk()
    rooT.geometry("600x450+500+100")
    rooT.title("basic pdf viewer")
    rooT.config(bg="white")
    
    def browseFile():
        filename=filedialog.askopenfilename(initialdir=os.getcwd(),title="select the file",filetypes=(("PDF Files (.pdf)",".PDF"),("PDF Files (.pdf)",".pdf")))
        v1=pdf.ShowPdf()
        v2=v1.pdf_view(root,pdf_location=open(filename,"r"),width=70,height=100)
        v2.pack()
        
        
    
    btn1=Button(rooT,text="Open the file",command=browseFile,width=50,font="arial 12",bd=4)
    btn1.pack(padx=100,pady=50)
    rooT.mainloop()

def xltopdf():
    # Import Module
    #from win32com import client

    # Open Microsoft Excel
    excel = client.Dispatch("Excel.Application")

    # Read Excel File
    sheets = excel.Workbooks.Open(r'C:\Users\hrith\OneDrive\Desktop\4th Sem\Pending_40_New Horizon College of Engineering.xlsx')
    work_sheets = sheets.Worksheets[0]

    # Convert into PDF File
    work_sheets.ExportAsFixedFormat(0,r'C:\Users\hrith\OneDrive\Desktop\pdf_converted.pdf')
    
def pdf_docx():
    pdf_file = 'merged_all_pages.pdf'
    docx_file = 'sample.docx'
    cv = Converter(pdf_file)
    cv.convert(docx_file)
    cv.close()

def docx_pdf():
    docx_file = 'project.docx'
    pdf_file = 'merged_all_pages.pdf'
    convert(docx_file)
    
def pdf_crop():   
    reader = PdfReader("id.pdf")
    page = reader.pages[0]
    print(page.cropbox.upper_right)
    writer = PdfWriter()

    for page in reader.pages:
        page.cropbox.upper_left = (100,200)
        page.cropbox.lower_right = (300,400)
        writer.add_page(page) 

    with open('result.pdf','wb') as fp:
        writer.write(fp) 
        
def pdf_edit():
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    can.setFillColorRGB(1, 0, 0)
    can.setFont("Times-Roman", 14)
    can.drawString(72, 655, "Hello from Python")
    can.save()
    packet.seek(0)
    new_pdf = PdfFileReader(packet)
    existing_pdf = PdfFileReader(open("C:\\Users\\hrith\\OneDrive\\Desktop\\4th Sem\\Fee payment reciept.pdf", "rb"))
    output = PdfFileWriter()
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)
    outputStream = open("name123.pdf", "wb")
    output.write(outputStream)
    outputStream.close()



root= Tk()
# creating a canvas to insert image
canv = Canvas(root, width=500, height=420, bg='white')
canv.grid(row=0, column=6)
root.title("Digital data reader")
# creating label and buttons to perform operations
Label(root, text="File Options", font=("Helvetica", 16), fg="blue").grid(row = 5, column = 5)
Button(root, text = "Open a File", command = open_file).grid(row=15, column =2)
Button(root, text = "Copy a File", command = copy_file).grid(row = 25, column = 2)
Button(root, text = "Delete a File", command = delete_file).grid(row = 35, column = 2)
Button(root, text = "Rename a File", command = rename_file).grid(row = 45, column = 2)
Button(root, text = "Move a File", command = move_file).grid(row = 55, column =2)
Button(root, text = "Make a Folder", command = make_folder).grid(row = 15, column = 3)
Button(root, text = "Remove a Folder", command = remove_folder).grid(row = 25, column =3)
Button(root, text = "List all Files in Directory", command = list_files).grid(row = 35,column = 3)
Button(root, text = "Text to audio", command = text_speech).grid(row = 45,column = 3)
Button(root, text = "Text to pdf", command = text_pdf).grid(row = 55,column = 3)
Button(root, text = "pdf splitter").grid(row = 15,column = 4)
Button(root, text = "pdf merger", command = pdf_merge).grid(row = 25,column = 4)
Button(root, text = "Excel to pdf", command = xltopdf).grid(row = 35,column = 4)
Button(root, text = "pdf to docx", command = pdf_docx).grid(row = 45,column = 4)
Button(root, text = "docx to pdf", command = docx_pdf).grid(row = 55,column = 4)
Button(root, text = "pdf crop", command = pdf_crop).grid(row = 15,column = 5)
Button(root, text = "pdf editor",).grid(row = 25,column = 5)

root.mainloop()
