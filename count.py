import docx
import re
import os
import csv
from pypdf import PdfReader
import customtkinter as c
from tkinter import filedialog


# Link to button which uploads file
def upload_file():
    global filename
    global file_extension
    filename = filedialog.askopenfilename()
    file_name, file_extension = os.path.splitext(filename)


# Deletes entry boxes
def clear():
    entry1.delete(0, len(str(entry1.get)))
    entry2.delete(0, len(str(entry2.get)))


def word_count():
    global filename
    global file_extension
    search = entry1.get().strip()
    if not search:
        counter = 0
    else:
        # Opens file and returns count of word
        counter = 0
        try:
            match file_extension:
                case ".docx":
                    file = docx.Document(filename)
                    for paragraph in file.paragraphs:
                        line = paragraph.text
                        if word := re.findall(f"\\b{search}", line, re.IGNORECASE):
                            counter += len(word)
                case ".csv":
                    with open(filename) as csvfile:
                        csvfile = csv.reader(csvfile)
                        for row in csvfile:
                            for i in range(len(row)):
                                if word := re.findall(f"\\b{search}", row[i], re.IGNORECASE):
                                    counter += len(word)
                case ".pdf":
                    reader = PdfReader(filename)
                    for i in range(len(reader.pages)):
                        page = reader.pages[i]
                        text_on_page = page.extract_text()
                        if word := re.findall(f"\\b{search}", text_on_page, re.IGNORECASE):
                            counter += len(word)
                case _:
                    with open(filename) as otherfile:
                        file_lines = otherfile.readlines()
                        for file_line in file_lines:
                            if word := re.findall(f"\\b{search}", file_line, re.IGNORECASE):
                                counter += len(word)
        except (NameError, UnicodeDecodeError):
            pass
        
    # Clears count entry box and returns the count of the word to it
    entry2.delete(0, len(str(entry2.get)))
    entry2.insert(0, counter)

# Sets appearance
c.set_appearance_mode("dark")
c.set_default_color_theme("dark-blue")

# Defines size
root = c.CTk()
root.geometry("320x200")
root.minsize(width=320, height=200)
root.maxsize(width=320, height=200)
root.title("Search Engine")

# project Lebel
label = c.CTkLabel(master=root, text="Word Counter", font=("Arial", 18))
label.grid(row=0, pady=8, padx=10, columnspan=2 , sticky="NSEW")

# Sets buttons
button_upload = c.CTkButton(master=root, text="Upload Document", command=upload_file)
button_upload.grid(row=1, columnspan=2, pady=4, padx=6)

entry1 = c.CTkEntry(master=root, placeholder_text="Enter word")
entry1.grid(row=2, column = 0, pady=4, padx=6)

entry2 = c.CTkEntry(master=root, placeholder_text="Count")
entry2.grid(row=2, column=1, pady=4, padx=6)


button_count = c.CTkButton(master=root, text="Count", command=word_count)
button_count.grid(row=4, column=0, pady=6, padx=8)

button_clear = c.CTkButton(master=root, text="Clear", command=clear)
button_clear.grid(row=4, column=1, pady=6, padx=8)


root.mainloop()