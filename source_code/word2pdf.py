from tkinter import filedialog, messagebox as mbox
import os
import win32com.client
import subprocess

# This program won't work on linux or mac os!

current_user = os.getlogin()  # GET THE USERNAME
print(f"Username: {current_user}\n")


def open_files(username):
    """
    A function to get the pdf files filepaths.

    :param username: The user's username. (self generated in the program)
    :return: A list of the filepaths which has been iterated through and
        converted the foward slashed to backward slashes.
    """

    file_paths = []
    # ASK THE USER FOR WHICH PDF FILES TO SELECT
    filenames = filedialog.askopenfilenames(title='Open PDFs', initialdir=f"C:\\Users\\{username}\\Downloads",
                                            filetypes=(("PDF", "*.pdf"), ("All files", "*.*")))

    # SIMPLE FOR LOOP TO REPLACE THE "/" TO "\" using os.path.normcase method.
    for filename_path in filenames:
        filepath = os.path.normcase(filename_path)
        file_paths.append(filepath)

    return file_paths


# INPUT/OUTPUT PATH
pdf_paths = open_files(current_user)
output_path = os.path.normcase(filedialog.askdirectory())

# CHECK THE PDF PATHS TO MAKE SURE THEY DID NOT PRESS CANCEL ON THE FILES SELECTION
if not pdf_paths:
    no_files = True
else:
    no_files = False
    word = win32com.client.Dispatch("Word.Application")
    word.visible = 0  # CHANGE TO 1 IF YOU WANT TO SEE WORD APPLICATION RUNNING AND ALL MESSAGES OR WARNINGS SHOWN BY
    # WORD
    print("Conversion:")

    # GET FILE NAME AND NORMALIZED PATH
    for pdf_path in pdf_paths:

        filename = pdf_path.split('\\')[-1]
        in_file = os.path.abspath(pdf_path)

        # CONVERT PDF TO DOCX AND SAVE IT ON THE OUTPUT PATH WITH THE SAME INPUT FILE NAME
        wb = word.Documents.Open(in_file)
        out_file = os.path.abspath(output_path + '\\' + filename[0:-4] + ".docx")

        print(f"{in_file}     -->      {out_file}")

        wb.SaveAs2(out_file, FileFormat=16)
        wb.Close()

    word.Quit()

    # OPEN THE OUTPUT PATH IN FILE EXPLORER AND SHOW IN A MESSAGE BOX THAT THE CONVERSION IS SUCCESSFUL
    subprocess.Popen(f'explorer "{output_path}')
    mbox.showinfo("PDF files converted!", "The pdf files are successfully converted to word documents!!!")

if no_files:
    # IF NO FILES IS TRUE THEN SHOWS THIS MESSAGE BOX ABOUT TO SELECT SOME FILES
    mbox.showinfo("No files selected", "Please select some files!!!")
else:
    print("Done!!!")
