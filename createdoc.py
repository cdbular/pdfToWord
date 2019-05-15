from docxtpl import DocxTemplate, RichText
from PyPDF2 import PdfFileReader
from tkinter import filedialog
import tkinter
import os
import re
import sys

def checkstring(s):
    letter_flag = False
    number_flag = False
    for i in s:
        if i.isalpha():
            letter_flag = True
        if i.isdigit():
            number_flag = True
    return letter_flag and number_flag

def checkstring_sym(s, sym_str):
    sym_flag = False
    number_flag = False
    for i in s:
        if i.isdigit():
            number_flag = True
    if s.find(sym_str) >0:
        sym_flag=True
       
                
    return sym_flag and number_flag


def delete_line_numbers(input_string):
    #delete line numbers
    question_words=input_string.split(" ")
    list_index=0
    
    for word in question_words:
        #check if the words contains numbers and letters
        if (checkstring(word) == True):      
            #Proceed to replace the numbers with a space
            numbers=getNumbers(word)
            for n in numbers:
                word=word.replace(n," ")
                question_words[list_index]=word

        list_index+=1                              
    print(question_words)
    out_string=' '.join(question_words)
            
    return out_string

def getNumbers(str): 
    array = re.findall(r'[0-9]+', str) 
    return array 
tpl=DocxTemplate('templates/richtext_tpl.docx')




# Build a list of tuples for each file type the file dialog should display
my_filetypes = [('PDF files','*.pdf'), ("All files", "*.*")]

application_window = tkinter.Tk()
# Ask the user to select a single file name.
answer = filedialog.askopenfilename(parent=application_window, initialdir=os.getcwd(), title="Please select a file:", filetypes=my_filetypes)
print(answer)

print(answer)

FILE_PATH = answer



input1 = PdfFileReader(open(FILE_PATH, mode='rb'))
n_pages=input1.getNumPages()
print("document1.pdf has %d pages." % n_pages)
end_of_document=False

rt = RichText("")
page_index=0;
inc_page=True
#n_pages=6

while (end_of_document== False):
    print("Página %d ." % page_index)    
    if (inc_page == True):   
        page=input1.getPage(page_index)
        txt=page.extractText()
        end_of_page=txt.find("\n\n")
        if (end_of_page >0):
            print("End of page found")
        txt=txt[0:end_of_page]
        page_index+=1
        inc_page=False
    


    #Look for last page
    if (txt.find("ERRATA SHEET")>0):
        end_of_document=True

    #Find a question
    print("Looking for Q.\r")
    Q=txt.find("    Q.")
    if (Q > 0):
        question=txt[Q:]
        mark=question.find("?")
        Amark=question.find("   A.")
        if (Amark > 0):
            question=question[0:Amark+1]
            print("Q. found!\r")
            
            
            question=delete_line_numbers(question)
            print(question)
            #add question to word output
            rt.add(question+"\r\n\r\n")
            txt=txt[Q+Amark+6:end_of_page]
           
        else:
            #if this happens is because the end of the question is in the next page
            print("Question continue in next page...")
            partial_q=question
            page=input1.getPage(page_index)
            txt=page.extractText()
        
            Amark=txt.find("    A.")
            if (Amark > 0):
                question=partial_q+txt[0:Amark+1]
                print("Q. found!\r")
                
                question=delete_line_numbers(question)
                print(question)
                rt.add(question+"\r\n\r\n")
                inc_page=True
            
                       
    else:
        print("no more question found in page")
        inc_page=True
    
    if (page_index > n_pages):
       break;


context = {
    'example' : rt,
}

tpl.render(context)
tpl.save('richtext.docx')
