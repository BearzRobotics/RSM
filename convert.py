#!/bin/usr/python3

import eng_to_ipa as ipa
# from openpyxl import load_workbook, Workbook
import xlsxwriter as xls
from array import *
from tkinter import *
import pyglet
from tkinter import ttk
from ctypes import byref, create_unicode_buffer, create_string_buffer
import epitran

try:
    from ctypes import windll
except:
    pass

FR_PRIVATE  = 0x10
FR_NOT_ENUM = 0x20

# ipa to Rosquin Empire
def ipaTreo():
    # copies in ipa text and creates an output text
    wipat = ipa_txt.get("1.0", END)
    nipat = ""

    # Creates Look up dictinoary
    ipa_reo_dict = {"m":"a","b":"s","p":"S","v":"d","f":"D","d":"f","t":"F","ð":"g","θ":"G","z":"h","s":"H","n":"j","ɾ":"k","l":"z","ɫ":"x","ʒ":"c","ʃ":"C","j":"b","ɹ":"n","ɡ":"m","k":"M","w":",","ʔ":">","h":"?","i":"1","ɪ":"2","e":"3","ɛ":"4","æ":"5","ə":"6","ɚ":"e","ɜ":"7","ɝ":"r","ʌ":"8","u":"!","ʊ":"@","o":"#","ɔ":"$","ɑ":"%","p":"I","t":"O","k":"P","ɪr":"q","ɛr":"w","ɚ":"e","ɝ":"r","ʊr":"t","oʊ":"W","ɔr":"y","ɔɪ":"E","ɑr":"u","aɪ":"R","aʊ":"T"}
    for  element in range(0, len(wipat)):
        try:
            nipat = nipat + ipa_reo_dict[wipat[element]]
        except:
            pass

    # update Glyphes table
    GlyphesTable_txt.configure(state="normal")
    GlyphesTable_txt.insert("1.0", nipat)
    GlyphesTable_txt.configure(state="disabled")

def engTipa():
    # Text to convert
    in_box = in_txt.get("1.0", END)
    ctxt = ipa.convert(in_box)
    ipa_txt.configure(state="normal")
    ipa_txt.insert('1.0', ctxt)
    ipa_txt.configure(state="disabled")

    ipaTreo()

def excel():
    # Create a workbook
    workbook = xls.Workbook( name_entry.get() + '.xlsx')
    worksheet = workbook.add_worksheet()

    # Start from the first cell.
    row = 0
    col = 0
    worksheet.write(row, col, in_txt.get(1.0, END))
    worksheet.write(row,col+1, ipa_txt.get(1.0, END))
    worksheet.write(row, col+2, GlyphesTable_txt.get(1.0, END))
    
    # Clean up
    workbook.close()

def clear_text():
    name_entry.delete(0, END)
    in_txt.delete(1.0, END)
    ipa_txt.delete(1.0, END)
    GlyphesTable_txt.delete(1.0, END)

def loadfont(fontpath, private=True, enumerable=False):
    '''
    Makes fonts located in file `fontpath` available to the font system.

    `private`     if True, other processes cannot see this font, and this 
                  font will be unloaded when the process dies
    `enumerable`  if True, this font will appear when enumerating fonts

    See https://msdn.microsoft.com/en-us/library/dd183327(VS.85).aspx

    '''
    
    

    if isinstance(fontpath, bytes):
        pathbuf = create_string_buffer(fontpath)
        AddFontResourceEx = windll.gdi32.AddFontResourceExA
    elif isinstance(fontpath, str):
        pathbuf = create_unicode_buffer(fontpath)
        AddFontResourceEx = windll.gdi32.AddFontResourceExW
    else:
        raise TypeError('fontpath must be of type str or unicode')

    flags = (FR_PRIVATE if private else 0) | (FR_NOT_ENUM if not enumerable else 0)
    numFontsAdded = AddFontResourceEx(byref(pathbuf), flags, 0)
    return bool(numFontsAdded)


# Create basic Window structure
win = Tk()
win.title("Bearz English to IPA")
win.geometry('1080x640')

try:
    loadfont("RĔ~O_001_AmEng.IPA_16210.tff")
except:
    pass
# File name label and entry.
name_text = StringVar()
name_label = Label(win, text='File name: ', font=('bold', 14), pady=20)
name_label.grid(row=0, column=0)
name_entry = Entry(win, textvariable=name_text)
name_entry.grid(row=1, column=0)

# Input text label and entry.
in_label = Label(win, text='Input Text: ', font=('bold', 14), pady=20)
in_label.grid(row=2, column=0)
in_txt = Text(win, height=20, width=30)
in_txt.grid(row=3, column=0)

# Output text label and entry.
out_label = Label(win, text='Output Text: ', font=('bold', 14), pady=20)
out_label.grid(row=1, column=3)
ipa_txt = Text(win, height=20, width=30)
ipa_txt.grid(row=3, column=3)
ipa_txt.configure(state="disabled")

# Glyphes
GlyphesTable_label = Label(win, text='Glyphes Table: ', font=('bold', 14), pady=20)
GlyphesTable_label.grid(row=1, column=5)
GlyphesTable_txt = Text(win, height=20, width=50)
GlyphesTable_txt.grid(row=3, column=5)
GlyphesTable_txt.configure(font=("RĔ~O_001_AmEng.IPA_16210", 18), state="disabled")


# Clear text Button
clsbt = Button(win, text="Clear", command=clear_text)
clsbt.grid(row=0, column=2)

# Save Button

sv = Button(win, text="Save", command=excel)
sv.grid(row=0, column=4)

#Convert Button
cb = Button(win, text="Convert", command=engTipa)
cb.grid(row=0, column=5)

win.mainloop()
