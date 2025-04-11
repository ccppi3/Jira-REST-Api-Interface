import tkinter.font as tkFont
import customtkinter as ctk
import webbrowser
# Make links work
def hyperlinkCallback(url,event):
    #print(url)
    webbrowser.open_new_tab(url)
def hyperlinkDoUnderline(ctkLabel,show,event):
    font = ctkLabel.cget("font")
    #print("ctkLabel:",ctkLabel,"show:",show)
    if show:
        font.configure(underline = True)
    else:
        font.configure(underline = False)
    ctkLabel.configure(font=font)
