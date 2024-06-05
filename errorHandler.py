import customtkinter
from tkinter import *

def createErrorWindow(errors):
    window = customtkinter.CTkToplevel()
    window.geometry("300x150")
    window.title("Errors Found")

    windowTitle = customtkinter.CTkLabel(window, text="Please fix the following and submit again.")
    windowTitle.grid(row=0, column=0, padx=0, pady=5, ipady=0)
    window.grid_rowconfigure(1, weight=0)
    window.grid_columnconfigure(0, weight=1)   
    rowNum = 1

    for x in errors:
        label = customtkinter.CTkLabel(window, text="- " + x)
        label.grid(row=rowNum, column=0, padx=(25,0), pady=0, ipady=0, sticky="w")
        rowNum = rowNum + 1