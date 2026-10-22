#!/usr/bin/python3
# keyboard.py by Barron Stone
# This is an exercise file from Python GUI Development with Tkinter on lynda.com

from tkinter import *


def key_press(event):
    print(f"type: {event.type}")
    print(f"widget: {event.widget}")
    print(f"char: {event.char}")
    print(f"keysym: {event.keysym}")
    print(f"keycode: {event.keycode}\n")


def shortcut(action):
    print(action)


root = Tk()

root.bind("<Control-c>", lambda e: shortcut("Copy"))
root.bind("<Control-v>", lambda e: shortcut("Paste"))

root.mainloop()
