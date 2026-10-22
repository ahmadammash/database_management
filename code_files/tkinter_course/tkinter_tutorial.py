# In[]:
# Import the required libraries
from tkinter import *
from tkinter import ttk

# In[]: CHAPTER 2
root = Tk()
button = ttk.Button(root, text="Click Me")  # create a button
button.pack()  # place the botton on the root window
button["text"] = "Press Me"  # change the text of the button
button.config(text="Push Me")  # another way to change the text of the button
print(button.config())  # print the config of the button
ttk.Label(
    root, text="Hello tkinter"
).pack()  # add a label widget without needing to assign it to a variable
root.mainloop()
"""
Geometry managers: pack, grid, place.

Pack is the simplest while grid uses a grid layout. Place is for placing widgets precisely in specific locations. It is required
to use the same geometry within the same master to ensure consistency
"""

"""
A common design practice is to have the application written as a class. All tk objects widgets for the GUI are created in the
__init__ constructor method
"""
# In[]: CHAPTER 3.1
root = Tk()
label = ttk.Label(root, text="Hello, Tkinter!")
label.pack()
label.config(
    text="Howdy, Tkinter! It's been a while since we last met. Great to see you again!"
)
label.config(
    wraplength=400,
    justify=CENTER,
    foreground="blue",
    background="yellow",
    font=("Courier", 24, "bold underline"),
)
label.config(text="Howdy, Tkinter!")
logo = PhotoImage(
    file="C:\\\\Users\\\\medaamm\\\\OneDrive - University of Leeds\\\\UOL Data Manager\\\\database_management\\\\code_files\\\\tkinter_course\\\\Exercise Files\\\\Ch03\\\\python_logo.gif"
)
label.config(image=logo)


""" A label has the text property to configure and also has an image property associated with it.
We can control which one is displayed by changing the compound property. """

label.config(compound="text")
label.config(compound="center")  # places the text on the center of the image
label.config(compound="left")  # image to the left of the text

# do the below to prevent losing the image if the image variable loses scope
label.img = logo  # creates an image variable within the label object
label.config(image=label.img)


root.mainloop()

# In[]: CHAPTER 3.2
root = Tk()
button = ttk.Button(root, text="Click Me")
button.pack()


# define the callback function to declare what is printed when a button is clicked
def callback():
    print("Clicked!")


# assign the callback function to the command property of the button for it to run when the button is clicked
button.config(command=callback)

# to simulate a button click
button.invoke()

# set the state of a widget
button.state(["disabled"])

# check the current state of the button
print(
    button.instate(["disabled"])
)  # this asks the question whether the button is disabled and it will return True or False

# enable the button
button.state(["!disabled"])

print(button.instate(["disabled"]))

# add an image to the button
logo = PhotoImage(
    file="C:\\Users\\medaamm\\OneDrive - University of Leeds\\UOL Data Manager\\database_management\\code_files\\tkinter_course\\Exercise Files\\Ch03\\python_logo.gif"
)
button.img = logo
button.config(image=button.img)
button.config(compound="left")
small_logo = logo.subsample(5, 5)  # to resize the image
button.config(image=small_logo)
root.mainloop()

# In[]: CHAPTER 3.3
root = Tk()
checkbutton = ttk.Checkbutton(root, text="SPAM?")
checkbutton.pack()
spam = StringVar()  # create a string variable
spam.set("SPAM!")  # assign the variable value
checkbutton.config(
    variable=spam, onvalue="SPAM Please!", offvalue="Boo SPAM"
)  # by default the checkbutton uses 1 when selected and 0 when not selected
print(spam.get())

breakfast = StringVar()
ttk.Radiobutton(root, text="SPAM", variable=breakfast, value="SPAM").pack()
ttk.Radiobutton(root, text="Eggs", variable=breakfast, value="Eggs").pack()
ttk.Radiobutton(root, text="Sausage", variable=breakfast, value="Sausage").pack()
ttk.Radiobutton(root, text="SPAM", variable=breakfast, value="SPAM").pack()

""" Configuring the textvariable property allows us to change the value of that text by just modifying the value of the variable"""
checkbutton.config(textvariable=breakfast)

root.mainloop()
# In[]: CHAPTER 3.4
root = Tk()
entry = ttk.Entry(root, width=30)
entry.pack()  # the width property controls the size of the entry field (as a number of characters)
entry.delete(0, 1)
root.mainloop()

# In[]: CHAPTER 3.5
root = Tk()
month = StringVar()
combobox = ttk.Combobox(root, textvariable=month)
combobox.pack()
month.set("Jan")
combobox.config(
    values=(
        "Jan",
        "Feb",
        "Mar",
        "Apr",
        "May",
        "Jun",
        "Jul",
        "Aug",
        "Sep",
        "Oct",
        "Nov",
        "Dec",
    )
)
year = StringVar()
year.set("2013")
Spinbox(root, from_=1990, to=2014, textvariable=year).pack()
root.mainloop()
year.get()

# In[]: CHAPTER 3.6
root = Tk()
progressbar = ttk.Progressbar(root, orient="horizontal", length=200)
progressbar.pack()
progressbar.config(mode="indeterminate")
progressbar.start()
progressbar.stop()
progressbar.config(mode="determinate", maximum=11.0, value=4.2)
progressbar.config(value=8.0)
progressbar.step(2)
value = DoubleVar()
progressbar.config(variable=value)
scale = ttk.Scale(
    root, orient="horizontal", length=200, variable=value, from_=0.0, to=11.0
).pack()
value.set(4.2)
root.mainloop()

# In[]: CHAPTER 4.1
root = Tk()
"""
Frames are useful to divide the GUI into sections we can work with separately - to organise the GUI. Each frame can also use a different geometry manager to organise the widgets within
that frame.
"""
frame = ttk.Frame(root)
frame.pack()
frame.config(height=200, width=300, relief="ridge")
ttk.Button(
    frame, text="Click Me"
).grid()  # using the grid manager rather than the pack manager here and for this frame
frame.config(padding=(30, 15))
"""
Accordingly, any other widgets added to this frame later will also need to use the grid geometry manager
"""
ttk.Labelframe(root, height=200, width=300, text="My Frame").pack()
root.mainloop()

# In[]: CHAPTER 4.2
root = Tk()
window = Toplevel(root)
window.title("New Window")
window.lift(root)
window.lower()
window.state("zoomed")
window.state("withdrawn")
window.geometry("150x500+50+150")
root.mainloop()

# In[]: CHAPTER 4.3
root = Tk()
panedWindow = ttk.Panedwindow(root, orient="horizontal")
panedWindow.pack(fill=BOTH, expand=True)
frame1 = ttk.Frame(panedWindow, width=100, height=300, relief="sunken")
frame2 = ttk.Frame(panedWindow, width=400, height=400, relief="sunken")
frame3 = ttk.Frame(panedWindow, width=400, height=50, relief="sunken")
panedWindow.add(
    frame1, weight=1
)  # weight specifies how much the frame will be scaled when the pane window is resized.
panedWindow.add(frame2, weight=4)
panedWindow.insert(1, frame3)
panedWindow.forget(1)
"""
the paned window uses the largest height of the two frames
"""
root.mainloop()

# In[]: CHAPTER 4.4
root = Tk()
# Notebook is like a browser tab
notebook = ttk.Notebook(root)
notebook.pack()
frame1 = ttk.Frame(notebook)
frame2 = ttk.Frame(notebook)
notebook.add(frame1, text="One")
notebook.add(frame2, text="Two")
button = ttk.Button(frame1, text="Click Me").pack()
frame3 = ttk.Frame(notebook)
notebook.insert(1, frame3, text="Three")
notebook.forget(1)
notebook.add(frame3, text="Three")
print(notebook.select())
root.mainloop()


# In[]: CHAPTER 5.1
root = Tk()
text = Text(root, width=40, height=10)
text.pack()
text.config(wrap="word")
text.delete("1.0")
root.mainloop()
