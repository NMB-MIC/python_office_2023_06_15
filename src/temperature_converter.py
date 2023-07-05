# %%
#!pip install tk

# %% [markdown]
# ## Tkinter

# %%
import tkinter as tk
from tkinter import ttk 

def fahrenheit_to_celsius(f):
    """convert fahrenhite to celsius"""
    result = round((f - 32) * 5/9,2)
    return result

# root window
window = tk.Tk()
window.title("Temperature Converter")
window.geometry("300x70")
window.resizable(False,False)

# frame
frame = ttk.Frame(window)

# field options
options = {'padx':5,'pady':5}

# temperature label
temperature_label = ttk.Label(frame,text="Fahrenheit")
temperature_label.grid(column=0,row=0,**options)

# temperatire input box
temperature = tk.StringVar()
temperature_input = ttk.Entry(frame,textvariable=temperature)
temperature_input.grid(column=1,row=0,**options)

def convert_button_clicked():
    """get temperature from entry then convert it"""
    f = float(temperature.get())
    c = fahrenheit_to_celsius(f)
    result_label.config(text=c)

# button
convert_button = ttk.Button(frame,text="Convert")
convert_button.grid(column=2,row=0,**options)
convert_button.configure(command=convert_button_clicked)

#result label
result_label = ttk.Label(frame,text="please input")
result_label.grid(columnspan=3,row=1,**options)

# add frame
frame.grid(pady=10,padx=10)

window.mainloop() #run


