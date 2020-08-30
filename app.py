from tkinter import *
from tkinter import messagebox,filedialog
import sender
import smtplib
import speech_recognition as sr
import os

root=Tk()
root.title("Voice mail system")
root.geometry("800x650")
root.configure(bg='cyan2')

# function for speech
def speak(str):
    from win32com.client import Dispatch
    speak = Dispatch("SAPI.SpVoice")
    speak.speak(str)

def get_body():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        speak("Ok, What's the message.'")
        audio=r.listen(source)
        try:
            voice_data=r.recognize_google(audio)
            txt_box.insert(END, voice_data)
           
        except sr.UnknownValueError:
            speak('sorry, i didnt hear you')
        except sr.RequestError:
            speak('Sorry, my speech service is down...')


def get_id():
    try:
        sender.reciver_id = box.get()
        messagebox.showinfo('E-mail Info',"Email id added successfully..")
        get_body()
    except ValueError:
        messagebox.showerror('E-mail Info','Incorrect email id..')

def send_email():
        sender.body=txt_box.get(1.0, END)
        messagebox.showinfo('E-mail Info',"Message added successfully..")        
        email=f"Subject:Update from TECH-deets.\n\n{sender.body}"
        try:
            server=smtplib.SMTP_SSL("smtp.gmail.com",465)         
            server.login(sender.email_id,sender.password)
            server.sendmail(sender.email_id,sender.reciver_id,email)
            server.quit()
            messagebox.showinfo('E-mail Info',"Email send successfully..")
        except ValueError:
            messagebox.showerror('E-mail Info','Email not send..')


header = Label(root, text='WELCOME TO PYTHON voice E-mail SENDER',bg="cyan4",fg="white",width=50,font=("Times", "16", "bold italic"),borderwidth=5, relief=SUNKEN)
header.pack(pady=10)

label = Label(root, text="Enter sender's E-mail Id",bg='black',fg='white',font=("Times", "14", "bold"),borderwidth=5,relief=SUNKEN)
label.pack(pady=10)

box = Entry(root,font=("Times", "12", "bold"),width=35,borderwidth=5)
box.pack()

button = Button(root, text='Confirm',bg="black",fg="white",font=("Times", "12", "bold"),borderwidth=5,command=get_id)
button.pack(pady=10)

txt_box = Text(root, width=60, height=10, font=("Helvetica",16,"italic"),selectbackground="white",selectforeground="blue")
txt_box.pack(pady=10)

# button = Button(root, text='Insert Image',bg="black",fg="white",font=("Times", "12", "bold"),borderwidth=5,command=insert_img)
# button.pack(pady=10)

button = Button(root, text='SEND Email',bg="black",fg="white",font=("Times", "12", "bold"),borderwidth=5,command=send_email)
button.pack(pady=20)

label = Label(root, text='SANJAY KUMAR (C) 2020',bg='cyan2',font=("Times", "12", "bold"),borderwidth=5)
label.pack(pady=20)

root.mainloop()