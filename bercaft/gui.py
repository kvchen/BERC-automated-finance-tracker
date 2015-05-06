import logging

from threading import Thread, Lock
# from multiprocessing import Process, Lock

import tkinter as tk
import tkinter.font
from tkinter import messagebox as tkmb


logger = logging.getLogger('root')



class DispatchGUI(tk.Tk):
    def __init__(self, dispatch):
        tk.Tk.__init__(self)

        self.dispatch = dispatch
        self.wm_protocol("WM_DELETE_WINDOW", self.destroy)

        # Initialize the main window
        self.title("BERC Automated Finance Tracker")
        padding = {"padx": 6, "pady": 6}

        # Initialize all Tkinter widgets
        frame = tk.Frame(self, bd=-2)
        inner_frame = tk.Frame(frame, bd=-2)
        action_frame = tk.LabelFrame(inner_frame, text="Actions", **padding)

        # Create logo widget
        logo_image = tk.PhotoImage(file="bercaft/logo.gif")
        logo_label = tk.Label(inner_frame, image=logo_image)
        logo_label.image = logo_image

        # Create action button widgets
        update_label = tk.Label(action_frame, text="Update the workbook")
        update_button = tk.Button(action_frame, text="Update", 
            command=self.update, height=1, width=10)

        backup_label = tk.Label(action_frame, text="Save a backup")
        backup_button = tk.Button(action_frame, text="Backup", 
            command=self.backup, height=1, width=10)

        exit_label = tk.Label(action_frame, text="Save and quit")
        exit_button = tk.Button(action_frame, text="Exit", 
            command=self.exit, height=1, width=10)

        # Create console widget
        console_font = tkinter.font.Font(family="Courier", size=10)
        console = tk.Text(frame, bg='#333', fg='white', bd=1, font=console_font, 
            height=8, width=60)
        self.console_handler = WidgetLogger(console)
        logger.addHandler(self.console_handler)

        # Add frames to grid
        frame.grid(column=0, row=0)
        inner_frame.grid(column=0, row=0, **padding)

        logo_label.grid(column=0, row=0, **padding)
        action_frame.grid(column=1, row=0)

        # Insert buttons and labels into grid
        update_button.grid(column=0, row=0)
        update_label.grid(column=1, row=0, sticky=tk.W)

        backup_button.grid(column=0, row=1)
        backup_label.grid(column=1, row=1, sticky=tk.W)

        exit_button.grid(column=0, row=2)
        exit_label.grid(column=1, row=2, sticky=tk.W)

        console.grid(column=0, row=1)
        self.resizable(width=False, height=False)

        logger.info("User interface created successfully")


    def backup(self):
        def backup_callback():
            tk.messagebox.showinfo("Backup",
                "Backup created successfully!")

        if self.dispatch.edit_lock.acquire(blocking=False):
            logger.debug("Creating backup thread")

            t = Thread(target=self.dispatch.backup, daemon=True, 
                args=(backup_callback,))
            
            logger.debug("Backup thread created")
            t.start()
        else:
            tk.messagebox.showerror("Backup", 
                "The spreadsheet is currently being modified!")


    def update(self):
        def update_callback():
            tk.messagebox.showinfo("Update",
                "Update completed successfully!")

        if self.dispatch.edit_lock.acquire(blocking=False):
            logger.debug("Creating update thread")

            t = Thread(target=self.dispatch.update, daemon=True, 
                args=(update_callback,))

            logger.debug("Update thread created.")
            t.start()
        else:
            tk.messagebox.showerror("Update", 
                "The spreadsheet is currently being modified!")


    def exit(self):
        if not self.dispatch.edit_lock.acquire(blocking=False):
            if not tk.messagebox.askyesno("Exit", 
                "The spreadsheet is currently being modified.\n"
                "Quitting now may leave it in an inconsistent state.\n"
                "Are you sure you wish to exit?"):
                return

        self.dispatch.edit_lock.release()
        self.destroy()


class WidgetLogger(logging.Handler):
    def __init__(self, widget):
        logging.Handler.__init__(self)
        self.setLevel(logging.INFO)
        self.widget = widget
        self.widget.config(state='disabled')
        self.widget.bind("<1>", lambda event: self.widget.focus_set())


    def emit(self, record):
        self.widget.config(state='normal')
        self.widget.insert(tk.END, self.format(record) + '\n')
        self.widget.see(tk.END)
        self.widget.config(state='disabled')


