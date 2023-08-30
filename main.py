import ast
import datetime
import os
import pathlib
import tkinter
import tkinter.messagebox
import customtkinter
from tkinter import filedialog
from tkinter import ttk
from workflow import *
from runner import *

customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"


class ToplevelWindow(customtkinter.CTkToplevel):
    def __init__(self, master, master_data, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.title("Task")
        self.geometry("400x300")
        # self.master = master

        self.frame = customtkinter.CTkScrollableFrame(self)
        self.frame.grid(row=0, column=0, sticky="nsew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.new_data = master_data
        # self.frame.grid_columnconfigure((0,1),weight=1)

        self.font = customtkinter.CTkFont("Roboto", weight="bold")
        self.description = customtkinter.CTkTextbox(self.frame, font=self.font, height=100)
        self.description.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
        self.description.insert("0.0", master_data["task_description"])
        self.description.grid_columnconfigure(0, weight=1)

        self.optional_label = customtkinter.CTkLabel(self.frame, text="Select Task Type: ", font=self.font)
        self.optional_label.grid(row=1, column=0, padx=10, pady=5)
        self.optionmenu_var = customtkinter.StringVar(value="Choose Task Type")
        self.optionmenu = customtkinter.CTkOptionMenu(self.frame, values=list(TASKS.keys()),
                                                      command=self.optionmenu_callback,
                                                      variable=self.optionmenu_var)
        self.optionmenu.grid(row=1, column=1, padx=10, pady=10)
        self.labels = []
        self.add_new_argument = customtkinter.CTkButton(self.frame, text="Add New Argument",
                                                        command=self.create_new_arguments)
        self.add_new_argument.grid(row=2, column=0, padx=10, pady=10)
        self.save_button = customtkinter.CTkButton(self.frame, text="Save this Task", command=self.save_data)
        if "task_type" in list(self.new_data.keys()):
            self.optionmenu.set(self.new_data["task_type"])
            self.load_data()

    def optionmenu_callback(self, choice):
        for label in self.labels:
            label[0].destroy(), label[1].destroy()

        self.labels.clear()
        for option in TASKS[choice]:
            self.create_new_arguments()
            self.labels[-1][0].insert(0, option)
            self.labels[-1][0].configure(state="disabled")

    def create_new_arguments(self):
        key = customtkinter.CTkEntry(self.frame, font=self.font)
        key.grid(row=3 + len(self.labels), column=0, padx=10, pady=2)
        value = customtkinter.CTkEntry(self.frame, font=self.font)
        value.grid(row=3 + len(self.labels), column=1, padx=5, pady=2)
        self.labels.append((key, value))
        self.save_button.grid_forget()
        self.save_button.grid(row=3 + len(self.labels), column=0, padx=10, pady=2)

    def load_data(self):
        data = self.new_data["task_data"]
        for key, value in data.items():
            self.create_new_arguments()
            self.labels[-1][0].insert(0, key)
            self.labels[-1][1].insert(0, value)
            if key in TASKS[self.new_data["task_type"]]:
                self.labels[-1][0].configure(state="disabled")
            else:
                pass

    def save_data(self):
        self.new_data = {
            "type": "task",
            "task_type": self.optionmenu_var.get(),
            "task_description": self.description.get("0.0", "end"),
            "task_data": {
                key.get(): self.convert_data_type(value.get()) for key, value in self.labels
            }
        }
        self.master.master_data = self.new_data
        self.master.set_description(self.description.get("0.0", "end")),
        self.destroy()

    def convert_data_type(self, value):
        try:
            ev = ast.literal_eval(value)
        except Exception as e:
            ev = value
        if isinstance(ev, str) and ev.isnumeric():
            ev = int(ev)
        return ev


class Task(customtkinter.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.grid_columnconfigure(0, weight=1)
        self.description = customtkinter.StringVar()
        self.description.set("Describe the task here...")
        self.entry = customtkinter.CTkEntry(self, textvariable=self.description, state="disabled")
        self.entry.grid(row=0, column=0, sticky="nsew")
        self.entry.bind("<Button-1>", self.press_task_tab)
        self.master_data = {
            "type": "task",
            "task_description": self.description.get(),
        }

    def set_description(self, describe):
        self.entry.configure(state="normal")
        self.description.set(describe)
        self.entry.configure(state="disabled")

    def press_task_tab(self, event):
        self.new_window = ToplevelWindow(self, master_data=self.master_data)
        self.new_window.focus_force()


class TaskCell(customtkinter.CTkScrollableFrame):
    def __init__(self, master, main_app, file_name, file_dir, **kwargs):
        super().__init__(master, **kwargs)
        self.grid_columnconfigure(1, weight=1)
        self.task_list = []  # contains (indexLabel, Task)
        self.add_new = customtkinter.CTkButton(self, text="Add Task", command=self.add_item)
        self.add_new.grid(row=0, column=0, padx=5, pady=5)
        self.savebutton = customtkinter.CTkButton(self, text="Compile & Save", command=self.save_data)
        self.savebutton.grid(row=0, column=1, padx=5, pady=5)

        self.runbutton = customtkinter.CTkButton(self, text="Run",command=self.run_data)
        self.runbutton.grid(row=0, column=2, padx=5, pady=5)

        self.file_name = file_name
        self.file_dir = file_dir
        self.main_app = main_app

    def add_item(self):
        indexLabel = customtkinter.CTkLabel(self, text=f"{len(self.task_list) + 1}")
        indexLabel.grid(row=len(self.task_list) + 1, column=0, padx=5, pady=5)
        TaskFrame = Task(self)
        TaskFrame.grid(row=len(self.task_list) + 1, column=1, padx=5, pady=5, sticky="we")
        self.task_list.append((indexLabel, TaskFrame))


    def delete_all(self):
        for (x,y) in self.task_list:
            x.destroy()
            y.destory()
        self.task_list.clear()


    def remove_item(self, index):
        if 0 <= index < len(self.task_list):
            y = self.task_list[index]
            y[0].destroy()
            y[1].destroy()
            counter = index
            while counter < len(self.task_list):
                x = self.task_list[counter]
                x[0].configure(text=f"{counter + 1}")

    def compile_all(self):
        y = []
        for task in self.task_list:
            y.append(task[1].master_data)

        data = {
            "cells": y
        }
        return data

    def save_data(self):
        try:
            data = self.compile_all()
            save_wf_file(self.file_name, self.file_dir, data)
            self.main_app.print_terminal(f"File {self.file_name} saved successfully!")
        except Exception as e:
            self.main_app.print_terminal(f"Not able to save file. Getting - {e}", "ERROR")

    def load_data(self):
        try:
            data = import_wf_file(self.file_name, self.file_dir)
            if "cells" not in list(data.keys()):
                return
            y = data["cells"]
            self.delete_all()
            for x in y:
                self.add_item()
                self.task_list[-1][1].master_data = x
                self.task_list[-1][1].set_description(x["task_description"])
            self.main_app.print_terminal(f"File {self.file_name} loaded successfully!")
        except Exception as e:
            self.main_app.print_terminal(f"Not able to load file. Getting - {e}","ERROR")


    def run_data(self):
        try:
            data = self.compile_all()
            work, err = load_workflow(data, name=self.file_name)
            if work is None:
                self.main_app.print_terminal(err)
                return
            work.run()
        except Exception as e:
            self.main_app.print_terminal(f"Not able to Run file. Getting - {e}", "ERROR")
class ScrollableLabelFrame(customtkinter.CTkScrollableFrame):
    def __init__(self, master, command=None, **kwargs):
        super().__init__(master, **kwargs)
        self.grid_columnconfigure(0, weight=1)

        self.command = command
        self.radiobutton_variable = customtkinter.StringVar()
        self.label_list = []

    def add_item(self, item, image=None):
        label = customtkinter.CTkLabel(self, text=item, image=image, compound="left", padx=5, anchor="w")
        if self.command is not None:
            label.bind("<Button-1>", lambda event: self.command(item))
        label.grid(row=len(self.label_list), column=0, pady=(0, 5), sticky="w")
        self.label_list.append(label)

    def remove_item(self, item):
        for label in self.label_list:
            if item == label.cget("text"):
                label.destroy()
                self.label_list.remove(label)
                return


class App(customtkinter.CTk):

    def __init__(self):
        super().__init__()

        # Variables
        self.FOLDER_PATH = None
        self.FILES = {}
        self.TABS = {}
        # configure window
        self.title("GUI Automation")
        self.geometry(f"{1100}x{580}")

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=2, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(2, weight=1)
        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="Files",
                                                 font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        self.sidebar_button_1 = customtkinter.CTkButton(self.sidebar_frame, text="Open Folder",
                                                        command=self.sidebar_button_event)
        self.sidebar_button_1.grid(row=1, column=0, padx=20, pady=10)
        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=3, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame,
                                                                       values=["Light", "Dark", "System"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=4, column=0, padx=20, pady=(10, 10))

        self.panwindow = ttk.PanedWindow(self, orient="vertical")
        self.panwindow.grid(row=0, column=1, sticky="nsew")
        self.panwindow.grid_rowconfigure(0, weight=1)
        self.panwindow.grid_columnconfigure(0, weight=1)
        # create textbox
        self.terminalFrame = customtkinter.CTkFrame(self.panwindow, height=200)
        self.terminalFrame.grid(row=1, column=0, sticky="nsew")
        self.terminalFrame.columnconfigure(0, weight=1)
        self.terminalFrame.rowconfigure(1, weight=1)

        self.textbox_label = customtkinter.CTkLabel(self.terminalFrame, text="Output Logs", anchor="w")
        self.textbox_label.grid(row=0, padx=40, sticky="nsew")

        self.textbox = customtkinter.CTkTextbox(self.terminalFrame)
        self.textbox.grid(row=1, column=0, padx=(10, 2), pady=(2, 2), sticky="nsew")
        self.textbox.columnconfigure(0, weight=1)
        self.textbox.configure(state="disabled")

        # create tabview
        self.tabview = customtkinter.CTkTabview(self.panwindow, width=250)
        self.tabview.grid(row=0, column=0, padx=(20, 0), pady=(20, 0), sticky="nsew")
        self.tabview.grid_rowconfigure(3, weight=1)
        self.tabview.add("+")  # should be at first place
        # Add new file label
        self.new_file_label = customtkinter.CTkButton(self.tabview.tab("+"), text="Create a New File",
                                                      command=self.create_new_file)
        self.new_file_label.grid(row=0, column=0, sticky="nsew")

        self.panwindow.add(self.tabview)
        self.panwindow.add(self.terminalFrame)
        # set default values
        self.appearance_mode_optionemenu.set("Dark")
        self.print_terminal("Welcome to GUI Automation")

    def open_input_dialog_event(self):
        dialog = customtkinter.CTkInputDialog(text="File Name: ", title="Create New File")
        print("CTkInputDialog: ", dialog.get_input())

    def open_file(self, file_name, file_dir):
        tab_name = file_name[:-3]
        if (tab_name, file_dir) in self.TABS.items():
            self.tabview.set(tab_name)
        else:
            try:
                # data = import_wkfw_file(path)
                self.create_tab(tab_name, file_dir)
                self.print_terminal(f"New File opened: {file_dir}/{tab_name}")
                self.tabview.set(tab_name)
                # process_data(data)
                if file_dir == self.FOLDER_PATH:
                    self.FILES[file_name] = file_dir

            except Exception as e:
                self.print_terminal(f"Not Able to open file, getting {e}", "ERROR")

    def open_new_file(self):
        dialog = filedialog.askopenfilename(defaultextension=".wf")
        file_dir, file_name = os.path.split(dialog)
        if file_name.endswith(".wf"):
            self.open_file(file_name, file_dir)

    def create_new_file(self):
        initdir = self.FOLDER_PATH
        if initdir is None:
            initdir = os.path.expanduser("~/Desktop")
        dialog = filedialog.asksaveasfilename(defaultextension="wf", initialdir=initdir,
                                              filetypes=[("Workflow Files", "*.wf")])
        file_dir, file_name = os.path.split(dialog)
        # print(file_dir,file_name)
        create_wf_file(file_name, file_dir)
        self.open_file(file_name, file_dir)

    def print_terminal(self, message, type="INFO"):
        self.textbox.configure(state="normal")
        formatted_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        new_message = f"[ {formatted_time} ] {type} :  {message}\n"
        self.textbox.insert("end", new_message)
        self.textbox.configure(state="disabled")

    def create_tab(self, tab_name, file_dir):
        # create scrollable frame
        self.tabview.insert(self.tabview._name_list.index("+"), tab_name)
        self.TABS[tab_name] = file_dir
        self.tabview.tab(tab_name).grid_columnconfigure(0, weight=1)
        self.tabview.tab(tab_name).grid_rowconfigure(0, weight=1)

        scrollable_frame = TaskCell(self.tabview.tab(tab_name),self, file_name=tab_name+".wf", file_dir=file_dir)
        scrollable_frame.grid(row=0, column=0, padx=(5, 0), pady=(5, 0), sticky="nsew")
        scrollable_frame.load_data()

        # scrollable_frame.grid_columnconfigure(0, weight=1)
        # scrollable_frame_switches = []
        # for i in range(100):
        #     switch = customtkinter.CTkSwitch(master=scrollable_frame, text=f"CTkSwitch {i}")
        #     switch.grid(row=i, column=0, padx=10, pady=(0, 20))
        #     scrollable_frame_switches.append(switch)
        # scrollable_frame_switches[0].select()

    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def sidebar_button_event(self):
        folder_path = filedialog.askdirectory()

        if folder_path:
            self.FOLDER_PATH = folder_path
            wf_files = [file for file in os.listdir(folder_path) if file.endswith(".wf")]
            self.print_terminal(f"Opened Folder {self.FOLDER_PATH}")
            if wf_files:
                self.updateFileList(wf_files)
            else:
                tkinter.messagebox.showinfo("No Workflow Files", "No workflow files found in the folder.")

    def command_file_open(self, event):
        self.open_file(file_name=event, file_dir=self.FOLDER_PATH)

    def updateFileList(self, files):
        files = [(file, self.FOLDER_PATH) for file in files]
        self.FILES.update(files)
        if hasattr(self, 'listFileWidget'):
            self.listFileWidget.destroy()
        else:
            self.sidebar_button_1.pack_forget()
        self.listFileWidget = ScrollableLabelFrame(self.sidebar_frame, command=self.command_file_open)
        self.listFileWidget.grid(row=1, column=0, sticky="nsew")
        for item in self.FILES.keys():
            self.listFileWidget.add_item(item)


if __name__ == "__main__":
    app = App()
    app.mainloop()
