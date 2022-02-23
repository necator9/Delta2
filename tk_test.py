import tkinter
import tkinter.messagebox
import tkinter.filedialog as fd
import customtkinter
import sys
import os
from datetime import datetime


customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("dark-blue")  # Themes: "blue" (standard), "green", "dark-blue"



class CreateToolTip(object):
    """
    create a tooltip for a given widget
    """
    def __init__(self, widget, text='widget info'):
        self.waittime = 500     #miliseconds
        self.wraplength = 180   #pixels
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
        self.id = None
        self.tw = None

    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hidetip()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(self.waittime, self.showtip)

    def unschedule(self):
        id = self.id
        self.id = None
        if id:
            self.widget.after_cancel(id)

    def showtip(self, event=None):
        x = y = 0
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        # creates a toplevel window
        self.tw = tkinter.Toplevel(self.widget)
        # Leaves only the label and removes the app window
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (x, y))
        label = tkinter.Label(self.tw, text=self.text, justify='left',
                       background="#ffffff", relief='solid', borderwidth=1,
                       wraplength = self.wraplength)
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tw
        self.tw= None
        if tw:
            tw.destroy()



class App(customtkinter.CTk):

    APP_NAME = "CustomTkinter complex example"
    WIDTH = 700
    HEIGHT = 500

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)


        self.tracked_cols = ['Description', 'Feature']

        self.title(App.APP_NAME)
        self.geometry(str(App.WIDTH) + "x" + str(App.HEIGHT))
        self.minsize(App.WIDTH, App.HEIGHT)

        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        if sys.platform == "darwin":
            self.bind("<Command-q>", self.on_closing)
            self.bind("<Command-w>", self.on_closing)
            self.createcommand('tk::mac::Quit', self.on_closing)

        # ============ create two CTkFrames ============

        self.frame_left = customtkinter.CTkFrame(master=self,
                                                 width=200,
                                                 height=App.HEIGHT-40,
                                                 corner_radius=5)
        self.frame_left.place(relx=0.32, rely=0.5, anchor=tkinter.E)

        self.frame_right = customtkinter.CTkFrame(master=self,
                                                  width=420,
                                                  height=App.HEIGHT-40,
                                                  corner_radius=5)
        self.frame_right.place(relx=0.365, rely=0.5, anchor=tkinter.W)

        # ============ frame_left ============



        

        # self.check_box_1 = customtkinter.CTkCheckBox(master=self.frame_left,
        #                                              text="CTkCheckBox")
        # self.check_box_1.place(relx=0.5, rely=0.92, anchor=tkinter.CENTER)

        # ============ frame_right ============

        self.frame_info = customtkinter.CTkFrame(master=self.frame_right,
                                                 width=350,
                                                 height=150,
                                                 corner_radius=1)
        self.frame_info.place(relx=0.5, rely=0.5, anchor=tkinter.N)

        # ============ frame_right -> frame_info ============

        self.label_info_1 = customtkinter.CTkLabel(master=self.frame_info,
                                                   text="Tracked columns",
                                                   width=150,
                                                   height=20,
                                                   corner_radius=8,
                                                   fg_color=("gray76", "gray38"),  # <- custom tuple-color
                                                   justify=tkinter.LEFT)
        self.label_info_1.place(relx=0.5, rely=0.1, anchor=tkinter.CENTER)

        # self.head1 = tkinter.Text(self.frame_left, height=2, width=30)
        self.head1 = tkinter.Entry(self.frame_left)
        # head1.insert(tkinter.END,'Tracked columns')
        self.head1.place(relx=0.06, rely=0.5, height=200, width=100, anchor=tkinter.NW)


        self.col_entry = customtkinter.CTkEntry(master=self.frame_info,
                               placeholder_text="Enter columns",
                               width=140,
                               height=29,
                               corner_radius=1)
        self.col_entry.place(relx=0.06, rely=0.22, anchor=tkinter.NW)

        self.add_button = customtkinter.CTkButton(master=self.frame_info,
                                                text="Add",
                                                command=self.add_column,
                                                border_width=0,
                                                corner_radius=3)
        self.add_button.place(relx=0.94, rely=0.2, anchor=tkinter.NE)

        add_tip = 'Symbol difference (\"Diff\" column in the output file) is shown only for the first column in the list. Arrange the order as needed.'
        button1_ttp = CreateToolTip(self.add_button, add_tip)


        self.delete_button = customtkinter.CTkButton(master=self.frame_info,
                                                text="Delete",
                                                command=self.delete_column,
                                                border_width=0,
                                                corner_radius=3)
        self.delete_button.place(relx=0.94, rely=0.66, anchor=tkinter.NE)



        self.col_listbox = tkinter.Listbox(self.frame_info, selectmode=tkinter.BROWSE)
        self.col_listbox.place(relx=0.06, rely=0.45, width=140, height=60, anchor=tkinter.NW)
        for entry in self.tracked_cols:
            self.col_listbox.insert(tkinter.END, entry)

        scrollbar = tkinter.Scrollbar(self.frame_info)
        scrollbar.config(command=self.col_listbox.yview)
        scrollbar.place(relx=0.4, rely=0.48, anchor=tkinter.NW)
        # scrollbar.pack(side="right", fill="y")

        self.col_listbox.config(yscrollcommand=scrollbar.set)
        # scrollbar.pack(side=tkinter.RIGHT, fill=tkinter.Y)





        # self.progressbar = customtkinter.CTkProgressBar(master=self.frame_info,
        #                                                 width=250,
        #                                                 height=12)
        # self.progressbar.place(relx=0.5, rely=0.85, anchor=tkinter.S)

        # ============ frame_right <- ============

        # self.slider_1 = customtkinter.CTkSlider(master=self.frame_right,
        #                                         width=160,
        #                                         height=16,
        #                                         border_width=5,
        #                                         from_=1,
        #                                         to=0,
        #                                         number_of_steps=3,
        #                                         command=self.progressbar.set)
        # self.slider_1.place(x=20, rely=0.6, anchor=tkinter.W)
        # self.slider_1.set(0.3)

        # self.slider_2 = customtkinter.CTkSlider(master=self.frame_right,
        #                                         width=160,
        #                                         height=16,
        #                                         border_width=5,
        #                                         command=self.progressbar.set)
        # self.slider_2.place(x=20, rely=0.7, anchor=tkinter.W)
        # self.slider_2.set(0.7)

        # self.label_info_2 = customtkinter.CTkLabel(master=self.frame_right,
        #                                            text="CTkLabel: Lorem ipsum",
        #                                            fg_color=None,
        #                                            width=180,
        #                                            height=20,
        #                                            justify=tkinter.CENTER)
        # self.label_info_2.place(x=310, rely=0.6, anchor=tkinter.CENTER)

        # self.button_4 = customtkinter.CTkButton(master=self.frame_right,
        #                                         height=25,
        #                                         text="CTkButton",
        #                                         command=self.button_event,
        #                                         border_width=0,
        #                                         corner_radius=8)
        # self.button_4.place(x=310, rely=0.7, anchor=tkinter.CENTER)

        # self.entry = customtkinter.CTkEntry(master=self.frame_right,
        #                                     width=120,
        #                                     height=25,
        #                                     corner_radius=8,
        #                                     placeholder_text="CTkEntry")
        # self.entry.place(relx=0.33, rely=0.92, anchor=tkinter.CENTER)



        # Frame file select ===================================================================

        self.frame_f_select = customtkinter.CTkFrame(master=self.frame_right,
                                                     width=350,
                                                     height=200,
                                                     corner_radius=1)
        self.frame_f_select.place(relx=0.5, rely=0.05, anchor=tkinter.N)


        self.select_label = customtkinter.CTkLabel(master=self.frame_f_select,
                                                   text="Select files",
                                                   width=150,
                                                   height=20,
                                                   corner_radius=8,
                                                   fg_color=("gray76", "gray38"),  # <- custom tuple-color
                                                  justify=tkinter.LEFT)
        self.select_label.place(relx=0.5, rely=0.1, anchor=tkinter.CENTER)

        self.old_f_listbox = tkinter.Listbox(self.frame_f_select, selectmode=tkinter.BROWSE)
        self.old_f_listbox.place(relx=0.06, rely=0.34, width=140, height=130, anchor=tkinter.NW)
        # for entry in self.tracked_cols:
        #     self.col_listbox.insert(tkinter.END, entry)

        old_f_scrollbar = tkinter.Scrollbar(self.frame_f_select)
        old_f_scrollbar.config(command=self.old_f_listbox.yview)
        old_f_scrollbar.place(relx=0.4, rely=0.37, height=120, anchor=tkinter.NW)
        self.old_f_listbox.config(yscrollcommand=old_f_scrollbar.set)

        self.old_file_b = customtkinter.CTkButton(master=self.frame_right,
                                                  text="+ Old file/s",
                                                  command=self.open_old_files,
                                                  border_width=0,
                                                  corner_radius=4,
                                                  width=140)
        self.old_file_b.place(relx=0.135, rely=0.12, anchor=tkinter.NW)

        self.new_file_b = customtkinter.CTkButton(master=self.frame_right,
                                                  text="+ New file/s",
                                                  command=self.open_new_files,
                                                  border_width=0,
                                                  corner_radius=4,
                                                  width=140)
        self.new_file_b.place(relx=0.535, rely=0.12, anchor=tkinter.NW)

        self.new_f_listbox = tkinter.Listbox(self.frame_f_select, selectmode=tkinter.BROWSE)
        self.new_f_listbox.place(relx=0.54, rely=0.34, width=140, height=130, anchor=tkinter.NW)
        # for entry in self.tracked_cols:
        #     self.col_listbox.insert(tkinter.END, entry)

        new_f_scrollbar = tkinter.Scrollbar(self.frame_f_select)
        new_f_scrollbar.config(command=self.new_f_listbox.yview)
        new_f_scrollbar.place(relx=0.88, rely=0.37, height=120, anchor=tkinter.NW)
        self.new_f_listbox.config(yscrollcommand=new_f_scrollbar.set)

        self.filetypes = (('Excel files', '*.xlsx'), ('All files', '*.*'))

        self.old_files = list()
        self.new_files = list()

        self.start_b = customtkinter.CTkButton(master=self.frame_right,
                                               text="Start",
                                               command=self.process,
                                               border_width=0,
                                               corner_radius=4,
                                               width=140)
        self.start_b.place(relx=0.535, rely=0.88, anchor=tkinter.NW)

        # self.out_dir_b = customtkinter.CTkButton(master=self.frame_right,
        #                                          text="Out dir",
        #                                          command=self.select_out_dir,
        #                                          border_width=0,
        #                                          corner_radius=4,
        #                                          width=140)
        # self.out_dir_b.place(relx=0.135, rely=0.88, anchor=tkinter.NW) # 

        # self.out_dir = str()


        # self.progressbar.set(0.65)

    def button_event(self):
        print("Button pressed")

    def on_closing(self, event=0):
        self.destroy()

    def start(self):
        self.mainloop()

    def print_entry(self):
    	# print(self.col_entry.get())
    	print(self.col_listbox.get(0, 100))
    	# print(self.col_listbox.curselection())


    def delete_column(self):
        selection = self.col_listbox.curselection()
        if selection:
            self.col_listbox.delete(selection[0])

    def add_column(self):
        new_col = self.col_entry.get()
        self.tracked_cols = self.get_all_columns()
        if new_col != '' and new_col not in self.tracked_cols:
            self.col_listbox.insert(0, new_col)

    def get_all_columns(self):
        self.tracked_cols = self.col_listbox.get(0, 100)
        return self.tracked_cols

    def add(self):
        self.head1.delete(0, tkinter.END)
        Ans = self.get_all_columns()
        self.head1.insert(0, Ans)

    def error(self):
        tkinter.messagebox.showerror("Title", "Message")

    def open_old_files(self):
        self.old_files = list()
        self.old_f_listbox.delete(0, tkinter.END)
        self.old_files = fd.askopenfilenames(parent=self.frame_f_select, title='Choose a file', filetypes=self.filetypes)
        for i, entry in enumerate(self.old_files):
            trim = os.path.basename(entry)
            self.old_f_listbox.insert(tkinter.END, f'{i + 1}. {trim}')

    def open_new_files(self):
        self.new_files = list()
        self.new_f_listbox.delete(0, tkinter.END)
        self.new_files = fd.askopenfilenames(parent=self.frame_f_select, title='Choose a file', filetypes=self.filetypes)
        for i, entry in enumerate(self.new_files):
            trim = os.path.basename(entry)
            self.new_f_listbox.insert(tkinter.END, f'{i + 1}. {trim}')


    # def select_out_dir(self):
    #     self.out_dir = fd.askdirectory(title='Choose an output directory')

    def process(self):
        # if len(self.old_files) != len(self.new_files):
        #     tkinter.messagebox.showerror('Configuration error', 'Amount of files does not match between old and new lists ')
        #     return 1
        # if len(self.old_files) == 0 or len(self.new_files) == 0:
        #     tkinter.messagebox.showerror('Configuration error', 'At least one file should be selected to compare')
        #     return 1
        # self.tracked_cols = self.get_all_columns()
        # if len(self.tracked_cols) == 0:
        #     tkinter.messagebox.showerror('Configuration error', 'At least one tracked column should be defined')
        #     return 1


        self.out_dir = fd.askdirectory(title='Choose an output directory')
        if len(self.out_dir) == 0:
            tkinter.messagebox.showerror('Configuration error', 'An output directory must be selected')
            return 1




        # if len(self.out_dir) == 0:
        #     self.out_dir = os.path.join(os.path.expanduser('~'), 'Documents', 'Delta2', datetime.now().strftime("%H-%M-%S_%d.%m.%Y"))
        #     tkinter.messagebox.showinfo('Info', f'Outputting into default directory: {self.out_dir}')
        #     if not os.path.exists(self.out_dir):
        #         os.makedirs(self.out_dir)




if __name__ == "__main__":
    app = App()
    app.start()