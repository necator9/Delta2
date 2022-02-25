import tkinter
import tkinter.messagebox
import tkinter.filedialog as fd
import customtkinter
import sys
import os
from PIL import Image, ImageTk
import pandas as pd
import traceback


import delta2

import warnings
warnings.filterwarnings('ignore')


customtkinter.set_appearance_mode('System')  # Modes: 'System' (standard), 'Dark', 'Light'
customtkinter.set_default_color_theme('dark-blue')  # Themes: 'blue' (standard), 'green', 'dark-blue'

PATH = os.path.dirname(os.path.realpath(__file__))


class CreateToolTip(object):
    '''
    create a tooltip for a given widget
    '''
    def __init__(self, widget, text='widget info'):
        self.waittime = 500     #miliseconds
        self.wraplength = 180   #pixels
        self.widget = widget
        self.text = text
        self.widget.bind('<Enter>', self.enter)
        self.widget.bind('<Leave>', self.leave)
        self.widget.bind('<ButtonPress>', self.leave)
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
        x, y, cx, cy = self.widget.bbox('insert')
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        # creates a toplevel window
        self.tw = tkinter.Toplevel(self.widget)
        # Leaves only the label and removes the app window
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry('+%d+%d' % (x, y))
        label = tkinter.Label(self.tw, text=self.text, justify='left',
                       background='#ffffff', relief='solid', borderwidth=1,
                       wraplength = self.wraplength)
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tw
        self.tw= None
        if tw:
            tw.destroy()



class App(customtkinter.CTk):
    APP_NAME = 'Delta2'
    WIDTH = 700
    HEIGHT = 470

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.tracked_cols = ['Description', 'Feature']  # Hardcoded default columns 
        self.title(App.APP_NAME)
        self.geometry(str(App.WIDTH) + 'x' + str(App.HEIGHT))
        self.minsize(App.WIDTH, App.HEIGHT)
        self.protocol('WM_DELETE_WINDOW', self.on_closing)
        images_dir = os.path.join(PATH, 'images')
        delta_image_path = os.path.join(images_dir, 'delta.png')
        self.iconphoto(False, tkinter.PhotoImage(file=delta_image_path))
        if sys.platform == 'darwin':
            self.bind('<Command-q>', self.on_closing)
            self.bind('<Command-w>', self.on_closing)
            self.createcommand('tk::mac::Quit', self.on_closing)


        self.build_menu()

        # ============ fr_right ============
        self.fr_right = customtkinter.CTkFrame(master=self,
                                                 width=170,
                                                 height=App.HEIGHT-30,
                                                 corner_radius=2)
        self.fr_right.place(relx=0.98, rely=0.5, anchor=tkinter.E)

        self.delta_image = ImageTk.PhotoImage(Image.open(delta_image_path).resize((120, 120), Image.ANTIALIAS))
        delta_sign = customtkinter.CTkLabel(self.fr_right, 
                                        image=self.delta_image,
                                        fg_color=('gray85', 'gray38'),
                                        height=120,
                                        width=120,)
        delta_sign.place(relx=0.14, rely=0.15, anchor=tkinter.W)

        self.progressbar = customtkinter.CTkProgressBar(master=self.fr_right,
                                                        width=120,
                                                        height=12)
        self.progressbar.place(relx=0.5, rely=0.91, anchor=tkinter.S)
        self.progressbar.set(0)

        self.proc_image = ImageTk.PhotoImage(Image.open(PATH + '/images/Update_thin.png').resize((30, 30), Image.ANTIALIAS))
        self.start_b = customtkinter.CTkButton(master=self.fr_right,
                                               image=self.proc_image,
                                               text='Start',
                                               height=40,
                                               command=self.wrap_process,
                                               border_width=0,
                                               corner_radius=4,
                                               width=140)
        self.start_b.place(relx=0.92, rely=0.75, anchor=tkinter.NE)
        self.out_dir = str()
        self.output_names = list()

        # ============ fr_left ============
        fr_left_width = 490
        fr_left_height = App.HEIGHT - 30
        self.fr_left = customtkinter.CTkFrame(master=self,
                                                  width=fr_left_width,
                                                  height=fr_left_height,
                                                  corner_radius=2)
        self.fr_left.place(relx=0.02, rely=0.5, anchor=tkinter.W)

        # ============ fr_left -> Frame select files ============ 
        frame_f_select_width = fr_left_width - 30
        self.frame_f_select = customtkinter.CTkFrame(master=self.fr_left,
                                                     width=frame_f_select_width,
                                                     height=240,
                                                     corner_radius=1)
        self.frame_f_select.place(relx=0.5, rely=0.03, anchor=tkinter.N)


        self.select_label = customtkinter.CTkLabel(master=self.frame_f_select,
                                                   text='Select files',
                                                   width=150,
                                                   height=20,
                                                   corner_radius=8,
                                                   fg_color=('gray76', 'gray38'),  # <- custom tuple-color
                                                  justify=tkinter.LEFT)
        self.select_label.place(relx=0.5, rely=0.1, anchor=tkinter.CENTER)


        listbox_width = 180
        listbox_x_offset = 0.04
        listbox_rely = 0.38

        self.old_f_listbox = tkinter.Listbox(self.frame_f_select, selectmode=tkinter.BROWSE)
        self.old_f_listbox.place(relx=listbox_x_offset, rely=listbox_rely, width=listbox_width, height=130, anchor=tkinter.NW)

        old_f_scrollbar = tkinter.Scrollbar(self.old_f_listbox)
        old_f_scrollbar.config(command=self.old_f_listbox.yview)
        old_f_scrollbar.pack(side=tkinter.RIGHT, fill=tkinter.Y)
        self.old_f_listbox.config(yscrollcommand=old_f_scrollbar.set)

        add_list_image = ImageTk.PhotoImage(Image.open(PATH + '/images/add-list.png').resize((20, 20), Image.ANTIALIAS))
        sel_but_rely = 0.12
        self.old_file_b = customtkinter.CTkButton(master=self.fr_left,
                                                  image=add_list_image,
                                                  text='Old file/s',
                                                  command=self.open_old_files,
                                                  border_width=0,
                                                  corner_radius=4,
                                                  height=40,
                                                  width=listbox_width)
        self.old_file_b.place(relx=0.07, rely=sel_but_rely, anchor=tkinter.NW)

        self.new_file_b = customtkinter.CTkButton(master=self.fr_left,
                                                  image=add_list_image,
                                                  text='New file/s',
                                                  command=self.open_new_files,
                                                  border_width=0,
                                                  corner_radius=4,
                                                  height=40,
                                                  width=listbox_width)
        self.new_file_b.place(relx=0.565, rely=sel_but_rely, anchor=tkinter.NW)

        self.new_f_listbox = tkinter.Listbox(self.frame_f_select, selectmode=tkinter.BROWSE)
        self.new_f_listbox.place(relx=1 - listbox_x_offset, rely=listbox_rely, width=listbox_width, height=130, anchor=tkinter.NE)
        new_f_scrollbar = tkinter.Scrollbar(self.new_f_listbox)
        new_f_scrollbar.config(command=self.new_f_listbox.yview)
        new_f_scrollbar.pack(side=tkinter.RIGHT, fill=tkinter.Y)
        self.new_f_listbox.config(yscrollcommand=new_f_scrollbar.set)

        button_size = 40
        self.del_image = ImageTk.PhotoImage(Image.open(PATH + '/images/delete.png').resize((30, 30), Image.ANTIALIAS))
        self.delete_file_b = customtkinter.CTkButton(master=self.frame_f_select,
                                                     image=self.del_image,
                                                     text='',
                                                     command=self.delete_file,
                                                     height=button_size,
                                                     width=button_size,
                                                     border_width=0,
                                                     corner_radius=2)
        self.delete_file_b.place(relx=0.5, rely=0.835, anchor=tkinter.CENTER)
        self.filetypes = (('Excel files', '*.xlsx'), ('All files', '*.*'))

        self.old_files = list()
        self.new_files = list()
        self.new_names_trimmed = list()

        # ============ fr_left -> frame_tracked_cols ============

        self.frame_tracked_cols = customtkinter.CTkFrame(master=self.fr_left,
                                                 width=frame_f_select_width,
                                                 height=160,
                                                 corner_radius=1)
        self.frame_tracked_cols.place(relx=0.5, rely=0.60, anchor=tkinter.N)

        self.label_info_1 = customtkinter.CTkLabel(master=self.frame_tracked_cols,
                                                   text='Tracked columns',
                                                   width=150,
                                                   height=20,
                                                   corner_radius=8,
                                                   fg_color=('gray76', 'gray38'),  # <- custom tuple-color
                                                   justify=tkinter.LEFT)
        self.label_info_1.place(relx=0.5, rely=0.1, anchor=tkinter.CENTER)

        col_width = 135
        self.col_entry = customtkinter.CTkEntry(master=self.frame_tracked_cols,
                               placeholder_text='Enter columns',
                               width=col_width,
                               height=40,
                               corner_radius=1)
        self.col_entry.place(relx=listbox_x_offset, rely=0.22, anchor=tkinter.NW)

        button_size = 40
        add_del_relx = 0.45
        plus_image = ImageTk.PhotoImage(Image.open(PATH + '/images/plus.png').resize((15, 15), Image.ANTIALIAS))
        self.add_button = customtkinter.CTkButton(master=self.frame_tracked_cols,
                                                  image=plus_image,
                                                  text='',
                                                command=self.add_column,
                                                border_width=0,
                                                height=button_size,
                                                width=button_size,
                                                corner_radius=2)
        self.add_button.place(relx=add_del_relx, rely=0.22, anchor=tkinter.NE)

        add_tip = 'Symbol difference (\"Diff\" column in the output file) is shown only for the first column in the list. Arrange the order as needed.'
        button1_ttp = CreateToolTip(self.add_button, add_tip)

        self.delete_button = customtkinter.CTkButton(master=self.frame_tracked_cols,
                                                image=self.del_image,
                                                text='',
                                                command=self.delete_column,
                                                height=button_size,
                                                width=button_size,
                                                border_width=0,
                                                corner_radius=2)
        self.delete_button.place(relx=add_del_relx, rely=0.69, anchor=tkinter.NE)

        self.col_listbox = tkinter.Listbox(self.frame_tracked_cols, selectmode=tkinter.BROWSE)
        self.col_listbox.place(relx=listbox_x_offset, rely=0.53, width=col_width, height=65, anchor=tkinter.NW)
        for entry in self.tracked_cols:
            self.col_listbox.insert(tkinter.END, entry)

        scrollbar = tkinter.Scrollbar(self.col_listbox)
        scrollbar.config(command=self.col_listbox.yview)
        scrollbar.pack(side=tkinter.RIGHT, fill=tkinter.Y)
        self.col_listbox.config(yscrollcommand=scrollbar.set)

        self.info_heading_cb = tkinter.IntVar()
        self.track_info_heading_cb = customtkinter.CTkCheckBox(master=self.frame_tracked_cols,
                                                               text='Track info & heading',
                                                               variable=self.info_heading_cb)
        self.track_info_heading_cb.place(relx=0.55, rely=0.47, anchor=tkinter.NW)

    def build_menu(self):
        menu = tkinter.Menu(self)
        self.config(menu=menu)

        file_menu = tkinter.Menu(menu)
        file_menu.add_command(label='Select old files', command=self.open_old_files)
        file_menu.add_command(label='Select new files', command=self.open_new_files)
        file_menu.add_command(label='Exit',  command=self.on_closing)
        menu.add_cascade(label='File', menu=file_menu)

        edit_menu = tkinter.Menu(menu)
        edit_menu.add_command(label='Start processing', command=self.wrap_process)
        menu.add_cascade(label='Edit', menu=edit_menu)

        help_menu = tkinter.Menu(menu)
        help_menu.add_command(label='About')
        menu.add_cascade(label='Help', menu=help_menu)


    def on_closing(self, event=0):
        self.destroy()

    def start(self):
        self.mainloop()

    def delete_file(self):
        def handle_delete(listbox, selection):
            listbox.delete(selection[0])
            tmp = listbox.get(0, 100)
            listbox.delete(0, tkinter.END)
            for i, item in enumerate(tmp):
                listbox.insert(tkinter.END, f'{i + 1}. {item.split(". ", 1)[1]}')
            listbox.activate(selection[0])
            listbox.see(selection[0])

        selection_old = self.old_f_listbox.curselection()
        selection_new = self.new_f_listbox.curselection()

        if selection_old:
            handle_delete(self.old_f_listbox, selection_old)
             
        elif selection_new:
            handle_delete(self.new_f_listbox, selection_new)

    def delete_column(self):
        selection = self.col_listbox.curselection()
        if selection:
            self.col_listbox.delete(selection[0])

    def add_column(self):
        new_col = self.col_entry.get()
        self.tracked_cols = self.get_all_columns()
        if new_col != '' and new_col not in self.tracked_cols:
            self.col_listbox.insert(tkinter.END, new_col)
            self.col_listbox.see(tkinter.END)

    def get_all_columns(self):
        self.tracked_cols = self.col_listbox.get(0, 100)
        return self.tracked_cols

    def open_old_files(self):
        self.old_files = list()
        self.old_f_listbox.delete(0, tkinter.END)
        self.old_files = fd.askopenfilenames(parent=self.frame_f_select, title='Choose old file/s', filetypes=self.filetypes)
        for i, entry in enumerate(self.old_files):
            trim = os.path.basename(entry)
            self.old_f_listbox.insert(tkinter.END, f'{i + 1}. {trim}')

    def open_new_files(self):
        self.new_files = list()
        self.new_f_listbox.delete(0, tkinter.END)
        self.new_files = fd.askopenfilenames(parent=self.frame_f_select, title='Choose new file/s', filetypes=self.filetypes)
        for i, entry in enumerate(self.new_files):
            trim = os.path.basename(entry)
            self.new_f_listbox.insert(tkinter.END, f'{i + 1}. {trim}')

    def wrap_process(self):
        self.start_b.configure(state=tkinter.DISABLED)
        err = self.process()
        if not err:
            tkinter.messagebox.showinfo(title='Info', message='Processing completed!')
        self.start_b.configure(state=tkinter.NORMAL)
        self.progressbar.set(0)

    def process(self):
        # Update lists after editing
        def filter_deleted(paths, listbox_items):
          tmp = list()
          for path in paths:
            for substr in listbox_items:
              if substr.split('. ', 1)[1] in path:
                tmp.append(path)
          return tmp

        self.old_files = filter_deleted(self.old_files, self.old_f_listbox.get(0, 100))
        self.new_files = filter_deleted(self.new_files, self.new_f_listbox.get(0, 100))

        self.new_names_trimmed = [os.path.basename(entry) for entry in self.new_files]

        if len(self.old_files) != len(self.new_files):
            tkinter.messagebox.showerror('Configuration error', 'Amount of files does not match between old and new lists ')
            return 1
        if len(self.old_files) == 0 or len(self.new_files) == 0:
            tkinter.messagebox.showerror('Configuration error', 'At least one file should be selected to compare')
            return 1
        self.tracked_cols = self.get_all_columns()
        if len(self.tracked_cols) == 0:
            tkinter.messagebox.showerror('Configuration error', 'At least one tracked column should be defined')
            return 1
        self.out_dir = fd.askdirectory(title='Choose an output directory')
        if len(self.out_dir) == 0:
            tkinter.messagebox.showerror('Configuration error', 'An output directory must be selected')
            return 1
        try:
            self.output_names = [os.path.join(self.out_dir, name.replace('.xlsx', '_delta.xlsx')) for name in self.new_names_trimmed]

            index_col = 'ID'
            filter_info_heading = not self.info_heading_cb.get()

            stats = list()

            self.progressbar.set(0)
            step = 1 / len(self.output_names)
            prgs = 0

            for f_old_path, f_new_path, out_file, module in zip(self.old_files, self.new_files, self.output_names, self.new_names_trimmed):
                prgs += step

                d = delta2.Delta2(f_old_path, f_new_path, out_file, index_col, self.tracked_cols, filter_info_heading)

                stats.append([module.split(' ')[0], *d.process()])

                self.progressbar.set(prgs)
                self.update()

            stats_df = pd.DataFrame(stats, columns=['Module', 'New', 'Changed', 'Deleted', 'Total'])
            sh = delta2.ExcelStatsHandler(os.path.join(self.out_dir, 'Statistics.xlsx'), stats_df)
            sh.format_exel_output(stats_df.shape[0])
            sh.quit()
        except PermissionError as err:
            tkinter.messagebox.showerror('Execution error', 'Permission denied: close related *.xlsx files and repeat')
            return 1
        except delta2.Delta2Exception as err:
            tkinter.messagebox.showerror('Execution error', err)
            return 1
        except Exception as err:
            tkinter.messagebox.showerror('Execution error', str(traceback.format_exc()))
            return 1


    
        

if __name__ == '__main__':
    app = App()
    app.start()