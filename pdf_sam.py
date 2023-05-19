
import sys, os
# TODO: delete in final version
sys.path.append(".\\submodules")

import ttkExt.table as table

import tkinter as tk
import tkinter.ttk as ttk

from tkinter.filedialog import askopenfilenames
from tkinter.filedialog import asksaveasfilename
from tkinter.filedialog import askdirectory

#import ttkExt.table as table
from pypdf import PdfWriter, PdfReader

PDF_FILE_TYPES = [('pdf files', "*.pdf")]

# get a string "1-[number of pages]" from an input filename
def get_page_range(pdfreader):
    return f"1-{len(pdfreader.pages)}"

# parse/provide iterable for the pages_string
class PageNumbers:
    def __init__(self, page_string):
        self.page_string = page_string
        self.i = 0
        self.N = len(page_string)
        self.cur = 0
        self.stop = 0
        self.step = 1
    def __iter__(self):
        return self
    def _set_start_stop(self, substr):
        vals = [int(i) for i in substr.split('-')]
        if len(vals) == 2:
            self.cur = vals[0]-1
            self.stop = vals[1]-1
        else:
            self.cur = vals[0]-1
            self.stop = self.cur
    def __next__(self):
        self.cur += 1
        if self.cur > self.stop:
            # find next range
            if self.i >= self.N:
                raise StopIteration
            j = self.i
            self.i += 1
            while self.i < self.N and self.page_string[self.i] != ',':
                self.i += 1
            self._set_start_stop(self.page_string[j:self.i])
            self.i += 1
        return self.cur

# move to ttkExt
class LineSpaces(tk.Frame):
    def __init__(self, master=None, nspaces=1, side="top", **kwargs):
        super().__init__(master, **kwargs)
        lbl = ttk.Label(self, text="\n"*nspaces)
        lbl.pack(side=side)

class PDFSAM(tk.Frame):
    def __init__(self, master=None, starting_folder=None, **kwargs):
        self._runnable = False
        if not master:
            self._runnable = True
            master = tk.Tk()
        self.master = master
        super().__init__(self.master, **kwargs)
        self.cwd = starting_folder if starting_folder and os.path.isdir(starting_folder) else None
        
        load_frame = ttk.Frame(self)
        from_folder_btn = ttk.Button(load_frame, text="Load From Folder", command=self.load_from_folder)
        from_folder_btn.pack(side="left")
        from_file_btn = ttk.Button(load_frame, text="Load From File(s)", command=self.load_from_files)
        from_file_btn.pack(side="left")
        load_frame.pack(side="top")

        self._reader_map = {}

        LineSpaces(self, 2).pack(side="top")

        file_table_frame = ttk.Frame(self)
        self.file_table = table.Table(file_table_frame, [("File","Entry"), ("Pages: e.g. 1,3-5,8","Entry"), ("","Button"), ("","Button"), ("","Button")], 0)
        self.file_table.set_column(0, var="", width="64")
        self.file_table.set_column(1, var="", width="16")
        self.file_table.set_column(2, var="Up", width="6", command_row=self.move_file_up)
        self.file_table.set_column(3, var="Down", width="6", command_row=self.move_file_down)
        self.file_table.set_column(4, var="Del", width="6", command_row=self.delete_file)
        #self.file_table.set_column(5, var="Preview", command_row=self.preview_file)
        self.file_table.pack()
        file_table_frame.pack(side="top")

        LineSpaces(self, 2).pack(side="top")

        file_out_frame = ttk.Frame(self)
        file_out_lbl = ttk.Label(file_out_frame, text="Output File:")
        file_out_lbl.pack(side="top")
        selection_frame = ttk.Frame(file_out_frame)
        self.file_out = tk.StringVar(value="")
        file_out_ent = ttk.Entry(selection_frame, textvariable=self.file_out, width="64")
        file_out_ent.pack(side="left")
        select_file_out_btn = ttk.Button(selection_frame, command=self.select_output_file, text="Select File")
        select_file_out_btn.pack(side="left")
        merge_btn = ttk.Button(selection_frame, command=self.save_output_file, text="Merge")
        merge_btn.pack(side="left")
        selection_frame.pack(side="top")
        file_out_frame.pack(side="top")

        if self._runnable:
            LineSpaces(self, 1).pack(side="top")
            exit_btn = ttk.Button(self, command=self.stop, text="Close")
            exit_btn.pack(side="top")

        self.pack()

    def preview_file(self, row):
        def preview_command():
            print("previewing file: ", self.file_table[row, 0])
        return preview_command
    
    def delete_file(self, row):
        def delete_command():
            self._reader_map.pop(self.file_table[row, 0], None)
            self.file_table.del_row(row)
        return delete_command
    
    def move_file_up(self, row):
        def move_up_command():
            if row:
                file = self.file_table[row, 0]
                pages = self.file_table[row, 1]
                self.file_table[row, 0] = self.file_table[row-1, 0]
                self.file_table[row, 1] = self.file_table[row-1, 1]
                self.file_table[row-1, 0] = file
                self.file_table[row-1, 1] = pages
        return move_up_command
    
    def move_file_down(self, row):
        def move_down_command():
            if row < self.file_table.n_rows-1:
                file = self.file_table[row, 0]
                pages = self.file_table[row, 1]
                self.file_table[row, 0] = self.file_table[row+1, 0]
                self.file_table[row, 1] = self.file_table[row+1, 1]
                self.file_table[row+1, 0] = file
                self.file_table[row+1, 1] = pages
        return move_down_command

    def load_from_folder(self):
        folder = askdirectory(initialdir = self.cwd, title = "Select folder", mustexist = True)
        if folder:
            print("loading from folder: ", folder)

            for file in os.listdir(folder):
                if os.path.splitext(file)[1] == ".pdf":
                    self._add_file(os.path.join(folder, file))
            self.cwd = folder
            print("current working directory: ", self.cwd)

    # filename must already be a valid pdf file
    def _add_file(self, filename):
        print("loading filename: ", filename)
        # TODO: might modify filename to not show full path
        reader = PdfReader(filename)
        self._reader_map[filename] = reader
        row = self.file_table.n_rows
        self.file_table.add_row()
        self.file_table[row, 0] = filename
        self.file_table[row, 1] = get_page_range(reader)

    def load_from_files(self):
        filenames = askopenfilenames(initialdir = self.cwd, title = "Select file",filetypes = PDF_FILE_TYPES)
        if filenames:
            for filename in filenames:
                self._add_file(filename)
            self.cwd = os.path.split(filenames[0])[0]
            print("current working directory: ", self.cwd)

    def select_output_file(self):
        filename = asksaveasfilename(initialdir = self.cwd, initialfile = "merged_file.pdf", filetypes = PDF_FILE_TYPES)
        print("saving to filename: ", filename)
        self.file_out.set(filename)

    def save_output_file(self):
        merger = PdfWriter()
        for row in range(self.file_table.n_rows):
            if self.file_table[row, 1] != "":
                filename = self.file_table[row, 0]
                if filename in self._reader_map:
                    merger.append(self._reader_map[filename], pages=list(PageNumbers(self.file_table[row, 1])), import_outline=False)
                elif os.path.isfile(filename) and os.path.splitext(filename)[1] == '.pdf':
                    merger.append(filename, pages=list(PageNumbers(self.file_table[row, 1])), import_outline=False)
        merger.write(self.file_out.get())
        
    def run(self):
        if self._runnable:
            self.master.mainloop()

    def stop(self):
        if self._runnable:
            self.master.destroy()

if __name__=="__main__":
    PDFSAM().run()