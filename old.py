from tkinter import *
import ctypes, sys
from PyPDF2 import PdfFileReader, PdfFileWriter
from tkinter.filedialog import askopenfile
from tkinter import filedialog, messagebox
from tkinter.filedialog import asksaveasfile


def is_admin():
    try:
        d = ctypes.windll.shell32.IsUserAnAdmin()
        return True
    except:
        return False


class PDF():
    name: str
    is_alive: bool
    pages: int

    def __init__(self, name):
        self.pages = 0
        self.name = name
        self.is_alive = True

    def kill(self):
        self.is_alive = False

    def get(self):
        return self.name

    def get_existence(self):
        return self.is_alive

    def add_more_page(self, p = 1):
        self.pages = self.pages + p

    def get_pages(self):
        return self.pages


class Main():
    pdfs: list
    pdf_holder: list
    root: Tk
    pdfside: Frame
    btnside: Frame
    dracula_back = "#282a36"
    total_pdf: StringVar
    total_pdff: int

    def __init__(self):
        self.pdfs = []
        self.pdf_holder = []
        self.root = Tk()
        self.configure_root()
        self.total_pdf = StringVar()
        self.total_pdf.set("Total Pages: 0")
        self.total_pdff = 0
        # self.pdf_frame = Frame(self.root, background="red", height=500)
        # self.acc_frame = Frame(self.root, background="blue", height=500, width=350)
        self.pdf_frame = Frame(self.root, background=self.dracula_back, height=500)
        self.acc_frame = Frame(self.root, background=self.dracula_back, height=500, width=350)

        self.acc_frame.grid_configure(row=3, column=1)
        self.acc_frame.grid_rowconfigure(0, weight=1)
        self.acc_frame.grid_rowconfigure(1, weight=1)
        self.acc_frame.grid_rowconfigure(2, weight=1)

        self.pdf_frame.grid(row=0, column=0)
        self.acc_frame.grid(row=0, column=1)

        self.btn_frame = Frame(self.acc_frame, background=self.dracula_back)
        self.btn_frame.grid(row=1, column=0, sticky=E + W)

        # self.add_btn = Button(self.btn_frame, text="Add PDF", command=lambda: self.add_pdf(f"hello {self.temp}"))
        self.add_btn = Button(self.btn_frame, text="Add PDF", command=self.add_pdf)
        self.clear_btn = Button(self.btn_frame, text="Clear", command=self.clear_pdfs)
        self.merge_and_save_btn = Button(self.acc_frame, text="Merge And Save", command=self.merge_pdfs)
        self.total_pdf_button = Button(self.acc_frame, textvariable=self.total_pdf, state="disabled")
        self.total_pdf_button.grid(row=0, column=0)
        print(self.total_pdf.get())
        self.clear_btn.grid(row=0, column=0)
        self.add_btn.grid(row=0, column=1)
        self.merge_and_save_btn.grid(row=2, column=0, sticky=E + W)

    def merge_pdfs(self):
        output = asksaveasfile(initialfile='Merged', defaultextension=".txt",
                               filetypes=[("PDF Files", "*.pdf*")]).name + ".pdf"
        pdf_writer = PdfFileWriter()
        for pdf in self.pdfs:
            if pdf.get_existence():
                pdf_reader = PdfFileReader(pdf.get())
                for page in range(pdf_reader.getNumPages()):
                    # Add each page to the writer object
                    pdf_writer.addPage(pdf_reader.getPage(page))
            # Write out the merged PDF
        with open(output, 'wb') as out:
            pdf_writer.write(out)

    def add_pdf(self, path=None):
        path = filedialog.askopenfile(mode='r', filetypes=[('PDF Files', '*.pdf')]).name
        if not path:
            return
        try:
            temp = self.total_pdff
            self.total_pdff = self.total_pdff + PdfFileReader(path).getNumPages()
            self.pdfs.append(PDF(path))
            self.pdfs[len(self.pdfs - 1)].add_more_page(self.total_pdff - temp)
            self.total_pdf.set("total pages: " + str(self.total_pdff))

        except Exception as e:
            if "PDF starts with" in str(e):
                messagebox.showerror("Error", "The pdf file is corrupted")
            else:
                messagebox.showerror("Error", e)
            return

        self.pdf_holder.append(Frame(self.pdf_frame, background=self.dracula_back, width=350))
        lenf = len(self.pdf_holder) - 1
        Label(self.pdf_holder[lenf],
              text=path.split('\\')[len(path.split('\\')) - 1],
              background=self.dracula_back).grid(row=0, column=0, padx=2, pady=2)
        Button(self.pdf_holder[lenf],
               text="❌",
               background=self.dracula_back,
               foreground="red",
               command=lambda: self.remove_pdf(lenf)
               ).grid(row=0, column=1, padx=2, pady=2)
        self.pdf_holder[lenf].grid()

    def start(self):
        self.root.mainloop()

    def configure_root(self):
        self.root.title("PDF merger By Subhranil Maity v2.0")
        self.root.geometry("700x500+100+100")
        self.root.configure(background=self.dracula_back)
        self.root.resizable(False, False)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)

    def remove_pdf(self, lenf):
        print(lenf)
        self.pdf_holder[lenf].destroy()
        self.total_pdff = self.total_pdff - self.pdfs[lenf].get_pages()
        self.total_pdf.set("total pages: " + str(self.total_pdff))
        self.pdfs[lenf].kill()

    def clear_pdfs(self):
        print("To be done")


if is_admin():
    pass
    # Code of your program here
    Window = Main()
    Window.start()
else:
    # Re-run the program with admin rights
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
