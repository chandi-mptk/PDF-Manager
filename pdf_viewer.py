from tkinter import Tk, Canvas, Scrollbar, Frame, Label, Scale, HORIZONTAL, VERTICAL, Menu, DISABLED, NORMAL
from tkinter.filedialog import askopenfilename
from PIL import Image, ImageTk
import fitz  # PyMuPDF


class PDFViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Viewer")
        height, width = self.root.winfo_screenheight(), self.root.winfo_screenwidth()
        self.root.geometry(f"{width}x{height}")

        # Create menubar
        self.menubar = Menu(self.root)
        self.root.config(menu=self.menubar)

        # Create File menu with Open and Exit option with separator
        self.file_menu = Menu(self.menubar, tearoff=0)
        self.file_menu.add_command(label="Open", command=self.load_pdf, state=NORMAL)
        self.file_menu.add_command(label="Close", command=self.close_pdf, state=DISABLED)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Exit", command=self.root.quit)
        self.menubar.add_cascade(label="File", menu=self.file_menu)

        self.canvas_frame = Frame(root)
        self.canvas_frame.pack(fill="both", expand=True)

        self.canvas = Canvas(self.canvas_frame)
        self.canvas.pack(side="left", fill="both", expand=True)

        self.scrollbar_y = Scrollbar(self.canvas_frame, orient=VERTICAL, command=self.canvas.yview)
        self.scrollbar_y.pack(side="right", fill="y")

        self.scrollbar_x = Scrollbar(root, orient=HORIZONTAL, command=self.canvas.xview)
        self.scrollbar_x.pack(side="bottom", fill="x")

        self.canvas.configure(yscrollcommand=self.scrollbar_y.set, xscrollcommand=self.scrollbar_x.set)

        self.page_label = Label(root, text="Page:")
        self.page_label.pack(side="top", pady=5)

        self.page_scale = Scale(root, from_=1, to=1, orient=HORIZONTAL, command=self.show_page)
        self.page_scale.pack(side="top", fill="x", padx=5)

        self.pdf_document = None

    def menu_manager(self):
        if self.pdf_document:
            self.file_menu.entryconfig("Open", state=DISABLED)
            self.file_menu.entryconfig("Close", state=NORMAL)
            
        else:
            self.file_menu.entryconfig("Open", state=NORMAL)
            self.file_menu.entryconfig("Close", state=DISABLED)

    def load_pdf(self):
        file_path = askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.pdf_document = fitz.open(file_path)
            num_pages = self.pdf_document.page_count
            self.page_scale.config(from_=1, to=num_pages, length=600, tickinterval=1, resolution=1)
            self.page_scale.set(1)
            self.show_page()
            self.menu_manager()

    def close_pdf(self):
        self.pdf_document = None
        self.page_scale.set(1)
        self.canvas.delete("all")
        self.page_label.config(text="Page:")
        self.page_scale.config(from_=1, to=1, length=600, tickinterval=1, resolution=1)
        self.page_scale.set(1)
        self.canvas.config(scrollregion=(0, 0, 0, 0))
        self.canvas.image = None
        self.canvas.pack_forget()
        self.canvas_frame.pack_forget()
        self.scrollbar_y.pack_forget()
        self.scrollbar_x.pack_forget()
        self.page_label.pack_forget()
        self.page_scale.pack_forget()
        self.root.geometry(f"{self.root.winfo_screenwidth()}x{self.root.winfo_screenheight()}")
        self.root.title("PDF Viewer")
        self.menu_manager()

    def show_page(self, *args):
        if self.pdf_document:
            page_number = int(self.page_scale.get())
            page = self.pdf_document[page_number - 1]

            # Convert PDF page to Pillow image
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img = ImageTk.PhotoImage(img)

            # Update canvas with the image in center of the canvas
            self.canvas.delete("all")
            self.canvas.image = img
            self.canvas.create_image(self.canvas.winfo_width() / 2, self.canvas.winfo_height() / 2, image=img)
            self.canvas.config(scrollregion=page.rect)
            self.page_label.config(text=f"Page {page_number}")
            self.root.title(f"PDF Viewer - Page {page_number}")
            


if __name__ == "__main__":
    root = Tk()
    pdf_viewer = PDFViewer(root)
    root.mainloop()
