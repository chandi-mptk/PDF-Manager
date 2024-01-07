from tkinter import (
    Tk, Canvas, Scrollbar, Frame, Label, HORIZONTAL, 
    VERTICAL, Menu, DISABLED, NORMAL, LEFT, 
    Y, Listbox, SINGLE, Button
    )
from tkinter.filedialog import askopenfilename
from tkinter.simpledialog import askstring
from tkinter.messagebox import showinfo, showerror, showwarning
from PIL import Image, ImageTk
import fitz  # PyMuPDF


class PDFViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Viewer")
        height, width = self.root.winfo_screenheight(), self.root.winfo_screenwidth()
        self.root.geometry(f"{width}x{height}")

        self.pdf_document = None
        self.current_page = 0
        self.current_zoom = 1.0
        self.num_pages = 0

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

        # Left Frame for page Preview and page list
        self.left_frame = Frame(root, width=width // 20)
        self.left_frame.pack(side=LEFT, fill=Y)

        self.page_list_label = Label(self.left_frame, text="Page List")
        self.page_list_label.pack(pady=5)

        self.page_listbox = Listbox(self.left_frame, selectmode=SINGLE, exportselection=0)
        self.page_listbox.pack(fill=Y, expand=True)

        self.zoom_in_button = Button(self.left_frame, text="Zoom In", command=self.zoom_in)
        self.zoom_in_button.pack(pady=5)

        self.zoom_out_button = Button(self.left_frame, text="Zoom Out", command=self.zoom_out)
        self.zoom_out_button.pack(pady=5)

        self.canvas_frame = Frame(root)
        self.canvas_frame.pack(fill="both", expand=True, padx=10)

        self.canvas = Canvas(self.canvas_frame)
        self.canvas.pack(side="left", fill="both", expand=True)

        self.scrollbar_y = Scrollbar(self.canvas_frame, orient=VERTICAL, command=self.canvas.yview)
        self.scrollbar_y.pack(side="right", fill="y")

        self.scrollbar_x = Scrollbar(root, orient=HORIZONTAL, command=self.canvas.xview)
        self.scrollbar_x.pack(side="bottom", fill="x")

        self.canvas.configure(yscrollcommand=self.scrollbar_y.set, xscrollcommand=self.scrollbar_x.set)

        self.page_label = Label(root, text="Page:")
        self.page_label.pack(side="top", pady=5)


    def menu_manager(self):
        if self.pdf_document:
            self.file_menu.entryconfig("Open", state=DISABLED)
            self.file_menu.entryconfig("Close", state=NORMAL)
            # self.page_listbox.bind("<ButtonRelease-1>", self.select_page_from_list)
            # self.page_listbox.bind("<ButtonRelease-4>", self.scroll_page_scale)
            self.page_listbox.bind("<ButtonRelease>", self.scroll_page_scale)
            
        else:
            self.file_menu.entryconfig("Open", state=NORMAL)
            self.file_menu.entryconfig("Close", state=DISABLED)
            # self.page_listbox.unbind("<ButtonRelease-1>")
            # self.page_listbox.unbind("<ButtonRelease-4>")
            self.page_listbox.unbind("<ButtonRelease>")

    def load_pdf(self):
        file_path = askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.pdf_document = fitz.open(file_path)
            # Check if pdf document is password protected
            if self.pdf_document.is_encrypted:
                while True:
                    password = askstring("Password", "Enter the password for the PDF file:")
                    if password is None:
                        break

                    # Attempt to decrypt the document
                    if self.pdf_document.authenticate(password):
                        break
                    else:
                        showerror("Error", "Incorrect password. Please try again.")

            self.num_pages = self.pdf_document.page_count
            
            # Populate the page listbox
            self.page_listbox.delete(0, 'end')
            for page_number in range(1, self.num_pages + 1):
                self.page_listbox.insert(page_number, f"Page {page_number}")
                # self.page_listbox.selection_set(page_number)
                # self.page_listbox.activate(page_number)
                # self.page_listbox.yview_moveto(page_number / num_pages)
                
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
        
        self.root.title("PDF Viewer")
        self.menu_manager()

    def show_page(self, *args):
        if self.pdf_document:
            
            # Update page listbox selection
            self.page_listbox.selection_clear(0, self.page_listbox.size() - 1)
            self.page_listbox.selection_set(self.current_page)
            self.page_listbox.activate(self.current_page)

            # Update Canvas with the image
            page = self.pdf_document[self.current_page]
            pix = page.get_pixmap()
            zoomed_width = int(pix.width * self.current_zoom)
            zoomed_height = int(pix.height * self.current_zoom)

            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img = img.resize((zoomed_width, zoomed_height), Image.ADAPTIVE)
            img = ImageTk.PhotoImage(img)

            offset_width = 0
            offset_height = 0

            # Check zoomed width and height is greater than canvas width and height
            if zoomed_width > self.canvas.winfo_width():
                offset_width = -(zoomed_width - self.canvas.winfo_width()) // 2
            if zoomed_height > self.canvas.winfo_height():
                offset_height = -(zoomed_height - self.canvas.winfo_height()) // 2
            
            # Screen center
            self.canvas.config(scrollregion=(offset_width, offset_height, zoomed_width, zoomed_height))
            self.canvas.create_image(self.canvas.winfo_width() / 2, self.canvas.winfo_height() / 2, image=img)
            self.canvas.image = img  # Keep a reference to prevent image from being garbage collected

    def scroll_page_scale(self, event):
        print("scroll", event)

        if event.num == 1:
            selected_index = self.page_listbox.curselection()
            if selected_index:
                self.current_page = selected_index[0]
                print("Current index",self.current_page)
                self.show_page()

        if event.num == 4:
            # Check if the current value is greater than 1
            if self.current_page > 0:
                # Decrement the value by 1
                self.current_page -= 1
                print("Current index",self.current_page)
                self.show_page()

        elif event.num == 5:
            print("last page", self.num_pages)
            # Check if the current value is less than the maximum value
            if self.current_page < self.num_pages:
                # Increment the value by 1
                self.current_page += 1
                print("Current index",self.current_page)
                self.show_page()
        
    def zoom_in(self):
        self.current_zoom *= 1.1
        self.show_page()

    def zoom_out(self):
        self.current_zoom /= 1.1
        self.show_page()


if __name__ == "__main__":
    root = Tk()
    pdf_viewer = PDFViewer(root)
    root.mainloop()
