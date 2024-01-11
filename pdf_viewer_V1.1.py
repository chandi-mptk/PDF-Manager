from tkinter import (
    Tk, Canvas, Scrollbar, Frame, Label, HORIZONTAL, 
    VERTICAL, Menu, DISABLED, NORMAL, LEFT, RIGHT, SOLID,
    Y, Listbox, SINGLE, Button, Toplevel, Entry, StringVar
    )
from tkinter.filedialog import askopenfilename, asksaveasfilename
from tkinter.simpledialog import askstring
from tkinter.messagebox import askyesnocancel
from tkinter.messagebox import showinfo, showerror, showwarning
from PIL import Image, ImageTk
import fitz  # PyMuPDF


class PDFViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Viewer")
        self.height, self.width = self.root.winfo_screenheight(), self.root.winfo_screenwidth()
        self.root.geometry(f"{self.width}x{self.height}")

        # Override clos button to quit_prog
        self.root.protocol("WM_DELETE_WINDOW", self.quit_prog)

        # Variables
        self.pdf_document = None
        self.current_page = 0
        self.current_zoom = 1.0
        self.num_pages = 0
        self.password = None

        # Create menubar
        self.create_menu()

        # Create Left Frame
        self.create_left_frame()

        # Create Right Frame
        self.create_right_frame()

        # Create Canvas Frame
        self.create_canvas_frame()

    # Create Menu Bar and Menu Items
    def create_menu(self):
        self.menubar = Menu(self.root)
        self.root.config(menu=self.menubar)

        # Create File menu with Open and Exit option with separator
        self.file_menu = Menu(self.menubar, tearoff=0)
        self.file_menu.add_command(label="Open", command=self.load_pdf, state=NORMAL)
        self.file_menu.add_command(label="Close", command=self.close_pdf, state=DISABLED)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Exit", command=self.quit_prog)
        self.menubar.add_cascade(label="File", menu=self.file_menu)

    # Create Left Frame for page Preview and page list
    def create_left_frame(self):
        # Left Frame for page Preview and page list
        self.left_frame = Frame(root, width=self.width // 20)
        self.left_frame.pack(side=LEFT, fill=Y)

        # Left Frame Label
        self.page_list_label = Label(self.left_frame, text="Page List")
        self.page_list_label.pack(pady=5)

        # Left Frame List Box for Page List 
        self.page_listbox = Listbox(self.left_frame, selectmode=SINGLE, exportselection=0)
        self.page_listbox.pack(fill=Y, expand=True)

    # Create Right Frame for tool Buttons
    def create_right_frame(self):
        # Right Frame for tool Buttons
        self.right_frame = Frame(root, width=self.width // 10)
        self.right_frame.pack(side=RIGHT, fill=Y)

        # Right Frame Widget
        # Zoom Frame
        self.zoom_frame = Frame(self.right_frame, padx=10, pady=10, bd=2, relief=SOLID)
        self.zoom_frame.pack(padx=10, pady=10)

        # Zoom Buttons Area
        # Heading
        self.zoom_label = Label(self.zoom_frame, text="Zoom")
        self.zoom_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

        self.zoom_in_button = Button(self.zoom_frame, text="Zoom In", command=self.zoom_in, state=DISABLED)
        self.zoom_in_button.grid(row=1, column=0, padx=5, pady=5)
        self.zoom_out_button = Button(self.zoom_frame, text="Zoom Out", command=self.zoom_out, state=DISABLED)
        self.zoom_out_button.grid(row=1, column=1, padx=5, pady=5)
    
    # Create Main Frame to show PDF File Content
    def create_canvas_frame(self):
        # Center View Area for viewing File
        self.canvas_frame = Frame(root)
        self.canvas_frame.pack(fill="both", expand=True, padx=10)

        # Main Frame Widgets
        self.canvas = Canvas(self.canvas_frame, state=DISABLED)
        self.canvas.pack(side="left", fill="both", expand=True)

        self.scrollbar_y = Scrollbar(self.canvas_frame, orient=VERTICAL, command=self.canvas.yview)
        self.scrollbar_y.pack(side="right", fill="y")

        self.scrollbar_x = Scrollbar(root, orient=HORIZONTAL, command=self.canvas.xview)
        self.scrollbar_x.pack(side="bottom", fill="x")

        self.canvas.configure(yscrollcommand=self.scrollbar_y.set, xscrollcommand=self.scrollbar_x.set)

        self.page_label = Label(root, text="Page:")
        self.page_label.pack(side="top", pady=5)

    # Menu item and Button Status Setter
    def status_manager(self):

        # If Document Available
        if self.pdf_document:

            # Disable Open Menu Enable Close Menu
            self.file_menu.entryconfig("Open", state=DISABLED)
            self.file_menu.entryconfig("Close", state=NORMAL)
            
            # Bind Scrool Button Function
            self.page_listbox.bind("<ButtonRelease>", self.scroll_page_scale)
            self.canvas.bind("<ButtonRelease>", self.scroll_page_scale)

            # Enable Zoon in and Zoom Out Button
            self.zoom_in_button.config(state=NORMAL)
            self.zoom_out_button.config(state=NORMAL)
        
        # If Document Not Availabl
        else:

            # Enable Open Menu Disable Close Menu
            self.file_menu.entryconfig("Open", state=NORMAL)
            self.file_menu.entryconfig("Close", state=DISABLED)

            # Unbind Scroll Button Function
            self.page_listbox.unbind("<ButtonRelease>")
            self.canvas.unbind("<ButtonRelease>")

            # Disable Zoom in and zoom out Buttons
            self.zoom_in_button.config(state=DISABLED)
            self.zoom_out_button.config(state=DISABLED)

    # Load PDF file when Clicked on Open Menu item
    def load_pdf(self):

        # Open File 
        file_path = askopenfilename(filetypes=[("PDF Files", "*.pdf")])

        # If File is Selected
        if file_path:

            # Load PyMuPDF Object to self.pdf_document
            self.pdf_document = fitz.open(file_path)

            # Check if pdf document is password protected
            if self.pdf_document.is_encrypted:
                def password_validate(event=None):
                    self.password = password.get()
                    password.set("")
                    password_popup.unbind("<Return>")
                    password_popup.unbind("<Escape>")
                    password_popup.destroy()

                def password_cancel(event=None):
                    self.password = None
                    self.pdf_document = None
                    password_popup.unbind("<Return>")
                    password_popup.unbind("<Escape>")
                    password_popup.destroy()
                    
                
                # Infinite Loop
                while self.pdf_document:
                    password = StringVar()

                    # Ask Password Using Toplevel
                    password_popup = Toplevel(self.root)

                    password_popup.title("Password")
                    password_label = Label(password_popup, text="Enter the password for the PDF file:")
                    password_label.grid(row=0, columnspan=2, padx=5, pady=5)

                    password_entry = Entry(password_popup, show="*", textvariable=password)
                    password_entry.grid(row=1, columnspan=2, padx=5, pady=5)
                    password_entry.focus()

                    password_submit_button = Button(password_popup, text="Submit", command=password_validate)
                    password_submit_button.grid(row=2, column=0, pady=10)
                    password_cancel_button = Button(password_popup, text="Cancel", command=password_cancel)
                    password_cancel_button.grid(row=2, column=1, pady=10)

                    # Bind enter key to password_validate function escape key to password_cancel function
                    password_popup.bind("<Return>", password_validate)
                    password_popup.bind("<Escape>", password_cancel)

                    password_popup.grab_set()
                    password_popup.wait_window()
                    
                    # Attempt to decrypt the document
                    if self.pdf_document.authenticate(self.password) > 0:
                        break
                    else:
                        showerror("Error", "Incorrect password. Please try again.")
                    
            if self.pdf_document is not None:
                self.num_pages = self.pdf_document.page_count
                self.root.title(f"PDF Viewer - {file_path}")
                self.show_page()
                self.status_manager()

    def close_pdf(self):
        
        # Set all Variables to default
        self.pdf_document = None
        self.current_zoom = 1.0
        self.num_pages = 0
        
        # Destroy All Data Showing Areas and Clear data
        self.page_listbox.delete(0, 'end')
        self.canvas.delete("all")
        self.page_label.config(text="Page:")
        self.canvas.config(scrollregion=(0, 0, 0, 0))
        self.canvas.image = None
        
        # Change Title to Default
        self.root.title("PDF Viewer")
        self.status_manager()

    # Show PDF File Data To Relevant areas
    def show_page(self, *args):
        if self.pdf_document:

            # Populate the page listbox
            self.page_listbox.delete(0, 'end')
            for page_number in range(1, self.num_pages + 1):
                self.page_listbox.insert(page_number, f"Page {page_number}")
            
            # Update page listbox selection
            self.page_listbox.selection_clear(0, 'end')
            self.page_listbox.selection_set(self.current_page)
            self.page_listbox.activate(self.current_page)

            # Update Canvas with the image
            page = self.pdf_document[self.current_page]
            pix = page.get_pixmap()

            # Set Zoom Height and width
            zoomed_width = int(pix.width * self.current_zoom)
            zoomed_height = int(pix.height * self.current_zoom)

            # Create Image
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img = img.resize((zoomed_width, zoomed_height), Image.ADAPTIVE)
            img = ImageTk.PhotoImage(img)

            # off set dimensions
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

    # Scroll Button 
    def scroll_page_scale(self, event):
        # If Button 1 Released
        if event.num == 1:
            selected_index = self.page_listbox.curselection()
            if selected_index:
                self.current_page = selected_index[0]
                self.show_page()
                self.status_manager()

        # If Scroll UP
        if event.num == 4:
            # Check if the current value is greater than 1
            if self.current_page > 0:
                # Decrement the value by 1
                self.current_page -= 1
                self.show_page()
                self.status_manager()

        # If Scroll Down
        elif event.num == 5:
            # Check if the current value is less than the maximum value
            if self.current_page < self.num_pages - 1:
                # Increment the value by 1
                self.current_page += 1
                self.show_page()
                self.status_manager()

    # Zoom in
    def zoom_in(self):
        self.current_zoom *= 1.1
        self.show_page()
        self.status_manager()

    # Zoom out
    def zoom_out(self):
        self.current_zoom /= 1.1
        self.show_page()
        self.status_manager()

    # Quit Program
    def quit_prog(self):
        self.close_pdf()
        self.root.destroy()

if __name__ == "__main__":
    root = Tk()
    pdf_viewer = PDFViewer(root)
    root.mainloop()
