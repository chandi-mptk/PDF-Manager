from tkinter import (
    Tk, Canvas, Scrollbar, Frame, Label, HORIZONTAL, 
    VERTICAL, Menu, DISABLED, NORMAL, LEFT, RIGHT, SOLID,
    Y, Listbox, SINGLE, Button, StringVar, Entry, Toplevel
    )
from tkinter.filedialog import askopenfilename, asksaveasfilename
from tkinter.simpledialog import askstring, askfloat
from tkinter.messagebox import askyesnocancel
from tkinter.messagebox import showinfo, showerror, showwarning
from pathlib import Path
from io import BytesIO
from PIL import Image, ImageTk
import fitz  # PyMuPDF


class GUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Viewer")
        self.height, self.width = self.root.winfo_screenheight(), self.root.winfo_screenwidth()
        self.root.geometry(f"{self.width // 2}x{self.height // 2}+0+0")

        # Override clos button to quit_prog
        self.root.protocol("WM_DELETE_WINDOW", self.quit_prog)

        # Variables
        self.pdf_object = PDFManager()
        self.current_page = 0
        self.current_zoom = 1.0
        self.num_pages = 0
        self.is_changed = False
        self.file_path = None

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

        # Create File menu with Open, Save, Close and Exit option with separator
        self.file_menu = Menu(self.menubar, tearoff=0)
        self.file_menu.add_command(label="Open", command=self.load_pdf, state=NORMAL)
        self.file_menu.add_command(label="Save", command=self.save_pdf, state=DISABLED)
        self.file_menu.add_command(label="Close", command=self.close_pdf, state=DISABLED)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Exit", command=self.quit_prog)
        self.menubar.add_cascade(label="File", menu=self.file_menu)

        # Create Merge Menu with Add to First and Add to Last
        self.merge_menu = Menu(self.menubar, tearoff=0)
        self.merge_menu.add_command(label="Add to First", command=lambda: self.merge_files(at_end=False))
        self.merge_menu.add_command(label="Add to Last", command=lambda: self.merge_files(at_end=True))
        self.menubar.add_cascade(label="Merge", menu=self.merge_menu, state=DISABLED)

        # Create Split Menu with range
        self.split_menu = Menu(self.menubar, tearoff=0)
        self.split_menu.add_command(label="Split by Range", command=self.split_pdf)
        self.menubar.add_cascade(label="Split", menu=self.split_menu, state=DISABLED)

        # Create Compress Menu with Compress
        self.compress_menu = Menu(self.menubar, tearoff=0)
        self.compress_menu.add_command(label="Compress", command=self.compress_pdf)
        self.menubar.add_cascade(label="Compress", menu=self.compress_menu, state=DISABLED)

        # Create About Menu with About
        self.about_menu = Menu(self.menubar, tearoff=0)
        self.about_menu.add_command(label="About", command=self.about_prog)
        self.menubar.add_cascade(label="About", menu=self.about_menu)

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

        # Delete Page Frame
        self.delete_page_frame = Frame(self.right_frame, padx=10, pady=10, bd=2, relief=SOLID)
        self.delete_page_frame.pack(padx=10, pady=10)
        
        # Delete Page Button Area
        # Heading
        self.delete_page_label = Label(self.delete_page_frame, text="Delete Page")
        self.delete_page_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5)
        
        # Delete Page Button
        self.delete_page_button = Button(self.delete_page_frame, text="Delete Page", command=self.delete_page_confirmation, state=DISABLED)
        self.delete_page_button.grid(row=1, column=0, columnspan=2, padx=5, pady=5)

        # Password Frame
        self.password_frame = Frame(self.right_frame, padx=10, pady=10, bd=2, relief=SOLID)
        self.password_frame.pack(padx=10, pady=10)

        # Password Button Area
        # Heading
        self.password_label = Label(self.password_frame, text="Password")
        self.password_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5)

        # Set Password Button
        self.password_button = Button(self.password_frame, text="Set Password", command=self.set_password, state=DISABLED)
        self.password_button.grid(row=1, column=0, columnspan=2, padx=5, pady=5)
    
        # Add File Frame
        self.add_file_frame = Frame(self.right_frame, padx=10, pady=10, bd=2, relief=SOLID)
        self.add_file_frame.pack(padx=10, pady=10)        

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

        self.page_label = Label(root, text=f"Page:")
        self.page_label.pack(side="top", pady=5)

    # Menu item and Button Status Setter
    def status_manager(self):

        # If Document Available
        if self.pdf_object.is_file_available:

            # Disable Open Menu Enable Close Menu
            self.file_menu.entryconfig("Open", state=DISABLED)
            self.file_menu.entryconfig("Close", state=NORMAL)

            # Enable Merge Menu
            if not self.pdf_object.need_password:
                self.menubar.entryconfig("Merge", state=NORMAL)

            # Enable Split Menu if More than 1 Page
            if self.num_pages > 1:
                self.menubar.entryconfig("Split", state=NORMAL)
            
            # Enable Compress Menu
            self.menubar.entryconfig("Compress", state=NORMAL)
            
            # Bind Scroll Button Function
            self.page_listbox.bind("<ButtonRelease>", self.scroll_page_scale)
            self.canvas.bind("<ButtonRelease>", self.scroll_page_scale)

            # Enable Zoon in and Zoom Out Button
            self.zoom_in_button.config(state=NORMAL)
            self.zoom_out_button.config(state=NORMAL)

            # Enable Delete Button if More than 1 Page    
            if self.num_pages > 1:
                self.delete_page_button.config(state=NORMAL)

            # Enable Set Password Button if need_password is False
            if not self.pdf_object.need_password:
                self.password_button.config(state=NORMAL)
        
        # If Document Not Available
        else:

            # Enable Open Menu Disable Close Menu
            self.file_menu.entryconfig("Open", state=NORMAL)
            self.file_menu.entryconfig("Close", state=DISABLED)

            # Disable Merge Menu
            self.menubar.entryconfig("Merge", state=DISABLED)

            # Disable Split Menu
            self.menubar.entryconfig("Split", state=DISABLED)

            # Disable Compress Menu
            self.menubar.entryconfig("Compress", state=DISABLED)

            # Unbind Scroll Button Function
            self.page_listbox.unbind("<ButtonRelease>")
            self.canvas.unbind("<ButtonRelease>")

            # Disable delete, Set Password, Zoom in and zoom out Buttons
            self.zoom_in_button.config(state=DISABLED)
            self.zoom_out_button.config(state=DISABLED)
            self.delete_page_button.config(state=DISABLED)
            self.password_button.config(state=DISABLED)

        # If Change in Document Enable Save Button else Disable
        if self.is_changed:
            self.file_menu.entryconfig("Save", state=NORMAL)
        else:
            self.file_menu.entryconfig("Save", state=DISABLED)

    # Load PDF file when Clicked on Open Menu item
    def load_pdf(self):

        # Open File 
        self.file_path = askopenfilename(filetypes=[("PDF Files", "*.pdf")])

        # If File is not Selected
        if not self.file_path:
            return
        
        # Load pdf file to pdf object
        self.pdf_object.load_pdf(self.file_path)

        # Check if pdf document is password protected
        if self.pdf_object.is_encrypted:
            password = self.password_prompt()

            # If Canceled or Not entered any Data
            if not password:
                self.pdf_object = PDFManager()
                showerror("Error", "Password is required to open the PDF file.")
                return
            
            # Try to Decrypt using password
            if not self.pdf_object.decrypt_pdf(password):
                showerror("Error", "Invalid password. PDF File cannot be opened.")
                return
            
        if self.pdf_object.is_file_available:
            self.num_pages = self.pdf_object.get_page_count
            self.root.title(f"PDF Viewer - {self.file_path}")
            self.show_page()
            self.status_manager()

    # Prompt Password
    def password_prompt(self):
        password_var = StringVar()

        def password_validate(event=None):
            password_popup.unbind("<Return>")
            password_popup.unbind("<Escape>")
            password_popup.destroy()

        def password_cancel(event=None):
            password_var.set("")
            password_popup.unbind("<Return>")
            password_popup.unbind("<Escape>")
            password_popup.destroy()

        # Ask Password Using Toplevel
        password_popup = Toplevel(self.root)
        password_popup.attributes("-topmost", True)
        password_popup.title("Password")
        password_label = Label(password_popup, text="Enter the password for the PDF file:")
        password_label.grid(row=0, columnspan=2, padx=5, pady=5)

        password_entry = Entry(password_popup, show="*", textvariable=password_var)
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
        
        return password_var.get()
    
    # Close PDF Document
    def close_pdf(self):

        # If any Change 
        if self.is_changed:

            # Ask to Save
            result = askyesnocancel("Warning", "The document has been changed. Do you want to save the changes?")
            if result == True:
                self.save_pdf()
                
            elif result == False:
                self.is_changed = False
            
        if not self.is_changed:

            # Set all Variables to default
            self.pdf_object = PDFManager()
            self.current_page = 0
            self.current_zoom = 1.0            
            self.num_pages = 0
            self.is_changed = False
            self.file_path = None

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
        if not self.pdf_object.is_file_available:
            return
        
        self.page_label.config(text=f"Page: {self.current_page + 1}")
        # Populate the page listbox
        self.page_listbox.delete(0, 'end')
        for page_number in range(1, self.num_pages + 1):
            self.page_listbox.insert(page_number, f"Page {page_number}")
        
        # Update page listbox selection
        self.page_listbox.selection_clear(0, 'end')
        self.page_listbox.selection_set(self.current_page)
        self.page_listbox.activate(self.current_page)

        # Update Canvas with the image
        page = self.pdf_object.get_page(self.current_page)
        pix = page.get_pixmap()

        # Set Zoom Height and width
        zoomed_width = int(pix.width * self.current_zoom)
        zoomed_height = int(pix.height * self.current_zoom)

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

    # Zoom Out
    def zoom_out(self):
        self.current_zoom /= 1.1
        self.show_page()
        self.status_manager()

    # PDF Document Page Delete Confirmation
    def delete_page_confirmation(self):

        # Only 1 Page Show Error
        if self.num_pages == 1:
            showerror("Error", "Cannot delete the only page.")
            return
        
        # Delete the page self.current_page from the PDF document
        self.pdf_object.delete_page_no(self.current_page)
        self.is_changed = True

        # Update the number of pages
        self.num_pages -= 1

        # Update the current page
        if self.current_page >= self.num_pages:
            self.current_page = self.num_pages - 1
        self.show_page()
        self.status_manager()

    # Set Password to PDF file
    def set_password(self):
        
        if self.pdf_object.need_password:
            showerror('Error', "File already have a Password.")
            return
        
        # Variables
        owner_pass = StringVar()
        user_pass = StringVar()

        def submit(event=None):
            if user_pass.get() == '' or owner_pass.get() == '':
                showerror('Error', "Password cannot be empty.")
                password_popup.destroy()
                
            elif user_pass.get() == owner_pass.get():
                showerror('Error', 'Both Passwords Cannot be Same.')
                password_popup.destroy()
                
            else:
                password_popup.destroy()
                self.save_pdf(owner_pass.get(), user_pass.get())

        def cancel(event=None):
            password_popup.destroy()
            
        password_popup = Toplevel(self.root)
        password_popup.attributes('-topmost', True)
        password_popup.title("Set Password")
        password_label = Label(password_popup, text="Enter the password for the PDF file:")
        password_label.grid(row=0, columnspan=2, padx=5, pady=5)

        owner_password_label = Label(password_popup, text="Owner Password:")
        owner_password_label.grid(row=1, column=0, padx=5, pady=5)

        owner_password_entry = Entry(password_popup, show="*", textvariable=owner_pass)
        owner_password_entry.grid(row=1, column=1, padx=5, pady=5)

        owner_password_entry.focus()

        user_password_label = Label(password_popup, text="User Password:")
        user_password_label.grid(row=2, column=0, padx=5, pady=5)

        user_password_entry = Entry(password_popup, show="*", textvariable=user_pass)
        user_password_entry.grid(row=2, column=1, padx=5, pady=5)

        password_submit_button = Button(password_popup, text="Submit", command=submit)
        password_submit_button.grid(row=3, column=0, padx=5, pady=10)
        password_cancel_button = Button(password_popup, text="Cancel", command=cancel)
        password_cancel_button.grid(row=3, column=1, padx=5, pady=10)

        password_popup.bind('<Return>', submit)
        password_popup.bind('<Escape>', cancel)

        password_popup.grab_set()
        password_popup.wait_window()

    # Merge Files
    def merge_files(self, at_end):
        
        # Open File 
        self.file_path = askopenfilename(filetypes=[("PDF Files", "*.pdf")])

        # If File is not Selected
        if not self.file_path:
            return
        pdf_obj_2 = PDFManager()
        pdf_obj_2.load_pdf(self.file_path)

        if pdf_obj_2.need_password:
            showerror('Error', 'Password Protected Files cannot be merged')
            return
        
        # Add File to first
        self.pdf_object.merge_pdf(pdf_obj_2, at_end=at_end)
        self.num_pages = self.pdf_object.get_page_count
        self.is_changed = True
        self.show_page()
    
    # Split PDF File into page ranges
    def split_pdf(self):
        if self.is_changed:
            showwarning("Warning", "Please Save all Changes Before Splitting the File.")
            return
        
        # Get Split ranges from user
        split_ranges = askstring("Split Ranges", f"Enter the split ranges (e.g. 1-3,5-5,7-9) Last Page is {self.num_pages}:")

        if split_ranges:
            split_ranges.strip()
            
            # Check ' in split ranges
            if ',' in split_ranges:

                # Split ranges by ',' and save in split_ranges_list
                split_ranges_list = split_ranges.split(',')
                split_ranges_list = [item.strip() for item in split_ranges_list]
            else:
                split_ranges_list = [split_ranges]

            try:
                # Convert split ranges to list of integers tuples which are separated by -
                split_ranges_list = [tuple(map(int, item.split('-'))) for item in split_ranges_list]
            except ValueError:
                showerror("Error", "Invalid split ranges. (e.g. 1-3,5-5,7-9)")
                return
            
            invalid_ranges = []

            for item in split_ranges_list[:]:
                
                # Check if tuple first element is greater than end page
                if 1 <= item[0] <= self.num_pages and 1 <= item[1] <= self.num_pages and item[0] <= item[1]:
                    
                    self.pdf_object.backup_pdf

                    # Split PDF
                    self.pdf_object.split_pdf(item)
                    
                    self.save_pdf(initial_file=f"{Path(self.file_path).stem}_{item[0]}_{item[1]}.pdf")
                    self.is_changed = False

                    self.pdf_object.restore_pdf
                
                else:
                    showerror("Error", f"Invalid split ranges.{item[0]}-{item[1]} not valid in {self.num_pages} Page PDF Document")
                    continue
    
    # Compress PDF File
    def compress_pdf(self):
        garbage = 4
        clean = True
        deflate = True
        deflate_images = True
        deflate_fonts = True
        self.save_pdf(garbage=garbage, clean=clean, deflate=deflate, deflate_images=deflate_images, deflate_fonts=deflate_fonts)
    
    # Save PDF File
    def save_pdf(self,
                output_file=None,
                 garbage=0, 
                 clean=False, 
                 deflate=False, 
                 deflate_images=False, 
                 deflate_fonts=False, 
                 owner_pass=None, 
                 user_pass=None, 
                 initial_file=None):

        if output_file:
            output_file = str(Path(self.file_path).parent / output_file)
        else:

            # Get Output File
            output_file = asksaveasfilename(title="Save PDF", filetypes=[("PDF Files", "*.pdf")], initialfile=initial_file)
        

        if output_file:

            # IF file name not ending with PDF then Set
            if not output_file.endswith(".pdf"):
                output_file += ".pdf"

            # If owner or user password set then set encryption as fitz.PDF_ENCRYPT_AES_256
            if owner_pass or user_pass:
                encryption = fitz.PDF_ENCRYPT_AES_256
            else:
                encryption = fitz.PDF_ENCRYPT_KEEP
            
            # Get Boolean and String tuple as result
            result =  self.pdf_object.save_file(output_file, 
                                                garbage=garbage, 
                                                clean=clean, 
                                                deflate=deflate, 
                                                deflate_images=deflate_images, 
                                                deflate_fonts=deflate_fonts, 
                                                owner_pass=owner_pass, 
                                                user_pass=user_pass, 
                                                encryption=encryption)
            
            # if Boolean result is True Show Success Message
            if result[0]:
                showinfo("Success", f"PDF file saved successfully as {Path(output_file).name}")
                self.is_changed = False
                self.status_manager()
                return
            
            # If Save Failed
            else:

                # check the error message in result tuple is save to original must be incremental
                # Then show error message as Please do not try to overwrite Source file.
                if result[1] == "save to original must be incremental":
                    showerror("Error", "Please do not try to overwrite Source file.")
                    
                else:
                    showerror("Error", "Failed to save PDF file.")
                self.is_changed = True
                self.status_manager()
                return
            
        self.is_changed = False
        self.status_manager()

    # About Menu
    def about_prog(self):
        about_popup = Toplevel(self.root)
        about_popup.title("About")
        # about_popup.geometry("300x200")
        about_popup.resizable(False, False)
        about_popup.grab_set()
        about_popup.focus_set()
        about_popup.transient(self.root)
        about_popup.protocol("WM_DELETE_WINDOW", lambda: about_popup.destroy())
        about_popup.bind('<Return>', lambda event: about_popup.destroy())
        about_popup.bind('<Escape>', lambda event: about_popup.destroy())

        about_label = Label(about_popup, text="PDF Manager V2.0", font=("Arial", 16))
        about_label.pack(pady=10)

        about_label = Label(about_popup, text="Developed By: Nishanth A C", font=("Arial", 12))
        about_label.pack(pady=10)
        about_label = Label(about_popup, text="Email: nishanth.ac@gmail.com", font=("Arial", 12))
        about_label.pack(pady=10)

    # Quit Program
    def quit_prog(self):
        self.close_pdf()
        if not self.is_changed:
            self.root.destroy()


class PDFManager:
    def __init__(self):
        self.pdf_document = None
        self.temp_document = None

    def load_pdf(self, file_path):
        self.pdf_document = fitz.open(file_path)

    @property
    def get_page_count(self):
        return self.pdf_document.page_count if self.is_file_available else 0
        
    @property
    def is_encrypted(self):
        return self.pdf_document.is_encrypted
    
    @property
    def is_file_available(self):
        return self.pdf_document is not None
    
    @property
    def need_password(self):
        return self.pdf_document.needs_pass

    @property
    def is_image_available(self):
        for page_number in range(self.get_page_count):
            image_list = self.pdf_document.get_page_images(page_number, full=True)
            if image_list:
                return True
            return False

    def decrypt_pdf(self, password):
        return self.pdf_document.authenticate(password)
    
    def get_page(self, page_number):
        return self.pdf_document[page_number]
    
    def delete_page_no(self, page_number):
        self.pdf_document.delete_page(page_number)

    def merge_pdf(self, pdf_doc, at_end):
        if at_end:
            self.pdf_document.insert_pdf(pdf_doc.pdf_document)
        else:
            self.pdf_document.insert_pdf(pdf_doc.pdf_document, start_at=0)

    def split_pdf(self, split_range):
        self.temp_document.insert_pdf(self.pdf_document, from_page=split_range[0] - 1, to_page=split_range[1] - 1)
        
    def save_file(self, output_file, 
                  garbage=0, 
                  clean=False, 
                  deflate=False, 
                  deflate_images=False, 
                  deflate_fonts=False, 
                  owner_pass=None, 
                  user_pass=None, 
                  encryption=fitz.PDF_ENCRYPT_KEEP):
        try:
            self.pdf_document.save(output_file, 
                                garbage=garbage,
                                clean=clean,
                                deflate=deflate,
                                deflate_images=deflate_images,
                                deflate_fonts=deflate_fonts,
                                encryption=encryption,
                                owner_pw=owner_pass, 
                                user_pw=user_pass)
            return True, ""
        except Exception as e:
            return False, str(e)
    


if __name__ == "__main__":
    root = Tk()
    pdf_viewer = GUI(root)
    root.mainloop()
