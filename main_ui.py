import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import platform
from ttkbootstrap import Style, DateEntry
from datetime import datetime


class MainUI:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.entries = {}
        self.document_entries = {}
        self.doc_frame_created = False
        self.os_type = platform.system()
        self.date_fields = [
            "Date ",
            "commenment Date",
            "Lease Expiry",
            "Lock in expiry",
        ]
        self.setup_styles()
        self.create_main_ui()
        self.setup_scrolling()

    def setup_styles(self):
        self.style = Style(theme='flatly')

        # Enhanced color palette
        self.colors = {
            'primary': '#2c3e50',
            'secondary': '#34495e',
            'accent': '#455a64',
            'background': '#f8f9fa',
            'surface': '#ffffff',
            'border': '#e9ecef',
            'text': '#2c3e50',
            'text_secondary': '#6c757d'
        }

        # Configure enhanced styles
        self.style.configure('Title.TLabel',
                             font=('Segoe UI', 24, 'normal'),
                             foreground=self.colors['primary'],
                             background=self.colors['surface'])

        self.style.configure('Subtitle.TLabel',
                             font=('Segoe UI', 11),
                             foreground=self.colors['text_secondary'],
                             background=self.colors['surface'])

        self.style.configure('Card.TFrame',
                             background=self.colors['surface'],
                             relief='flat')

        self.style.configure('Content.TFrame',
                             background=self.colors['surface'])

        self.style.configure('Section.TLabel',
                             font=('Segoe UI', 16, 'bold'),
                             foreground=self.colors['primary'],
                             background=self.colors['surface'],
                             padding=(0, 10))

        self.style.configure('Field.TLabel',
                             font=('Segoe UI', 11),
                             foreground=self.colors['text'],
                             background=self.colors['surface'])

        self.style.configure('Field.TEntry',
                             font=('Segoe UI', 11),
                             fieldbackground=self.colors['surface'])

        # Configure Combobox style
        self.style.configure('TCombobox',
                             font=('Segoe UI', 11),
                             fieldbackground=self.colors['surface'])

    def create_main_ui(self):
        # Main container with professional padding
        main_container = ttk.Frame(self.parent, style='Card.TFrame')
        main_container.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)

        # Enhanced header section
        header_frame = ttk.Frame(main_container, style='Content.TFrame')
        header_frame.pack(fill=tk.X, pady=(20, 30), padx=20)

        title = ttk.Label(header_frame,
                         text="Enterprise Asset Management",
                         style='Title.TLabel')
        title.pack(anchor=tk.W)

        subtitle = ttk.Label(header_frame,
                           text="Add and manage enterprise assets with supporting documentation",
                           style='Subtitle.TLabel')
        subtitle.pack(anchor=tk.W, pady=(5, 0))

        # Scrollable content with improved styling
        self.container_frame = ttk.Frame(main_container, style='Content.TFrame')
        self.container_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(self.container_frame,
                              bg=self.colors['surface'],
                              highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.container_frame,
                                     orient="vertical",
                                     command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas, style='Content.TFrame')

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        canvas_window = self.canvas.create_window(
            (0, 0),
            window=self.scrollable_frame,
            anchor="nw"
        )

        self.canvas.bind(
            '<Configure>',
            lambda e: self.canvas.itemconfig(canvas_window, width=e.width)
        )

        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Create content sections
        self.create_input_fields()

        # Action buttons with enhanced styling
        button_frame = ttk.Frame(self.scrollable_frame, style='Content.TFrame')
        button_frame.pack(fill=tk.X, pady=20, padx=30)

        self.upload_docs_button = ttk.Button(
            button_frame,
            text="Upload Documents",
            style="primary.TButton",
            command=self.show_document_upload
        )
        self.upload_docs_button.pack(side=tk.LEFT, padx=5)

        self.submit_button = ttk.Button(
            button_frame,
            text="Submit Asset",
            style="success.TButton",
            command=self.submit_data
        )
        self.submit_button.pack(side=tk.LEFT, padx=5)
        self.submit_button.pack_forget()

    def create_input_fields(self):
        # Enhanced section header
        section_frame = ttk.Frame(self.scrollable_frame, style='Content.TFrame')
        section_frame.pack(fill=tk.X, pady=(0, 20), padx=30)

        section_label = ttk.Label(section_frame,
                                  text="Asset Information",
                                  style='Section.TLabel')
        section_label.pack(side=tk.LEFT)

        # Main content with card-like appearance
        content_frame = ttk.Frame(self.scrollable_frame,
                                  style='Card.TFrame',
                                  padding=25)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=(0, 20))

        # Two-column layout with proper spacing
        columns_frame = ttk.Frame(content_frame, style='Content.TFrame')
        columns_frame.pack(fill=tk.BOTH, expand=True)

        # Create columns with equal width
        left_column = ttk.Frame(columns_frame, style='Content.TFrame')
        right_column = ttk.Frame(columns_frame, style='Content.TFrame')

        # Configure equal weights for columns
        columns_frame.columnconfigure(0, weight=1)
        columns_frame.columnconfigure(1, weight=1)

        left_column.grid(row=0, column=0, sticky='nsew', padx=(0, 15))
        right_column.grid(row=0, column=1, sticky='nsew', padx=(15, 0))

        # Add the new fields to the headers list
        headers_to_use = self.app.headers[:20] + [
            self.app.headers[20],
            "Lease Manager",
            "Tenant type",
            "Brokerage"
        ]
        midpoint = (len(headers_to_use) + 1) // 2

        self.entry_order = []

        # Configure styles for consistent appearance
        field_width = 35  # Consistent width for all fields
        label_width = 20  # Consistent width for all labels

        # Configure grid weights for the columns
        left_column.columnconfigure(1, weight=1)
        right_column.columnconfigure(1, weight=1)

        for i, header in enumerate(headers_to_use):
            parent_frame = left_column if i < midpoint else right_column
            row_index = i if i < midpoint else i - midpoint

            field_frame = ttk.Frame(parent_frame, style='Content.TFrame')
            field_frame.grid(row=row_index, column=0, sticky='ew', pady=8)
            field_frame.columnconfigure(1, weight=1)  # Make entry column expandable

            label = ttk.Label(field_frame,
                              text=header,
                              style='Field.TLabel',
                              width=label_width,
                              anchor='e')  # Right-align labels
            label.grid(row=0, column=0, padx=(0, 10), sticky='e')

            if header == "Brokerage":
                # Create dropdown for Brokerage
                entry = ttk.Combobox(field_frame,
                                     values=["Pending", "Received"],
                                     state="readonly",
                                     width=field_width)
                entry.grid(row=0, column=1, sticky='ew')

            elif header in self.date_fields:
                # Create DateEntry with MM-DD-YYYY format and consistent styling
                entry = DateEntry(
                    field_frame,
                    width=field_width,
                    bootstyle="primary",
                    dateformat="%m-%d-%Y",
                    firstweekday=6,
                    startdate=None
                )
                entry.entry.delete(0, tk.END)
                # Apply consistent font to the entry part
                entry.entry.configure(font=('Segoe UI', 11))
                entry.grid(row=0, column=1, sticky='ew')

            else:
                # Regular entry for non-date fields
                entry = ttk.Entry(field_frame,
                                  style='Field.TEntry',
                                  width=field_width)
                entry.grid(row=0, column=1, sticky='ew')

            self.entries[header] = entry
            self.entry_order.append(entry)

            # Set tab order
            entry.bind('<Tab>', self._handle_tab)
            entry.bind('<Shift-Tab>', self._handle_shift_tab)

        # Bind keyboard navigation
        self._setup_keyboard_navigation()

    def _setup_keyboard_navigation(self):
        def handle_arrow_keys(event):
            current_widget = self.parent.focus_get()

            # Check if the current widget is one of our entry fields
            # or is the entry part of a DateEntry
            current_index = -1
            for i, widget in enumerate(self.entry_order):
                if isinstance(widget, DateEntry):
                    if widget.entry == current_widget:
                        current_index = i
                        break
                elif widget == current_widget:
                    current_index = i
                    break

            if current_index != -1:
                if event.keysym == 'Up' and current_index > 0:
                    # Move to previous entry
                    next_widget = self.entry_order[current_index - 1]
                    if isinstance(next_widget, DateEntry):
                        next_widget.entry.focus_set()
                    else:
                        next_widget.focus_set()
                    return 'break'

                elif event.keysym == 'Down' and current_index < len(self.entry_order) - 1:
                    # Move to next entry
                    next_widget = self.entry_order[current_index + 1]
                    if isinstance(next_widget, DateEntry):
                        next_widget.entry.focus_set()
                    else:
                        next_widget.focus_set()
                    return 'break'

        # Bind arrow key navigation to all entry widgets
        for entry in self.entry_order:
            if isinstance(entry, DateEntry):
                entry.entry.bind('<Up>', handle_arrow_keys)
                entry.entry.bind('<Down>', handle_arrow_keys)
            else:
                entry.bind('<Up>', handle_arrow_keys)
                entry.bind('<Down>', handle_arrow_keys)

    def _handle_tab(self, event):
        current = event.widget
        for i, widget in enumerate(self.entry_order):
            if (isinstance(widget, DateEntry) and widget.entry == current) or widget == current:
                next_idx = (i + 1) % len(self.entry_order)
                next_widget = self.entry_order[next_idx]
                if isinstance(next_widget, DateEntry):
                    next_widget.entry.focus_set()
                else:
                    next_widget.focus_set()
                break
        return 'break'

    def _handle_shift_tab(self, event):
        current = event.widget
        for i, widget in enumerate(self.entry_order):
            if (isinstance(widget, DateEntry) and widget.entry == current) or widget == current:
                prev_idx = (i - 1) % len(self.entry_order)
                prev_widget = self.entry_order[prev_idx]
                if isinstance(prev_widget, DateEntry):
                    prev_widget.entry.focus_set()
                else:
                    prev_widget.focus_set()
                break
        return 'break'

    def validate_date_format(self, event, header):
        """Validate and format the date entry"""
        date_widget = self.entries[header]
        try:
            # Get the date from the widget
            date_str = date_widget.entry.get()
            # Parse the date to validate format
            datetime.strptime(date_str, '%m-%d-%Y')
            return True
        except ValueError:
            messagebox.showerror(
                "Invalid Date",
                f"Please enter a valid date in MM-DD-YYYY format for {header}"
            )
            return False

    def _setup_keyboard_navigation(self):
        def handle_arrow_keys(event):
            current_widget = self.parent.focus_get()

            # Check if the current widget is one of our entry fields
            if current_widget in self.entry_order:
                current_index = self.entry_order.index(current_widget)

                if event.keysym == 'Up' and current_index > 0:
                    # Move to previous entry
                    self.entry_order[current_index - 1].focus_set()
                    self.entry_order[current_index - 1].icursor(tk.END)
                    return 'break'

                elif event.keysym == 'Down' and current_index < len(self.entry_order) - 1:
                    # Move to next entry
                    self.entry_order[current_index + 1].focus_set()
                    self.entry_order[current_index + 1].icursor(tk.END)
                    return 'break'

        # Bind arrow key navigation to all entry widgets
        for entry in self.entry_order:
            entry.bind('<Up>', handle_arrow_keys)
            entry.bind('<Down>', handle_arrow_keys)

    def show_document_upload(self):
        if not self.doc_frame_created:
            # Define document headers in order of columns
            self.doc_headers = [
                "KYC",
                "Tenent Verification",
                "Property tax",
                "Lease upload",
                "Cheque PDC"
            ]

            # Section title
            upload_section = ttk.Label(self.scrollable_frame,
                                       text="Document Upload",
                                       style='Section.TLabel')
            upload_section.pack(pady=(20, 20), padx=30)

            # Upload container with refined styling
            self.doc_frame = ttk.Frame(self.scrollable_frame,
                                       style='Card.TFrame',
                                       padding=25)
            self.doc_frame.pack(fill=tk.X, padx=30, pady=(0, 20))

            # Create a grid layout for document uploads
            for header in self.doc_headers:
                frame = ttk.Frame(self.doc_frame, style='Content.TFrame')
                frame.pack(fill=tk.X, pady=12)

                # Document type label with refined styling
                label = ttk.Label(frame,
                                  text=header,
                                  style='Field.TLabel',
                                  width=25)
                label.pack(side=tk.LEFT, padx=(0, 15))

                # Status label to show selected file
                status_label = ttk.Label(frame,
                                         text="No file selected",
                                         style='Subtitle.TLabel')
                status_label.pack(side=tk.LEFT, expand=True, fill=tk.X)

                # Button frame to hold both buttons
                button_frame = ttk.Frame(frame, style='Content.TFrame')
                button_frame.pack(side=tk.RIGHT)

                # NA button
                na_button = ttk.Button(button_frame,
                                       text="NA",
                                       style="secondary.Outline.TButton",
                                       command=lambda h=header, s=status_label: self.mark_as_na(h, s))
                na_button.pack(side=tk.RIGHT, padx=(5, 0))

                # Upload button with refined styling
                upload_button = ttk.Button(button_frame,
                                           text="Browse",
                                           style="secondary.Outline.TButton",
                                           command=lambda h=header, s=status_label: self.upload_file(h, s))
                upload_button.pack(side=tk.RIGHT)

                self.document_entries[header] = None

            # Add submit button at the bottom of document upload section
            submit_frame = ttk.Frame(self.scrollable_frame, style='Content.TFrame')
            submit_frame.pack(fill=tk.X, pady=(20, 30), padx=30)

            self.submit_button = ttk.Button(
                submit_frame,
                text="Submit Asset",
                style="success.TButton",
                command=self.submit_data
            )
            self.submit_button.pack(side=tk.RIGHT)

            self.doc_frame_created = True

    def mark_as_na(self, header, status_label):
        """Mark a document as Not Applicable"""
        status_label.configure(text="NA")
        self.document_entries[header] = "NA"

    def upload_file(self, header, status_label):
        """Handle file upload and update status"""
        file_path = filedialog.askopenfilename(
            title=f"Select a file for {header}",
            filetypes=[
                ("All Files", "*.*"),
                ("PDF Files", "*.pdf"),
                ("Image Files", "*.png *.jpg *.jpeg"),
                ("Document Files", "*.doc *.docx")
            ]
        )
        if file_path:
            status_label.configure(text=os.path.basename(file_path))
            self.document_entries[header] = file_path

    def submit_data(self):
        # First validate all date fields before proceeding
        for header in self.date_fields:
            if header in self.entries:
                entry = self.entries[header]
                try:
                    if isinstance(entry, DateEntry):
                        date_value = entry.entry.get()
                        if date_value:
                            datetime.strptime(date_value, '%m-%d-%Y')
                except ValueError:
                    messagebox.showerror(
                        "Invalid Date Format",
                        f"Please enter a valid date in MM-DD-YYYY format for {header}"
                    )
                    return False

        # Collect all data from entry fields
        data = {}
        for header, entry in self.entries.items():
            if isinstance(entry, DateEntry):
                value = entry.entry.get()
                data[header] = value if value else ''
            elif isinstance(entry, ttk.Combobox):
                data[header] = entry.get()
            else:
                data[header] = entry.get()

        try:
            # Initialize row data
            row = []

            # Process first 21 columns (A through U)
            for header in self.app.headers[:21]:
                value = data.get(header, '')

                if header in self.date_fields and value:
                    try:
                        # Parse and reformat date to ensure MM-DD-YYYY format
                        date_obj = datetime.strptime(value, '%m-%d-%Y')
                        # Store as formula to force date format
                        value = f'=DATE({date_obj.year},{date_obj.month},{date_obj.day})'
                    except ValueError:
                        messagebox.showerror(
                            "Date Format Error",
                            f"Invalid date format in {header}. Please use MM-DD-YYYY format."
                        )
                        return False
                elif value and value.replace('.', '').isdigit():
                    # Handle numeric values
                    try:
                        if '.' in value:
                            value = float(value)
                        else:
                            value = int(value)
                    except ValueError:
                        pass

                row.append(value)

            # Add Tenant type (Column V)
            tenant_type_value = data.get("Tenant type", "")
            row.append(tenant_type_value)

            # Add Brokerage (Column W)
            brokerage_value = data.get("Brokerage", "")
            if isinstance(brokerage_value, ttk.Combobox):
                brokerage_value = brokerage_value.get()
            row.append(brokerage_value)

            # Create folder name components
            folder_components = [
                str(row[10]),  # Column K
                str(row[11]),  # Column L
                str(row[12]),  # Column M
                str(row[4])  # Column E
            ]

            # Create folder name
            asset_folder_name = f"{folder_components[0]}-{folder_components[1]}{folder_components[2]} {folder_components[3]}"
            asset_folder_name = asset_folder_name.strip()

            # Process document uploads (Columns X through AB)
            doc_headers = [
                "KYC",  # Column X
                "Tenent Verification",  # Column Y
                "Property tax",  # Column Z
                "Lease upload",  # Column AA
                "Cheque PDC"  # Column AB
            ]

            has_uploads = False
            asset_folder_id = None

            # Handle document uploads
            for doc_header in doc_headers:
                file_path = self.document_entries.get(doc_header)

                if file_path == "NA":
                    row.append("NA")
                elif file_path:
                    try:
                        if not has_uploads:
                            asset_folder_id = self.app.create_subfolder(asset_folder_name)
                            has_uploads = True

                        subfolder_id = self.app.create_subfolder(doc_header, asset_folder_id)

                        if not os.path.exists(file_path):
                            raise FileNotFoundError(f"File not found: {file_path}")

                        upload_success = self.app.upload_file_to_folder(file_path, subfolder_id)

                        if upload_success:
                            row.append("UPLOADED")
                        else:
                            raise Exception("File upload failed")

                    except Exception as e:
                        print(f"Error processing document {doc_header}: {str(e)}")
                        messagebox.showerror(
                            "Document Processing Error",
                            f"Failed to process {doc_header}: {str(e)}"
                        )
                        return False
                else:
                    row.append("")

            # Add Lease Manager (Column AC)
            lease_manager_value = data.get("Lease Manager", "")
            row.append(lease_manager_value)

            # Add folder link (Column AD)
            folder_link = f"https://drive.google.com/drive/folders/{asset_folder_id}" if has_uploads else ""
            row.append(folder_link)

            # Debug print to verify data
            print("Row data before submission:")
            for idx, value in enumerate(row):
                print(f"Column {idx}: {value}")

            try:
                # Use insert_row instead of append_row for better control
                next_row = len(self.app.sheet.get_all_values()) + 1
                self.app.sheet.insert_row(row, next_row, value_input_option='USER_ENTERED')

                messagebox.showinfo(
                    "Success",
                    "Asset information and documents have been successfully submitted."
                )
                self.clear_input_fields()

                if self.doc_frame_created:
                    self.doc_frame.destroy()
                    self.doc_frame_created = False
                    self.submit_button.pack_forget()

                return True

            except Exception as e:
                messagebox.showerror(
                    "Sheet Update Error",
                    f"Failed to update spreadsheet: {str(e)}"
                )
                return False

        except Exception as e:
            messagebox.showerror(
                "Submission Error",
                f"An unexpected error occurred while submitting the data:\n{str(e)}"
            )
            return False
    def clear_input_fields(self):
        # Clear all entry fields
        for header, entry in self.entries.items():
            if header in self.date_fields:
                # Clear DateEntry widgets by deleting the entry contents
                entry.entry.delete(0, tk.END)
            else:
                # Clear regular Entry widgets
                entry.delete(0, tk.END)

        # Reset document entries dictionary
        self.document_entries = {}

        # Remove all document upload related widgets
        if self.doc_frame_created:
            # Destroy the document upload section title if it exists
            for widget in self.scrollable_frame.winfo_children():
                if isinstance(widget, ttk.Label) and widget.cget("text") == "Document Upload":
                    widget.destroy()
            # Destroy the document frame
            self.doc_frame.destroy()
            # Reset the doc_frame_created flag
            self.doc_frame_created = False

        # Hide the submit button
        if hasattr(self, 'submit_button') and self.submit_button.winfo_ismapped():
            self.submit_button.pack_forget()

    # In MainUI class, replace the existing setup_scrolling method with:
    def setup_scrolling(self):
        def _on_mousewheel(event):
            if not self.canvas.winfo_exists():
                return

            current_pos = self.canvas.yview()

            if self.os_type == "Windows":
                delta = -int(event.delta / 120)
            elif self.os_type == "Darwin":  # macOS
                delta = -event.delta
            else:  # Linux
                if event.num == 4:
                    delta = -1
                elif event.num == 5:
                    delta = 1
                else:
                    return

            scroll_amount = delta * 0.05

            if ((delta > 0 and current_pos[1] < 1) or
                    (delta < 0 and current_pos[0] > 0)):
                self.canvas.yview_moveto(current_pos[0] + scroll_amount)

        def _on_shift_mousewheel(event):
            if not self.canvas.winfo_exists():
                return

            current_pos = self.canvas.xview()

            if self.os_type == "Windows":
                delta = -int(event.delta / 120)
            elif self.os_type == "Darwin":  # macOS
                delta = -event.delta
            else:  # Linux
                if event.num == 4:
                    delta = -1
                elif event.num == 5:
                    delta = 1
                else:
                    return

            scroll_amount = delta * 0.05

            if ((delta > 0 and current_pos[1] < 1) or
                    (delta < 0 and current_pos[0] > 0)):
                self.canvas.xview_moveto(current_pos[0] + scroll_amount)

        # Store the handlers as instance variables
        self._on_mousewheel = _on_mousewheel
        self._on_shift_mousewheel = _on_shift_mousewheel

        def restore_scrolling(event=None):
            # Bind all scroll events
            self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
            self.canvas.bind_all("<Shift-MouseWheel>", self._on_shift_mousewheel)
            self.canvas.bind_all("<Button-4>", self._on_mousewheel)
            self.canvas.bind_all("<Button-5>", self._on_mousewheel)
            self.canvas.bind_all("<Shift-Button-4>", self._on_shift_mousewheel)
            self.canvas.bind_all("<Shift-Button-5>", self._on_shift_mousewheel)

        # Store the restore function as instance variable
        self._restore_scrolling = restore_scrolling

        # Initial binding
        restore_scrolling()

        def _clean_bindings(event=None):
            if self.canvas.winfo_exists():
                self.canvas.unbind_all("<MouseWheel>")
                self.canvas.unbind_all("<Shift-MouseWheel>")
                self.canvas.unbind_all("<Button-4>")
                self.canvas.unbind_all("<Button-5>")
                self.canvas.unbind_all("<Shift-Button-4>")
                self.canvas.unbind_all("<Shift-Button-5>")

        # Add these bindings to handle focus changes
        self.parent.bind("<FocusIn>", restore_scrolling)
        self.parent.bind("<FocusOut>", _clean_bindings)
        self.parent.bind("<Destroy>", _clean_bindings)