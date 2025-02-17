import tkinter as tk
from tkinter import ttk, messagebox
from ttkbootstrap import Style
import webbrowser

class AssetList:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.style = Style(theme='flatly')
        self.setup_styles()
        self.create_asset_list_ui()

    def setup_styles(self):
        # Corporate color palette matching MainUI
        self.colors = {
            'primary': '#2c3e50',  # Dark blue-gray
            'secondary': '#34495e',  # Lighter blue-gray
            'accent': '#455a64',  # Medium blue-gray
            'background': '#f5f6fa',  # Light gray
            'surface': '#ffffff',  # White
            'border': '#dfe4ea',  # Light border
            'text': '#2c3e50',  # Dark text
            'text_secondary': '#576574',  # Secondary text
            'warning': '#ffebee',  # Light red for warning
            'selection': '#e3f2fd'  # Light blue for selection overlay
        }

        self.style.configure('Warning.Treeview',
                           background=self.colors['warning'])

        # Title styling
        self.style.configure('Title.TLabel',
                           font=('Segoe UI', 24, 'normal'),
                           foreground=self.colors['primary'],
                           background=self.colors['surface'])

        self.style.configure('Subtitle.TLabel',
                           font=('Segoe UI', 10),
                           foreground=self.colors['text_secondary'],
                           background=self.colors['surface'])

        # Frame styles
        self.style.configure('Card.TFrame',
                           background=self.colors['surface'],
                           relief='flat')

        self.style.configure('Content.TFrame',
                           background=self.colors['surface'])

        # Custom Treeview styling
        self.style.configure('Custom.Treeview',
                           background=self.colors['surface'],
                           fieldbackground=self.colors['surface'],
                           font=('Segoe UI', 10),
                           rowheight=40)

        self.style.configure('Custom.Treeview.Heading',
                           font=('Segoe UI', 10, 'bold'),
                           background=self.colors['background'],
                           foreground=self.colors['text'])

        # Treeview selection colors
        self.style.map('Custom.Treeview',
                      background=[
                          ('selected', '#cfd8dc')  # Medium gray-blue for selection
                      ],
                      foreground=[
                          ('selected', self.colors['text'])  # Keep text color consistent
                      ])

        # Detail window styles
        self.style.configure('DetailHeader.TFrame',
                           background=self.colors['primary'])

        self.style.configure('DetailHeader.TLabel',
                           font=('Segoe UI', 18, 'normal'),
                           foreground=self.colors['surface'],
                           background=self.colors['primary'])

        self.style.configure('DetailRow.TFrame',
                           background=self.colors['surface'])

        self.style.configure('DetailLabel.TLabel',
                           font=('Segoe UI', 10, 'bold'),
                           foreground=self.colors['text'])

        self.style.configure('DetailValue.TLabel',
                           font=('Segoe UI', 10),
                           foreground=self.colors['text_secondary'])

        # Search and button styles
        self.style.configure('Search.TEntry',
                           font=('Segoe UI', 10),
                           fieldbackground=self.colors['surface'])

    def open_google_sheet(self):
        sheet_url = "https://docs.google.com/spreadsheets/d/1orIbEddJvC9PExzZnfxct8xW-fz_w9PVk32u3QO5694/edit?gid=0#gid=0"
        webbrowser.open(sheet_url)

    def create_asset_list_ui(self):
        # Main container with professional padding
        main_container = ttk.Frame(self.parent, style='Card.TFrame')
        main_container.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)

        # Enhanced header section
        header_frame = ttk.Frame(main_container, style='Content.TFrame')
        header_frame.pack(fill=tk.X, pady=(20, 30), padx=20)

        title = ttk.Label(header_frame,
                         text="Welcome to AsseTRA",
                         style='Title.TLabel')
        title.pack(anchor=tk.W)

        subtitle = ttk.Label(header_frame,
                           text="View and manage your enterprise assets on assetra",
                           style='Subtitle.TLabel')
        subtitle.pack(anchor=tk.W, pady=(5, 0))

        # Enhanced search section
        search_frame = ttk.Frame(main_container, style='Content.TFrame')
        search_frame.pack(fill=tk.X, padx=20, pady=(0, 20))

        search_container = ttk.Frame(search_frame, style='Card.TFrame', padding=15)
        search_container.pack(fill=tk.X)

        # Search entry with improved styling
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_container,
                               textvariable=self.search_var,
                               style='Search.TEntry',
                               width=40)
        search_entry.pack(side=tk.LEFT, padx=(0, 10))
        search_entry.insert(0, "Search assets...")
        search_entry.bind('<FocusIn>', lambda e: search_entry.delete(0, tk.END)
                         if search_entry.get() == "Search assets..." else None)
        search_entry.bind('<FocusOut>', lambda e: search_entry.insert(0, "Search assets...")
                         if search_entry.get() == "" else None)

        # Action buttons
        search_button = ttk.Button(search_container,
                                 text="Search",
                                 style="primary.TButton",
                                 command=self.search_assets)
        search_button.pack(side=tk.LEFT, padx=5)

        refresh_button = ttk.Button(search_container,
                                  text="↻ Refresh List",
                                  style="primary.TButton",
                                  command=self.update_asset_list)
        refresh_button.pack(side=tk.LEFT, padx=5)

        open_sheet_button = ttk.Button(search_container,
                                      text="📄 Open Sheet",
                                      style="primary.TButton",
                                      command=self.open_google_sheet)
        open_sheet_button.pack(side=tk.LEFT, padx=5)

        # Assets list container
        list_frame = ttk.Frame(main_container, style='Card.TFrame', padding=20)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))

        # Configure display columns
        display_columns = [
            self.app.headers[0],  # Asset ID
            self.app.headers[2],  # Property Name
            self.app.headers[3],  # Location
            self.app.headers[4],  # Column E
            self.app.headers[5],  # Column F
            self.app.headers[6],  # Category
            self.app.headers[10],  # Column K
            self.app.headers[11],  # Status
            self.app.headers[12],  # Column M
            self.app.headers[28]  # Column AC
        ]

        self.asset_list = ttk.Treeview(list_frame,
                                     columns=display_columns,
                                     show='headings',
                                     style='Custom.Treeview',
                                     height=15)

        # Column configurations
        column_widths = {
            self.app.headers[0]: 60,  # Asset ID
            self.app.headers[2]: 180,  # Property Name
            self.app.headers[3]: 130,  # Location
            self.app.headers[4]: 160,  # Column E
            self.app.headers[5]: 120,  # Column F
            self.app.headers[6]: 150,  # Category
            self.app.headers[10]: 120,  # Column K
            self.app.headers[11]: 100,  # Status
            self.app.headers[12]: 120,  # Column M
            self.app.headers[28]: 100  # Column AC
        }

        for header in display_columns:
            self.asset_list.heading(header, text=header, anchor=tk.W)
            self.asset_list.column(header, width=column_widths[header], minwidth=50, anchor=tk.W)

        scrollbar = ttk.Scrollbar(list_frame,
                                orient="vertical",
                                command=self.asset_list.yview)
        self.asset_list.configure(yscrollcommand=scrollbar.set)


        # Bind mousewheel events directly to the Treeview
        self.asset_list.bind('<MouseWheel>', self._on_mousewheel)  # Windows
        self.asset_list.bind('<Button-4>', self._on_mousewheel)    # Linux up
        self.asset_list.bind('<Button-5>', self._on_mousewheel)    # Linux down

        self.asset_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Event bindings
        self.asset_list.bind("<Double-1>", self.on_treeview_double_click)
        self.asset_list.bind('<Motion>', self.on_hover)

        # Initialize the list
        self.update_asset_list()

    def show_asset_details(self, values):
        assets = self.app.sheet.get_all_values()
        headers = assets[0]
        full_row = next((row for row in assets[1:] if row[0] == values[0] and row[2] == values[1]), None)

        if full_row:
            details_window = tk.Toplevel(self.parent)
            details_window.title("Asset Details")
            details_window.geometry("700x600")
            details_window.minsize(700, 600)

            def on_window_close():
                # Clean up canvas bindings
                canvas.unbind_all("<MouseWheel>")
                canvas.unbind_all("<Button-4>")
                canvas.unbind_all("<Button-5>")
                details_window.destroy()
                # Restore main window scrolling
                if hasattr(self.app.main_ui, '_restore_scrolling'):
                    self.app.main_ui._restore_scrolling()

            details_window.protocol("WM_DELETE_WINDOW", on_window_close)

            # Main container
            main_container = ttk.Frame(details_window, style='Card.TFrame')
            main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

            # Header with property name and asset ID
            header_frame = ttk.Frame(main_container, style='DetailHeader.TFrame')
            header_frame.pack(fill=tk.X, pady=(0, 15))

            header_content = ttk.Frame(header_frame, style='DetailHeader.TFrame', padding=(15, 10))
            header_content.pack(fill=tk.X)

            title = ttk.Label(header_content,
                              text=f"{full_row[4]}",
                              style='DetailHeader.TLabel')
            title.pack(side=tk.LEFT)

            asset_id = ttk.Label(header_content,
                                 text=f"ID: {full_row[0]}",
                                 style='DetailHeader.TLabel')
            asset_id.pack(side=tk.RIGHT)

            # Create content area with canvas for scrolling
            content_frame = ttk.Frame(main_container, style='Card.TFrame')
            content_frame.pack(fill=tk.BOTH, expand=True)

            canvas = tk.Canvas(content_frame,
                               bg=self.colors['surface'],
                               highlightthickness=0)
            scrollbar = ttk.Scrollbar(content_frame,
                                      orient="vertical",
                                      command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas, style='Content.TFrame')

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )

            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=canvas.winfo_width())
            canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas.find_all()[0], width=e.width))

            # Original groups
            groups = {
                'Basic Info': [0, 1, 2, 3, 4, 5],
                'Property Details': [6, 7, 8, 9, 10],
                'Financial Info': [11, 12, 13, 14, 15],
                'Additional Details': [16, 17, 18, 19, 20]
            }

            # Add groups
            for group_name, field_indices in groups.items():
                group_frame = ttk.Frame(scrollable_frame, style='DetailRow.TFrame')
                group_frame.pack(fill=tk.X, pady=5, padx=10)

                ttk.Label(group_frame,
                          text=group_name,
                          style='DetailLabel.TLabel',
                          font=('Segoe UI', 11, 'bold')).pack(anchor=tk.W, pady=(10, 5))

                content_frame = ttk.Frame(group_frame, style='DetailRow.TFrame')
                content_frame.pack(fill=tk.X, padx=15)

                for idx in field_indices:
                    if idx < len(headers) and idx < len(full_row):
                        row = ttk.Frame(content_frame, style='DetailRow.TFrame')
                        row.pack(fill=tk.X, pady=2)

                        header_label = ttk.Label(row,
                                                 text=f"{headers[idx]}:",
                                                 style='DetailLabel.TLabel',
                                                 width=20,
                                                 anchor='e')
                        header_label.pack(side=tk.LEFT, padx=(5, 10))

                        value_label = ttk.Label(row,
                                                text=str(full_row[idx] or '-'),
                                                style='DetailValue.TLabel',
                                                anchor='w')
                        value_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

            # Add Brokerage Status Section
            brokerage_frame = ttk.Frame(scrollable_frame, style='DetailRow.TFrame')
            brokerage_frame.pack(fill=tk.X, pady=5, padx=10)

            ttk.Label(brokerage_frame,
                      text="Brokerage Status",
                      style='DetailLabel.TLabel',
                      font=('Segoe UI', 11, 'bold')).pack(anchor=tk.W, pady=(10, 5))

            brokerage_content = ttk.Frame(brokerage_frame, style='DetailRow.TFrame')
            brokerage_content.pack(fill=tk.X, padx=15)

            status_row = ttk.Frame(brokerage_content, style='DetailRow.TFrame')
            status_row.pack(fill=tk.X, pady=2)

            ttk.Label(status_row,
                      text=f"{headers[22]}:",  # Column W header
                      style='DetailLabel.TLabel',
                      width=20,
                      anchor='e').pack(side=tk.LEFT, padx=(5, 10))

            # Dropdown for brokerage status
            status_var = tk.StringVar(value=full_row[22] or 'Pending')
            status_dropdown = ttk.Combobox(status_row,
                                           textvariable=status_var,
                                           values=['Pending', 'Received'],
                                           state='readonly',
                                           width=20)
            status_dropdown.pack(side=tk.LEFT, padx=(0, 10))

            def update_brokerage_status():
                try:
                    # Find the row index in the sheet
                    row_index = next(i for i, row in enumerate(assets[1:], start=2)
                                     if row[0] == values[0] and row[2] == values[1])

                    # Update the cell in column W (index 22)
                    self.app.sheet.update_cell(row_index, 23, status_var.get())  # 23 because sheets are 1-indexed

                    messagebox.showinfo("Success", "Brokerage status updated successfully!")

                    # Update the main list view
                    self.update_asset_list()
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to update brokerage status: {str(e)}")

            # Add update button
            update_button = ttk.Button(status_row,
                                       text="Update Status",
                                       style="primary.TButton",
                                       command=update_brokerage_status)
            update_button.pack(side=tk.LEFT)

            # Document Details Section
            doc_frame = ttk.Frame(scrollable_frame, style='DetailRow.TFrame')
            doc_frame.pack(fill=tk.X, pady=5, padx=10)

            ttk.Label(doc_frame,
                      text="Document Details",
                      style='DetailLabel.TLabel',
                      font=('Segoe UI', 11, 'bold')).pack(anchor=tk.W, pady=(10, 5))

            doc_content_frame = ttk.Frame(doc_frame, style='DetailRow.TFrame')
            doc_content_frame.pack(fill=tk.X, padx=15)

            # Document columns (X, Y, Z, AA, AB) - indices 23 to 27
            for i in range(23, 28):
                row = ttk.Frame(doc_content_frame, style='DetailRow.TFrame')
                row.pack(fill=tk.X, pady=2)

                header_label = ttk.Label(row,
                                         text=f"{headers[i]}:",
                                         style='DetailLabel.TLabel',
                                         width=20,
                                         anchor='e')
                header_label.pack(side=tk.LEFT, padx=(5, 10))

                value = str(full_row[i] or '-')
                value_label = ttk.Label(row,
                                        text=value,
                                        style='DetailValue.TLabel',
                                        anchor='w')

                if value.upper().strip() not in ["UPLOADED", "NA"] and value.strip():
                    row.configure(style='Warning.TFrame')
                    value_label.configure(background=self.colors['warning'])

                value_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

            # Bottom button frame
            button_frame = ttk.Frame(main_container, style='Content.TFrame')
            button_frame.pack(fill=tk.X, pady=(15, 0))

            def open_drive_documents():
                drive_link = full_row[29]  # Column AD contains the Google Drive folder link
                if drive_link:
                    response = messagebox.askyesno(
                        "Open Documents",
                        "You will be redirected to Google Drive to view the documents. Continue?",
                        icon='question'
                    )
                    if response:
                        webbrowser.open(drive_link)
                else:
                    messagebox.showwarning("No Documents", "No document folder link available for this asset.")

            view_docs_button = ttk.Button(button_frame,
                                          text="View Documents",
                                          style="primary.TButton",
                                          command=open_drive_documents,
                                          width=15)
            view_docs_button.pack(side=tk.LEFT, padx=(0, 10))

            close_button = ttk.Button(button_frame,
                                      text="Close",
                                      style="primary.TButton",
                                      command=on_window_close,
                                      width=15)
            close_button.pack(side=tk.RIGHT)

            # Set up scrolling
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=(0, 15))
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            canvas.configure(yscrollcommand=scrollbar.set)

            self._setup_scroll_bindings(canvas)
            self._center_window(details_window, 700, 600)
        else:
            messagebox.showerror("Error", "Asset details not found.")

    def show_documents_window(self, drive_link, property_name, on_window_close=None):
        # Create new window for documents
        docs_window = tk.Toplevel(self.parent)
        docs_window.title(f"Documents - {property_name}")
        docs_window.geometry("800x600")
        docs_window.minsize(800, 600)

        # Main container
        main_container = ttk.Frame(docs_window, style='Card.TFrame')
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Header
        header_frame = ttk.Frame(main_container, style='DetailHeader.TFrame')
        header_frame.pack(fill=tk.X, pady=(0, 15))

        header_content = ttk.Frame(header_frame, style='DetailHeader.TFrame', padding=(15, 10))
        header_content.pack(fill=tk.X)

        title = ttk.Label(header_content,
                          text=f"Documents for {property_name}",
                          style='DetailHeader.TLabel')
        title.pack(side=tk.LEFT)

        # Create listbox for documents with scrollbar
        docs_frame = ttk.Frame(main_container, style='Card.TFrame')
        docs_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        # You'll need to implement the logic to fetch documents from Google Drive here
        # For now, we'll show a message about the drive link
        info_label = ttk.Label(docs_frame,
                               text=f"Google Drive Folder Link:\n{drive_link}",
                               style='DetailValue.TLabel',
                               wraplength=700)
        info_label.pack(pady=20)

        # Add information about implementation
        note_label = ttk.Label(docs_frame,
                               text="Note: To implement document fetching, you'll need to:\n" +
                                    "1. Use Google Drive API to fetch documents\n" +
                                    "2. List PDF files in the drive folder\n" +
                                    "3. Provide preview/download functionality",
                               style='DetailValue.TLabel',
                               wraplength=700)
        note_label.pack(pady=20)

        # Bottom button frame
        button_frame = ttk.Frame(main_container, style='Content.TFrame')
        button_frame.pack(fill=tk.X, pady=(15, 0))

        # Add Open in Browser button
        open_browser_button = ttk.Button(button_frame,
                                         text="Open in Browser",
                                         style="primary.TButton",
                                         command=lambda: webbrowser.open(drive_link),
                                         width=15)
        open_browser_button.pack(side=tk.LEFT, padx=(0, 10))

        close_button = ttk.Button(button_frame,
                                  text="Close",
                                  style="primary.TButton",
                                  command=on_window_close,
                                  width=15)
        close_button.pack(side=tk.RIGHT)

        self._center_window(docs_window, 800, 600)

    # Rest of the methods remain unchanged
    def update_asset_list(self):
        self.asset_list.delete(*self.asset_list.get_children())
        assets = self.app.sheet.get_all_values()[1:]  # Skip header row

        # Configure warning background tag
        self.asset_list.tag_configure('warning', background=self.colors['warning'])
        self.asset_list.tag_configure('evenrow', background=self.colors['background'])
        self.asset_list.tag_configure('selected_warning', background='#ffcdd2')  # Darker warning color for selection


        for i, row in enumerate(assets):
            values = [
                row[0],  # Asset ID
                row[2],  # Property Name
                row[3],  # Location
                row[4],  # Column E
                row[5],  # Column F
                row[6],  # Category
                row[10],  # Column K
                row[11],  # Status
                row[12],  # Column M
                row[28]  # Column AC
            ]

            # Check document status in columns X, Y, Z, AA, AB (indices 23-27)
            needs_warning = False
            for idx in range(23, 28):
                status = str(row[idx]).upper().strip()
                # If status is empty OR not "UPLOADED" or "NA", mark for warning
                if not status or (status not in ["UPLOADED", "NA"]):
                    needs_warning = True
                    break

            # Insert the row
            item = self.asset_list.insert('', tk.END, values=values)

            # Apply warning tag if needed, otherwise apply evenrow tag for alternating rows
            if needs_warning:
                self.asset_list.item(item, tags=('warning',))
            elif i % 2 == 0:
                self.asset_list.item(item, tags=('evenrow',))

    def search_assets(self):
        search_term = self.search_var.get().lower()
        if search_term == "search assets...":
            return

        self.asset_list.delete(*self.asset_list.get_children())
        assets = self.app.sheet.get_all_values()[1:]

        for i, row in enumerate(assets):
            if any(search_term in str(value).lower() for value in row):
                values = [
                    row[0],  # Asset ID
                    row[2],  # Property Name
                    row[3],  # Location
                    row[4],  # Column E
                    row[5],  # Column F
                    row[6],  # Category
                    row[10],  # Column K
                    row[11],  # Status
                    row[12],  # Column M
                    row[28]  # Column AC
                ]

                # Check document status in columns X, Y, Z, AA, AB (indices 23-27)
                needs_warning = False
                for idx in range(23, 28):
                    status = str(row[idx]).upper().strip()
                    # If status is empty OR not "UPLOADED" or "NA", mark for warning
                    if not status or (status not in ["UPLOADED", "NA"]):
                        needs_warning = True
                        break

                # Insert the row
                item = self.asset_list.insert('', tk.END, values=values)

                # Apply warning tag if needed, otherwise apply evenrow tag
                if needs_warning:
                    self.asset_list.item(item, tags=('warning',))
                elif i % 2 == 0:
                    self.asset_list.item(item, tags=('evenrow',))

    def on_hover(self, event):
        region = self.asset_list.identify_region(event.x, event.y)
        if region == "cell":
            self.asset_list.configure(cursor="hand2")
        else:
            self.asset_list.configure(cursor="")

    def on_treeview_double_click(self, event):
        selected_item = self.asset_list.selection()
        if selected_item:
            item = selected_item[0]
            values = self.asset_list.item(item, 'values')
            self.show_asset_details(values)

    def _on_mousewheel(self, event):
        """Handle mousewheel scrolling for the Treeview"""
        # For Windows
        if event.delta:
            move = -1 * (event.delta // 120)
        # For Linux
        elif event.num == 4:
            move = -1
        elif event.num == 5:
            move = 1
        else:
            return

        self.asset_list.yview_scroll(move, "units")

    def _setup_scroll_bindings(self, canvas):
        """Setup scrolling for detail window canvas"""

        def _bind_to_mousewheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.bind_all("<Button-4>", _on_mousewheel)
            canvas.bind_all("<Button-5>", _on_mousewheel)

        def _unbind_from_mousewheel(event=None):  # Made event optional
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        def _on_mousewheel(event):
            if event.num == 4 or event.delta > 0:
                canvas.yview_scroll(-1, "units")
            elif event.num == 5 or event.delta < 0:
                canvas.yview_scroll(1, "units")

        canvas.bind('<Enter>', _bind_to_mousewheel)
        canvas.bind('<Leave>', _unbind_from_mousewheel)

        # Add this: ensure bindings are cleaned up when window is closed
        canvas.winfo_toplevel().bind('<Destroy>', _unbind_from_mousewheel)

    def _center_window(self, window, width, height):
        # Get screen dimensions
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()

        # Calculate center position
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2

        # Set window geometry
        window.geometry(f"{width}x{height}+{x}+{y}")