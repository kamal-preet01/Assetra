import tkinter as tk
from tkinter import ttk, messagebox
import datetime
from dateutil.parser import parse
from ttkbootstrap import Style
import webbrowser



class AssetManagementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Enterprise Asset Management System")

        # Configure window
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        window_width = int(screen_width * 0.85)
        window_height = int(screen_height * 0.85)
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # Initialize data

        # Create and style notebook
        self.style = Style(theme='flatly')
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=2, pady=2)

        # Create tabs
        self.main_tab = ttk.Frame(self.notebook)
        self.asset_list_tab = ttk.Frame(self.notebook)
        self.reminders_tab = ttk.Frame(self.notebook)

        self.notebook.add(self.main_tab, text='Dashboard')
        self.notebook.add(self.asset_list_tab, text='Asset Management')
        self.notebook.add(self.reminders_tab, text='Maintenance Schedule')

        # Initialize reminders
        self.reminders = Reminders(self.reminders_tab, self)


class Reminders:
    def __init__(self, parent, app):
        self.status_colors = None
        self.colors = None
        self.parent = parent
        self.app = app
        self.style = Style(theme='flatly')
        self.selected_days = tk.StringVar(value="30")
        self.setup_styles()
        self.create_reminders_ui()

    def setup_styles(self):
        # Corporate color palette
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


        self.status_colors = {
            'expired': '#DC143C',  # Crimson red for expired
            'urgent': '#FF2400',  # Very light red
            'warning': '#fff3e0',  # Very light orange
            'normal': '#e8f5e9'  # Very light green
        }


        # Title styling
        self.style.configure('Title.TLabel',
                             font=('Segoe UI', 24, 'normal'),
                             foreground=self.colors['primary'],
                             background=self.colors['surface'])

        # Frame styling
        self.style.configure('Card.TFrame',
                             background=self.colors['surface'],
                             relief='flat')

        self.style.configure('Header.TFrame',
                             background=self.colors['surface'])

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

        # TreeView styling - more compact and professional
        self.style.configure('Custom.Treeview',
                             background=self.colors['surface'],
                             fieldbackground=self.colors['surface'],
                             font=('Segoe UI', 10),
                             rowheight=35)

        self.style.configure('Custom.Treeview.Heading',
                             font=('Segoe UI', 10, 'bold'),
                             background=self.colors['background'],
                             foreground=self.colors['primary'],
                             relief='flat')

        self.style.map('Custom.Treeview.Heading',
                       background=[('pressed', self.colors['background']),
                                   ('active', self.colors['background'])])

        # Filter styling
        self.style.configure('Filter.TFrame',
                             background=self.colors['surface'])

        self.style.configure('Filter.TLabel',
                             font=('Segoe UI', 10),
                             background=self.colors['surface'],
                             foreground=self.colors['text_secondary'])

        # Button styling
        self.style.configure('Action.TButton',
                             font=('Segoe UI', 10),
                             background=self.colors['primary'],
                             padding=(15, 8))

    def create_reminders_ui(self):
        # Main container
        container = ttk.Frame(self.parent, style='Card.TFrame')
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Header with filter
        header_frame = ttk.Frame(container, style='Header.TFrame')
        header_frame.pack(fill=tk.X, padx=20, pady=(20, 10))

        # Title and subtitle
        title_frame = ttk.Frame(header_frame, style='Header.TFrame')
        title_frame.pack(side=tk.LEFT)

        title = ttk.Label(title_frame,
                          text="Maintenance Schedule",
                          style='Title.TLabel')
        title.pack(anchor=tk.W)

        subtitle = ttk.Label(title_frame,
                             text="Track and manage upcoming maintenance activities",
                             font=('Segoe UI', 10),
                             foreground=self.colors['text_secondary'],
                             background=self.colors['surface'])
        subtitle.pack(anchor=tk.W, pady=(5, 0))

        # Filter section
        filter_frame = ttk.Frame(header_frame, style='Filter.TFrame')
        filter_frame.pack(side=tk.RIGHT, pady=(10, 0))

        filter_label = ttk.Label(filter_frame,
                                 text="Time Range:",
                                 style='Filter.TLabel')
        filter_label.pack(side=tk.LEFT, padx=(0, 10))

        days_options = ["30", "60", "90", "120", "180", "All", "Expired"]
        days_dropdown = ttk.Combobox(filter_frame,
                                     textvariable=self.selected_days,
                                     values=days_options,
                                     width=8,
                                     font=('Segoe UI', 10),
                                     state="readonly")
        days_dropdown.pack(side=tk.LEFT, padx=(0, 5))

        def update_status_indicators():
            for widget in legend_frame.winfo_children():
                widget.destroy()

            if self.selected_days.get() == "Expired":
                status_indicators = [("Expired", 'expired')]
            else:
                status_indicators = [
                    ("Critical", 'urgent'),
                    ("Upcoming", 'warning'),
                    ("Scheduled", 'normal')
                ]

            for text, status in status_indicators:
                indicator_frame = ttk.Frame(legend_frame, style='Filter.TFrame')
                indicator_frame.pack(side=tk.LEFT, padx=(0, 15))

                status_dot = ttk.Label(indicator_frame,
                                       text="●",
                                       font=('Segoe UI', 14),
                                       foreground=self.status_colors[status],
                                       background=self.colors['surface'])
                status_dot.pack(side=tk.LEFT, padx=(0, 5))

                label = ttk.Label(indicator_frame,
                                  text=text,
                                  font=('Segoe UI', 9),
                                  foreground=self.colors['text_secondary'],
                                  background=self.colors['surface'])
                label.pack(side=tk.LEFT)

        # Bind the dropdown selection to update both the reminders and status indicators
        days_dropdown.bind('<<ComboboxSelected>>', lambda e: (self.refresh_reminders(), update_status_indicators()))

        ttk.Label(filter_frame,
                  text="days",
                  style='Filter.TLabel').pack(side=tk.LEFT)

        days_dropdown.bind('<<ComboboxSelected>>', lambda e: self.refresh_reminders())

        # Status indicators
        legend_frame = ttk.Frame(container, style='Filter.TFrame')
        legend_frame.pack(fill=tk.X, padx=20, pady=10)

        status_indicators = [
            ("Critical", 'urgent'),
            ("Upcoming", 'warning'),
            ("Scheduled", 'normal')
        ]

        for text, status in status_indicators:
            indicator_frame = ttk.Frame(legend_frame, style='Filter.TFrame')
            indicator_frame.pack(side=tk.LEFT, padx=(0, 15))

            status_dot = ttk.Label(indicator_frame,
                                   text="●",
                                   font=('Segoe UI', 14),
                                   foreground=self.status_colors[status],
                                   background=self.colors['surface'])
            status_dot.pack(side=tk.LEFT, padx=(0, 5))

            label = ttk.Label(indicator_frame,
                              text=text,
                              font=('Segoe UI', 9),
                              foreground=self.colors['text_secondary'],
                              background=self.colors['surface'])
            label.pack(side=tk.LEFT)

        # TreeView
        tree_frame = ttk.Frame(container)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(10, 20))

        columns = ['Asset ID','Lease Manager','Micro Market', 'Location', 'Project', 'Tower', 'Floor', 'Unit Number',
                   'Lease Expiry', 'Days Remaining' ]
        self.tree = ttk.Treeview(tree_frame,
                                 columns=columns,
                                 show="headings",
                                 style='Custom.Treeview',
                                 height=15)
        # Add these bindings right after creating the Treeview
        self.tree.bind('<Double-1>', self.on_treeview_double_click)
        self.tree.bind('<Motion>', self.on_hover)


        col_widths = {
            'Asset ID': 60,
            'Lease Manager': 60,
            'Micro Market': 110,
            'Location': 100,
            'Project': 100,
            'Tower': 70,
            'Floor': 80,
            'Unit Number': 80,
            'Lease Expiry': 80,
            'Days Remaining': 70

        }

        for col in columns:
            self.tree.heading(col, text=col, anchor=tk.W)
            self.tree.column(col, width=col_widths[col], minwidth=col_widths[col])

        # Scrollbar
        # Scrollbar with proper binding
        scrollbar = ttk.Scrollbar(tree_frame,
                                orient="vertical",
                                command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.bind('<MouseWheel>', self._on_mousewheel)       # Windows
        self.tree.bind('<Button-4>', self._on_mousewheel)         # Linux up
        self.tree.bind('<Button-5>', self._on_mousewheel)         # Linux down

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Footer
        footer_frame = ttk.Frame(container, style='Filter.TFrame')
        footer_frame.pack(fill=tk.X, padx=20, pady=(0, 20))

        refresh_button = ttk.Button(filter_frame,
                                    text="Refresh",
                                    style='Action.TButton',
                                    command=self.refresh_reminders)
        refresh_button.pack(side=tk.LEFT, padx=(10, 0))

        # Initialize data
        self.refresh_reminders()

    def refresh_reminders(self):
        self.tree.delete(*self.tree.get_children())
        all_rows = self.app.sheet.get_all_values()
        today = datetime.date.today()

        days_filter = self.selected_days.get()
        show_expired = days_filter.lower() == 'expired'

        # Set max_days for non-expired view
        if days_filter.lower() == 'all':
            max_days = float('inf')
        elif days_filter.lower() == 'expired':
            max_days = 0  # Not used for expired view
        else:
            max_days = int(days_filter)

        for row in all_rows[1:]:
            try:
                date_str = row[14]
                if date_str:
                    date = parse(date_str).date()
                    days_left = (date - today).days

                    # Show row based on filter selection
                    show_row = False
                    if show_expired:
                        # Only show expired leases when "Expired" is selected
                        show_row = days_left <= 0
                    else:
                        # Only show non-expired leases for all other filters
                        show_row = days_left > 0 and (days_filter.lower() == 'all' or days_left <= max_days)

                    if show_row:
                        values = [
                            f"AST-{row[0]}",  # Formatted Asset ID
                            row[28],  # Lease Manager
                            row[2],  # Property Name
                            row[3],  # Location
                            row[4],  # project
                            row[10],  # tower
                            row[11],  # Floor
                            row[12],  # unit no.
                            row[14],  # lease expiry date
                            days_left,  # Days remaining as integer
                        ]

                        item = self.tree.insert("", "end", values=values)

                        # Apply status styles
                        if show_expired:
                            self.tree.tag_configure('expired', background=self.status_colors['expired'])
                            self.tree.item(item, tags=('expired',))
                        else:
                            if days_left <= 15:
                                self.tree.tag_configure('urgent', background=self.status_colors['urgent'])
                                self.tree.item(item, tags=('urgent',))
                            elif days_left <= 31:
                                self.tree.tag_configure('warning', background=self.status_colors['warning'])
                                self.tree.item(item, tags=('warning',))
                            else:
                                self.tree.tag_configure('normal', background=self.status_colors['normal'])
                                self.tree.item(item, tags=('normal',))

            except (ValueError, IndexError):
                continue

        # Sort by days remaining
        sorted_items = sorted(self.tree.get_children(''),
                              key=lambda x: int(self.tree.item(x)['values'][-1]))

        # After sorting, update the display to show formatted values
        for index, item in enumerate(sorted_items):
            values = list(self.tree.item(item)['values'])
            values[-1] = f"{values[-1]:,}"  # Format the days remaining with commas
            self.tree.item(item, values=values)
            self.tree.move(item, '', index)

    def show_asset_details(self, values):
        assets = self.app.sheet.get_all_values()
        headers = assets[0]
        # Find the full row data using Asset ID and Property Name
        full_row = next((row for row in assets[1:] if row[0] == values[0].replace("AST-", "") and row[2] == values[2]),
                        None)

        if full_row:
            details_window = tk.Toplevel(self.parent)
            details_window.title("Asset Details")
            details_window.geometry("700x600")
            details_window.minsize(700, 600)

            def on_window_close():
                canvas.unbind_all("<MouseWheel>")
                canvas.unbind_all("<Button-4>")
                canvas.unbind_all("<Button-5>")
                details_window.destroy()

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

            # Create scrollable content area
            content_frame = ttk.Frame(main_container, style='Card.TFrame')
            content_frame.pack(fill=tk.BOTH, expand=True)

            canvas = tk.Canvas(content_frame,
                               bg=self.colors['surface'],
                               highlightthickness=0)
            scrollbar = ttk.Scrollbar(content_frame,
                                      orient="vertical",
                                      command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas, style='Card.TFrame')

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )

            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=canvas.winfo_width())
            canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas.find_all()[0], width=e.width))

            # Information groups
            groups = {
                'Basic Info': [0, 1, 2, 3, 4, 5],
                'Property Details': [6, 7, 8, 9, 10],
                'Lease Information': [11, 12, 13, 14, 15],
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
            button_frame = ttk.Frame(main_container, style='Card.TFrame')
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

    def on_treeview_double_click(self, event):
        selected_item = self.tree.selection()
        if selected_item:
            item = selected_item[0]
            values = self.tree.item(item, 'values')
            self.show_asset_details(values)

    def on_hover(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            self.tree.configure(cursor="hand2")
        else:
            self.tree.configure(cursor="")

    def _setup_scroll_bindings(self, canvas):
        def _bind_to_mousewheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.bind_all("<Button-4>", _on_mousewheel)
            canvas.bind_all("<Button-5>", _on_mousewheel)

        def _unbind_from_mousewheel(event=None):
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
        canvas.winfo_toplevel().bind('<Destroy>', _unbind_from_mousewheel)

    def _center_window(self, window, width, height):
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        window.geometry(f"{width}x{height}+{x}+{y}")

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

        self.tree.yview_scroll(move, "units")
