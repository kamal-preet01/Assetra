import tkinter as tk
from tkinter import ttk
from ttkbootstrap import Style
import datetime
from dateutil.parser import parse
from tkinter import ttk, messagebox  # Added messagebox import
import webbrowser



class BrokerageManagement:
    def __init__(self, parent, app):
        self.status_colors = None
        self.colors = None
        self.parent = parent
        self.app = app
        self.style = Style(theme='flatly')
        self.selected_status = tk.StringVar(value="All")
        self.selected_period = tk.StringVar(value="All Time")
        self.stat_labels = {}
        self.setup_styles()
        self.create_brokerage_ui()

    def setup_styles(self):
        # Professional color palette
        self.colors = {
            'primary': '#1a365d',  # Deep navy blue
            'secondary': '#2d3748',  # Dark slate
            'accent': '#2563eb',  # Professional blue
            'background': '#f8fafc',  # Light gray background
            'surface': '#ffffff',  # Pure white
            'border': '#e2e8f0',  # Light gray
            'text': '#1e293b',  # Dark text
            'text_secondary': '#64748b',  # Medium gray text
            'success': '#16a34a',  # Professional green
            'warning': '#dc2626'  # Professional red
        }

        # Status colors - professional and subtle
        self.status_colors = {
            'received': '#6AB187',  # Very light green background
            'received_text': '#000000',  # Black text for received
            'pending': '#CED2CC',  # White background
            'pending_text': '#000000'  # Black text for pending
        }

        # Configure styles
        self.style.configure('Title.TLabel',
                             font=('Segoe UI Semibold', 24, 'normal'),
                             foreground=self.colors['primary'],
                             background=self.colors['surface'])

        self.style.configure('Card.TFrame',
                             background=self.colors['surface'],
                             relief='flat')

        self.style.configure('Header.TFrame',
                             background=self.colors['surface'])

        # Enhanced TreeView styling
        self.style.configure('Custom.Treeview',
                             background=self.colors['surface'],
                             fieldbackground=self.colors['surface'],
                             font=('Segoe UI', 10),
                             rowheight=40)

        self.style.configure('Custom.Treeview.Heading',
                             font=('Segoe UI Semibold', 10),
                             background=self.colors['background'],
                             foreground=self.colors['primary'])

        # Professional refresh button style
        self.style.configure('Refresh.TButton',
                             font=('Segoe UI Semibold', 10),
                             background="black",
                             foreground='white',
                             padding=(15, 8))

        # Hover effects
        self.style.map('Custom.Treeview',
                       background=[('selected', '#e2e8f0')],
                       foreground=[('selected', self.colors['primary'])])

        self.style.map('Refresh.TButton',
                       background=[('active', '#1d4ed8')],
                       foreground=[('active', 'white')])

    def create_stats_card(self, parent, title, key):
        card = ttk.Frame(parent, style='Card.TFrame', padding=(15, 10))
        card.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=8)

        # Add subtle border and shadow effect
        inner_frame = ttk.Frame(card, style='Card.TFrame')
        inner_frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        ttk.Label(inner_frame,
                 text=title,
                 font=('Segoe UI', 11),
                 foreground=self.colors['text_secondary'],
                 background=self.colors['surface']).pack(pady=(5, 0))

        value_label = ttk.Label(inner_frame,
                              text="0",
                              font=('Segoe UI Semibold', 20),
                              foreground=self.colors['accent'],
                              background=self.colors['surface'])
        value_label.pack(pady=(5, 5))

        self.stat_labels[key] = value_label

    def create_brokerage_ui(self):
        container = ttk.Frame(self.parent, style='Card.TFrame')
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Enhanced header section
        header_frame = ttk.Frame(container, style='Header.TFrame')
        header_frame.pack(fill=tk.X, padx=20, pady=(20, 15))

        # Left side: Title and subtitle
        title_frame = ttk.Frame(header_frame, style='Header.TFrame')
        title_frame.pack(side=tk.LEFT)

        title = ttk.Label(title_frame,
                          text="Brokerage Management",
                          style='Title.TLabel')
        title.pack(anchor=tk.W)

        subtitle = ttk.Label(title_frame,
                             text="Track and manage your property brokerage status",
                             font=('Segoe UI', 11),
                             foreground=self.colors['text_secondary'],
                             background=self.colors['surface'])
        subtitle.pack(anchor=tk.W, pady=(5, 0))

        # Right side: Filters and refresh button
        control_frame = ttk.Frame(header_frame, style='Filter.TFrame')
        control_frame.pack(side=tk.RIGHT, pady=(10, 0))

        # Status filter
        status_label = ttk.Label(control_frame,
                                 text="Status:",
                                 font=('Segoe UI', 11),
                                 background=self.colors['surface'])
        status_label.pack(side=tk.LEFT, padx=(0, 10))

        status_options = ["All", "Received", "Pending"]
        status_dropdown = ttk.Combobox(control_frame,
                                       textvariable=self.selected_status,
                                       values=status_options,
                                       width=12,
                                       font=('Segoe UI', 10),
                                       state="readonly")
        status_dropdown.pack(side=tk.LEFT, padx=(0, 20))

        # Refresh button
        refresh_button = ttk.Button(control_frame,
                                    text="⟳ Refresh",
                                    style='Refresh.TButton',
                                    command=self.refresh_brokerage_data)
        refresh_button.pack(side=tk.LEFT, padx=(0, 10))

        # Add status indicators legend
        legend_frame = ttk.Frame(container, style='Filter.TFrame')
        legend_frame.pack(fill=tk.X, padx=20, pady=10)

        status_indicators = [
            ("Received", 'received', self.status_colors['received']),
            ("Pending", 'pending', self.status_colors['pending'])
        ]

        for text, status, color in status_indicators:
            indicator_frame = ttk.Frame(legend_frame, style='Filter.TFrame')
            indicator_frame.pack(side=tk.LEFT, padx=(0, 15))

            status_dot = ttk.Label(indicator_frame,
                                   text="●",
                                   font=('Segoe UI', 14),
                                   foreground=color,
                                   background=self.colors['surface'])
            status_dot.pack(side=tk.LEFT, padx=(0, 5))

            label = ttk.Label(indicator_frame,
                              text=text,
                              font=('Segoe UI', 9),
                              foreground=self.colors['text_secondary'],
                              background=self.colors['surface'])
            label.pack(side=tk.LEFT)

        # Stats cards section
        stats_frame = ttk.Frame(container, style='Card.TFrame')
        stats_frame.pack(fill=tk.X, padx=20, pady=15)

        self.create_stats_card(stats_frame, "Total Assets", "total_assets")
        self.create_stats_card(stats_frame, "Brokerage Received", "received_count")
        self.create_stats_card(stats_frame, "Brokerage Pending", "pending_count")
        self.create_stats_card(stats_frame, "Received %", "received_percentage")

        # Enhanced TreeView section
        tree_frame = ttk.Frame(container)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(15, 20))

        # [Rest of the tree view code remains the same...]
        columns = ['S.No', 'Project', 'Tower', 'Floor', 'Unit No', 'Owner Name', 'Tenant', 'Brokerage', 'Lease Manager']

        self.tree = ttk.Treeview(tree_frame,
                                 columns=columns,
                                 show="headings",
                                 style='Custom.Treeview',
                                 height=15)

        col_widths = {
            'S.No': 60,
            'Project': 130,
            'Tower': 50,
            'Floor': 50,
            'Unit No': 50,
            'Owner Name': 220,
            'Tenant': 240,  # New column
            'Brokerage': 80,
            'Lease Manager': 70
        }

        for col in columns:
            self.tree.heading(col, text=col, anchor=tk.W)
            self.tree.column(col, width=col_widths[col], minwidth=col_widths[col])

        # Enhanced scrollbar
        scrollbar = ttk.Scrollbar(tree_frame,
                                  orient="vertical",
                                  command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Configure row colors with new professional styling
        self.tree.tag_configure('received',
                                background=self.status_colors['received'],
                                foreground=self.status_colors['received_text'])
        self.tree.tag_configure('pending',
                                background=self.status_colors['pending'],
                                foreground=self.status_colors['pending_text'])

        status_dropdown.bind('<<ComboboxSelected>>', lambda e: self.refresh_brokerage_data())
        self.tree.bind('<Double-1>', self.on_treeview_double_click)
        self.tree.bind('<Motion>', self.on_hover)
        self.tree.bind('<MouseWheel>', self._on_mousewheel)


        self.refresh_brokerage_data()

    def show_asset_details(self, values):
        # Find the full asset data from the sheet
        assets = self.app.sheet.get_all_values()
        headers = assets[0]

        # Match based on multiple columns to ensure correct asset
        full_row = next((row for row in assets[1:]
                         if row[0] == values[0]  # S.No
                         and row[4] == values[1]  # Project
                         and row[10] == values[2]), None)  # Tower

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
                              text=f"{full_row[4]}",  # Project name
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

            # Content groups
            groups = {
                'Basic Info': [0, 1, 2, 3, 4, 5],
                'Property Details': [6, 7, 8, 9, 10],
                'Unit Information': [11, 12, 13, 14, 15],
                'Tenant Details': [16, 17, 18, 19, 20],
                'Brokerage Information': [21, 22, 23]
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

            # Brokerage Status Update Section
            status_frame = ttk.Frame(scrollable_frame, style='DetailRow.TFrame')
            status_frame.pack(fill=tk.X, pady=5, padx=10)

            ttk.Label(status_frame,
                      text="Update Brokerage Status",
                      style='DetailLabel.TLabel',
                      font=('Segoe UI', 11, 'bold')).pack(anchor=tk.W, pady=(10, 5))

            status_content = ttk.Frame(status_frame, style='DetailRow.TFrame')
            status_content.pack(fill=tk.X, padx=15)

            # Status dropdown
            status_var = tk.StringVar(value=full_row[22] or 'Pending')
            status_dropdown = ttk.Combobox(status_content,
                                           textvariable=status_var,
                                           values=['Pending', 'Received'],
                                           state='readonly',
                                           width=20)
            status_dropdown.pack(side=tk.LEFT, padx=(0, 10))

            def update_brokerage_status():
                try:
                    # Find the row index in the sheet
                    row_index = next(i for i, row in enumerate(assets[1:], start=2)
                                     if row[0] == values[0])

                    # Update the cell in column W (index 22)
                    self.app.sheet.update_cell(row_index, 23, status_var.get())
                    messagebox.showinfo("Success", "Brokerage status updated successfully!")
                    self.refresh_brokerage_data()
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to update brokerage status: {str(e)}")

            update_button = ttk.Button(status_content,
                                       text="Update Status",
                                       style="primary.TButton",
                                       command=update_brokerage_status)
            update_button.pack(side=tk.LEFT)

            # Bottom button frame
            button_frame = ttk.Frame(main_container, style='Content.TFrame')
            button_frame.pack(fill=tk.X, pady=(15, 0))

            # View documents button
            def open_drive_documents():
                drive_link = full_row[29]  # Google Drive folder link
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



    def refresh_brokerage_data(self):
        self.tree.delete(*self.tree.get_children())
        all_rows = self.app.sheet.get_all_values()

        # Calculate total statistics
        total_assets = len(all_rows[1:])
        total_received = 0
        total_pending = 0

        # First pass for statistics
        for row in all_rows[1:]:
            try:
                status = row[22].lower() if row[22] else "pending"
                if status == "received":
                    total_received += 1
                elif status == "pending":
                    total_pending += 1
            except (ValueError, IndexError):
                continue

        # Calculate percentage
        received_percentage = (total_received / total_assets * 100) if total_assets > 0 else 0

        # Update statistics
        self.stat_labels["total_assets"].configure(text=str(total_assets))
        self.stat_labels["received_count"].configure(text=str(total_received))
        self.stat_labels["pending_count"].configure(text=str(total_pending))
        self.stat_labels["received_percentage"].configure(text=f"{received_percentage:.1f}%")

        # Populate TreeView with filtered data
        for row in all_rows[1:]:
            try:
                status = row[22].lower() if row[22] else "pending"

                if self.selected_status.get() == "All" or status.lower() == self.selected_status.get().lower():
                    values = [
                        row[0],   # S.No
                        row[4],   # Project
                        row[10],  # Tower
                        row[11],  # Floor
                        row[12],  # Unit No
                        row[19],  # Owner Name
                        row[20],  # Tenant (Column U)
                        row[22] if row[22] else "Pending",  # Brokerage
                        row[28],  # Lease Manager
                    ]

                    item = self.tree.insert("", "end", values=values)
                    self.tree.item(item, tags=(status,))

            except (ValueError, IndexError):
                continue

    def _on_mousewheel(self, event):
        if event.delta:
            move = -1 * (event.delta // 120)
        elif event.num == 4:
            move = -1
        elif event.num == 5:
            move = 1
        else:
            return

        self.tree.yview_scroll(move, "units")