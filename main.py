import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import date, timedelta
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, mm
import os
import tempfile
import subprocess
import uuid
import io
import sys
from PIL import Image as PILImage
from PIL import ImageTk

class QuotationGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Ajith Iron Works - Advanced Quotation Generator")
        self.root.geometry("1000x700")
        self.root.configure(bg="#f5f5f5")
        
        # Create a style
        style = ttk.Style()
        style.configure('TFrame', background='#f5f5f5')
        style.configure('TLabelframe', background='#f5f5f5')
        style.configure('TLabelframe.Label', font=('Arial', 10, 'bold'))
        style.configure('TButton', font=('Arial', 10, 'bold'))
        
        # Main frame with scrollbar
        container = ttk.Frame(root)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create a canvas with scrollbar
        self.canvas = tk.Canvas(container, bg="#f5f5f5")
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Logo section
        self.logo_path = None
        self.create_logo_section(self.scrollable_frame)
        
        # Title
        title_frame = ttk.Frame(self.scrollable_frame)
        title_frame.pack(fill=tk.X, padx=10, pady=5)
        
        title_label = ttk.Label(title_frame, text="Ajith Iron Works - Advanced Quotation Generator", 
                               font=("Arial", 16, "bold"))
        title_label.pack(anchor="w")
        
        # Create form
        self.create_form(self.scrollable_frame)
        
        # Create buttons
        button_frame = ttk.Frame(self.scrollable_frame)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        generate_btn = ttk.Button(button_frame, text="Generate Quotation", 
                                 command=self.generate_quotation)
        generate_btn.pack(side=tk.LEFT, padx=5)
        
        clear_btn = ttk.Button(button_frame, text="Clear Form", 
                              command=self.clear_form)
        clear_btn.pack(side=tk.LEFT, padx=5)
        
        add_item_btn = ttk.Button(button_frame, text="Add Item", 
                                command=self.add_item_row)
        add_item_btn.pack(side=tk.LEFT, padx=5)
        
        remove_item_btn = ttk.Button(button_frame, text="Remove Last Item", 
                                   command=self.remove_item_row)
        remove_item_btn.pack(side=tk.LEFT, padx=5)
        
        # Bind mouse wheel to canvas for scrolling
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
    
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def create_logo_section(self, parent):
        logo_frame = ttk.LabelFrame(parent, text="Company Logo", padding="10")
        logo_frame.pack(fill=tk.X, padx=10, pady=5)
        
        logo_btn_frame = ttk.Frame(logo_frame)
        logo_btn_frame.pack(side=tk.LEFT, padx=5)
        
        self.logo_label = ttk.Label(logo_btn_frame, text="No logo selected")
        self.logo_label.pack(side=tk.LEFT, padx=5)
        
        select_logo_btn = ttk.Button(logo_btn_frame, text="Select Logo", 
                                   command=self.select_logo)
        select_logo_btn.pack(side=tk.LEFT, padx=5)
        
        remove_logo_btn = ttk.Button(logo_btn_frame, text="Remove Logo", 
                                   command=self.remove_logo)
        remove_logo_btn.pack(side=tk.LEFT, padx=5)
        
        # Logo preview
        self.logo_preview_frame = ttk.Frame(logo_frame)
        self.logo_preview_frame.pack(side=tk.RIGHT, padx=5)
        self.logo_preview = ttk.Label(self.logo_preview_frame)
        self.logo_preview.pack()
    
    def select_logo(self):
        file_path = filedialog.askopenfilename(
            title="Select Logo Image",
            filetypes=(("Image files", "*.png *.jpg *.jpeg *.gif *.bmp"), ("All files", "*.*"))
        )
        
        if file_path:
            self.logo_path = file_path
            self.logo_label.config(text=os.path.basename(file_path))
            
            # Display logo preview
            try:
                img = PILImage.open(file_path)
                img = img.resize((100, 100), PILImage.LANCZOS)
                photo_img = ImageTk.PhotoImage(img)
                self.logo_preview.config(image=photo_img)
                self.logo_preview.image = photo_img  # Keep a reference
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load image: {str(e)}")
    
    def remove_logo(self):
        self.logo_path = None
        self.logo_label.config(text="No logo selected")
        self.logo_preview.config(image="")
    
    def create_form(self, parent):
        # Quotation details frame
        details_frame = ttk.LabelFrame(parent, text="Quotation Details", padding="10")
        details_frame.pack(fill=tk.X, padx=10, pady=5)
        
        details_inner_frame = ttk.Frame(details_frame)
        details_inner_frame.pack(fill=tk.X)
        
        # Left side
        left_frame = ttk.Frame(details_inner_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Issue date
        date_frame = ttk.Frame(left_frame)
        date_frame.pack(fill=tk.X, pady=2)
        ttk.Label(date_frame, text="Issue Date:").pack(side=tk.LEFT)
        self.issue_date = ttk.Entry(date_frame, width=20)
        self.issue_date.pack(side=tk.LEFT, padx=5)
        self.issue_date.insert(0, date.today().strftime("%Y-%m-%d"))
        
        # Quote number
        quote_frame = ttk.Frame(left_frame)
        quote_frame.pack(fill=tk.X, pady=2)
        ttk.Label(quote_frame, text="Quote #:").pack(side=tk.LEFT)
        self.quote_number = ttk.Entry(quote_frame, width=20)
        self.quote_number.pack(side=tk.LEFT, padx=5)
        
        # Right side
        right_frame = ttk.Frame(details_inner_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        
        # Customer ID
        cust_id_frame = ttk.Frame(right_frame)
        cust_id_frame.pack(fill=tk.X, pady=2)
        ttk.Label(cust_id_frame, text="Customer ID:").pack(side=tk.LEFT)
        self.customer_id = ttk.Entry(cust_id_frame, width=20)
        self.customer_id.pack(side=tk.LEFT, padx=5)
        
        # Valid until
        valid_frame = ttk.Frame(right_frame)
        valid_frame.pack(fill=tk.X, pady=2)
        ttk.Label(valid_frame, text="Valid Until:").pack(side=tk.LEFT)
        self.valid_until = ttk.Entry(valid_frame, width=20)
        self.valid_until.pack(side=tk.LEFT, padx=5)
        # Set default valid until date (7 days from today)
        valid_until_date = date.today() + timedelta(days=7)
        self.valid_until.insert(0, valid_until_date.strftime("%Y-%m-%d"))
        
        # Customer frame
        customer_frame = ttk.LabelFrame(parent, text="Customer Details", padding="10")
        customer_frame.pack(fill=tk.X, padx=10, pady=5)
        
        customer_inner_frame = ttk.Frame(customer_frame)
        customer_inner_frame.pack(fill=tk.X)
        
        # Customer name
        name_frame = ttk.Frame(customer_inner_frame)
        name_frame.pack(fill=tk.X, pady=2)
        ttk.Label(name_frame, text="Customer Name:").pack(side=tk.LEFT)
        self.customer_name = ttk.Entry(name_frame, width=40)
        self.customer_name.pack(side=tk.LEFT, padx=5)
        
        # Customer phone
        phone_frame = ttk.Frame(customer_inner_frame)
        phone_frame.pack(fill=tk.X, pady=2)
        ttk.Label(phone_frame, text="Customer Phone:").pack(side=tk.LEFT)
        self.customer_phone = ttk.Entry(phone_frame, width=40)
        self.customer_phone.pack(side=tk.LEFT, padx=5)
        
        # Customer address
        address_frame = ttk.Frame(customer_inner_frame)
        address_frame.pack(fill=tk.X, pady=2)
        ttk.Label(address_frame, text="Customer Address:").pack(side=tk.LEFT)
        self.customer_address = ttk.Entry(address_frame, width=40)
        self.customer_address.pack(side=tk.LEFT, padx=5)
        
        # Items frame
        items_frame = ttk.LabelFrame(parent, text="Item Details", padding="10")
        items_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Column headers
        headers_frame = ttk.Frame(items_frame)
        headers_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(headers_frame, text="No.", width=5).pack(side=tk.LEFT, padx=2)
        ttk.Label(headers_frame, text="Description", width=40).pack(side=tk.LEFT, padx=2)
        ttk.Label(headers_frame, text="Quantity", width=10).pack(side=tk.LEFT, padx=2)
        ttk.Label(headers_frame, text="Unit Price (Rs)", width=15).pack(side=tk.LEFT, padx=2)
        ttk.Label(headers_frame, text="Amount (Rs)", width=15).pack(side=tk.LEFT, padx=2)
        
        # Container for item rows
        self.items_container = ttk.Frame(items_frame)
        self.items_container.pack(fill=tk.X)
        
        # List to store item widgets
        self.item_rows = []
        
        # Add 10 initial item rows
        for i in range(10):
            self.add_item_row()
        
        # Summary frame
        summary_frame = ttk.LabelFrame(parent, text="Summary", padding="10")
        summary_frame.pack(fill=tk.X, padx=10, pady=5)
        
        summary_grid = ttk.Frame(summary_frame)
        summary_grid.pack(side=tk.RIGHT)
        
        # Subtotal
        ttk.Label(summary_grid, text="Subtotal:").grid(row=0, column=0, sticky="e", pady=2)
        self.subtotal = ttk.Entry(summary_grid, width=20)
        self.subtotal.grid(row=0, column=1, sticky="w", pady=2, padx=5)
        self.subtotal.insert(0, "0.00")
        
        # Discount
        ttk.Label(summary_grid, text="Discount (Rs):").grid(row=1, column=0, sticky="e", pady=2)
        self.discount = ttk.Entry(summary_grid, width=20)
        self.discount.grid(row=1, column=1, sticky="w", pady=2, padx=5)
        self.discount.insert(0, "0.00")
        
        # Tax rate
        ttk.Label(summary_grid, text="Tax Rate (%):").grid(row=2, column=0, sticky="e", pady=2)
        self.tax_rate = ttk.Entry(summary_grid, width=20)
        self.tax_rate.grid(row=2, column=1, sticky="w", pady=2, padx=5)
        self.tax_rate.insert(0, "0.00")
        
        # Calculate button
        calculate_btn = ttk.Button(summary_grid, text="Calculate Total", 
                                 command=self.calculate_total)
        calculate_btn.grid(row=3, column=0, columnspan=2, pady=10)
        
        # Total
        ttk.Label(summary_grid, text="TOTAL (Rs):").grid(row=4, column=0, sticky="e", pady=2)
        self.total = ttk.Entry(summary_grid, width=20)
        self.total.grid(row=4, column=1, sticky="w", pady=2, padx=5)
        self.total.insert(0, "0.00")
        
        # Terms and conditions frame
        terms_frame = ttk.LabelFrame(parent, text="Terms and Conditions", padding="10")
        terms_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.terms = tk.Text(terms_frame, width=80, height=5)
        self.terms.pack(fill=tk.X, pady=5)
        self.terms.insert(tk.END, "1. The prices in this quotation are valid for the period mentioned above.\n2. 50% advance payment is required to commence the work.\n3. You should pay the remaining amount of the bill upon completion of the work.\n4. Any modifications to the design after approval may incur additional charges.\n5. Delivery timeline will be confirmed upon receipt of advance payment.")
    
    def add_item_row(self):
        row_num = len(self.item_rows) + 1
        row_frame = ttk.Frame(self.items_container)
        row_frame.pack(fill=tk.X, pady=2)
        
        # Item number
        item_num = ttk.Label(row_frame, text=f"{row_num}.", width=5)
        item_num.pack(side=tk.LEFT, padx=2)
        
        # Description
        description = ttk.Entry(row_frame, width=40)
        description.pack(side=tk.LEFT, padx=2)
        
        # Quantity
        quantity = ttk.Entry(row_frame, width=10)
        quantity.pack(side=tk.LEFT, padx=2)
        quantity.insert(0, "1")
        
        # Unit price
        unit_price = ttk.Entry(row_frame, width=15)
        unit_price.pack(side=tk.LEFT, padx=2)
        unit_price.insert(0, "0.00")
        
        # Amount (auto-calculated)
        amount = ttk.Entry(row_frame, width=15)
        amount.pack(side=tk.LEFT, padx=2)
        amount.insert(0, "0.00")
        amount.config(state="readonly")
        
        # Bind calculation events
        def calculate_row_amount(event=None):
            try:
                qty = float(quantity.get() or 0)
                price = float(unit_price.get() or 0)
                row_amount = qty * price
                
                # Update the amount field
                amount.config(state="normal")
                amount.delete(0, tk.END)
                amount.insert(0, f"{row_amount:.2f}")
                amount.config(state="readonly")
            except ValueError:
                pass
        
        quantity.bind("<KeyRelease>", calculate_row_amount)
        unit_price.bind("<KeyRelease>", calculate_row_amount)
        
        # Store row widgets
        row_data = {
            "frame": row_frame,
            "number": item_num,
            "description": description,
            "quantity": quantity,
            "unit_price": unit_price,
            "amount": amount
        }
        
        self.item_rows.append(row_data)
    
    def remove_item_row(self):
        if len(self.item_rows) > 1:  # Keep at least one row
            row_data = self.item_rows.pop()
            row_data["frame"].destroy()
            
            # Renumber remaining rows
            for i, row in enumerate(self.item_rows, 1):
                row["number"].config(text=f"{i}.")
    
    def calculate_total(self):
        try:
            # Calculate subtotal from all item rows
            subtotal = 0
            for row in self.item_rows:
                try:
                    amount_str = row["amount"].get()
                    if amount_str:
                        subtotal += float(amount_str)
                except ValueError:
                    pass
            
            # Update subtotal field
            self.subtotal.delete(0, tk.END)
            self.subtotal.insert(0, f"{subtotal:.2f}")
            
            # Calculate final total
            discount = float(self.discount.get() or 0)
            tax_rate = float(self.tax_rate.get() or 0)
            
            net_amount = subtotal - discount
            tax_amount = net_amount * (tax_rate / 100)
            total = net_amount + tax_amount
            
            # Update total field
            self.total.delete(0, tk.END)
            self.total.insert(0, f"{total:.2f}")
        except Exception as e:
            messagebox.showerror("Calculation Error", f"Error calculating total: {str(e)}")
    
    def clear_form(self):
        # Reset dates
        self.issue_date.delete(0, tk.END)
        self.issue_date.insert(0, date.today().strftime("%Y-%m-%d"))
        self.valid_until.delete(0, tk.END)
        valid_until_date = date.today() + timedelta(days=7)
        self.valid_until.insert(0, valid_until_date.strftime("%Y-%m-%d"))
        
        # Clear basic fields
        self.quote_number.delete(0, tk.END)
        self.customer_id.delete(0, tk.END)
        self.customer_name.delete(0, tk.END)
        self.customer_phone.delete(0, tk.END)
        self.customer_address.delete(0, tk.END)
        
        # Clear all item rows
        for row_data in self.item_rows:
            row_data["description"].delete(0, tk.END)
            row_data["quantity"].delete(0, tk.END)
            row_data["quantity"].insert(0, "1")
            row_data["unit_price"].delete(0, tk.END)
            row_data["unit_price"].insert(0, "0.00")
            
            # Reset amount
            row_data["amount"].config(state="normal")
            row_data["amount"].delete(0, tk.END)
            row_data["amount"].insert(0, "0.00")
            row_data["amount"].config(state="readonly")
        
        # Reset summary fields
        self.subtotal.delete(0, tk.END)
        self.subtotal.insert(0, "0.00")
        self.discount.delete(0, tk.END)
        self.discount.insert(0, "0.00")
        self.tax_rate.delete(0, tk.END)
        self.tax_rate.insert(0, "0.00")
        self.total.delete(0, tk.END)
        self.total.insert(0, "0.00")
        
        # Reset terms
        self.terms.delete("1.0", tk.END)
        self.terms.insert(tk.END, "1. The prices in this quotation are valid for the period mentioned above.\n2. 50% advance payment is required to commence the work.\n3. You should pay the remaining amount of the bill upon completion of the work.\n4. Any modifications to the design after approval may incur additional charges.\n5. Delivery timeline will be confirmed upon receipt of advance payment.")
    
    def generate_quotation(self):
        try:
            # Calculate totals first
            self.calculate_total()
            
            # Validate inputs
            if not self.quote_number.get():
                messagebox.showerror("Error", "Quote number is required")
                return
            
            # Extract values for the PDF
            subtotal = float(self.subtotal.get() or 0)
            discount = float(self.discount.get() or 0)
            tax_rate = float(self.tax_rate.get() or 0)
            net_amount = subtotal - discount
            tax_amount = net_amount * (tax_rate / 100)
            total = net_amount + tax_amount
            
            # Create a unique filename
            filename = f"Ajith_Iron_Works_Quotation_{self.quote_number.get()}_{uuid.uuid4().hex[:6]}.pdf"
            file_path = os.path.join(os.path.expanduser("~"), "Documents", filename)
            
            # Ensure directory exists
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            
            # Create the PDF
            doc = SimpleDocTemplate(
                file_path,
                pagesize=A4,
                rightMargin=72,
                leftMargin=72,
                topMargin=72,
                bottomMargin=72
            )
            
            styles = getSampleStyleSheet()
            elements = []
            
            # Custom styles
            title_style = ParagraphStyle(
                'Title',
                parent=styles['Heading1'],
                fontSize=18,
                alignment=1,
                spaceAfter=6,
            )
            
            subtitle_style = ParagraphStyle(
                'Subtitle',
                parent=styles['Normal'],
                fontSize=14,
                alignment=1,
                spaceAfter=12,
            )
            
            address_style = ParagraphStyle(
                'Address',
                parent=styles['Normal'],
                fontSize=10,
                alignment=1,
                spaceAfter=0.1*inch,
            )
            
            section_title_style = ParagraphStyle(
                'SectionTitle',
                parent=styles['Heading3'],
                fontSize=12,
                spaceBefore=0.2*inch,
                spaceAfter=0.1*inch,
            )
            
            # Create style for table cell paragraphs
            table_cell_style = ParagraphStyle(
                'TableCell',
                parent=styles['Normal'],
                fontSize=10,
                leading=12
            )
            
            # Create header with logo and company info
            header_data = []
            
            # Add logo if available
            if self.logo_path:
                try:
                    logo_img = Image(self.logo_path, width=50, height=50)
                    header_data.append([logo_img, ""])
                except Exception as e:
                    print(f"Error loading logo: {e}")
                    # Add title without logo
                    header_data.append(["", ""])


            else:
                # Add empty cell for alignment
                header_data.append(["", ""])
            
            # Company info
            company_info = [
                Paragraph("Ajith Iron Works", title_style),
                Paragraph("QUOTATION", subtitle_style),
                Paragraph("No, 167/4, Bogahalandhawatta, Brahakmanagama,<br/>Pannipitiya.10230", address_style),
                Paragraph("Website: pending!...<br/>Phone: 0789926314<br/>Whatsapp: 0789926314", address_style)
            ]
            
            header_data[0][1] = company_info
            
            # Create header table
            header_table = Table(header_data, colWidths=[1.5*inch, 4*inch])
            header_table.setStyle(TableStyle([ 
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('ALIGN', (0, 0), (0, -1), 'CENTER'),
                ('ALIGN', (1, 0), (1, -1), 'CENTER'),
                ('LEFTPADDING', (0, 0), (-1, -1), 10),
                ('RIGHTPADDING', (0, 0), (-1, -1), 10),
            ]))
            
            elements.append(header_table)
            elements.append(Spacer(1, 0.2*inch))
            
            # Quote details table
            data = [
                ["ISSUE DATE", self.issue_date.get(), "QUOTE #", self.quote_number.get()],
                ["VALID UNTIL", self.valid_until.get(), "CUSTOMER ID", self.customer_id.get()]
            ]
            
            quote_table = Table(data, colWidths=[1.2*inch, 1.3*inch, 1.2*inch, 1.3*inch])
            quote_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (0, -1), colors.lightgrey),
                ('BACKGROUND', (2, 0), (2, -1), colors.lightgrey),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('PADDING', (0, 0), (-1, -1), 6),
            ]))
            elements.append(quote_table)
            elements.append(Spacer(1, 0.2*inch))
            
            # Customer details
            elements.append(Paragraph("CUSTOMER DETAILS", section_title_style))
            
            customer_data = []
            if self.customer_name.get():
                customer_data.append(["Customer Name:", self.customer_name.get()])
            if self.customer_phone.get():
                customer_data.append(["Phone:", self.customer_phone.get()])
            if self.customer_address.get():
                customer_data.append(["Address:", self.customer_address.get()])
                
            if customer_data:
                customer_table = Table(customer_data, colWidths=[1.5*inch, 3.5*inch])
                customer_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (0, -1), colors.lightgrey),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                    ('PADDING', (0, 0), (-1, -1), 6),
                ]))
                elements.append(customer_table)
            else:
                elements.append(Paragraph("No customer details provided", styles['Normal']))
            
            elements.append(Spacer(1, 0.2*inch))
            
            # Items table
            elements.append(Paragraph("ITEM DETAILS", section_title_style))
            
            # Prepare item data for table with paragraphs for wrapping text
            item_data = [["No.", "Description", "Qty", "Unit Price (Rs)", "Amount (Rs)"]]
            
            for i, row in enumerate(self.item_rows, 1):
                desc = row["description"].get()
                qty = row["quantity"].get()
                unit_price = row["unit_price"].get()
                amount = row["amount"].get()
                
                # Only include rows with a description
                if desc.strip():
                    # Use Paragraph objects to allow text wrapping
                    desc_para = Paragraph(desc, table_cell_style)
                    item_data.append([i, desc_para, qty, unit_price, amount])
            
            # If no items were added, add a placeholder row
            if len(item_data) == 1:
                item_data.append(["", "No items added", "", "", ""])
            
            # Create item table with dynamic row heights
            col_widths = [0.4*inch, 3.0*inch, 0.6*inch, 1.0*inch, 1.0*inch]
            item_table = Table(item_data, colWidths=col_widths)
            item_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('PADDING', (0, 0), (-1, -1), 6),
                ('ALIGN', (0, 0), (0, -1), 'CENTER'),
                ('ALIGN', (2, 0), (4, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ]))
            elements.append(item_table)
            elements.append(Spacer(1, 0.1*inch))
            
            # Summary table
            summary_data = [
                ["", "", "Subtotal:", f"{subtotal:,.2f}"],
                ["", "", "Discount:", f"{discount:,.2f}"],
                ["", "", "Net Amount:", f"{net_amount:,.2f}"],
                ["", "", f"Tax ({tax_rate}%):", f"{tax_amount:,.2f}"],
                ["", "", "TOTAL:", f"{total:,.2f}"]
            ]
            
            summary_table = Table(summary_data, colWidths=[3.0*inch, 1.0*inch, 1.0*inch, 1.0*inch])
            summary_table.setStyle(TableStyle([
                ('GRID', (2, 0), (3, -1), 0.5, colors.black),
                ('BACKGROUND', (2, 0), (2, -1), colors.lightgrey),
                ('BACKGROUND', (2, -1), (2, -1), colors.grey),
                ('TEXTCOLOR', (2, -1), (3, -1), colors.red),
                ('ALIGN', (2, 0), (3, -1), 'RIGHT'),
                ('PADDING', (2, 0), (3, -1), 6),
                ('FONTNAME', (2, -1), (3, -1), 'Helvetica-Bold'),
            ]))
            elements.append(summary_table)
            elements.append(Spacer(1, 0.2*inch))
            
            # Terms and conditions
            elements.append(Paragraph("TERMS AND CONDITIONS", section_title_style))
            terms_text = self.terms.get("1.0", tk.END).strip()
            
            # Split terms into paragraphs
            terms_paragraphs = terms_text.split('\n')
            for term in terms_paragraphs:
                if term.strip():
                    elements.append(Paragraph(term, styles['Normal']))
                    elements.append(Spacer(1, 0.05*inch))
            
            # Add signature section
            elements.append(Spacer(1, 0.3*inch))
            sig_data = [
                ["_______________________", "", "_______________________"],
                ["Authorized Signature", "", "Customer Signature"],
                ["Date: ________________", "", "Date: ________________"]
            ]
            
            sig_table = Table(sig_data, colWidths=[2.0*inch, 1.0*inch, 2.0*inch])
            sig_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (0, -1), 'CENTER'),
                ('ALIGN', (2, 0), (2, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ]))
            elements.append(sig_table)
            
            # Build PDF
            doc.build(elements)
            
            # Open the PDF
            if sys.platform.startswith('win'):
                os.startfile(file_path)
            elif sys.platform.startswith('darwin'):  # macOS
                subprocess.run(['open', file_path])
            else:  # Linux
                subprocess.run(['xdg-open', file_path])
            
            messagebox.showinfo("Success", f"Quotation has been generated and saved as:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate quotation: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = QuotationGenerator(root)
    root.mainloop()

