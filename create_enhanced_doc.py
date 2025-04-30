from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
import io
import requests
from PIL import Image

def set_cell_background(cell, color):
    """Set the background color of a cell."""
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), color)
    cell._tc.get_or_add_tcPr().append(shading_elm)

def add_page_break(doc):
    """Add a page break to the document."""
    doc.add_page_break()

def add_heading_with_style(doc, text, level=1, font_size=24, font_bold=True, font_color=RGBColor(0, 0, 0), align=WD_ALIGN_PARAGRAPH.LEFT):
    """Add a heading with custom styling."""
    heading = doc.add_heading(text, level=level)
    heading.alignment = align
    
    # Set font properties for the heading
    for run in heading.runs:
        run.font.size = Pt(font_size)
        run.font.bold = font_bold
        run.font.color.rgb = font_color
    
    return heading

def add_paragraph_with_style(doc, text, font_size=12, font_bold=False, font_color=RGBColor(0, 0, 0), align=WD_ALIGN_PARAGRAPH.LEFT, space_after=6):
    """Add a paragraph with custom styling."""
    paragraph = doc.add_paragraph()
    paragraph.alignment = align
    paragraph.space_after = Pt(space_after)
    
    run = paragraph.add_run(text)
    run.font.size = Pt(font_size)
    run.font.bold = font_bold
    run.font.color.rgb = font_color
    
    return paragraph

def add_test_case_table(doc, title, header_row, data_rows):
    """Add a test case table with proper formatting."""
    # Add a header for the table
    doc.add_heading(title, level=2)
    
    # Create the table
    table = doc.add_table(rows=1, cols=len(header_row))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Fill header row
    header_cells = table.rows[0].cells
    for i, header_text in enumerate(header_row):
        header_cells[i].text = header_text
        set_cell_background(header_cells[i], "3498DB")  # Blue background
        
        # Format the header text
        for paragraph in header_cells[i].paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(11)
                run.font.color.rgb = RGBColor(255, 255, 255)  # White text
    
    # Fill data rows
    for row_data in data_rows:
        row_cells = table.add_row().cells
        for i, cell_text in enumerate(row_data):
            row_cells[i].text = cell_text
            
            # Format the data cell text
            for paragraph in row_cells[i].paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for run in paragraph.runs:
                    run.font.size = Pt(10)
                    
                    # Highlight Pass/Fail status
                    if i == len(row_data) - 1:  # Last column (Status)
                        if cell_text.lower() == "pass":
                            run.font.color.rgb = RGBColor(39, 174, 96)  # Green
                            run.font.bold = True
                        elif cell_text.lower() == "fail":
                            run.font.color.rgb = RGBColor(231, 76, 60)  # Red
                            run.font.bold = True
    
    # Set column widths
    if len(header_row) == 6:  # If standard test case table with 6 columns
        table.columns[0].width = Inches(0.5)   # ID
        table.columns[1].width = Inches(1.5)   # Test Steps
        table.columns[2].width = Inches(1.5)   # Test Data
        table.columns[3].width = Inches(1.5)   # Expected Result
        table.columns[4].width = Inches(1.5)   # Actual Result
        table.columns[5].width = Inches(0.8)   # Status
    
    doc.add_paragraph('')  # Add some space after table
    return table

def add_bulleted_list(doc, items, font_size=12):
    """Add a bulleted list with custom formatting."""
    for item in items:
        paragraph = doc.add_paragraph(style='List Bullet')
        run = paragraph.add_run(item)
        run.font.size = Pt(font_size)
    
    doc.add_paragraph('')  # Add some space after list
    return

def download_image(url, filename):
    """Download an image from a URL and save it to the specified filename."""
    response = requests.get(url)
    if response.status_code == 200:
        with open(filename, 'wb') as f:
            f.write(response.content)
        return True
    else:
        print(f"Failed to download image from {url}")
        return False

def add_image_to_document(doc, image_path, width=None, height=None, centered=True):
    """Add an image to the document with optional width and height."""
    if width and height:
        doc.add_picture(image_path, width=width, height=height)
    elif width:
        doc.add_picture(image_path, width=width)
    elif height:
        doc.add_picture(image_path, height=height)
    else:
        doc.add_picture(image_path)
    
    if centered:
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    return doc.paragraphs[-1]

def add_table_with_images(doc, title, num_cols, data):
    """
    Add a table with images and text.
    
    Args:
        doc: Document object
        title: Table title
        num_cols: Number of columns
        data: List of tuples (image_path, title, description)
    """
    doc.add_heading(title, level=2)
    
    num_rows = (len(data) + num_cols - 1) // num_cols  # Ceiling division
    
    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Set even spacing for all columns
    for col in table.columns:
        col.width = Inches(6.0 / num_cols)
    
    # Fill table with data
    for i, (image_path, item_title, description) in enumerate(data):
        row_idx = i // num_cols
        col_idx = i % num_cols
        
        cell = table.cell(row_idx, col_idx)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        
        # Add image
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run()
        run.add_picture(image_path, width=Inches(2.0))
        
        # Add title
        paragraph = cell.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(item_title)
        run.font.bold = True
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(52, 152, 219)
        
        # Add description
        paragraph = cell.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = paragraph.add_run(description)
        run.font.size = Pt(10)
    
    doc.add_paragraph()
    return table

def create_enhanced_document():
    """Create a comprehensive document for ShopSleek E-commerce platform with visual enhancements."""
    doc = Document()
    
    # Set default paragraph spacing
    style = doc.styles['Normal']
    style.font.size = Pt(12)
    style.paragraph_format.space_after = Pt(6)
    
    # Download and prepare images
    os.makedirs("images", exist_ok=True)
    
    # Background for title page
    title_bg_url = "https://images.unsplash.com/photo-1557821552-17105176677c?q=80&w=1000"
    title_bg_path = "images/title_bg.jpg"
    download_image(title_bg_url, title_bg_path)
    
    # Feature icons
    feature_icons = {
        "ui": "https://cdn-icons-png.flaticon.com/512/2165/2165289.png",
        "search": "https://cdn-icons-png.flaticon.com/512/149/149852.png",
        "payment": "https://cdn-icons-png.flaticon.com/512/2169/2169854.png",
        "order": "https://cdn-icons-png.flaticon.com/512/411/411762.png",
        "user": "https://cdn-icons-png.flaticon.com/512/1077/1077063.png",
        "admin": "https://cdn-icons-png.flaticon.com/512/1705/1705331.png",
        "analytics": "https://cdn-icons-png.flaticon.com/512/2091/2091764.png"
    }
    
    for key, url in feature_icons.items():
        download_image(url, f"images/{key}_icon.png")
    
    # Architecture diagram
    arch_url = "https://www.softwaretestinghelp.com/wp-content/qa/uploads/2020/01/3-TIER-ARCHITECTURE.png"
    arch_path = "images/architecture.png"
    download_image(arch_url, arch_path)
    
    # Security icons
    security_icons = {
        "encryption": "https://cdn-icons-png.flaticon.com/512/2889/2889676.png",
        "access": "https://cdn-icons-png.flaticon.com/512/2037/2037333.png",
        "session": "https://cdn-icons-png.flaticon.com/512/6357/6357048.png",
        "audit": "https://cdn-icons-png.flaticon.com/512/3027/3027408.png"
    }
    
    for key, url in security_icons.items():
        download_image(url, f"images/{key}_icon.png")
    
    # Deployment icons
    deployment_icons = {
        "hosting": "https://cdn-icons-png.flaticon.com/512/1373/1373585.png",
        "server": "https://cdn-icons-png.flaticon.com/512/9691/9691568.png",
        "database": "https://cdn-icons-png.flaticon.com/512/8425/8425464.png",
        "backup": "https://cdn-icons-png.flaticon.com/512/9494/9494132.png"
    }
    
    for key, url in deployment_icons.items():
        download_image(url, f"images/{key}_icon.png")
    
    # Create a section for document properties
    section = doc.sections[0]
    section.page_width = Cm(21)  # A4 width
    section.page_height = Cm(29.7)  # A4 height
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    
    # Title Page with background
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Add background image to title page
    title_run = title_para.add_run()
    title_run.add_picture(title_bg_path, width=Inches(7))
    
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_para.space_before = Pt(40)
    title_run = title_para.add_run("ShopSleek")
    title_run.font.size = Pt(44)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(52, 152, 219)
    
    subtitle_para = doc.add_paragraph()
    subtitle_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle_run = subtitle_para.add_run("E-commerce Platform")
    subtitle_run.font.size = Pt(32)
    subtitle_run.font.bold = True
    subtitle_run.font.color.rgb = RGBColor(44, 62, 80)
    
    # Add space
    doc.add_paragraph()
    doc.add_paragraph()
    
    # Author information with styled formatting
    author_para = doc.add_paragraph()
    author_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    author_run = author_para.add_run("Developed by: ")
    author_run.font.size = Pt(14)
    author_run.font.italic = True
    
    author_name_run = author_para.add_run("Hardik Kalathiya")
    author_name_run.font.size = Pt(14)
    author_name_run.font.bold = True
    
    guide_para = doc.add_paragraph()
    guide_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    guide_run = guide_para.add_run("Under the Guidance of: ")
    guide_run.font.size = Pt(14)
    guide_run.font.italic = True
    
    guide_name_run = guide_para.add_run("Dr. Maitri Jhaveri")
    guide_name_run.font.size = Pt(14)
    guide_name_run.font.bold = True
    
    dept_para = doc.add_paragraph()
    dept_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    dept_run = dept_para.add_run("Department of Computer Science")
    dept_run.font.size = Pt(14)
    dept_run.font.bold = True
    
    univ_para = doc.add_paragraph()
    univ_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    univ_run = univ_para.add_run("Gujarat University")
    univ_run.font.size = Pt(14)
    univ_run.font.bold = True
    
    add_page_break(doc)
    
    # Table of Contents
    toc_title = add_heading_with_style(doc, "Table of Contents", font_size=18, font_color=RGBColor(52, 152, 219), align=WD_ALIGN_PARAGRAPH.CENTER)
    
    toc_items = [
        "1. Project Overview",
        "2. Key Features",
        "3. System Architecture",
        "4. Test Cases",
        "   4.1 User Authentication Test Cases",
        "   4.2 Product Search Test Cases",
        "   4.3 Shopping Cart Test Cases",
        "   4.4 Checkout Process Test Cases",
        "   4.5 Test Results Summary",
        "5. Security Aspects",
        "   5.1 Data Protection",
        "   5.2 Security Implementation",
        "6. Deployment Architecture",
        "   6.1 Hosting Infrastructure",
        "   6.2 Server Configuration",
        "   6.3 Database Setup",
        "   6.4 Backup and Monitoring",
        "7. Future Scope"
    ]
    
    # Create styled table of contents
    for item in toc_items:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        para.paragraph_format.left_indent = Pt(20)
        para.paragraph_format.space_after = Pt(10)
        
        run = para.add_run(item)
        run.font.size = Pt(12)
        
        # Primary headings in bold and blue
        if len(item) <= 20 and not item.startswith("   "):
            run.font.bold = True
            run.font.color.rgb = RGBColor(52, 152, 219)
    
    add_page_break(doc)
    
    # 1. Project Overview
    overview_title = add_heading_with_style(doc, "1. Project Overview", font_size=18, font_color=RGBColor(52, 152, 219), align=WD_ALIGN_PARAGRAPH.CENTER)
    
    # Add decorative line
    line_para = doc.add_paragraph()
    line_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    line_run = line_para.add_run("_________________________________________")
    line_run.font.color.rgb = RGBColor(52, 152, 219)
    line_run.font.size = Pt(14)
    
    ecommerce_img_url = "https://img.freepik.com/free-vector/e-commerce-icon-robotic-hand-online-shop-logo-shopping-cart-icon-digital-marketing-concept-blue-background-3d-rendering_56104-1477.jpg"
    ecommerce_img_path = "images/ecommerce.jpg"
    download_image(ecommerce_img_url, ecommerce_img_path)
    
    # Add image
    add_image_to_document(doc, ecommerce_img_path, width=Inches(4))
    
    overview_text = """
ShopSleek is a comprehensive e-commerce platform developed to provide users with a seamless online shopping experience. The platform is designed to handle various aspects of online retail, from product browsing and search to secure checkout and order management.

The system aims to provide an intuitive interface for customers while offering robust backend management capabilities for administrators. ShopSleek supports multiple device compatibility through responsive design and incorporates industry standard security practices to protect user data and transactions.
"""
    add_paragraph_with_style(doc, overview_text)
    
    # Core Functionalities with two-column format
    add_heading_with_style(doc, "Core Functionalities:", level=2, font_size=16, font_color=RGBColor(44, 62, 80))
    
    core_functions = [
        "Online shopping platform for various products",
        "User registration and account management",
        "Guest shopping capabilities",
        "Product search and filtering",
        "Shopping cart management",
        "Secure checkout process",
        "Order tracking",
        "Admin dashboard for system management",
        "Multi-vendor support system",
        "Reporting and analytics",
        "Secure payment processing",
        "Responsive design for all devices"
    ]
    
    # Create two columns for core functions
    table = doc.add_table(rows=len(core_functions)//2, cols=2)
    table.style = 'Light Grid Accent 1'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    for i in range(len(core_functions)//2):
        for j in range(2):
            idx = i*2 + j
            if idx < len(core_functions):
                cell = table.cell(i, j)
                paragraph = cell.paragraphs[0]
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                run = paragraph.add_run("✓ " + core_functions[idx])
                run.font.size = Pt(11)
    
    doc.add_paragraph()
    
    add_page_break(doc)
    
    # 2. Key Features with visual icons
    add_heading_with_style(doc, "2. Key Features", font_size=18, font_color=RGBColor(52, 152, 219), align=WD_ALIGN_PARAGRAPH.CENTER)
    
    # Add decorative line
    line_para = doc.add_paragraph()
    line_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    line_run = line_para.add_run("_________________________________________")
    line_run.font.color.rgb = RGBColor(52, 152, 219)
    line_run.font.size = Pt(14)
    
    features_text = """
ShopSleek offers a range of features designed to enhance both the user shopping experience and the administrative capabilities of the platform. The following are the key features that set ShopSleek apart from other e-commerce solutions:
"""
    add_paragraph_with_style(doc, features_text)
    
    # Add key features with icons in a table
    feature_data = [
        (
            "images/ui_icon.png", 
            "Intuitive User Interface", 
            "Clean, responsive design optimized for all devices with thoughtful navigation paths and minimal friction points."
        ),
        (
            "images/search_icon.png", 
            "Advanced Search & Filtering", 
            "Find products quickly with dynamic filtering capabilities that include category browsing, price range filters, and brand selection."
        ),
        (
            "images/payment_icon.png", 
            "Secure Payment Processing", 
            "PCI-DSS compliant payment gateway integration supporting multiple payment methods including credit/debit cards and digital wallets."
        ),
        (
            "images/order_icon.png", 
            "Order Management", 
            "Track orders from placement to delivery with real-time updates and notifications for both users and administrators."
        ),
        (
            "images/user_icon.png",

            "User Accounts", 
            "Personal profiles with wishlists, order history, saved addresses, and preference settings for a personalized experience."
        ),
        (
            "images/admin_icon.png", 
            "Admin Dashboard", 
            "Complete control over products, orders, users, and inventory with an intuitive administrative interface."
        ),
        (
            "images/analytics_icon.png", 
            "Analytics & Reporting", 
            "Comprehensive sales reports, user activity tracking, and inventory level monitoring with visual data representation."
        )
    ]
    
    add_table_with_images(doc, "Key Feature Details", 3, feature_data)
    
    add_page_break(doc)
    
    # 3. System Architecture with diagram
    add_heading_with_style(doc, "3. System Architecture", font_size=18, font_color=RGBColor(52, 152, 219), align=WD_ALIGN_PARAGRAPH.CENTER)
    
    # Add decorative line
    line_para = doc.add_paragraph()
    line_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    line_run = line_para.add_run("_________________________________________")
    line_run.font.color.rgb = RGBColor(52, 152, 219)
    line_run.font.size = Pt(14)
    
    architecture_text = """
ShopSleek follows a modern three-tier architecture that separates the application into distinct layers, each with its own responsibilities. This architectural approach enhances maintainability, scalability, and security by isolating different components of the system.
"""
    add_paragraph_with_style(doc, architecture_text)
    
    # Add architecture diagram
    add_image_to_document(doc, arch_path, width=Inches(6))
    
    # Add caption for the architecture diagram
    caption_para = doc.add_paragraph()
    caption_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption_run = caption_para.add_run("Figure 1: ShopSleek Three-Tier Architecture Diagram")
    caption_run.font.italic = True
    caption_run.font.size = Pt(10)
    
    # Add Three-Tier Architecture section
    add_heading_with_style(doc, "Three-Tier Architecture:", level=2, font_size=16, font_color=RGBColor(44, 62, 80))
    
    three_tier = [
        "Presentation Layer: The top-most layer that directly interacts with users through the User Interface, Admin Dashboard, and Vendor Panel. This layer is responsible for rendering data in a user-friendly format and collecting user inputs.",
        
        "Application Layer: The middle layer that processes business logic, including User Management, Product Management, Order Processing, and Payment Gateway integration. This layer acts as an intermediary between the presentation and data layers, applying business rules and ensuring data integrity.",
        
        "Data Layer: The foundation layer that handles data storage and retrieval through MySQL Database and File Storage. This layer is optimized for data persistence, security, and efficient query processing."
    ]
    
    # Create a visually appealing table for the three-tier architecture
    table = doc.add_table(rows=3, cols=2)
    table.style = 'Light Grid Accent 1'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Set column widths
    table.columns[0].width = Inches(1.5)
    table.columns[1].width = Inches(4.5)
    
    # Add layer icons and descriptions
    layer_icons = {
        "Presentation Layer": "https://cdn-icons-png.flaticon.com/512/1055/1055687.png",
        "Application Layer": "https://cdn-icons-png.flaticon.com/512/2822/2822774.png",
        "Data Layer": "https://cdn-icons-png.flaticon.com/512/1045/1045458.png"
    }
    
    for i, item in enumerate(three_tier):
        layer_name = item.split(": ")[0]
        layer_desc = item.split(": ")[1]
        
        # Download layer icon if not already downloaded
        icon_path = f"images/{layer_name.lower().replace(' ', '_')}_icon.png"
        if not os.path.exists(icon_path):
            download_image(layer_icons[layer_name], icon_path)
        
        # Add icon to left cell
        left_cell = table.cell(i, 0)
        left_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        left_para = left_cell.paragraphs[0]
        left_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        left_run = left_para.add_run()
        left_run.add_picture(icon_path, width=Inches(1.0))
        
        # Add text to right cell
        right_cell = table.cell(i, 1)
        right_para = right_cell.paragraphs[0]
        right_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        # Add layer name
        name_run = right_para.add_run(layer_name + ":\n")
        name_run.font.bold = True
        name_run.font.size = Pt(13)
        name_run.font.color.rgb = RGBColor(52, 152, 219)
        
        # Add layer description
        desc_run = right_para.add_run(layer_desc)
        desc_run.font.size = Pt(12)
    
    doc.add_paragraph()
    
    # Add Technology Stack section
    add_heading_with_style(doc, "Technology Stack:", level=2, font_size=16, font_color=RGBColor(44, 62, 80))
    
    tech_stack = {
        "Frontend": ("React.js, Tailwind CSS", "Modern JavaScript framework with utility-first CSS for building responsive user interfaces"),
        "Backend": ("Node.js, Express.js", "JavaScript runtime environment with a minimalist web framework for building APIs and server-side logic"),
        "Database": ("MySQL", "Relational database management system for structured data storage with powerful query capabilities"),
        "Server": ("Node.js", "Lightweight, efficient, and scalable server environment for handling HTTP requests"),
        "Hosting": ("HostingRaja", "Reliable hosting service with high uptime guarantee and scalable infrastructure")
    }
    
    # Download tech stack icons
    tech_icons = {
        "Frontend": "https://cdn-icons-png.flaticon.com/512/5968/5968672.png",
        "Backend": "https://cdn-icons-png.flaticon.com/512/5968/5968322.png",
        "Database": "https://cdn-icons-png.flaticon.com/512/3161/3161158.png",
        "Server": "https://cdn-icons-png.flaticon.com/512/6213/6213822.png",
        "Hosting": "https://cdn-icons-png.flaticon.com/512/5946/5946387.png"
    }
    
    for tech, url in tech_icons.items():
        tech_icon_path = f"images/{tech.lower()}_icon.png"
        if not os.path.exists(tech_icon_path):
            download_image(url, tech_icon_path)
    
    # Create tech stack table
    table = doc.add_table(rows=len(tech_stack), cols=3)
    table.style = 'Medium Grid 1 Accent 1'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Set column widths
    table.columns[0].width = Inches(1.0)
    table.columns[1].width = Inches(1.5)
    table.columns[2].width = Inches(3.5)
    
    # Fill the tech stack table
    for i, (tech, (components, description)) in enumerate(tech_stack.items()):
        # Add icon
        icon_cell = table.cell(i, 0)
        icon_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        icon_para = icon_cell.paragraphs[0]
        icon_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        icon_run = icon_para.add_run()
        icon_run.add_picture(f"images/{tech.lower()}_icon.png", width=Inches(0.8))
        
        # Add technology name
        name_cell = table.cell(i, 1)
        name_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        name_para = name_cell.paragraphs[0]
        name_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        # Technology name
        tech_run = name_para.add_run(tech + ":\n")
        tech_run.font.bold = True
        tech_run.font.size = Pt(12)
        
        # Components
        comp_run = name_para.add_run(components)
        comp_run.font.size = Pt(11)
        comp_run.font.italic = True
        
        # Add description
        desc_cell = table.cell(i, 2)
        desc_para = desc_cell.paragraphs[0]
        desc_run = desc_para.add_run(description)
        desc_run.font.size = Pt(11)
    
    add_page_break(doc)
    
    # 4. Test Cases
    add_heading_with_style(doc, "4. Test Cases", font_size=18, font_color=RGBColor(52, 152, 219), align=WD_ALIGN_PARAGRAPH.CENTER)
    
    # Add decorative line
    line_para = doc.add_paragraph()
    line_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    line_run = line_para.add_run("_________________________________________")
    line_run.font.color.rgb = RGBColor(52, 152, 219)
    line_run.font.size = Pt(14)
    
    # Download test case image
    test_image_url = "https://www.techslang.com/wp-content/uploads/2021/11/what-is-test-case-definition-1024x536.png"
    test_image_path = "images/test_case.png"
    download_image(test_image_url, test_image_path)
    
    # Add test case image
    add_image_to_document(doc, test_image_path, width=Inches(5))
    
    test_intro_text = """
Comprehensive testing was conducted to ensure the platform's functionality, reliability, and security. Each test case was carefully designed to validate specific features and identify potential issues before deployment. The following sections present the test cases organized by functional categories.
"""
    add_paragraph_with_style(doc, test_intro_text)
    
    # 4.1 User Authentication Test Cases
    add_heading_with_style(doc, "4.1 User Authentication Test Cases", level=2, font_size=16, font_color=RGBColor(44, 62, 80))
    
    auth_test_header = ["ID", "Test Steps", "Test Data", "Expected Result", "Actual Result", "Status"]
    
    auth_test_data = [
        ["TC01", "1. Navigate to login page\n2. Enter valid credentials\n3. Click on Login button", "Email: user@example.com\nPassword: Correct123", "User logged in and redirected to dashboard", "User successfully logged in and redirected to dashboard", "Pass"],
        
        ["TC02", "1. Navigate to login page\n2. Enter invalid credentials\n3. Click on Login button", "Email: user@example.com\nPassword: WrongPass123", "Error message displayed", "Error message: 'Invalid email or password'", "Pass"],
        
        ["TC03", "1. Navigate to login page\n2. Leave fields blank\n3. Click on Login button", "Email: [blank]\nPassword: [blank]", "Form validation errors", "Error messages displayed for all required fields", "Pass"],
        
        ["TC04", "1. Navigate to registration page\n2. Fill all required fields\n3. Click Register button", "Name: John Doe\nEmail: john@example.com\nPassword: SecurePass123", "Account created with confirmation email sent", "Account created and confirmation email received", "Pass"],
        
        ["TC05", "1. Navigate to registration page\n2. Leave all fields empty\n3. Click Register button", "All fields empty", "Error message displayed for each required field", "Error message displayed: 'All fields are required'", "Pass"],
        
        ["TC06", "1. Navigate to registration page\n2. Enter invalid email format\n3. Click Register button", "Name: Jane Smith\nEmail: jane.emailcom\nPassword: Pass123", "Validation error shown for email field", "Validation error: 'Please enter a valid email address'", "Pass"]
    ]
    
    add_test_case_table(doc, "User Authentication Test Cases", auth_test_header, auth_test_data)
    
    # 4.2 Product Search Test Cases
    add_heading_with_style(doc, "4.2 Product Search Test Cases", level=2, font_size=16, font_color=RGBColor(44, 62, 80))
    
    search_test_header = ["ID", "Test Steps", "Test Data", "Expected Result", "Actual Result", "Status"]
    
    search_test_data = [
        ["TC07", "1. Navigate to search bar\n2. Enter valid search term\n3. Press search button", "Search term: 'smartphone'", "Relevant smartphone products displayed", "5 smartphone products displayed with images and prices", "Pass"],
        
        ["TC08", "1. Navigate to search bar\n2. Enter search term with no matches\n3. Press search button", "Search term: 'nonexistentitem123'", "Empty results message shown", "Message: 'No products match your search criteria'", "Pass"],
        
        ["TC09", "1. Navigate to category menu\n2. Select Electronics category\n3. View filtered results", "Category: Electronics", "Products from Electronics category displayed", "15 Electronics products displayed correctly", "Pass"],
        
        ["TC10", "1. Apply price filter\n2. Set minimum price to ₹5000\n3. Set maximum price to ₹15000\n4. Apply filter", "Price range: ₹5000 - ₹15000", "Products within price range displayed", "8 products displayed with prices between ₹5000-₹15000", "Pass"],
        
        ["TC11", "1. Apply multiple filters\n2. Select Electronics category\n3. Select brand 'Samsung'\n4. Set price range ₹10000-₹20000", "Category: Electronics\nBrand: Samsung\nPrice: ₹10000-₹20000", "Filtered products matching all criteria", "3 Samsung Electronics products in price range displayed", "Pass"],
        
        ["TC12", "1. Search for out of stock item\n2. Locate an item marked as 'Out of Stock'\n3. Attempt to add to cart", "Product: 'Limited Edition Smartphone'", "Add to Cart button disabled with message", "Button disabled with 'Out of Stock' message overlay", "Fail"]
    ]
    
    add_test_case_table(doc, "Product Search Test Cases", search_test_header, search_test_data)
    
    add_page_break(doc)
    
    # 4.3 Shopping Cart Test Cases
    add_heading_with_style(doc, "4.3 Shopping Cart Test Cases", level=2, font_size=16, font_color=RGBColor(44, 62, 80))
    
    cart_test_header = ["ID", "Test Steps", "Test Data", "Expected Result", "Actual Result", "Status"]
    
    cart_test_data = [
        ["TC13", "1. View product details\n2. Click 'Add to Cart'\n3. Verify cart updates", "Product: Smartphone XYZ\nQuantity: 1", "Product added to cart with correct price", "Product successfully added to cart with price ₹15,999", "Pass"],
        
        ["TC14", "1. View shopping cart\n2. Update quantity to 3\n3. Check total price", "Product: Smartphone XYZ\nNew Quantity: 3", "Quantity updated and total price recalculated", "Quantity updated to 3, total price ₹47,997", "Pass"],
        
        ["TC15", "1. View shopping cart\n2. Click 'Remove' button\n3. Verify item removed", "Product: Smartphone XYZ", "Item removed from cart", "Item successfully removed from cart, cart empty", "Pass"],
        
        ["TC16", "1. Add multiple items to cart\n2. Verify cart shows correct items\n3. Check total price", "Products: Smartphone (₹15,999), Headphones (₹1,999)", "Cart shows both items with correct total", "Cart displays both items, total price ₹17,998", "Pass"],
        
        ["TC17", "1. View cart with items\n2. Close browser\n3. Reopen site\n4. Check if cart persists", "Products: Smartphone (₹15,999)", "Cart state should persist across sessions", "Cart items restored upon returning to site", "Pass"],
        
        ["TC18", "1. Add item to cart\n2. Attempt to add more units than available stock\n3. Observe system response", "Product: Limited Edition Watch\nAvailable Stock: 2\nAttempted Quantity: 5", "System should limit to available quantity", "Error message 'Only 2 units available'", "Pass"]
    ]
    
    add_test_case_table(doc, "Shopping Cart Test Cases", cart_test_header, cart_test_data)
    
    # 4.4 Checkout Process Test Cases
    add_heading_with_style(doc, "4.4 Checkout Process Test Cases", level=2, font_size=16, font_color=RGBColor(44, 62, 80))
    
    checkout_test_header = ["ID", "Test Steps", "Test Data", "Expected Result", "Actual Result", "Status"]
    
    checkout_test_data = [
        ["TC19", "1. Add products to cart\n2. Click 'Checkout' button\n3. Enter shipping info\n4. Select payment method\n5. Confirm order", "Valid shipping & payment details", "Order placed successfully", "Order placed with confirmation #12345", "Pass"],
        
        ["TC20", "1. Add products to cart\n2. Click 'Checkout' button\n3. Enter invalid card details\n4. Attempt to confirm", "Invalid card number: 1234 5678 9012 3456", "Payment error displayed", "Error: 'Invalid card number'", "Pass"],
        
        ["TC21", "1. Add products to cart\n2. Click 'Checkout' button\n3. Leave shipping fields empty\n4. Attempt to confirm", "Empty shipping fields", "Validation errors for required fields", "Error messages shown for all required fields", "Pass"],
        
        ["TC22", "1. Complete checkout process\n2. Check email for order confirmation\n3. Verify order details", "Completed order with confirmation", "Order confirmation email received", "Email received with correct order details and tracking info", "Pass"],
        
        ["TC23", "1. Add products to cart\n2. Begin checkout process\n3. Use coupon code\n4. Verify discount applied", "Coupon code: WELCOME10", "10% discount applied to order total", "Discount successfully applied, total reduced by 10%", "Pass"],
        
        ["TC24", "1. Add products to cart\n2. Begin checkout process\n3. Select cash on delivery\n4. Complete order", "Payment method: Cash on Delivery", "Order placed with COD payment method", "Order tracking not updating status correctly", "Fail"]
    ]
    
    add_test_case_table(doc, "Checkout Process Test Cases", checkout_test_header, checkout_test_data)
    
    add_page_break(doc)
    
    # 4.5 Test Results Summary
    add_heading_with_style(doc, "4.5 Test Results Summary", level=2, font_size=16, font_color=RGBColor(44, 62, 80))
    
    summary_text = """
A total of 36 test cases were executed across different functional areas of the ShopSleek platform. The following table summarizes the test results by category:
"""
    add_paragraph_with_style(doc, summary_text)
    
    summary_header = ["Category", "Total Tests", "Passed", "Failed", "Success Rate"]
    
    summary_data = [
        ["User Authentication", "6", "6", "0", "100%"],
        ["Product Browsing", "6", "5", "1", "83.3%"],
        ["Shopping & Checkout", "6", "5", "1", "83.3%"],
        ["Admin Functions", "8", "7", "1", "87.5%"],
        ["API Endpoints", "10", "9", "1", "90%"],
        ["Overall System", "36", "32", "4", "88.9%"]
    ]
    
    # Create a summary table
    table = doc.add_table(rows=1, cols=len(summary_header))
    table.style = 'Colorful Grid Accent 1'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Fill header row
    header_cells = table.rows[0].cells
    for i, header_text in enumerate(summary_header):
        header_cells[i].text = header_text
        set_cell_background(header_cells[i], "3498DB")  # Blue background
        
        # Format the header text
        for paragraph in header_cells[i].paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(12)
                run.font.color.rgb = RGBColor(255, 255, 255)  # White text
    
    # Fill data rows
    for row_data in summary_data:
        row_cells = table.add_row().cells
        
        # Fill cell data
        for i, cell_text in enumerate(row_data):
            row_cells[i].text = cell_text
            
            # Format the cell text
            for paragraph in row_cells[i].paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.size = Pt(11)
                    
                    # For "Overall System" row, make it bold
                    if row_data[0] == "Overall System":
                        run.font.bold = True
                        set_cell_background(row_cells[i], "E5F0FF")  # Light blue background
    
    # Add analysis of test results
    doc.add_paragraph("")
    
    # Add test results chart
    chart_url = "https://quickchart.io/chart?c={type:'pie',data:{labels:['Passed','Failed'],datasets:[{data:[32,4],backgroundColor:['%2327AE60','%23E74C3C']}]}}"
    chart_path = "images/test_chart.png"
    download_image(chart_url, chart_path)
    
    add_image_to_document(doc, chart_path, width=Inches(4))
    
    # Add caption for the chart
    caption_para = doc.add_paragraph()
    caption_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption_run = caption_para.add_run("Figure 2: Test Results Summary Chart (88.9% Pass Rate)")
    caption_run.font.italic = True
    caption_run.font.size = Pt(10)
    
    test_analysis = """
The test results indicate an overall success rate of 88.9%, which demonstrates that the ShopSleek platform is functioning well in most areas. The User Authentication functionality achieved a perfect score, indicating robust security for user accounts. 

Areas requiring attention include specific features in Product Browsing and Shopping & Checkout, where failures were identified. These issues have been documented and prioritized for resolution before the final deployment.
"""
    add_paragraph_with_style(doc, test_analysis)
    
    add_page_break(doc)
    
    # 5. Security Aspects
    add_heading_with_style(doc, "5. Security Aspects", font_size=18, font_color=RGBColor(52, 152, 219), align=WD_ALIGN_PARAGRAPH.CENTER)
    
    # Add decorative line
    line_para = doc.add_paragraph()
    line_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    line_run = line_para.add_run("_________________________________________")
    line_run.font.color.rgb = RGBColor(52, 152, 219)
    line_run.font.size = Pt(14)
    
    # Add security image
    security_url = "https://img.freepik.com/premium-vector/online-shop-security-image-lock-icon-safety-concept-shield-with-check-mark-symbol-cyber-security-ecommerce-store_185386-324.jpg"
    security_path = "images/security.jpg"
    download_image(security_url, security_path)
    
    add_image_to_document(doc, security_path, width=Inches(4))
    
    security_intro = """
Security is a critical aspect of the ShopSleek e-commerce platform, especially considering the sensitive nature of user data and payment information being processed. The platform implements multiple layers of security measures to protect both user data and system integrity.
"""
    add_paragraph_with_style(doc, security_intro)
    
    # 5.1 Data Protection with visual icons
    add_heading_with_style(doc, "5.1 Data Protection", level=2, font_size=16, font_color=RGBColor(44, 62, 80))
    
    # Create data protection info with icons
    protection_data = [
        (
            "images/encryption_icon.png",
            "Password Encryption",
            "All user passwords are encrypted using bcrypt hashing algorithm with salt, ensuring that even in the event of a database breach, passwords remain secure."
        ),
        (
            "images/access_icon.png",
            "Role-based Access Control",
            "The system implements granular access controls that restrict users to only the data and functions necessary for their role."
        ),
        (
            "images/session_icon.png",
            "Session Management",
            "Secure session handling with automatic timeout after periods of inactivity. Sessions are invalidated upon logout and implement secure cookies."
        ),
        (
            "images/audit_icon.png",
            "Regular Security Audits",
            "Scheduled security assessments and penetration testing by third-party security experts to identify and address vulnerabilities."
        )
    ]
    
    add_table_with_images(doc, "Data Protection Measures", 2, protection_data)
    
    # 5.2 Security Implementation
    add_heading_with_style(doc, "5.2 Security Implementation", level=2, font_size=16, font_color=RGBColor(44, 62, 80))
    
    # Download security implementation icons
    security_impl_icons = {
        "sql": "https://cdn-icons-png.flaticon.com/512/6119/6119867.png",
        "xss": "https://cdn-icons-png.flaticon.com/512/2374/2374510.png",
        "csrf": "https://cdn-icons-png.flaticon.com/512/2885/2885417.png",
        "brute": "https://cdn-icons-png.flaticon.com/512/5527/5527938.png"
    }
    
    for key, url in security_impl_icons.items():
        icon_path = f"images/{key}_icon.png"
        if not os.path.exists(icon_path):
            download_image(url, icon_path)
    
    # Create security implementation table with visuals
    security_implementation_data = [
        (
            "images/sql_icon.png",
            "SQL Injection Protection",
            "All database interactions use prepared statements and parameterized queries rather than direct string concatenation."
        ),
        (
            "images/xss_icon.png",
            "XSS Attack Prevention",
            "Implementation of Content Security Policy (CSP) headers and output encoding for all user-generated content."
        ),
        (
            "images/csrf_icon.png",
            "CSRF Attack Mitigation",
            "Unique tokens are generated for each session and validated for all state-changing operations."
        ),
        (
            "images/brute_icon.png",
            "Brute Force Attack Prevention",
            "Progressive account lockout after multiple failed login attempts, with CAPTCHA implementation to prevent automated attacks."
        )
    ]
    
    add_table_with_images(doc, "Security Implementation Features", 2, security_implementation_data)
    
    add_page_break(doc)
    
    # 6. Deployment Architecture
    add_heading_with_style(doc, "6. Deployment Architecture", font_size=18, font_color=RGBColor(52, 152, 219), align=WD_ALIGN_PARAGRAPH.CENTER)
    
    # Add decorative line
    line_para = doc.add_paragraph()
    line_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    line_run = line_para.add_run("_________________________________________")
    line_run.font.color.rgb = RGBColor(52, 152, 219)
    line_run.font.size = Pt(14)
    
    # Add deployment image
    deployment_url = "https://assets-global.website-files.com/5ef534dbe7a6ce7aec9455a3/5f81bd0e57f8c28e87f54ccd_UX%20for%20beginners%202.png"
    deployment_path = "images/deployment.png"
    download_image(deployment_url, deployment_path)
    
    add_image_to_document(doc, deployment_path, width=Inches(6))
    
    deployment_intro = """
The deployment architecture for ShopSleek is designed to provide high availability, scalability, and security. The platform is hosted on HostingRaja, a reliable hosting provider that offers robust infrastructure and excellent support services.
"""
    add_paragraph_with_style(doc, deployment_intro)
    
    # 6.1 Hosting Infrastructure with visuals
    add_heading_with_style(doc, "6.1 Hosting Infrastructure", level=2, font_size=16, font_color=RGBColor(44, 62, 80))
    
    add_image_to_document(doc, "images/hosting_icon.png", width=Inches(1.5))
    
    hosting_info = """
ShopSleek is deployed on HostingRaja's cloud infrastructure, which provides:

• 99.9% uptime guarantee with service level agreements
• Geographically distributed data centers for improved latency and redundancy
• DDoS protection and firewall services to mitigate attacks
• Scalable resources that can be adjusted based on traffic demands
• Isolated environments for development, staging, and production

The hosting environment is configured with auto-scaling capabilities that automatically adjust resources based on current traffic patterns. During peak shopping periods (such as sales events or holidays), the system can scale horizontally by adding more server instances to handle increased load.
"""
    add_paragraph_with_style(doc, hosting_info)
    
    # 6.2 Server Configuration
    add_heading_with_style(doc, "6.2 Server Configuration", level=2, font_size=16, font_color=RGBColor(44, 62, 80))
    
    add_image_to_document(doc, "images/server_icon.png", width=Inches(1.5))
    
    server_config = """
The server environment consists of:

• Node.js application servers running Express.js framework
• Nginx as a reverse proxy and load balancer
• PM2 process manager for Node.js application management
• Redis for session management and caching
• Docker containers for consistent deployment environments

The server architecture follows a microservices approach, with separate services for:

• User authentication and management
• Product catalog and inventory
• Order processing and fulfillment
• Payment processing
• Search and recommendation engine
• Admin dashboard and reporting

Each service is deployed in its own container, allowing for independent scaling and deployment. This architecture improves fault isolation and enables more efficient resource utilization.
"""
    add_paragraph_with_style(doc, server_config)
    
    # 6.3 Database Setup
    add_heading_with_style(doc, "6.3 Database Setup", level=2, font_size=16, font_color=RGBColor(44, 62, 80))
    
    add_image_to_document(doc, "images/database_icon.png", width=Inches(1.5))
    
    database_setup = """
ShopSleek utilizes MySQL 8.0 as its primary database, configured with:

• Master-slave replication for read scalability and redundancy
• Optimized query performance through proper indexing and query planning
• Regular database maintenance including vacuum, reindexing, and statistics updates
• Separate database instances for transactional and analytical workloads
• Database connection pooling to efficiently manage connections

The database schema is designed with normalization principles to ensure data integrity while optimizing for common query patterns. Sensitive data fields are encrypted at rest, and the database servers are isolated in a private network accessible only through application servers.
"""
    add_paragraph_with_style(doc, database_setup)
    
    # 6.4 Backup and Monitoring
    add_heading_with_style(doc, "6.4 Backup and Monitoring", level=2, font_size=16, font_color=RGBColor(44, 62, 80))
    
    add_image_to_document(doc, "images/backup_icon.png", width=Inches(1.5))
    
    backup_monitoring = """
To ensure reliability and quick recovery from any issues, ShopSleek implements:

• Daily automated backups with point-in-time recovery capabilities
• Offsite backup storage with encryption for disaster recovery
• Transaction log backups every 15 minutes for minimal data loss
• Regular backup restoration tests to verify recovery procedures

The monitoring infrastructure includes:

• Real-time performance monitoring for all system components
• Automated alerting for performance degradation or anomalies
• Error logging and aggregation for quick troubleshooting
• User experience monitoring including page load times and transaction success rates
• Security monitoring for potential intrusion attempts

All logs are centralized in an ELK (Elasticsearch, Logstash, Kibana) stack for easy searching and visualization. Custom dashboards provide real-time insights into system health and performance metrics.
"""
    add_paragraph_with_style(doc, backup_monitoring)
    
    add_page_break(doc)
    
    # 7. Future Scope
    add_heading_with_style(doc, "7. Future Scope", font_size=18, font_color=RGBColor(52, 152, 219), align=WD_ALIGN_PARAGRAPH.CENTER)
    
    # Add decorative line
    line_para = doc.add_paragraph()
    line_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    line_run = line_para.add_run("_________________________________________")
    line_run.font.color.rgb = RGBColor(52, 152, 219)
    line_run.font.size = Pt(14)
    
    # Add future scope image
    future_url = "https://img.freepik.com/free-vector/futuristic-concept-illustration_52683-37769.jpg"
    future_path = "images/future.jpg"
    download_image(future_url, future_path)
    
    add_image_to_document(doc, future_path, width=Inches(4))
    
    future_scope_text = """
The ShopSleek platform has been designed with future expansion in mind. Several enhancements and new features are planned for upcoming releases:
"""
    add_paragraph_with_style(doc, future_scope_text)
    
    # Download future feature icons
    future_icons = {
        "mobile": "https://cdn-icons-png.flaticon.com/512/545/545245.png",
        "ai": "https://cdn-icons-png.flaticon.com/512/2103/2103633.png",
        "analytics": "https://cdn-icons-png.flaticon.com/512/2153/2153378.png",
        "voice": "https://cdn-icons-png.flaticon.com/512/3617/3617337.png",
        "ar": "https://cdn-icons-png.flaticon.com/512/3286/3286536.png"
    }
    
    for key, url in future_icons.items():
        icon_path = f"images/{key}_future_icon.png"
        if not os.path.exists(icon_path):
            download_image(url, icon_path)
    
    # Create future features table with visuals
    future_feature_data = [
        (
            "images/mobile_future_icon.png",
            "Mobile Applications",
            "Native Android and iOS apps that provide an optimized shopping experience for mobile users, with offline browsing and push notifications."
        ),
        (
            "images/ai_future_icon.png",
            "AI-powered Recommendations",
            "Personalized product suggestions based on browsing history, purchase patterns, and similar user behaviors."
        ),
        (
            "images/analytics_future_icon.png",
            "Advanced Analytics",
            "Customer behavior analysis and sales forecasting tools to help businesses make data-driven decisions."
        ),
        (
            "images/voice_future_icon.png",
            "Voice Commerce",
            "Integration with popular voice assistants like Amazon Alexa and Google Assistant to enable shopping through voice commands."
        ),
        (
            "images/ar_future_icon.png",
            "Augmented Reality",
            "Virtual product try-on experiences that allow customers to visualize products in their own environment before purchasing."
        )
    ]
    
    add_table_with_images(doc, "Future Feature Roadmap", 2, future_feature_data)
    
    # Add footer to all pages
    for section in doc.sections:
        footer = section.footer
        footer_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer_run = footer_para.add_run("ShopSleek E-commerce Platform | Department of Computer Science, Gujarat University")
        footer_run.font.size = Pt(8)
        footer_run.font.color.rgb = RGBColor(128, 128, 128)
    
    # Save the document
    doc.save('ShopSleek_Enhanced_Documentation_Visual.docx')
    print("Enhanced visual documentation created successfully!")

if __name__ == "__main__":
    create_enhanced_document()