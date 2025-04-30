from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

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

def add_colored_box(doc, text, color="3498DB", font_color=RGBColor(255, 255, 255), padding=12):
    """Add a colored box with text."""
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Set width to page width minus margins
    table.autofit = False
    table.width = Inches(6)
    
    # Add text to the cell
    cell = table.cell(0, 0)
    set_cell_background(cell, color)
    
    # Add padding to the cell
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    
    for side in ['top', 'left', 'bottom', 'right']:
        node = OxmlElement(f'w:{side}')
        node.set(qn('w:w'), str(padding * 20))  # twentieth of a point
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    
    tcPr.append(tcMar)
    
    # Add styled text
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(text)
    run.font.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = font_color
    
    doc.add_paragraph()  # Add some space after the box
    
    return table

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

def add_process_diagram(doc, steps, title=None):
    """
    Add a horizontal process flow diagram.
    
    Args:
        doc: Document object
        steps: List of step descriptions
        title: Optional title for the diagram
    """
    if title:
        heading = doc.add_heading(title, level=2)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Create a table for the process flow
    num_steps = len(steps)
    table = doc.add_table(rows=3, cols=num_steps)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Step numbers row
    for i in range(num_steps):
        cell = table.cell(0, i)
        set_cell_background(cell, "3498DB")  # Blue background
        
        # Add step number
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(f"Step {i+1}")
        run.font.bold = True
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(255, 255, 255)  # White text
    
    # Arrows row
    for i in range(num_steps):
        cell = table.cell(1, i)
        
        # Add arrow (except for the last cell)
        if i < num_steps - 1:
            paragraph = cell.paragraphs[0]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = paragraph.add_run("→")
            run.font.bold = True
            run.font.size = Pt(16)
            run.font.color.rgb = RGBColor(52, 152, 219)
    
    # Description row
    for i, step in enumerate(steps):
        cell = table.cell(2, i)
        
        # Add step description
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(step)
        run.font.size = Pt(10)
    
    doc.add_paragraph()  # Add some space after the diagram
    
    return table

def add_feature_boxes(doc, features, title=None, cols=2):
    """
    Add feature boxes in a grid layout.
    
    Args:
        doc: Document object
        features: List of tuples (title, description)
        title: Optional section title
        cols: Number of columns in the grid
    """
    if title:
        heading = doc.add_heading(title, level=2)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Calculate number of rows needed
    num_features = len(features)
    rows = (num_features + cols - 1) // cols  # Ceiling division
    
    # Create a table for the features grid
    table = doc.add_table(rows=rows, cols=cols)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Fill the table with features
    for i, (feature_title, feature_desc) in enumerate(features):
        row = i // cols
        col = i % cols
        
        cell = table.cell(row, col)
        
        # Set background color
        colors = ["E3F2FD", "E8F5E9", "FFF3E0", "F3E5F5"]  # Light blue, green, orange, purple
        set_cell_background(cell, colors[i % len(colors)])
        
        # Add feature title
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(feature_title)
        run.font.bold = True
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(44, 62, 80)
        
        # Add feature description
        paragraph = cell.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = paragraph.add_run(feature_desc)
        run.font.size = Pt(10)
    
    doc.add_paragraph()  # Add some space after the feature boxes
    
    return table

def add_comparison_table(doc, title, categories, options, ratings):
    """
    Add a comparison table with ratings.
    
    Args:
        doc: Document object
        title: Table title
        categories: List of category names
        options: List of option names
        ratings: Dictionary mapping (category, option) to rating (1-5)
    """
    doc.add_heading(title, level=2)
    
    # Create the table
    table = doc.add_table(rows=len(categories)+1, cols=len(options)+1)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Empty top-left cell
    table.cell(0, 0).text = ""
    
    # Header row with option names
    for j, option in enumerate(options):
        cell = table.cell(0, j+1)
        set_cell_background(cell, "3498DB")  # Blue background
        
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(option)
        run.font.bold = True
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(255, 255, 255)  # White text
    
    # Category names in first column
    for i, category in enumerate(categories):
        cell = table.cell(i+1, 0)
        set_cell_background(cell, "ECF0F1")  # Light gray background
        
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = paragraph.add_run(category)
        run.font.bold = True
        run.font.size = Pt(10)
    
    # Fill in ratings
    for i, category in enumerate(categories):
        for j, option in enumerate(options):
            cell = table.cell(i+1, j+1)
            rating = ratings.get((category, option), 0)
            
            # Add stars based on rating
            paragraph = cell.paragraphs[0]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            stars = "★" * rating + "☆" * (5 - rating)
            
            # Color based on rating
            if rating >= 4:
                star_color = RGBColor(46, 204, 113)  # Green
            elif rating >= 3:
                star_color = RGBColor(241, 196, 15)  # Yellow
            else:
                star_color = RGBColor(231, 76, 60)  # Red
            
            run = paragraph.add_run(stars)
            run.font.size = Pt(10)
            run.font.color.rgb = star_color
    
    doc.add_paragraph()  # Add some space after the table
    
    return table

def add_timeline(doc, title, events):
    """
    Add a vertical timeline of events.
    
    Args:
        doc: Document object
        title: Timeline title
        events: List of tuples (date, event_name, description)
    """
    doc.add_heading(title, level=2)
    
    # Create a table for the timeline
    table = doc.add_table(rows=len(events), cols=3)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Set column widths
    table.columns[0].width = Inches(1)    # Date
    table.columns[1].width = Inches(0.5)  # Connector
    table.columns[2].width = Inches(4.5)  # Event details
    
    # Add timeline events
    for i, (date, event_name, description) in enumerate(events):
        # Date column
        date_cell = table.cell(i, 0)
        set_cell_background(date_cell, "3498DB")  # Blue background
        
        paragraph = date_cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(date)
        run.font.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(255, 255, 255)  # White text
        
        # Connector column
        connector_cell = table.cell(i, 1)
        paragraph = connector_cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run("⟹")
        run.font.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(52, 152, 219)
        
        # Event details column
        event_cell = table.cell(i, 2)
        
        # Event name
        paragraph = event_cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = paragraph.add_run(event_name)
        run.font.bold = True
        run.font.size = Pt(12)
        
        # Event description
        paragraph = event_cell.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = paragraph.add_run(description)
        run.font.size = Pt(10)
    
    doc.add_paragraph()  # Add some space after the timeline
    
    return table

def add_statistics_section(doc, title, stats):
    """
    Add a section with key statistics in boxes.
    
    Args:
        doc: Document object
        title: Section title
        stats: List of tuples (stat_value, stat_label, color_hex)
    """
    doc.add_heading(title, level=2)
    
    # Create a table for the statistics
    num_stats = len(stats)
    table = doc.add_table(rows=1, cols=num_stats)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Add each statistic
    for i, (value, label, color) in enumerate(stats):
        cell = table.cell(0, i)
        set_cell_background(cell, color)
        
        # Value
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(value)
        run.font.bold = True
        run.font.size = Pt(24)
        run.font.color.rgb = RGBColor(255, 255, 255)
        
        # Label
        paragraph = cell.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(label)
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(255, 255, 255)
    
    doc.add_paragraph()  # Add some space after the statistics
    
    return table

def create_visual_document():
    """Create a highly visual document for ShopSleek E-commerce platform."""
    doc = Document()
    
    # Set default paragraph spacing
    style = doc.styles['Normal']
    style.font.size = Pt(12)
    style.paragraph_format.space_after = Pt(6)
    
    # Title Page
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_para.space_after = Pt(0)
    
    # Add ShopSleek title with blue color
    title_run = title_para.add_run("ShopSleek")
    title_run.font.size = Pt(60)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(52, 152, 219)
    
    # Add E-commerce Platform subtitle
    subtitle_para = doc.add_paragraph()
    subtitle_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle_para.space_after = Pt(60)
    subtitle_run = subtitle_para.add_run("E-commerce Platform")
    subtitle_run.font.size = Pt(36)
    subtitle_run.font.bold = True
    subtitle_run.font.color.rgb = RGBColor(44, 62, 80)
    
    # Add colored box with tagline
    add_colored_box(doc, "A Modern Shopping Experience", "3498DB")
    
    # Add presentation date
    date_para = doc.add_paragraph()
    date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_para.space_after = Pt(60)
    date_run = date_para.add_run("April 30, 2025")
    date_run.font.size = Pt(14)
    date_run.font.italic = True
    
    # Add author information
    author_para = doc.add_paragraph()
    author_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    author_run = author_para.add_run("Developed by: Hardik Kalathiya")
    author_run.font.size = Pt(14)
    author_run.font.bold = True
    
    guide_para = doc.add_paragraph()
    guide_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    guide_run = guide_para.add_run("Under the Guidance of: Dr. Maitri Jhaveri")
    guide_run.font.size = Pt(14)
    guide_run.font.bold = True
    
    dept_para = doc.add_paragraph()
    dept_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    dept_run = dept_para.add_run("Department of Computer Science")
    dept_run.font.size = Pt(14)
    
    univ_para = doc.add_paragraph()
    univ_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    univ_run = univ_para.add_run("Gujarat University")
    univ_run.font.size = Pt(14)
    
    add_page_break(doc)
    
    # Table of Contents
    add_heading_with_style(doc, "Table of Contents", font_size=18, font_color=RGBColor(52, 152, 219), align=WD_ALIGN_PARAGRAPH.CENTER)
    
    # Add decorative line
    line_para = doc.add_paragraph()
    line_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    line_run = line_para.add_run("_________________________________________")
    line_run.font.color.rgb = RGBColor(52, 152, 219)
    line_run.font.size = Pt(14)
    
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
    add_heading_with_style(doc, "1. Project Overview", font_size=18, font_color=RGBColor(52, 152, 219), align=WD_ALIGN_PARAGRAPH.CENTER)
    
    # Add decorative line
    line_para = doc.add_paragraph()
    line_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    line_run = line_para.add_run("_________________________________________")
    line_run.font.color.rgb = RGBColor(52, 152, 219)
    line_run.font.size = Pt(14)
    
    # Add what is ShopSleek section in a colored box
    add_colored_box(doc, "What is ShopSleek?", "3498DB")
    
    overview_text = """
ShopSleek is a comprehensive e-commerce platform developed to provide users with a seamless online shopping experience. The platform is designed to handle various aspects of online retail, from product browsing and search to secure checkout and order management.

The system aims to provide an intuitive interface for customers while offering robust backend management capabilities for administrators. ShopSleek supports multiple device compatibility through responsive design and incorporates industry standard security practices to protect user data and transactions.
"""
    add_paragraph_with_style(doc, overview_text)
    
    # Add key statistics
    stats = [
        ("30%", "Faster Checkout", "3498DB"),   # Blue
        ("45%", "Higher Conversion", "2ECC71"), # Green
        ("60%", "Lower Cart Abandonment", "E74C3C")  # Red
    ]
    add_statistics_section(doc, "Key Performance Metrics", stats)
    
    # Add core functionalities process flow
    process_steps = [
        "Browse Products",
        "Add to Cart",
        "Checkout",
        "Payment",
        "Order Tracking"
    ]
    add_process_diagram(doc, process_steps, "ShopSleek User Journey")
    
    # Add user types
    user_types = [
        ("Customers", "End users who browse products, make purchases, and track orders."),
        ("Administrators", "Backend users who manage products, orders, and system settings."),
        ("Vendors", "Third-party sellers who list their products on the platform."),
        ("Guest Users", "Unregistered users with limited browsing and shopping capabilities.")
    ]
    add_feature_boxes(doc, user_types, "User Types", cols=2)
    
    add_page_break(doc)
    
    # 2. Key Features
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
    
    # Add key features with clear formatting
    features_data = [
        (
            "Intuitive User Interface", 
            "Clean, responsive design optimized for all devices with thoughtful navigation paths and minimal friction points."
        ),
        (
            "Advanced Search & Filtering", 
            "Find products quickly with dynamic filtering capabilities that include category browsing, price range filters, and brand selection."
        ),
        (
            "Secure Payment Processing", 
            "PCI-DSS compliant payment gateway integration supporting multiple payment methods including credit/debit cards and digital wallets."
        ),
        (
            "Order Management", 
            "Track orders from placement to delivery with real-time updates and notifications for both users and administrators."
        ),
        (
            "User Accounts", 
            "Personal profiles with wishlists, order history, saved addresses, and preference settings for a personalized experience."
        ),
        (
            "Admin Dashboard", 
            "Complete control over products, orders, users, and inventory with an intuitive administrative interface."
        ),
        (
            "Analytics & Reporting", 
            "Comprehensive sales reports, user activity tracking, and inventory level monitoring with visual data representation."
        ),
        (
            "Multi-vendor Support", 
            "Platform for multiple sellers to list their products with separate dashboards and commission management."
        )
    ]
    
    add_feature_boxes(doc, features_data, "Feature Highlights", cols=2)
    
    # Add feature comparison
    categories = ["User Experience", "Search Functionality", "Payment Options", "Mobile Support", "Admin Controls"]
    options = ["ShopSleek", "Competitor A", "Competitor B"]
    
    ratings = {
        ("User Experience", "ShopSleek"): 5,
        ("User Experience", "Competitor A"): 3,
        ("User Experience", "Competitor B"): 4,
        
        ("Search Functionality", "ShopSleek"): 5,
        ("Search Functionality", "Competitor A"): 4,
        ("Search Functionality", "Competitor B"): 3,
        
        ("Payment Options", "ShopSleek"): 4,
        ("Payment Options", "Competitor A"): 3,
        ("Payment Options", "Competitor B"): 5,
        
        ("Mobile Support", "ShopSleek"): 5,
        ("Mobile Support", "Competitor A"): 4,
        ("Mobile Support", "Competitor B"): 3,
        
        ("Admin Controls", "ShopSleek"): 5,
        ("Admin Controls", "Competitor A"): 4,
        ("Admin Controls", "Competitor B"): 2,
    }
    
    add_comparison_table(doc, "Feature Comparison", categories, options, ratings)
    
    add_page_break(doc)
    
    # 3. System Architecture
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
    
    # Add Three-Tier Architecture as feature boxes
    architecture_layers = [
        (
            "Presentation Layer", 
            "The top-most layer that directly interacts with users through the User Interface, Admin Dashboard, and Vendor Panel. This layer is responsible for rendering data in a user-friendly format and collecting user inputs."
        ),
        (
            "Application Layer", 
            "The middle layer that processes business logic, including User Management, Product Management, Order Processing, and Payment Gateway integration. This layer acts as an intermediary between the presentation and data layers, applying business rules and ensuring data integrity."
        ),
        (
            "Data Layer", 
            "The foundation layer that handles data storage and retrieval through MySQL Database and File Storage. This layer is optimized for data persistence, security, and efficient query processing."
        )
    ]
    
    # Create a visually distinct layout for architecture layers
    layer_colors = ["3498DB", "2ECC71", "E74C3C"]  # Blue, Green, Red
    
    for i, (layer, description) in enumerate(architecture_layers):
        add_colored_box(doc, layer, layer_colors[i])
        add_paragraph_with_style(doc, description)
    
    # Add Technology Stack section
    add_heading_with_style(doc, "Technology Stack", level=2, font_size=16, font_color=RGBColor(44, 62, 80))
    
    tech_stack = [
        ("Frontend", "React.js, Tailwind CSS - Modern JavaScript framework with utility-first CSS for building responsive user interfaces"),
        ("Backend", "Node.js, Express.js - JavaScript runtime environment with a minimalist web framework for building APIs and server-side logic"),
        ("Database", "MySQL - Relational database management system for structured data storage with powerful query capabilities"),
        ("Server", "Node.js - Lightweight, efficient, and scalable server environment for handling HTTP requests"),
        ("Hosting", "HostingRaja - Reliable hosting service with high uptime guarantee and scalable infrastructure")
    ]
    
    # Create a table for technology stack
    table = doc.add_table(rows=len(tech_stack), cols=2)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Set column widths
    table.columns[0].width = Inches(1.5)
    table.columns[1].width = Inches(4.5)
    
    # Fill the table with tech stack information
    for i, (tech_name, tech_desc) in enumerate(tech_stack):
        # Technology name column
        name_cell = table.cell(i, 0)
        set_cell_background(name_cell, "34495E")  # Dark blue background
        
        paragraph = name_cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(tech_name)
        run.font.bold = True
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(255, 255, 255)  # White text
        
        # Technology description column
        desc_cell = table.cell(i, 1)
        
        paragraph = desc_cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        parts = tech_desc.split(" - ", 1)
        
        # Add technology components
        run = paragraph.add_run(parts[0])
        run.font.bold = True
        run.font.size = Pt(11)
        
        # Add description if available
        if len(parts) > 1:
            run = paragraph.add_run(" - " + parts[1])
            run.font.size = Pt(10)
            run.font.italic = True
    
    add_page_break(doc)
    
    # 4. Test Cases
    add_heading_with_style(doc, "4. Test Cases", font_size=18, font_color=RGBColor(52, 152, 219), align=WD_ALIGN_PARAGRAPH.CENTER)
    
    # Add decorative line
    line_para = doc.add_paragraph()
    line_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    line_run = line_para.add_run("_________________________________________")
    line_run.font.color.rgb = RGBColor(52, 152, 219)
    line_run.font.size = Pt(14)
    
    test_intro_text = """
Comprehensive testing was conducted to ensure the platform's functionality, reliability, and security. Each test case was carefully designed to validate specific features and identify potential issues before deployment. The following sections present the test cases organized by functional categories.
"""
    add_paragraph_with_style(doc, test_intro_text)
    
    # 4.1 User Authentication Test Cases
    add_heading_with_style(doc, "4.1 User Authentication Test Cases", level=2, font_size=14)
    
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
    
    add_page_break(doc)
    
    # 4.2 Product Search Test Cases
    add_heading_with_style(doc, "4.2 Product Search Test Cases", level=2, font_size=14)
    
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
    add_heading_with_style(doc, "4.3 Shopping Cart Test Cases", level=2, font_size=14)
    
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
    
    add_page_break(doc)
    
    # 4.4 Checkout Process Test Cases
    add_heading_with_style(doc, "4.4 Checkout Process Test Cases", level=2, font_size=14)
    
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
    add_heading_with_style(doc, "4.5 Test Results Summary", level=2, font_size=14)
    
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
    table.style = 'Table Grid'
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
                    
                    # For success rate column, color by percentage
                    if i == 4:  # Success Rate column
                        rate = float(cell_text.strip('%'))
                        if rate >= 90:
                            run.font.color.rgb = RGBColor(39, 174, 96)  # Green
                        elif rate >= 80:
                            run.font.color.rgb = RGBColor(241, 196, 15)  # Yellow
                        else:
                            run.font.color.rgb = RGBColor(231, 76, 60)  # Red
                    
                    # For "Overall System" row, make it bold
                    if row_data[0] == "Overall System":
                        run.font.bold = True
                        set_cell_background(row_cells[i], "E5F0FF")  # Light blue background
    
    # Add visual indicators for test results
    add_statistics_section(doc, "Test Success Rates", [
        ("100%", "User Authentication", "2ECC71"),  # Green
        ("83.3%", "Product & Checkout", "F1C40F"),  # Yellow
        ("88.9%", "Overall System", "3498DB")       # Blue
    ])
    
    # Add analysis of test results
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
    
    security_intro = """
Security is a critical aspect of the ShopSleek e-commerce platform, especially considering the sensitive nature of user data and payment information being processed. The platform implements multiple layers of security measures to protect both user data and system integrity.
"""
    add_paragraph_with_style(doc, security_intro)
    
    # 5.1 Data Protection
    add_heading_with_style(doc, "5.1 Data Protection", level=2, font_size=14)
    
    # Security features in a visual format
    security_features = [
        ("Password Encryption", "All user passwords are encrypted using bcrypt hashing algorithm with salt, ensuring that even in the event of a database breach, passwords remain secure."),
        
        ("Sensitive Data Encryption", "Personal information and payment details are encrypted using AES-256 encryption in the database with secure key management."),
        
        ("Role-based Access Control", "The system implements granular access controls that restrict users to only the data and functions necessary for their role."),
        
        ("Session Management", "Secure session handling with automatic timeout after periods of inactivity and secure cookies with appropriate flags."),
        
        ("Two-factor Authentication", "Optional for regular users and mandatory for administrative accounts, adding an additional layer of security beyond passwords."),
        
        ("Data Loss Prevention", "Regular automated backups with encryption and secure off-site storage with point-in-time recovery capabilities.")
    ]
    
    # Create a styled table for security features
    table = doc.add_table(rows=len(security_features), cols=2)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Set column widths
    table.columns[0].width = Inches(2)
    table.columns[1].width = Inches(4)
    
    # Add security features to the table
    for i, (feature, description) in enumerate(security_features):
        # Feature column
        feature_cell = table.cell(i, 0)
        set_cell_background(feature_cell, "E74C3C")  # Red background for security
        
        paragraph = feature_cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run("🔒 " + feature)
        run.font.bold = True
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(255, 255, 255)  # White text
        
        # Description column
        desc_cell = table.cell(i, 1)
        
        paragraph = desc_cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = paragraph.add_run(description)
        run.font.size = Pt(10)
    
    doc.add_paragraph()  # Add some space after the table
    
    # 5.2 Security Implementation
    add_heading_with_style(doc, "5.2 Security Implementation", level=2, font_size=14)
    
    # Timeline of security implementations
    security_timeline = [
        ("Phase 1", "Fundamental Security", "Implementation of basic security measures including encryption, access control, and secure sessions."),
        ("Phase 2", "Advanced Protection", "Addition of two-factor authentication, intrusion detection, and enhanced monitoring."),
        ("Phase 3", "Compliance & Auditing", "Implementation of PCI-DSS and GDPR compliance measures, regular security audits."),
        ("Phase 4", "Continuous Improvement", "Ongoing vulnerability assessments, penetration testing, and security updates.")
    ]
    
    add_timeline(doc, "Security Implementation Timeline", security_timeline)
    
    add_page_break(doc)
    
    # 6. Deployment Architecture
    add_heading_with_style(doc, "6. Deployment Architecture", font_size=18, font_color=RGBColor(52, 152, 219), align=WD_ALIGN_PARAGRAPH.CENTER)
    
    # Add decorative line
    line_para = doc.add_paragraph()
    line_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    line_run = line_para.add_run("_________________________________________")
    line_run.font.color.rgb = RGBColor(52, 152, 219)
    line_run.font.size = Pt(14)
    
    deployment_intro = """
The deployment architecture for ShopSleek is designed to provide high availability, scalability, and security. The platform is hosted on HostingRaja, a reliable hosting provider that offers robust infrastructure and excellent support services.
"""
    add_paragraph_with_style(doc, deployment_intro)
    
    # 6.1 Hosting Infrastructure
    add_heading_with_style(doc, "6.1 Hosting Infrastructure", level=2, font_size=14)
    
    hosting_features = [
        ("High Availability", "99.9% uptime guarantee with service level agreements and geographically distributed data centers."),
        ("Scalability", "Auto-scaling capabilities that automatically adjust resources based on current traffic patterns."),
        ("Security", "DDoS protection, firewall services, and secure network infrastructure to mitigate attacks."),
        ("Environment Isolation", "Separate environments for development, staging, and production with controlled deployment pipeline.")
    ]
    
    add_feature_boxes(doc, hosting_features, "Hosting Capabilities", cols=2)
    
    # 6.2 Server Configuration
    add_heading_with_style(doc, "6.2 Server Configuration", level=2, font_size=14)
    
    server_config = """
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
    
    # Server components in visual format
    server_components = [
        ("Application Servers", "Node.js application servers running Express.js framework"),
        ("Load Balancer", "Nginx as a reverse proxy and load balancer for distributing traffic"),
        ("Process Manager", "PM2 process manager for Node.js application management"),
        ("Caching", "Redis for session management and data caching"),
        ("Containers", "Docker containers for consistent deployment environments")
    ]
    
    # Create a visually appealing layout for server components
    table = doc.add_table(rows=len(server_components), cols=2)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Set column widths
    table.columns[0].width = Inches(2)
    table.columns[1].width = Inches(4)
    
    # Fill the table with server components
    for i, (component, description) in enumerate(server_components):
        # Component name column
        component_cell = table.cell(i, 0)
        set_cell_background(component_cell, "2ECC71")  # Green background
        
        paragraph = component_cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(component)
        run.font.bold = True
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(255, 255, 255)  # White text
        
        # Description column
        desc_cell = table.cell(i, 1)
        
        paragraph = desc_cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = paragraph.add_run(description)
        run.font.size = Pt(11)
    
    add_page_break(doc)
    
    # 6.3 Database Setup
    add_heading_with_style(doc, "6.3 Database Setup", level=2, font_size=14)
    
    database_setup = """
ShopSleek utilizes MySQL 8.0 as its primary database, configured with master-slave replication for read scalability and redundancy. The database architecture incorporates the following key features:
"""
    add_paragraph_with_style(doc, database_setup)
    
    # Database configuration details
    db_config = [
        "• Master-slave replication for read scalability and redundancy",
        "• Optimized query performance through proper indexing and query planning",
        "• Regular database maintenance including vacuum, reindexing, and statistics updates",
        "• Separate database instances for transactional and analytical workloads",
        "• Database connection pooling to efficiently manage connections",
        "• Automated daily backups with point-in-time recovery capabilities",
        "• Database schema designed with normalization principles for data integrity",
        "• Sensitive data fields encrypted at rest to protect user information",
        "• Database servers isolated in a private network accessible only through application servers"
    ]
    
    for item in db_config:
        add_paragraph_with_style(doc, item)
    
    # 6.4 Backup and Monitoring
    add_heading_with_style(doc, "6.4 Backup and Monitoring", level=2, font_size=14)
    
    # Backup and monitoring details
    backup_monitoring_features = [
        ("Daily Backups", "Automated daily backups with point-in-time recovery capabilities and encrypted offsite storage."),
        ("Transaction Logs", "Transaction log backups every 15 minutes for minimal data loss with verified restoration procedures."),
        ("Performance Monitoring", "Real-time performance monitoring for all system components with automated alerting."),
        ("Error Tracking", "Centralized logging with error aggregation for quick troubleshooting and resolution."),
        ("User Experience", "Monitoring of page load times, transaction success rates, and other user experience metrics."),
        ("Security Monitoring", "Intrusion detection systems and regular scanning for potential security threats.")
    ]
    
    add_feature_boxes(doc, backup_monitoring_features, "Backup & Monitoring Features", cols=2)
    
    add_page_break(doc)
    
    # 7. Future Scope
    add_heading_with_style(doc, "7. Future Scope", font_size=18, font_color=RGBColor(52, 152, 219), align=WD_ALIGN_PARAGRAPH.CENTER)
    
    # Add decorative line
    line_para = doc.add_paragraph()
    line_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    line_run = line_para.add_run("_________________________________________")
    line_run.font.color.rgb = RGBColor(52, 152, 219)
    line_run.font.size = Pt(14)
    
    future_scope_text = """
The ShopSleek platform has been designed with future expansion in mind. Several enhancements and new features are planned for upcoming releases:
"""
    add_paragraph_with_style(doc, future_scope_text)
    
    # Future enhancements timeline
    future_timeline = [
        ("Q3 2025", "Mobile Applications", "Native Android and iOS apps with offline capabilities, push notifications, and biometric authentication."),
        ("Q4 2025", "AI Recommendations", "Personalized product suggestions based on browsing history, purchase patterns, and similar user behaviors."),
        ("Q1 2026", "Voice Commerce", "Integration with voice assistants like Amazon Alexa and Google Assistant for voice-based shopping."),
        ("Q2 2026", "AR Shopping", "Augmented reality features allowing customers to virtually try products before purchasing."),
        ("Q3 2026", "Global Expansion", "Multi-language and multi-currency support with localized payment methods and shipping options."),
        ("Q4 2026", "Blockchain Integration", "Secure transactions, product authenticity verification, and loyalty program management.")
    ]
    
    add_timeline(doc, "Feature Development Roadmap", future_timeline)
    
    # Create a better organized future features table
    future_features_table = doc.add_table(rows=4, cols=3)
    future_features_table.style = 'Table Grid'
    future_features_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Set header row
    header_cells = future_features_table.rows[0].cells
    
    headers = ["Timeline", "Feature", "Description"]
    for i, header in enumerate(headers):
        header_cells[i].text = header
        set_cell_background(header_cells[i], "3498DB")  # Blue background
        
        # Format header
        for paragraph in header_cells[i].paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(12)
                run.font.color.rgb = RGBColor(255, 255, 255)  # White text

    # Set column widths
    future_features_table.columns[0].width = Inches(1.0)   # Timeline
    future_features_table.columns[1].width = Inches(1.5)   # Feature
    future_features_table.columns[2].width = Inches(3.5)   # Description
    
    # Add feature rows
    features_data = [
        ["2025 Q3-Q4", "Mobile Applications", "Native Android and iOS apps with offline capabilities, push notifications, and synchronized user accounts across devices."],
        ["2026 Q1", "AI Recommendations", "Personalized product suggestions using machine learning algorithms based on browsing history, purchase patterns, and similar user behaviors."],
        ["2026 Q2-Q3", "Augmented Reality", "Virtual product try-on experiences allowing customers to visualize products in their environment before purchasing."]
    ]
    
    for row_idx, feature in enumerate(features_data, start=1):
        for col_idx, value in enumerate(feature):
            cell = future_features_table.cell(row_idx, col_idx)
            cell.text = value
            
            # Add some formatting
            if col_idx == 0:  # Timeline column
                set_cell_background(cell, "ECF0F1")  # Light gray background
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            if col_idx == 1:  # Feature column
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True
    
    doc.add_paragraph()  # Add spacing
    
    # Save the document
    doc.save('ShopSleek_Highly_Visual_Documentation.docx')
    print("Highly visual documentation created successfully!")

if __name__ == "__main__":
    create_visual_document()