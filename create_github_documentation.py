from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
from PIL import Image, ImageDraw, ImageFont
import io

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
        paragraph = doc.add_paragraph(item, style='List Bullet')
        paragraph.paragraph_format.space_after = Pt(6)
        
        # Format the text
        for run in paragraph.runs:
            run.font.size = Pt(font_size)
    
    return doc

def create_chart_image(title, data, filename, width=800, height=500):
    """Create a chart image and save it to a file."""
    # Create blank image
    img = Image.new('RGB', (width, height), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)
    
    # Try to load a font, fallback to default if not available
    try:
        title_font = ImageFont.truetype("arial.ttf", 24)
        label_font = ImageFont.truetype("arial.ttf", 16)
        value_font = ImageFont.truetype("arial.ttf", 14)
    except IOError:
        title_font = ImageFont.load_default()
        label_font = ImageFont.load_default()
        value_font = ImageFont.load_default()
    
    # Draw title
    draw.text((width//2, 30), title, font=title_font, fill=(44, 62, 80), anchor="mm")
    
    # Bar chart for test cases
    if "Test" in title:
        margin = 80
        chart_width = width - 2*margin
        chart_height = 300
        bar_width = chart_width // (len(data) + 2)  # +2 for spacing
        max_value = 100  # percentage
        
        # Draw Y-axis
        draw.line([(margin, margin), (margin, margin+chart_height)], fill=(44, 62, 80), width=2)
        
        # Draw X-axis
        draw.line([(margin, margin+chart_height), (margin+chart_width, margin+chart_height)], fill=(44, 62, 80), width=2)
        
        # Draw Y-axis labels and grid lines
        for i in range(6):  # 0%, 20%, 40%, 60%, 80%, 100%
            y = margin + chart_height - (i * chart_height // 5)
            value = i * 20
            draw.text((margin-10, y), f"{value}%", font=value_font, fill=(44, 62, 80), anchor="rm")
            draw.line([(margin, y), (margin+chart_width, y)], fill=(200, 200, 200), width=1)
        
        # Draw bars
        for i, (category, success_rate) in enumerate(data):
            x = margin + (i+1) * bar_width
            success_rate = float(success_rate.strip('%'))
            bar_height = (success_rate / 100) * chart_height
            
            # Choose color based on success rate
            if success_rate >= 90:
                color = (46, 204, 113)  # Green
            elif success_rate >= 80:
                color = (241, 196, 15)  # Yellow
            else:
                color = (231, 76, 60)  # Red
            
            # Draw bar
            draw.rectangle(
                [(x, margin+chart_height-bar_height), (x+bar_width*0.8, margin+chart_height)], 
                fill=color
            )
            
            # Draw category label
            draw.text(
                (x + bar_width*0.4, margin+chart_height+15), 
                category if len(category) < 15 else category[:12] + "...", 
                font=label_font, 
                fill=(44, 62, 80), 
                anchor="mt"
            )
            
            # Draw success rate on top of bar
            draw.text(
                (x + bar_width*0.4, margin+chart_height-bar_height-10), 
                f"{success_rate}%", 
                font=value_font, 
                fill=(44, 62, 80), 
                anchor="mb"
            )
    
    # Save the image
    if not os.path.exists('images'):
        os.makedirs('images')
    
    img.save(f"images/{filename}")
    return f"images/{filename}"

def create_architecture_image(filename, width=800, height=600):
    """Create a simple architecture diagram image."""
    img = Image.new('RGB', (width, height), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)
    
    # Try to load a font, fallback to default if not available
    try:
        title_font = ImageFont.truetype("arial.ttf", 24)
        header_font = ImageFont.truetype("arial.ttf", 20)
        label_font = ImageFont.truetype("arial.ttf", 16)
        desc_font = ImageFont.truetype("arial.ttf", 14)
    except IOError:
        title_font = ImageFont.load_default()
        header_font = ImageFont.load_default()
        label_font = ImageFont.load_default()
        desc_font = ImageFont.load_default()
    
    # Draw title
    draw.text((width//2, 30), "ShopSleek Architecture", font=title_font, fill=(44, 62, 80), anchor="mm")
    
    # Draw three-tier architecture
    layer_height = 120
    layer_width = 600
    start_y = 100
    margin_x = (width - layer_width) // 2
    
    # Layer 1: Presentation Layer
    draw.rectangle(
        [(margin_x, start_y), (margin_x+layer_width, start_y+layer_height)],
        fill=(52, 152, 219),  # Blue
        outline=(41, 128, 185),
        width=2
    )
    draw.text(
        (width//2, start_y+30), 
        "Presentation Layer", 
        font=header_font, 
        fill=(255, 255, 255), 
        anchor="mm"
    )
    draw.text(
        (width//2, start_y+70), 
        "User Interface | Admin Dashboard | Vendor Portal", 
        font=label_font, 
        fill=(255, 255, 255), 
        anchor="mm"
    )
    
    # Layer 2: Application Layer
    start_y += layer_height + 40
    draw.rectangle(
        [(margin_x, start_y), (margin_x+layer_width, start_y+layer_height)],
        fill=(46, 204, 113),  # Green
        outline=(39, 174, 96),
        width=2
    )
    draw.text(
        (width//2, start_y+30), 
        "Application Layer", 
        font=header_font, 
        fill=(255, 255, 255), 
        anchor="mm"
    )
    draw.text(
        (width//2, start_y+70), 
        "Business Logic | User Management | Order Processing | Payment Gateway", 
        font=label_font, 
        fill=(255, 255, 255), 
        anchor="mm"
    )
    
    # Layer 3: Data Layer
    start_y += layer_height + 40
    draw.rectangle(
        [(margin_x, start_y), (margin_x+layer_width, start_y+layer_height)],
        fill=(231, 76, 60),  # Red
        outline=(192, 57, 43),
        width=2
    )
    draw.text(
        (width//2, start_y+30), 
        "Data Layer", 
        font=header_font, 
        fill=(255, 255, 255), 
        anchor="mm"
    )
    draw.text(
        (width//2, start_y+70), 
        "MySQL Database | File Storage | Data Access Layer", 
        font=label_font, 
        fill=(255, 255, 255), 
        anchor="mm"
    )
    
    # Draw connecting arrows between layers
    arrow_x = width // 2
    
    # Arrow 1
    arrow_y1 = start_y - 40
    arrow_y2 = start_y - layer_height - 40
    draw.line([(arrow_x, arrow_y1), (arrow_x, arrow_y2)], fill=(44, 62, 80), width=3)
    draw.polygon([(arrow_x-10, arrow_y2+10), (arrow_x+10, arrow_y2+10), (arrow_x, arrow_y2)], fill=(44, 62, 80))
    
    # Arrow 2
    arrow_y1 = start_y - layer_height - 40 - 40
    arrow_y2 = 100 + layer_height
    draw.line([(arrow_x, arrow_y1), (arrow_x, arrow_y2)], fill=(44, 62, 80), width=3)
    draw.polygon([(arrow_x-10, arrow_y2-10), (arrow_x+10, arrow_y2-10), (arrow_x, arrow_y2)], fill=(44, 62, 80))
    
    # Save the image
    if not os.path.exists('images'):
        os.makedirs('images')
    
    img.save(f"images/{filename}")
    return f"images/{filename}"

def create_security_image(filename, width=800, height=500):
    """Create a security diagram image."""
    img = Image.new('RGB', (width, height), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)
    
    # Try to load a font, fallback to default if not available
    try:
        title_font = ImageFont.truetype("arial.ttf", 24)
        header_font = ImageFont.truetype("arial.ttf", 18)
        label_font = ImageFont.truetype("arial.ttf", 16)
        desc_font = ImageFont.truetype("arial.ttf", 14)
    except IOError:
        title_font = ImageFont.load_default()
        header_font = ImageFont.load_default()
        label_font = ImageFont.load_default()
        desc_font = ImageFont.load_default()
    
    # Draw title
    draw.text((width//2, 30), "ShopSleek Security Framework", font=title_font, fill=(44, 62, 80), anchor="mm")
    
    # Draw central shield
    center_x, center_y = width//2, height//2
    shield_width, shield_height = 120, 160
    
    # Shield outline
    shield_points = [
        (center_x, center_y - shield_height//2),  # Top
        (center_x + shield_width//2, center_y - shield_height//4),  # Top right
        (center_x + shield_width//2, center_y + shield_height//4),  # Bottom right
        (center_x, center_y + shield_height//2),  # Bottom
        (center_x - shield_width//2, center_y + shield_height//4),  # Bottom left
        (center_x - shield_width//2, center_y - shield_height//4),  # Top left
    ]
    draw.polygon(shield_points, fill=(52, 152, 219, 70), outline=(41, 128, 185), width=3)
    
    # Shield inner
    inner_shield_points = [
        (center_x, center_y - shield_height//2 + 20),  # Top
        (center_x + shield_width//2 - 20, center_y - shield_height//4 + 10),  # Top right
        (center_x + shield_width//2 - 20, center_y + shield_height//4 - 10),  # Bottom right
        (center_x, center_y + shield_height//2 - 20),  # Bottom
        (center_x - shield_width//2 + 20, center_y + shield_height//4 - 10),  # Bottom left
        (center_x - shield_width//2 + 20, center_y - shield_height//4 + 10),  # Top left
    ]
    draw.polygon(inner_shield_points, fill=(52, 152, 219, 120), outline=(41, 128, 185), width=2)
    
    # Shield text
    draw.text((center_x, center_y), "Core", font=header_font, fill=(44, 62, 80), anchor="mm")
    draw.text((center_x, center_y+25), "Security", font=header_font, fill=(44, 62, 80), anchor="mm")
    
    # Draw security elements around the shield
    radius = 200
    elements = [
        {"name": "Authentication", "angle": 45, "color": (52, 152, 219)},  # Blue
        {"name": "Encryption", "angle": 135, "color": (231, 76, 60)},  # Red
        {"name": "Access Control", "angle": 225, "color": (46, 204, 113)},  # Green
        {"name": "Network Security", "angle": 315, "color": (243, 156, 18)}  # Orange
    ]
    
    import math
    for element in elements:
        angle_rad = math.radians(element["angle"])
        x = center_x + radius * math.cos(angle_rad)
        y = center_y + radius * math.sin(angle_rad)
        
        # Draw circle
        circle_size = 60
        draw.ellipse(
            [(x-circle_size//2, y-circle_size//2), (x+circle_size//2, y+circle_size//2)],
            fill=element["color"] + (100,),  # Add alpha
            outline=element["color"],
            width=2
        )
        
        # Draw text
        draw.text((x, y), element["name"], font=label_font, fill=(44, 62, 80), anchor="mm")
        
        # Draw connecting line to shield
        line_end_x = center_x + (shield_width//2) * 0.8 * math.cos(angle_rad)
        line_end_y = center_y + (shield_height//2) * 0.8 * math.sin(angle_rad)
        draw.line([(x, y), (line_end_x, line_end_y)], fill=element["color"], width=2)
    
    # Save the image
    if not os.path.exists('images'):
        os.makedirs('images')
    
    img.save(f"images/{filename}")
    return f"images/{filename}"

def create_diagram_images():
    """Create all diagram images needed for the documentation."""
    # Test Results chart
    test_data = [
        ("User Auth", "100%"),
        ("Products", "83.3%"),
        ("Checkout", "83.3%"),
        ("Admin", "87.5%"),
        ("API", "90%"),
        ("Overall", "88.9%")
    ]
    test_chart_path = create_chart_image("Test Results Summary", test_data, "test_results_chart.png")
    
    # Architecture diagram
    architecture_path = create_architecture_image("architecture_diagram.png")
    
    # Security diagram
    security_path = create_security_image("security_diagram.png")
    
    return {
        "test_chart": test_chart_path,
        "architecture": architecture_path,
        "security": security_path
    }

def add_image_with_caption(doc, image_path, width=None, caption=None, centered=True):
    """Add an image with an optional caption."""
    if os.path.exists(image_path):
        paragraph = doc.add_paragraph()
        if centered:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        run = paragraph.add_run()
        if width:
            run.add_picture(image_path, width=Inches(width))
        else:
            run.add_picture(image_path)
        
        if caption:
            caption_paragraph = doc.add_paragraph()
            caption_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            caption_paragraph.style = 'Caption'
            caption_run = caption_paragraph.add_run(caption)
            caption_run.font.italic = True
            caption_run.font.size = Pt(11)
    else:
        paragraph = doc.add_paragraph(f"[Image not found: {image_path}]")
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    return doc

def create_features_with_icons(doc, features_data):
    """
    Create a section with features and their icons.
    
    Args:
        doc: Document object
        features_data: List of tuples (icon_path, title, description)
    """
    # Create a table for features
    table = doc.add_table(rows=len(features_data), cols=2)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Set column widths
    table.columns[0].width = Inches(1.5)  # Icon column
    table.columns[1].width = Inches(5.0)  # Description column
    
    # Add each feature
    for i, (icon_path, title, description) in enumerate(features_data):
        # Icon cell
        icon_cell = table.cell(i, 0)
        icon_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        if os.path.exists(icon_path):
            icon_run = icon_cell.paragraphs[0].add_run()
            icon_run.add_picture(icon_path, width=Inches(1.0))
        
        # Description cell
        desc_cell = table.cell(i, 1)
        
        # Add title
        title_paragraph = desc_cell.paragraphs[0]
        title_run = title_paragraph.add_run(title)
        title_run.font.bold = True
        title_run.font.size = Pt(14)
        title_run.font.color.rgb = RGBColor(52, 152, 219)  # Blue
        
        # Add description
        desc_paragraph = desc_cell.add_paragraph()
        desc_run = desc_paragraph.add_run(description)
        desc_run.font.size = Pt(12)
    
    return doc

def create_document():
    """Create a comprehensive document for ShopSleek E-commerce platform with GitHub-compatible formatting."""
    doc = Document()
    
    # Create diagram images
    diagrams = create_diagram_images()
    
    # Title Page
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run("ShopSleek")
    title_run.font.size = Pt(36)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(52, 152, 219)  # Blue
    
    subtitle_para = doc.add_paragraph()
    subtitle_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle_run = subtitle_para.add_run("E-commerce Platform Documentation")
    subtitle_run.font.size = Pt(24)
    subtitle_run.font.italic = True
    
    doc.add_paragraph()  # Spacing
    
    # Date
    date_para = doc.add_paragraph()
    date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_run = date_para.add_run("April 30, 2025")
    date_run.font.size = Pt(14)
    
    # Author
    author_para = doc.add_paragraph()
    author_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    author_run = author_para.add_run("Developed by: Hardik Kalathiya")
    author_run.font.size = Pt(14)
    author_run.font.bold = True
    
    # Guide
    guide_para = doc.add_paragraph()
    guide_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    guide_run = guide_para.add_run("Under the Guidance of: Dr. Maitri Jhaveri")
    guide_run.font.size = Pt(14)
    
    # Department and University
    dept_para = doc.add_paragraph()
    dept_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    dept_run = dept_para.add_run("Department of Computer Science\nGujarat University")
    dept_run.font.size = Pt(14)
    
    add_page_break(doc)
    
    # Table of Contents
    add_heading_with_style(doc, "Table of Contents", font_size=16, font_color=RGBColor(52, 152, 219))
    
    toc_items = [
        "1. Introduction",
        "2. Key Features",
        "3. System Architecture",
        "4. Test Cases",
        "5. Security Aspects",
        "6. Deployment Architecture",
        "7. Future Scope"
    ]
    
    for item in toc_items:
        para = doc.add_paragraph()
        para.paragraph_format.left_indent = Pt(24)
        para.paragraph_format.space_after = Pt(8)
        run = para.add_run(item)
        run.font.size = Pt(12)
        
    add_page_break(doc)
    
    # 1. Introduction
    add_heading_with_style(doc, "1. Introduction", font_size=16, font_color=RGBColor(52, 152, 219))
    
    intro_text = """
ShopSleek is a comprehensive e-commerce platform developed to provide users with a seamless online shopping experience. The platform aims to handle all aspects of online retail, from product browsing and search to secure checkout and order management.

Key highlights of the ShopSleek platform include:
"""
    add_paragraph_with_style(doc, intro_text)
    
    intro_bullets = [
        "User-friendly interface with responsive design for all devices",
        "Robust backend admin portal for inventory and order management",
        "Secure payment processing with multiple payment options",
        "Advanced search and filtering capabilities",
        "Detailed analytics and reporting features",
        "Multi-vendor support with commission management"
    ]
    
    add_bulleted_list(doc, intro_bullets)
    
    add_paragraph_with_style(doc, "\nThe platform is built with a focus on performance, security, and scalability, making it suitable for businesses of all sizes.")
    
    add_page_break(doc)
    
    # 2. Key Features
    add_heading_with_style(doc, "2. Key Features", font_size=16, font_color=RGBColor(52, 152, 219))
    
    features_text = """
ShopSleek offers a comprehensive set of features designed to enhance both the customer shopping experience and administrative capabilities. The following are key features that set ShopSleek apart from other e-commerce solutions:
"""
    add_paragraph_with_style(doc, features_text)
    
    # Features with icons (simulated with basic shapes in generated images)
    # In a real implementation, you would use actual icons
    os.makedirs('images/icons', exist_ok=True)
    
    # Create simple icon images
    icon_size = (100, 100)
    icon_bg_colors = [
        (52, 152, 219),   # Blue
        (46, 204, 113),   # Green
        (231, 76, 60),    # Red
        (243, 156, 18),   # Orange
        (155, 89, 182),   # Purple
        (52, 73, 94)      # Dark
    ]
    
    icon_names = ["ui", "search", "payment", "orders", "users", "admin"]
    icon_paths = []
    
    for i, name in enumerate(icon_names):
        icon_img = Image.new('RGB', icon_size, color=(255, 255, 255))
        draw = ImageDraw.Draw(icon_img)
        
        # Draw colored circle background
        margin = 10
        draw.ellipse(
            [(margin, margin), (icon_size[0]-margin, icon_size[1]-margin)],
            fill=icon_bg_colors[i] + (150,),
            outline=icon_bg_colors[i],
            width=2
        )
        
        # Save icon
        icon_path = f"images/icons/{name}_icon.png"
        icon_img.save(icon_path)
        icon_paths.append(icon_path)
    
    # Feature data with icons
    features_data = [
        (icon_paths[0], "Intuitive User Interface", 
         "Clean, responsive design optimized for all devices with thoughtful navigation paths and minimal friction points."),
        
        (icon_paths[1], "Advanced Search & Filtering", 
         "Find products quickly with dynamic filtering capabilities that include category browsing, price range filters, and brand selection."),
        
        (icon_paths[2], "Secure Payment Processing", 
         "PCI-DSS compliant payment gateway integration supporting multiple payment methods including credit/debit cards and digital wallets."),
        
        (icon_paths[3], "Order Management", 
         "Track orders from placement to delivery with real-time updates and notifications for both users and administrators."),
        
        (icon_paths[4], "User Account Management", 
         "Personal profiles with wishlists, order history, saved addresses, and preference settings for a personalized experience."),
        
        (icon_paths[5], "Admin Dashboard", 
         "Complete control over products, orders, users, and inventory with an intuitive administrative interface and comprehensive analytics.")
    ]
    
    create_features_with_icons(doc, features_data)
    
    add_page_break(doc)
    
    # 3. System Architecture
    add_heading_with_style(doc, "3. System Architecture", font_size=16, font_color=RGBColor(52, 152, 219))
    
    architecture_text = """
ShopSleek follows a modern three-tier architecture that separates the application into distinct layers, each with its own responsibilities. This architectural approach enhances maintainability, scalability, and security by isolating different components of the system.
"""
    add_paragraph_with_style(doc, architecture_text)
    
    # Add architecture diagram
    add_image_with_caption(doc, diagrams["architecture"], width=6, 
                         caption="ShopSleek Three-Tier Architecture", centered=True)
    
    # Architecture details
    arch_details = [
        "Presentation Layer: The top-most layer that directly interacts with users through the User Interface, Admin Dashboard, and Vendor Panel.",
        
        "Application Layer: The middle layer that processes business logic, including User Management, Product Management, Order Processing, and Payment Gateway integration.",
        
        "Data Layer: The foundation layer that handles data storage and retrieval through MySQL Database and File Storage."
    ]
    
    add_bulleted_list(doc, arch_details)
    
    add_page_break(doc)
    
    # 4. Test Cases
    add_heading_with_style(doc, "4. Test Cases", font_size=16, font_color=RGBColor(52, 152, 219))
    
    test_intro = """
Comprehensive testing was conducted to ensure the platform's functionality, reliability, and security. Each test case was carefully designed to validate specific features and identify potential issues before deployment.
"""
    add_paragraph_with_style(doc, test_intro)
    
    # Test Results Chart
    add_image_with_caption(doc, diagrams["test_chart"], width=6, 
                         caption="Test Results Summary by Category", centered=True)
    
    # User Authentication test cases
    auth_test_header = ["ID", "Test Steps", "Test Data", "Expected Result", "Actual Result", "Status"]
    
    auth_test_data = [
        ["TC01", "Login with valid credentials", "user@example.com\nCorrect123", "Successful login", "User logged in and redirected to dashboard", "Pass"],
        
        ["TC02", "Login with invalid credentials", "user@example.com\nWrongPass123", "Error message displayed", "Error message: 'Invalid email or password'", "Pass"],
        
        ["TC03", "Leave fields blank and attempt login", "Email: [blank]\nPassword: [blank]", "Form validation errors", "Error messages displayed for all required fields", "Pass"],
        
        ["TC04", "Register with valid information", "Name: John Doe\nEmail: john@example.com\nPassword: SecurePass123", "Account created successfully", "Account created and confirmation email received", "Pass"]
    ]
    
    add_test_case_table(doc, "User Authentication Test Cases", auth_test_header, auth_test_data)
    
    # Product Search test cases
    search_test_header = ["ID", "Test Steps", "Test Data", "Expected Result", "Actual Result", "Status"]
    
    search_test_data = [
        ["TC05", "Search with valid term", "Search: 'smartphone'", "Relevant smartphone products displayed", "5 smartphone products displayed with correct details", "Pass"],
        
        ["TC06", "Search with no matches", "Search: 'nonexistentitem123'", "Empty results message shown", "Message: 'No products match your search criteria'", "Pass"],
        
        ["TC07", "Filter by category", "Category: Electronics", "Electronics products shown", "15 Electronics products displayed correctly", "Pass"],
        
        ["TC08", "Filter by price range", "Price: ₹5000 - ₹15000", "Products in price range", "8 products displayed with prices between ₹5000-₹15000", "Pass"]
    ]
    
    add_test_case_table(doc, "Product Search Test Cases", search_test_header, search_test_data)
    
    add_page_break(doc)
    
    # Shopping Cart test cases
    cart_test_header = ["ID", "Test Steps", "Test Data", "Expected Result", "Actual Result", "Status"]
    
    cart_test_data = [
        ["TC09", "Add item to cart", "Product: Smartphone XYZ", "Product added correctly", "Added with price ₹15,999", "Pass"],
        
        ["TC10", "Update quantity", "Increase quantity to 3", "Price recalculated", "Total price ₹47,997", "Pass"],
        
        ["TC11", "Multiple items", "Smartphone, Headphones", "Correct items and total", "Both items with correct total ₹17,998", "Pass"],
        
        ["TC12", "Limited stock handling", "Available: 2, Request: 5", "System limits quantity", "Error: 'Only 2 units available'", "Pass"]
    ]
    
    add_test_case_table(doc, "Shopping Cart Test Cases", cart_test_header, cart_test_data)
    
    add_page_break(doc)
    
    # 5. Security Aspects
    add_heading_with_style(doc, "5. Security Aspects", font_size=16, font_color=RGBColor(52, 152, 219))
    
    security_intro = """
Security is a critical aspect of the ShopSleek e-commerce platform, especially considering the sensitive nature of user data and payment information being processed. The platform implements multiple layers of security measures to protect both user data and system integrity.
"""
    add_paragraph_with_style(doc, security_intro)
    
    # Add security diagram
    add_image_with_caption(doc, diagrams["security"], width=6, 
                         caption="ShopSleek Security Framework", centered=True)
    
    security_features = [
        "User Authentication: Password hashing with bcrypt, two-factor authentication, and secure session management",
        
        "Data Encryption: AES-256 encryption for sensitive data with secure key management",
        
        "Access Control: Role-based permissions with principle of least privilege",
        
        "PCI-DSS Compliance: For secure payment card handling",
        
        "GDPR Compliance: For personal data protection and privacy",
        
        "Regular Security Audits: Third-party penetration testing and vulnerability assessments",
        
        "Network Security: TLS 1.3 encryption, firewalls, and DDoS mitigation measures"
    ]
    
    add_bulleted_list(doc, security_features)
    
    add_page_break(doc)
    
    # 6. Deployment Architecture
    add_heading_with_style(doc, "6. Deployment Architecture", font_size=16, font_color=RGBColor(52, 152, 219))
    
    deployment_intro = """
The deployment architecture for ShopSleek is designed to provide high availability, scalability, and security. The platform is hosted on HostingRaja's cloud infrastructure with the following key components:
"""
    add_paragraph_with_style(doc, deployment_intro)
    
    # Create simple deployment diagram
    deployment_path = "images/deployment_diagram.png"
    deploy_img = Image.new('RGB', (800, 500), color=(255, 255, 255))
    draw = ImageDraw.Draw(deploy_img)
    
    try:
        title_font = ImageFont.truetype("arial.ttf", 24)
        header_font = ImageFont.truetype("arial.ttf", 18)
        label_font = ImageFont.truetype("arial.ttf", 14)
    except IOError:
        title_font = ImageFont.load_default()
        header_font = ImageFont.load_default()
        label_font = ImageFont.load_default()
    
    # Title
    draw.text((400, 30), "ShopSleek Deployment Architecture", font=title_font, fill=(44, 62, 80), anchor="mm")
    
    # Cloud outline
    cloud_points = [
        (150, 100), (250, 80), (350, 100), (400, 80), (500, 100), (550, 80),
        (600, 100), (650, 150), (600, 200), (650, 250), (600, 300), (500, 320),
        (400, 300), (350, 320), (250, 300), (200, 320), (150, 300), (100, 250),
        (150, 200), (100, 150)
    ]
    draw.polygon(cloud_points, fill=(240, 240, 240), outline=(200, 200, 200), width=2)
    draw.text((400, 120), "HostingRaja Cloud Infrastructure", font=header_font, fill=(52, 152, 219), anchor="mm")
    
    # Server components
    components = [
        {"name": "Load Balancer", "x": 300, "y": 180, "width": 120, "height": 40, "color": (52, 152, 219)},
        {"name": "App Server 1", "x": 200, "y": 250, "width": 100, "height": 40, "color": (46, 204, 113)},
        {"name": "App Server 2", "x": 400, "y": 250, "width": 100, "height": 40, "color": (46, 204, 113)},
        {"name": "App Server 3", "x": 600, "y": 250, "width": 100, "height": 40, "color": (46, 204, 113)},
        {"name": "Database", "x": 300, "y": 320, "width": 200, "height": 40, "color": (231, 76, 60)},
        {"name": "Redis Cache", "x": 550, "y": 320, "width": 100, "height": 40, "color": (243, 156, 18)},
        {"name": "Monitoring", "x": 200, "y": 390, "width": 100, "height": 40, "color": (155, 89, 182)},
        {"name": "Backup", "x": 400, "y": 390, "width": 100, "height": 40, "color": (52, 73, 94)}
    ]
    
    # Draw components
    for comp in components:
        x1 = comp["x"] - comp["width"]//2
        y1 = comp["y"] - comp["height"]//2
        x2 = comp["x"] + comp["width"]//2
        y2 = comp["y"] + comp["height"]//2
        
        draw.rectangle([(x1, y1), (x2, y2)], fill=comp["color"], outline=(44, 62, 80), width=2)
        draw.text((comp["x"], comp["y"]), comp["name"], font=label_font, fill=(255, 255, 255), anchor="mm")
    
    # Draw connections
    # Load balancer to app servers
    draw.line([(300, 200), (200, 230)], fill=(44, 62, 80), width=2)
    draw.line([(300, 200), (400, 230)], fill=(44, 62, 80), width=2)
    draw.line([(300, 200), (600, 230)], fill=(44, 62, 80), width=2)
    
    # App servers to database and cache
    draw.line([(200, 270), (300, 300)], fill=(44, 62, 80), width=2)
    draw.line([(400, 270), (300, 300)], fill=(44, 62, 80), width=2)
    draw.line([(600, 270), (300, 300)], fill=(44, 62, 80), width=2)
    
    draw.line([(200, 270), (550, 300)], fill=(44, 62, 80), width=2)
    draw.line([(400, 270), (550, 300)], fill=(44, 62, 80), width=2)
    draw.line([(600, 270), (550, 300)], fill=(44, 62, 80), width=2)
    
    # Monitoring and backup connections
    draw.line([(200, 370), (300, 340)], fill=(44, 62, 80), width=2, joint="curve")
    draw.line([(400, 370), (300, 340)], fill=(44, 62, 80), width=2, joint="curve")
    
    # Save deployment diagram
    deploy_img.save(deployment_path)
    
    # Add deployment diagram
    add_image_with_caption(doc, deployment_path, width=6, 
                         caption="ShopSleek Deployment Architecture", centered=True)
    
    # Deployment details
    deployment_features = [
        "Load Balancing: Nginx for distributing traffic across multiple application servers",
        
        "Application Servers: Node.js with Express running the ShopSleek application",
        
        "Database: MySQL with master-slave replication for redundancy and read scaling",
        
        "Caching: Redis for session management and data caching to improve performance",
        
        "Monitoring: ELK stack (Elasticsearch, Logstash, Kibana) for comprehensive monitoring",
        
        "Backup System: Automated daily backups with point-in-time recovery capabilities",
        
        "Security: Web Application Firewall (WAF) and DDoS protection"
    ]
    
    add_bulleted_list(doc, deployment_features)
    
    add_page_break(doc)
    
    # 7. Future Scope
    add_heading_with_style(doc, "7. Future Scope", font_size=16, font_color=RGBColor(52, 152, 219))
    
    future_scope_text = """
The ShopSleek platform has been designed with future expansion in mind. Several enhancements and new features are planned for upcoming releases:
"""
    add_paragraph_with_style(doc, future_scope_text)
    
    future_features = [
        "Mobile Applications: Native Android and iOS apps with offline capabilities",
        
        "AI-powered Recommendations: Personalized product suggestions using machine learning",
        
        "Voice Commerce: Integration with popular voice assistants for voice-based shopping",
        
        "Augmented Reality: Virtual product try-on experiences",
        
        "Internationalization: Multi-language and multi-currency support",
        
        "Subscription-based Models: Recurring billing capabilities",
        
        "Blockchain Integration: For secure transactions and product authenticity verification"
    ]
    
    add_bulleted_list(doc, future_features)
    
    # Conclusion
    conclusion_text = """
ShopSleek provides a robust, secure, and user-friendly e-commerce solution suitable for businesses of all sizes. With its comprehensive feature set, modern architecture, and focus on security, ShopSleek addresses the challenges of today's competitive online retail landscape while laying the groundwork for future innovations and enhancements.
"""
    add_paragraph_with_style(doc, conclusion_text)
    
    # Save the document
    doc.save('ShopSleek_Github_Documentation.docx')
    print("GitHub-compatible documentation created successfully!")

if __name__ == "__main__":
    create_document()