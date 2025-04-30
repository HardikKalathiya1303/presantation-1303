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

def create_document():
    """Create a comprehensive document for ShopSleek E-commerce platform with GitHub-compatible formatting."""
    doc = Document()
    
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
    
    features_bullets = [
        "Intuitive User Interface: Clean, responsive design optimized for all devices with thoughtful navigation paths and minimal friction points.",
        
        "Advanced Search & Filtering: Find products quickly with dynamic filtering capabilities that include category browsing, price range filters, and brand selection.",
        
        "Secure Payment Processing: PCI-DSS compliant payment gateway integration supporting multiple payment methods including credit/debit cards and digital wallets.",
        
        "Order Management: Track orders from placement to delivery with real-time updates and notifications for both users and administrators.",
        
        "User Account Management: Personal profiles with wishlists, order history, saved addresses, and preference settings for a personalized experience.",
        
        "Admin Dashboard: Complete control over products, orders, users, and inventory with an intuitive administrative interface and comprehensive analytics."
    ]
    
    add_bulleted_list(doc, features_bullets)
    
    add_page_break(doc)
    
    # 3. System Architecture
    add_heading_with_style(doc, "3. System Architecture", font_size=16, font_color=RGBColor(52, 152, 219))
    
    architecture_text = """
ShopSleek follows a modern three-tier architecture that separates the application into distinct layers, each with its own responsibilities. This architectural approach enhances maintainability, scalability, and security by isolating different components of the system.
"""
    add_paragraph_with_style(doc, architecture_text)
    
    # Architecture details
    arch_details = [
        "Presentation Layer: The top-most layer that directly interacts with users through the User Interface, Admin Dashboard, and Vendor Panel.",
        
        "Application Layer: The middle layer that processes business logic, including User Management, Product Management, Order Processing, and Payment Gateway integration.",
        
        "Data Layer: The foundation layer that handles data storage and retrieval through MySQL Database and File Storage."
    ]
    
    add_bulleted_list(doc, arch_details)
    
    technology_stack = """
Technology Stack:
• Frontend: React.js with Tailwind CSS
• Backend: Node.js with Express.js
• Database: MySQL for structured data storage
• Caching: Redis for session management and performance improvement
• Server: Node.js application server with PM2 process manager
• Load Balancer: Nginx for traffic distribution and SSL termination
• Hosting: HostingRaja cloud infrastructure
"""
    add_paragraph_with_style(doc, technology_stack)
    
    add_page_break(doc)
    
    # 4. Test Cases
    add_heading_with_style(doc, "4. Test Cases", font_size=16, font_color=RGBColor(52, 152, 219))
    
    test_intro = """
Comprehensive testing was conducted to ensure the platform's functionality, reliability, and security. Each test case was carefully designed to validate specific features and identify potential issues before deployment.
"""
    add_paragraph_with_style(doc, test_intro)
    
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
    
    # Test Results Summary
    test_summary = """
Test Results Summary:
• User Authentication Tests: 6 executed, 6 passed (100% success rate)
• Product Browsing Tests: 6 executed, 5 passed (83.3% success rate)
• Shopping & Checkout Tests: 6 executed, 5 passed (83.3% success rate)
• Admin Functions Tests: 8 executed, 7 passed (87.5% success rate)
• API Endpoint Tests: 10 executed, 9 passed (90% success rate)
• Overall System: 36 executed, 32 passed (88.9% success rate)
"""
    add_paragraph_with_style(doc, test_summary, font_size=12)
    
    add_page_break(doc)
    
    # 5. Security Aspects
    add_heading_with_style(doc, "5. Security Aspects", font_size=16, font_color=RGBColor(52, 152, 219))
    
    security_intro = """
Security is a critical aspect of the ShopSleek e-commerce platform, especially considering the sensitive nature of user data and payment information being processed. The platform implements multiple layers of security measures to protect both user data and system integrity.
"""
    add_paragraph_with_style(doc, security_intro)
    
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
    
    deployment_benefits = """
Benefits of this Architecture:
• High Availability: Redundancy at every level to eliminate single points of failure
• Scalability: Ability to scale horizontally by adding more application servers
• Performance: Caching and load balancing to optimize response times
• Security: Multi-layered security approach with encryption and access controls
• Maintainability: Separation of concerns for easier updates and maintenance
"""
    add_paragraph_with_style(doc, deployment_benefits)
    
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
    
    # Create a better organized future features table instead of feature assessment
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