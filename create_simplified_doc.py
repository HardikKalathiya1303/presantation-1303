from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

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

def create_document():
    """Create a comprehensive document for ShopSleek E-commerce platform."""
    doc = Document()
    
    # Set default paragraph spacing
    style = doc.styles['Normal']
    style.font.size = Pt(12)
    style.paragraph_format.space_after = Pt(6)
    
    # Title Page
    add_heading_with_style(doc, "ShopSleek", font_size=40, font_color=RGBColor(52, 152, 219), align=WD_ALIGN_PARAGRAPH.CENTER)
    add_heading_with_style(doc, "E-commerce Platform", level=2, font_size=30, font_color=RGBColor(52, 152, 219), align=WD_ALIGN_PARAGRAPH.CENTER)
    
    doc.add_paragraph()
    doc.add_paragraph()
    
    add_paragraph_with_style(doc, "Developed by: Hardik Kalathiya", font_size=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_paragraph_with_style(doc, "Under the Guidance of: Dr. Maitri Jhaveri", font_size=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_paragraph_with_style(doc, "Department of Computer Science", font_size=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_paragraph_with_style(doc, "Gujarat University", font_size=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    
    add_page_break(doc)
    
    # Table of Contents
    add_heading_with_style(doc, "Table of Contents", font_size=16, font_color=RGBColor(52, 152, 219))
    
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
        add_paragraph_with_style(doc, item)
    
    add_page_break(doc)
    
    # 1. Project Overview
    add_heading_with_style(doc, "1. Project Overview", font_size=16, font_color=RGBColor(52, 152, 219))
    
    overview_text = """
ShopSleek is a comprehensive e-commerce platform developed to provide users with a seamless online shopping experience. The platform is designed to handle various aspects of online retail, from product browsing and search to secure checkout and order management.

The system aims to provide an intuitive interface for customers while offering robust backend management capabilities for administrators. ShopSleek supports multiple device compatibility through responsive design and incorporates industry standard security practices to protect user data and transactions.
"""
    add_paragraph_with_style(doc, overview_text)
    
    add_heading_with_style(doc, "Core Functionalities:", level=2, font_size=14)
    
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
    
    add_bulleted_list(doc, core_functions)
    
    add_page_break(doc)
    
    # 2. Key Features
    add_heading_with_style(doc, "2. Key Features", font_size=16, font_color=RGBColor(52, 152, 219))
    
    features_text = """
ShopSleek offers a range of features designed to enhance both the user shopping experience and the administrative capabilities of the platform. The following are the key features that set ShopSleek apart from other e-commerce solutions:
"""
    add_paragraph_with_style(doc, features_text)
    
    # Add key features with descriptions
    feature_descriptions = [
        "Intuitive User Interface - Clean, responsive design optimized for all devices with thoughtful navigation paths and minimal friction points. The UI adapts dynamically to different screen sizes while maintaining usability and aesthetic appeal.",
        
        "Advanced Search & Filtering - Find products quickly with dynamic filtering capabilities that include category browsing, price range filters, brand selection, and attribute-based filtering. The search engine implements fuzzy matching to account for typos and returns relevant results even with partial queries.",
        
        "Secure Payment Processing - PCI-DSS compliant payment gateway integration that supports multiple payment methods including credit/debit cards, digital wallets, and bank transfers. All transactions are encrypted end-to-end and processed through secure channels.",
        
        "Order Management - Track orders from placement to delivery with real-time updates and notifications. Both users and administrators have access to comprehensive order history with filtering and search capabilities.",
        
        "User Accounts - Personal profiles with wishlists, order history, saved addresses, and preference settings. Users can manage multiple shipping addresses, payment methods, and notification preferences from their dashboard.",
        
        "Admin Dashboard - Complete control over products, orders, users, and inventory with an intuitive administrative interface. Administrators can monitor sales, manage product listings, handle order fulfillment, and generate various operational reports.",
        
        "Analytics & Reporting - Comprehensive sales reports, user activity tracking, and inventory level monitoring with visual data representation. The platform provides insights into customer behavior, popular products, and sales patterns to inform business decisions."
    ]
    
    for feature in feature_descriptions:
        parts = feature.split(" - ", 1)
        paragraph = doc.add_paragraph()
        paragraph.space_after = Pt(12)
        run = paragraph.add_run(parts[0])
        run.font.bold = True
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(52, 152, 219)
        
        if len(parts) > 1:
            run = paragraph.add_run(" - " + parts[1])
            run.font.size = Pt(12)
    
    add_page_break(doc)
    
    # 3. System Architecture
    add_heading_with_style(doc, "3. System Architecture", font_size=16, font_color=RGBColor(52, 152, 219))
    
    architecture_text = """
ShopSleek follows a modern three-tier architecture that separates the application into distinct layers, each with its own responsibilities. This architectural approach enhances maintainability, scalability, and security by isolating different components of the system.
"""
    add_paragraph_with_style(doc, architecture_text)
    
    # Add Three-Tier Architecture section
    add_heading_with_style(doc, "Three-Tier Architecture:", level=2, font_size=14)
    
    three_tier = [
        "Presentation Layer: The top-most layer that directly interacts with users through the User Interface, Admin Dashboard, and Vendor Panel. This layer is responsible for rendering data in a user-friendly format and collecting user inputs.",
        
        "Application Layer: The middle layer that processes business logic, including User Management, Product Management, Order Processing, and Payment Gateway integration. This layer acts as an intermediary between the presentation and data layers, applying business rules and ensuring data integrity.",
        
        "Data Layer: The foundation layer that handles data storage and retrieval through MySQL Database and File Storage. This layer is optimized for data persistence, security, and efficient query processing."
    ]
    
    for item in three_tier:
        parts = item.split(": ", 1)
        paragraph = doc.add_paragraph()
        paragraph.space_after = Pt(12)
        run = paragraph.add_run(parts[0] + ": ")
        run.font.bold = True
        run.font.size = Pt(13)
        
        if len(parts) > 1:
            run = paragraph.add_run(parts[1])
            run.font.size = Pt(12)
    
    # Add Technology Stack section
    add_heading_with_style(doc, "Technology Stack:", level=2, font_size=14)
    
    tech_stack = [
        "Frontend: React.js, Tailwind CSS - Modern JavaScript framework with utility-first CSS for building responsive user interfaces",
        "Backend: Node.js, Express.js - JavaScript runtime environment with a minimalist web framework for building APIs and server-side logic",
        "Database: MySQL - Relational database management system for structured data storage with powerful query capabilities",
        "Server: Node.js - Lightweight, efficient, and scalable server environment for handling HTTP requests",
        "Hosting: HostingRaja - Reliable hosting service with high uptime guarantee and scalable infrastructure"
    ]
    
    for item in tech_stack:
        parts = item.split(": ", 1)
        paragraph = doc.add_paragraph()
        paragraph.space_after = Pt(12)
        run = paragraph.add_run(parts[0] + ": ")
        run.font.bold = True
        run.font.size = Pt(13)
        
        if len(parts) > 1:
            secondary_parts = parts[1].split(" - ", 1)
            run = paragraph.add_run(secondary_parts[0])
            run.font.size = Pt(12)
            
            if len(secondary_parts) > 1:
                run = paragraph.add_run(" - " + secondary_parts[1])
                run.font.size = Pt(12)
                run.font.italic = True
    
    add_page_break(doc)
    
    # 4. Test Cases
    add_heading_with_style(doc, "4. Test Cases", font_size=16, font_color=RGBColor(52, 152, 219))
    
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
                    
                    # For "Overall System" row, make it bold
                    if row_data[0] == "Overall System":
                        run.font.bold = True
                        set_cell_background(row_cells[i], "E5F0FF")  # Light blue background
    
    # Add analysis of test results
    doc.add_paragraph("")
    
    test_analysis = """
The test results indicate an overall success rate of 88.9%, which demonstrates that the ShopSleek platform is functioning well in most areas. The User Authentication functionality achieved a perfect score, indicating robust security for user accounts. 

Areas requiring attention include specific features in Product Browsing and Shopping & Checkout, where failures were identified. These issues have been documented and prioritized for resolution before the final deployment.
"""
    add_paragraph_with_style(doc, test_analysis)
    
    add_page_break(doc)
    
    # 5. Security Aspects
    add_heading_with_style(doc, "5. Security Aspects", font_size=16, font_color=RGBColor(52, 152, 219))
    
    security_intro = """
Security is a critical aspect of the ShopSleek e-commerce platform, especially considering the sensitive nature of user data and payment information being processed. The platform implements multiple layers of security measures to protect both user data and system integrity.
"""
    add_paragraph_with_style(doc, security_intro)
    
    # 5.1 Data Protection
    add_heading_with_style(doc, "5.1 Data Protection", level=2, font_size=14)
    
    data_protection = [
        "Password Encryption: All user passwords are encrypted using bcrypt hashing algorithm with salt, ensuring that even in the event of a database breach, passwords remain secure. The system enforces strong password policies including minimum length, complexity requirements, and regular password rotation for admin accounts.",
        
        "Sensitive Data Encryption: Personal information and payment details are encrypted using AES-256 encryption in the database. Encryption keys are managed using a secure key management system with regular rotation.",
        
        "Role-based Access Control: The system implements granular access controls that restrict users to only the data and functions necessary for their role. This limits the potential damage from compromised accounts and enforces the principle of least privilege.",
        
        "Session Management: Secure session handling with automatic timeout after periods of inactivity. Sessions are invalidated upon logout and implement secure cookies with appropriate flags (Secure, HttpOnly, SameSite).",
        
        "Regular Security Audits: Scheduled security assessments and penetration testing by third-party security experts to identify and address vulnerabilities. This includes code reviews, dependency scanning, and infrastructure security testing.",
        
        "Two-factor Authentication: Optional for regular users and mandatory for administrative accounts, adding an additional layer of security beyond passwords.",
        
        "Data Loss Prevention: Regular automated backups with encryption and secure off-site storage. Point-in-time recovery capabilities and disaster recovery procedures are tested regularly."
    ]
    
    for item in data_protection:
        parts = item.split(": ", 1)
        paragraph = doc.add_paragraph()
        paragraph.space_after = Pt(12)
        run = paragraph.add_run(parts[0] + ": ")
        run.font.bold = True
        run.font.size = Pt(12)
        
        if len(parts) > 1:
            run = paragraph.add_run(parts[1])
            run.font.size = Pt(12)
    
    # 5.2 Security Implementation
    add_heading_with_style(doc, "5.2 Security Implementation", level=2, font_size=14)
    
    security_implementation = [
        "SQL Injection Protection: All database interactions use prepared statements and parameterized queries rather than direct string concatenation. Input validation occurs at both client and server sides, with ORM frameworks providing additional protection layers.",
        
        "XSS Attack Prevention: Implementation of Content Security Policy (CSP) headers and output encoding for all user-generated content. React's built-in XSS protections provide additional security by automatically escaping content.",
        
        "CSRF Attack Mitigation: Unique tokens are generated for each session and validated for all state-changing operations. These tokens are rotated regularly and tied to the user's session.",
        
        "Session Hijacking Prevention: Secure cookies with appropriate flags and IP binding for administrative sessions. Session regeneration occurs after login, privilege changes, and at regular intervals.",
        
        "Brute Force Attack Prevention: Progressive account lockout after multiple failed login attempts, with CAPTCHA implementation to prevent automated attacks. Notifications are sent to users when suspicious login activities are detected.",
        
        "Data Protection Compliance: Implementation of PCI-DSS requirements for payment card handling and GDPR compliance for personal data protection. This includes consent management, data minimization, and right to be forgotten capabilities.",
        
        "Network Security: Implementation of TLS 1.3 for all communications with HSTS to prevent downgrade attacks. Regular scans for misconfigured certificates and weak cipher suites.",
        
        "Infrastructure Security: Regular security patches and updates for all system components. Automated vulnerability scanning and intrusion detection systems monitor for suspicious activities."
    ]
    
    for item in security_implementation:
        parts = item.split(": ", 1)
        paragraph = doc.add_paragraph()
        paragraph.space_after = Pt(12)
        run = paragraph.add_run(parts[0] + ": ")
        run.font.bold = True
        run.font.size = Pt(12)
        
        if len(parts) > 1:
            run = paragraph.add_run(parts[1])
            run.font.size = Pt(12)
    
    add_page_break(doc)
    
    # 6. Deployment Architecture
    add_heading_with_style(doc, "6. Deployment Architecture", font_size=16, font_color=RGBColor(52, 152, 219))
    
    deployment_intro = """
The deployment architecture for ShopSleek is designed to provide high availability, scalability, and security. The platform is hosted on HostingRaja, a reliable hosting provider that offers robust infrastructure and excellent support services.
"""
    add_paragraph_with_style(doc, deployment_intro)
    
    # 6.1 Hosting Infrastructure
    add_heading_with_style(doc, "6.1 Hosting Infrastructure", level=2, font_size=14)
    
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
    add_heading_with_style(doc, "6.2 Server Configuration", level=2, font_size=14)
    
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
    add_heading_with_style(doc, "6.3 Database Setup", level=2, font_size=14)
    
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
    add_heading_with_style(doc, "6.4 Backup and Monitoring", level=2, font_size=14)
    
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
    add_heading_with_style(doc, "7. Future Scope", font_size=16, font_color=RGBColor(52, 152, 219))
    
    future_scope_text = """
The ShopSleek platform has been designed with future expansion in mind. Several enhancements and new features are planned for upcoming releases:
"""
    add_paragraph_with_style(doc, future_scope_text)
    
    future_scope = [
        "Mobile Applications - Native Android and iOS apps that provide an optimized shopping experience for mobile users. These applications will feature offline browsing capabilities, push notifications for order updates, and biometric authentication for enhanced security.",
        
        "AI-powered Recommendations - Personalized product suggestions based on browsing history, purchase patterns, and similar user behaviors. The recommendation engine will use machine learning algorithms that improve over time as more data is collected.",
        
        "Advanced Analytics - Customer behavior analysis and sales forecasting tools to help businesses make data-driven decisions. This will include cohort analysis, funnel visualization, and predictive analytics for inventory management.",
        
        "Voice Commerce - Integration with popular voice assistants like Amazon Alexa and Google Assistant to enable shopping through voice commands. Users will be able to search for products, check order status, and place orders using natural language.",
        
        "Augmented Reality - Virtual product try-on experiences that allow customers to visualize products in their own environment before purchasing. This will be particularly useful for furniture, home decor, and fashion items.",
        
        "Internationalization - Multi-language and multi-currency support to expand the platform's reach to global markets. This includes localized payment methods, shipping options, and compliance with regional regulations.",
        
        "Marketplace Expansion - Enhanced vendor management tools to transform the platform into a comprehensive marketplace where multiple sellers can list and sell their products. This will include vendor onboarding workflows, commission management, and seller analytics.",
        
        "Integration with Social Commerce - Ability to buy products directly from social media platforms through API integrations. This will reduce friction in the purchase journey and tap into social discovery as a sales channel.",
        
        "Subscription-based Models - Support for recurring billing and subscription management to capture recurring revenue streams. This will include flexible billing cycles, automatic renewals, and subscription analytics.",
        
        "Blockchain Integration - Implementation of blockchain technology for secure transactions, product authenticity verification, and loyalty program management. This will enhance trust and provide new opportunities for customer engagement."
    ]
    
    for item in future_scope:
        parts = item.split(" - ", 1)
        paragraph = doc.add_paragraph()
        paragraph.space_after = Pt(12)
        run = paragraph.add_run(parts[0] + " - ")
        run.font.bold = True
        run.font.size = Pt(12)
        
        if len(parts) > 1:
            run = paragraph.add_run(parts[1])
            run.font.size = Pt(12)
    
    # Save the document
    doc.save('ShopSleek_Enhanced_Documentation.docx')
    print("Documentation created successfully!")

if __name__ == "__main__":
    create_document()