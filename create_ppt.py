from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

def add_title_slide(prs, title, subtitle):
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    slide.placeholders[1].text = subtitle
    title_shape = slide.shapes.title
    title_shape.text_frame.paragraphs[0].font.size = Pt(44)
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(44, 62, 80)
    return slide

def add_section_slide(prs, title):
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = title
    title_shape = slide.shapes.title
    title_shape.text_frame.paragraphs[0].font.size = Pt(40)
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(52, 152, 219)
    return slide

def add_content_slide(prs, title, bullet_points):
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = title
    
    content = slide.placeholders[1]
    text_frame = content.text_frame
    
    for point in bullet_points:
        p = text_frame.add_paragraph()
        p.text = point
        p.level = 0
        p.font.size = Pt(24)
    
    return slide

def add_two_column_slide(prs, title, left_content, right_content, left_title=None, right_title=None):
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    slide.shapes.title.text = title
    
    # Define shape dimensions
    left_column_width = Inches(4.5)
    right_column_width = Inches(4.5)
    column_height = Inches(5)
    left_column_left = Inches(0.5)
    right_column_left = Inches(5.5)
    column_top = Inches(1.8)
    
    # Add left column
    left_shape = slide.shapes.add_textbox(left_column_left, column_top, left_column_width, column_height)
    left_text_frame = left_shape.text_frame
    left_text_frame.word_wrap = True
    
    if left_title:
        p = left_text_frame.add_paragraph()
        p.text = left_title
        p.font.bold = True
        p.font.size = Pt(24)
        p.font.color.rgb = RGBColor(41, 128, 185)
        p.alignment = PP_ALIGN.LEFT
    
    for point in left_content:
        p = left_text_frame.add_paragraph()
        p.text = f"• {point}"
        p.font.size = Pt(20)
        p.alignment = PP_ALIGN.LEFT
    
    # Add right column
    right_shape = slide.shapes.add_textbox(right_column_left, column_top, right_column_width, column_height)
    right_text_frame = right_shape.text_frame
    right_text_frame.word_wrap = True
    
    if right_title:
        p = right_text_frame.add_paragraph()
        p.text = right_title
        p.font.bold = True
        p.font.size = Pt(24)
        p.font.color.rgb = RGBColor(41, 128, 185)
        p.alignment = PP_ALIGN.LEFT
    
    for point in right_content:
        p = right_text_frame.add_paragraph()
        p.text = f"• {point}"
        p.font.size = Pt(20)
        p.alignment = PP_ALIGN.LEFT
    
    return slide

def add_test_case_slide(prs, title, test_cases):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title
    
    rows = len(test_cases) + 1  # +1 for header row
    cols = 6
    
    left = Inches(0.5)
    top = Inches(1.5)
    width = Inches(9)
    height = Inches(5)
    
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table
    
    # Define column widths
    table.columns[0].width = Inches(0.5)  # ID
    table.columns[1].width = Inches(1.8)  # Test Steps
    table.columns[2].width = Inches(1.7)  # Test Data
    table.columns[3].width = Inches(1.8)  # Expected Result
    table.columns[4].width = Inches(1.8)  # Actual Result
    table.columns[5].width = Inches(1.0)  # Status
    
    # Add header row
    headers = ["ID", "Test Steps", "Test Data", "Expected Result", "Actual Result", "Status"]
    for i, header in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = header
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(52, 152, 219)
        
        # Format all paragraphs in the cell
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.bold = True
            paragraph.font.size = Pt(14)
            paragraph.font.color.rgb = RGBColor(255, 255, 255)
            paragraph.alignment = PP_ALIGN.CENTER
    
    # Add test case rows
    for i, test_case in enumerate(test_cases):
        row_index = i + 1
        for j, value in enumerate(test_case):
            cell = table.cell(row_index, j)
            cell.text = value
            
            # Format all paragraphs in the cell
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.size = Pt(12)
                
                # Highlight Pass/Fail status
                if j == 5:  # Status column
                    if value.lower() == "pass":
                        paragraph.font.color.rgb = RGBColor(39, 174, 96)
                        paragraph.font.bold = True
                    elif value.lower() == "fail":
                        paragraph.font.color.rgb = RGBColor(231, 76, 60)
                        paragraph.font.bold = True
    
    return slide

def create_presentation():
    prs = Presentation()
    
    # Add title slide
    add_title_slide(prs, "ShopSleek", "E-commerce Platform\n\nDeveloped by: Hardik Kalathiya\nUnder the Guidance of: Dr. Maitri Jhaveri\nDepartment of Computer Science\nGujarat University")
    
    # Project Overview
    project_overview = [
        "Online shopping platform for various products",
        "User registration and account management",
        "Guest shopping capabilities",
        "Product search and filtering",
        "Shopping cart management",
        "Secure checkout process",
        "Order tracking",
        "Admin dashboard for system management",
        "Multi-vendor support system"
    ]
    add_content_slide(prs, "Project Overview", project_overview)
    
    # Key Features
    key_features = [
        "Intuitive User Interface - Clean, responsive design optimized for all devices",
        "Advanced Search & Filtering - Find products quickly with dynamic filtering",
        "Secure Payment Processing - PCI-DSS compliant payment gateway integration",
        "Order Management - Track orders from placement to delivery",
        "User Accounts - Personal profiles, wishlists, and order history",
        "Admin Dashboard - Complete control over products, orders, and users",
        "Analytics & Reporting - Sales reports, user activity, and inventory levels"
    ]
    add_content_slide(prs, "Key Features", key_features)
    
    # System Architecture
    three_tier = [
        "Presentation Layer: User Interface, Admin Dashboard, Vendor Panel",
        "Application Layer: User Management, Product Management, Order Processing, Payment Gateway",
        "Data Layer: MySQL Database, File Storage"
    ]
    
    tech_stack = [
        "Frontend: React.js, Tailwind CSS",
        "Backend: Node.js, Express.js",
        "Database: MySQL",
        "Server: Node.js",
        "Hosting: HostingRaja"
    ]
    
    add_two_column_slide(prs, "System Architecture", three_tier, tech_stack, "Three-Tier Architecture", "Technology Stack")
    
    # Test Cases - User Authentication
    add_section_slide(prs, "Test Cases")
    
    # Test Case 1: User Login
    login_test_cases = [
        ["TC01", "1. Navigate to login page\n2. Enter valid credentials\n3. Click on Login button", "Email: user@example.com\nPassword: Correct123", "User logged in and redirected to dashboard", "User successfully logged in and redirected to dashboard", "Pass"],
        ["TC02", "1. Navigate to login page\n2. Enter invalid credentials\n3. Click on Login button", "Email: user@example.com\nPassword: WrongPass123", "Error message displayed", "Error message: 'Invalid email or password'", "Pass"],
        ["TC03", "1. Navigate to login page\n2. Leave fields blank\n3. Click on Login button", "Email: [blank]\nPassword: [blank]", "Form validation errors", "Error messages displayed for all required fields", "Pass"]
    ]
    add_test_case_slide(prs, "Test Cases - User Login", login_test_cases)
    
    # Test Case 2: User Registration
    registration_test_cases = [
        ["TC04", "1. Navigate to registration page\n2. Fill all required fields\n3. Click Register button", "Name: John Doe\nEmail: john@example.com\nPassword: SecurePass123", "Account created with confirmation email sent", "Account created and confirmation email received", "Pass"],
        ["TC05", "1. Navigate to registration page\n2. Leave all fields empty\n3. Click Register button", "All fields empty", "Error message displayed for each required field", "Error message displayed: 'All fields are required'", "Pass"],
        ["TC06", "1. Navigate to registration page\n2. Enter invalid email format\n3. Click Register button", "Name: Jane Smith\nEmail: jane.emailcom\nPassword: Pass123", "Validation error shown for email field", "Validation error: 'Please enter a valid email address'", "Pass"]
    ]
    add_test_case_slide(prs, "Test Cases - User Registration", registration_test_cases)
    
    # Test Case 3: Product Search
    product_test_cases = [
        ["TC07", "1. Navigate to search bar\n2. Enter valid search term\n3. Press search button", "Search term: 'smartphone'", "Relevant smartphone products displayed", "5 smartphone products displayed with images and prices", "Pass"],
        ["TC08", "1. Navigate to search bar\n2. Enter search term with no matches\n3. Press search button", "Search term: 'nonexistentitem123'", "Empty results message shown", "Message: 'No products match your search criteria'", "Pass"],
        ["TC09", "1. Navigate to category menu\n2. Select Electronics category\n3. View filtered results", "Category: Electronics", "Products from Electronics category displayed", "15 Electronics products displayed correctly", "Pass"]
    ]
    add_test_case_slide(prs, "Test Cases - Product Search", product_test_cases)
    
    # Test Case 4: Shopping Cart
    cart_test_cases = [
        ["TC10", "1. View product details\n2. Click 'Add to Cart'\n3. Verify cart updates", "Product: Smartphone XYZ\nQuantity: 1", "Product added to cart with correct price", "Product successfully added to cart with price ₹15,999", "Pass"],
        ["TC11", "1. View shopping cart\n2. Update quantity to 3\n3. Check total price", "Product: Smartphone XYZ\nNew Quantity: 3", "Quantity updated and total price recalculated", "Quantity updated to 3, total price ₹47,997", "Pass"],
        ["TC12", "1. View shopping cart\n2. Click 'Remove' button\n3. Verify item removed", "Product: Smartphone XYZ", "Item removed from cart", "Item successfully removed from cart, cart empty", "Pass"]
    ]
    add_test_case_slide(prs, "Test Cases - Shopping Cart", cart_test_cases)
    
    # Test Case 5: Checkout Process
    checkout_test_cases = [
        ["TC13", "1. Add products to cart\n2. Click 'Checkout' button\n3. Enter shipping info\n4. Select payment method\n5. Confirm order", "Valid shipping & payment details", "Order placed successfully", "Order placed with confirmation #12345", "Pass"],
        ["TC14", "1. Add products to cart\n2. Click 'Checkout' button\n3. Enter invalid card details\n4. Attempt to confirm", "Invalid card number: 1234 5678 9012 3456", "Payment error displayed", "Error: 'Invalid card number'", "Pass"],
        ["TC15", "1. Add products to cart\n2. Click 'Checkout' button\n3. Leave shipping fields empty\n4. Attempt to confirm", "Empty shipping fields", "Validation errors for required fields", "Error messages shown for all required fields", "Pass"]
    ]
    add_test_case_slide(prs, "Test Cases - Checkout Process", checkout_test_cases)
    
    # Test Results Summary
    summary_data = [
        ["Category", "Total Tests", "Passed", "Failed", "Success Rate"],
        ["User Authentication", "6", "6", "0", "100%"],
        ["Product Browsing", "6", "5", "1", "83.3%"],
        ["Shopping & Checkout", "6", "5", "1", "83.3%"],
        ["Admin Functions", "8", "7", "1", "87.5%"],
        ["API Endpoints", "10", "9", "1", "90%"],
        ["Overall System", "36", "32", "4", "88.9%"]
    ]
    
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Test Results Summary"
    
    rows = len(summary_data)
    cols = len(summary_data[0])
    
    left = Inches(1.2)
    top = Inches(2)
    width = Inches(8)
    height = Inches(3.5)
    
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table
    
    # Add header row
    for i, header_text in enumerate(summary_data[0]):
        cell = table.cell(0, i)
        cell.text = header_text
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(52, 152, 219)
        
        # Format all paragraphs in the cell
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.bold = True
            paragraph.font.size = Pt(16)
            paragraph.font.color.rgb = RGBColor(255, 255, 255)
            paragraph.alignment = PP_ALIGN.CENTER
    
    # Add data rows
    for i in range(1, rows):
        for j in range(cols):
            cell = table.cell(i, j)
            cell.text = summary_data[i][j]
            
            # Format all paragraphs in the cell
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.size = Pt(14)
                paragraph.alignment = PP_ALIGN.CENTER
                
                # Highlight last row (Overall)
                if i == rows - 1:
                    paragraph.font.bold = True
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(235, 245, 255)
    
    # Enhanced Security Section
    security_aspects = [
        "Password encryption using strong hashing algorithms (bcrypt with salt)",
        "Sensitive data encryption in database (AES-256)",
        "Role-based access control mechanisms",
        "Session management and timeout controls",
        "Regular security audits and vulnerability scanning",
        "Two-factor authentication for admin access",
        "IP-based rate limiting to prevent brute force attacks",
        "Data loss prevention policies and backup strategies"
    ]
    
    security_implementation = [
        "SQL Injection: Prepared statements and parameterized queries",
        "XSS Attacks: Output encoding and content security policy",
        "CSRF Attacks: Unique tokens for each session",
        "Session Hijacking: Secure cookies and session regeneration",
        "Brute Force: Account lockout and CAPTCHA integration",
        "Data Protection: PCI-DSS compliance for payment data",
        "Network Security: TLS 1.3 for all communications",
        "Infrastructure: Regular security patches and updates"
    ]
    
    add_two_column_slide(prs, "Security Aspects", security_aspects, security_implementation, "Data Protection Measures", "Security Implementation")
    
    # Enhanced Deployment Architecture
    deployment_aspects = [
        "Hosting: Deployed on HostingRaja with 99.9% uptime guarantee",
        "Server Configuration: Node.js server with Express.js framework",
        "Database: MySQL 8.0 with optimized query performance",
        "SSL Implementation: HTTPS encryption for secure data transmission",
        "Scalability: Load balancing for high traffic management",
        "Backup System: Daily automated backups with point-in-time recovery",
        "Monitoring: Real-time performance and security monitoring",
        "CDN Integration: Content delivery network for optimized asset delivery",
        "CI/CD Pipeline: Automated testing and deployment process",
        "Containerization: Docker for consistent deployment environments",
        "Microservices Architecture: For improved scalability and maintainability",
        "Multi-region Deployment: For improved global performance and failover"
    ]
    
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    slide.shapes.title.text = "Enhanced Deployment Architecture"
    
    content = slide.placeholders[1]
    text_frame = content.text_frame
    
    for aspect in deployment_aspects:
        p = text_frame.add_paragraph()
        p.text = f"• {aspect}"
        p.level = 0
        p.font.size = Pt(18)
    
    # Future Scope
    future_scope = [
        "Mobile Applications - Native Android and iOS apps",
        "AI-powered Recommendations - Personalized product suggestions",
        "Advanced Analytics - Customer behavior and sales forecasting",
        "Voice Commerce - Integration with voice assistants",
        "Augmented Reality - Virtual product try-on experiences",
        "Internationalization - Multi-language and multi-currency support",
        "Marketplace Expansion - Platform for multiple vendors",
        "Integration with Social Commerce - Buy directly from social platforms",
        "Subscription-based Models - Recurring revenue streams",
        "Blockchain Integration - For secure transactions and product authenticity"
    ]
    
    add_content_slide(prs, "Future Scope", future_scope)
    
    # Thank You slide
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Thank You!"
    slide.placeholders[1].text = "Hardik Kalathiya\nDepartment of Computer Science\nGujarat University"
    title_shape = slide.shapes.title
    title_shape.text_frame.paragraphs[0].font.size = Pt(54)
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(44, 62, 80)
    
    # Save the presentation
    prs.save('ShopSleek_Enhanced_Presentation.pptx')
    print("Presentation created successfully!")

if __name__ == "__main__":
    create_presentation()