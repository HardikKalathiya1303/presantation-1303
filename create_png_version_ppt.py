from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
import os

def add_title_slide(prs, title, subtitle):
    """Add a title slide to the presentation."""
    slide_layout = prs.slide_layouts[0]  # Title Slide layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Set the slide background to a light blue color
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(235, 245, 255)
    
    # Add title
    title_shape = slide.shapes.title
    title_shape.text = title
    
    # Format title
    title_text_frame = title_shape.text_frame
    title_text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    title_run = title_text_frame.paragraphs[0].runs[0]
    title_run.font.size = Pt(44)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(52, 152, 219)
    
    # Add subtitle
    subtitle_shape = slide.placeholders[1]
    subtitle_shape.text = subtitle
    
    # Format subtitle
    subtitle_text_frame = subtitle_shape.text_frame
    subtitle_text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    subtitle_run = subtitle_text_frame.paragraphs[0].runs[0]
    subtitle_run.font.size = Pt(28)
    subtitle_run.font.bold = True
    subtitle_run.font.color.rgb = RGBColor(44, 62, 80)
    
    # Add author information
    author_shape = slide.shapes.add_textbox(
        Inches(1), Inches(5), Inches(8), Inches(1))
    text_frame = author_shape.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # Add developer info
    p = text_frame.add_paragraph()
    p.alignment = PP_ALIGN.CENTER
    p.text = "Developed by: Hardik Kalathiya"
    p.runs[0].font.size = Pt(14)
    p.runs[0].font.bold = True
    
    # Add guide info
    p = text_frame.add_paragraph()
    p.alignment = PP_ALIGN.CENTER
    p.text = "Under the Guidance of: Dr. Maitri Jhaveri"
    p.runs[0].font.size = Pt(14)
    p.runs[0].font.bold = True
    
    # Add department info
    p = text_frame.add_paragraph()
    p.alignment = PP_ALIGN.CENTER
    p.text = "Department of Computer Science"
    p.runs[0].font.size = Pt(14)
    
    # Add university info
    p = text_frame.add_paragraph()
    p.alignment = PP_ALIGN.CENTER
    p.text = "Gujarat University"
    p.runs[0].font.size = Pt(14)
    
    return slide

def add_section_slide(prs, title):
    """Add a section divider slide with a title."""
    slide_layout = prs.slide_layouts[2]  # Section Header layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Set the slide background to a gradient color
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(235, 245, 255)
    
    # Add title
    title_shape = slide.shapes.title
    title_shape.text = title
    
    # Format title
    title_text_frame = title_shape.text_frame
    title_text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    title_run = title_text_frame.paragraphs[0].runs[0]
    title_run.font.size = Pt(40)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(52, 152, 219)
    
    # Add decorative line
    line = slide.shapes.add_shape(
        1, Inches(2.5), Inches(3.5), Inches(5), Inches(0.05))
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(52, 152, 219)
    line.line.fill.background()
    
    return slide

def add_content_slide(prs, title, bullet_points):
    """Add a content slide with a title and bullet points."""
    slide_layout = prs.slide_layouts[1]  # Title and Content layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Add title
    title_shape = slide.shapes.title
    title_shape.text = title
    
    # Format title
    title_text_frame = title_shape.text_frame
    title_text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
    title_run = title_text_frame.paragraphs[0].runs[0]
    title_run.font.size = Pt(32)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(52, 152, 219)
    
    # Add bullet points
    content_placeholder = slide.placeholders[1]
    text_frame = content_placeholder.text_frame
    
    # Clear any existing text
    text_frame.clear()
    
    # Add each bullet point
    for point in bullet_points:
        p = text_frame.add_paragraph()
        p.text = point
        p.level = 0
        p.runs[0].font.size = Pt(18)
        p.runs[0].font.color.rgb = RGBColor(44, 62, 80)
    
    return slide

def add_two_column_slide(prs, title, left_content, right_content, left_title=None, right_title=None):
    """Add a slide with two columns of content."""
    slide_layout = prs.slide_layouts[3]  # Two Content layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Add title
    title_shape = slide.shapes.title
    title_shape.text = title
    
    # Format title
    title_text_frame = title_shape.text_frame
    title_text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    title_run = title_text_frame.paragraphs[0].runs[0]
    title_run.font.size = Pt(32)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(52, 152, 219)
    
    # Get the content placeholders
    left_placeholder = slide.placeholders[1]
    right_placeholder = slide.placeholders[2]
    
    # Add left content with title if provided
    left_text_frame = left_placeholder.text_frame
    left_text_frame.clear()
    
    if left_title:
        p = left_text_frame.add_paragraph()
        p.text = left_title
        p.alignment = PP_ALIGN.CENTER
        p.runs[0].font.bold = True
        p.runs[0].font.size = Pt(20)
        p.runs[0].font.color.rgb = RGBColor(52, 152, 219)
        
        # Add a blank line
        left_text_frame.add_paragraph()
    
    # Add left content
    for point in left_content:
        p = left_text_frame.add_paragraph()
        p.text = point
        p.level = 0
        p.runs[0].font.size = Pt(16)
        p.runs[0].font.color.rgb = RGBColor(44, 62, 80)
    
    # Add right content with title if provided
    right_text_frame = right_placeholder.text_frame
    right_text_frame.clear()
    
    if right_title:
        p = right_text_frame.add_paragraph()
        p.text = right_title
        p.alignment = PP_ALIGN.CENTER
        p.runs[0].font.bold = True
        p.runs[0].font.size = Pt(20)
        p.runs[0].font.color.rgb = RGBColor(52, 152, 219)
        
        # Add a blank line
        right_text_frame.add_paragraph()
    
    # Add right content
    for point in right_content:
        p = right_text_frame.add_paragraph()
        p.text = point
        p.level = 0
        p.runs[0].font.size = Pt(16)
        p.runs[0].font.color.rgb = RGBColor(44, 62, 80)
    
    return slide

def add_test_case_slide(prs, title, test_cases):
    """Add a slide with a test case table."""
    slide_layout = prs.slide_layouts[5]  # Title Only layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Add title
    title_shape = slide.shapes.title
    title_shape.text = title
    
    # Format title
    title_text_frame = title_shape.text_frame
    title_text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    title_run = title_text_frame.paragraphs[0].runs[0]
    title_run.font.size = Pt(28)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(52, 152, 219)
    
    # Define table dimensions and position
    rows = len(test_cases) + 1  # Header row + data rows
    cols = 6  # ID, Steps, Data, Expected, Actual, Status
    
    table = slide.shapes.add_table(rows, cols, Inches(0.3), Inches(1.5), 
                                   Inches(9.4), Inches(5)).table
    
    # Define column widths
    table.columns[0].width = Inches(0.6)  # ID
    table.columns[1].width = Inches(2.0)  # Test Steps
    table.columns[2].width = Inches(1.7)  # Test Data
    table.columns[3].width = Inches(1.8)  # Expected Result
    table.columns[4].width = Inches(2.3)  # Actual Result
    table.columns[5].width = Inches(1.0)  # Status
    
    # Header row
    header_cells = ['ID', 'Test Steps', 'Test Data', 'Expected Result', 'Actual Result', 'Status']
    for i, text in enumerate(header_cells):
        cell = table.cell(0, i)
        cell.text = text
        
        # Format header cell
        for paragraph in cell.text_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.CENTER
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(12)
                run.font.color.rgb = RGBColor(255, 255, 255)
        
        # Set header background color
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(52, 152, 219)
    
    # Data rows
    for row_idx, test_case in enumerate(test_cases, start=1):
        for col_idx, value in enumerate(test_case):
            cell = table.cell(row_idx, col_idx)
            cell.text = str(value)
            
            # Format data cell
            for paragraph in cell.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.LEFT
                for run in paragraph.runs:
                    run.font.size = Pt(10)
                    
                    # Highlight Pass/Fail status
                    if col_idx == 5:  # Status column
                        if value.lower() == "pass":
                            run.font.color.rgb = RGBColor(39, 174, 96)  # Green
                            run.font.bold = True
                        elif value.lower() == "fail":
                            run.font.color.rgb = RGBColor(231, 76, 60)  # Red
                            run.font.bold = True
    
    return slide

def create_visual_presentation():
    """Create a comprehensive visual presentation for ShopSleek E-commerce platform."""
    prs = Presentation()
    
    # Set slide width and height (16:9 aspect ratio)
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)
    
    # 1. Title Slide
    add_title_slide(prs, "ShopSleek", "E-commerce Platform")
    
    # 2. Overview Section
    add_section_slide(prs, "Project Overview")
    
    # 3. Overview Content
    overview_points = [
        "✓ Comprehensive e-commerce platform for seamless online shopping",
        "✓ Intuitive user interface with responsive design for all devices",
        "✓ Robust backend management capabilities for administrators",
        "✓ Secure transaction processing with industry-standard protocols",
        "✓ Comprehensive product search and filtering capabilities",
        "✓ Multi-vendor support with dedicated portals",
        "✓ Real-time analytics and reporting dashboard"
    ]
    add_content_slide(prs, "ShopSleek Overview", overview_points)
    
    # 4. Features Section
    add_section_slide(prs, "Key Features")
    
    # 5. Features Two Column
    left_features = [
        "• Intuitive User Interface",
        "• Advanced Search & Filtering",
        "• Secure Payment Processing",
        "• Order Management"
    ]
    
    right_features = [
        "• User Account Management",
        "• Admin Dashboard",
        "• Analytics & Reporting",
        "• Multi-vendor Support"
    ]
    add_two_column_slide(prs, "Key Features", left_features, right_features, 
                         "Customer-Facing Features", "Administrative Features")
    
    # 6. Architecture Section
    add_section_slide(prs, "System Architecture")
    
    # 7. Architecture Details
    architecture_points = [
        "➊ Presentation Layer: User interfaces, admin dashboards, and vendor panels",
        "➋ Application Layer: Business logic including user management, product catalog, order processing, and payment integration",
        "➌ Data Layer: MySQL database for structured data storage with robust query capabilities",
        "➍ Technology Stack: React.js, Node.js, Express.js, MySQL, and HostingRaja hosting infrastructure",
        "➎ Designed for scalability, maintainability, and security through component isolation"
    ]
    add_content_slide(prs, "Architecture Details", architecture_points)
    
    # 8. Test Cases Section
    add_section_slide(prs, "Test Cases")
    
    # 9. User Authentication Test Cases
    auth_test_cases = [
        ["TC01", "Login with valid credentials", "user@example.com\nCorrect123", "Successful login", "User logged in and redirected to dashboard", "Pass"],
        ["TC02", "Login with invalid credentials", "user@example.com\nWrongPass123", "Error message", "Error: 'Invalid email or password'", "Pass"],
        ["TC03", "Empty login fields", "Email: [blank]\nPassword: [blank]", "Form validation errors", "Error messages for required fields", "Pass"],
        ["TC04", "Valid registration", "Name: John Doe\nEmail: john@example.com", "Account created", "Account created and email sent", "Pass"]
    ]
    add_test_case_slide(prs, "User Authentication Test Cases", auth_test_cases)
    
    # 10. Product Search Test Cases
    search_test_cases = [
        ["TC07", "Search with valid term", "Search: 'smartphone'", "Relevant products displayed", "5 smartphone products displayed", "Pass"],
        ["TC08", "Search with no matches", "Search: 'nonexistentitem123'", "Empty results message", "Message: 'No products match'", "Pass"],
        ["TC09", "Category filtering", "Category: Electronics", "Electronics products", "15 Electronics products displayed", "Pass"],
        ["TC12", "Out of stock handling", "Product: 'Limited Edition'", "Disabled add button", "Incorrect behavior observed", "Fail"]
    ]
    add_test_case_slide(prs, "Product Search Test Cases", search_test_cases)
    
    # 11. Shopping Cart Test Cases
    cart_test_cases = [
        ["TC13", "Add to cart", "Product: Smartphone XYZ", "Product added correctly", "Added with price ₹15,999", "Pass"],
        ["TC14", "Update quantity", "Increase quantity to 3", "Price recalculated", "Total price ₹47,997", "Pass"],
        ["TC16", "Multiple items", "Smartphone, Headphones", "Correct items and total", "Both items with correct total", "Pass"],
        ["TC18", "Stock limitation", "Limited stock: 2, Request: 5", "System limits quantity", "Error: 'Only 2 available'", "Pass"]
    ]
    add_test_case_slide(prs, "Shopping Cart Test Cases", cart_test_cases)
    
    # 12. Security Section
    add_section_slide(prs, "Security Aspects")
    
    # 13. Security Measures
    security_measures = [
        "🔒 User Authentication: Password hashing, two-factor authentication, and secure session management",
        "🔒 Data Encryption: AES-256 encryption for sensitive data with secure key management",
        "🔒 Access Control: Role-based permissions with principle of least privilege",
        "🔒 PCI-DSS Compliance: For secure payment card handling",
        "🔒 GDPR Compliance: For personal data protection and privacy",
        "🔒 Regular Security Audits: Third-party penetration testing and vulnerability assessments",
        "🔒 Network Security: TLS 1.3, firewalls, and DDoS mitigation measures"
    ]
    add_content_slide(prs, "Security Implementation", security_measures)
    
    # 14. Deployment Section
    add_section_slide(prs, "Deployment Architecture")
    
    # 15. Deployment Details Two Column
    left_deployment = [
        "• High-availability cloud hosting",
        "• Load balancing with Nginx",
        "• Auto-scaling capabilities",
        "• Multiple application servers",
        "• Redundant database setup"
    ]
    
    right_deployment = [
        "• Master-slave DB replication",
        "• Redis for caching",
        "• Daily encrypted backups",
        "• ELK stack for monitoring",
        "• Automated deployment pipeline"
    ]
    add_two_column_slide(prs, "Deployment Details", left_deployment, right_deployment, 
                         "Infrastructure", "Maintenance & Monitoring")
    
    # 16. Future Scope Section
    add_section_slide(prs, "Future Scope")
    
    # 17. Future Enhancements
    future_scope = [
        "⟹ Mobile Applications: Native Android and iOS apps with offline capabilities",
        "⟹ AI-powered Recommendations: Personalized product suggestions using machine learning",
        "⟹ Voice Commerce: Integration with popular voice assistants",
        "⟹ Augmented Reality: Virtual product try-on experiences",
        "⟹ Internationalization: Multi-language and multi-currency support",
        "⟹ Subscription-based Models: Recurring billing capabilities",
        "⟹ Social Commerce: Direct integration with social media platforms"
    ]
    add_content_slide(prs, "Future Enhancements", future_scope)
    
    # 18. Thank You Slide
    slide_layout = prs.slide_layouts[0]  # Title Slide layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Set the slide background to a light blue color
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(235, 245, 255)
    
    # Add title
    title_shape = slide.shapes.title
    title_shape.text = "Thank You"
    
    # Format title
    title_text_frame = title_shape.text_frame
    title_text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    title_run = title_text_frame.paragraphs[0].runs[0]
    title_run.font.size = Pt(54)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(52, 152, 219)
    
    # Add contact information
    subtitle_shape = slide.placeholders[1]
    subtitle_shape.text = "For more information, please contact:\nemail@example.com"
    
    # Format subtitle
    subtitle_text_frame = subtitle_shape.text_frame
    subtitle_text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    subtitle_run = subtitle_text_frame.paragraphs[0].runs[0]
    subtitle_run.font.size = Pt(24)
    subtitle_run.font.color.rgb = RGBColor(44, 62, 80)
    
    # Save the presentation
    prs.save('ShopSleek_Visual_Presentation.pptx')
    print("Visual PowerPoint presentation created successfully!")

if __name__ == "__main__":
    create_visual_presentation()