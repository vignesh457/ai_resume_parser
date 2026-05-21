import io
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import qn, nsdecls

def add_hyperlink(paragraph, text, url, color="0B5394", underline=True):
    """Helper function to inject real, clickable hyperlinks into paragraphs"""
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')

    # Color Setup
    c = OxmlElement('w:color')
    c.set(qn('w:val'), color)
    rPr.append(c)

    # Underline Setup
    if underline:
        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'single')
        rPr.append(u)

    new_run.append(rPr)
    text_node = OxmlElement('w:t')
    text_node.text = text
    new_run.append(text_node)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink

def build_docx_binary(user_data):
    """Compiles structured JSON data into a precise replica of the original document format"""
    doc = docx.Document()
    
    # 0.5 inch margins all around for compact professional layout
    for section in doc.sections:
        section.page_width, section.page_height = Inches(8.5), Inches(11.0)
        section.top_margin = section.bottom_margin = Inches(0.5)
        section.left_margin = section.right_margin = Inches(0.5)

    # Explicit Color Standards
    COLOR_PRIMARY = RGBColor(11, 83, 148)  # Deep Professional Blue
    COLOR_TEXT = RGBColor(0, 0, 0)         # Stark Black

    def set_run_font(run, name='Calibri', size=10, bold=False, italic=False, color=COLOR_TEXT):
        run.font.name = name
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color

    def add_section_heading(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        
        run = p.add_run(text.upper())
        set_run_font(run, name='Calibri', size=10.5, bold=True, color=COLOR_PRIMARY)
        
        # Crisp blue border line directly underlying the text heading
        pBdr = parse_xml(r'<w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                         r'<w:bottom w:val="single" w:sz="6" w:space="1" w:color="0B5394"/>'
                         r'</w:pBdr>')
        p._p.get_or_add_pPr().append(pBdr)

    # ==================== HEADER SECTION ====================
    p_name = doc.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_name.paragraph_format.space_after = Pt(2)
    run_name = p_name.add_run(user_data.get('name', 'VIGNESHWAR REDDY DONAPATI').upper())
    set_run_font(run_name, name='Calibri', size=18, bold=True, color=COLOR_PRIMARY)

    # Contact Info Line
    p_contact = doc.add_paragraph()
    p_contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_contact.paragraph_format.space_after = Pt(2)
    
    phone = user_data.get('phone', '+91 9908731210')
    email = user_data.get('email', 'donativigneshwar@gmail.com')
    
    run_phone = p_contact.add_run(f"{phone} | ")
    set_run_font(run_phone, name='Calibri', size=9.5, color=COLOR_TEXT)
    add_hyperlink(p_contact, email, f"mailto:{email}", color="0B5394", underline=True)

    # Links Line
    p_links = doc.add_paragraph()
    p_links.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_links.paragraph_format.space_after = Pt(6)
    
    portfolio = user_data.get('portfolio', 'imvignesh.in')
    linkedin = user_data.get('linkedin', 'linkedin.com/in/vignesh457')
    github = user_data.get('github', 'github.com/vignesh457')

    add_hyperlink(p_links, portfolio, f"https://{portfolio}", color="0B5394", underline=True)
    r_sep1 = p_links.add_run(" | ")
    set_run_font(r_sep1, name='Calibri', size=9.5, color=COLOR_PRIMARY)
    
    add_hyperlink(p_links, linkedin, f"https://{linkedin}", color="0B5394", underline=True)
    r_sep2 = p_links.add_run(" | ")
    set_run_font(r_sep2, name='Calibri', size=9.5, color=COLOR_PRIMARY)
    
    add_hyperlink(p_links, github, f"https://{github}", color="0B5394", underline=True)


    # ==================== EXPERIENCE SECTION ====================
    if 'experience' in user_data and user_data['experience']:
        add_section_heading("Experience")
        for job in user_data['experience']:
            p_job = doc.add_paragraph()
            p_job.paragraph_format.space_before = Pt(4)
            p_job.paragraph_format.space_after = Pt(2)
            p_job.paragraph_format.keep_with_next = True
            
            # Left block text metadata
            r_title = p_job.add_run(f"{job.get('title', '')} | {job.get('company', '')}")
            set_run_font(r_title, name='Calibri', size=10, bold=True, color=COLOR_TEXT)
            
            # Right-aligned tab stop for dates
            p_job.paragraph_format.tab_stops.add_tab_stop(Inches(7.5), docx.enum.text.WD_TAB_ALIGNMENT.RIGHT)
            p_job.add_run("\t")
            r_dates = p_job.add_run(job.get('dates', ''))
            set_run_font(r_dates, name='Calibri', size=10, bold=True, color=COLOR_TEXT)

            for bullet in job.get('bullets', []):
                bp = doc.add_paragraph(style='List Bullet')
                bp.paragraph_format.space_after = Pt(2)
                bp.paragraph_format.line_spacing = 1.05
                
                # Removed keyword highlighting inside experience strings completely
                clean_bullet = bullet.replace("**", "")
                br = bp.add_run(clean_bullet)
                set_run_font(br, name='Calibri', size=10, bold=False, color=COLOR_TEXT)


    # ==================== SKILLS SECTION ====================
    if 'skills' in user_data and user_data['skills']:
        add_section_heading("Skills")
        skills_data = user_data['skills']
        
        if isinstance(skills_data, dict):
            for category, items in skills_data.items():
                p_sk = doc.add_paragraph()
                p_sk.paragraph_format.space_after = Pt(2)
                
                cat_run = p_sk.add_run(f"{category}: ")
                set_run_font(cat_run, name='Calibri', size=10, bold=True, color=COLOR_TEXT)
                
                items_str = ", ".join(items) if isinstance(items, list) else str(items)
                item_run = p_sk.add_run(items_str)
                set_run_font(item_run, name='Calibri', size=10, color=COLOR_TEXT)


    # ==================== PROJECTS SECTION ====================
    if 'projects' in user_data and user_data['projects']:
        add_section_heading("Projects")
        for proj in user_data['projects']:
            p_proj = doc.add_paragraph()
            p_proj.paragraph_format.space_before = Pt(4)
            p_proj.paragraph_format.space_after = Pt(2)
            p_proj.paragraph_format.keep_with_next = True
            
            r_pname = p_proj.add_run(f"{proj.get('title', '')} ")
            set_run_font(r_pname, name='Calibri', size=10, bold=True, color=COLOR_TEXT)
            
            # Handle clickable link integrations inside project layouts smoothly
            if proj.get('video_demo') or proj.get('link'):
                r_pipe = p_proj.add_run("| ")
                set_run_font(r_pipe, name='Calibri', size=10, color=COLOR_TEXT)
                
                if proj.get('video_demo'):
                    add_hyperlink(p_proj, "Video demo", proj.get('video_demo'), color="0B5394", underline=True)
                    if proj.get('link'):
                        r_space = p_proj.add_run(" | ")
                        set_run_font(r_space, name='Calibri', size=10, color=COLOR_TEXT)
                
                if proj.get('link'):
                    add_hyperlink(p_proj, "Link", proj.get('link'), color="0B5394", underline=True)
                
            for bullet in proj.get('bullets', []):
                bp = doc.add_paragraph(style='List Bullet')
                bp.paragraph_format.space_after = Pt(2)
                bp.paragraph_format.line_spacing = 1.05
                
                clean_bullet = bullet.replace("**", "")
                br = bp.add_run(clean_bullet)
                set_run_font(br, name='Calibri', size=10, bold=False, color=COLOR_TEXT)


    # ==================== CERTIFICATIONS ====================
    if 'certifications' in user_data and user_data['certifications']:
        add_section_heading("Certifications")
        for cert in user_data['certifications']:
            p_cert = doc.add_paragraph(style='List Bullet')
            p_cert.paragraph_format.space_after = Pt(2)
            
            if isinstance(cert, dict):
                title = cert.get('title', '')
                url = cert.get('link', '')
                r_ctitle = p_cert.add_run(f"{title} ")
                set_run_font(r_ctitle, name='Calibri', size=10, color=COLOR_TEXT)
                if url:
                    p_cert.add_run("| ")
                    add_hyperlink(p_cert, "Certificate", url, color="0B5394", underline=True)
            else:
                r_text = p_cert.add_run(str(cert))
                set_run_font(r_text, name='Calibri', size=10, color=COLOR_TEXT)


    # ==================== EDUCATION SECTION ====================
    if 'education' in user_data and user_data['education']:
        add_section_heading("Education")
        for edu in user_data['education']:
            p_edu = doc.add_paragraph()
            p_edu.paragraph_format.space_after = Pt(1)
            p_edu.paragraph_format.keep_with_next = True
            
            r_inst = p_edu.add_run(f"{edu.get('institution', '')}")
            set_run_font(r_inst, name='Calibri', size=10, bold=True, color=COLOR_TEXT)
            
            p_edu.paragraph_format.tab_stops.add_tab_stop(Inches(7.5), docx.enum.text.WD_TAB_ALIGNMENT.RIGHT)
            p_edu.add_run("\t")
            r_edates = p_edu.add_run(edu.get('dates', ''))
            set_run_font(r_edates, name='Calibri', size=10, bold=True, color=COLOR_TEXT)
            
            p_det = doc.add_paragraph()
            p_det.paragraph_format.space_after = Pt(3)
            
            degree_str = f"{edu.get('degree', '')} in {edu.get('field', '')} | GPA: {edu.get('gpa', '')}"
            r_det = p_det.add_run(degree_str)
            # Standardized layout font styling to uniform Calibri
            set_run_font(r_det, name='Calibri', size=10, color=COLOR_TEXT)


    # ==================== CODING PROFILES ====================
    if 'coding_profiles' in user_data and user_data['coding_profiles']:
        add_section_heading("Coding Profiles")
        p_cp = doc.add_paragraph()
        p_cp.paragraph_format.space_after = Pt(3)
        
        profiles = user_data['coding_profiles']
        if isinstance(profiles, list):
            # Flat row format assembly
            for idx, item in enumerate(profiles):
                r_item = p_cp.add_run(item)
                set_run_font(r_item, name='Calibri', size=10, color=COLOR_TEXT)
                if idx < len(profiles) - 1:
                    r_pipe = p_cp.add_run(" | ")
                    set_run_font(r_pipe, name='Calibri', size=10, color=COLOR_TEXT)
        else:
            r_text = p_cp.add_run(str(profiles))
            set_run_font(r_text, name='Calibri', size=10, color=COLOR_TEXT)

    target_stream = io.BytesIO()
    doc.save(target_stream)
    return target_stream.getvalue().hex()