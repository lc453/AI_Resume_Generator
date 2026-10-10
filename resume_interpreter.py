from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn
from validators import url, email
from docx.opc import constants
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from markdown_it import MarkdownIt

import docx

def main() -> None:
    save_name = "cvetko_luke_resume.docx"
    resume = Document()

    vert_margin: float = 0.5
    horiz_margin: float = 0.5
    line_spacing: float = 1.1

    name: str = "Luke Cvetko"

    # setting the page margins
    for section in resume.sections:
        section.top_margin = Inches(vert_margin)
        section.bottom_margin = Inches(vert_margin)
        section.left_margin = Inches(horiz_margin)
        section.right_margin = Inches(horiz_margin)

    body_font: str = 'Garamond'
    '''styles: dict[str, list[str|int|bool]] = {
        'Resume Normal': [body_font, 12],
        'Resume Title': [body_font, 28],
        'Resume Subtitle': [body_font, 13],
        'Resume Heading 1': [body_font, 14]
    }'''
    
    styles: dict[str, dict[str,str|int]] = {
        'Resume Normal': {
            "font": body_font,
            "size": 12,
            "bold": None,
            "ital": None,
            "underline": None
        },
        'Resume Title': {
            "font": body_font,
            "size": 24,
            "bold": True,
            "ital": False,
            "underline": False
        },
        'Resume Subtitle': {
            "font": body_font,
            "size": 13,
            "bold": False,
            "ital": False,
            "underline": None
        },
        'Resume Heading 1': {
            "font": body_font,
            "size": 14,
            "bold": True,
            "ital": False,
            "underline": False
        },
        'Resume Heading 2': {
            "font": body_font,
            "size": 12,
            "bold": True,
            "ital": False,
            "underline": False
        },
        'Resume Heading 3': {
            "font": body_font,
            "size": 12,
            "bold": False,
            "ital": True,
            "underline": False
        },
        'Resume Section Break': {
            "font": body_font,
            "size": 4,
            "bold": None,
            "ital": None,
            "underline": None
        }
    }
    define_styles(resume, styles)

    generate_header(resume, name, ["Highland, UT", "(385) 219-6631", "cvetkoluke54@gmail.com", "https://linkedin.com/in/cvetko_luke"])

    professional_summary = "Software Engineer with MS in physics and 1+ years of professional experience building production-level web applications at Epic Systems. 3+ years in scientific software development, including physical simulations, hardware integration, and data processing."

    generate_text(resume, professional_summary)

    generate_experience_section(resume, "education", {})

    work_experience = [
        {
            'institution': 'University of Wisconsin - Madison',
            'dates': 'Dec. 2024 - Dec. 2025',
            'subtitle': 'Research Associate, Soley Lab & Song Lab',
            'location': 'Madison, WI',
            'bullets': [
                'Developed a Python workflow to automate a motorized microscope imaging system for sample analysis',
                'Implemented multithreading to simultaneously communicate with the digital camera and individual microscope components',
                'Built pipelines to organize and analyze image data collected from automated microscope systems',
                'Contributed to a quantum algorithm estimating molecular bond strength from interatomic potentials'
            ]
        }
    ]

    generate_experience_section(resume, "work experience", work_experience)

    skills = [
        "**Programming**: C#, TypeScript, Python, Julia, Bash scripting, object-oriented programming, threading",
        "**Scientific Computing**: MATLAB, molecular simulation, high-performance computing, Linux",
        "**Professional**: Technical communication & documentation, collaboration, MS Office"
    ]

    interests = [
        "Tormenting kittens, eating soup, being nonsense"
    ]

    generate_simple_section(resume, "skills", skills)

    generate_simple_section(resume, "interests", interests)

    for paragraph in resume.paragraphs:
        paragraph.paragraph_format.line_spacing = line_spacing
        paragraph.paragraph_format.space_after = Pt(0)

    delete_paragraph(resume.paragraphs[-1]) 

    resume.save(save_name)

def set_margins(document: Document, v_margin: float, h_margin: float):
    """
    set_doc_params sets the margins and line spacing of the document
    """
    
    for section in document.sections:
        section.top_margin = Inches(v_margin)
        section.bottom_margin = Inches(v_margin)
        section.left_margin = Inches(h_margin)
        section.right_margin = Inches(h_margin)

def set_spacing(document: Document, line_spacing: float):
    """
    set_spacing sets the line spacing of the entire document
    """

    for paragraph in document.paragraphs:
        paragraph.paragraph_format.line_spacing = line_spacing
        paragraph.paragraph_format.space_after = Pt(0)

def save_document(document: Document, filepath: str, cutlines: int = 0):
    for i in range(cutlines):
        delete_paragraph(document.paragraphs[-1]) 

    document.save(filepath)

    

def define_styles(document: Document, style_dict: dict[str, dict]) -> None:
    """
    define_styles defines the styles to be used in the document by font and text size
    :param document: the document the main program is writing to, which we are setting styles for
    :param style_dict: dictionary that contains the style formatting for each style. Entries have format
    'key': [font name, font size], where key is the style name, i.e. 'Normal', 'Title', 'Header', etc
    """
    for key in style_dict:
        style = document.styles.add_style(key, docx.enum.style.WD_STYLE_TYPE.PARAGRAPH)
        font = style.font
        config = style_dict[key]
        font.name = config['font']
        font.size = Pt(config['size'])
        font.bold = config['bold']
        font.italic = config['ital']
        font.underline = config['underline']
        if 'normal' in key.lower():
            style = document.styles['Normal']
            font = style.font
            config = style_dict[key]
            font.name = config['font']
            font.size = Pt(config['size'])
            font.bold = config['bold']
            font.italic = config['ital']
            font.underline = config['underline']


def generate_header(document: Document, name: str, contact_info: list[str]) -> None:
    """
    generate_header creates the header of the document, made up of the user name as a title
    and the contact info as a subtitle
    :param document: the document to create a header for
    :param name: string containing the user name (first name + last name)
    :param contact info: list of strings containing the user contact information
    """
    # Add the name
    p_title = document.add_paragraph(name)
    p_title.style = document.styles['Resume Title']
    p_subtitle = document.add_paragraph("")
    for info in contact_info:
        if url(info):
            add_hyperlink(p_subtitle, info.split("//")[-1], info)
        else:
            p_subtitle.add_run(info)
        separator = p_subtitle.add_run(' | ')
    cutoff = p_subtitle.runs[-1]._element
    cutoff.getparent().remove(cutoff)
    p_subtitle.style = document.styles['Resume Subtitle']
    p_subtitle.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    #insertHR(p_subtitle)
    insert_empty_line(document, True)

def generate_name(document: Document, name:str) -> None:
    """
    generate_name creates the name on top of the document
    :param document: the document to create a header for
    :param name: string containing the user name (first name + last name)
    """
    # Add the name
    p_title = document.add_paragraph(name)
    p_title.style = document.styles['Resume Title']

def generate_contact_info(document: Document, contact_info: list[str]) -> None:
    """
    generate_contact_info creates the contact info on top of the document under the name
    :param document: the document to create a header for
    :param contact info: list of strings containing the user contact information
    """
    p_subtitle = document.add_paragraph("")
    for info in contact_info:
        if url(info):
            add_hyperlink(p_subtitle, info.split("//")[-1], info)
        else:
            p_subtitle.add_run(info)
        separator = p_subtitle.add_run(' | ')
    cutoff = p_subtitle.runs[-1]._element
    cutoff.getparent().remove(cutoff)
    p_subtitle.style = document.styles['Resume Subtitle']
    p_subtitle.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    #insertHR(p_subtitle)
    insert_empty_line(document, True)

def generate_text(document: Document, text: str) -> None:
    """
    generate_text inserts a section of unformatted plaintext into the document
    :param document: document into which the text will be inserted
    :param text: string containing the text to add
    """
    p_text = document.add_paragraph(text)
    #insertHR(p_text)
    p_text.style = document.styles['Resume Normal']
    insert_empty_line(document, False)

def generate_experience_section(document: Document, section_title: str, experience_list: list[dict]) -> None:
    """
    generate_experience_section is the meat of the resume, responsible for creating the experience and education sections
    Within the section, it has room for a list of experience items to add to the section.
    These experience items can be potentially complex, with titles, subtitles, dates, and bullets
    :param document: document to which the section will be added
    :param section_title: should be either "Work experience" or "education"
    :param experience_list: A list of dicts, each dict containing all the relevant information on the particular experience
    experience dict:
        - "institution": either the school attended or company worked for
        - "subtitle": job title or degree information
        - "dates": string containing the relevant date range
        - "location": string containing the location (might be None for remote work)
        - "bullets": A list of bullets to be put in the section
    """
    p_section = document.add_paragraph(section_title.upper())
    p_section.style = document.styles['Resume Heading 1']
    insert_empty_line(document, True)

    # get the page width for tab stuff later on
    section = document.sections[0]

    usable_width = section.page_width - section.left_margin - section.right_margin



    for experience in experience_list:
        # First output the institution with the date range. 
        p_inst = document.add_paragraph()
        p_inst.paragraph_format.tab_stops.add_tab_stop(usable_width, docx.enum.text.WD_TAB_ALIGNMENT.RIGHT)
        p_inst.add_run(experience['institution'])
        p_inst.add_run('\t')
        p_inst.add_run(experience['dates'])
        p_inst.style = document.styles['Resume Heading 2']

        # Next, output the subtitle and location (if location exists)
        p_sub = document.add_paragraph()
        p_sub.add_run(experience['subtitle'])
        if 'location' in experience:
            p_sub.paragraph_format.tab_stops.add_tab_stop(usable_width, docx.enum.text.WD_TAB_ALIGNMENT.RIGHT)
            p_sub.add_run('\t')
            p_sub.add_run(experience['location'])
            p_sub.style = document.styles['Resume Heading 3']

        # Finally, if the experience has bullets, add the bullet points
        dash_num_id = add_dash_bullet_numbering(document)
        if 'bullets' in experience:
            for bullet in experience['bullets']:
                p_bullet = document.add_paragraph(bullet)
                p_bullet.style = document.styles['Resume Normal']
                # apply_numbering(p_bullet, square_num_id)
                apply_numbering(p_bullet, dash_num_id)
                #p_bullet = document.add_paragraph(bullet, style='List Bullet')
                #p_bullet.latent_style = document.styles['List Bullet']
    insert_empty_line(document, False)
        
def generate_simple_section(document: Document, section_title: str, items: list[str]):
    """
    generate_simple_section creates a simple section comprised only of a title and bulleted list
    :param document: The document to add the simple section to
    :param section_title: The title for this resume section
    :param items: The list of items to add in a bulleted list. These will be in markdown, and will have to be parsed as such
    """
    document.add_paragraph(section_title.upper(), style = 'Resume Heading 1')
    insert_empty_line(document, True)
    
    dash_num_id = add_dash_bullet_numbering(document)
    for item in items:
        p_bullet = document.add_paragraph(style = 'Resume Normal')
        add_markdown_paragraph(item, p_bullet)
        apply_numbering(p_bullet, dash_num_id)
    insert_empty_line(document, False)

def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None
#region FORMATTING HELPER FUNCTIONS
def insert_empty_line(document, add_HR: bool) -> None:
    """
    insert_empty_line creates an empty line of whitespace
    :param document: Document object to add the whitespace to
    :param add_HR" bool that controls whether or not a horizontal rule is input as part of the empty line
    """
    p = document.add_paragraph()
    p.style = document.styles['Resume Section Break']
    if add_HR:
        insertHR(p)

def insertHR(paragraph: docx.text.paragraph.Paragraph) -> None:
    """
    insertHR inserts a horizontal rule directly after the paragraph
    :param paragraph: the paragraph after which to insert the horizontal rule
    """
    p = paragraph._p  # p is the <w:p> XML element
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    pPr.insert_element_before(pBdr,
        'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap',
        'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN',
        'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind',
        'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc',
        'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap',
        'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr',
        'w:pPrChange'
    )
    bottom = OxmlElement('w:top')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'auto')
    pBdr.append(bottom)


def add_hyperlink(paragraph: docx.text.paragraph.Paragraph, text: str, url: str) -> None:
    """
    add_hyperlink adds a hyperlink to the paragraph formatted as a standard hyperlink
    :param paragraph: paragraph to which the hyperlink is being appended
    :param text: text to show for the hyperlink
    :param url: url that the hyperlink directs to
    """
    # This gets access to the document.xml.rels file and gets a new relation id value
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    # Create the w:hyperlink tag and add needed values
    hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
    hyperlink.set(docx.oxml.shared.qn('r:id'), r_id, )

    # Create a new run object (a wrapper over a 'w:r' element)
    new_run = docx.text.run.Run(
        docx.oxml.shared.OxmlElement('w:r'), paragraph)
    new_run.text = text

    # Set the run's style to the builtin hyperlink style, defining it if necessary
    new_run.style = get_or_create_hyperlink_style(part.document)
    # Alternatively, set the run's formatting explicitly
    # new_run.font.color.rgb = docx.shared.RGBColor(0, 0, 255)
    # new_run.font.underline = True

    # Join all the xml elements together
    hyperlink.append(new_run._element)
    paragraph._p.append(hyperlink)
    return hyperlink

def get_or_create_hyperlink_style(d):
    """If this document had no hyperlinks so far, the builtin
       Hyperlink style will likely be missing and we need to add it.
       There's no predefined value, different Word versions
       define it differently.
       This version is how Word 2019 defines it in the
       default theme, excluding a theme reference.
    """
    if "Hyperlink" not in d.styles:
        if "Default Character Font" not in d.styles:
            ds = d.styles.add_style("Default Character Font",
                                    docx.enum.style.WD_STYLE_TYPE.CHARACTER,
                                    True)
            ds.element.set(docx.oxml.shared.qn('w:default'), "1")
            ds.priority = 1
            ds.hidden = True
            ds.unhide_when_used = True
            del ds
        hs = d.styles.add_style("Hyperlink",
                                docx.enum.style.WD_STYLE_TYPE.CHARACTER,
                                True)
        hs.base_style = d.styles["Default Character Font"]
        hs.unhide_when_used = True
        hs.font.color.rgb = docx.shared.RGBColor(0x05, 0x63, 0xC1)
        hs.font.underline = True
        del hs

    return "Hyperlink"


def add_dash_bullet_numbering(document):
    numbering = document.part.numbering_part.element

    # Find the next available abstractNumId and numId
    abstract_ids = [
        int(el.get(qn("w:abstractNumId")))
        for el in numbering.findall(qn("w:abstractNum"))
    ]
    num_ids = [
        int(el.get(qn("w:numId")))
        for el in numbering.findall(qn("w:num"))
    ]

    abstract_num_id = max(abstract_ids, default=0) + 1
    num_id = max(num_ids, default=0) + 1

    # Create abstract numbering definition
    abstract_num = OxmlElement("w:abstractNum")
    abstract_num.set(qn("w:abstractNumId"), str(abstract_num_id))

    multi_level_type = OxmlElement("w:multiLevelType")
    multi_level_type.set(qn("w:val"), "singleLevel")
    abstract_num.append(multi_level_type)

    level = OxmlElement("w:lvl")
    level.set(qn("w:ilvl"), "0")

    start = OxmlElement("w:start")
    start.set(qn("w:val"), "1")
    level.append(start)

    num_fmt = OxmlElement("w:numFmt")
    num_fmt.set(qn("w:val"), "bullet")
    level.append(num_fmt)

    # bullet character
    lvl_text = OxmlElement("w:lvlText")
    lvl_text.set(qn("w:val"), "-")
    level.append(lvl_text)

    # Controls bullet/text spacing
    level_p_pr = OxmlElement("w:pPr")

    indent = OxmlElement("w:ind")
    indent.set(qn("w:left"), "360")
    indent.set(qn("w:hanging"), "180")
    level_p_pr.append(indent)

    tab = OxmlElement("w:tabs")
    tab_stop = OxmlElement("w:tab")
    tab_stop.set(qn("w:val"), "num")
    tab_stop.set(qn("w:pos"), "360")
    tab.append(tab_stop)
    level_p_pr.append(tab)

    level.append(level_p_pr)

    # Use a standard bullet font
    lvl_font = OxmlElement("w:rPr")
    r_fonts = OxmlElement("w:rFonts")
    r_fonts.set(qn("w:ascii"), "Arial")
    r_fonts.set(qn("w:hAnsi"), "Arial")
    lvl_font.append(r_fonts)
    level.append(lvl_font)

    abstract_num.append(level)
    numbering.append(abstract_num)

    # Create numbering instance
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))

    abstract_num_ref = OxmlElement("w:abstractNumId")
    abstract_num_ref.set(qn("w:val"), str(abstract_num_id))
    num.append(abstract_num_ref)

    numbering.append(num)

    return num_id


def apply_numbering(paragraph, num_id, level=0):
    p_pr = paragraph._p.get_or_add_pPr()

    num_pr = p_pr.find(qn("w:numPr"))
    if num_pr is None:
        num_pr = OxmlElement("w:numPr")
        p_pr.append(num_pr)

    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), str(level))
    num_pr.append(ilvl)

    num_id_element = OxmlElement("w:numId")
    num_id_element.set(qn("w:val"), str(num_id))
    num_pr.append(num_id_element)


def add_markdown_paragraph(text:str, paragraph):
    """
    Add text containing Markdown bold and italic formatting
    to a python-docx paragraph.

    Supported:
        **bold**
        __bold__
        *italic*
        _italic_
        ***bold italic***
        ___bold italic___
    """
    if isinstance(text, dict):
        return text
    md = MarkdownIt("commonmark")
    parsed = md.parseInline(text)

    # The inline tokens are stored in the first inline token.
    children = parsed[0].children or []

    bold = False
    italic = False

    for token in children:
        if token.type == "strong_open":
            bold = True

        elif token.type == "strong_close":
            bold = False

        elif token.type == "em_open":
            italic = True

        elif token.type == "em_close":
            italic = False

        elif token.type == "text":
            run = paragraph.add_run(token.content)
            run.bold = bold
            run.italic = italic

        elif token.type in {"softbreak", "hardbreak"}:
            paragraph.add_run("\n")

        elif token.type == "code_inline":
            # Treat inline code as ordinary unformatted text.
            paragraph.add_run(token.content)

    return paragraph
#endregion




if __name__ == "__main__":
    main()