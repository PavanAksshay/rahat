import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# =============================================================================
# COLOR PALETTE (Exact match to web design system)
# =============================================================================
COLOR_IVORY = RGBColor(247, 242, 234)        # #F7F2EA
COLOR_IVORY_WARM = RGBColor(250, 246, 240)   # #FAF6F0
COLOR_NAVY_DEEP = RGBColor(8, 13, 28)        # #080D1C
COLOR_NAVY_CARD = RGBColor(24, 34, 56)       # #182238
COLOR_RAHAT_RED = RGBColor(231, 25, 63)      # #E7193F
COLOR_WHITE = RGBColor(255, 255, 255)        # #FFFFFF
COLOR_SUBTLE_GOLD = RGBColor(214, 168, 93)   # #D6A85D
COLOR_MUTED_TEXT = RGBColor(100, 105, 120)   # Secondary muted
COLOR_BORDER_LIGHT = RGBColor(225, 220, 212) # Clean border

FONT_SERIF = "Georgia"
FONT_SANS = "Arial"

def set_slide_background(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

def create_textbox(slide, left, top, width, height):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.04)
    tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.04)
    tf.margin_bottom = Inches(0.04)
    return tf

def add_eyebrow(slide, text, left, top, color=COLOR_RAHAT_RED):
    tf = create_textbox(slide, left, top, Inches(8.0), Inches(0.35))
    p = tf.paragraphs[0]
    p.text = "—  " + text.upper()
    p.font.name = FONT_SANS
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = color
    return tf

def build_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # =========================================================================
    # SLIDE 01: HERO OPENING
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_background(s1, COLOR_IVORY)
    
    # Hero Logo (Large, prominent)
    if os.path.exists('assets/logos/rahat_logo.jpeg'):
        s1.shapes.add_picture('assets/logos/rahat_logo.jpeg', Inches(5.166), Inches(0.8), width=Inches(3.0), height=Inches(3.0))
    
    tf1 = create_textbox(s1, Inches(1.0), Inches(3.9), Inches(11.333), Inches(3.2))
    
    p = tf1.paragraphs[0]
    p.text = "ESTABLISHED 2023  •  MIT BENGALURU"
    p.font.name = FONT_SANS
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_MUTED_TEXT
    p.alignment = PP_ALIGN.CENTER
    p.space_after = Pt(8)
    
    p = tf1.add_paragraph()
    p.text = "RAHAT"
    p.font.name = FONT_SERIF
    p.font.size = Pt(64)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DEEP
    p.alignment = PP_ALIGN.CENTER
    p.space_after = Pt(12)
    
    p = tf1.add_paragraph()
    p.text = "A student-led club and media house focused on mental well-being."
    p.font.name = FONT_SERIF
    p.font.size = Pt(18)
    p.font.italic = True
    p.font.color.rgb = COLOR_NAVY_DEEP
    p.alignment = PP_ALIGN.CENTER
    p.space_after = Pt(14)

    p = tf1.add_paragraph()
    p.text = "MAKING SPACE  •  Rahat to all."
    p.font.name = FONT_SANS
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_RAHAT_RED
    p.alignment = PP_ALIGN.CENTER

    # =========================================================================
    # SLIDE 02: THE PREMISE
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_background(s2, COLOR_IVORY)
    add_eyebrow(s2, "Act I — Making Space", Inches(0.8), Inches(0.45))
    
    # Left Visuals (2 photos from booth outreach & classroom)
    if os.path.exists('assets/08_RAHAT_BOOTH_OUTREACH/rahat_booth_students_01.jpeg'):
        s2.shapes.add_picture('assets/08_RAHAT_BOOTH_OUTREACH/rahat_booth_students_01.jpeg', Inches(0.8), Inches(1.0), width=Inches(3.6), height=Inches(5.8))
    if os.path.exists('assets/08_RAHAT_BOOTH_OUTREACH/classroom_talk.jpeg'):
        s2.shapes.add_picture('assets/08_RAHAT_BOOTH_OUTREACH/classroom_talk.jpeg', Inches(4.55), Inches(1.0), width=Inches(2.4), height=Inches(5.8))
    
    # Right Editorial Text
    tf2 = create_textbox(s2, Inches(7.35), Inches(1.3), Inches(5.2), Inches(5.4))
    
    p = tf2.paragraphs[0]
    p.text = "THE PREMISE"
    p.font.name = FONT_SANS
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_RAHAT_RED
    p.space_after = Pt(12)
    
    p = tf2.add_paragraph()
    p.text = "What does feeling okay look like?"
    p.font.name = FONT_SERIF
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DEEP
    p.space_after = Pt(18)
    
    p = tf2.add_paragraph()
    p.text = "“Sometimes, it doesn’t look like a formal conversation about mental health.”"
    p.font.name = FONT_SERIF
    p.font.size = Pt(20)
    p.font.italic = True
    p.font.color.rgb = COLOR_NAVY_DEEP
    p.space_after = Pt(18)
    
    p = tf2.add_paragraph()
    p.text = "Sometimes, it simply looks like finding a space to be."
    p.font.name = FONT_SANS
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLOR_RAHAT_RED

    # =========================================================================
    # SLIDE 03: RAHAT — THE EVOLUTION (GIANT LOGOS)
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_background(s3, COLOR_IVORY)
    add_eyebrow(s3, "Act II — Evolution of Identity", Inches(0.8), Inches(0.4))

    tf_s3_h = create_textbox(s3, Inches(0.8), Inches(0.75), Inches(11.7), Inches(0.55))
    p = tf_s3_h.paragraphs[0]
    p.text = "FROM WHERE WE STARTED  ➔  TO WHERE WE ARE GOING"
    p.font.name = FONT_SERIF
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DEEP

    # Left Card: Old Logo (Huge)
    card_old = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(5.1), Inches(4.5))
    card_old.fill.solid()
    card_old.fill.fore_color.rgb = COLOR_WHITE
    card_old.line.color.rgb = COLOR_BORDER_LIGHT
    
    tf_o_tag = create_textbox(s3, Inches(0.8), Inches(1.5), Inches(5.1), Inches(0.35))
    p = tf_o_tag.paragraphs[0]
    p.text = "THEN  •  THE FOUNDATION (ESTD. 2023)"
    p.font.name = FONT_SANS
    p.font.size = Pt(11)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = COLOR_MUTED_TEXT

    if os.path.exists('assets/logos/rahat_old_logo.jpeg'):
        s3.shapes.add_picture('assets/logos/rahat_old_logo.jpeg', Inches(1.6), Inches(1.9), width=Inches(3.5), height=Inches(3.5))
    
    tf_o_b = create_textbox(s3, Inches(0.8), Inches(5.5), Inches(5.1), Inches(0.35))
    p = tf_o_b.paragraphs[0]
    p.text = "STUDENT WELL-BEING COMMUNITY"
    p.font.name = FONT_SANS
    p.font.size = Pt(11)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = COLOR_NAVY_DEEP

    # Center Bridge Arrow
    tf_bridge = create_textbox(s3, Inches(5.9), Inches(3.0), Inches(1.5), Inches(1.2))
    p = tf_bridge.paragraphs[0]
    p.text = "➔"
    p.font.name = FONT_SANS
    p.font.size = Pt(36)
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = COLOR_RAHAT_RED
    p = tf_bridge.add_paragraph()
    p.text = "EXPANSION"
    p.font.name = FONT_SANS
    p.font.size = Pt(9)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = COLOR_RAHAT_RED

    # Right Card: New Logo (Huge)
    card_new = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.433), Inches(1.4), Inches(5.1), Inches(4.5))
    card_new.fill.solid()
    card_new.fill.fore_color.rgb = COLOR_WHITE
    card_new.line.color.rgb = COLOR_RAHAT_RED
    
    tf_n_tag = create_textbox(s3, Inches(7.433), Inches(1.5), Inches(5.1), Inches(0.35))
    p = tf_n_tag.paragraphs[0]
    p.text = "NOW  •  THE NEW CHAPTER"
    p.font.name = FONT_SANS
    p.font.size = Pt(11)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = COLOR_RAHAT_RED

    if os.path.exists('assets/logos/rahat_logo.jpeg'):
        s3.shapes.add_picture('assets/logos/rahat_logo.jpeg', Inches(8.233), Inches(1.9), width=Inches(3.5), height=Inches(3.5))

    tf_n_b = create_textbox(s3, Inches(7.433), Inches(5.5), Inches(5.1), Inches(0.35))
    p = tf_n_b.paragraphs[0]
    p.text = "STUDENT CLUB & MEDIA HOUSE"
    p.font.name = FONT_SANS
    p.font.size = Pt(11)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = COLOR_RAHAT_RED

    # Manifesto Footer
    tf_s3_f = create_textbox(s3, Inches(0.8), Inches(6.05), Inches(11.7), Inches(1.1))
    p = tf_s3_f.paragraphs[0]
    p.text = "“Same purpose. A wider canvas.”"
    p.font.name = FONT_SERIF
    p.font.size = Pt(19)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = COLOR_NAVY_DEEP
    p = tf_s3_f.add_paragraph()
    p.text = "From a student-led mental well-being community to a space for conversations, creativity, connection and stories that matter."
    p.font.name = FONT_SERIF
    p.font.size = Pt(13)
    p.font.italic = True
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = COLOR_MUTED_TEXT

    # =========================================================================
    # SLIDE 04: WHAT WE'VE DONE (CREATING SPACE - 5 EVENT PHOTOS)
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_background(s4, COLOR_IVORY)
    add_eyebrow(s4, "Act II — What We Have Built", Inches(0.8), Inches(0.4))
    
    tf_s4_h = create_textbox(s4, Inches(0.8), Inches(0.72), Inches(11.7), Inches(0.5))
    p = tf_s4_h.paragraphs[0]
    p.text = "“We started by creating space for expression, creativity and play.”"
    p.font.name = FONT_SERIF
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DEEP

    pics_s4 = [
        ('assets/08_RAHAT_BOOTH_OUTREACH/booth_polaroids_table.jpeg', 'POLAROIDS & NOTES'),
        ('assets/02_GRATITUDE_FEELINGS/feelings_board_activity.jpeg', 'FEELINGS BOARD'),
        ('assets/03_MEMORY_BOX/memory_box_activity.png', 'MEMORY BOX'),
        ('assets/06_GAMES_TREASURE_HUNT/treasure_hunt_crowd_01.jpeg', 'TREASURE HUNT'),
        ('assets/05_TECHSOLSTICE_NEXUS/techsolstice_nexus_group_activity.jpeg', 'COMMUNITY SPACE')
    ]
    col_w = Inches(2.23)
    col_gap = Inches(0.14)
    start_x = Inches(0.8)
    y_pos = Inches(1.35)
    h_pos = Inches(5.6)
    
    for i, (path, tag) in enumerate(pics_s4):
        x = start_x + i * (col_w + col_gap)
        if os.path.exists(path):
            s4.shapes.add_picture(path, x, y_pos, width=col_w, height=h_pos)
            
        tbox = create_textbox(s4, x, y_pos + h_pos - Inches(0.45), col_w, Inches(0.4))
        p = tbox.paragraphs[0]
        p.text = tag
        p.font.name = FONT_SANS
        p.font.size = Pt(9)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE
        p.alignment = PP_ALIGN.CENTER

    # =========================================================================
    # SLIDE 05: WHAT COMES NEXT (EXPANDING THE CANVAS)
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_background(s5, COLOR_IVORY)
    add_eyebrow(s5, "Act III — The Horizon", Inches(0.8), Inches(0.4))

    # Left Box
    box_l = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.85), Inches(5.6), Inches(6.1))
    box_l.fill.solid()
    box_l.fill.fore_color.rgb = COLOR_WHITE
    box_l.line.color.rgb = COLOR_BORDER_LIGHT
    
    tf_l = create_textbox(s5, Inches(1.0), Inches(0.95), Inches(5.2), Inches(0.5))
    p = tf_l.paragraphs[0]
    p.text = "WE STARTED BY CREATING SPACE"
    p.font.name = FONT_SERIF
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DEEP

    if os.path.exists('assets/04_MANOVRITTI/manovritti_group_photo.jpeg'):
        s5.shapes.add_picture('assets/04_MANOVRITTI/manovritti_group_photo.jpeg', Inches(1.0), Inches(1.55), width=Inches(3.0), height=Inches(4.6))
    if os.path.exists('assets/02_GRATITUDE_FEELINGS/feelings_board_activity.jpeg'):
        s5.shapes.add_picture('assets/02_GRATITUDE_FEELINGS/feelings_board_activity.jpeg', Inches(4.1), Inches(1.55), width=Inches(2.1), height=Inches(2.2))
    if os.path.exists('assets/08_RAHAT_BOOTH_OUTREACH/rahat_booth_students_02.jpeg'):
        s5.shapes.add_picture('assets/08_RAHAT_BOOTH_OUTREACH/rahat_booth_students_02.jpeg', Inches(4.1), Inches(3.9), width=Inches(2.1), height=Inches(2.25))

    # Right Box
    box_r = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(0.85), Inches(5.6), Inches(6.1))
    box_r.fill.solid()
    box_r.fill.fore_color.rgb = COLOR_WHITE
    box_r.line.color.rgb = COLOR_RAHAT_RED
    
    tf_r = create_textbox(s5, Inches(7.1), Inches(0.95), Inches(5.2), Inches(0.5))
    p = tf_r.paragraphs[0]
    p.text = "NOW, WE'RE MAKING THE SPACE BIGGER"
    p.font.name = FONT_SERIF
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = COLOR_RAHAT_RED

    pillars = [
        ("01", "RAHAT AS A MEDIA HOUSE", "Conversations with people who create social impact."),
        ("02", "RAHAT CARNIVAL", "A day of games, activities, creativity and connection."),
        ("03", "OPEN THEATRE", "Movies, conversations and shared experiences under one roof."),
        ("04", "CONTINUING THE CORE", "Mental well-being remains at the heart of everything we do.")
    ]
    
    for j, (num, title, desc) in enumerate(pillars):
        p_box = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.1), Inches(1.55 + j * 1.2), Inches(5.2), Inches(1.05))
        p_box.fill.solid()
        p_box.fill.fore_color.rgb = COLOR_IVORY
        p_box.line.color.rgb = COLOR_BORDER_LIGHT
        
        tf_p = create_textbox(s5, Inches(7.2), Inches(1.6 + j * 1.2), Inches(5.0), Inches(0.95))
        p1 = tf_p.paragraphs[0]
        p1.text = f"{num}  •  {title}"
        p1.font.name = FONT_SANS
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_NAVY_DEEP
        p1.space_after = Pt(3)
        
        p2 = tf_p.add_paragraph()
        p2.text = desc
        p2.font.name = FONT_SANS
        p2.font.size = Pt(10)
        p2.font.color.rgb = COLOR_MUTED_TEXT

    # =========================================================================
    # SLIDE 06: CINEMATIC SCALE (EMBEDDED PLAYABLE VIDEOS)
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_background(s6, COLOR_NAVY_DEEP)
    add_eyebrow(s6, "Act IV — Scale & Momentum", Inches(0.8), Inches(0.4), color=COLOR_RAHAT_RED)

    # Left Stat Block
    tf6 = create_textbox(s6, Inches(0.8), Inches(0.9), Inches(4.3), Inches(3.2))
    p = tf6.paragraphs[0]
    p.text = "700+"
    p.font.name = FONT_SERIF
    p.font.size = Pt(64)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE
    p.space_after = Pt(2)
    
    p = tf6.add_paragraph()
    p.text = "REGISTRATIONS"
    p.font.name = FONT_SANS
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLOR_RAHAT_RED
    p.space_after = Pt(4)

    p = tf6.add_paragraph()
    p.text = "ONE PACKED HOUSE"
    p.font.name = FONT_SANS
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_SUBTLE_GOLD
    p.space_after = Pt(8)

    p = tf6.add_paragraph()
    p.text = "“A room full of people who showed up, watched, laughed, and shared the moment.”"
    p.font.name = FONT_SERIF
    p.font.size = Pt(13)
    p.font.italic = True
    p.font.color.rgb = COLOR_WHITE

    if os.path.exists('assets/09_AUDITORIUM_SCREENING/700_plus_registrations_packed_house.png'):
        s6.shapes.add_picture('assets/09_AUDITORIUM_SCREENING/700_plus_registrations_packed_house.png', Inches(0.8), Inches(4.3), width=Inches(4.3), height=Inches(2.65))

    # Right Playable Videos (Embedded MP4 with poster frames)
    # 1. Wide Video (Top Right / Middle)
    vid_wide = 'assets/11_VIDEOS/video_02_audience_wide.mp4'
    poster_wide = 'assets/video_stills/video_02_audience_wide_still.jpg'
    if os.path.exists(vid_wide):
        s6.shapes.add_movie(vid_wide, Inches(5.35), Inches(0.9), Inches(4.2), Inches(6.0), poster_frame_image=poster_wide if os.path.exists(poster_wide) else None, mime_type='video/mp4')
    elif os.path.exists(poster_wide):
        s6.shapes.add_picture(poster_wide, Inches(5.35), Inches(0.9), width=Inches(4.2), height=Inches(6.0))

    # 2. Vertical Video 03 (Auditorium screen)
    vid_v3 = 'assets/11_VIDEOS/video_03_auditorium_screen_vertical.mp4'
    poster_v3 = 'assets/video_stills/video_03_auditorium_screen_vertical_still.jpg'
    if os.path.exists(vid_v3):
        s6.shapes.add_movie(vid_v3, Inches(9.7), Inches(0.9), Inches(1.6), Inches(6.0), poster_frame_image=poster_v3 if os.path.exists(poster_v3) else None, mime_type='video/mp4')
    elif os.path.exists(poster_v3):
        s6.shapes.add_picture(poster_v3, Inches(9.7), Inches(0.9), width=Inches(1.6), height=Inches(6.0))

    # 3. Vertical Video 06 (Event & Audience)
    vid_v6 = 'assets/11_VIDEOS/video_06_event_vertical.mp4'
    poster_v6 = 'assets/video_stills/video_06_event_vertical_still.jpg'
    if os.path.exists(vid_v6):
        s6.shapes.add_movie(vid_v6, Inches(11.45), Inches(0.9), Inches(1.6), Inches(6.0), poster_frame_image=poster_v6 if os.path.exists(poster_v6) else None, mime_type='video/mp4')
    elif os.path.exists(poster_v6):
        s6.shapes.add_picture(poster_v6, Inches(11.45), Inches(0.9), width=Inches(1.6), height=Inches(6.0))

    # =========================================================================
    # SLIDE 07: CORE COMMITTEE (THE PEOPLE BEHIND RAHAT)
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_background(s7, COLOR_IVORY)
    add_eyebrow(s7, "Act V — The Leadership", Inches(0.8), Inches(0.35))

    tf_s7_h = create_textbox(s7, Inches(0.8), Inches(0.65), Inches(11.7), Inches(0.45))
    p = tf_s7_h.paragraphs[0]
    p.text = "THE PEOPLE BEHIND RAHAT"
    p.font.name = FONT_SERIF
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DEEP

    # Top Row: President & Vice President
    # Tanvi (President)
    card_p1 = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.2), Inches(1.15), Inches(3.3), Inches(2.75))
    card_p1.fill.solid()
    card_p1.fill.fore_color.rgb = COLOR_WHITE
    card_p1.line.color.rgb = COLOR_BORDER_LIGHT

    tf_p1 = create_textbox(s7, Inches(3.3), Inches(1.25), Inches(3.1), Inches(2.55))
    p = tf_p1.paragraphs[0]
    p.text = "TANVI"
    p.font.name = FONT_SERIF
    p.font.size = Pt(18)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = COLOR_NAVY_DEEP
    p = tf_p1.add_paragraph()
    p.text = "PRESIDENT"
    p.font.name = FONT_SANS
    p.font.size = Pt(11)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = COLOR_RAHAT_RED
    p.space_after = Pt(6)
    p = tf_p1.add_paragraph()
    p.text = "Executive Leadership & Community Vision"
    p.font.name = FONT_SANS
    p.font.size = Pt(10)
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = COLOR_MUTED_TEXT

    # Pavan (Vice President)
    card_p2 = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.833), Inches(1.15), Inches(3.3), Inches(2.75))
    card_p2.fill.solid()
    card_p2.fill.fore_color.rgb = COLOR_WHITE
    card_p2.line.color.rgb = COLOR_BORDER_LIGHT

    if os.path.exists('assets/committee/pavan_pic.jpeg'):
        s7.shapes.add_picture('assets/committee/pavan_pic.jpeg', Inches(7.0), Inches(1.25), width=Inches(1.3), height=Inches(1.8))

    tf_p2 = create_textbox(s7, Inches(8.4), Inches(1.45), Inches(1.65), Inches(1.6))
    p = tf_p2.paragraphs[0]
    p.text = "PAVAN"
    p.font.name = FONT_SERIF
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DEEP
    p = tf_p2.add_paragraph()
    p.text = "VICE PRESIDENT"
    p.font.name = FONT_SANS
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_RAHAT_RED

    # Bottom Row: 5 Leads
    bottom_leads = [
        ('assets/committee/tanish_pic.jpeg', 'TANISH', 'GENERAL SECRETARY'),
        ('assets/committee/akshita_pic.jpeg', 'AKSHITA', 'CREATIVE LEAD'),
        ('assets/committee/haasini_pic.jpeg', 'HAASINI', 'OPERATIONS LEAD'),
        ('assets/committee/keshav_pic.jpeg', 'KESHAV', 'TREASURER'),
        ('assets/committee/manya_pic.jpeg', 'MANYA', 'WEB & SOCIAL')
    ]
    b_w = Inches(2.23)
    b_gap = Inches(0.14)
    b_start_x = Inches(0.8)
    b_y = Inches(4.1)
    
    for k, (img_path, name, role) in enumerate(bottom_leads):
        bx = b_start_x + k * (b_w + b_gap)
        lead_card = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, b_y, b_w, Inches(3.0))
        lead_card.fill.solid()
        lead_card.fill.fore_color.rgb = COLOR_WHITE
        lead_card.line.color.rgb = COLOR_BORDER_LIGHT

        if os.path.exists(img_path):
            s7.shapes.add_picture(img_path, bx + Inches(0.2), b_y + Inches(0.15), width=Inches(1.83), height=Inches(2.0))

        tf_l_info = create_textbox(s7, bx, b_y + Inches(2.2), b_w, Inches(0.75))
        p = tf_l_info.paragraphs[0]
        p.text = name
        p.font.name = FONT_SERIF
        p.font.size = Pt(13)
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER
        p.font.color.rgb = COLOR_NAVY_DEEP
        p = tf_l_info.add_paragraph()
        p.text = role
        p.font.name = FONT_SANS
        p.font.size = Pt(8.5)
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER
        p.font.color.rgb = COLOR_RAHAT_RED

    # =========================================================================
    # SLIDE 08: CONNECT & JOIN US (HUGE QR CODES)
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_background(s8, COLOR_IVORY)
    add_eyebrow(s8, "Act VI — Connect With Us", Inches(0.8), Inches(0.45))

    tf_s8_h = create_textbox(s8, Inches(0.8), Inches(0.8), Inches(11.7), Inches(0.55))
    p = tf_s8_h.paragraphs[0]
    p.text = "BE A PART OF THE CONVERSATION"
    p.font.name = FONT_SERIF
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY_DEEP

    # Instagram Card (Left - Huge QR)
    card_insta = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(1.5), Inches(5.0), Inches(5.3))
    card_insta.fill.solid()
    card_insta.fill.fore_color.rgb = COLOR_WHITE
    card_insta.line.color.rgb = COLOR_BORDER_LIGHT

    if os.path.exists('assets/qrs/insta_qr_cropped.png'):
        s8.shapes.add_picture('assets/qrs/insta_qr_cropped.png', Inches(2.0), Inches(1.75), width=Inches(3.4), height=Inches(3.4))

    tf_insta = create_textbox(s8, Inches(1.2), Inches(5.25), Inches(5.0), Inches(1.4))
    p = tf_insta.paragraphs[0]
    p.text = "INSTAGRAM"
    p.font.name = FONT_SANS
    p.font.size = Pt(14)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = COLOR_RAHAT_RED
    p = tf_insta.add_paragraph()
    p.text = "@rahat.mitb"
    p.font.name = FONT_SERIF
    p.font.size = Pt(18)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = COLOR_NAVY_DEEP
    p = tf_insta.add_paragraph()
    p.text = "Follow for stories, event updates, reels & community initiatives"
    p.font.name = FONT_SANS
    p.font.size = Pt(10)
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = COLOR_MUTED_TEXT

    # LinkedIn Card (Right - Huge QR)
    card_link = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.133), Inches(1.5), Inches(5.0), Inches(5.3))
    card_link.fill.solid()
    card_link.fill.fore_color.rgb = COLOR_WHITE
    card_link.line.color.rgb = COLOR_BORDER_LIGHT

    if os.path.exists('assets/qrs/linkedin_qr_cropped.png'):
        s8.shapes.add_picture('assets/qrs/linkedin_qr_cropped.png', Inches(7.933), Inches(1.75), width=Inches(3.4), height=Inches(3.4))

    tf_link = create_textbox(s8, Inches(7.133), Inches(5.25), Inches(5.0), Inches(1.4))
    p = tf_link.paragraphs[0]
    p.text = "LINKEDIN"
    p.font.name = FONT_SANS
    p.font.size = Pt(14)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = COLOR_RAHAT_RED
    p = tf_link.add_paragraph()
    p.text = "Rahat MIT Bengaluru"
    p.font.name = FONT_SERIF
    p.font.size = Pt(18)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = COLOR_NAVY_DEEP
    p = tf_link.add_paragraph()
    p.text = "Connect for collaborations, panels & student leadership networks"
    p.font.name = FONT_SANS
    p.font.size = Pt(10)
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = COLOR_MUTED_TEXT

    # =========================================================================
    # SLIDE 09: CLOSING SCREEN
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_background(s9, COLOR_IVORY)

    if os.path.exists('assets/logos/rahat_logo.jpeg'):
        s9.shapes.add_picture('assets/logos/rahat_logo.jpeg', Inches(5.416), Inches(1.0), width=Inches(2.5), height=Inches(2.5))

    tf9 = create_textbox(s9, Inches(1.0), Inches(3.8), Inches(11.333), Inches(3.2))
    p = tf9.paragraphs[0]
    p.text = "“MAKING SPACE”"
    p.font.name = FONT_SERIF
    p.font.size = Pt(44)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = COLOR_NAVY_DEEP
    p.space_after = Pt(8)

    p = tf9.add_paragraph()
    p.text = "Rahat to all."
    p.font.name = FONT_SERIF
    p.font.size = Pt(26)
    p.font.italic = True
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = COLOR_RAHAT_RED
    p.space_after = Pt(14)

    p = tf9.add_paragraph()
    p.text = "mitbengaluru.rahat@gmail.com  •  MIT Bengaluru"
    p.font.name = FONT_SANS
    p.font.size = Pt(12)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = COLOR_MUTED_TEXT

    # Save
    out_path = '/Users/pavanaksshay/rahat/RAHAT_Inauguration_Presentation.pptx'
    prs.save(out_path)
    print(f"Presentation saved successfully to {out_path}")

if __name__ == '__main__':
    build_presentation()
