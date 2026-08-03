import os
import streamlit as st
from google import genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="BAMS Seminar Generator",
    page_icon="🌿",
    layout="centered"
)

# --- APP UI ---
st.markdown("<h1 style='color: #2e7d32;'>🌿 BAMS Seminar Presentation Generator</h1>", unsafe_allow_html=True)
st.markdown("*Developed by: RAJA KUMAR SAHU@93*")
st.write("Generate comprehensive, content-rich academic presentations (15+ slides) complete with deep core study, classical references, and modern scientific correlations for Ayurveda topics.")

topic = st.text_input("Enter Seminar Topic", "BENEFITS OF BREASTFEEDING ON OBSTETRIC")

if st.button("Generate Seminar Presentation (15+ Slides)"):
    if not topic.strip():
        st.warning("Please enter a valid topic.")
    else:
        # Check for Gemini API key via Streamlit Secrets or environment variable
        api_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")
        
        if not api_key:
            st.error("Gemini API Key missing! Please configure it in your Streamlit Secrets.")
        else:
            try:
                with st.spinner("Synthesizing deep core study, classical Samhitas, and modern insights... Please wait."):
                    # Initialize GenAI Client using the correct google-genai SDK
                    client = genai.Client(api_key=api_key)
                    
                    # Prompt designed to extract deep academic prose rather than simple outlines
                    prompt = f"""
                    Create a comprehensive, highly detailed academic presentation structure for a Bachelor of Ayurvedic Medicine and Surgery (BAMS) seminar on the topic: '{topic}'.
                    
                    You must generate exactly 15 slides. For each slide, provide:
                    1. A clear, academic Slide Title.
                    2. 3 to 4 comprehensive, information-dense bullet points. Each bullet point must contain rich explanatory text, deep clinical analysis, classical terminology (where applicable), mechanisms of action, or modern scientific correlation rather than short phrases.
                    
                    Format your response strictly as follows for each slide:
                    SLIDE X: [Title]
                    - [Detailed paragraph/bullet point with deep core study text]
                    - [Detailed paragraph/bullet point with deep core study text]
                    - [Detailed paragraph/bullet point with deep core study text]
                    
                    Ensure the 15 slides cover: Introduction, Classical References & Samhitas, Fundamental Principles (Siddhanta), Sharira Kriya (Anatomy/Physiology), Samprapti (Pathogenesis), Bheda (Classification), Purva Karma (Pre-protocols), Pradhana Karma (Execution), Paschat Karma (Post-care), Indications/Therapeutic Uses, Contraindications, Dravyaguna (Pharmacology), Modern Scientific Correlation, Critical Analysis & Discussion, and Conclusion & Future Scope.
                    """
                    
                    response = client.models.generate_content(
                        model='gemini-1.5-flash',
                        contents=prompt,
                    )
                    
                    raw_text = response.text
                    
                    # --- POWERPOINT GENERATION ---
                    prs = Presentation()
                    prs.slide_width = Inches(10)
                    prs.slide_height = Inches(5.625) # 16:9 widescreen ratio
                    
                    slides_data = []
                    current_title = ""
                    current_bullets = []
                    
                    for line in raw_text.split('\n'):
                        line = line.strip()
                        if line.startswith("SLIDE "):
                            if current_title:
                                slides_data.append((current_title, current_bullets))
                            parts = line.split(":", 1)
                            current_title = parts[1].strip() if len(parts) > 1 else line
                            current_bullets = []
                        elif line.startswith("-") or line.startswith("*"):
                            bullet_text = line.lstrip("-* ").strip()
                            if bullet_text:
                                current_bullets.append(bullet_text)
                                
                    if current_title:
                        slides_data.append((current_title, current_bullets))
                        
                    # Fallback if parsing fails to capture structured text properly
                    if not slides_data:
                        slides_data = [("Overview of " + topic, [raw_text[:500]])]

                    # 1. Title Slide Creation
                    blank_layout = prs.slide_layouts[6]
                    slide_1 = prs.slides.add_slide(blank_layout)
                    
                    # Background banner
                    bg_shape = slide_1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(5.625))
                    bg_shape.fill.solid()
                    bg_shape.fill.fore_color.rgb = RGBColor(235, 245, 235)
                    bg_shape.line.fill.background()
                    
                    title_box = slide_1.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8), Inches(2.5))
                    tf_1 = title_box.text_frame
                    tf_1.word_wrap = True
                    
                    p_main = tf_1.paragraphs[0]
                    p_main.text = f"Seminar on:\n{topic}"
                    p_main.font.size = Pt(28)
                    p_main.font.bold = True
                    p_main.font.color.rgb = RGBColor(46, 125, 50)
                    
                    p_sub = tf_1.add_paragraph()
                    p_sub.text = "BAMS Comprehensive Academic Presentation\nPrepared by: RAJA KUMAR SAHU@93"
                    p_sub.font.size = Pt(14)
                    p_sub.font.color.rgb = RGBColor(80, 80, 80)
                    
                    # 2. Content Slides Creation
                    for title, bullets in slides_data:
                        slide = prs.slides.add_slide(blank_layout)
                        
                        # Header title
                        t_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.4), Inches(8.8), Inches(0.8))
                        tf = t_box.text_frame
                        tf.word_wrap = True
                        p = tf.paragraphs[0]
                        p.text = title
                        p.font.size = Pt(22)
                        p.font.bold = True
                        p.font.color.rgb = RGBColor(46, 125, 50)
                        
                        # Content text box with deep core study paragraphs
                        c_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.3), Inches(8.8), Inches(3.8))
                        c_tf = c_box.text_frame
                        c_tf.word_wrap = True
                        
                        for idx, bullet in enumerate(bullets):
                            if idx == 0:
                                bp = c_tf.paragraphs[0]
                            else:
                                bp = c_tf.add_paragraph()
                            bp.text = "• " + bullet
                            bp.font.size = Pt(13)
                            bp.font.color.rgb = RGBColor(40, 40, 40)
                            bp.space_after = Pt(10)
                            
                    # Save presentation file locally
                    file_name = f"{topic.replace(' ', '_')}_BAMS_Seminar.pptx"
                    prs.save(file_name)
                    
                    st.success("Presentation generated successfully with deep core study content!")
                    
                    with open(file_name, "rb") as file:
                        st.download_link = st.download_button(
                            label="📥 Download Detailed Presentation (.pptx)",
                            data=file,
                            file_name=file_name,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
            except Exception as e:
                st.error(f"An error occurred during generation: {e}")
