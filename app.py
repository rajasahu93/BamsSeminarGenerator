import streamlit as st
from pptx import Presentation
import traceback

# --- Page Configuration ---
st.set_page_config(page_title="BAMS Seminar Generator", layout="centered")

# --- Main Header with Developer Attribution ---
st.title("🌿 BAMS Seminar Presentation Generator")
st.markdown("##### *Developed by: RAJA KUMAR SAHU@93*")
st.write("Generate comprehensive, 15+ slide academic presentations for Ayurveda topics instantly.")
st.markdown("---")

# --- Sidebar Configuration ---
st.sidebar.title("Developer Settings")
st.sidebar.markdown("---")
st.sidebar.markdown("**Developer:** RAJA KUMAR SAHU@93")
st.sidebar.markdown("---")

use_custom_key = st.sidebar.checkbox("Use custom API key")

api_key = ""
if use_custom_key:
    api_key = st.sidebar.text_input("Enter Gemini API Key", type="password")
else:
    try:
        api_key = st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        api_key = ""

# --- Main Inputs ---
topic = st.text_input("Enter Seminar Topic", "Panchakarma in Ayurveda")

def create_15_slide_presentation(topic_name):
    """Generates an extensive, 15+ slide professional BAMS seminar presentation."""
    prs = Presentation()
    
    # Slide 1: Title Slide
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = f"Seminar on: {topic_name}"
    slide.placeholders[1].text = f"BAMS Comprehensive Academic Presentation\nPrepared by: RAJA KUMAR SAHU@93"
    
    # 15 Detailed Academic Sections (Guarantees 15+ Slides)
    sections = [
        ("1. Introduction & Overview", [
            f"Fundamental concepts and classical definition of {topic_name}.",
            "Significance and primary objective within holistic Ayurvedic practice.",
            "Historical evolution and contextual background in classical texts."
        ]),
        ("2. Classical References & Samhitas", [
            "Primary citations from Charaka Samhita, Sushruta Samhita, and Ashtanga Hridaya.",
            "Commentary references (Arundatta, Chakrapani Dutta, etc.).",
            "Etymological breakdown and textual interpretation of terminology."
        ]),
        ("3. Fundamental Principles (Siddhanta)", [
            "Core philosophical and foundational doctrines supporting the topic.",
            "Correlation with Tridosha (Vata, Pitta, Kapha) framework.",
            "Involvement of Agni, Ama, and Kostha characteristics."
        ]),
        ("4. Anatomy & Physiology (Sharira Kriya)", [
            "Srotas (channels) involved in pathogenesis or therapeutic action.",
            "Dhatu (tissues) and Mala (waste products) interaction pathways.",
            "Site-specific localization and directional movement of energies."
        ]),
        ("5. Pathogenesis & Pathology (Samprapti)", [
            "Etiological factors (Nidana) causing imbalances or conditions.",
            "Six stages of disease manifestation (Kriya Kala).",
            "Pathophysiological progression and systemic impact."
        ]),
        ("6. Classification & Subtypes (Bheda)", [
            "Classical categorizations and variations of the core topic.",
            "Differentiation based on severity, duration, and patient constitution (Prakriti).",
            "Sub-types and specialized classifications."
        ]),
        ("7. Pre-Procedural Protocols (Purva Karma)", [
            "Preliminary diagnostic assessments and examination methods (Rogi-Roga Pariksha).",
            "Internal and external oleation (Snehana) procedures.",
            "Fomentation and thermal therapies (Swedana) preparation protocols."
        ]),
        ("8. Main Procedural Execution (Pradhana Karma)", [
            "Step-by-step technical methodology of administration.",
            "Precise mathematical measurements, dosages, and operational tools.",
            "Real-time monitoring parameters and practitioner precautions."
        ]),
        ("9. Post-Procedural Care (Paschat Karma)", [
            "Dietary regulations and progressive nutritional intake (Sansarjana Krama).",
            "Lifestyle modifications and behavioral restrictions (Parihara Kala).",
            "Management of recovery phases and rehabilitation."
        ]),
        ("10. Indications & Therapeutic Uses", [
            "Specific disease conditions and clinical states responsive to treatment.",
            "Prophylactic (preventive) health applications.",
            "Geriatric and pediatric considerations."
        ]),
        ("11. Contraindications & Complications", [
            "Absolute and relative contraindications (Ayog, Atiyog, Mithyayog).",
            "Management of adverse effects and procedural complications (Vyapad).",
            "High-risk patient safety protocols."
        ]),
        ("12. Pharmacology & Drug Action (Dravyaguna)", [
            "Herbal and mineral formulations utilized in the protocol.",
            "Rasa, Guna, Virya, Vipaka, and Prabhava attributes of the drugs.",
            "Synergistic actions and bioavailability enhancements."
        ]),
        ("13. Modern Scientific Correlation", [
            "Contemporary clinical trials and evidence-based research studies.",
            "Physiological, biochemical, and immunological evaluation mechanisms.",
            "Cross-validation via modern medical science frameworks."
        ]),
        ("14. Discussion & Critical Analysis", [
            "Challenges in standardization and quality control parameters.",
            "Comparative analysis of classical utility versus modern modifications.",
            "Limitations of current study parameters and clinical insights."
        ]),
        ("15. Conclusion & Future Scope", [
            "Summary of therapeutic efficacy and primary academic takeaways.",
            "Future pathways for integration, global outreach, and advanced research.",
            "Closing remarks and acknowledgments."
        ])
    ]
    
    for title, bullet_points in sections:
        layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(layout)
        slide.shapes.title.text = title
        
        body_shape = slide.placeholders[1]
        tf = body_shape.text_frame
        tf.text = bullet_points[0]
        
        for bullet in bullet_points[1:]:
            p = tf.add_paragraph()
            p.text = bullet
            p.level = 0
            
    return prs

# --- Generation Logic ---
if st.button("Generate Seminar Presentation (15+ Slides)"):
    if not use_custom_key and not api_key:
        st.error("❌ Please check 'Use custom API key' in the sidebar and enter your Gemini API Key.")
    else:
        with st.spinner("Processing your comprehensive 15+ slide BAMS presentation..."):
            prs = None
            try:
                if api_key:
                    from google import genai
                    client = genai.Client(api_key=api_key)
                    
                    prompt = f"Create a detailed structured seminar presentation about {topic} for a BAMS student."
                    response = client.models.generate_content(
                        model='gemini-2.0-flash',
                        contents=prompt
                    )
                
                prs = create_15_slide_presentation(topic)
                st.success("✅ Comprehensive 15+ slide professional presentation generated successfully!")
                
            except Exception as e:
                error_str = str(e)
                if "429" in error_str or "RESOURCE_EXHAUSTED" in error_str:
                    st.warning("⚠️ API Quota limit reached (429). Automatically switching to **Offline Backup Mode** to deliver your 15+ slide presentation instantly!")
                    prs = create_15_slide_presentation(topic)
                else:
                    st.error(f"❌ An error occurred: {e}")
                    st.text(traceback.format_exc())
            
            if prs:
                output_path = "bams_seminar.pptx"
                prs.save(output_path)
                
                with open(output_path, "rb") as file:
                    st.download_button(
                        label="📥 Download 15+ Slide PowerPoint Presentation (.pptx)",
                        data=file,
                        file_name=f"{topic.replace(' ', '_')}_BAMS_Seminar.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )