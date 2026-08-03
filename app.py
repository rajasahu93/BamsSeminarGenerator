import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
import io

st.set_page_config(page_title="BAMS Seminar Presentation Generator", layout="centered")

st.markdown("""
    <h2 style='color: #2e7d32;'>🌿 BAMS Seminar Presentation Generator</h2>
    <p style='color: #555;'>Generating comprehensive academic presentations with deep core study prose, classical references, and modern scientific correlations.</p>
""", unsafe_allow_html=True)

topic = st.text_input("Enter Seminar Topic", "BENEFITS OF BREASTFEEDING ON OBSTETRICS")

def create_full_page_presentation(topic_name):
    prs = Presentation()
    
    # 15 detailed slides designed with full paragraph prose per slide
    slides_data = [
        ("Title Slide", [
            f"Comprehensive Academic Seminar on: {topic_name}",
            "Presented in Partial Fulfillment of BAMS Curriculum",
            "Specialized Study in Classical Samhitas, Physiological Mechanics, and Contemporary Clinical Evidence",
            "Department of Prasuti Tantra Evam Striroga / Samhita Adhyayana"
        ]),
        ("1. Comprehensive Introduction & Background", [
            f"The subject of {topic_name} occupies a fundamental position in traditional Ayurvedic literature as well as contemporary neonatal care.",
            "Ayurveda approaches this domain through holistic paradigms, emphasizing maternal health, infant nutrition, and the vital continuity of life energy (Prana).",
            "Modern scientific validation increasingly aligns with these ancient observations, confirming distinct immunological, psychological, and physiological advantages.",
            "This seminar aims to bridge classical textual references with rigorous evidence-based scientific correlation to establish a unified clinical perspective."
        ]),
        ("2. Etymology & Classical Nirukti", [
            "Linguistic derivation forms the cornerstone of Ayurvedic epistemology, uncovering the inherent functional attributes of bodily processes and therapeutic interventions.",
            "Classical Sanskrit nomenclature is carefully structured to express deep physiological mechanics rather than mere descriptive labels.",
            "Analyzing the root terms utilized in Samhitas sheds light on how ancient scholars conceptualized the dynamics of lactation, maternal bonding, and infant nourishment.",
            "This etymological depth ensures that clinical applications remain true to fundamental classical principles while adapting to modern medical terminology."
        ]),
        ("3. Samhita Vachana & Classical References", [
            "Comprehensive review of primary classical treatises including Charaka Samhita, Sushruta Samhita, and Ashtanga Hridaya reveals explicit citations regarding maternal lactation.",
            "Specific passages across Sutra Sthana, Sharira Sthana, and Uttara Tantra detail the precise qualities of pure breast milk (Stanya) and the physiological rules governing its production.",
            "Prominent commentators such as Chakrapani Datta and Dalhana provide profound insights into how disturbances in maternal Rasa dhatu directly impact infant development.",
            "These scriptural citations form an unbroken lineage of medical wisdom, validating the timeless applicability of Ayurvedic obstetrics and pediatrics (Kaumarabhritya)."
        ]),
        ("4. Fundamental Siddhanta (Core Philosophy)", [
            "Ayurvedic physiology views lactation as an advanced physiological transformation, driven primarily by the high-grade nourishment of Rasa dhatu synthesized from wholesome maternal diet.",
            "The delicate equilibrium of Tridosha—specifically Vyana Vata governing circulatory movement, Sadhaka Pitta managing metabolic transformation, and Tarpaka Kapha providing structural fluid support—is essential.",
            "Any imbalance in Agni (digestive fire) or the presence of Ama (toxins) at the tissue level compromises the quality of Stanya, leading to downstream pediatric complications.",
            "Understanding these foundational principles enables practitioners to formulate effective, root-cause-oriented interventions for both mother and child."
        ]),
        ("5. Samprapti (Pathogenesis & Process Dynamics)", [
            "Pathogenesis in the context of lactation disorders follows a clearly defined progression through classical stages, starting from dietary indiscretion (Aahara hetu) to tissue depletion.",
            "Improper lifestyle or emotional stress vitiates Prana Vata and Udana Vata, obstructing the natural downward and outward flow pathways required for unobstructed milk ejection.",
            "The interaction between vitiated Doshas and Dhatus results in qualitative or quantitative defects in Stanya production, presenting clinically as sub-optimal infant growth or maternal fatigue.",
            "Early identification of these pathological milestones allows for timely therapeutic corrections using classical dietary and herbal protocols."
        ]),
        ("6. Classification & Clinical Typology (Bheda)", [
            "Classical texts categorize lactation abnormalities and related obstetric parameters into distinct typologies based on Dosha predominance and severity indices.",
            "Stanya Dushti (vitiations of breast milk) is systematically classified into Vataja, Pittaja, Kaphaja, and Sannipatika variants, each displaying unique physical attributes and symptom profiles in the infant.",
            "Recognizing these specific classifications is critical for tailoring precise line-of-treatment strategies rather than applying generalized interventions.",
            "Differential diagnostic parameters outlined in Samhitas help clinicians distinguish between primary glandular insufficiencies and secondary lifestyle-induced lactation failures."
        ]),
        ("7. Comprehensive Diagnostic Protocol (Nidana Panchaka)", [
            "Diagnosis in Ayurvedic clinical practice utilizes the five-fold framework of Nidana Panchaka to map out the complete clinical trajectory of maternal-infant health conditions.",
            "Nidana identifies root causes such as improper maternal diet, emotional distress, and physical overexertion that disrupt normal lactation cycles.",
            "Purvaroopa identifies prodromal signs like heaviness in breasts or subtle drops in milk supply before full clinical manifestation occurs.",
            "Roopa, Upashaya, and Samprapti collectively complete the diagnostic matrix, empowering practitioners to secure an accurate clinical evaluation."
        ]),
        ("8. Management & Chikitsa Siddhanta", [
            "The strategic management of conditions related to {topic_name} centers on restoring metabolic harmony, purifying bodily channels, and augmenting maternal vitality.",
            "Chikitsa Sutra emphasizes the use of Deepana and Pachana measures to correct digestive fire before administering specialized Galactagogues (Stanyajanana drugs).",
            "Internal administration of classical formulations enriched with nutritive herbs supports tissue regeneration and ensures optimal milk composition for infant consumption.",
            "Integrative therapeutic planning combines restorative dietary habits with psychological reassurance to maximize overall clinical efficacy."
        ]),
        ("9. Pathya-Apathya (Diet & Lifestyle Regimen)", [
            "Dietary discipline (Pathya) is critical during the postpartum and lactation period to sustain high-quality milk production without depleting maternal somatic reserves.",
            "Recommended dietary articles include easily digestible, nutrient-dense foods cooked with traditional spices that stimulate digestive fire and balance Vata dosha.",
            "Prohibited habits (Apathya), such as excessive dry, cold, or processed foods {topic_name}, are strictly discouraged to prevent Ama accumulation and channel blockage.",
            "Adherence to structured daily regimens (Dinacharya) ensures sustained maternal energy levels and robust immune defense."
        ]),
        ("10. Modern Scientific Correlation", [
            "Contemporary medical science strongly corroborates classical Ayurvedic principles regarding the profound biochemical and immunological superiority of natural infant feeding.",
            "Clinical research highlights the presence of live leukocytes, secretory antibodies, and bioactive growth factors in breast milk that mirror the classical concept of enhancing Ojas and immunity.",
            "Endocrinological studies validate the role of oxytocin and prolactin release during feeding, aligning with the ancient understanding of neuro-hormonal and energetic pathways.",
            "Bridging these perspectives establishes a powerful scientific foundation for integrating traditional wisdom into modern obstetric guidelines."
        ]),
        ("11. Pharmacology & Drug Action (Dravya Guna)", [
            "Pharmacological evaluation of galactagogue herbs relies on understanding their unique Rasa, Guna, Virya, Vipaka, and Prabhava profiles.",
            "Substances possessing Madhura (sweet) and Snigdha (unctuous) properties act synergistically to nourish Rasa dhatu and stimulate specialized mammary tissue receptors.",
            "Cellular-level actions focus on enhancing metabolic assimilation and supporting systemic micro-circulation to facilitate unobstructed glandular secretion.",
            "Standardized quality control and bio-availability assessments ensure safety, purity, and predictable therapeutic outcomes in clinical practice."
        ]),
        ("12. Clinical Observations & Case Study Framework", [
            "Structured clinical observation frameworks provide empirical validation for classical Ayurvedic interventions applied within modern obstetric settings.",
            "Documenting patient cohorts undergoing traditional management reveals statistically significant improvements in milk volume, infant weight-gain trajectories, and maternal recovery indices.",
            "Standardized scoring systems assess parameters such as fatigue levels, emotional well-being, and digestive comfort throughout the postnatal phase.",
            "These findings contribute valuable data to the broader academic community, encouraging further evidence-based exploration of Ayurvedic gynecology."
        ]),
        ("13. Preventive Aspects & Public Health (Swasthavritta)", [
            "Preventive healthcare principles embedded in Swasthavritta emphasize proactive management of maternal health to avert long-term chronic complications.",
            "Promoting natural lactation protocols serves as a primary public health intervention against infant malnutrition, allergies, and childhood obesity.",
            "Rejuvenation (Rasayana) therapies administered during the postnatal period secure long-term vitality, preventing premature aging and constitutional depletion.",
            "Community-level awareness campaigns rooted in classical traditions play an essential role in improving societal health metrics."
        ]),
        ("14. Conclusion & Summary", [
            "The systematic exploration of {topic_name} demonstrates the remarkable depth, precision, and enduring relevance of classical Ayurvedic science.",
            "Synthesizing scriptural wisdom with modern scientific correlations confirms that traditional paradigms offer sophisticated solutions for contemporary health challenges.",
            "Future research directions must focus on large-scale clinical trials and standardized molecular evaluations to further validate ancient medical doctrines.",
            "The ultimate clinical objective remains the holistic well-being of both mother and child through balanced, compassionate, and evidence-based care."
        ]),
        ("15. References & Bibliography", [
            "1. Charaka Samhita, Ayurveda Dipika Commentary by Chakrapani Datta, edited by Jadavaji Trikamji Acharya, Chaukhamba Surbharati Prakashan.",
            "2. Sushruta Samhita, Nibandha Samgraha Commentary by Dalhana, edited by Vaidya Jadavaji Trikamji Acharya, Chaukhamba Orientalia.",
            "3. Ashtanga Hridaya of Vagbhata, Vidyotini Hindi Commentary by Kaviraj Atrideva Gupta, Chaukhamba Sanskrit Series Office.",
            "4. Contemporary peer-reviewed journals in Obstetrics, Gynecology, and Integrative Ayurvedic Medicine."
        ])
    ]

    for title, paragraphs in slides_data:
        slide_layout = prs.slide_layouts[1] # Title and Content layout
        slide = prs.slides.add_slide(slide_layout)
        
        # Title formatting
        title_shape = slide.shapes.title
        title_shape.text = title
        
        # Content formatting - ensuring full prose paragraphs per slide
        body_shape = slide.placeholders[1]
        tf = body_shape.text_frame
        tf.clear()
        
        for i, para_text in enumerate(paragraphs):
            p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
            p.text = para_text
            p.level = 0
            p.font.size = Pt(13)  # Optimized font size to fit full text blocks cleanly per slide
            
    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

if st.button("Generate Seminar Presentation (15+ Slides - Full Page Content)"):
    with st.spinner("Compiling full-page academic prose, classical texts, and detailed slides..."):
        try:
            pptx_file = create_full_page_presentation(topic)
            st.success("Presentation successfully generated with comprehensive full-page content across all slides!")
            st.download_button(
                label="📥 Download Detailed Presentation (.pptx)",
                data=pptx_file,
                file_name=f"{topic.replace(' ', '_')}_Full_BAMS_Seminar.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        except Exception as e:
            st.error(f"Error generating presentation: {e}")
