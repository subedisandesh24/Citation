# app.py
import re
import streamlit as st
import pandas as pd

# Import our registered modules
from modules import (
    generate_nepal_citations,
    generate_international_citations,
    generate_mixture_citations,
    clean_document_citations,
    run_document_integrity_check,
    AIDetectorEngine,
    run_paraphrase_document,
    ParaphraserEngine
)

st.set_page_config(layout="wide", page_title="Thesis Integrity & Citation Toolkit")

# --- Resource Caching for Transformer Models ---
@st.cache_resource
def load_ai_detector():
    return AIDetectorEngine()

@st.cache_resource
def load_paraphraser():
    return ParaphraserEngine()

# Initialize resources
st.sidebar.subheader("Model Status")
with st.sidebar:
    with st.spinner("Initializing Deep Learning Models..."):
        ai_detector = load_ai_detector()
        paraphraser = load_paraphraser()
    st.success("All models loaded!")

st.title("Academic Paper Processing & Integrity Suite")
st.write("A unified workspace to clean, paraphrase, analyze, and enrich your academic documents.")

# --- Tab Layout for the 6 Core Workflows ---
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "🇳🇵 Regional Citations",
    "🌍 International Q1 Citations",
    "⚖️ 50/50 Balanced Mixture",
    "🧹 Remove Citations",
    "✍️ Document Paraphraser",
    "🔍 Integrity & AI Scan"
])

# ==========================================
# TAB 1: REGIONAL CITATIONS (NEPAL)
# ==========================================
with tab1:
    st.header("Nepalese Regional Citation Generator")
    st.write("Inserts contextually relevant Nepalese studies and literature into your document.")
    
    file_t1 = st.file_uploader("Upload Word Document (.docx)", type=["docx"], key="file_t1")
    count_t1 = st.slider("Target Citation Count", 1, 100, 50, key="count_t1")
    
    if file_t1 and st.button("Generate Regional Citations", key="btn_t1"):
        prog_bar = st.progress(0.0)
        status = st.empty()
        
        def cb_t1(curr, total, query):
            prog_bar.progress(curr / total)
            status.text(f"Querying [{curr}/{total}]: Searching database for '{query}'...")
            
        out_buf, paras, refs, report = generate_nepal_citations(file_t1, target_citations=count_t1, progress_callback=cb_t1)
        status.success("Completed!")
        
        if out_buf:
            st.download_button("📥 Download Document", out_buf, "Nepal_Cited_Document.docx", key="dl_t1")
            st.subheader("Verification Checklist")
            st.dataframe(pd.DataFrame(report)[["id", "context", "title", "author", "doi", "status"]], use_container_width=True)

# ==========================================
# TAB 2: INTERNATIONAL Q1 CITATIONS
# ==========================================
with tab2:
    st.header("International Q1 Citation Generator")
    st.write("Enriches your text with highly cited peer-reviewed articles from globally prestigious publishers.")
    
    file_t2 = st.file_uploader("Upload Word Document (.docx)", type=["docx"], key="file_t2")
    count_t2 = st.slider("Target Citation Count", 1, 100, 50, key="count_t2")
    
    if file_t2 and st.button("Generate International Citations", key="btn_t2"):
        prog_bar = st.progress(0.0)
        status = st.empty()
        
        def cb_t2(curr, total, query):
            prog_bar.progress(curr / total)
            status.text(f"Querying [{curr}/{total}]: Searching database for '{query}'...")
            
        out_buf, paras, refs, report = generate_international_citations(file_t2, target_citations=count_t2, progress_callback=cb_t2)
        status.success("Completed!")
        
        if out_buf:
            st.download_button("📥 Download Document", out_buf, "Global_Cited_Document.docx", key="dl_t2")
            st.subheader("Verification Checklist")
            st.dataframe(pd.DataFrame(report)[["id", "context", "title", "author", "doi", "impact_rating", "status"]], use_container_width=True)

# ==========================================
# TAB 3: balanced MIXTURE CITATIONS (50/50)
# ==========================================
with tab3:
    st.header("50/50 Balanced Mixture Citation Generator")
    st.write("Alternates queries to balance regional context and high-impact global studies evenly.")
    
    file_t3 = st.file_uploader("Upload Word Document (.docx)", type=["docx"], key="file_t3")
    count_t3 = st.slider("Target Citation Count", 1, 100, 50, key="count_t3")
    
    if file_t3 and st.button("Generate Balanced Mixture", key="btn_t3"):
        prog_bar = st.progress(0.0)
        status = st.empty()
        
        def cb_t3(curr, total, query):
            prog_bar.progress(curr / total)
            status.text(f"Querying [{curr}/{total}]: Searching database for '{query}'...")
            
        out_buf, paras, refs, report = generate_mixture_citations(file_t3, target_citations=count_t3, progress_callback=cb_t3)
        status.success("Completed!")
        
        if out_buf:
            st.download_button("📥 Download Document", out_buf, "Mixture_Cited_Document.docx", key="dl_t3")
            st.subheader("Verification Checklist")
            st.dataframe(pd.DataFrame(report)[["id", "mode", "context", "title", "author", "doi", "rating", "status"]], use_container_width=True)

# ==========================================
# TAB 4: CITATION & REFERENCE REMOVER
# ==========================================
with tab4:
    st.header("Citation & Reference Remover")
    st.write("Strips trailing reference lists and removes inline citations from paragraphs and tables.")
    
    file_t4 = st.file_uploader("Upload Word Document (.docx)", type=["docx"], key="file_t4")
    
    if file_t4 and st.button("Clean Citations", key="btn_t4"):
        with st.spinner("Cleaning document..."):
            out_buf, logs = clean_document_citations(file_t4)
        
        st.success("Document cleanup finished.")
        for log in logs:
            st.info(log)
            
        st.download_button("📥 Download Cleaned Document", out_buf, "Cleaned_Document.docx", key="dl_t4")

# ==========================================
# TAB 5: PARAPHRASER (PLAGIARISM REMOVER)
# ==========================================
with tab5:
    st.header("Deep Learning Document Paraphraser")
    st.write("Rewrites non-formatting sentences using a sequence-to-sequence transformer model to assist with original framing.")
    
    file_t5 = st.file_uploader("Upload Word Document (.docx)", type=["docx"], key="file_t5")
    
    if file_t5 and st.button("Run Paraphraser", key="btn_t5"):
        prog_bar = st.progress(0.0)
        status = st.empty()
        
        def cb_t5(curr, total, text_preview):
            prog_bar.progress(curr / total)
            status.text(f"Paraphrasing [{curr}/{total}]: \"{text_preview}\"")
            
        with st.spinner("Running deep learning model inference..."):
            out_buf, changes, logs = run_paraphrase_document(file_t5, engine=paraphraser, progress_callback=cb_t5)
        
        status.success(f"Processing Complete! Paraphrased {changes} paragraphs.")
        
        if out_buf:
            st.download_button("📥 Download Paraphrased Document", out_buf, "Paraphrased_Document.docx", key="dl_t5")

# ==========================================
# TAB 6: PLAGIARISM & AI INTEGRITY SCAN
# ==========================================
with tab6:
    st.header("Document Integrity & AI Scanner")
    st.write("Scans for machine-generated writing probabilities and queries live web engines to identify text matches.")
    
    file_t6 = st.file_uploader("Upload Word Document (.docx)", type=["docx"], key="file_t6")
    
    if file_t6 and st.button("Run Integrity Check", key="btn_t6"):
        prog_bar = st.progress(0.0)
        status = st.empty()
        
        def cb_t6(step, percent):
            prog_bar.progress(percent / 100)
            status.text(f"Currently processing: {step} ({percent}%)")
            
        with st.spinner("Analyzing document structure..."):
            results = run_document_integrity_check(file_t6, detector_engine=ai_detector, progress_callback=cb_t6)
            
        status.success("Scan Completed!")
        
        if results["success"]:
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Estimated AI Score", f"{results['ai_score']}%")
                st.info(f"AI Verdict: {results['ai_verdict']}")
                st.write(f"Scanned Paragraphs: {results['paragraphs_scanned']}")
            with col2:
                st.metric("Similarity Index", f"{results['similarity_index']}%")
                st.info(f"Similarity Verdict: {results['similarity_verdict']}")
                st.write(f"Evaluated Sentences: {results['sentences_checked']}")
                
            if results["matched_sources"]:
                st.subheader("Matched Online Sources")
                for source in results["matched_sources"]:
                    st.write(f"- [{source}]({source})")
