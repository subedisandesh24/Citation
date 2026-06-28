# modules/__init__.py

# 1. Citation and Reference Remover
from .citation_remover import clean_document_citations

# 2. International Q1 Citation Generator
from .int_generator import generate_international_citations

# 3. 50/50 Regional/International Mixture Citation Generator
from .mixture_generator import generate_mixture_citations

# 4. Nepalese Regional Citation Generator
from .nepal_citation_generator import generate_nepal_citations

# 5. Plagiarism & AI Integrity Checker
from .plag_checker import run_document_integrity_check, AIDetectorEngine

# 6. Deep Learning Paraphraser
from .plag_remover import run_paraphrase_document, ParaphraserEngine

# Expose these utilities to the application level
__all__ = [
    "clean_document_citations",
    "generate_international_citations",
    "generate_mixture_citations",
    "generate_nepal_citations",
    "run_document_integrity_check",
    "AIDetectorEngine",
    "run_paraphrase_document",
    "ParaphraserEngine",
]
