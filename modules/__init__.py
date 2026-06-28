# Import the main functions/classes from each module to make them accessible
from .citation_remover import remove_citations
from .nepal_generator import generate_nepal_text
from .int_generator import generate_int_text
from .mixture_generator import generate_mixture_text
from .plag_remover import remove_plagiarism
from .plag_checker import check_plagiarism

# Define what is exported when someone does: from modules import *
__all__ = [
    "remove_citations",
    "generate_nepal_text",
    "generate_int_text",
    "generate_mixture_text",
    "remove_plagiarism",
    "check_plagiarism",
]
