from importlib.metadata import PackageNotFoundError
import os
import re
import torch
import logging
import docx
from sklearn.metrics.pairwise import cosine_similarity
from sentence_transformers import SentenceTransformer
from typing import List, Dict

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

model = SentenceTransformer('all-MiniLM-L6-v2')  # Light, fast BERT model

def extract_text_from_docx(file_path):
    try:
        document = docx.Document(file_path)
        full_text = []
        for para in document.paragraphs:
            full_text.append(para.text)
        return "\n".join(full_text)
    except PackageNotFoundError:
        logger.error("The uploaded file is not a valid .docx file.")
        raise ValueError("The uploaded file is not a valid .docx file. Please upload a proper Word document.")
    except Exception as e:
        logger.error(f"Failed to extract text: {e}")
        raise ValueError(f"Error reading document: {str(e)}")

def split_into_sentences(text: str) -> List[str]:
    # Basic sentence splitting (you can improve with SpaCy/NLTK)
    sentences = re.split(r'(?<=[.!?])\s+', text.strip())
    return [s.strip() for s in sentences if s.strip()]

def extract_references(text: str) -> List[str]:
    # Extract references block based on "References" section
    if "references" not in text.lower():
        return []
    refs = text.lower().split("references")[-1]
    return split_into_sentences(refs)

def check_citations(text: str, references: List[str]) -> Dict[str, bool]:
    citation_pattern = r"\[(\d+)\]"
    found = re.findall(citation_pattern, text)
    ref_map = {str(i+1): ref for i, ref in enumerate(references)}
    citation_check = {}
    for c in found:
        citation_check[f"[{c}]"] = c in ref_map
    return citation_check

def compute_semantic_similarity(sentences: List[str], threshold: float = 0.85) -> List[Dict]:
    embeddings = model.encode(sentences, convert_to_tensor=True)
    sims = cosine_similarity(embeddings.cpu(), embeddings.cpu())
    
    flagged = []
    for i in range(len(sentences)):
        for j in range(i+1, len(sentences)):
            if sims[i][j] > threshold:
                flagged.append({
                    "sentence_1": sentences[i],
                    "sentence_2": sentences[j],
                    "similarity": float(sims[i][j])
                })
    return flagged

def analyze_plagiarism(docx_path: str, threshold: float = 0.85) -> Dict:
    try:
        logger.info("Extracting text...")
        text = extract_text_from_docx(docx_path)
        sentences = split_into_sentences(text)
        references = extract_references(text)
        citations = check_citations(text, references)
        logger.info("Performing semantic analysis...")
        similar_pairs = compute_semantic_similarity(sentences, threshold)

        return {
            "total_sentences": len(sentences),
            "citation_validation": citations,
            "similar_sentences": similar_pairs,
            "plagiarism_score": round(len(similar_pairs) / max(1, len(sentences)), 2)
        }

    except Exception as e:
        logger.error(f"Plagiarism analysis failed: {e}")
        raise RuntimeError("Plagiarism analysis failed.")
