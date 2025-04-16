from plagiarism_checker import similarity_report

external_sources = [
    {
        "title": "Original Paper A",
        "content": "This paper presents a methodology for fraud detection using AI.",
        "references": ["A. Author, ‘AI in Fraud Detection’, 2023."]
    },
    {
        "title": "Original Paper B",
        "content": "We explore neural firewalls for online security applications.",
        "references": ["B. Author, ‘Neural Firewalls’, 2022."]
    }
]

docx_path = "output/ieee_generated_paper.docx"

report = similarity_report(docx_path, external_sources)

print("Semantic Similarity:")
for entry in report['semantic_similarity']:
    print(f" - {entry['title']}: {entry['score']}")

print("\nReference Overlap:")
for entry in report['reference_overlap']:
    print(f" - {entry['title']} matched: {entry['matched_references']}")

print("\nCitations in Paper:")
print(f"Total: {report['citation_analysis']['total']}")
print(f"Found: {report['citation_analysis']['citations_found']}")
