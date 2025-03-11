from bertalign import Bertalign
from docx import Document
import os

src_file = "doc/2023 ASPAC DBS Regional Report.docx"
tgt_file = "doc/2023 ASPAC DBS Regional Report zh.docx"
output_file = "doc/alignment_result.txt"

def read_docx(file_path):
    doc = Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        if para.text.strip():  # Only include non-empty paragraphs
            full_text.append(para.text)
    return "\n".join(full_text)

def save_alignment(src_sents, tgt_sents, alignments, output_file):
    with open(output_file, 'w', encoding='utf-8') as f:
        for (src_idx, tgt_idx) in alignments:
            # Convert indices to lists if they're not already
            if isinstance(src_idx, int):
                src_idx = [src_idx]
            if isinstance(tgt_idx, int):
                tgt_idx = [tgt_idx]
                
            # Get the sentences for these indices
            src_text = ' '.join([src_sents[i] for i in src_idx])
            tgt_text = ' '.join([tgt_sents[i] for i in tgt_idx])
            
            # Write to file
            # f.write(f"SOURCE: {src_text}\n")
            # f.write(f"TARGET: {tgt_text}\n")
            # f.write("-" * 80 + "\n")

            f.write(f"{src_text}\n{tgt_text}\n\n")

def main():
    # Read the Word documents
    src_text = read_docx(src_file)
    tgt_text = read_docx(tgt_file)
    
    # Create aligner and align sentences
    aligner = Bertalign(src_text, tgt_text)
    aligner.align_sents()
    
    # Save the alignment results
    save_alignment(aligner.src_sents, aligner.tgt_sents, aligner.result, output_file)
    print(f"Alignment results have been saved to {output_file}")

if __name__ == "__main__":
    main()
