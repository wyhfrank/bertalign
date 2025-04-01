from bertalign import Bertalign
from docx import Document
import os
import pandas as pd

file_pairs = [
    # ("doc/2023 ASPAC DBS Regional Report.docx", "doc/2023 ASPAC DBS Regional Report zh.docx"),
    ("doc/parallel data/2023 ASPAC DBS Regional Report.docx", "doc/parallel data/2023 ASPAC DBS Regional Report_Chinese.docx"),
    ("doc/parallel data/IRS for duty-free Imported Cigarettes and Cigars_V1.0.docx", "doc/parallel data/行业调控项目进口卷烟及雪茄烟应用模块系统集成方案【免税】V1.0-2_rupY.docx"),
]


def main():
    for src_file, tgt_file in file_pairs:
        align_doc(src_file, tgt_file)


def align_doc(src_file, tgt_file):
    # Read the Word documents
    src_text = read_docx(src_file)
    tgt_text = read_docx(tgt_file)
    
    excel_output = os.path.splitext(src_file)[0] + "_alignment.xlsx"
    
    # Create aligner and align sentences
    aligner = Bertalign(src_text, tgt_text)
    aligner.align_sents()
    
    # Save the alignment results to both text and Excel files
    # save_alignment(aligner.src_sents, aligner.tgt_sents, aligner.result, output_file)
    save_alignment_to_excel(aligner.src_sents, aligner.tgt_sents, aligner.result, excel_output)
    # print(f"Alignment results have been saved to {output_file}")



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

def save_alignment_to_excel(src_sents, tgt_sents, alignments, output_file):
    # Create lists to store the aligned sentences
    source_sentences = []
    target_sentences = []
    
    for (src_idx, tgt_idx) in alignments:
        # Convert indices to lists if they're not already
        if isinstance(src_idx, int):
            src_idx = [src_idx]
        if isinstance(tgt_idx, int):
            tgt_idx = [tgt_idx]
            
        # Get the sentences for these indices
        src_text = ' '.join([src_sents[i] for i in src_idx])
        tgt_text = ' '.join([tgt_sents[i] for i in tgt_idx])
        
        source_sentences.append(src_text)
        target_sentences.append(tgt_text)
    
    # Create a DataFrame
    df = pd.DataFrame({
        'Source Text': source_sentences,
        'Target Text': target_sentences
    })
    
    # Save to Excel
    df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"Alignment results have been saved to Excel file: {output_file}")

        
if __name__ == "__main__":
    main()
