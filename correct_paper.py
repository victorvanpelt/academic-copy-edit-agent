import openai
import os
from docx import Document
import win32com.client as win32
import re

# Load OpenAI API key from environment variable
api_key = os.getenv("OPENAI_API_KEY")
if not api_key:
    raise ValueError("OpenAI API key not found. Set OPENAI_API_KEY as an environment variable.")

openai.api_key = api_key  # Set OpenAI API Key for compatibility

def edit_paragraph_with_chatgpt(text, instruction, model="GPT-4o", debug=False):
    """
    Edit a single paragraph with OpenAI's Chat API for grammar and style correction.
    Ensures that paragraph structure is preserved and formatting is not changed.
    """
    if text.count(".") < 2:  # Skip if it's not a paragraph (less than 2 sentences)
        return text

    messages = [
        {"role": "system", "content": (
            "You are a professional academic editor. "
            "Improve grammar, spelling, style, and professional language use while preserving original formatting. "
            "Do NOT split, merge, reorder, duplicate, or alter paragraph breaks. "
            "Do NOT change citations, equations, or references. "
            "Do NOT modify anything inside parentheses that looks like references (e.g., (Choi et al. 2018)) or square brackets (e.g., [12]). "
            "Preserve all numbers, author names, and citation formatting exactly as they are."
            "ONLY return the corrected paragraph text, nothing else."
        )},
        {"role": "user", "content": f"Correct this paragraph without changing formatting:\n\n{text}\n\n{instruction}"}
    ]

    try:
        response = openai.chat.completions.create(
            model=model,
            messages=messages,
            temperature=0.1
        )

        # Extract and return response
        edited_text = response.choices[0].message.content.strip()

        if debug:
            print(f"Original: {text}")
            print(f"Edited: {edited_text}\n{'-'*40}")

        return edited_text

    except Exception as e:
        print(f"Error calling OpenAI API: {e}")
        return text  # Return original text if API fails

# Instruction for editing
instruction = "Edit grammar. Correct for comma splices. No run-on sentences. Correct preposition usage. Ensure verb tense agreement. Correct spelling. No contractions. Improve academic tone and clarity. Ensure professional and concise writing. Do NOT change domain-specific terminology, technical terms, or proper nouns. Do not change in-line references."

# Load document
input_path = "0_input/paper.docx"
output_path = "1_output/edited_paper.docx"
doc = Document(input_path)

# Track processing
processing = False  # Flag to start processing after "Introduction"
found_abstract = False  # Track if we found the "Abstract" section
stop_keywords = ["References", "Bibliography"]
start_keywords = ["Introduction"]  # Flexible introduction headings

# Function to detect if a line is likely a heading
def is_heading(text):
    """ Returns True if the text is likely a heading or subheading. """
    # Check for numbered headings (e.g., "2.1 Methodology", "3 Results", "4.1 Hypothesis Testing")
    if re.match(r"^\d+(\.\d+)*\s+\w+", text):
        return True
    # Check if the text is short and does not contain full sentences
    if len(text.split()) < 10 and text.count(".") <= 1:
        return True
    return False

# Adjust total paragraph count (ONLY counting actual paragraphs before "References")
paragraph_count = 0

for p in doc.paragraphs:
    text = p.text.strip()

    # Detect and include the paragraph right after "Abstract"
    if re.match(r"^Abstract$", text, re.IGNORECASE):
        found_abstract = True
        continue  # Skip the "Abstract" heading itself

    if found_abstract and text.count(".") >= 2:
        paragraph_count += 1
        found_abstract = False  # Reset flag so only ONE paragraph after Abstract is counted

    # Start counting after "Introduction" (or numbered versions like "1. Introduction")
    if re.match(r"^\d*\.?\s*Introduction$", text, re.IGNORECASE):
        processing = True
        continue

    # Stop counting after "References" (or numbered versions like "6. References")
    if re.match(r"^\d*\.?\s*References$", text, re.IGNORECASE) or text in stop_keywords:
        break  # Stop counting at "References" or "Bibliography"

    # Skip headings/subheadings while processing
    if processing and not is_heading(text) and text.count(".") >= 2:
        paragraph_count += 1

processed_count = 0

for i, p in enumerate(doc.paragraphs):
    text = p.text.strip()

    # Detect and include the paragraph right after "Abstract"
    if re.match(r"^Abstract$", text, re.IGNORECASE):
        found_abstract = True
        continue  # Skip the "Abstract" heading itself

    if found_abstract and text.count(".") >= 2:
        edited_text = edit_paragraph_with_chatgpt(text, instruction, model="gpt-4-turbo", debug=False)
        p.text = edited_text
        processed_count += 1
        print(f"Processed paragraph {processed_count}/{paragraph_count}")
        found_abstract = False  # Reset flag so only ONE paragraph after Abstract is edited

    # Start normal processing after "Introduction" (or numbered versions like "1. Introduction")
    if re.match(r"^\d*\.?\s*Introduction$", text, re.IGNORECASE):
        processing = True
        print("Starting edits after 'Introduction'")
        continue

    # Stop processing after "References" (or numbered versions like "6. References")
    if re.match(r"^\d*\.?\s*References$", text, re.IGNORECASE) or text in stop_keywords:
        print(f"Stopping edits at '{text}'")
        break  # Stop processing at "References" or "Bibliography"

    if processing and text.count(".") >= 2:
        edited_text = edit_paragraph_with_chatgpt(text, instruction, model="gpt-4-turbo", debug=False)
        p.text = edited_text
        processed_count += 1
        print(f"Processed paragraph {processed_count}/{paragraph_count}")

print("Editing complete. Saving document...")

doc.save(output_path)
print(f"Document saved to {output_path}")

# Compare original vs. edited document
original_doc_path = os.path.abspath("0_input/paper.docx")
edited_doc_path = os.path.abspath("1_output/edited_paper.docx")
output_doc_path = os.path.abspath("1_output/trackchanges_paper.docx")

def compare_documents(original, edited, output):
    """
    Automates Microsoft Word's 'Compare Documents' function while ignoring citation changes.
    Only rejects changes inside square brackets if they contain a number.
    """
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False

        original_doc = word.Documents.Open(original)
        edited_doc = word.Documents.Open(edited)

        # Perform document comparison
        compared_doc = word.CompareDocuments(
            OriginalDocument=original_doc,
            RevisedDocument=edited_doc,
            CompareFormatting=False,
            IgnoreAllComparisonWarnings=True
        )

        # Reject citation-related changes
        for change in compared_doc.Revisions:
            text = change.Range.Text

            # Reject if text is inside parentheses (likely an inline citation)
            if re.match(r"\(.*\)", text):  
                change.Reject()

            # Reject if text is inside square brackets AND contains a number (likely a numbered reference)
            elif re.match(r"\[\d+\]", text):  
                change.Reject()

        # Save the cleaned comparison document
        compared_doc.SaveAs(output, FileFormat=16)

        # Close documents properly
        compared_doc.Close(False)
        original_doc.Close(False)
        edited_doc.Close(False)

        word.Quit()  # Fully close Word

        print(f"Comparison completed. Document saved to: {output}")

    except Exception as e:
        print(f"Error comparing documents: {e}")

    finally:
        word.Quit()  # Ensure Word fully exits

compare_documents(original_doc_path, edited_doc_path, output_doc_path)