import openai
import os
from docx import Document
import win32com.client as win32
import re

###############################################################################
# Paths and housekeeping
###############################################################################
original_doc_path = os.path.abspath("0_input/paper.docx")
edited_doc_path   = os.path.abspath("1_output/edited_paper.docx")
output_doc_path   = os.path.abspath("1_output/trackchanges_paper.docx")

# Choose your model (or set the GPT_MODEL env var)
gpt_model = os.getenv("GPT_MODEL", "gpt-4o")

# Remove old output files if they exist
for path in [edited_doc_path, output_doc_path]:
    if os.path.exists(path):
        try:
            os.remove(path)
        except PermissionError:
            print(f"❌ Error: The file '{path}' is currently open or locked. Please close it and try again.")
        except Exception as e:
            print(f"⚠️ Unexpected error deleting '{path}': {e}")

if not os.path.exists(original_doc_path):
    raise FileNotFoundError(
        "❌ 'paper.docx' does not exist in the input folder. "
        "Please ensure this file is named correctly and placed in the '0_input' directory."
    )

# Load OpenAI API key from environment variable
api_key = os.getenv("OPENAI_API_KEY")
if not api_key:
    raise ValueError("OpenAI API key not found. Set OPENAI_API_KEY as an environment variable.")
openai.api_key = api_key

###############################################################################
# 1) Helper functions
###############################################################################

def is_heading(text: str) -> bool:
    """Return True for paragraphs that look like headings."""
    # 1) Numeric heading like "2.1" or "3" at the start
    if re.match(r"^\d+(\.\d+)*\s+\w+", text):
        return True
    # 2) Very short lines (likely section titles)
    if len(text.split()) < 10 and text.count(".") <= 1:
        return True
    return False


def split_into_sentences(text: str):
    """Split text keeping the sentence punctuation as separate items."""
    return re.split(r'([.?!])', text)


def reassemble_sentences(parts):
    """Re-join split text pieces into a clean paragraph."""
    result = []
    for i in range(0, len(parts), 2):
        sentence = parts[i].strip()
        punctuation = parts[i + 1] if i + 1 < len(parts) else ""
        # Remove trailing punctuation to avoid duplication after edit
        if sentence and sentence[-1] in ".!?":
            sentence = sentence.rstrip(".?!")
        if sentence:
            combined = (sentence + punctuation).strip()
            combined = re.sub(r"\.+", ".", combined)  # collapse double dots
            result.append(combined)
    joined = " ".join(result)
    joined = re.sub(r"\s([.?!])", r"\\1", joined)  # remove space before punctuation
    return joined


def edit_sentence_with_chatgpt(sentence: str, model: str = gpt_model) -> str:
    """Make a minimal copy‑edit call for one sentence."""
    # Skip citations / very short sentences
    if re.search(r"\(.*?\)", sentence) or re.search(r"\[.*?\]", sentence):
        return sentence
    if len(sentence.split()) < 3:
        return sentence

    system_prompt = (
        "You are a professional academic copy editor. Improve grammar, spelling, "
        "concision, clarity, and academic style in American English while "
        "preserving meaning and terminology.\n"
        "Follow strictly: 1) Do not alter citations/references/footnotes. 2) Do not merge, "
        "split, or reorder sentences. 3) Return ONLY the corrected sentence."
    )

    try:
        response = openai.chat.completions.create(
            model=model,
            temperature=0,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": sentence},
            ],
        )
        edited = response.choices[0].message.content.strip()
        edited = re.sub(r"\.+$", ".", edited)  # trim trailing dots
        return edited
    except Exception as e:
        print(f"⚠️ Error calling OpenAI API: {e}")
        return sentence


def edit_paragraph_sentencewise(paragraph_text: str, model: str = gpt_model) -> str:
    """Edit each sentence in a paragraph and reassemble."""
    parts = split_into_sentences(paragraph_text)
    edited_parts = []
    for i in range(0, len(parts), 2):
        text_chunk = parts[i].strip() if i < len(parts) else ""
        punctuation = parts[i + 1] if i + 1 < len(parts) else ""
        if text_chunk:
            edited_text = edit_sentence_with_chatgpt(text_chunk, model=model)
            edited_parts.append(edited_text)
            edited_parts.append(punctuation)
    return reassemble_sentences(edited_parts)

###############################################################################
# 2) Document processing logic
###############################################################################

doc = Document(original_doc_path)

processing = False  # We start editing AFTER we hit "Introduction"
stop_keywords = {"references", "bibliography"}

processed_count = 0

print("🚀 Starting copy‑edit…")

for para in doc.paragraphs:
    raw_text = para.text.strip()

    # Detect the section boundaries --------------------------------------
    if re.match(r"^(?:\d+\.?\s*)?Introduction$", raw_text, re.IGNORECASE):
        processing = True  # Start editing from the *next* paragraph
        print("   ↳ Entering main body after 'Introduction'")
        continue

    if raw_text.lower() in stop_keywords or re.match(r"^\d*\.?\s*References$", raw_text, re.IGNORECASE):
        print(f"   ↳ Reached '{raw_text}', stopping edits")
        break

    # --------------------------------------------------------------------
    if processing:
        # Skip headings but EDIT EVERYTHING ELSE (no sentence‑count filter)
        if not is_heading(raw_text) and raw_text:
            new_text = edit_paragraph_sentencewise(raw_text)
            para.text = new_text
            processed_count += 1
            print(f"      • Edited paragraph {processed_count}")

print("✅ Editing done. Saving…")
doc.save(edited_doc_path)
print(f"✅ Saved to {edited_doc_path}")

###############################################################################
# 3) Compare documents in Word (optional — Windows only)
###############################################################################

def compare_documents(original: str, edited: str, output: str):
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False

        original_doc = word.Documents.Open(original)
        edited_doc = word.Documents.Open(edited)
        compared_doc = word.CompareDocuments(
            OriginalDocument=original_doc,
            RevisedDocument=edited_doc,
            CompareFormatting=False,
            IgnoreAllComparisonWarnings=True,
        )
        compared_doc.SaveAs(output, FileFormat=16)  # wdFormatXMLDocument
        compared_doc.Close(False)
        original_doc.Close(False)
        edited_doc.Close(False)
        word.Quit()
        print(f"✅ Track‑changes doc saved to {output}")
    except Exception as e:
        print(f"⚠️ Word comparison failed: {e}")

# Attempt comparison (silently skip if Word/pywin32 unavailable)
try:
    compare_documents(original_doc_path, edited_doc_path, output_doc_path)
except Exception:
    pass

print("🏁 All done!")
