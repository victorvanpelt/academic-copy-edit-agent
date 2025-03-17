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
gpt_model = "gpt-4o"

# Remove old output files if they exist
for path in [edited_doc_path, output_doc_path]:
    if os.path.exists(path):
        os.remove(path)

if not os.path.exists(original_doc_path):
    print("paper.docx does not exist in the input folder. "
          "Please ensure this file exists under that name.")

# Load OpenAI API key from environment variable
api_key = os.getenv("OPENAI_API_KEY")
if not api_key:
    raise ValueError("OpenAI API key not found. Set OPENAI_API_KEY as an environment variable.")

openai.api_key = api_key

###############################################################################
# 1) Helper functions
###############################################################################
def is_heading(text):
    """
    Returns True if the text is likely a heading or subheading.
    Adjust or refine as needed for your specific doc style.
    """
    # Example checks:
    # 1) If it starts with a numeric heading like "2.1" or "3"
    if re.match(r"^\d+(\.\d+)*\s+\w+", text):
        return True
    # 2) If it is short and lacks typical sentence structure
    if len(text.split()) < 10 and text.count(".") <= 1:
        return True
    return False

def split_into_sentences(text):
    """
    Splits text on '.', '?', or '!', capturing punctuation separately.
    This is a simplistic approach. Improve as needed for edge cases.
    """
    parts = re.split(r'([.?!])', text)
    return parts

def reassemble_sentences(parts):
    """
    Re-join split text pieces into a single string, ensuring correct spacing
    and avoiding double punctuation.
    """
    result = []
    # parts is something like: [sentence_text, '.', next_sentence_text, '.', ...]
    for i in range(0, len(parts), 2):
        sentence = parts[i].strip()
        punctuation = parts[i+1] if i+1 < len(parts) else ""

        # If the edited sentence already ends with punctuation, remove it to avoid duplication
        if sentence and sentence[-1] in ".!?":
            sentence = sentence.rstrip(".!?")

        # Join the sentence and its punctuation
        if sentence:
            combined = sentence + punctuation
            # Trim any accidental double or triple dots (e.g. ".." -> ".")
            combined = re.sub(r"\.\.+", ".", combined)
            result.append(combined.strip())

    # Join them with a space, then do one final cleanup
    joined = " ".join(result)
    joined = re.sub(r"\s([.?!])", r"\1", joined)  # remove space before punctuation if any
    return joined.strip()

def edit_sentence_with_chatgpt(sentence, model=gpt_model):
    """
    Calls OpenAI with minimal editing instructions to fix grammar/spelling
    """
    if re.search(r"\(.*?\)", sentence) or re.search(r"\[.*?\]", sentence):
        return sentence

    # Also skip if trivially short
    if len(sentence.split()) < 3:
        return sentence

    system_prompt = (
        "You are a professional academic editor. Your job is to improve grammar, spelling, style, and professional language use "
        "while preserving original paragraph breaks and formatting.\n"
        "\n"
        "Readability & Clarity: Improve sentence structure, enhance logical flow, and eliminate unnecessary complexity "
        "while preserving academic rigor.\n"
        "Active Voice: Convert passive voice to active voice wherever possible, unless passive is necessary.\n"
        "Punctuation & Grammar: Correct punctuation, grammar, and syntax errors to ensure fluency.\n"
        "Consistency in Terminology & Style: Ensure uniform usage of terms and maintain consistent American English spelling.\n"
        "Precision & Objectivity: Remove vague language, strengthen claims with precise wording, and avoid subjective or exaggerated statements.\n"
        "Avoid Wordiness: Eliminate redundant words while preserving meaning.\n"
        "Logical Flow & Transitions: Improve coherence between sentences.\n"
        "\n"
        "Additionally, if a sentence contains any footnote marker at the end, or if it contains parentheses/brackets (likely references), "
        "do NOT edit that sentence. Leave such sentences entirely intact.\n"
        "\n"
        "Follow these rules meticulously:\n\n"
        "1) Do NOT split, merge, reorder, duplicate, or remove paragraphs. You must keep the exact paragraph structure.\n"
        "2) Do NOT alter or remove domain-specific terminology, technical terms, capitalized terms, or proper nouns.\n"
        "3) Do NOT delete or alter citations, equations, footnotes, bracketed references, or any text in parentheses or brackets. "
        "If a sentence contains them, skip editing that sentence.\n"
        "4) If a paragraph contains sentences with references or footnotes, skip those sentences entirely, unchanged.\n"
        "5) Preserve all numbers, author names, and citation formatting exactly as they are.\n"
        "6) Return only the corrected text, with no extra explanations, no new paragraph breaks, and no additional comments.\n"
        "7) Do not contract words (e.g., use 'do not' instead of 'don’t').\n"
        "8) Focus purely on grammar, style, and clarity—do not add or remove any content.\n"
    )

    try:
        response = openai.ChatCompletion.create(
            model=model,
            temperature = 0.1,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": sentence}
            ]
        )
        edited = response.choices[0].message.content.strip()

        # Post-processing step: remove any accidental trailing double punctuation
        edited = re.sub(r"\.\.+$", ".", edited)  # if it ends with multiple periods
        edited = edited.strip()
        return edited

    except Exception as e:
        print(f"⚠️  Error calling OpenAI API on sentence: {e}")
        return sentence

def edit_paragraph_sentencewise(paragraph_text, model=gpt_model):
    """
    Splits a paragraph into sentences, runs minimal-edits on each one,
    then reassembles them back into one paragraph string.
    """
    # If fewer than 2 sentences, skip
    if paragraph_text.count(".") < 1:
        return paragraph_text

    parts = split_into_sentences(paragraph_text)
    edited_parts = []

    for i in range(0, len(parts), 2):
        text_chunk = parts[i].strip() if i < len(parts) else ""
        punctuation = parts[i+1] if i+1 < len(parts) else ""
        if text_chunk:
            edited_text = edit_sentence_with_chatgpt(text_chunk, model=model)
            edited_parts.append(edited_text)
            edited_parts.append(punctuation)

    reassembled = reassemble_sentences(edited_parts)
    return reassembled

###############################################################################
# 2) Document processing logic
###############################################################################
doc = Document(original_doc_path)

processing     = False  # Flag to start processing after "Introduction"
found_abstract = False  # Track if we found the "Abstract" section
stop_keywords  = ["References", "Bibliography"]
start_keywords = ["Introduction"]  # Flexible introduction headings

paragraph_count = 1

# First pass: count how many paragraphs will be edited
for p in doc.paragraphs:
    text = p.text.strip()

    # Detect "Abstract" heading
    if re.match(r"^Abstract$", text, re.IGNORECASE):
        found_abstract = True
        continue

    # If we are in the paragraph right after Abstract
    if found_abstract and text.count(".") >= 2:
        paragraph_count += 1
        found_abstract = False  # Only do 1 paragraph after Abstract

    # If "Introduction" is in text, start editing
    if "Introduction" in text:
        processing = True
        continue

    # Stop if we see "References" or "Bibliography"
    if "References" in text or "Bibliography" in text:
        break

    # If processing, skip headings but count paragraphs with 2+ sentences
    if processing and not is_heading(text) and text.count(".") >= 2:
        paragraph_count += 1

processed_count = 0

# Second pass: actually edit relevant paragraphs
for p in doc.paragraphs:
    raw_text = p.text.strip()

    # Detect "Abstract" heading
    if re.match(r"^Abstract$", raw_text, re.IGNORECASE):
        found_abstract = True
        continue

    # Paragraph after Abstract
    if found_abstract and raw_text.count(".") >= 2:
        new_text = edit_paragraph_sentencewise(raw_text, model=gpt_model)
        p.text = new_text
        processed_count += 1
        print(f"Processed paragraph {processed_count}/{paragraph_count}")
        found_abstract = False

    if "Introduction" in raw_text:
        processing = True
        print("✅ Starting edits after 'Introduction'")
        continue

    if re.match(r"^\d*\.?\s*References$", raw_text, re.IGNORECASE) or raw_text in stop_keywords:
        print(f"Stopping edits at '{raw_text}'")
        break

    if processing and raw_text.count(".") >= 2 and not is_heading(raw_text):
        new_text = edit_paragraph_sentencewise(raw_text, model=gpt_model)
        p.text = new_text
        processed_count += 1
        print(f"Processed paragraph {processed_count}/{paragraph_count}")

print("Editing complete. Saving document...")

doc.save(edited_doc_path)
print(f"Document saved to {edited_doc_path}")

###############################################################################
# 3) Compare documents in Word
###############################################################################
def compare_documents(original, edited, output):
    """
    Automates Microsoft Word's 'Compare Documents' function while ignoring citation changes.
    Ensures Word properly opens both files before attempting comparison.
    """
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False

        # Open Original Document
        try:
            original_doc = word.Documents.Open(original)
        except Exception as e:
            print(f"⚠️ Error opening original document: {e}")
            word.Quit()
            return
        
        # Open Edited Document
        try:
            edited_doc = word.Documents.Open(edited)
        except Exception as e:
            print(f"⚠️ Error opening edited document: {e}")
            original_doc.Close(False)
            word.Quit()
            return

        # Perform document comparison
        try:
            compared_doc = word.CompareDocuments(
                OriginalDocument=original_doc,
                RevisedDocument=edited_doc,
                CompareFormatting=False,
                IgnoreAllComparisonWarnings=True
            )
        except Exception as e:
            print(f"⚠️ Word comparison failed: {e}")
            original_doc.Close(False)
            edited_doc.Close(False)
            word.Quit()
            return

        # Reject citation-related changes
        for change in compared_doc.Revisions:
            txt = change.Range.Text
            if re.match(r"\(.*\)", txt):
                change.Reject()
            elif re.match(r"\[\d+\]", txt):
                change.Reject()

        compared_doc.SaveAs(output, FileFormat=16)

        compared_doc.Close(False)
        original_doc.Close(False)
        edited_doc.Close(False)

        word.Quit()
        print(f"✅ Comparison completed. Document saved to: {output}")

    except Exception as e:
        print(f"❌ Critical error comparing documents: {e}")
        word.Quit()

compare_documents(original_doc_path, edited_doc_path, output_doc_path)
