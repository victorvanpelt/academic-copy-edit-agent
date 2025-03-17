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
        try:
            os.remove(path)
        except PermissionError:
            print(f"❌ Error: The file '{path}' is currently open or locked. Please close it and try again.")
        except Exception as e:
            print(f"⚠️ Unexpected error deleting '{path}': {e}")

if not os.path.exists(original_doc_path):
    print("❌ Error: 'paper.docx' does not exist in the input folder. "
          "Please ensure this file is named correctly and placed in the '0_input' directory.")

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
        "You are a professional academic editor. Improve grammar, spelling, and style while preserving paragraph breaks. "
        "Follow these rules strictly:\n"
        "1) Readability & Clarity: refine sentence structure, enhance logical flow, and remove unnecessary complexity (maintain academic rigor).\n"
        "2) Active Voice: convert passive to active whenever possible, unless truly needed.\n"
        "3) Punctuation & Grammar: correct errors for fluency.\n"
        "4) Consistency & Style: keep terms uniform, use consistent American English spelling.\n"
        "5) Precision & Objectivity: remove vague language, strengthen claims, avoid subjectivity.\n"
        "6) Avoid Wordiness: cut redundant words while preserving meaning.\n"
        "7) Logical Flow & Transitions: ensure coherent transitions between sentences.\n"
        "8) If a sentence has footnotes at the end or parentheses/brackets (references), skip editing that sentence. Leave it intact.\n"
        "9) Do NOT merge, split, or reorder paragraphs. Preserve domain terminology, citations, numbers, and equations.\n"
        "10) Use typographic (curly) apostrophes (’ instead of ').\n"
        "11) Return only the corrected text, with no explanations or new paragraph breaks.\n"
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

found_abstract = False  # Track if we found the "Abstract" section
processing     = False  # Flag to start processing after "Introduction"
stop_keywords  = ["References", "Bibliography"]
start_keywords = ["Introduction"]

paragraph_count = 1

# First pass: count how many paragraphs will be edited
for p in doc.paragraphs:
    text = p.text.strip()

    # Detect "Abstract" heading
    if re.match(r"^Abstract$", text, re.IGNORECASE):
        found_abstract = True
        print("✅ Found the Abstract")
        continue

    # If we are in the paragraph right after Abstract
    if found_abstract and text.count(".") >= 2:
        paragraph_count += 1
        found_abstract = False  # Only do 1 paragraph after Abstract

    # If "Introduction" is in text, start editing
    if re.match(r"^Introduction$", text, re.IGNORECASE):
        processing = True
        print("✅ Now counting paragraphs after 'Introduction'")
        continue


    # Stop if we see "References" or "Bibliography"
    if re.match(r"^\d*\.?\s*References$", text, re.IGNORECASE) or text in stop_keywords:
        print(f"✅ Stopping counting at '{text}'")
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
    if found_abstract and not processing and raw_text.count(".") >= 2:
        new_text = edit_paragraph_sentencewise(raw_text, model=gpt_model)
        p.text = new_text
        processed_count += 1
        print(f"Processed paragraph {processed_count}/{paragraph_count}")

    if re.match(r"^Introduction$", raw_text, re.IGNORECASE):
        processing = True
        print("✅ Now processing paragraphs after 'Introduction'")
        continue

    if re.match(r"^\d*\.?\s*References$", raw_text, re.IGNORECASE) or raw_text in stop_keywords:
        print(f"✅ Stopping processing at '{raw_text}'")
        break

    if processing and raw_text.count(".") >= 2 and not is_heading(raw_text):
        new_text = edit_paragraph_sentencewise(raw_text, model=gpt_model)
        p.text = new_text
        processed_count += 1
        print(f"Processed paragraph {processed_count}/{paragraph_count}")

print("✅ Editing complete. Saving document...")

doc.save(edited_doc_path)
print(f"✅ Document saved to {edited_doc_path}")

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

        # Reject footnote-related changes
        for footnote in compared_doc.Footnotes:
            for change in footnote.Range.Revisions:
                change.Reject()  # Reject any modifications in footnotes

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
