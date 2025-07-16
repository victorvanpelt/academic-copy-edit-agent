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

gpt_model = os.getenv("GPT_MODEL", "gpt-4o")

for path in [edited_doc_path, output_doc_path]:
    if os.path.exists(path):
        try:
            os.remove(path)
        except PermissionError:
            print(f"Close '{path}' and retry.")
        except Exception as e:
            print(f"Couldn't delete '{path}': {e}")

if not os.path.exists(original_doc_path):
    raise FileNotFoundError("Missing 0_input/paper.docx")

api_key = os.getenv("OPENAI_API_KEY")
if not api_key:
    raise ValueError("Set OPENAI_API_KEY env var")
openai.api_key = api_key

###############################################################################
# Helper functions
###############################################################################

def is_heading(text: str) -> bool:
    """Heuristic heading detector."""
    if re.match(r"^\d+(?:\.\d+)*\s+\w+", text):
        return True
    return len(text.split()) < 10 and text.count(".") <= 1


SENT_SPLIT = re.compile(r"([.?!])")


def split_into_sentences(text: str):
    return SENT_SPLIT.split(text)


def reassemble(parts):
    joined = " ".join(parts)
    joined = re.sub(r"\s+([.?!])", r"\1", joined)
    joined = re.sub(r"([.?!]){2,}", r"\1", joined)
    return joined.strip()


CITE_PATTERNS = [
    re.compile(r"\(.*?\)"),      # (Smith, 2022)
    re.compile(r"\[.*?\]"),      # [15]
    re.compile(r"\{.*?\}"),      # {Smith, 2022 #45}  (EndNote temporary)
]


def contains_citation(text: str) -> bool:
    return any(p.search(text) for p in CITE_PATTERNS)


def edit_sentence_with_chatgpt(sentence: str, model: str = gpt_model) -> str:
    if contains_citation(sentence) or len(sentence.split()) < 3:
        return sentence

    prompt = (
        "You are a professional academic copy editor. Improve grammar, spelling, "
        "concision, clarity, and academic style in American English while "
        "preserving meaning and terminology.\n"
        "Rules: 1) Do NOT change citations or footnotes. 2) Do NOT merge, split, "
        "or reorder sentences. 3) Return ONLY the corrected sentence."
    )

    try:
        resp = openai.chat.completions.create(
            model=model,
            temperature=0,
            messages=[
                {"role": "system", "content": prompt},
                {"role": "user", "content": sentence},
            ],
        )
        return resp.choices[0].message.content.strip()
    except Exception as e:
        print(f"⚠️  OpenAI error: {e}")
        return sentence


def edit_paragraph(paragraph_text: str) -> str:
    parts    = split_into_sentences(paragraph_text)
    edited   = []
    for i in range(0, len(parts), 2):
        chunk = parts[i].strip()
        punct = parts[i + 1] if i + 1 < len(parts) else ""
        if chunk:
            edited.append(edit_sentence_with_chatgpt(chunk))
            edited.append(punct)
    return reassemble(edited)

###############################################################################
# Document processing
###############################################################################

doc = Document(original_doc_path)
processing      = False  # becomes True after Introduction
after_abstract  = False  # edit the first paragraph after "Abstract"
stop_keywords   = {"references", "bibliography"}
count           = 0

print("Starting copy‑edit…")

for para in doc.paragraphs:
    text = para.text.strip()

    # --- Section boundary logic ----------------------------------------
    if re.match(r"^(?:\d+\.?)?\s*Abstract$", text, re.IGNORECASE):
        after_abstract = True
        print("   ↳ Found 'Abstract' heading")
        continue

    if after_abstract and text and not is_heading(text):
        para.text = edit_paragraph(text)
        count += 1
        after_abstract = False  # only edit the first paragraph after Abstract
        print("      • Edited paragraph in Abstract section")
        continue

    if re.match(r"^(?:\d+\.?)?\s*Introduction$", text, re.IGNORECASE):
        processing = True
        print("   ↳ Entering main body after 'Introduction'")
        continue

    if text.lower() in stop_keywords or re.match(r"^\d*\.?\s*References$", text, re.IGNORECASE):
        print(f"   ↳ Reached '{text}', stopping edits")
        break

    # ------------------------------------------------------------------
    if processing and text and not is_heading(text):
        para.text = edit_paragraph(text)
        count += 1
        print(f"      • Edited paragraph {count}")

print("Edited", count, "paragraphs. Saving…")
doc.save(edited_doc_path)
print("Saved to", edited_doc_path)

###############################################################################
# Optional Word comparison (Windows only)
###############################################################################

def compare_docs(orig: str, edited: str, output: str):
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        o = word.Documents.Open(orig)
        e = word.Documents.Open(edited)
        c = word.CompareDocuments(o, e, CompareFormatting=False, IgnoreAllComparisonWarnings=True)
        c.SaveAs(output, FileFormat=16)
        c.Close(False); o.Close(False); e.Close(False); word.Quit()
        print("Track‑changes doc saved to", output)
    except Exception as exc:
        print("ℹ️  Word compare skipped:", exc)

try:
    compare_docs(original_doc_path, edited_doc_path, output_doc_path)
except Exception:
    pass

print("All done!")
