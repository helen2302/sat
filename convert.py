from pathlib import Path
import pandas as pd
from docx import Document
import re

# === File paths ===
# Use the current directory of this script if available, otherwise use working directory
current_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
input_file = current_dir / "input.docx"      # Input Word file
output_file = current_dir / "output.xlsx"    # Output Excel file

# Load the Word document
doc = Document(input_file)

# This will store all processed question data
data = []

# === Regex patterns ===
module_pattern = re.compile(r"Module\s*(\d+):\s*(.+)", re.IGNORECASE)   # Detect "Module X: Topic"
question_num_pattern = re.compile(r"^(\d+)\.$")                         # Detect "1.", "2.", etc.
english_option_pattern = re.compile(r"^[ABCD]$")                        # Detect single letter option for English
math_option_start_pattern = re.compile(r"^\(?([ABCD])\)?[.)]\s*(.*)$")  # Detect "A. text" or "(B) text" for Math

# === State tracking variables ===
current_module = ""
current_module_topic = ""
current_passage_paragraphs = []
current_question_num = ""
current_question_text = ""
current_question_text_lines = []
options = {"A": "", "B": "", "C": "", "D": ""}
state = "start"
current_option = None
option_format = None  # "english" or "math"

question_has_image = False  # Track if the current question contains any image


# === Helper functions ===

def para_has_image(para):
    """Check if a paragraph contains an embedded image"""
    for run in para.runs:
        # Detect image elements in Word XML
        if run.element.xpath(".//pic:pic") or run.element.xpath(".//a:blip"):
            return True
    return False


def save_question():
    """Save the current question into the data list"""
    global data, question_has_image

    # Join collected text parts
    passage_text = "\n".join(current_passage_paragraphs).strip()
    math_question_text = "\n".join(current_question_text_lines).strip()

    # Skip if there's no question number
    if not current_question_num:
        return

    # English questions have passage + question text
    # Math questions have only question text (no passage)
    if option_format == "english":
        qtext = current_question_text.strip()
        passage = passage_text
    else:
        qtext = math_question_text
        passage = ""

    # Append row to the data list
    data.append([
        current_question_num,
        current_module,
        passage,
        qtext,
        options.get("A", ""),
        options.get("B", ""),
        options.get("C", ""),
        options.get("D", ""),
        "",  # Placeholder for correct answer
        "Yes" if question_has_image else "No"  # Image presence flag
    ])

    # Reset image tracker for next question
    question_has_image = False


def reset_question():
    """Reset all question-related variables before reading a new question"""
    global current_passage_paragraphs, current_question_text, current_question_text_lines, options, current_option
    current_passage_paragraphs = []
    current_question_text = ""
    current_question_text_lines = []
    options = {"A": "", "B": "", "C": "", "D": ""}
    current_option = None


# === Main parsing loop ===
for para in doc.paragraphs:
    text = para.text.strip()

    # Detect if this paragraph contains an image
    if para_has_image(para):
        question_has_image = True

    # Skip empty paragraphs
    if not text:
        continue

    # Detect new module (e.g., "Module 1: English - Reading")
    m_mod = module_pattern.match(text)
    if m_mod:
        save_question()  # Save any question before starting new module
        current_module = f"Module {m_mod.group(1)}: {m_mod.group(2)}"
        current_module_topic = m_mod.group(2).strip().lower()
        reset_question()
        current_question_num = ""
        state = "start"
        # Decide question type based on module name
        option_format = "english" if "english" in current_module_topic else "math"
        continue

    # Detect new question number (for ex: "5.")
    m_qnum = question_num_pattern.match(text)
    if m_qnum:
        save_question()  # Save any previous question
        reset_question()
        current_question_num = m_qnum.group(1)
        state = "reading_passage" if option_format == "english" else "reading_question_math"
        continue

    # === English question processing ===
    if option_format == "english":
        if state == "reading_passage":
            # Detect start of options
            if english_option_pattern.match(text):
                # Extract question from passage if it contains '?'
                question_idx = -1
                for i, p_text in enumerate(current_passage_paragraphs):
                    if "?" in p_text:
                        question_idx = i
                        break
                if question_idx >= 0:
                    current_question_text = current_passage_paragraphs[question_idx]
                    current_passage_paragraphs = current_passage_paragraphs[:question_idx]
                else:
                    current_question_text = ""
                state = "reading_options"
                current_option = text
                options[current_option] = ""
                continue
            else:
                # Still reading passage text
                current_passage_paragraphs.append(text)
                continue

        elif state == "reading_options":
            # New option letter
            if english_option_pattern.match(text):
                current_option = text
                options[current_option] = ""
                continue
            # Append to the current option text
            elif current_option:
                options[current_option] += (" " if options[current_option] else "") + text
                continue

    # === Math question processing ===
    else:
        if state == "reading_question_math":
            m_opt_start = math_option_start_pattern.match(text)
            if m_opt_start:
                # First option found
                opt_letter = m_opt_start.group(1)
                opt_text = m_opt_start.group(2).strip()
                options[opt_letter] = opt_text
                current_option = opt_letter
                state = "reading_options_math"
            else:
                # Still reading question text
                current_question_text_lines.append(text)
            continue

        elif state == "reading_options_math":
            m_opt_start = math_option_start_pattern.match(text)
            if m_opt_start:
                # Found new option
                current_option = m_opt_start.group(1)
                opt_text = m_opt_start.group(2).strip()
                options[current_option] = opt_text
            elif current_option:
                # Append to the current option text
                options[current_option] += (" " if options[current_option] else "") + text
            continue

# Save the final question after the loop ends
save_question()

# === Create DataFrame and export to Excel ===
columns = [
    "Question Number",
    "Section",
    "Passage Content",
    "Question Text",
    "Option A",
    "Option B",
    "Option C",
    "Option D",
    "Correct Answer",
    "Additional Image"
]

df = pd.DataFrame(data, columns=columns)
df.to_excel(output_file, index=False)

print(f"✅ Excel file saved to: {output_file}")
