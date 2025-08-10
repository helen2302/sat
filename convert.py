# from pathlib import Path
# import pandas as pd
# from docx import Document
# import re

# # === Paths ===
# current_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
# input_file = current_dir / "input.docx"
# output_file = current_dir / "output.xlsx"

# doc = Document(input_file)

# data = []

# module_pattern = re.compile(r"Module\s*(\d+):\s*(.+)", re.IGNORECASE)
# question_num_pattern = re.compile(r"^(\d+)\.$")

# # Option patterns
# english_option_pattern = re.compile(r"^[ABCD]$")
# math_option_start_pattern = re.compile(r"^\(?([ABCD])\)?\.?\s*(.*)$")  # (A) or A. or A with optional text
# math_option_start_pattern = re.compile(r"^\(?([ABCD])\)?[.)]\s*(.*)$")

# current_module = ""
# current_module_topic = ""
# current_passage_paragraphs = []  # list of passage paragraphs (English)
# current_question_num = ""
# current_question_text = ""  # For English, a single paragraph with '?'
# current_question_text_lines = []  # For math, multiple lines
# options = {"A": "", "B": "", "C": "", "D": ""}
# state = "start"
# current_option = None
# option_format = None  # "english" or "math"

# def save_question():
#     global current_question_num, current_module, current_passage_paragraphs, current_question_text, current_question_text_lines, options, data

#     passage_text = "\n".join(current_passage_paragraphs).strip()
#     math_question_text = "\n".join(current_question_text_lines).strip()

#     if not current_question_num:
#         return

#     if option_format == "english":
#         qtext = current_question_text.strip()
#         passage = passage_text
#     else:
#         qtext = math_question_text
#         passage = ""

#     data.append([
#         current_question_num,
#         current_module,
#         passage,
#         qtext,
#         options.get("A", ""),
#         options.get("B", ""),
#         options.get("C", ""),
#         options.get("D", ""),
#         "",
#         ""  # Additional Picture placeholder
#     ])

# def reset_question():
#     global current_passage_paragraphs, current_question_text, current_question_text_lines, options, current_option
#     current_passage_paragraphs = []
#     current_question_text = ""
#     current_question_text_lines = []
#     options = {"A": "", "B": "", "C": "", "D": ""}
#     current_option = None

# for para in doc.paragraphs:
#     text = para.text.strip()
#     if not text:
#         continue

#     # Detect module line
#     m_mod = module_pattern.match(text)
#     if m_mod:
#         save_question()
#         current_module = f"Module {m_mod.group(1)}: {m_mod.group(2)}"
#         current_module_topic = m_mod.group(2).strip().lower()
#         reset_question()
#         current_question_num = ""
#         state = "start"
#         # Decide option format by module topic
#         if "english" in current_module_topic:
#             option_format = "english"
#         elif "math" in current_module_topic:
#             option_format = "math"
#         else:
#             option_format = "english"  # safer default
#         continue

#     # Detect question number line
#     m_qnum = question_num_pattern.match(text)
#     if m_qnum:
#         save_question()
#         reset_question()
#         current_question_num = m_qnum.group(1)
#         # For math: start reading question text immediately
#         if option_format == "english":
#             state = "reading_passage"
#         else:
#             state = "reading_question_math"
#         continue

#     # English module processing
#     if option_format == "english":
#         if state == "reading_passage":
#             # Accumulate paragraphs as passage until we find options start
#             if english_option_pattern.match(text):
#                 # Find question text paragraph in passage paragraphs (first with '?')
#                 question_idx = -1
#                 for i, p_text in enumerate(current_passage_paragraphs):
#                     if "?" in p_text:
#                         question_idx = i
#                         break

#                 if question_idx >= 0:
#                     current_question_text = current_passage_paragraphs[question_idx]
#                     current_passage_paragraphs = current_passage_paragraphs[:question_idx]
#                 else:
#                     current_question_text = ""

#                 # Now start reading options
#                 state = "reading_options"
#                 current_option = text
#                 options[current_option] = ""
#                 continue
#             else:
#                 current_passage_paragraphs.append(text)
#                 continue

#         elif state == "reading_options":
#             if english_option_pattern.match(text):
#                 current_option = text
#                 options[current_option] = ""
#                 continue
#             elif current_option:
#                 # Append multi-line option text
#                 if options[current_option]:
#                     options[current_option] += " " + text
#                 else:
#                     options[current_option] = text
#                 continue

#     # Math module processing
#     else:
#         if state == "reading_question_math":
#             m_opt_start = math_option_start_pattern.match(text)
#             if m_opt_start:
#                 # Found options start, switch to reading options
#                 opt_letter = m_opt_start.group(1)
#                 opt_text = m_opt_start.group(2).strip()
#                 options[opt_letter] = opt_text
#                 current_option = opt_letter
#                 state = "reading_options_math"
#             else:
#                 # Append question text lines (multiple lines)
#                 current_question_text_lines.append(text)
#             continue

#         elif state == "reading_options_math":
#             m_opt_start = math_option_start_pattern.match(text)
#             if m_opt_start:
#                 current_option = m_opt_start.group(1)
#                 opt_text = m_opt_start.group(2).strip()
#                 options[current_option] = opt_text
#             elif current_option:
#                 # Append multiline option text
#                 if options[current_option]:
#                     options[current_option] += " " + text
#                 else:
#                     options[current_option] = text
#             continue

# # Save last question after loop ends
# save_question()

# columns = [
#     "Question Number",
#     "Section",
#     "Passage Content",
#     "Question Text",
#     "Option A",
#     "Option B",
#     "Option C",
#     "Option D",
#     "Correct Answer",
#     "Additional Picture"
# ]

# df = pd.DataFrame(data, columns=columns)
# df.to_excel(output_file, index=False)

# print(f"✅ Excel file saved to: {output_file}")

from pathlib import Path
import pandas as pd
from docx import Document
import re

# === Paths ===
current_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
input_file = current_dir / "input.docx"
output_file = current_dir / "output.xlsx"

doc = Document(input_file)

data = []

module_pattern = re.compile(r"Module\s*(\d+):\s*(.+)", re.IGNORECASE)
question_num_pattern = re.compile(r"^(\d+)\.$")

english_option_pattern = re.compile(r"^[ABCD]$")
math_option_start_pattern = re.compile(r"^\(?([ABCD])\)?[.)]\s*(.*)$")

current_module = ""
current_module_topic = ""
current_passage_paragraphs = []
current_question_num = ""
current_question_text = ""
current_question_text_lines = []
options = {"A": "", "B": "", "C": "", "D": ""}
state = "start"
current_option = None
option_format = None

question_has_image = False  # NEW: track if any part of the question has an image


def para_has_image(para):
    """Check if a paragraph contains an image"""
    for run in para.runs:
        if run.element.xpath(".//pic:pic") or run.element.xpath(".//a:blip"):
            return True
    return False


def save_question():
    global data, question_has_image
    passage_text = "\n".join(current_passage_paragraphs).strip()
    math_question_text = "\n".join(current_question_text_lines).strip()

    if not current_question_num:
        return

    if option_format == "english":
        qtext = current_question_text.strip()
        passage = passage_text
    else:
        qtext = math_question_text
        passage = ""

    data.append([
        current_question_num,
        current_module,
        passage,
        qtext,
        options.get("A", ""),
        options.get("B", ""),
        options.get("C", ""),
        options.get("D", ""),
        "",
        "Yes" if question_has_image else "No"
    ])

    question_has_image = False  # reset for next question


def reset_question():
    global current_passage_paragraphs, current_question_text, current_question_text_lines, options, current_option
    current_passage_paragraphs = []
    current_question_text = ""
    current_question_text_lines = []
    options = {"A": "", "B": "", "C": "", "D": ""}
    current_option = None


for para in doc.paragraphs:
    text = para.text.strip()

    # Detect if paragraph has image
    if para_has_image(para):
        question_has_image = True

    if not text:
        continue

    # Detect module
    m_mod = module_pattern.match(text)
    if m_mod:
        save_question()
        current_module = f"Module {m_mod.group(1)}: {m_mod.group(2)}"
        current_module_topic = m_mod.group(2).strip().lower()
        reset_question()
        current_question_num = ""
        state = "start"
        option_format = "english" if "english" in current_module_topic else "math"
        continue

    # Detect question number
    m_qnum = question_num_pattern.match(text)
    if m_qnum:
        save_question()
        reset_question()
        current_question_num = m_qnum.group(1)
        state = "reading_passage" if option_format == "english" else "reading_question_math"
        continue

    # English processing
    if option_format == "english":
        if state == "reading_passage":
            if english_option_pattern.match(text):
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
                current_passage_paragraphs.append(text)
                continue

        elif state == "reading_options":
            if english_option_pattern.match(text):
                current_option = text
                options[current_option] = ""
                continue
            elif current_option:
                if options[current_option]:
                    options[current_option] += " " + text
                else:
                    options[current_option] = text
                continue

    # Math processing
    else:
        if state == "reading_question_math":
            m_opt_start = math_option_start_pattern.match(text)
            if m_opt_start:
                opt_letter = m_opt_start.group(1)
                opt_text = m_opt_start.group(2).strip()
                options[opt_letter] = opt_text
                current_option = opt_letter
                state = "reading_options_math"
            else:
                current_question_text_lines.append(text)
            continue

        elif state == "reading_options_math":
            m_opt_start = math_option_start_pattern.match(text)
            if m_opt_start:
                current_option = m_opt_start.group(1)
                opt_text = m_opt_start.group(2).strip()
                options[current_option] = opt_text
            elif current_option:
                if options[current_option]:
                    options[current_option] += " " + text
                else:
                    options[current_option] = text
            continue

# Save last question
save_question()

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
    "Has Image"
]

df = pd.DataFrame(data, columns=columns)
df.to_excel(output_file, index=False)

print(f"✅ Excel file saved to: {output_file}")

