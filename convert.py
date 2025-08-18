from pathlib import Path
import pandas as pd
from docx import Document
import re

# === File paths ===
current_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
input_file = current_dir / "input.docx"
output_file = current_dir / "output.xlsx"
images_dir = current_dir / "images"
images_dir.mkdir(exist_ok=True)

# Load Word document
doc = Document(input_file)
data = []

# === Regex patterns ===
module_pattern = re.compile(r"Module\s*(\d+):\s*(.+)", re.IGNORECASE)
question_num_pattern = re.compile(r"^(\d+)\.$")
english_option_pattern = re.compile(r"^[ABCD]$")
math_option_start_pattern = re.compile(r"^\(?([ABCD])\)?[.)]\s*(.*)$")

# === State variables ===
current_module = ""
current_module_num = ""
current_module_topic = ""
current_passage_paragraphs = []
current_question_num = ""
current_question_text = ""
current_question_text_lines = []
options = {"A": "", "B": "", "C": "", "D": ""}
state = "start"
current_option = None
option_format = None  # "english" or "math"

question_has_image = False
question_images = []

# Keep track of image index per question
image_idx = 1

# === Helper: extract images from paragraph ===
def extract_images_from_para(para, module_num, qnum):
    global image_idx
    imgs = []
    for run in para.runs:
        blips = run.element.xpath(".//a:blip")
        for blip in blips:
            rId = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
            if not rId:
                continue
            image_part = doc.part.related_parts[rId]
            image_bytes = image_part.blob

            # Save image with module_question_index format
            fname = f"module{module_num}_question{qnum}_{image_idx}.png"
            fpath = images_dir / fname
            with open(fpath, "wb") as f:
                f.write(image_bytes)

            # Store relative path for Excel
            imgs.append(str(fpath.relative_to(current_dir)))
            image_idx += 1
    return imgs

# === Save question to data list ===
def save_question():
    global data, question_has_image, question_images, image_idx
    if not current_question_num:
        return

    passage_text = "\n".join(current_passage_paragraphs).strip()
    math_question_text = "\n".join(current_question_text_lines).strip()
    qtext = current_question_text.strip() if option_format == "english" else ""
    passage = passage_text if option_format == "english" else math_question_text

    data.append([
        current_module,
        passage,
        current_question_num,
        qtext,
        options.get("A", ""),
        options.get("B", ""),
        options.get("C", ""),
        options.get("D", ""),
        "",  # Answer
        "; ".join(question_images) if question_has_image else "",  # Additional Picture (relative path)
        "",  # Note
        ""   # Explanation
    ])

    question_has_image = False
    question_images.clear()
    image_idx = 1  # reset for next question

# === Reset question variables ===
def reset_question():
    global current_passage_paragraphs, current_question_text, current_question_text_lines, options, current_option
    current_passage_paragraphs.clear()
    current_question_text = ""
    current_question_text_lines.clear()
    options = {"A": "", "B": "", "C": "", "D": ""}
    current_option = None

# === Main loop over paragraphs ===
for para in doc.paragraphs:
    text = para.text.strip()

    # Extract images if question is active
    if current_module_num and current_question_num:
        imgs = extract_images_from_para(para, current_module_num, current_question_num)
        if imgs:
            question_has_image = True
            question_images.extend(imgs)

    if not text:
        continue

    # Module detection
    m_mod = module_pattern.match(text)
    if m_mod:
        save_question()
        current_module_num = m_mod.group(1)
        current_module = f"Module {current_module_num}: {m_mod.group(2)}"
        current_module_topic = m_mod.group(2).strip().lower()
        reset_question()
        current_question_num = ""
        state = "start"
        option_format = "english" if "english" in current_module_topic else "math"
        continue

    # Question number detection
    m_qnum = question_num_pattern.match(text)
    if m_qnum:
        save_question()
        reset_question()
        current_question_num = m_qnum.group(1)
        state = "reading_passage" if option_format == "english" else "reading_question_math"
        continue

    # === English question parsing ===
    if option_format == "english":
        if state == "reading_passage":
            if english_option_pattern.match(text):
                if current_passage_paragraphs:
                    current_question_text = current_passage_paragraphs.pop(-1)
                state = "reading_options"
                current_option = text
                options[current_option] = ""
            else:
                current_passage_paragraphs.append(text)
            continue
        elif state == "reading_options":
            if english_option_pattern.match(text):
                current_option = text
                options[current_option] = ""
            elif current_option:
                options[current_option] += (" " if options[current_option] else "") + text
            continue

    # === Math question parsing ===
    else:
        if state == "reading_question_math":
            m_opt_start = math_option_start_pattern.match(text)
            if m_opt_start:
                opt_letter, opt_text = m_opt_start.group(1), m_opt_start.group(2).strip()
                options[opt_letter] = opt_text
                current_option = opt_letter
                state = "reading_options_math"
            else:
                current_question_text_lines.append(text)
            continue
        elif state == "reading_options_math":
            m_opt_start = math_option_start_pattern.match(text)
            if m_opt_start:
                current_option, opt_text = m_opt_start.group(1), m_opt_start.group(2).strip()
                options[current_option] = opt_text
            elif current_option:
                options[current_option] += (" " if options[current_option] else "") + text
            continue

# Save the last question
save_question()

# === Export to Excel ===
columns = [
    "Section", "Passage Content", "Question Number", "Question Text",
    "Option A", "Option B", "Option C", "Option D",
    "Answer", "Additional Picture", "Note", "Explanation"
]
df = pd.DataFrame(data, columns=columns)
df.to_excel(output_file, index=False)

print(f"Excel saved: {output_file}")
print(f"Extracted images in: {images_dir}")
