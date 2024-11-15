from docx import Document
from docx.shared import Pt
import sqlite3
import random
from datetime import datetime
from docx2pdf import convert

def replace_top_text(doc, subject):
    """Replace the top text in the document based on the selected subject."""
    for para in doc.paragraphs:
        if "22MCA101" in para.text:
            para.clear()  # Clear the existing paragraph
            title_run = set_font(para, 'Times New Roman', 13, bold=True)  # Set new title with desired style
            title_run.text = "Course Code: 22MCA204" if subject == "java" else "Course Code: 22MCA101"
            break  # Stop after replacing the title

# Function to extract 10-mark questions from a specific module in the database
def extract_10_mark_questions(db_path, module_name):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    query = f"SELECT question_name, marks FROM {module_name} WHERE marks = 10"
    cursor.execute(query)
    questions = cursor.fetchall()
    conn.close()
    return questions

# Function to set font style
def set_font(para, font_name, size, bold=False):
    run = para.add_run()  # Add a new run to the paragraph
    run.font.name = font_name
    run.font.size = Pt(size)
    run.bold = bold
    return run

# Function to replace keywords in the DOCX with questions (without question numbers)
# and set the font to Times New Roman, size 13
def replace_keywords_in_doc(doc, questions_by_module):
    for para in doc.paragraphs:
        for keyword, questions in questions_by_module.items():
            if questions:  # Ensure there are questions left in the module
                question = random.choice(questions)  # Select a random question
                question_text = f"{question[0]}"  # Only keep the question text without the question number
                
                # Replace the keyword with the selected question
                if keyword in para.text:
                    # Clear the paragraph and add a new run with the correct font settings
                    para.clear()
                    run = para.add_run(question_text)
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(13)
                    
                    questions.remove(question)  # Remove the selected question to avoid repetition

    # Check tables if any keywords are present there
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for keyword, questions in questions_by_module.items():
                    if questions:
                        question = random.choice(questions)  # Select a random question
                        question_text = f"{question[0]}"  # Only keep the question text without the question number
                        if keyword in cell.text:
                            # Clear the cell and add new text with the correct font settings
                            cell.text = ""  # Clear the cell content
                            run = cell.paragraphs[0].add_run(question_text)
                            run.font.name = 'Times New Roman'
                            run.font.size = Pt(13)

                            questions.remove(question)  # Remove the selected question to avoid repetition


# Function to replace CO text based on subject
def replace_co_text(doc, subject):
    if subject == 'mfca':
        co_replacements = {
            "C011": "Apply the concepts of set theory and propositional logic to solve problems.",
            "C022": "Compute statistical measures of random variables and probability distributions.",
            "C033": "Solve problems using concepts of relations.",
            "C044": "Apply the abstract concepts of graph theory to solve problems.",
            "C055": "Fit an appropriate curve for the given data."
        }
    elif subject == 'java':
        co_replacements = {
            "C011": "Demonstrate the basic programming constructs of Java and OOP concepts to develop Java applications.",
            "C022": "Illustrate the concepts of generalization and run time polymorphism to develop reusable components.",
            "C033": "Exemplify the usage of Multithreading in building efficient applications.",
            "C044": "Build web applications using Servlets and JSP.",
            "C055": "Design applications using JDBC and Enterprise Java Beans."
        }
    else:
        return  # No action for invalid subjects

    # Replace in paragraphs
    for para in doc.paragraphs:
        for co, replacement_text in co_replacements.items():
            if co in para.text:
                para.text = para.text.replace(co, replacement_text)

    # Replace in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for co, replacement_text in co_replacements.items():
                    if co in cell.text:
                        cell.text = cell.text.replace(co, replacement_text)


def main():
    try:
        # Ask the user to choose the subject
        subject = input("Choose the subject (Java or MFCA): ").strip().lower()

        # Set database path based on user input
        if subject == 'java':
            db_path = './database/java.db'
        elif subject == 'mfca':
            db_path = './database/MFCA.db'
        else:
            print("Invalid subject chosen. Please choose either 'Java' or 'MFCA'.")
            return

        # Extract 10-mark questions for each module from the chosen database
        questions_by_module = {
            "M11": extract_10_mark_questions(db_path, 'module1'),
            "M12": extract_10_mark_questions(db_path, 'module1'),
            "M13": extract_10_mark_questions(db_path, 'module1'),
            "M14": extract_10_mark_questions(db_path, 'module1'),
            "M15": extract_10_mark_questions(db_path, 'module1'),
            "M21": extract_10_mark_questions(db_path, 'module2'),
            "M22": extract_10_mark_questions(db_path, 'module2'),
            "M23": extract_10_mark_questions(db_path, 'module2'),
            "M24": extract_10_mark_questions(db_path, 'module2'),
            "M31": extract_10_mark_questions(db_path, 'module3'),
            "M32": extract_10_mark_questions(db_path, 'module3'),
            "M33": extract_10_mark_questions(db_path, 'module3'),
            "M34": extract_10_mark_questions(db_path, 'module3'),
            "M41": extract_10_mark_questions(db_path, 'module4'),
            "M42": extract_10_mark_questions(db_path, 'module4'),
            "M43": extract_10_mark_questions(db_path, 'module4'),
            "M44": extract_10_mark_questions(db_path, 'module4'),
            "M51": extract_10_mark_questions(db_path, 'module5'),
            "M52": extract_10_mark_questions(db_path, 'module5'),
            "M53": extract_10_mark_questions(db_path, 'module5'),
            "M54": extract_10_mark_questions(db_path, 'module5'),
        }

        # Load the DOCX template
        template_path = './pattern/pattern.docx'
        doc = Document(template_path)

        # Update the title based on the chosen subject
        for para in doc.paragraphs:
            if "MATHEMATICAL FOUNDATION FOR COMPUTER APPLICATIONS" in para.text:
                para.clear()  # Clear the existing paragraph
                title_run = set_font(para, 'Times New Roman', 16, bold=True)  # Set new title with desired style
                title_run.text = "JAVA PROGRAMMING" if subject == "java" else "MATHEMATICAL FOUNDATION FOR COMPUTER APPLICATIONS"
                break  # Stop after replacing the title

        # Replace the keywords in the DOCX with questions
        replace_keywords_in_doc(doc, questions_by_module)

        # Replace CO text based on subject
        replace_co_text(doc, subject)
        replace_top_text(doc, subject)

        # Generate a timestamp for the output file name
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        docx_output_path = f'./generatedPapers/{subject}_{timestamp}.docx'

        # Save the updated DOCX
        doc.save(docx_output_path)

        # Convert DOCX to PDF
        pdf_output_path = f'./generatedPapers/{subject}_{timestamp}.pdf'
        convert(docx_output_path, pdf_output_path)  # Convert to PDF

        print(f"Generated question paper saved as: {docx_output_path}")
        print(f"Generated question paper PDF saved as: {pdf_output_path}")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
