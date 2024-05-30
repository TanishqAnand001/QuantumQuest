import random
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches

def select_questions(questions_list, num_questions):
    """
    Select the specified number of questions from the given list, allowing repetitions if necessary.

    :param questions_list: List of questions to select from.
    :param num_questions: Number of questions to select.
    :return: List of selected questions.
    """
    if num_questions <= len(questions_list):
        return random.sample(questions_list, num_questions)
    else:
        return questions_list + random.choices(questions_list, k=num_questions - len(questions_list))

def add_image_to_document(doc, image_path):
    """
    Add an image to the given document.

    :param doc: Document object.
    :param image_path: Path to the image file.
    """
    doc.add_picture(image_path, width=Inches(4.0))  # Adjust width as needed

def read_question_bank(filename):
    """
    Read the question bank from a file and return a dictionary of topics and their associated questions.

    :param filename: Name of the question bank file.
    :return: Dictionary of topics and their associated questions.
    """
    topics = {}
    with open(filename, 'r', encoding='utf-8') as file:
        for line in file:
            parts = line.strip().split('|')
            topic = parts[0]
            question_text = parts[1].replace("\\u03A9", "Ω")  # Replace escape sequence with symbol
            marks = int(parts[2])
            image_path = parts[3] if len(parts) > 3 else None
            options = parts[4:] if len(parts) > 4 else []

            if topic not in topics:
                topics[topic] = {}
            if marks not in topics[topic]:
                topics[topic][marks] = []
            topics[topic][marks].append((question_text, image_path, options, marks))
    return topics

def set_document_style(doc):
    """
    Set the document style to Verdana, font size 11, and margins.
    """
    # Set the font to Verdana and size to 11 for the entire document
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Verdana'
    font.size = Pt(11)
    
    # Set document margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

def create_question_paper(topics, num_questions_per_type, specific_topic_questions, output_filename):
    """
    Create a Word document with questions from specified topics.

    :param topics: Dictionary of topics and their associated questions.
    :param num_questions_per_type: Dictionary specifying how many questions of each type to select.
    :param specific_topic_questions: Dictionary specifying how many questions of each topic to select.
    :param output_filename: Name of the output Word document file.
    """
    # Create a new Document
    doc = Document()
    set_document_style(doc)  # Apply the styles to the document

    doc.add_heading('Question Paper', 0).alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    selected_questions = []

    # Select questions for the specific topic
    for topic, num_questions_by_marks in specific_topic_questions.items():
        for marks, num_questions in num_questions_by_marks.items():
            if marks in num_questions_per_type:
                topic_questions = topics.get(topic, {}).get(marks, [])
                selected_questions.extend(select_questions(topic_questions, num_questions))

    # Calculate remaining questions after topic selection
    remaining_questions_per_type = {mark: count - sum(q[3] == mark for q in selected_questions)
                                    for mark, count in num_questions_per_type.items()}

    # Check if any remaining questions need selection
    if any(count > 0 for count in remaining_questions_per_type.values()):
        # Select remaining questions randomly from all topics (excluding a specific topic, if desired)
        for marks, num_questions in remaining_questions_per_type.items():
            if num_questions > 0:
                all_topic_questions = []
                for topic, topic_questions in topics.items():
                    if topic != "Current Electricity":  # Exclude a specific topic (if needed)
                        all_topic_questions.extend(topic_questions.get(marks, []))
                selected_questions.extend(select_questions(all_topic_questions, num_questions))

    # Add questions to the document
    questions_by_marks = {}
    for question in selected_questions:
        marks = question[3]
        if marks not in questions_by_marks:
            questions_by_marks[marks] = []
        questions_by_marks[marks].append(question)

    for marks, questions in questions_by_marks.items():
        doc.add_heading(f"{marks}-Mark Questions", level=1)
        for question_text, image_path, options, _ in questions:
            doc.add_paragraph(f"Question: {question_text}")
            if marks == 1 and options:  # Only 1-mark questions are MCQs
                for i, option in enumerate(options):
                    doc.add_paragraph(f"{chr(65 + i)}. {option}", style='List Bullet')
            if image_path:
                add_image_to_document(doc, image_path)

    # Save the document
    doc.save(output_filename)


def prompt_user_for_input(topics):
    """
    Prompt the user for input about the question paper configuration.

    This function gathers information about the desired number of questions
    for each mark value and for each topic. It also prompts the user for
    the output filename.

    Args:
        topics: A dictionary containing the available topics and their questions.

    Returns:
        A tuple containing three elements:
            - num_questions_per_type: A dictionary mapping mark values (e.g., 1, 2)
                                      to the desired number of questions for that mark.
            - num_questions_per_topic: A dictionary mapping topics to another
                                      dictionary that maps mark values to the desired
                                      number of questions from that topic for that mark.
            - output_filename: The filename specified by the user for the output Word document.
    """
    num_questions_per_type = {}
    remaining_questions_per_type = {}  # Track remaining questions per mark

    # Prompt for total count of questions for each mark type
    for mark_type in [1, 2, 3, 5]:
        while True:
            try:
                num_questions = int(input(f"Enter the total number of {mark_type}-mark questions: "))
                if num_questions < 0:
                    print("Error: Please enter a non-negative number of questions.")
                else:
                    num_questions_per_type[mark_type] = num_questions
                    remaining_questions_per_type[mark_type] = num_questions  # Initialize remaining count
                    break
            except ValueError:
                print("Error: Please enter a valid integer number of questions.")

    num_questions_per_topic = {}

    # Prompt for count of questions for each topic
    for topic in topics:
        print(f"\nFor topic '{topic}':")
        topic_questions = {}

        # Subtract the count of questions for each topic from the remaining count for each mark type
        for mark_type in num_questions_per_type:
            max_questions = remaining_questions_per_type[mark_type]
            while True:
                try:
                    num_questions = int(input(f"Enter the number of {mark_type}-mark questions you want from topic '{topic}' (maximum {max_questions}): "))
                    if num_questions < 0:
                        print("Error: Please enter a non-negative number of questions.")
                    elif num_questions > max_questions:
                        print(f"Error: The number of {mark_type}-mark questions from '{topic}' exceeds the maximum allowed ({max_questions}).")
                    else:
                        topic_questions[mark_type] = num_questions
                        remaining_questions_per_type[mark_type] -= num_questions  # Subtract the selected count
                        break
                except ValueError:
                    print("Error: Please enter a valid integer number of questions.")

        num_questions_per_topic[topic] = topic_questions

    output_filename = input("Enter the output filename (e.g., Question_Paper.docx): ").strip()
    return num_questions_per_type, num_questions_per_topic, output_filename



# Main execution
if __name__ == "__main__":
    # Read questions from the question bank file
    question_bank_filename = "question.txt"
    topics = read_question_bank(question_bank_filename)

    # Prompt the user for input
    num_questions_per_type, specific_topic_questions, output_filename = prompt_user_for_input(topics)

    # Create the question paper
    create_question_paper(topics, num_questions_per_type, specific_topic_questions, output_filename)
	