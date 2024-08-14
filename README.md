**Case-Note Document Generator**
🖥️ Tech Stack 🖥️

Built with Python, Openpyxl, Docx

💡 **Inspiration** 💡

In educational settings, maintaining comprehensive and well-organized documentation for student cases is crucial. This process, however, can be time-consuming and prone to human error. Our solution automates the creation of reflection documents, transforming raw data into well-structured Word documents. By leveraging the power of Python libraries and language models, we aim to simplify documentation tasks, enhance accuracy, and free up time for more critical activities.

**What is Case-Note Document Generator?**
The Case-Note Document Generator is a Python-based tool designed to automate the creation of student reflection documents. It extracts data from an Excel file and generates a formatted Word document, including sections for student details, issues and resolutions, and records of support. The tool uses advanced language models to summarize and convert text, ensuring clear and concise output.

**Key Features:**
Excel Integration: Monitors an Excel file for updates and automatically generates a Word document when new data is detected.
Dynamic Content: Uses language models to summarize input text and generate past tense descriptions.
Custom Formatting: Creates tables with borders and formatted text for easy readability and professional presentation.

🔧 **How I Built It** 🔧

The script is designed to take rows from an Excel file, and create a Word document with the following components:

Frontend: The user interacts with an Excel file, which serves as the data source.
Backend: Python script processes the data using libraries such as Openpyxl for reading Excel files and Docx for creating Word documents.
Machine Learning: Utilizes Langchain's language models to generate summaries and convert text to past tense, enhancing document content.
Script Components:

cutPretext(result): Cleans up and extracts relevant text from language model output.
summarize(sentence_input): Converts input sentences from first person to third person past tense.
desiredOutcome(sentence_input): Provides a summary of the expected desired outcome based on input text.
create_borders(table): Adds borders to tables in the Word document for improved readability.
create_word_document(data, filename): Generates a Word document with formatted sections based on the extracted data.
monitor_excel_file(excel_filename, word_filename): Monitors the specified Excel file for updates and triggers document creation.

👀 What's Next for Case-Note Document Generator? 👀

Enhanced Functionality: Explore additional features such as reading from emails to extend on the existing case notes.
User Interface: Consider developing a user-friendly interface for easier interaction and configuration.
Performance Optimization: Refine the script for faster processing and better handling.
