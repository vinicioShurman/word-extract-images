import os
import sys
import docx2txt

# This code extracts all images of all files in the directory the script is running

def find_word_files(directory):
    word_files = []

    # List all files in the directory
    files = os.listdir(directory)

    # Filter files with ".docx" and ".doc" extensions
    word_files = [file for file in files if file.endswith((".docx", ".doc"))]

    return word_files

# Get the directory of the current script '__file__' is special variable in Python holds the path of the script that is currently being executed.
if getattr(sys, 'frozen', False):
    # Running as compiled executable
    directory_path = os.path.dirname(sys.executable)
else:
    # Running as a regular Python script
    directory_path = os.path.dirname(os.path.abspath(__file__))

word_files_found = find_word_files(directory_path)

for word in word_files_found:
    folder_path = os.path.join(directory_path, "Imagens_" + word)

    os.makedirs(folder_path, exist_ok=True)

    text = docx2txt.process(word, folder_path)