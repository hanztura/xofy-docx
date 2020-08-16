from docx import Document
from pathlib import Path

from settings import *


def create_document_directory(name):
    path = Path("{}/{}/".format(DOCUMENTS_DIRECTORY, name))
    if not path.exists():
        Path("{}/{}/outputs".format(DOCUMENTS_DIRECTORY, name)).mkdir(
            parents=True,
            exist_ok=True,
        )

        Path("{}/{}/variables.txt".format(
            DOCUMENTS_DIRECTORY, name)).write_text("")

        template = Path("{}/{}/template.docx".format(
            DOCUMENTS_DIRECTORY, name))

        document = Document()
        document.save(str(template))
    else:
        print("Folder '{}' already exist.")


if __name__ == '__main__':
    user_input = input("Document Name: ")
    create_document_directory(user_input.lower().replace(' ', '_'))
