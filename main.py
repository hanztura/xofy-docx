# create documents from a MS Word .docx template
import datetime
import os
import stringcase

from docxtpl import DocxTemplate

from settings import *


def get_user_input(msg, is_required=False):
    user_input = input(msg)

    # ask input again if required and not valid input
    while is_required and len(user_input) == 0:
        print("***Need a valid input***")
        user_input = input(msg)

    return user_input


def get_variables(variables_filename):
    with open(variables_filename) as f:
        variables = [line.strip().split() for line in f.readlines()]

    return variables


def get_variable_attributes(variable):
    """Returns a dict {
        "name": "variable_name",
        "is_required": True
    }

    variable argument is in [name, required]
    """
    return {
        "name": variable[0],
        "is_required": '--required' in variable,
    }


def get_label(name, is_required):
    label = '{} (required): ' if is_required else '{} (optional): '
    return label.format(stringcase.titlecase(name))


def get_context(variables):
    context = {}
    for variable in variables:
        variable = get_variable_attributes(variable)
        name = variable["name"]
        is_required = variable['is_required']

        label = get_label(name, is_required)

        user_input = get_user_input(label, is_required)

        context[name] = user_input

    return context


def get_available_documents(documents_name='documents'):
    """Return a list of the sub directories of the documents directory"""
    documents = [name for name in os.listdir(documents_name) if os.path.isdir(
        os.path.join(documents_name, name))]
    return documents


def print_available_documents(documents):
    for index, doc in enumerate(documents):
        print(index, stringcase.titlecase(doc))


def get_output_filename():
    return "output--{}.docx".format(
        datetime.datetime.now().strftime(TIMESTAMP_FORMAT)
    )


def main():
    # get the document user wants to workon
    documents = get_available_documents(DOCUMENTS_DIRECTORY)
    print_available_documents(documents)
    workon = int(get_user_input(
        "Type number of document to workon: ",
        True
    ))

    template = os.path.join(
        DOCUMENTS_DIRECTORY,
        documents[workon],
        TEMPLATES_FILENAME,
    )
    variables = os.path.join(
        DOCUMENTS_DIRECTORY,
        documents[workon],
        VARIABLES_FILENAME,
    )
    output = os.path.join(
        DOCUMENTS_DIRECTORY,
        documents[workon],
        "outputs",
        get_output_filename(),
    )

    # get the template document
    doc = DocxTemplate(template)

    # render
    doc.render(get_context(get_variables(variables)))
    doc.save(output)


if __name__ == '__main__':
    main()
