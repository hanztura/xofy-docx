# create documents from a MS Word .docx template
import datetime
import os
import stringcase
import json

from docxtpl import DocxTemplate
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from settings import *
from models import Base, Document


engine = create_engine(DB_NAME, echo=True)


def create_db_engine():
    engine = create_engine(DB_NAME, echo=True)
    return engine


def create_db_tables(engine):
    Base.metadata.create_all(engine)


def save_to_db(output, context, created):
    d = Document(filename=output, data=json.dumps(context), created=created)

    Session = sessionmaker(bind=engine)
    session = Session()
    session.add(d)
    session.commit()

    return d


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


def get_output_filename(timestamp):
    return "output--{}.docx".format(
        timestamp.strftime(TIMESTAMP_FORMAT)
    )


def main():
    # get the document user wants to workon
    documents = get_available_documents(DOCUMENTS_DIRECTORY)
    print_available_documents(documents)
    workon = int(get_user_input(
        "Type number of document to workon: ",
        True
    ))
    timestamp = datetime.datetime.now()

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
        get_output_filename(timestamp),
    )

    # get the template document
    doc = DocxTemplate(template)

    # render
    context = get_context(get_variables(variables))
    doc.render(context)
    doc.save(output)

    create_db_tables(engine)
    context['timestamp'] = str(timestamp)
    save_to_db(output, context, timestamp)


if __name__ == '__main__':
    main()
