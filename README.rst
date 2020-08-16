***********************************************
Create multiple .docx document using a template
***********************************************

How to install
##############


* (Optional) Create and activate virtual environment
    .. code-block:: bash

        python3 -m virtualenv env
        source env/bin/activate

* Install Python required packages.
    .. code-block:: bash

        pip install -r requirements.txt


How to use
##########

* Create a document. Only required on first time use.
    .. code-block:: bash

        python new_document.py

* Edit the template.docx file inside the document folder. Use Jinja2 template language. See `sample-template.docx` for example.

* List your variable names in kebab case. See `sample-variables.txt` for examples.
    - Required fields should have '--required' appended.

* Run the program.
    .. code-block:: bash

        python main.py


Notes to developer
#################

* optional or required variables [Done]
    - in requirements.txt file append ` --required` if variable is required
* input label should be user friendly text [Done]
* multiple templates [Done]
* filename should have timestamp [Done]
* user can create another document with bootstrapped files/folders [DONE]
    - run `python new_document.py`