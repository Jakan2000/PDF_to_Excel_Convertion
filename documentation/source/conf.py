# Configuration file for the Sphinx documentation builder.
#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information

project = 'pdf_to_excel_convertion'
copyright = '2024, jakan'
author = 'jakan'
release = '1.0.0'

# -- General configuration ---------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#general-configuration

# extensions = []
# templates_path = []
# exclude_patterns = []
#=======================================================================================================================
# from chatgpt
import os
import sys

sys.path.insert(0, os.path.abspath('..'))
scripts = "C:/Users/Admin/PycharmProjects/pythonProject1/KSV/FormatingExcelFiles/documentation"

extensions = [
    'sphinx.ext.autodoc',
    'sphinx.ext.viewcode',
    'sphinx.ext.todo',
    # ... other extensions
]
templates_path = ['_templates']


exclude_patterns = ['_build', 'Thumbs.db', '.DS_Store']
#=======================================================================================================================


# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output

html_theme = 'alabaster'
html_static_path = ['_static']
