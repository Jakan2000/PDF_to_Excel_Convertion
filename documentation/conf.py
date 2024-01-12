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
import os
import sys

# Add the path to your project
sys.path.insert(0, os.path.abspath('C:/Users/Admin/PycharmProjects/pythonProject1/KSV/FormatingExcelFiles/documentation'))
extensions = ['sphinx.ext.autodoc',]

templates_path = ['_templates']
exclude_patterns = ['_build', 'Thumbs.db', '.DS_Store']



# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output

# html_theme = 'alabaster'
html_theme = "furo"
# html_theme = 'groundwork'
# html_theme = 'cloud'

html_static_path = ['_static']
