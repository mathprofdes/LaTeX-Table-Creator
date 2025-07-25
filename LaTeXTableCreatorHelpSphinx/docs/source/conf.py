# Configuration file for the Sphinx documentation builder.
#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information

from __future__ import annotations

import importlib.util
import os
import sys
import time
import datetime

from sphinx.application import Sphinx

project = 'LaTeX Table Creator'
copyright = '2025, Don Spickler'
author = 'Don Spickler'
release = '2.6.1'

# -- General configuration ---------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#general-configuration

extensions = []

templates_path = ['_templates']
exclude_patterns = []



# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output

extensions = [
    "sphinx.ext.autodoc",
    "sphinx.ext.viewcode",
    "sphinx.ext.napoleon",  # has to be loaded before sphinx_autodoc_typehints
    "sphinx_autodoc_typehints",
    "sphinx.ext.intersphinx",
    #"sphinx_qt_documentation",
    "sphinx_design",
    "sphinx_favicon",
    #"sphinxext.rediraffe",
    #"sphinxcontrib.images",
    #'sphinx.ext.inheritance_diagram',
    # 'sphinxcontrib.spelling'  # commenting out to allow for easy usage locally
    "sphinx.ext.imgmath"
]

templates_path = ['_templates']
html_static_path = ['_static']
html_css_files = ["custom.css"]
html_last_updated_fmt = '%A, %B %d, %Y  -  %I:%M %p'
exclude_patterns = []
master_doc = 'index'

html_show_sourcelink = False
html_show_sphinx = True

# -- Options for HTML output -------------------------------------------------

html_logo = os.path.join("", "ProgramIcon2.png")
html_theme = 'pydata_sphinx_theme'

html_theme_options = {
    "collapse_navigation": True,
    "icon_links": [
        {
            "name": "GitHub",
            "url": "https://github.com/mathprofdes/Latex-Table-Creator",
            "icon": "fa-brands fa-square-github",
            "type": "fontawesome",
        }
    ],
#    "navbar_end": ["theme-switcher", "navbar-icon-links"],
    "navbar_end": ["navbar-icon-links"],
    "footer_start": ["copyright", "last-updated"],
    "footer_end": ["sphinx-version", "theme-version"],
    "navigation_depth": 2,
    "navigation_with_keys": False,
    "secondary_sidebar_items": ["page-toc"],
    "show_toc_level": 3,
    "show_nav_level": 2,
    "use_edit_page_button": False,

    "navbar_align": "left",
    "primary_sidebar_end": [],
    "back_to_top_button": False,
}

html_sidebars = {
  "**": []
}
imgmath_use_preview=True

