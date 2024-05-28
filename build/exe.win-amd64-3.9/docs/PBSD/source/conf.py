# Configuration file for the Sphinx documentation builder.
#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information

from docutils.writers.latex2e import Babel
Babel.language_codes = {'ko':'korean', 'en':'english'}
project = '성능기반 내진설계 업무절차서'
copyright = '2023, CNP Dongyang'
author = 'CNP Dongyang'
release = 'v1.0'

# -- General configuration ---------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#general-configuration

extensions = ["sphinx_design"]

templates_path = ['_templates']
exclude_patterns = []

language = 'ko'

# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output

html_theme = 'pydata_sphinx_theme'
html_title = " 성능기반 내진설계 업무절차서 "
# html_logo = "_static/images/myimage.png"
html_static_path = ['_static']
html_css_files = ['css/custom.css']

# inclusion for the color support
# https://stackoverflow.com/questions/3702865/sphinx-restructuredtext-set-color-for-a-single-word
rst_prolog = """
.. include:: <s5defs.txt>

 """

# -- LaTeX -------------------------------------------------
release = ''

latex_use_xindy = True
sd_fontawesome_latex = True
latex_engine = 'xelatex'
latex_documents=[('pbd_p3d_manual_latex', 'manual.tex', '성능기반 내진설계 업무절차서', 'CNP Dongyang', 'manual')]
# latex_logo='_static/images/CNP_logo.png'
latex_elements = {
    # The paper size ('letterpaper' or 'a4paper').
    'papersize': 'a4paper',

    # The font size ('10pt', '11pt' or '12pt').
    'pointsize': '11pt',

    # Additional stuff for the LaTeX preamble.
    'preamble': '',

    # Latex figure (float) alignment
    'figure_align': 'htbp',

    # Remove blank pages
    'extraclassoptions': 'openany,oneside',

    # Delete Release
    'releasename': '',

    # kotex config
    'figure_align': 'htbp',

    'fontpkg': r'''
\usepackage{kotex}
\usepackage{setspace}
\onehalfspacing
\usepackage[skip=10pt plus1pt]{parskip}
\usepackage[bottom]{footmisc}

\setmainfont{Noto Serif KR}
\setsansfont{Noto Sans KR}
\setmonofont{Noto Sans KR}
''',
}