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
html_title = " 성능기반 내진설계 업무절차서 " + release
# html_logo = "_static/images/myimage.png"
# html_static_path = [_static]

# -- LaTeX -------------------------------------------------
sd_fontawesome_latex = True
latex_engine = 'xelatex'
latex_documents=[('pbd_p3d_manual', 'main.tex', '성능기반 내진설계 업무절차서', 'CNP Dongyang', 'manual')]
# latex_logo='_static/images/CNP_logo.png'
latex_elements = {
    # The paper size ('letterpaper' or 'a4paper').
    'papersize': 'a4paper',

    # The font size ('10pt', '11pt' or '12pt').
    'pointsize': '10pt',

    # Additional stuff for the LaTeX preamble.
    'preamble': '',

    # Latex figure (float) alignment
    'figure_align': 'htbp',

    # kotex config
    'figure_align': 'htbp',

    'fontpkg': r'''
\usepackage{kotex}
\usepackage{setspace}
\singlespacing
\usepackage[skip=10pt plus1pt]{parskip}

\setmainfont{Noto Serif KR}
\setsansfont{Noto Sans KR}
\setmonofont{Noto Sans KR}
''',
}