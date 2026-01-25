# Configuration file for the Sphinx documentation builder.
#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information

project = 'Answer-Sheet-Studio'
copyright = '2026, Abieskawa and Calamus'
author = 'Abieskawa and Calamus'
release = '0.0'

# -- General configuration ---------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#general-configuration

import os
from datetime import datetime, timezone

extensions = []

templates_path = ['_templates']
exclude_patterns = []

_rtd_language = os.environ.get('READTHEDOCS_LANGUAGE', '').strip().lower()
if _rtd_language in {'zh-tw', 'zh_tw'}:
    language = 'zh_TW'
elif _rtd_language:
    language = _rtd_language.replace('-', '_')
else:
    # Default to zh_TW so local builds match the RTD default language.
    language = 'zh_TW'

# Date helpers (show both zh/en at the top of index.rst).
_now = datetime.now(timezone.utc).astimezone()
today_zh = _now.strftime("%Y 年 %m 月 %d 日")
today_en = _now.strftime("%Y-%m-%d")

rst_epilog = f"""
.. |today_zh| replace:: {today_zh}
.. |today_en| replace:: {today_en}
"""

# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output

html_theme = 'sphinx_rtd_theme'
html_static_path = ['_static']

# -- Options for LaTeX / PDF output ------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-latex-output

latex_engine = 'xelatex'

latex_elements = {
    'preamble': r"""
\usepackage{fontspec}
\usepackage{xeCJK}

\newcommand{\sphinxsetcjkfonts}[1]{
  \setCJKmainfont{#1}
  \setCJKsansfont{#1}
  \setCJKmonofont{#1}
}

% Prefer Noto CJK on Read the Docs; fall back to common macOS fonts.
\IfFontExistsTF{Noto Sans CJK TC}{
  \sphinxsetcjkfonts{Noto Sans CJK TC}
}{
  \IfFontExistsTF{PingFang TC}{
    \sphinxsetcjkfonts{PingFang TC}
  }{
    \IfFontExistsTF{Noto Sans CJK SC}{
      \sphinxsetcjkfonts{Noto Sans CJK SC}
    }{
      \setCJKmainfont{FandolSong-Regular}
      \setCJKsansfont{FandolHei-Regular}
      \setCJKmonofont{FandolFang-Regular}
    }
  }
}

% Sphinx may write \selectlanguage*{...} into .aux/.toc; babel doesn't support
% the star-form, so make it accept and ignore the star.
\makeatletter
\@ifpackageloaded{babel}{
  \let\sphinx@selectlanguage\selectlanguage
  \renewcommand{\selectlanguage}{\@ifstar{\sphinx@selectlanguage}{\sphinx@selectlanguage}}
}{}
\makeatother
""",
}
