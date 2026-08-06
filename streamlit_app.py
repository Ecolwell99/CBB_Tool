"""
Streamlit host for the CBB lineup tool.

This file is a HOST, not the app. All the lineup maths, the player join and the
UI live in index.html and run in the browser exactly as they do locally, so
clicking players in and out never touches the server -- important because the
people using this are spread across the country.

The three <script src> tags in index.html are inlined at runtime here rather
than committed as a pre-built bundle, so the repo keeps normal separate files
and a data.js refresh stays a clean diff.
"""

from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

HERE = Path(__file__).parent

# Inlined in this order -- index.html loads them as three sequential <script
# src> tags, and data.js/index_torvik.js are read by the inline script that
# follows them.
BUNDLED = ("sheetjs.js", "data.js", "index_torvik.js")

# components.html needs a fixed pixel height; the page inside is an app shell
# whose content region scrolls on its own.
FRAME_HEIGHT = 1000

st.set_page_config(page_title="CBB Lineup Tool", layout="wide")

# Reclaim the vertical space Streamlit reserves for its own chrome, so the
# tool's sidebar and table get close to the full window.
st.markdown(
    """
    <style>
      .block-container {padding: 0 !important; max-width: 100% !important;}
      header[data-testid="stHeader"] {display: none;}
      footer {display: none;}
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def build_page() -> str:
    """index.html with its external scripts inlined, ready for a srcdoc iframe."""
    html = (HERE / "index.html").read_text(encoding="utf-8")

    for name in BUNDLED:
        tag = f'<script src="{name}"></script>'
        if tag not in html:
            raise RuntimeError(
                f"expected {tag} in index.html but did not find it -- "
                "the script tags were renamed or reordered"
            )

        js = (HERE / name).read_text(encoding="utf-8")
        # A literal </script> in the payload would close the tag early. None of
        # the current files contain one; this keeps a future data.js honest.
        js = js.replace("</script", "<\\/script")
        html = html.replace(tag, f"<script>\n{js}\n</script>")

    return html


try:
    page = build_page()
except (OSError, RuntimeError) as exc:
    st.error(f"Could not assemble the tool: {exc}")
    st.stop()

components.html(page, height=FRAME_HEIGHT, scrolling=True)
