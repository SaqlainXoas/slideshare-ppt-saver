from __future__ import annotations

import warnings

import requests
import streamlit as st
from urllib3.exceptions import NotOpenSSLWarning

from slideshare_to_pptx import (
    SlideShareError,
    export_slideshare_pptx_to_tempfile,
    normalize_slideshare_url,
)

warnings.filterwarnings("ignore", category=NotOpenSSLWarning)

APP_TITLE = "SlideShare PPT Saver"


def inject_styles() -> None:
    st.markdown(
        """
        <style>
        .stApp {
            background:
                radial-gradient(circle at top left, rgba(13, 148, 136, 0.10), transparent 28%),
                radial-gradient(circle at bottom right, rgba(148, 163, 184, 0.12), transparent 32%),
                linear-gradient(180deg, #f5f7fb 0%, #edf1f6 100%);
        }
        .block-container {
            max-width: 680px;
            padding-top: 3.2rem;
            padding-bottom: 3rem;
        }
        [data-testid="stAppViewContainer"] .main .block-container {
            gap: 0 !important;
        }
        .app-shell {
            background: rgba(255, 255, 255, 0.88);
            border: 1px solid rgba(15, 23, 42, 0.08);
            border-radius: 28px;
            box-shadow: 0 24px 60px rgba(15, 23, 42, 0.08);
            padding: 1.35rem;
        }
        .hero-block,
        .result-card,
        .footnote {
            border-radius: 22px;
        }
        .hero-block {
            padding: 0.55rem 0.35rem 0.85rem;
        }
        .form-panel,
        .result-card {
            background: linear-gradient(180deg, rgba(255, 255, 255, 0.96) 0%, rgba(247, 250, 252, 0.94) 100%);
            border: 1px solid rgba(15, 23, 42, 0.08);
            border-radius: 22px;
        }
        .form-panel {
            padding: 0.95rem;
        }
        .eyebrow {
            display: inline-block;
            color: #0f766e;
            background: rgba(15, 118, 110, 0.08);
            border: 1px solid rgba(15, 118, 110, 0.10);
            padding: 0.45rem 0.75rem;
            border-radius: 999px;
            font-size: 0.76rem;
            font-weight: 700;
            letter-spacing: 0.08em;
            text-transform: uppercase;
        }
        .hero-title {
            color: #0f172a;
            font-size: clamp(2.15rem, 6vw, 3.4rem);
            line-height: 0.98;
            font-weight: 800;
            letter-spacing: -0.05em;
            margin: 0.95rem 0 0.35rem;
        }
        .hero-copy {
            color: #526073;
            font-size: 0.97rem;
            line-height: 1.6;
            margin: 0;
            max-width: 26rem;
        }
        .field-label {
            color: #0f172a;
            font-size: 0.9rem;
            font-weight: 700;
            margin: 0 0 0.5rem;
        }
        .result-card {
            padding: 0.95rem 1rem;
            margin-top: 1rem;
        }
        .result-title {
            color: #0f172a;
            font-size: 0.94rem;
            font-weight: 700;
            margin-bottom: 0.3rem;
        }
        .result-meta {
            color: #526073;
            font-size: 0.9rem;
            line-height: 1.45;
            margin: 0;
        }
        div[data-testid="stTextInputRootElement"] > div {
            border-radius: 16px !important;
        }
        div[data-testid="stTextInputRootElement"] input {
            font-size: 1rem !important;
            padding-top: 0.95rem !important;
            padding-bottom: 0.95rem !important;
        }
        div[data-testid="stTextInputRootElement"] label {
            display: none !important;
        }
        .stTextInput > label {
            display: none !important;
        }
        div[data-testid="stForm"] {
            border: none !important;
            padding: 0 !important;
        }
        div[data-testid="stButton"] > button,
        div[data-testid="stFormSubmitButton"] > button,
        button[kind="primary"] {
            border-radius: 16px !important;
            min-height: 3.25rem !important;
            font-weight: 700 !important;
            font-size: 1rem !important;
            border: none !important;
            background: linear-gradient(135deg, #0f766e 0%, #115e59 100%) !important;
            box-shadow: 0 14px 28px rgba(15, 118, 110, 0.22) !important;
        }
        div[data-testid="stButton"] > button:hover,
        div[data-testid="stFormSubmitButton"] > button:hover,
        button[kind="primary"]:hover {
            background: linear-gradient(135deg, #115e59 0%, #134e4a 100%) !important;
        }
        div[data-testid="stDownloadButton"] > button {
            border-radius: 16px !important;
            min-height: 3.15rem !important;
            font-weight: 700 !important;
            font-size: 1rem !important;
            border: 1px solid rgba(15, 23, 42, 0.08) !important;
            background: #0f172a !important;
            color: white !important;
            box-shadow: 0 14px 28px rgba(15, 23, 42, 0.16) !important;
        }
        div[data-testid="stDownloadButton"] > button:hover {
            background: #111827 !important;
        }
        @media (max-width: 720px) {
            .block-container {
                padding-top: 1.35rem;
                padding-bottom: 1.75rem;
            }
            .app-shell {
                border-radius: 24px;
                padding: 1rem;
            }
            .hero-block {
                padding: 0.2rem 0 0.75rem;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False, ttl=1800)
def build_download_payload(url: str) -> tuple[bytes, str, int, str]:
    result = export_slideshare_pptx_to_tempfile(normalize_slideshare_url(url))
    return (
        result.pptx_path.read_bytes(),
        result.pptx_path.name,
        result.slide_count,
        result.title,
    )


def render_header() -> None:
    st.markdown(
        """
        <div class="app-shell">
            <div class="hero-block">
            <div class="eyebrow">SlideShare</div>
            <h1 class="hero-title">Paste the link.<br/>Download.</h1>
            <p class="hero-copy">
                One link in. One file out.
            </p>
            </div>
        """,
        unsafe_allow_html=True,
    )


def close_shell() -> None:
    st.markdown("</div>", unsafe_allow_html=True)


def render_app() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="S", layout="centered")
    inject_styles()
    render_header()

    st.markdown(
        """
        <div class="form-panel">
            <p class="field-label">SlideShare link</p>
        """,
        unsafe_allow_html=True,
    )
    with st.form("download-form", clear_on_submit=False):
        url = st.text_input(
            "SlideShare link",
            placeholder="https://www.slideshare.net/slideshow/...",
            help="Public SlideShare slideshow links work best.",
        )
        submitted = st.form_submit_button("Generate", use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

    if not submitted:
        close_shell()
        return

    try:
        normalized_url = normalize_slideshare_url(url)
    except SlideShareError as exc:
        st.error(str(exc))
        close_shell()
        return

    with st.spinner("Downloading slides and building your PPTX..."):
        try:
            pptx_bytes, filename, slide_count, title = build_download_payload(normalized_url)
        except (SlideShareError, OSError, ValueError, requests.RequestException) as exc:
            st.error(f"Could not build the PPTX: {exc}")
            close_shell()
            return
        except Exception as exc:  # pragma: no cover
            st.error(f"Unexpected error: {exc}")
            close_shell()
            return

    st.markdown(
        f"""
        <div class="result-card">
            <div class="result-title">{title}</div>
            <p class="result-meta">{slide_count} slides</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.download_button(
        "Download",
        data=pptx_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
    )
    close_shell()
