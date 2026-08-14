"""Diseño responsivo global del Sistema CAV."""

import streamlit as st


def aplicar_diseno_responsivo_inteligente():
    css_responsivo = """
    <style>
    :root {
        --cav-touch-size: 46px;
        --cav-radius-responsive: clamp(10px, 1.2vw, 18px);
        --cav-gap-responsive: clamp(.55rem, 1.1vw, 1.15rem);
        --cav-page-padding-x: clamp(.65rem, 2.4vw, 2.4rem);
        --cav-page-padding-y: clamp(.75rem, 1.8vw, 1.8rem);
        --cav-content-width: 1760px;
    }

    html {
        text-size-adjust: 100%;
        -webkit-text-size-adjust: 100%;
        scroll-behavior: smooth;
    }

    body,
    [data-testid="stAppViewContainer"] {
        overflow-x: hidden;
    }

    .block-container {
        width: 100% !important;
        max-width: var(--cav-content-width) !important;
        padding:
            var(--cav-page-padding-y)
            var(--cav-page-padding-x)
            calc(var(--cav-page-padding-y) + env(safe-area-inset-bottom))
            var(--cav-page-padding-x)
            !important;
    }

    h1 {
        font-size: clamp(1.75rem, 3.3vw, 3rem) !important;
        line-height: 1.08 !important;
        text-wrap: balance;
    }

    h2 {
        font-size: clamp(1.4rem, 2.6vw, 2.25rem) !important;
        line-height: 1.12 !important;
        text-wrap: balance;
    }

    h3 {
        font-size: clamp(1.12rem, 2vw, 1.65rem) !important;
        line-height: 1.16 !important;
    }

    p,
    label,
    [data-testid="stMarkdownContainer"] {
        overflow-wrap: anywhere;
    }

    img,
    video,
    canvas,
    svg {
        max-width: 100% !important;
        height: auto;
    }

    iframe {
        max-width: 100% !important;
    }

    /* streamlit-autorefresh usa un iframe invisible.
       No se debe convertir en un bloque alto con height:auto. */
    iframe[title*="autorefresh"],
    iframe[src*="streamlit_autorefresh"] {
        position: absolute !important;
        width: 0 !important;
        min-width: 0 !important;
        height: 0 !important;
        min-height: 0 !important;
        border: 0 !important;
        opacity: 0 !important;
        pointer-events: none !important;
        overflow: hidden !important;
    }

    div[data-testid="stElementContainer"]:has(
        iframe[title*="autorefresh"]
    ),
    div[data-testid="stElementContainer"]:has(
        iframe[src*="streamlit_autorefresh"]
    ) {
        display: none !important;
        width: 0 !important;
        height: 0 !important;
        min-height: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
        overflow: hidden !important;
    }

    [data-testid="stImage"] img {
        object-fit: contain;
    }

    [data-testid="stHorizontalBlock"] {
        gap: var(--cav-gap-responsive) !important;
        align-items: stretch;
    }

    [data-testid="column"] {
        min-width: 0 !important;
    }

    [data-testid="stForm"],
    [data-testid="stExpander"],
    [data-testid="stMetric"],
    [data-testid="stAlert"],
    [data-testid="stDataFrame"],
    [data-testid="stFileUploader"] {
        border-radius: var(--cav-radius-responsive) !important;
    }

    [data-testid="stForm"] {
        padding: clamp(.85rem, 1.8vw, 1.5rem) !important;
    }

    [data-testid="stMetric"] {
        height: 100%;
        padding: clamp(.7rem, 1.25vw, 1.05rem) !important;
    }

    [data-testid="stMetricValue"] {
        font-size: clamp(1.3rem, 2.2vw, 2rem) !important;
        overflow-wrap: anywhere;
    }

    [data-testid="stMetricLabel"] {
        white-space: normal !important;
    }

    .stButton > button,
    .stDownloadButton > button,
    [data-testid="stLinkButton"] a,
    button[kind] {
        min-height: var(--cav-touch-size) !important;
        border-radius: clamp(9px, 1vw, 14px) !important;
        white-space: normal !important;
        line-height: 1.2 !important;
    }

    input,
    textarea,
    [data-baseweb="select"] > div {
        min-height: var(--cav-touch-size) !important;
        font-size: max(16px, 1rem) !important;
        border-radius: clamp(8px, .9vw, 13px) !important;
    }

    textarea {
        min-height: 110px !important;
    }

    [data-testid="stTabs"] [role="tablist"] {
        overflow-x: auto;
        overflow-y: hidden;
        scrollbar-width: thin;
        flex-wrap: nowrap;
    }

    [data-testid="stTabs"] [role="tab"] {
        flex: 0 0 auto;
        min-height: var(--cav-touch-size);
        white-space: nowrap;
    }

    [data-testid="stDataFrame"],
    [data-testid="stTable"] {
        width: 100% !important;
        overflow-x: auto !important;
    }

    [data-testid="stDataFrame"] > div {
        max-width: 100% !important;
    }

    [data-testid="stSidebar"] {
        width: clamp(270px, 25vw, 370px) !important;
    }

    [data-testid="stSidebarContent"] {
        padding-bottom: env(safe-area-inset-bottom);
    }

    [data-testid="stDialog"] > div {
        width: min(94vw, 760px) !important;
        max-height: 92dvh !important;
        overflow-y: auto !important;
    }

    /* Tablets y notebooks pequeños */
    @media (max-width: 1180px) {
        :root {
            --cav-content-width: 100%;
            --cav-gap-responsive: .75rem;
        }

        [data-testid="stHorizontalBlock"] {
            flex-wrap: wrap !important;
        }

        [data-testid="column"] {
            flex: 1 1 min(320px, 100%) !important;
            width: auto !important;
        }
    }

    /* Celulares y tablets verticales */
    @media (max-width: 760px) {
        :root {
            --cav-touch-size: 48px;
            --cav-page-padding-x: .7rem;
            --cav-page-padding-y: .8rem;
        }

        .block-container {
            max-width: 100% !important;
        }

        h1,
        h2 {
            text-align: center;
        }

        [data-testid="stHorizontalBlock"] {
            display: flex !important;
            flex-direction: column !important;
            gap: .7rem !important;
        }

        [data-testid="column"] {
            flex: 1 1 100% !important;
            width: 100% !important;
            min-width: 100% !important;
        }

        .stButton,
        .stDownloadButton,
        [data-testid="stLinkButton"] {
            width: 100%;
        }

        .stButton > button,
        .stDownloadButton > button,
        [data-testid="stLinkButton"] a {
            width: 100% !important;
            justify-content: center !important;
            padding-left: .8rem !important;
            padding-right: .8rem !important;
        }

        [data-testid="stSidebar"] {
            width: min(88vw, 340px) !important;
        }

        [data-testid="stMetric"] {
            min-height: 88px;
        }

        [data-testid="stForm"] {
            padding: .85rem !important;
        }

        [data-testid="stFileUploader"] section {
            padding: .8rem !important;
        }

        [data-testid="stFileUploaderDropzoneInstructions"] {
            min-width: 0 !important;
        }

        [data-testid="stToolbar"],
        [data-testid="stDecoration"] {
            max-width: 100vw !important;
        }

        div[data-testid="stPlotlyChart"] {
            overflow-x: auto;
        }

        div[data-testid="stPlotlyChart"] > div {
            min-width: 100%;
        }
    }

    /* Celulares muy pequeños */
    @media (max-width: 390px) {
        :root {
            --cav-page-padding-x: .5rem;
        }

        h1 {
            font-size: 1.62rem !important;
        }

        h2 {
            font-size: 1.32rem !important;
        }

        [data-testid="stMetricValue"] {
            font-size: 1.22rem !important;
        }
    }

    /* Pantallas grandes y televisores */
    @media (min-width: 1700px) {
        :root {
            --cav-content-width: 1900px;
            --cav-page-padding-x: 2.8rem;
        }

        body {
            font-size: 1.05rem;
        }

        .stButton > button,
        .stDownloadButton > button,
        [data-testid="stLinkButton"] a {
            min-height: 50px !important;
        }
    }

    /* Dispositivos táctiles */
    @media (pointer: coarse) {
        button,
        a,
        input,
        textarea,
        [role="tab"],
        [role="option"] {
            touch-action: manipulation;
        }

        [data-baseweb="select"] [role="option"] {
            min-height: 46px;
        }
    }

    /* Pantallas horizontales de poca altura */
    @media (orientation: landscape) and (max-height: 650px) {
        :root {
            --cav-page-padding-y: .45rem;
        }

        h1 {
            margin-top: .2rem !important;
            margin-bottom: .45rem !important;
        }

        [data-testid="stForm"] {
            padding-top: .65rem !important;
            padding-bottom: .65rem !important;
        }
    }

    /* Accesibilidad: reducir animaciones cuando el dispositivo lo solicita */
    @media (prefers-reduced-motion: reduce) {
        *,
        *::before,
        *::after {
            scroll-behavior: auto !important;
            animation-duration: .01ms !important;
            animation-iteration-count: 1 !important;
            transition-duration: .01ms !important;
        }
    }
    </style>
    """

    if hasattr(st, "html"):
        st.html(css_responsivo)
    else:
        st.markdown(css_responsivo, unsafe_allow_html=True)

