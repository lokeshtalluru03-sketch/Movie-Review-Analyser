"""
The Odyssey — TMDB Review Analyzer (Demo Mode)
-----------------------------------------------
Looks and feels like a live TMDB fetch + AI categorization pipeline
(progress bars, staged status text), but the actual category /
sentiment / reasoning values are read from a pre-analyzed master
Excel file (odyssey_reviews_analysis.xlsx) sitting next to this
script. This avoids depending on a live Groq/TMDB call at click time.

If you later want this to hit real APIs again, swap `simulate_fetch`
and `simulate_analysis` for real TMDB/Groq calls and use their
results instead of `load_master_data()`.

Install:
    pip install streamlit pandas openpyxl

Run:
    streamlit run review_monitor.py

Make sure odyssey_reviews_analysis.xlsx is in the same folder.
"""

import os
import time
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# =============================================================================
# CONFIG
# =============================================================================

MOVIE_TITLE = "The Odyssey"
MOVIE_ID = 1368337

EXCEL_FILE = "odyssey_reviews_analysis.xlsx"  # pre-analyzed master dataset

CATEGORIES = [
    "Overall Experience",
    "Casting & Diversity Controversy",
    "Historical/Cultural Accuracy",
    "Acting Performances",
    "Story/Adaptation Fidelity",
    "Writing & Dialogue",
]

SENTIMENTS = ["Positive", "Negative", "Mixed"]

BASE_COLUMNS = [
    "review_id", "movie", "movie_id", "reviewer", "review_text",
    "rating", "created_at", "updated_at", "review_url", "source",
]

FULL_COLUMNS = BASE_COLUMNS + [
    "category", "sentiment", "reasoning", "status", "analyzed_at",
]


# =============================================================================
# DATA LOADING — reads the pre-built master Excel file
# =============================================================================

@st.cache_data(show_spinner=False)
def load_master_data(path: str = EXCEL_FILE) -> pd.DataFrame:
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Could not find '{path}'. Place it in the same folder as this script."
        )

    df = pd.read_excel(path)

    for col in FULL_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    return df[FULL_COLUMNS]


# =============================================================================
# SIMULATED "LIVE" FETCH / ANALYZE STAGES
# (real API calls could be dropped in here later)
# =============================================================================

def simulate_fetch(df: pd.DataFrame):
    n = len(df)

    progress = st.progress(0, text="Connecting to TMDB...")

    intro_steps = [
        "Searching for 'The Odyssey' (2026)...",
        f"Movie found — TMDB ID {MOVIE_ID}...",
        "Requesting review pages...",
    ]

    for i, step_text in enumerate(intro_steps):
        progress.progress(int((i + 1) / (len(intro_steps) + 3) * 100), text=step_text)
        time.sleep(0.45)

    batch_size = max(1, n // 5)
    base_pct = len(intro_steps)
    total_stages = len(intro_steps) + 3

    for start in range(0, n, batch_size):
        fetched = min(start + batch_size, n)
        pct = int(((base_pct + (fetched / n) * 3) / total_stages) * 100)
        progress.progress(min(pct, 99), text=f"Fetched {fetched}/{n} reviews...")
        time.sleep(0.35)

    progress.progress(100, text=f"Done — {n} reviews fetched.")
    time.sleep(0.3)
    progress.empty()


def simulate_analysis(df: pd.DataFrame):
    n = len(df)
    batch_size = 5

    progress = st.progress(0, text="Warming up AI model...")
    time.sleep(0.4)

    for start in range(0, n, batch_size):
        completed = min(start + batch_size, n)
        progress.progress(
            int(completed / n * 100),
            text=f"Categorizing + scoring sentiment... {completed}/{n} reviews",
        )
        time.sleep(0.3)

    progress.progress(100, text="Analysis complete.")
    time.sleep(0.3)
    progress.empty()


# =============================================================================
# EXCEL EXPORT — rebuild a clean, formatted copy for download
# =============================================================================

def save_excel(df: pd.DataFrame, path: str = EXCEL_FILE):
    df = df[FULL_COLUMNS]
    df.to_excel(path, index=False, engine="openpyxl")

    wb = load_workbook(path)
    ws = wb.active
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    widths = {
        "A": 25, "B": 18, "C": 12, "D": 22, "E": 70, "F": 10,
        "G": 24, "H": 24, "I": 55, "J": 12, "K": 28, "L": 12,
        "M": 60, "N": 12, "O": 24,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    wb.save(path)


# =============================================================================
# STREAMLIT APP
# =============================================================================

def init_state():
    defaults = {
        "fetched": False,
        "analyzed": False,
        "base_df": pd.DataFrame(),
        "results": pd.DataFrame(),
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def render_header():
    st.title("🎬 The Odyssey — Review Analyzer")
    st.caption("TMDB reviews → AI category + sentiment → Excel")
    st.markdown(
        f"""
        **Movie:** {MOVIE_TITLE}  
        **Source:** TMDB  
        **Outputs:** Category + Sentiment + Reasoning
        """
    )


def render_fetch_section(master_df: pd.DataFrame):
    st.subheader("1️⃣ Fetch The Odyssey Reviews")

    if st.button(
        "📥 Fetch Odyssey reviews from TMDB",
        type="primary",
        use_container_width=True,
    ):
        with st.spinner("Talking to TMDB..."):
            simulate_fetch(master_df)

        base_df = master_df[BASE_COLUMNS].copy()

        st.session_state.base_df = base_df
        st.session_state.fetched = True
        st.session_state.analyzed = False
        st.session_state.results = pd.DataFrame()

        st.success(
            f"Fetched {len(base_df)} review(s) for {MOVIE_TITLE} "
            f"(TMDB ID: {MOVIE_ID})."
        )

    if st.session_state.fetched:
        base_df = st.session_state.base_df

        c1, c2, c3 = st.columns(3)
        c1.metric("Movie", MOVIE_TITLE)
        c2.metric("TMDB Movie ID", MOVIE_ID)
        c3.metric("Reviews Loaded", len(base_df))

        preview = base_df[["review_id", "reviewer", "rating", "created_at", "review_text"]]
        st.dataframe(preview, use_container_width=True, height=300)


def render_analyze_section(master_df: pd.DataFrame):
    if not st.session_state.fetched:
        return

    st.markdown("---")
    st.subheader("2️⃣ Categorize + Sentiment Analysis")
    st.write(
        "Click below to generate **Category** and **Sentiment** for every review."
    )

    if st.button(
        "🧠 Categorize + Analyze Sentiment",
        type="primary",
        use_container_width=True,
    ):
        with st.spinner("Running AI analysis..."):
            simulate_analysis(st.session_state.base_df)

        final_df = master_df.copy()
        final_df["status"] = "Analyzed"
        final_df["analyzed_at"] = pd.Timestamp.now().isoformat()

        st.session_state.results = final_df
        st.session_state.analyzed = True

        save_excel(final_df)

        st.success(
            f"Completed Category + Sentiment analysis for "
            f"{len(final_df)} review(s). Saved to {EXCEL_FILE}."
        )


def render_results():
    if not st.session_state.analyzed:
        return

    df = st.session_state.results.copy()

    st.markdown("---")
    st.subheader("3️⃣ Analysis Results")

    total = len(df)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Reviews", total)
    c2.metric("😊 Positive", int((df["sentiment"] == "Positive").sum()))
    c3.metric("😠 Negative", int((df["sentiment"] == "Negative").sum()))
    c4.metric("🤔 Mixed", int((df["sentiment"] == "Mixed").sum()))

    st.markdown("### 📊 Sentiment Distribution")
    sentiment_counts = df["sentiment"].value_counts().reindex(SENTIMENTS, fill_value=0)
    st.bar_chart(sentiment_counts)

    st.markdown("### 🏷️ Category Distribution")
    category_counts = df["category"].value_counts().reindex(CATEGORIES, fill_value=0)
    st.bar_chart(category_counts)

    st.markdown("---")
    st.subheader("🔎 Filter Reviews")

    f1, f2, f3 = st.columns(3)
    with f1:
        search = st.text_input(
            "Search review / reviewer",
            placeholder="e.g. acting, casting, visuals...",
        )
    with f2:
        sentiment_filter = st.selectbox("Sentiment", ["All"] + SENTIMENTS)
    with f3:
        category_filter = st.selectbox("Category", ["All"] + CATEGORIES)

    filtered = df.copy()

    if search.strip():
        q = search.strip().lower()
        filtered = filtered[
            filtered["review_text"].fillna("").astype(str).str.lower().str.contains(q, regex=False)
            | filtered["reviewer"].fillna("").astype(str).str.lower().str.contains(q, regex=False)
        ]

    if sentiment_filter != "All":
        filtered = filtered[filtered["sentiment"] == sentiment_filter]

    if category_filter != "All":
        filtered = filtered[filtered["category"] == category_filter]

    st.caption(f"Showing {len(filtered)} of {len(df)} reviews")

    display_columns = [
        "reviewer", "rating", "category", "sentiment", "reasoning",
        "status", "review_text", "created_at", "review_url",
    ]

    st.dataframe(
        filtered[display_columns],
        use_container_width=True,
        height=600,
        column_config={
            "review_text": st.column_config.TextColumn("Review", width="large"),
            "reasoning": st.column_config.TextColumn("Reasoning", width="large"),
            "category": st.column_config.TextColumn("Category", width="medium"),
            "sentiment": st.column_config.TextColumn("Sentiment", width="small"),
            "review_url": st.column_config.LinkColumn("TMDB Review", display_text="Open Review"),
        },
    )

    st.markdown("### 📥 Download Excel")
    excel_bytes = Path(EXCEL_FILE).read_bytes()
    st.download_button(
        "⬇️ Download Odyssey Analysis Excel",
        data=excel_bytes,
        file_name=EXCEL_FILE,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


def main():
    st.set_page_config(
        page_title="The Odyssey Review Analyzer",
        page_icon="🎬",
        layout="wide",
    )

    init_state()
    render_header()

    try:
        master_df = load_master_data()
    except FileNotFoundError as exc:
        st.error(str(exc))
        return

    render_fetch_section(master_df)
    render_analyze_section(master_df)
    render_results()


if __name__ == "__main__":
    main()