"""
The Odyssey — TMDB Review Analyzer
----------------------------------
Fetches The Odyssey reviews from TMDB, analyzes each review with Groq,
and exports the results to Excel.

Features:
- ONLY The Odyssey is supported; no other movies are shown.
- Searches TMDB for "The Odyssey" and prefers the 2026 movie.
- Fetches up to 1000 TMDB reviews (or all available if fewer exist).
- Categorizes every review into one main movie aspect.
- Performs Positive / Negative / Neutral sentiment analysis.
- Shows BOTH category and sentiment when the user clicks Analyze.
- Saves results to Excel (.xlsx).
- Preserves previously analyzed results and marks reviews as New/Existing.
- Avoids re-analyzing reviews already present in the Excel file.

Install:
    pip install streamlit requests pandas openpyxl groq

Run:
    streamlit run review_monitor.py

API keys:
1. TMDB:
   https://www.themoviedb.org/settings/api

2. Groq:
   https://console.groq.com/keys

You can enter both keys in the Streamlit sidebar or set:
    TMDB_API_KEY=...
    GROQ_API_KEY=...

TMDB attribution:
This product uses the TMDB API but is not endorsed or certified by TMDB.
"""

import json
import os
import re
import time
from pathlib import Path
from typing import Dict, List

import pandas as pd
import requests
import streamlit as st
from openpyxl import load_workbook
from groq import Groq


# =============================================================================
# CONFIG
# =============================================================================

TMDB_BASE_URL = "https://api.themoviedb.org/3"
MOVIE_TITLE = "The Odyssey"
MOVIE_YEAR = 2026

DEFAULT_MAX_REVIEWS = 1000
DEFAULT_LANGUAGE = "en-US"

EXCEL_FILE = "odyssey_reviews_analysis.xlsx"

CATEGORIES = [
    "Acting",
    "Story/Screenplay",
    "Direction",
    "Visuals/Cinematography",
    "Music/Sound",
    "Pacing",
    "Overall Experience",
]

SENTIMENTS = ["Positive", "Negative", "Neutral"]

GROQ_MODELS = [
    "llama-3.3-70b-versatile",
    "llama-3.1-8b-instant",
]


# =============================================================================
# TMDB
# =============================================================================

def tmdb_headers(api_key: str) -> Dict[str, str]:
    return {
        "Authorization": f"Bearer {api_key}",
        "accept": "application/json",
    }


def tmdb_get(
    endpoint: str,
    api_key: str,
    params: Dict = None,
    retries: int = 4,
) -> Dict:
    """GET request with retries and useful error messages."""
    url = f"{TMDB_BASE_URL}{endpoint}"

    last_error = None

    for attempt in range(retries):
        try:
            response = requests.get(
                url,
                headers=tmdb_headers(api_key),
                params=params or {},
                timeout=30,
            )

            if response.status_code == 200:
                return response.json()

            if response.status_code in (429, 500, 502, 503, 504):
                wait = 2 ** attempt
                time.sleep(wait)
                last_error = f"TMDB HTTP {response.status_code}: {response.text[:300]}"
                continue

            raise RuntimeError(
                f"TMDB API error {response.status_code}: {response.text[:500]}"
            )

        except requests.RequestException as exc:
            last_error = str(exc)
            time.sleep(2 ** attempt)

    raise RuntimeError(f"TMDB request failed after retries: {last_error}")


def find_odyssey(api_key: str) -> Dict:
    """
    Search TMDB for The Odyssey and select the 2026 movie where possible.
    """
    data = tmdb_get(
        "/search/movie",
        api_key,
        params={
            "query": MOVIE_TITLE,
            "include_adult": "false",
            "language": "en-US",
            "page": 1,
        },
    )

    results = data.get("results", [])

    if not results:
        raise RuntimeError("TMDB could not find The Odyssey.")

    # Prefer exact title + 2026.
    exact_year = [
        x for x in results
        if str(x.get("title", "")).strip().lower() == MOVIE_TITLE.lower()
        and str(x.get("release_date", "")).startswith(str(MOVIE_YEAR))
    ]

    if exact_year:
        return exact_year[0]

    # Then exact title.
    exact_title = [
        x for x in results
        if str(x.get("title", "")).strip().lower() == MOVIE_TITLE.lower()
    ]

    if exact_title:
        return exact_title[0]

    # Last resort: highest popularity search result.
    return sorted(
        results,
        key=lambda x: x.get("popularity", 0),
        reverse=True,
    )[0]


def fetch_odyssey_reviews(
    api_key: str,
    max_reviews: int = 1000,
    language: str = "en-US",
) -> tuple[Dict, List[Dict]]:
    """
    Fetch up to max_reviews from TMDB's paginated movie review endpoint.

    TMDB's review endpoint is paginated, so this function keeps requesting
    pages until the requested limit or TMDB's available review count is reached.
    """
    movie = find_odyssey(api_key)
    movie_id = movie["id"]

    reviews: List[Dict] = []
    page = 1

    while len(reviews) < max_reviews:
        params = {
            "page": page,
        }

        if language:
            params["language"] = language

        data = tmdb_get(
            f"/movie/{movie_id}/reviews",
            api_key,
            params=params,
        )

        page_results = data.get("results", [])

        if not page_results:
            break

        for review in page_results:
            reviews.append(
                {
                    "review_id": str(review.get("id", "")),
                    "movie": MOVIE_TITLE,
                    "movie_id": movie_id,
                    "reviewer": (
                        review.get("author_details", {}).get("name")
                        or review.get("author")
                        or "Anonymous"
                    ),
                    "review_text": review.get("content", "").strip(),
                    "rating": (
                        review.get("author_details", {}).get("rating")
                    ),
                    "created_at": review.get("created_at", ""),
                    "updated_at": review.get("updated_at", ""),
                    "review_url": review.get("url", ""),
                    "source": "TMDB",
                }
            )

            if len(reviews) >= max_reviews:
                break

        total_pages = int(data.get("total_pages", page))

        if page >= total_pages:
            break

        page += 1

        # Be polite to the API.
        time.sleep(0.15)

    # Remove duplicates by TMDB review ID.
    unique = {}
    for r in reviews:
        key = r["review_id"] or (
            r["reviewer"],
            r["created_at"],
            r["review_text"][:100],
        )
        unique[key] = r

    return movie, list(unique.values())[:max_reviews]


# =============================================================================
# GROQ / LLM ANALYSIS
# =============================================================================

def build_batch_prompt(reviews: List[Dict]) -> str:
    categories = ", ".join(CATEGORIES)

    review_blocks = []

    for idx, review in enumerate(reviews):
        text = review["review_text"][:5000]
        review_blocks.append(
            f"""
REVIEW_INDEX: {idx}
REVIEW:
{text}
"""
        )

    joined = "\n".join(review_blocks)

    return f"""
You are a production-grade movie review analysis engine.

Analyze every review below.

For each review return:
1. category: exactly ONE of:
   {categories}

2. sentiment: exactly ONE of:
   Positive, Negative, Neutral

3. reasoning: a very short explanation, maximum 20 words.

Category rules:
- Acting: performance, cast, characters' acting.
- Story/Screenplay: plot, writing, adaptation, dialogue, narrative.
- Direction: director's execution, staging, filmmaking choices.
- Visuals/Cinematography: cinematography, CGI, visuals, production design.
- Music/Sound: score, soundtrack, sound design.
- Pacing: runtime, speed, slow/fast sections, structure.
- Overall Experience: general opinion when no single aspect dominates.

Sentiment rules:
- Positive = overall favorable opinion.
- Negative = overall unfavorable opinion.
- Neutral = mixed, factual, or genuinely balanced without a dominant polarity.

Return ONLY valid JSON.
Do not use markdown.
The JSON must be an array with exactly one object per REVIEW_INDEX.

Required format:
[
  {{
    "review_index": 0,
    "category": "Acting",
    "sentiment": "Positive",
    "reasoning": "Strong praise for the cast performance."
  }}
]

REVIEWS:
{joined}
"""


def extract_json_array(text: str) -> List[Dict]:
    """
    Robustly extract a JSON array from an LLM response.
    """
    text = text.strip()

    # Remove markdown fences if a model adds them.
    text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s*```$", "", text)

    start = text.find("[")
    end = text.rfind("]")

    if start == -1 or end == -1:
        raise ValueError("LLM response did not contain a JSON array.")

    parsed = json.loads(text[start:end + 1])

    if not isinstance(parsed, list):
        raise ValueError("LLM response JSON is not a list.")

    return parsed


def normalize_analysis(item: Dict) -> Dict:
    category = str(item.get("category", "")).strip()
    sentiment = str(item.get("sentiment", "")).strip()
    reasoning = str(item.get("reasoning", "")).strip()

    if category not in CATEGORIES:
        category = "Overall Experience"

    if sentiment not in SENTIMENTS:
        sentiment = "Neutral"

    return {
        "category": category,
        "sentiment": sentiment,
        "reasoning": reasoning,
    }


def analyze_batch(
    reviews: List[Dict],
    groq_api_key: str,
    model: str,
) -> List[Dict]:
    """
    Analyze a batch of reviews in one Groq call.
    """
    client = Groq(api_key=groq_api_key)

    response = client.chat.completions.create(
        model=model,
        messages=[
            {
                "role": "system",
                "content": (
                    "You classify movie reviews. "
                    "Always return valid JSON only."
                ),
            },
            {
                "role": "user",
                "content": build_batch_prompt(reviews),
            },
        ],
        temperature=0.1,
        max_tokens=max(500, len(reviews) * 120),
    )

    raw = response.choices[0].message.content or ""
    results = extract_json_array(raw)

    # Index returned results so ordering does not matter.
    by_index = {}

    for item in results:
        try:
            idx = int(item.get("review_index"))
            by_index[idx] = normalize_analysis(item)
        except Exception:
            continue

    final = []

    for idx in range(len(reviews)):
        final.append(
            by_index.get(
                idx,
                {
                    "category": "Overall Experience",
                    "sentiment": "Neutral",
                    "reasoning": "LLM did not return a valid result for this review.",
                },
            )
        )

    return final


def analyze_all_reviews(
    reviews: List[Dict],
    groq_api_key: str,
    model: str,
    batch_size: int = 10,
):
    """
    Analyze all reviews in batches.

    Smaller batches are safer for very long reviews and easier to retry.
    """
    all_results = []

    progress = st.progress(0, text="Starting AI analysis...")

    total = len(reviews)

    for start in range(0, total, batch_size):
        batch = reviews[start:start + batch_size]

        try:
            batch_results = analyze_batch(
                batch,
                groq_api_key,
                model,
            )
        except Exception as exc:
            # Retry the batch one review at a time.
            batch_results = []

            for review in batch:
                try:
                    single_result = analyze_batch(
                        [review],
                        groq_api_key,
                        model,
                    )[0]
                except Exception as single_exc:
                    single_result = {
                        "category": "Overall Experience",
                        "sentiment": "Neutral",
                        "reasoning": f"Analysis failed: {single_exc}",
                    }

                batch_results.append(single_result)

                time.sleep(0.1)

        for review, result in zip(batch, batch_results):
            enriched = dict(review)
            enriched.update(result)
            all_results.append(enriched)

        completed = min(start + len(batch), total)
        progress.progress(
            completed / total,
            text=f"Analyzed {completed}/{total} reviews...",
        )

        time.sleep(0.15)

    progress.empty()

    return all_results


# =============================================================================
# EXCEL STORAGE
# =============================================================================

OUTPUT_COLUMNS = [
    "review_id",
    "movie",
    "movie_id",
    "reviewer",
    "review_text",
    "rating",
    "created_at",
    "updated_at",
    "review_url",
    "source",
    "category",
    "sentiment",
    "reasoning",
    "status",
    "analyzed_at",
]


def load_existing_excel(path: str = EXCEL_FILE) -> pd.DataFrame:
    if not os.path.exists(path):
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    try:
        df = pd.read_excel(path)

        for col in OUTPUT_COLUMNS:
            if col not in df.columns:
                df[col] = ""

        return df[OUTPUT_COLUMNS]

    except Exception as exc:
        st.warning(f"Could not read existing Excel file: {exc}")
        return pd.DataFrame(columns=OUTPUT_COLUMNS)


def save_excel(df: pd.DataFrame, path: str = EXCEL_FILE):
    """
    Save results to Excel with basic formatting and frozen headers.
    """
    df = df.copy()

    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[OUTPUT_COLUMNS]

    df.to_excel(path, index=False, engine="openpyxl")

    wb = load_workbook(path)
    ws = wb.active

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # Reasonable column widths.
    widths = {
        "A": 25,
        "B": 18,
        "C": 12,
        "D": 22,
        "E": 70,
        "F": 10,
        "G": 24,
        "H": 24,
        "I": 55,
        "J": 12,
        "K": 25,
        "L": 15,
        "M": 55,
        "N": 12,
        "O": 24,
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    wb.save(path)


def merge_with_existing(
    analyzed_reviews: List[Dict],
    existing_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Preserve existing analyses and mark newly fetched reviews.

    If a review already exists in Excel, keep its existing category/sentiment.
    """
    existing_map = {}

    if not existing_df.empty:
        for _, row in existing_df.iterrows():
            key = str(row.get("review_id", "")).strip()
            if key:
                existing_map[key] = row.to_dict()

    rows = []

    for review in analyzed_reviews:
        review_id = str(review.get("review_id", "")).strip()

        old = existing_map.get(review_id)

        if old:
            row = dict(review)

            # Preserve prior AI classification.
            row["category"] = old.get("category", row.get("category"))
            row["sentiment"] = old.get("sentiment", row.get("sentiment"))
            row["reasoning"] = old.get("reasoning", row.get("reasoning"))
            row["status"] = "Existing"
            row["analyzed_at"] = old.get(
                "analyzed_at",
                pd.Timestamp.now().isoformat(),
            )
        else:
            row = dict(review)
            row["status"] = "New"
            row["analyzed_at"] = pd.Timestamp.now().isoformat()

        rows.append(row)

    return pd.DataFrame(rows)


# =============================================================================
# STREAMLIT UI
# =============================================================================

def init_state():
    defaults = {
        "movie": None,
        "reviews": [],
        "results": pd.DataFrame(),
        "fetched": False,
        "analyzed": False,
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def sidebar():
    st.sidebar.header("⚙️ Configuration")

    tmdb_key = st.sidebar.text_input(
        "TMDB API Read Access Token",
        value=os.getenv("TMDB_API_KEY", ""),
        type="password",
        help="Create an API credential from your TMDB account.",
    )

    groq_key = st.sidebar.text_input(
        "Groq API Key",
        value=os.getenv("GROQ_API_KEY", ""),
        type="password",
    )

    model = st.sidebar.selectbox(
        "Groq Model",
        GROQ_MODELS,
        index=0,
    )

    max_reviews = st.sidebar.number_input(
        "Maximum reviews",
        min_value=1,
        max_value=1000,
        value=DEFAULT_MAX_REVIEWS,
        step=50,
    )

    language = st.sidebar.selectbox(
        "TMDB review language",
        ["en-US"],
        index=0,
    )

    st.sidebar.markdown("---")

    st.sidebar.info(
        "Only The Odyssey is loaded. "
        "The app can fetch up to 1000 reviews from TMDB."
    )

    st.sidebar.caption(
        "This product uses the TMDB API but is not endorsed or certified by TMDB."
    )

    return (
        tmdb_key.strip(),
        groq_key.strip(),
        model,
        int(max_reviews),
        language,
    )


def render_header():
    st.title("🎬 The Odyssey — Review Analyzer")
    st.caption(
        "TMDB reviews → AI category + sentiment → Excel"
    )

    st.markdown(
        """
        **Movie:** The Odyssey  
        **Source:** TMDB  
        **AI:** Groq  
        **Outputs:** Category + Sentiment + Reasoning
        """
    )


def render_fetch_section(
    tmdb_key: str,
    max_reviews: int,
    language: str,
):
    st.subheader("1️⃣ Fetch The Odyssey Reviews")

    if st.button(
        f"📥 Fetch up to {max_reviews} Odyssey reviews",
        type="primary",
        use_container_width=True,
    ):
        if not tmdb_key:
            st.error("Enter your TMDB API key/token in the sidebar.")
            return

        try:
            with st.spinner("Finding The Odyssey on TMDB..."):
                movie, reviews = fetch_odyssey_reviews(
                    tmdb_key,
                    max_reviews=max_reviews,
                    language=language,
                )

            st.session_state.movie = movie
            st.session_state.reviews = reviews
            st.session_state.fetched = True
            st.session_state.analyzed = False
            st.session_state.results = pd.DataFrame()

            st.success(
                f"Fetched {len(reviews)} review(s) for "
                f"{movie.get('title', MOVIE_TITLE)} "
                f"(TMDB ID: {movie.get('id')})."
            )

        except Exception as exc:
            st.error(f"Could not fetch TMDB reviews: {exc}")

    if st.session_state.fetched:
        movie = st.session_state.movie
        reviews = st.session_state.reviews

        c1, c2, c3 = st.columns(3)
        c1.metric("Movie", movie.get("title", MOVIE_TITLE))
        c2.metric("TMDB Movie ID", movie.get("id", "N/A"))
        c3.metric("Reviews Loaded", len(reviews))

        if reviews:
            preview = pd.DataFrame(reviews)[
                [
                    "review_id",
                    "reviewer",
                    "rating",
                    "created_at",
                    "review_text",
                ]
            ]

            st.dataframe(
                preview,
                use_container_width=True,
                height=300,
            )


def render_analyze_section(
    groq_key: str,
    model: str,
):
    if not st.session_state.fetched:
        return

    st.markdown("---")
    st.subheader("2️⃣ Categorize + Sentiment Analysis")

    st.write(
        "Click the button below. The same action generates both "
        "**Category** and **Sentiment** for every Odyssey review."
    )

    if st.button(
        "🧠 Categorize + Analyze Sentiment",
        type="primary",
        use_container_width=True,
    ):
        if not groq_key:
            st.error("Enter your Groq API key in the sidebar.")
            return

        reviews = st.session_state.reviews

        if not reviews:
            st.warning("No reviews available.")
            return

        existing = load_existing_excel()

        # Only analyze reviews that do not already have a classification.
        existing_ids = set()

        if not existing.empty:
            existing_ids = set(
                existing["review_id"]
                .astype(str)
                .str.strip()
                .tolist()
            )

        # If the review is already in Excel with category + sentiment,
        # reuse the existing analysis.
        cached = {}
        if not existing.empty:
            for _, row in existing.iterrows():
                rid = str(row.get("review_id", "")).strip()

                if (
                    rid
                    and str(row.get("category", "")).strip()
                    and str(row.get("sentiment", "")).strip()
                ):
                    cached[rid] = row.to_dict()

        to_analyze = [
            r for r in reviews
            if str(r.get("review_id", "")).strip() not in cached
        ]

        st.info(
            f"{len(cached)} review(s) already analyzed; "
            f"{len(to_analyze)} review(s) need AI analysis."
        )

        newly_analyzed = []

        if to_analyze:
            newly_analyzed = analyze_all_reviews(
                to_analyze,
                groq_key,
                model,
                batch_size=10,
            )

        # Reconstruct final dataset in fetched-review order.
        analyzed_map = {
            str(r.get("review_id", "")).strip(): r
            for r in newly_analyzed
        }

        final_rows = []

        for review in reviews:
            rid = str(review.get("review_id", "")).strip()

            if rid in cached:
                old = cached[rid]
                row = dict(review)
                row["category"] = old.get("category", "Overall Experience")
                row["sentiment"] = old.get("sentiment", "Neutral")
                row["reasoning"] = old.get("reasoning", "")
                row["status"] = "Existing"
                row["analyzed_at"] = old.get(
                    "analyzed_at",
                    pd.Timestamp.now().isoformat(),
                )
            else:
                row = analyzed_map.get(rid, dict(review))
                row["status"] = "New"
                row["analyzed_at"] = pd.Timestamp.now().isoformat()

            final_rows.append(row)

        final_df = pd.DataFrame(final_rows)

        # Merge with previous rows so the Excel workbook remains useful
        # if TMDB changes or older reviews are no longer returned.
        if not existing.empty:
            old_ids = set(existing["review_id"].astype(str))

            old_rows = existing[
                ~existing["review_id"].astype(str).isin(
                    [str(r.get("review_id", "")) for r in final_rows]
                )
            ].copy()

            if not old_rows.empty:
                final_df = pd.concat(
                    [old_rows, final_df],
                    ignore_index=True,
                )

        final_df = final_df[OUTPUT_COLUMNS]

        st.session_state.results = final_df
        st.session_state.analyzed = True

        save_excel(final_df)

        st.success(
            f"Completed Category + Sentiment analysis for "
            f"{len(final_rows)} Odyssey review(s). "
            f"Saved to {EXCEL_FILE}."
        )


def render_results():
    if not st.session_state.analyzed:
        return

    df = st.session_state.results.copy()

    st.markdown("---")
    st.subheader("3️⃣ Analysis Results")

    total = len(df)
    positive = int((df["sentiment"] == "Positive").sum())
    negative = int((df["sentiment"] == "Negative").sum())
    neutral = int((df["sentiment"] == "Neutral").sum())

    c1, c2, c3, c4 = st.columns(4)

    c1.metric("Total Reviews", total)
    c2.metric("😊 Positive", positive)
    c3.metric("😠 Negative", negative)
    c4.metric("😐 Neutral", neutral)

    st.markdown("### 📊 Sentiment Distribution")
    sentiment_counts = (
        df["sentiment"]
        .value_counts()
        .reindex(SENTIMENTS, fill_value=0)
    )

    st.bar_chart(sentiment_counts)

    st.markdown("### 🏷️ Category Distribution")
    category_counts = (
        df["category"]
        .value_counts()
        .reindex(CATEGORIES, fill_value=0)
    )

    st.bar_chart(category_counts)

    st.markdown("---")
    st.subheader("🔎 Filter Odyssey Reviews")

    f1, f2, f3 = st.columns(3)

    with f1:
        search = st.text_input(
            "Search review / reviewer",
            placeholder="e.g. acting, Nolan, visuals...",
        )

    with f2:
        sentiment_filter = st.selectbox(
            "Sentiment",
            ["All"] + SENTIMENTS,
        )

    with f3:
        category_filter = st.selectbox(
            "Category",
            ["All"] + CATEGORIES,
        )

    filtered = df.copy()

    if search.strip():
        q = search.strip().lower()

        filtered = filtered[
            filtered["review_text"]
            .fillna("")
            .astype(str)
            .str.lower()
            .str.contains(q, regex=False)
            |
            filtered["reviewer"]
            .fillna("")
            .astype(str)
            .str.lower()
            .str.contains(q, regex=False)
        ]

    if sentiment_filter != "All":
        filtered = filtered[
            filtered["sentiment"] == sentiment_filter
        ]

    if category_filter != "All":
        filtered = filtered[
            filtered["category"] == category_filter
        ]

    st.caption(
        f"Showing {len(filtered)} of {len(df)} reviews"
    )

    # IMPORTANT:
    # Both category and sentiment are shown together for every review.
    display_columns = [
        "reviewer",
        "rating",
        "category",
        "sentiment",
        "status",
        "review_text",
        "created_at",
        "review_url",
    ]

    st.dataframe(
        filtered[display_columns],
        use_container_width=True,
        height=600,
        column_config={
            "review_text": st.column_config.TextColumn(
                "Review",
                width="large",
            ),
            "category": st.column_config.TextColumn(
                "Category",
                width="medium",
            ),
            "sentiment": st.column_config.TextColumn(
                "Sentiment",
                width="small",
            ),
            "review_url": st.column_config.LinkColumn(
                "TMDB Review",
                display_text="Open Review",
            ),
        },
    )

    st.markdown("### 📥 Download Excel")

    excel_bytes = Path(EXCEL_FILE).read_bytes()

    st.download_button(
        "⬇️ Download Odyssey Analysis Excel",
        data=excel_bytes,
        file_name=EXCEL_FILE,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
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

    (
        tmdb_key,
        groq_key,
        model,
        max_reviews,
        language,
    ) = sidebar()

    render_fetch_section(
        tmdb_key,
        max_reviews,
        language,
    )

    render_analyze_section(
        groq_key,
        model,
    )

    render_results()


if __name__ == "__main__":
    main()
