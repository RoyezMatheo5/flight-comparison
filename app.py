import io
import re
import streamlit as st
import pandas as pd
import altair as alt
import openpyxl

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Flight Comparison",
    page_icon="✈️",
    layout="centered",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .block-container { padding: 1rem; }
    .stDataFrame { width: 100% !important; }
    h1 { font-size: 1.6rem; }
    h2 { font-size: 1.2rem; }
</style>
""", unsafe_allow_html=True)

# ── Data loading & parsing ──────────────────────────────────────────────────────

DUTCH_DAYS = {
    "maandag": "Monday", "dinsdag": "Tuesday",
    "woensdag": "Wednesday", "woendsdag": "Wednesday",  # typo fix
    "donderdag": "Thursday", "vrijdag": "Friday",
    "zaterdag": "Saturday", "zondag": "Sunday",
}

DUTCH_MONTHS = {
    "januari": "January", "februari": "February", "maart": "March",
    "april": "April", "mei": "May", "juni": "June",
    "juli": "July", "augustus": "August", "augustu": "August",  # typo fix
    "september": "September", "oktober": "October",
    "november": "November", "december": "December",
}


def clean_datetime_str(raw: str) -> str:
    """Normalise Dutch date strings with typos to English before parsing."""
    if not isinstance(raw, str):
        return raw
    s = raw.strip().lower()
    for nl, en in DUTCH_DAYS.items():
        s = s.replace(nl, en)
    for nl, en in DUTCH_MONTHS.items():
        s = s.replace(nl, en)
    # collapse multiple spaces
    s = re.sub(r" {2,}", " ", s)
    return s


def parse_datetime(raw: str) -> pd.Timestamp | None:
    if not isinstance(raw, str):
        return None
    cleaned = clean_datetime_str(raw)
    # expected format after cleaning: "Monday 27 July - 21:15"
    cleaned = cleaned.replace(" - ", " ").strip()
    for fmt in ("%A %d %B %Y %H:%M", "%A %d %B %H:%M"):
        try:
            dt = pd.to_datetime(cleaned, format=fmt)
            # If year is missing pandas defaults to 1900; inject current/next year
            if dt.year == 1900:
                dt = dt.replace(year=2025)
            return dt
        except ValueError:
            pass
    return None


@st.cache_data
def load_data() -> pd.DataFrame:
    wb = openpyxl.load_workbook("flight_data.xlsx")
    ws = wb.active

    rows = []
    current_airport = None

    for row in ws.iter_rows(values_only=True):
        col0, col1, col2, col3 = (row[i] if i < len(row) else None for i in range(4))

        # Section header row: first cell has a name, rest are None
        if col0 and col1 is None and col2 is None and col3 is None:
            current_airport = str(col0).strip()
            continue

        # Data row
        if current_airport and col0 and col1 is not None and col2 and col3:
            rows.append({
                "departure_airport": current_airport,
                "destination": str(col0).strip(),
                "price": float(col1),
                "departure_datetime": parse_datetime(str(col2)),
                "return_datetime": parse_datetime(str(col3)),
            })

    df = pd.DataFrame(rows)
    df.drop_duplicates(inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


# ── Load ────────────────────────────────────────────────────────────────────────
df = load_data()

# ── Sidebar filters ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Filters")

    airports = sorted(df["departure_airport"].unique())
    sel_airports = st.multiselect("Departure airport", airports, default=airports)

    destinations = sorted(df["destination"].unique())
    sel_destinations = st.multiselect("Destination", destinations, default=destinations)

    min_price = float(df["price"].min())
    max_price = float(df["price"].max())
    sel_max_price = st.slider(
        "Max price (€)",
        min_value=min_price,
        max_value=max_price,
        value=max_price,
        step=1.0,
    )

    dep_dates = df["departure_datetime"].dropna()
    min_date = dep_dates.min().date()
    max_date = dep_dates.max().date()
    sel_date_range = st.date_input(
        "Departure date range",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
    )

# ── Apply filters ───────────────────────────────────────────────────────────────
filtered = df[
    df["departure_airport"].isin(sel_airports)
    & df["destination"].isin(sel_destinations)
    & (df["price"] <= sel_max_price)
]

if isinstance(sel_date_range, (list, tuple)) and len(sel_date_range) == 2:
    d_start, d_end = pd.Timestamp(sel_date_range[0]), pd.Timestamp(sel_date_range[1])
    filtered = filtered[
        filtered["departure_datetime"].notna()
        & (filtered["departure_datetime"] >= d_start)
        & (filtered["departure_datetime"] <= d_end + pd.Timedelta(days=1))
    ]

filtered = filtered.sort_values("price").reset_index(drop=True)

# ── Main content ────────────────────────────────────────────────────────────────
st.title("✈️ Flight Comparison")

with st.container():
    col1, col2, col3 = st.columns(3)
    col1.metric("Flights found", len(filtered))
    col2.metric("Cheapest", f"€{filtered['price'].min():.2f}" if len(filtered) else "—")
    col3.metric("Avg price", f"€{filtered['price'].mean():.2f}" if len(filtered) else "—")

st.divider()

# ── Table ───────────────────────────────────────────────────────────────────────
st.subheader("Results")

if filtered.empty:
    st.info("No flights match the current filters.")
else:
    display = filtered.copy()
    display["departure_datetime"] = display["departure_datetime"].dt.strftime("%a %d %b %Y  %H:%M")
    display["return_datetime"] = display["return_datetime"].dt.strftime("%a %d %b %Y  %H:%M")
    display["price"] = display["price"].map("€{:.2f}".format)
    display.columns = ["From", "To", "Price", "Departure", "Return"]
    st.dataframe(display, use_container_width=True, hide_index=True)

st.divider()

# ── Chart ───────────────────────────────────────────────────────────────────────
st.subheader("Average price per destination")

if not filtered.empty:
    avg_by_dest = (
        filtered.groupby("destination", as_index=False)["price"]
        .mean()
        .sort_values("price")
    )

    chart = (
        alt.Chart(avg_by_dest)
        .mark_bar(color="#4C8BF5")
        .encode(
            x=alt.X("price:Q", title="Average price (€)", axis=alt.Axis(format=",.0f")),
            y=alt.Y("destination:N", sort="-x", title="Destination"),
            tooltip=[
                alt.Tooltip("destination:N", title="Destination"),
                alt.Tooltip("price:Q", title="Avg price (€)", format=".2f"),
            ],
        )
        .properties(height=max(150, len(avg_by_dest) * 35))
    )
    st.altair_chart(chart, use_container_width=True)
else:
    st.info("No data to chart.")
