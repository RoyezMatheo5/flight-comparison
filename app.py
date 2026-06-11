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
    h1 { font-size: 1.6rem; }
    h2 { font-size: 1.2rem; }

    /* Destination card */
    .dest-card {
        background: #f8f9fb;
        border: 1px solid #e0e4ea;
        border-radius: 10px;
        padding: 0.9rem 1.1rem 0.6rem 1.1rem;
        margin-bottom: 1rem;
    }
    .dest-title {
        font-size: 1.05rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
        color: #1a1a2e;
    }

    /* Flight row */
    .flight-row {
        background: white;
        border: 1px solid #e8ecf0;
        border-radius: 8px;
        padding: 0.55rem 0.8rem;
        margin-bottom: 0.4rem;
        display: flex;
        flex-wrap: wrap;
        align-items: center;
        gap: 0.5rem;
    }
    .price-badge {
        background: #4C8BF5;
        color: white;
        border-radius: 6px;
        padding: 0.2rem 0.55rem;
        font-weight: 700;
        font-size: 0.95rem;
        white-space: nowrap;
    }
    .cheapest-badge {
        background: #28a745;
    }
    .flight-detail {
        font-size: 0.82rem;
        color: #444;
        white-space: nowrap;
    }
    .flight-sep {
        color: #aaa;
        font-size: 0.75rem;
    }
    .duration-tag {
        font-size: 0.75rem;
        color: #888;
        background: #f0f2f5;
        border-radius: 4px;
        padding: 0.1rem 0.4rem;
        white-space: nowrap;
    }
</style>
""", unsafe_allow_html=True)

# ── Data loading & parsing ──────────────────────────────────────────────────────

DUTCH_DAYS = {
    "maandag": "Monday", "dinsdag": "Tuesday",
    "woensdag": "Wednesday", "woendsdag": "Wednesday",
    "donderdag": "Thursday", "vrijdag": "Friday",
    "zaterdag": "Saturday", "zondag": "Sunday",
}

DUTCH_MONTHS = {
    "januari": "January", "februari": "February", "maart": "March",
    "april": "April", "mei": "May", "juni": "June",
    "juli": "July", "augustus": "August", "augustu": "August",
    "september": "September", "oktober": "October",
    "november": "November", "december": "December",
}


def clean_datetime_str(raw: str) -> str:
    if not isinstance(raw, str):
        return raw
    s = raw.strip().lower()
    for nl, en in DUTCH_DAYS.items():
        s = s.replace(nl, en)
    for nl, en in DUTCH_MONTHS.items():
        s = s.replace(nl, en)
    return re.sub(r" {2,}", " ", s)


def parse_datetime(raw: str):
    if not isinstance(raw, str):
        return None
    cleaned = clean_datetime_str(raw).replace(" - ", " ").strip()
    for fmt in ("%A %d %B %Y %H:%M", "%A %d %B %H:%M"):
        try:
            dt = pd.to_datetime(cleaned, format=fmt)
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
        c0, c1, c2, c3 = (row[i] if i < len(row) else None for i in range(4))
        if c0 and c1 is None and c2 is None and c3 is None:
            current_airport = str(c0).strip()
            continue
        if current_airport and c0 and c1 is not None and c2 and c3:
            rows.append({
                "departure_airport": current_airport,
                "destination": str(c0).strip(),
                "price": float(c1),
                "departure_datetime": parse_datetime(str(c2)),
                "return_datetime": parse_datetime(str(c3)),
            })
    df = pd.DataFrame(rows).drop_duplicates().reset_index(drop=True)
    return df


# ── Load ────────────────────────────────────────────────────────────────────────
df = load_data()

# ── Sidebar filters ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("🔍 Filters")

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

# ── Apply filters ───────────────────────────────────────────────────────────────
filtered = df[
    df["departure_airport"].isin(sel_airports)
    & df["destination"].isin(sel_destinations)
    & (df["price"] <= sel_max_price)
].sort_values("price").reset_index(drop=True)

# ── Header ───────────────────────────────────────────────────────────────────────
st.title("✈️ Flight Comparison")

with st.container():
    c1, c2, c3 = st.columns(3)
    c1.metric("Flights found", len(filtered))
    c2.metric("Cheapest", f"€{filtered['price'].min():.2f}" if len(filtered) else "—")
    c3.metric("Avg price", f"€{filtered['price'].mean():.2f}" if len(filtered) else "—")

st.divider()

if filtered.empty:
    st.info("No flights match the current filters.")
    st.stop()

# ── Per-airport tabs ──────────────────────────────────────────────────────────
airports_in_view = sorted(filtered["departure_airport"].unique())
tabs = st.tabs([f"🛫 {ap}" for ap in airports_in_view])

for tab, airport in zip(tabs, airports_in_view):
    with tab:
        ap_df = filtered[filtered["departure_airport"] == airport]
        destinations_in_ap = ap_df.sort_values("price")["destination"].unique()

        # ── Per-destination cards ──────────────────────────────────────────────
        for dest in destinations_in_ap:
            dest_df = ap_df[ap_df["destination"] == dest].sort_values("price")
            cheapest_price = dest_df["price"].min()
            n = len(dest_df)

            flight_rows_html = ""
            for _, row in dest_df.iterrows():
                dep = row["departure_datetime"]
                ret = row["return_datetime"]

                dep_str  = dep.strftime("%a %d %b  %H:%M") if pd.notna(dep) else "—"
                ret_str  = ret.strftime("%a %d %b  %H:%M") if pd.notna(ret) else "—"

                # trip duration in nights
                if pd.notna(dep) and pd.notna(ret):
                    nights = (ret.date() - dep.date()).days
                    duration = f"{nights} nights"
                else:
                    duration = ""

                is_cheapest = row["price"] == cheapest_price
                badge_class = "price-badge cheapest-badge" if is_cheapest else "price-badge"

                flight_rows_html += f"""
                <div class="flight-row">
                    <span class="{badge_class}">€{row['price']:.2f}</span>
                    <span class="flight-detail">🛫 {dep_str}</span>
                    <span class="flight-sep">→</span>
                    <span class="flight-detail">🛬 {ret_str}</span>
                    {"<span class='duration-tag'>🌙 " + duration + "</span>" if duration else ""}
                </div>"""

            st.markdown(f"""
            <div class="dest-card">
                <div class="dest-title">🌍 {dest}
                    <span style="font-weight:400;font-size:0.82rem;color:#666;margin-left:0.5rem;">
                        {n} option{"s" if n > 1 else ""} · from €{cheapest_price:.2f}
                    </span>
                </div>
                {flight_rows_html}
            </div>
            """, unsafe_allow_html=True)

st.divider()

# ── Chart ────────────────────────────────────────────────────────────────────────
st.subheader("📊 Average price per destination")

avg_by_dest = (
    filtered.groupby(["destination", "departure_airport"], as_index=False)["price"]
    .mean()
    .sort_values("price")
)

chart = (
    alt.Chart(avg_by_dest)
    .mark_bar()
    .encode(
        x=alt.X("price:Q", title="Average price (€)", axis=alt.Axis(format=",.0f")),
        y=alt.Y("destination:N", sort="-x", title=None),
        color=alt.Color(
            "departure_airport:N",
            title="Airport",
            scale=alt.Scale(scheme="tableau10"),
        ),
        tooltip=[
            alt.Tooltip("departure_airport:N", title="From"),
            alt.Tooltip("destination:N", title="To"),
            alt.Tooltip("price:Q", title="Avg price (€)", format=".2f"),
        ],
    )
    .properties(height=max(180, len(avg_by_dest) * 30))
)

st.altair_chart(chart, use_container_width=True)
