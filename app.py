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
    .block-container { padding: 1rem 1rem 3rem 1rem; }
    h1 { font-size: 1.6rem; }
    /* Tighten space between destination sections */
    .dest-header {
        margin-top: 0.2rem;
        margin-bottom: 0.1rem;
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
                "destination":       str(c0).strip(),
                "price":             float(c1),
                "departure_datetime": parse_datetime(str(c2)),
                "return_datetime":    parse_datetime(str(c3)),
            })
    return pd.DataFrame(rows).drop_duplicates().reset_index(drop=True)


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
filtered = (
    df[
        df["departure_airport"].isin(sel_airports)
        & df["destination"].isin(sel_destinations)
        & (df["price"] <= sel_max_price)
    ]
    .sort_values("price")
    .reset_index(drop=True)
)

# ── Header metrics ───────────────────────────────────────────────────────────────
st.title("✈️ Flight Comparison")

c1, c2, c3 = st.columns(3)
c1.metric("Flights found", len(filtered))
c2.metric("Cheapest", f"€{filtered['price'].min():.2f}" if len(filtered) else "—")
c3.metric("Avg price",  f"€{filtered['price'].mean():.2f}" if len(filtered) else "—")

st.divider()

if filtered.empty:
    st.info("No flights match the current filters.")
    st.stop()


# ── Helper: build display table for one destination block ────────────────────────
def build_display(dest_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    cheapest = dest_df["price"].min()
    for _, r in dest_df.iterrows():
        dep = r["departure_datetime"]
        ret = r["return_datetime"]
        nights = (ret.date() - dep.date()).days if pd.notna(dep) and pd.notna(ret) else None
        rows.append({
            "Price (€)":      r["price"],
            "Dep. date":      dep.strftime("%a %d %b") if pd.notna(dep) else "—",
            "Dep. time":      dep.strftime("%H:%M")    if pd.notna(dep) else "—",
            "Return date":    ret.strftime("%a %d %b") if pd.notna(ret) else "—",
            "Return time":    ret.strftime("%H:%M")    if pd.notna(ret) else "—",
            "Nights":         int(nights) if nights is not None else None,
            "_cheapest":      r["price"] == cheapest,
        })
    out = pd.DataFrame(rows)
    return out


# ── Per-airport tabs ──────────────────────────────────────────────────────────────
airport_list = sorted(filtered["departure_airport"].unique())
tabs = st.tabs([f"🛫 {ap}" for ap in airport_list])

for tab, airport in zip(tabs, airport_list):
    with tab:
        ap_df = filtered[filtered["departure_airport"] == airport]

        # Order destinations by cheapest available price
        dest_order = (
            ap_df.groupby("destination")["price"].min()
            .sort_values()
            .index.tolist()
        )

        for dest in dest_order:
            dest_df = ap_df[ap_df["destination"] == dest].sort_values("price")
            cheapest = dest_df["price"].min()
            n = len(dest_df)

            # Section header
            st.markdown(
                f"### 🌍 {dest} &nbsp;"
                f"<span style='font-size:0.82rem;font-weight:400;color:#666;'>"
                f"{n} option{'s' if n>1 else ''} &nbsp;·&nbsp; from €{cheapest:.2f}"
                f"</span>",
                unsafe_allow_html=True,
            )

            display = build_display(dest_df)
            is_cheapest = display.pop("_cheapest")

            st.dataframe(
                display,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Price (€)": st.column_config.NumberColumn(
                        "💶 Price",
                        format="€%.2f",
                        width="small",
                    ),
                    "Dep. date": st.column_config.TextColumn(
                        "🛫 Departs",
                        width="medium",
                    ),
                    "Dep. time": st.column_config.TextColumn(
                        "⏰ Time",
                        width="small",
                    ),
                    "Return date": st.column_config.TextColumn(
                        "🛬 Returns",
                        width="medium",
                    ),
                    "Return time": st.column_config.TextColumn(
                        "⏰ Time",
                        width="small",
                    ),
                    "Nights": st.column_config.NumberColumn(
                        "🌙 Nights",
                        width="small",
                        format="%d",
                    ),
                },
            )

        st.divider()

# ── Chart ────────────────────────────────────────────────────────────────────────
st.subheader("📊 Average price per destination")

avg_df = (
    filtered
    .groupby(["destination", "departure_airport"], as_index=False)["price"]
    .mean()
    .sort_values("price")
)

chart = (
    alt.Chart(avg_df)
    .mark_bar(cornerRadiusEnd=4)
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
            alt.Tooltip("destination:N",       title="To"),
            alt.Tooltip("price:Q", title="Avg price (€)", format=".2f"),
        ],
    )
    .properties(height=max(180, len(avg_df) * 32))
)

st.altair_chart(chart, use_container_width=True)
