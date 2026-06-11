import re
import streamlit as st
import pandas as pd
import altair as alt
import openpyxl

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Vluchten Vergelijken",
    page_icon="✈️",
    layout="centered",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .block-container { padding: 1rem 1rem 3rem 1rem; }
    h1 { font-size: 1.5rem; }
    h3 { font-size: 1.05rem; margin-top: 1.2rem; margin-bottom: 0.2rem; }
    /* make metric values slightly smaller so 3 fit on mobile */
    [data-testid="stMetricValue"] { font-size: 1.2rem !important; }
</style>
""", unsafe_allow_html=True)

# ── Dutch → English lookup tables ─────────────────────────────────────────────
DUTCH_DAYS = {
    "maandag": "Monday",   "dinsdag": "Tuesday",
    "woensdag": "Wednesday", "woendsdag": "Wednesday",
    "donderdag": "Thursday", "vrijdag": "Friday",
    "zaterdag": "Saturday",  "zondag": "Sunday",
}
DUTCH_MONTHS = {
    "januari": "January",  "februari": "February", "maart": "March",
    "april": "April",      "mei": "May",            "juni": "June",
    "juli": "July",        "augustus": "August",    "augustu": "August",
    "september": "September", "oktober": "October",
    "november": "November",   "december": "December",
}


def clean_dt(raw: str) -> str:
    if not isinstance(raw, str):
        return raw
    s = raw.strip().lower()
    for nl, en in DUTCH_DAYS.items():
        s = s.replace(nl, en)
    for nl, en in DUTCH_MONTHS.items():
        s = s.replace(nl, en)
    return re.sub(r" {2,}", " ", s)


def parse_dt(raw: str):
    if not isinstance(raw, str):
        return None
    c = clean_dt(raw).replace(" - ", " ").strip()
    # Strip the leading weekday name (e.g. "Monday ") — day names in the
    # source data can be wrong, so we derive the correct weekday from the date.
    c = re.sub(r"^(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\s+", "", c, flags=re.IGNORECASE)
    for fmt in ("%d %B %Y %H:%M", "%d %B %H:%M"):
        try:
            dt = pd.to_datetime(c, format=fmt)
            return dt.replace(year=2026) if dt.year == 1900 else dt
        except ValueError:
            pass
    return None


@st.cache_data
def load_data() -> pd.DataFrame:
    wb = openpyxl.load_workbook("flight_data.xlsx")
    ws = wb.active
    rows, airport = [], None
    for row in ws.iter_rows(values_only=True):
        c0, c1, c2, c3 = (row[i] if i < len(row) else None for i in range(4))
        if c0 and c1 is None and c2 is None and c3 is None:
            airport = str(c0).strip()
            continue
        if airport and c0 and c1 is not None and c2 and c3:
            rows.append({
                "departure_airport":  airport,
                "destination":        str(c0).strip(),
                "price":              float(c1),
                "departure_datetime": parse_dt(str(c2)),
                "return_datetime":    parse_dt(str(c3)),
            })
    return pd.DataFrame(rows).drop_duplicates().reset_index(drop=True)


# ── Load & filter ──────────────────────────────────────────────────────────────
df = load_data()

with st.sidebar:
    st.header("🔍 Filters")
    airports     = sorted(df["departure_airport"].unique())
    sel_airports = st.multiselect("Vertrekluchthaven", airports, default=airports)

    destinations     = sorted(df["destination"].unique())
    sel_destinations = st.multiselect("Bestemming", destinations, default=destinations)

    min_p, max_p = float(df["price"].min()), float(df["price"].max())
    sel_max_price = st.slider("Max prijs (€)", min_p, max_p, max_p, step=1.0)

filtered = (
    df[
        df["departure_airport"].isin(sel_airports)
        & df["destination"].isin(sel_destinations)
        & (df["price"] <= sel_max_price)
    ]
    .sort_values("price")
    .reset_index(drop=True)
)

# ── Page header ────────────────────────────────────────────────────────────────
st.title("✈️ Vluchten Vergelijken")

c1, c2, c3 = st.columns(3)
c1.metric("Vluchten", len(filtered))
c2.metric("Goedkoopste", f"€{filtered['price'].min():.2f}" if len(filtered) else "—")
c3.metric("Gemiddeld",   f"€{filtered['price'].mean():.2f}" if len(filtered) else "—")

st.divider()

if filtered.empty:
    st.info("Geen vluchten gevonden voor de huidige filters.")
    st.stop()


NL_DAYS = {
    "Mon": "Ma", "Tue": "Di", "Wed": "Wo",
    "Thu": "Do", "Fri": "Vr", "Sat": "Za", "Sun": "Zo",
}
NL_MONTHS = {
    "Jan": "jan", "Feb": "feb", "Mar": "mrt", "Apr": "apr",
    "May": "mei", "Jun": "jun", "Jul": "jul", "Aug": "aug",
    "Sep": "sep", "Oct": "okt", "Nov": "nov", "Dec": "dec",
}


def nl_date(dt) -> str:
    """Return a Dutch short date string, e.g. 'Za 26 jul'."""
    s = dt.strftime("%a %d %b")
    for en, nl in NL_DAYS.items():
        s = s.replace(en, nl)
    for en, nl in NL_MONTHS.items():
        s = s.replace(en, nl)
    return s


# ── Flight card (one per flight row) ──────────────────────────────────────────
def flight_card(row: pd.Series, is_cheapest: bool) -> None:
    dep = row["departure_datetime"]
    ret = row["return_datetime"]
    nights = (ret.date() - dep.date()).days if pd.notna(dep) and pd.notna(ret) else None

    dep_date = nl_date(dep)          if pd.notna(dep) else "—"
    dep_time = dep.strftime("%H:%M") if pd.notna(dep) else "—"
    ret_date = nl_date(ret)          if pd.notna(ret) else "—"
    ret_time = ret.strftime("%H:%M") if pd.notna(ret) else "—"

    with st.container(border=True):
        # Price row
        price_label = f"✅ €{row['price']:.2f}  ← goedkoopste" if is_cheapest else f"€{row['price']:.2f}"
        st.markdown(
            f"<p style='font-size:1.2rem;font-weight:700;margin:0 0 0.5rem 0;"
            f"color:{'#1a7f37' if is_cheapest else '#1a1a2e'};'>{price_label}</p>",
            unsafe_allow_html=True,
        )

        # Two columns: departure | return
        left, right = st.columns(2)
        with left:
            st.markdown("**🛫 Vertrek**")
            st.markdown(f"{dep_date}")
            st.markdown(f"🕐 **{dep_time}**")
        with right:
            st.markdown("**🛬 Terug**")
            st.markdown(f"{ret_date}")
            st.markdown(f"🕐 **{ret_time}**")

        if nights is not None:
            st.caption(f"🌙 {nights} nachten")


# ── Per-airport tabs ───────────────────────────────────────────────────────────
airport_list = sorted(filtered["departure_airport"].unique())
tabs = st.tabs([f"🛫 {ap}" for ap in airport_list])

for tab, airport in zip(tabs, airport_list):
    with tab:
        ap_df = filtered[filtered["departure_airport"] == airport]

        dest_order = (
            ap_df.groupby("destination")["price"].min()
            .sort_values().index.tolist()
        )

        for dest in dest_order:
            dest_df   = ap_df[ap_df["destination"] == dest].sort_values("price")
            cheapest  = dest_df["price"].min()
            n         = len(dest_df)

            st.markdown(
                f"### 🌍 {dest} "
                f"<span style='font-size:0.8rem;font-weight:400;color:#888;'>"
                f"{n} optie{'s' if n > 1 else ''} · vanaf €{cheapest:.2f}"
                f"</span>",
                unsafe_allow_html=True,
            )

            for _, row in dest_df.iterrows():
                flight_card(row, is_cheapest=(row["price"] == cheapest))

        st.write("")  # breathing room at bottom of tab

# ── Chart ──────────────────────────────────────────────────────────────────────
st.divider()
st.subheader("📊 Gemiddelde prijs per bestemming")

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
            title="Luchthaven",
            scale=alt.Scale(scheme="tableau10"),
        ),
        tooltip=[
            alt.Tooltip("departure_airport:N", title="Van"),
            alt.Tooltip("destination:N",       title="Naar"),
            alt.Tooltip("price:Q", title="Gem. prijs (€)", format=".2f"),
        ],
    )
    .properties(height=max(180, len(avg_df) * 32))
)

st.altair_chart(chart, use_container_width=True)
