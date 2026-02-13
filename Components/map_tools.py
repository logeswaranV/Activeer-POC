from io import BytesIO

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

US_STATE_ABBREVIATIONS = [
    "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA",
    "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD",
    "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ",
    "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC",
    "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY",
]


def render_map_image(highlight_states: list[str], width: int, height: int) -> BytesIO:
    normalized = {s.upper() for s in highlight_states or []}
    df = pd.DataFrame(
        {
            "state": US_STATE_ABBREVIATIONS,
            "highlight": [1 if st in normalized else 0 for st in US_STATE_ABBREVIATIONS],
        }
    )

    fig = px.choropleth(
        df,
        locations="state",
        locationmode="USA-states",
        color="highlight",
        scope="usa",
        color_continuous_scale=["#f2f2f2", "#D9544D"],
        range_color=(0, 1),
    )
    fig.update_layout(
        margin=dict(l=0, r=0, t=0, b=0),
        coloraxis_showscale=False,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )

    fig.add_trace(
        go.Scattergeo(
            locations=US_STATE_ABBREVIATIONS,
            locationmode="USA-states",
            mode="text",
            text=US_STATE_ABBREVIATIONS,
            textfont=dict(size=12, color="black"),
            showlegend=False,
            hoverinfo="skip",
        )
    )

    buf = BytesIO()
    fig.write_image(
        buf,
        format="png",
        engine="kaleido",
        width=width,
        height=height,
        scale=1,
    )
    buf.seek(0)
    return buf
