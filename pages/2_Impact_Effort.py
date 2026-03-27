import io
import math
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Impact vs Effort", layout="wide")
st.title("Recursive Impact vs Effort Matrix")

# =============================
# LOAD DATA
# =============================
def require_upload():
    if "excel_file_bytes" not in st.session_state:
        st.warning("Upload file in Home page")
        st.stop()
    return st.session_state["excel_file_bytes"], st.session_state["sheet_name"]

def load_data(file_bytes, sheet):
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet)
    df = df.dropna(how="all")
    df["Step"] = pd.to_numeric(df["Step"], errors="coerce")
    df = df[df["Step"].notna()]
    df["Duration Seconds"] = pd.to_numeric(df["Duration (Sec)"], errors="coerce").fillna(0)
    return df.reset_index(drop=True)

# =============================
# IMPACT + EFFORT SCORES
# =============================
def compute_scores(df):
    max_dur = df["Duration Seconds"].max() or 1
    df["Effort Score"] = df["Duration Seconds"] / max_dur

    df["Impact Score"] = df.apply(
        lambda r: 1 if r["Activity Type"] == "VA" else 0.3, axis=1
    )
    return df

# =============================
# LOCAL QUADRANT SPLIT
# =============================
def split_quadrants(df, xmin, xmax, ymin, ymax, depth=0):
    if len(df) <= 1:
        return [(df, xmin, xmax, ymin, ymax)]

    x_mid = (xmin + xmax) / 2
    y_mid = (ymin + ymax) / 2

    q1 = df[(df["Effort Score"] <= x_mid) & (df["Impact Score"] > y_mid)]
    q2 = df[(df["Effort Score"] > x_mid) & (df["Impact Score"] > y_mid)]
    q3 = df[(df["Effort Score"] <= x_mid) & (df["Impact Score"] <= y_mid)]
    q4 = df[(df["Effort Score"] > x_mid) & (df["Impact Score"] <= y_mid)]

    results = []
    for sub_df, bounds in zip(
        [q1, q2, q3, q4],
        [
            (xmin, x_mid, y_mid, ymax),
            (x_mid, xmax, y_mid, ymax),
            (xmin, x_mid, ymin, y_mid),
            (x_mid, xmax, ymin, y_mid),
        ],
    ):
        if len(sub_df) <= 1:
            results.append((sub_df, *bounds))
        else:
            results.extend(split_quadrants(sub_df, *bounds, depth + 1))

    return results

# =============================
# PLACE POINTS IN CELL
# =============================
def place_points(df, xmin, xmax, ymin, ymax):
    n = len(df)
    if n == 0:
        return df

    cols = math.ceil(math.sqrt(n))
    rows = math.ceil(n / cols)

    xs, ys = [], []
    for i in range(n):
        r = i // cols
        c = i % cols

        x = xmin + (c + 1) / (cols + 1) * (xmax - xmin)
        y = ymin + (r + 1) / (rows + 1) * (ymax - ymin)

        xs.append(x)
        ys.append(y)

    df["x"] = xs
    df["y"] = ys
    return df

# =============================
# DRAW GRID BOXES
# =============================
def draw_boxes(fig, cells):
    for df_cell, xmin, xmax, ymin, ymax in cells:
        fig.add_shape(
            type="rect",
            x0=xmin,
            x1=xmax,
            y0=ymin,
            y1=ymax,
            line=dict(color="#00FF88", width=2),
            fillcolor="rgba(0,255,100,0.03)"
        )

# =============================
# BUILD CHART
# =============================
def make_chart(df):
    df = compute_scores(df)

    cells = split_quadrants(df, 0, 1, 0, 1)

    plot_list = []
    for cell_df, xmin, xmax, ymin, ymax in cells:
        temp = place_points(cell_df.copy(), xmin, xmax, ymin, ymax)
        plot_list.append(temp)

    plot_df = pd.concat(plot_list)

    fig = go.Figure()

    # draw subdivision grid
    draw_boxes(fig, cells)

    fig.add_trace(
        go.Scatter(
            x=plot_df["x"],
            y=plot_df["y"],
            mode="markers+text",
            text=plot_df["Step"],
            textposition="middle center",
            marker=dict(
                size=24,
                color="rgba(0,0,0,0)",
                line=dict(color="cyan", width=2),
            ),
            customdata=plot_df[["Step","Description","Activity","Therblig / Motion Type"]],
            hovertemplate=(
                "Step: %{customdata[0]}"
                "<br>Description: %{customdata[1]}"
                "<br>Activity: %{customdata[2]}"
                "<br>Therblig: %{customdata[3]}"
                "<extra></extra>"
            ),
        )
    )

    # main quadrant lines
    fig.add_shape(type="line", x0=0.5, x1=0.5, y0=0, y1=1,
                  line=dict(color="white", width=3, dash="dash"))
    fig.add_shape(type="line", x0=0, x1=1, y0=0.5, y1=0.5,
                  line=dict(color="white", width=3, dash="dash"))

    fig.update_layout(
        xaxis=dict(range=[0,1], visible=False),
        yaxis=dict(range=[0,1], visible=False),
        plot_bgcolor="#2E4A73",
        paper_bgcolor="#2E4A73",
        height=900,
    )

    return fig

# =============================
# RUN
# =============================
file_bytes, sheet = require_upload()
df = load_data(file_bytes, sheet)

fig = make_chart(df)
st.plotly_chart(fig, use_container_width=True)