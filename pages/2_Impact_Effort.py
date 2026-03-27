import math
import pandas as pd
import plotly.graph_objects as go

# ===============================
# POSITIONING INSIDE CELL
# ===============================
def assign_positions(df, xmin, xmax, ymin, ymax):
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


# ===============================
# LOCAL RECURSIVE SPLIT
# ===============================
def recursive_split(df, xmin, xmax, ymin, ymax):
    # stop if only 1 activity
    if len(df) <= 1:
        return [(df, xmin, xmax, ymin, ymax)]

    xmid = (xmin + xmax) / 2
    ymid = (ymin + ymax) / 2

    # split INSIDE this quadrant
    q1 = df[(df["Effort Score"] <= 0.5) & (df["Impact Score"] > 0.5)]
    q2 = df[(df["Effort Score"] > 0.5) & (df["Impact Score"] > 0.5)]
    q3 = df[(df["Effort Score"] <= 0.5) & (df["Impact Score"] <= 0.5)]
    q4 = df[(df["Effort Score"] > 0.5) & (df["Impact Score"] <= 0.5)]

    cells = []

    for q, bounds in zip(
        [q1, q2, q3, q4],
        [
            (xmin, xmid, ymid, ymax),
            (xmid, xmax, ymid, ymax),
            (xmin, xmid, ymin, ymid),
            (xmid, xmax, ymin, ymid),
        ],
    ):
        if len(q) == 0:
            continue
        elif len(q) == 1:
            cells.append((q, *bounds))
        else:
            cells.extend(recursive_split(q, *bounds))

    return cells


# ===============================
# MAIN PLOT FUNCTION
# ===============================
def make_plot(df):
    # normalize scores (VERY IMPORTANT)
    df["Effort Score"] = df["Duration Seconds"] / df["Duration Seconds"].max()
    df["Impact Score"] = df["Impact"].map({"High": 1.0, "Low": 0.2})

    cells = recursive_split(df, 0, 1, 0, 1)

    fig = go.Figure()
    plot_df_list = []

    for cell_df, xmin, xmax, ymin, ymax in cells:

        # draw strong borders
        fig.add_shape(
            type="rect",
            x0=xmin,
            x1=xmax,
            y0=ymin,
            y1=ymax,
            line=dict(color="black", width=2)
        )

        temp = assign_positions(cell_df.copy(), xmin, xmax, ymin, ymax)
        plot_df_list.append(temp)

    plot_df = pd.concat(plot_df_list)

    # scatter
    fig.add_trace(
        go.Scatter(
            x=plot_df["x"],
            y=plot_df["y"],
            mode="markers+text",
            text=plot_df["Step"],
            textposition="middle center",
            marker=dict(size=26, color="white", line=dict(width=2)),
            customdata=plot_df[
                ["Step", "Description", "Activity", "Therblig / Motion Type"]
            ].values,
            hovertemplate=
                "Step: %{customdata[0]}<br>" +
                "Description: %{customdata[1]}<br>" +
                "Activity: %{customdata[2]}<br>" +
                "Therblig: %{customdata[3]}<extra></extra>",
        )
    )

    fig.update_layout(
        title="Recursive Impact vs Effort Matrix",
        xaxis=dict(range=[0, 1], visible=False),
        yaxis=dict(range=[0, 1], visible=False),
        plot_bgcolor="#E6E6E6",
        height=900
    )

    return fig