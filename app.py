import re
from io import BytesIO
from typing import List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
from plotly.subplots import make_subplots
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="周报数据可视化看板", layout="wide")


# -----------------------------
# Helpers
# -----------------------------
REQUIRED_SHEETS = ["视频维度数据", "账号维度数据"]

VIDEO_COLS = {
    "week": "日期",
    "video_id": "唯一标识符",
    "account": "用户名",
    "publish_time": "发布的日期和时间",
    "url": "可分享 URL",
    "caption": "字幕/描述",
    "playlist": "关联片单",
    "views": "播放次数",
    "duration": "视频时长（以秒为单位）",
    "likes": "点赞数",
    "comments": "评论数",
    "shares": "分享的次数",
    "unique_viewers": "观看视频的观众数（去重后）",
    "completion_rate": "观看完整视频的观众百分比",
    "total_watch_time": "基于所有播放量的视频播放时长",
    "avg_watch_time": "基于所有播放量的视频平均播放时长",
    "favorites": "视频被收藏的总次数",
    "new_followers": "该视频带来的新粉丝总数",
}

ACCOUNT_COLS = {
    "week": "日期",
    "account": "用户名",
    "display_name": "显示名称",
    "post_count_1": "发布的公开视频总数",
    "post_count_2": "发布的公开视频总数v2",
    "new_followers": "粉丝数",
    "total_followers": "总粉丝数",
    "likes": "视频收到的点赞数",
    "comments": "视频收到的评论数",
    "shares": "视频被分享的次数",
    "profile_views": "主页被浏览的总次数",
    "views": "视频的播放次数",
}


@st.cache_data(show_spinner=False)
def load_workbook(file_bytes: bytes) -> Tuple[pd.DataFrame, pd.DataFrame]:
    excel_file = pd.ExcelFile(BytesIO(file_bytes))
    missing_sheets = [s for s in REQUIRED_SHEETS if s not in excel_file.sheet_names]
    if missing_sheets:
        raise ValueError(f"缺少工作表: {', '.join(missing_sheets)}")

    video_df = pd.read_excel(BytesIO(file_bytes), sheet_name="视频维度数据")
    account_df = pd.read_excel(BytesIO(file_bytes), sheet_name="账号维度数据")

    missing_video_cols = [v for v in VIDEO_COLS.values() if v not in video_df.columns]
    missing_account_cols = [v for v in ACCOUNT_COLS.values() if v not in account_df.columns]
    if missing_video_cols:
        raise ValueError(f"视频维度数据缺少字段: {', '.join(missing_video_cols)}")
    if missing_account_cols:
        raise ValueError(f"账号维度数据缺少字段: {', '.join(missing_account_cols)}")

    return preprocess_video(video_df), preprocess_account(account_df)


def parse_week_start(week_str: str) -> pd.Timestamp:
    match = re.match(r"\s*(\d{4}-\d{2}-\d{2})\s*~\s*(\d{4}-\d{2}-\d{2})\s*", str(week_str))
    if match:
        return pd.to_datetime(match.group(1), errors="coerce")
    return pd.NaT


def format_week_label(week_str: str) -> str:
    return str(week_str)


def preprocess_video(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for key, col in VIDEO_COLS.items():
        if key in {"week", "account", "caption", "url", "playlist", "video_id", "publish_time"}:
            continue
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["week_start"] = df[VIDEO_COLS["week"]].apply(parse_week_start)
    df[VIDEO_COLS["publish_time"]] = pd.to_datetime(df[VIDEO_COLS["publish_time"]], errors="coerce")
    df["互动率"] = np.where(
        df[VIDEO_COLS["views"]] > 0,
        (df[VIDEO_COLS["likes"]] + df[VIDEO_COLS["comments"]]) / df[VIDEO_COLS["views"]],
        np.nan,
    )
    df["完播率_decimal"] = np.where(
        df[VIDEO_COLS["completion_rate"]] > 1,
        df[VIDEO_COLS["completion_rate"]] / 100,
        df[VIDEO_COLS["completion_rate"]],
    )
    df["字幕简写"] = df[VIDEO_COLS["caption"]].fillna("").astype(str).str.slice(0, 80)
    return df


def preprocess_account(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    numeric_cols = [
        ACCOUNT_COLS["post_count_1"],
        ACCOUNT_COLS["post_count_2"],
        ACCOUNT_COLS["new_followers"],
        ACCOUNT_COLS["total_followers"],
        ACCOUNT_COLS["likes"],
        ACCOUNT_COLS["comments"],
        ACCOUNT_COLS["shares"],
        ACCOUNT_COLS["profile_views"],
        ACCOUNT_COLS["views"],
    ]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["week_start"] = df[ACCOUNT_COLS["week"]].apply(parse_week_start)
    df["发帖量"] = df[ACCOUNT_COLS["post_count_2"]].fillna(df[ACCOUNT_COLS["post_count_1"]])
    df["发帖量"] = df["发帖量"].fillna(df[ACCOUNT_COLS["post_count_1"]])
    return df


def get_week_options(account_df: pd.DataFrame) -> List[str]:
    weeks = (
        account_df[[ACCOUNT_COLS["week"], "week_start"]]
        .drop_duplicates()
        .sort_values("week_start")
    )
    return weeks[ACCOUNT_COLS["week"]].tolist()


def get_previous_week(weeks: List[str], selected_week: str) -> Optional[str]:
    try:
        idx = weeks.index(selected_week)
    except ValueError:
        return None
    return weeks[idx - 1] if idx > 0 else None


def fmt_number(value, digits: int = 0) -> str:
    if pd.isna(value):
        return "-"
    if digits == 0:
        return f"{value:,.0f}"
    return f"{value:,.{digits}f}"


def pct_change(current: float, previous: Optional[float]) -> Optional[float]:
    if previous is None or pd.isna(previous) or previous == 0:
        return None
    return (current - previous) / previous


def metric_html(label: str, value: str, delta: Optional[float]) -> str:
    if delta is None or pd.isna(delta):
        delta_html = '<span style="color:#6b7280;font-size:0.9rem;">无上周数据</span>'
    else:
        arrow = "▲" if delta >= 0 else "▼"
        color = "#16a34a" if delta >= 0 else "#dc2626"
        delta_html = f'<span style="color:{color};font-size:0.95rem;font-weight:600;">{arrow} {abs(delta):.1%}</span>'
    return f"""
    <div style="padding:18px 18px 12px 18px;border:1px solid #e5e7eb;border-radius:16px;background:#ffffff;box-shadow:0 1px 3px rgba(0,0,0,0.04);">
        <div style="font-size:0.95rem;color:#6b7280;margin-bottom:6px;">{label}</div>
        <div style="font-size:1.9rem;font-weight:700;color:#111827;line-height:1.2;">{value}</div>
        <div style="margin-top:6px;">{delta_html}</div>
    </div>
    """


def extract_hashtags(text: str) -> List[str]:
    if pd.isna(text):
        return []
    tags = re.findall(r"#([\w\u4e00-\u9fff]+)", str(text))
    return [f"#{tag.lower()}" for tag in tags]


def build_overview_metrics(account_week: pd.DataFrame, video_week: pd.DataFrame, prev_account_week: Optional[pd.DataFrame]):
    current = {
        "总播放量": float(account_week[ACCOUNT_COLS["views"]].sum()),
        "总新增粉丝": float(account_week[ACCOUNT_COLS["new_followers"]].sum()),
        "总点赞数": float(account_week[ACCOUNT_COLS["likes"]].sum()),
        "活跃账号数": float((account_week["发帖量"].fillna(0) > 0).sum()),
    }

    previous = {
        "总播放量": float(prev_account_week[ACCOUNT_COLS["views"]].sum()) if prev_account_week is not None else None,
        "总新增粉丝": float(prev_account_week[ACCOUNT_COLS["new_followers"]].sum()) if prev_account_week is not None else None,
        "总点赞数": float(prev_account_week[ACCOUNT_COLS["likes"]].sum()) if prev_account_week is not None else None,
        "活跃账号数": float((prev_account_week["发帖量"].fillna(0) > 0).sum()) if prev_account_week is not None else None,
    }
    return current, previous


def account_summary_table(
    account_df: pd.DataFrame,
    selected_week: str,
    previous_week: Optional[str]
) -> pd.DataFrame:
    """
    生成账号维度汇总表，包含：
    - 周播放量
    - 周播放量环比%
    - 日播
    - 日播环比%
    - 发帖量
    - 发帖量环比%
    - 条均
    - 条均环比%

    参数：
    account_df : pd.DataFrame
        账号维度原始数据表。

    selected_week : str
        当前选中的周次。

    previous_week : Optional[str]
        上一周周次；如果没有上一周，则传 None。

    返回：
    pd.DataFrame
        按周播放量降序排列的账号汇总表。
    """

    # 当前周数据
    current = account_df[account_df[ACCOUNT_COLS["week"]] == selected_week].copy()
    current = current[[ACCOUNT_COLS["account"], ACCOUNT_COLS["views"], "发帖量"]].copy()

    current.rename(columns={ACCOUNT_COLS["views"]: "周播放量"}, inplace=True)
    current["日播"] = current["周播放量"] / 7
    current["条均"] = np.where(
        current["发帖量"] > 0,
        current["周播放量"] / current["发帖量"],
        np.nan
    )

    # 如果有上周，则构造上周同口径指标
    if previous_week:
        prev = account_df[account_df[ACCOUNT_COLS["week"]] == previous_week].copy()
        prev = prev[[ACCOUNT_COLS["account"], ACCOUNT_COLS["views"], "发帖量"]].copy()

        prev.rename(columns={ACCOUNT_COLS["views"]: "上周周播放量", "发帖量": "上周发帖量"}, inplace=True)
        prev["上周日播"] = prev["上周周播放量"] / 7
        prev["上周条均"] = np.where(
            prev["上周发帖量"] > 0,
            prev["上周周播放量"] / prev["上周发帖量"],
            np.nan
        )

        current = current.merge(prev, on=ACCOUNT_COLS["account"], how="left")

        # 周播放量环比
        current["周播放量环比%"] = np.where(
            current["上周周播放量"] > 0,
            (current["周播放量"] - current["上周周播放量"]) / current["上周周播放量"],
            np.nan
        )

        # 日播环比
        current["日播环比%"] = np.where(
            current["上周日播"] > 0,
            (current["日播"] - current["上周日播"]) / current["上周日播"],
            np.nan
        )

        # 发帖量环比
        current["发帖量环比%"] = np.where(
            current["上周发帖量"] > 0,
            (current["发帖量"] - current["上周发帖量"]) / current["上周发帖量"],
            np.nan
        )

        # 条均环比
        current["条均环比%"] = np.where(
            current["上周条均"] > 0,
            (current["条均"] - current["上周条均"]) / current["上周条均"],
            np.nan
        )

    else:
        current["周播放量环比%"] = np.nan
        current["日播环比%"] = np.nan
        current["发帖量环比%"] = np.nan
        current["条均环比%"] = np.nan

    # 调整输出顺序，让环比紧跟在各字段后面
    result = current[
        [
            ACCOUNT_COLS["account"],
            "周播放量", "周播放量环比%",
            "日播", "日播环比%",
            "发帖量", "发帖量环比%",
            "条均", "条均环比%"
        ]
    ].copy()

    return result.sort_values("周播放量", ascending=False)


def style_account_table(df: pd.DataFrame):
    def color_change(val):
        if pd.isna(val):
            return "color:#6b7280;"
        return "color:#16a34a;font-weight:600;" if val >= 0 else "color:#dc2626;font-weight:600;"

    show_df = df[
        [
            ACCOUNT_COLS["account"],
            "周播放量", "周播放量环比%",
            "日播", "日播环比%",
            "发帖量", "发帖量环比%",
            "条均", "条均环比%",
        ]
    ].copy()

    ratio_cols = ["周播放量环比%", "日播环比%", "发帖量环比%", "条均环比%"]

    styler = (
        show_df.style
        .format(
            {
                "周播放量": "{:,.0f}",
                "周播放量环比%": "{:+.1%}",
                "日播": "{:,.0f}",
                "日播环比%": "{:+.1%}",
                "发帖量": "{:,.0f}",
                "发帖量环比%": "{:+.1%}",
                "条均": "{:,.0f}",
                "条均环比%": "{:+.1%}",
            },
            na_rep="-",
        )
        .map(color_change, subset=ratio_cols)
    )

    return styler


def render_tab_overview(account_df: pd.DataFrame, video_df: pd.DataFrame, weeks: List[str]):
    st.subheader("Tab1 整体概览")
    selected_week = st.selectbox("选择周次", weeks, index=len(weeks) - 1, key="overview_week")
    previous_week = get_previous_week(weeks, selected_week)

    account_week = account_df[account_df[ACCOUNT_COLS["week"]] == selected_week]
    video_week = video_df[video_df[VIDEO_COLS["week"]] == selected_week]
    prev_account_week = account_df[account_df[ACCOUNT_COLS["week"]] == previous_week] if previous_week else None

    current, previous = build_overview_metrics(account_week, video_week, prev_account_week)

    cols = st.columns(4)
    labels = ["总播放量", "总新增粉丝", "总点赞数", "活跃账号数"]
    for c, label in zip(cols, labels):
        value = fmt_number(current[label])
        delta = pct_change(current[label], previous[label])
        c.markdown(metric_html(label, value, delta), unsafe_allow_html=True)

    st.markdown("### 当周趋势速览")
    weekly_trend = (
        account_df.groupby(ACCOUNT_COLS["week"], as_index=False)
        .agg({
            ACCOUNT_COLS["views"]: "sum",
            ACCOUNT_COLS["new_followers"]: "sum",
            ACCOUNT_COLS["likes"]: "sum",
        })
    )
    week_order_df = account_df[[ACCOUNT_COLS["week"], "week_start"]].drop_duplicates().sort_values("week_start")
    weekly_trend = weekly_trend.merge(week_order_df, on=ACCOUNT_COLS["week"], how="left").sort_values("week_start")

    weekly_trend = weekly_trend.sort_values("week_start").tail(4)
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    # 左轴：播放量（折线图）
    fig.add_trace(
        go.Scatter(
            x=weekly_trend[ACCOUNT_COLS["week"]],
            y=weekly_trend[ACCOUNT_COLS["views"]],
            mode="lines+markers",
            name="播放量",
        ),
        secondary_y=False,
    )

    # 右轴：新增粉丝数（柱形图）
    fig.add_trace(
        go.Bar(
            x=weekly_trend[ACCOUNT_COLS["week"]],
            y=weekly_trend[ACCOUNT_COLS["new_followers"]],
            name="新增粉丝数",
            opacity=0.6,
        ),
        secondary_y=True,
    )

    fig.update_layout(
        title="周度播放量与新增粉丝趋势(<=4周)",
        height=400,
        legend_title_text="指标",
        xaxis_title="周次",
    )

    left_max = weekly_trend[ACCOUNT_COLS["views"]].max()
    left_min = weekly_trend[ACCOUNT_COLS["views"]].min()
    fig.update_yaxes(title_text="播放量", secondary_y=False,range=[left_min*0.4,left_max*1.2])
    fig.update_yaxes(title_text="新增粉丝数", secondary_y=True)

    st.plotly_chart(fig, use_container_width=True)
    with st.expander("查看当前周底层明细", expanded=False):
        c1, c2 = st.columns(2)
        account_week_display=account_week.sort_values("视频的播放次数", ascending=False).drop(columns=["week_start"])
        video_week_display=video_week.sort_values("播放次数", ascending=False).drop(columns=["week_start"])
        with c1:
            st.caption("账号维度数据")
            st.dataframe(account_week_display, use_container_width=True, hide_index=True)
        with c2:
            st.caption("视频维度数据")
            st.dataframe(video_week_display.head(100), use_container_width=True, hide_index=True)


def render_tab_accounts(account_df: pd.DataFrame, weeks: List[str]):
    st.subheader("Tab2 账号维度")

    # =========================
    # 顶部全局筛选
    # =========================
    selected_week = st.selectbox(
        "选择周次",
        weeks,
        index=len(weeks) - 1,
        key="account_week"
    )
    previous_week = get_previous_week(weeks, selected_week)

    current = account_df[account_df[ACCOUNT_COLS["week"]] == selected_week].copy()
    current["条均"] = np.where(
        current["发帖量"] > 0,
        current[ACCOUNT_COLS["views"]] / current["发帖量"],
        np.nan
    )

    # =========================
    # 账号汇总表
    # =========================
    summary_df = account_summary_table(account_df, selected_week, previous_week)
    st.markdown("### 账号汇总表")
    st.dataframe(
        style_account_table(summary_df),
        use_container_width=True,
        hide_index=True
    )

    st.markdown("---")

    # =========================
    # 当前周账号横截面表现
    # =========================
    st.markdown("### 当前周账号表现分布")
    st.caption("以下图表均基于当前所选周次，用于观察账号在本周的结构分布与表现差异。")

    left, right = st.columns([1.05, 1.25], gap="large")

    # 左侧：饼图
    with left:
        with st.container(border=True):
            st.markdown("#### 账号播放量占比")

            pie_fig = px.pie(
                current,
                names=ACCOUNT_COLS["account"],
                values=ACCOUNT_COLS["views"],
                hole=0.5,
            )

            pie_fig.update_traces(
                textposition="inside",
                textinfo="percent",
                hovertemplate=(
                    "<b>%{label}</b><br>"
                    "播放量：%{value:,.0f}<br>"
                    "占比：%{percent}<extra></extra>"
                ),
            )

            pie_fig.update_layout(
                height=520,
                margin=dict(l=20, r=20, t=20, b=40),
                showlegend=True,
                legend=dict(
                    orientation="h",
                    yanchor="top",
                    y=-0.12,
                    xanchor="center",
                    x=0.5,
                    title_text="",
                    font=dict(size=11),
                ),
            )

            st.plotly_chart(pie_fig, use_container_width=True)

    # 右侧：两个散点图上下排列
    with right:
        with st.container(border=True):
            st.markdown("#### 账号表现关系图")

            scatter1 = px.scatter(
                current,
                x=ACCOUNT_COLS["views"],
                y=ACCOUNT_COLS["new_followers"],
                color=ACCOUNT_COLS["account"],
                size="发帖量",
                hover_name=ACCOUNT_COLS["account"],
                title="播放量 vs 新增粉丝",
                labels={
                    ACCOUNT_COLS["views"]: "播放量",
                    ACCOUNT_COLS["new_followers"]: "新增粉丝"
                },
            )
            scatter1.update_traces(
                marker=dict(opacity=0.8),
                hovertemplate=(
                    "<b>%{hovertext}</b><br>"
                    "播放量：%{x:,.0f}<br>"
                    "新增粉丝：%{y:,.0f}<extra></extra>"
                ),
            )
            scatter1.update_layout(
                height=250,
                margin=dict(l=20, r=20, t=50, b=20),
                showlegend=False,
            )
            st.plotly_chart(scatter1, use_container_width=True)

            scatter2 = px.scatter(
                current,
                x=ACCOUNT_COLS["views"],
                y="条均",
                color=ACCOUNT_COLS["account"],
                size="发帖量",
                hover_name=ACCOUNT_COLS["account"],
                title="播放量 vs 条均",
                labels={
                    ACCOUNT_COLS["views"]: "播放量",
                    "条均": "条均播放"
                },
            )
            scatter2.update_traces(
                marker=dict(opacity=0.8),
                hovertemplate=(
                    "<b>%{hovertext}</b><br>"
                    "播放量：%{x:,.0f}<br>"
                    "条均播放：%{y:,.0f}<extra></extra>"
                ),
            )
            scatter2.update_layout(
                height=250,
                margin=dict(l=20, r=20, t=50, b=20),
                showlegend=False,
            )
            st.plotly_chart(scatter2, use_container_width=True)

    st.markdown("---")

    # =========================
    # 账号整体趋势变化
    # =========================
    st.markdown("### 账号整体变化趋势")
    st.caption("以下筛选器仅影响下方趋势图，用于查看不同账号在各周的播放量变化。")

    with st.container(border=True):
        all_accounts = sorted(account_df[ACCOUNT_COLS["account"]].dropna().unique().tolist())
        chosen_accounts = st.multiselect(
            "选择想查看趋势的账号",
            options=all_accounts,
            default=all_accounts,
            key="account_trend_selector",
        )

        trend = account_df[account_df[ACCOUNT_COLS["account"]].isin(chosen_accounts)].copy()
        trend = trend.sort_values("week_start")

        latest_5_weeks = trend["week_start"].drop_duplicates().sort_values().tail(5)
        trend = trend[trend["week_start"].isin(latest_5_weeks)].copy()
        trend = trend.sort_values("week_start")

        line_fig = px.line(
            trend,
            x=ACCOUNT_COLS["week"],
            y=ACCOUNT_COLS["views"],
            color=ACCOUNT_COLS["account"],
            markers=True,
            title="多账号播放量趋势(<=5周)",
        )

        line_fig.update_traces(
            line=dict(width=3),
            marker=dict(size=7),
            hovertemplate=(
                "<b>%{fullData.name}</b><br>"
                "周次：%{x}<br>"
                "播放量：%{y:,.0f}<extra></extra>"
            ),
        )

        line_fig.update_layout(
            height=460,
            margin=dict(l=20, r=20, t=50, b=20),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.5,
                xanchor="left",
                x=0,
                title_text="",
                font=dict(size=11),
            ),
            xaxis_title="周次",
            yaxis_title="播放量",
        )

        st.plotly_chart(line_fig, use_container_width=True)


def render_tab_videos(video_df: pd.DataFrame, weeks: List[str]):
    st.subheader("Tab3 视频维度")
    c1, c2 = st.columns([1, 2])
    with c1:
        selected_weeks = st.multiselect(
            "选择周次",
            options=weeks,
            default=[weeks[-1]] if weeks else [],
            key="video_week"
        )

    if selected_weeks:
        filtered_by_week = video_df[video_df[VIDEO_COLS["week"]].isin(selected_weeks)].copy()
    else:
        filtered_by_week = video_df.iloc[0:0].copy()  # 没选时返回空表

    account_options = sorted(filtered_by_week[VIDEO_COLS["account"]].dropna().unique().tolist())

    with c2:
        selected_accounts = st.multiselect(
            "选择账号（可多选）",
            options=account_options,
            default=account_options,
            key="video_accounts",
        )

    filtered = filtered_by_week[filtered_by_week[VIDEO_COLS["account"]].isin(selected_accounts)].copy()

    top10 = filtered.nlargest(10, VIDEO_COLS["views"]).copy()
    top10[VIDEO_COLS["video_id"]] = top10[VIDEO_COLS["video_id"]].astype(str)

    top10["hover_text"] = (
        "账号：" + top10[VIDEO_COLS["account"]].astype(str)
        + "<br>播放量：" + top10[VIDEO_COLS["views"]].fillna(0).astype(int).astype(str)
        + "<br>字幕/描述：" + top10[VIDEO_COLS["caption"]].fillna("").astype(str).str.slice(0, 200)
    )

    top10 = top10.sort_values(VIDEO_COLS["views"], ascending=True)

    bar_fig = px.bar(
        top10,
        x=VIDEO_COLS["views"],
        y=VIDEO_COLS["video_id"],
        orientation="h",
        color=VIDEO_COLS["account"],
        hover_name=VIDEO_COLS["account"],
        hover_data={VIDEO_COLS["caption"]: True, VIDEO_COLS["views"]: ":,.0f"},
        title="Top10 视频播放量",
        labels={VIDEO_COLS["video_id"]: "视频ID", VIDEO_COLS["views"]: "播放量"},
    )

    bar_fig.update_layout(
        height=460,
        yaxis=dict(
            categoryorder="total ascending",
            type="category"
        )
    )

    st.plotly_chart(bar_fig, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        bubble_fig = px.scatter(
            filtered,
            x=VIDEO_COLS["views"],
            y="互动率",
            size=VIDEO_COLS["favorites"],
            color=VIDEO_COLS["account"],
            hover_name=filtered[VIDEO_COLS["video_id"]].astype(str),
            hover_data={
                VIDEO_COLS["caption"]: True,
                VIDEO_COLS["favorites"]: ":,.0f",
                "互动率": ":.2%",
            },
            title="互动率气泡图",
            labels={VIDEO_COLS["views"]: "播放量", "互动率": "互动率"},
        )
        bubble_fig.update_layout(height=430)
        st.plotly_chart(bubble_fig, use_container_width=True)

    with c4:
        comp_fig = px.scatter(
            filtered,
            x=VIDEO_COLS["views"],
            y="完播率_decimal",
            color=VIDEO_COLS["account"],
            size=VIDEO_COLS["favorites"],
            hover_name=filtered[VIDEO_COLS["video_id"]].astype(str),
            hover_data={VIDEO_COLS["caption"]: True, "完播率_decimal": ":.2%"},
            title="完播率 vs 播放量",
            labels={VIDEO_COLS["views"]: "播放量", "完播率_decimal": "完播率"},
        )
        comp_fig.update_layout(height=430)
        st.plotly_chart(comp_fig, use_container_width=True)


    c5, c6 = st.columns(2)
    with c5:
        st.markdown("### 话题标签表格")
        hashtag_rows = []
        for _, row in filtered.iterrows():
            tags = extract_hashtags(row[VIDEO_COLS["caption"]])
            for tag in tags:
                hashtag_rows.append({
                    "话题标签": tag,
                    "播放量": row[VIDEO_COLS["views"]],
                    "发布量": 1,
                })
        if hashtag_rows:
            hashtag_df = pd.DataFrame(hashtag_rows).groupby("话题标签", as_index=False).agg({"播放量": "sum", "发布量": "sum"})
            hashtag_df["条均"] = hashtag_df["播放量"] / hashtag_df["发布量"]
            hashtag_df = hashtag_df.sort_values("播放量", ascending=False)
            st.dataframe(
                hashtag_df.style.format({"播放量": "{:,.0f}", "发布量": "{:,.0f}", "条均": "{:,.0f}"}),
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.info("当前筛选条件下未提取到 #话题标签。")

    with c6:
        st.markdown("### 各时长区间条均播放量")

        duration_bins = [-np.inf, 15, 30, 60, np.inf]
        duration_labels = ["0~15s", "15~30s", "30~60s", "60s+"]

        filtered["时长分桶"] = pd.cut(
            filtered[VIDEO_COLS["duration"]],
            bins=duration_bins,
            labels=duration_labels
        )

        duration_perf = (
            filtered.groupby("时长分桶", as_index=False)
            .agg(
                视频数=(VIDEO_COLS["video_id"], "count"),
                总播放量=(VIDEO_COLS["views"], "sum")
            )
        )

        duration_perf["条均播放量"] = np.where(
            duration_perf["视频数"] > 0,
            duration_perf["总播放量"] / duration_perf["视频数"],
            np.nan
        )

        hist_fig = px.bar(
            duration_perf,
            x="时长分桶",
            y="条均播放量",
            text="条均播放量",
            title="各时长区间条均播放量",
            labels={"时长分桶": "时长区间", "条均播放量": "条均播放量"},
        )

        hist_fig.update_traces(
            texttemplate="%{text:,.0f}",
            textposition="outside",
            hovertemplate=(
                "时长区间: %{x}<br>"
                "条均播放量: %{y:,.0f}<extra></extra>"
            )
        )

        y_max = duration_perf["条均播放量"].max()
        hist_fig.update_yaxes(
            range=[0, y_max * 1.2]
            )
        
        hist_fig.update_layout(height=450)
        st.plotly_chart(hist_fig, use_container_width=True)

    with st.expander("查看解析出标签的视频明细", expanded=False):
        show_cols = [
            VIDEO_COLS["video_id"], VIDEO_COLS["account"], VIDEO_COLS["views"], VIDEO_COLS["likes"],
            VIDEO_COLS["comments"], VIDEO_COLS["favorites"], VIDEO_COLS["completion_rate"], VIDEO_COLS["caption"]
        ]
        st.dataframe(filtered[show_cols], use_container_width=True, hide_index=True)


def main():
    st.title("周报数据可视化看板")
    st.caption("上传包含“视频维度数据”和“账号维度数据”两个工作表的 Excel 文件后，即可自动生成可视化看板。")

    uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx"])

    if uploaded_file is None:
        st.info("请先上传 Excel 文件。")
        st.stop()

    try:
        video_df = None
        account_df = None
        file_bytes = uploaded_file.read()
        video_df, account_df = load_workbook(file_bytes)
    except Exception as e:
        st.error(f"文件读取失败：{e}")
        return

    weeks = get_week_options(account_df)
    if not weeks:
        st.error("未识别到有效周次。请检查“日期”列格式是否为 `YYYY-MM-DD ~ YYYY-MM-DD`。")
        st.stop()

    tab1, tab2, tab3 = st.tabs(["整体概览", "账号维度", "视频维度"])

    with tab1:
        render_tab_overview(account_df, video_df, weeks)

    with tab2:
        render_tab_accounts(account_df, weeks)

    with tab3:
        render_tab_videos(video_df, weeks)


if __name__ == "__main__":
    main()
