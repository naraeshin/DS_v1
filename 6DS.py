# app.py
# 요구 라이브러리: streamlit, pandas, matplotlib, openpyxl, matplotlib-font-manager (옵션)
# 실행:  streamlit run app.py
import io
import math
import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from matplotlib.ticker import PercentFormatter

# --------------------------
# 한글 폰트 설정 (가능한 경우)
# --------------------------
def set_korean_font():
    try:
        import matplotlib.font_manager as fm
        # 자주 쓰는 무료 폰트 후보
        candidates = ["NanumGothic", "AppleGothic", "Malgun Gothic", "Noto Sans CJK KR", "Noto Sans KR"]
        available = {f.name for f in fm.fontManager.ttflist}
        for name in candidates:
            if name in available:
                plt.rcParams["font.family"] = name
                break
        plt.rcParams["axes.unicode_minus"] = False
    except Exception:
        pass

set_korean_font()

st.set_page_config(page_title="엑셀 차트 대시보드", layout="wide")
st.title("엑셀 대시보드 📊")
st.caption("업로드한 파일의 각 시트를 적절한 차트로 한 화면에서 시각화합니다.")

# --------------------------
# 파일 업로드
# --------------------------
uploaded = st.file_uploader("엑셀 파일(.xlsx)를 업로드하세요", type=["xlsx"])

# 표시 옵션
with st.sidebar:
    st.header("표시 옵션")
    show_titles = st.checkbox("차트 제목 표시", value=True)
    tight_layout = st.checkbox("여백 자동 조정(tight_layout)", value=True)
    height = st.number_input("차트 높이(px)", min_value=250, max_value=900, value=360, step=10)
    alpha_bubble = st.slider("버블 투명도", 0.1, 1.0, 0.6, 0.05)
    st.divider()
    st.markdown("**시트 표시 선택**")

def make_figure(figsize=(6,4), title=None):
    fig, ax = plt.subplots(figsize=(figsize[0], figsize[1]))
    if title:
        ax.set_title(title)
    return fig, ax

def draw_bar_hist(df, title):
    # 1열: 범주, 이후: 값(들)
    x = df.columns[0]
    ycols = df.columns[1:]
    fig, ax = make_figure(figsize=(7, height/120), title=title if show_titles else None)
    df.plot(kind="bar", x=x, y=ycols, ax=ax, legend=True)
    ax.set_xlabel(x)
    ax.set_ylabel("값")
    if tight_layout: fig.tight_layout()
    st.pyplot(fig)

def draw_timeseries(df, title):
    x = df.columns[0]
    ycols = df.columns[1:]
    # x가 날짜처럼 보이면 파싱
    try:
        df[x] = pd.to_datetime(df[x])
    except Exception:
        pass
    fig, ax = make_figure(figsize=(7, height/120), title=title if show_titles else None)
    for col in ycols:
        ax.plot(df[x], df[col], label=col)
    ax.set_xlabel(x)
    ax.set_ylabel("값")
    ax.legend(loc="best")
    if tight_layout: fig.tight_layout()
    st.pyplot(fig)

def draw_pie(df, title):
    # 1열: 레이블, 2열: 값
    labels = df.columns[0]
    values = df.columns[1]
    fig = plt.figure(figsize=(6, height/120))
    if show_titles:
        plt.title(title)
    plt.pie(df[values], labels=df[labels], autopct="%1.1f%%", startangle=90, counterclock=False)
    plt.axis('equal')
    if tight_layout: fig.tight_layout()
    st.pyplot(fig)

def draw_scatter(df, title):
    # 1열: x, 2열: y
    x, y = df.columns[0], df.columns[1]
    fig, ax = make_figure(figsize=(6.5, height/120), title=title if show_titles else None)
    ax.scatter(df[x], df[y])
    ax.set_xlabel(x); ax.set_ylabel(y)
    if tight_layout: fig.tight_layout()
    st.pyplot(fig)

def draw_pareto(df, title):
    # 1열: 항목, 2열: 값
    label, value = df.columns[0], df.columns[1]
    s = df[[label, value]].dropna().sort_values(by=value, ascending=False)
    cum = s[value].cumsum()
    total = s[value].sum() if s[value].sum() != 0 else 1
    cum_pct = cum / total * 100

    fig, ax = make_figure(figsize=(7.5, height/120), title=title if show_titles else None)
    ax.bar(s[label], s[value])
    ax.set_ylabel("값")
    ax.set_xticklabels(s[label], rotation=45, ha='right')

    ax2 = ax.twinx()
    ax2.plot(range(len(s)), cum_pct.values, marker="o")
    ax2.yaxis.set_major_formatter(PercentFormatter())
    ax2.set_ylabel("누적 비율(%)")
    ax2.set_ylim(0, 110)
    # 80% 기준선
    ax2.axhline(80, linestyle="--")

    if tight_layout: fig.tight_layout()
    st.pyplot(fig)

def draw_bubble(df, title, alpha=0.6):
    # 1열: x, 2열: y, 3열: size
    x, y, size = df.columns[:3]
    sizes = df[size]
    # 버블 크기 스케일링(시각적으로 적당히)
    if sizes.max() > 0:
        s_scaled = (sizes - sizes.min()) / (sizes.max() - sizes.min() + 1e-9)
        s_scaled = 100 + 2000 * s_scaled
    else:
        s_scaled = np.full_like(sizes, 300)

    fig, ax = make_figure(figsize=(6.5, height/120), title=title if show_titles else None)
    scatter = ax.scatter(df[x], df[y], s=s_scaled, alpha=alpha)
    ax.set_xlabel(x); ax.set_ylabel(y)
    if tight_layout: fig.tight_layout()
    st.pyplot(fig)

def safe_sheet(df, needed_cols=2):
    return df is not None and isinstance(df, pd.DataFrame) and df.shape[1] >= needed_cols and df.dropna(how="all").shape[0] > 0

if uploaded:
    # 파일 읽기
    xls = pd.ExcelFile(uploaded)
    sheets = xls.sheet_names

    # 시트별 필터 체크박스
    desired_order = ["바차트_히스토그램", "시계열차트", "파이차트", "산점도", "파레토차트", "버블차트"]
    # 업로드 파일의 실제 시트 순서/이름을 유지하되 가독성 위해 정렬
    ordered = [s for s in desired_order if s in sheets] + [s for s in sheets if s not in desired_order]

    chosen = {}
    with st.sidebar:
        for s in ordered:
            chosen[s] = st.checkbox(s, value=True)

    # 그리드 배치: 2열 그리드
    def two_col_grid(items):
        # items: (sheet_name, df) 리스트
        rows = math.ceil(len(items) / 2)
        idx = 0
        for _ in range(rows):
            c1, c2 = st.columns(2, gap="large")
            for c in (c1, c2):
                if idx < len(items):
                    s, df = items[idx]
                    with c:
                        # 시트명에 따라 그리기
                        try:
                            if s == "바차트_히스토그램" and safe_sheet(df, 2):
                                draw_bar_hist(df, s)
                            elif s == "시계열차트" and safe_sheet(df, 2):
                                draw_timeseries(df, s)
                            elif s == "파이차트" and safe_sheet(df, 2):
                                draw_pie(df, s)
                            elif s == "산점도" and safe_sheet(df, 2):
                                draw_scatter(df, s)
                            elif s == "파레토차트" and safe_sheet(df, 2):
                                draw_pareto(df, s)
                            elif s == "버블차트" and safe_sheet(df, 3):
                                draw_bubble(df, s, alpha=alpha_bubble)
                            else:
                                st.info(f"'{s}' 시트의 데이터 형식이 예상과 달라 차트를 건너뜁니다.")
                        except Exception as e:
                            st.error(f"'{s}' 차트 생성 중 오류: {e}")
                    idx += 1

    # 데이터 읽고 그리기
    loaded = []
    for s in ordered:
        if chosen.get(s, False):
            try:
                df = pd.read_excel(xls, sheet_name=s)
            except Exception as e:
                st.error(f"시트 '{s}'를 읽는 중 오류: {e}")
                df = None
            loaded.append((s, df))

    two_col_grid(loaded)

    st.success("대시보드 렌더링 완료!")
    st.caption("필요 시 사이드바에서 표시/옵션을 조정하세요.")
else:
    st.info("좌측 상단 버튼으로 엑셀 파일(.xlsx)을 업로드하면 대시보드가 생성됩니다.")
