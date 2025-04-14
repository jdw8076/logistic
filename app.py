import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# 한글 폰트 설정
font_path = 'C:/Windows/Fonts/malgun.ttf'  # 말랑한 Malgun Gothic (윈도우 기본)
font_name = fm.FontProperties(fname=font_path).get_name()
plt.rc('font', family=font_name)

# 마이너한 깨짐 방지
plt.rcParams['axes.unicode_minus'] = False


# Load and preprocess data
@st.cache_data
def load_data():
    df = pd.read_excel("data/수출입물류주요업무조사표-2022Ver.xlsx", skiprows=5)
    df = df[[
        '名前', '主要業務', 'イッシュー及び主要事項', '業務進行手続き',
        '현재목표', '지향하는 목표', '報告(対応)周期', '被報告者',
        '備考', '工数確認(hr/月)', '業務対応可能者', 'Memo']]
    df.columns = [
        "이름", "주요업무", "이슈및주요사항", "업무진행절차",
        "현재목표", "지향목표", "보고주기", "피보고자",
        "비고", "공수(hr/月)", "업무대응가능자", "메모"
    ]
    df = df.dropna(subset=["이름", "주요업무"])
    return df

# Sidebar filters
def filter_data(df):
    names = st.sidebar.multiselect("담당자", sorted(df["이름"].unique()))
    duties = st.sidebar.multiselect("업무유형", sorted(df["주요업무"].unique()))
    cycles = st.sidebar.multiselect("보고주기", sorted(df["보고주기"].dropna().unique()))
    responders = st.sidebar.text_input("업무대응가능자 포함 검색어")

    if names:
        df = df[df["이름"].isin(names)]
    if duties:
        df = df[df["주요업무"].isin(duties)]
    if cycles:
        df = df[df["보고주기"].isin(cycles)]
    if responders:
        df = df[df["업무대응가능자"].str.contains(responders, na=False)]
    return df

def plot_charts(df):
    import matplotlib.pyplot as plt

    st.subheader("📊 담당자별 업무 건수")
    fig1, ax1 = plt.subplots()
    df["이름"].value_counts().plot(kind="bar", ax=ax1)
    st.pyplot(fig1)

    st.subheader("🍰 업무유형별 공수 비중")
    fig2, ax2 = plt.subplots()
    pie_data = df.groupby("주요업무")["공수(hr/月)"].sum().dropna()
    if not pie_data.empty:
        pie_data.plot(kind="pie", ax=ax2, autopct='%1.1f%%')
        ax2.set_ylabel("")
        st.pyplot(fig2)
    else:
        st.warning("⚠️ 공수 데이터가 없습니다.")

    st.subheader("📈 보고주기별 분포")
    fig3, ax3 = plt.subplots()
    df["보고주기"].value_counts().plot(kind="bar", ax=ax3)
    st.pyplot(fig3)


from io import BytesIO

# Main app
def main():
    st.title("📦 물류 업무 분장 대시보드")
    df = load_data()
    filtered_df = filter_data(df)

    st.subheader("📋 필터링된 업무 목록")
    st.dataframe(filtered_df)

    plot_charts(filtered_df)

    # Download report
    st.subheader("📁 리포트 다운로드")
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        filtered_df.to_excel(writer, index=False, sheet_name='Report')
    st.download_button(
        label="📥 엑셀로 다운로드",
        data=output.getvalue(),
        file_name="filtered_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    main()
