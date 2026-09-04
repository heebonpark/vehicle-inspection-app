import streamlit as st

def apply_custom_styles():
    st.markdown("""
        <style>
        @import url("https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.8/dist/web/static/pretendard.css");

        /* 전역 폰트 및 배경 적용 */
        html, body, [class*="css"] {
            font-family: "Pretendard", -apple-system, BlinkMacSystemFont, system-ui, Roboto, "Helvetica Neue", "Segoe UI", "Apple SD Gothic Neo", "Noto Sans KR", "Malgun Gothic", "Apple Color Emoji", "Segoe UI Emoji", "Segoe UI Symbol", sans-serif;
        }

        /* 메인 배경색 - 그라디언트 or 깨끗한 화이트/그레이 */
        .stApp {
            background-color: #f8fafc;
        }

        /* Streamlit 기본 헤더/푸터 숨김 */
        #MainMenu {visibility: hidden;}
        header {visibility: hidden;}
        footer {visibility: hidden;}

        /* 타이틀 색상 (Dark Navy) */
        h1, h2, h3 {
            color: #0f172a !important;
        }

        /* 버튼 프리미엄 스타일링 (Royal Blue) */
        .stButton>button {
            background-color: #2563eb !important;
            color: white !important;
            border-radius: 12px !important;
            border: none !important;
            padding: 0.5rem 1rem !important;
            box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.2), 0 2px 4px -1px rgba(37, 99, 235, 0.1) !important;
            transition: all 0.2s ease-in-out !important;
            font-weight: 600 !important;
        }

        .stButton>button:hover {
            background-color: #1d4ed8 !important;
            box-shadow: 0 10px 15px -3px rgba(37, 99, 235, 0.3), 0 4px 6px -2px rgba(37, 99, 235, 0.1) !important;
            transform: translateY(-2px) !important;
        }

        /* 인풋창 라운드 처리 */
        .stTextInput>div>div>input {
            border-radius: 12px !important;
            border: 1px solid #cbd5e1 !important;
            padding: 0.7rem !important;
            transition: border-color 0.2s ease-in-out;
        }
        .stTextInput>div>div>input:focus {
            border-color: #2563eb !important;
            box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.2) !important;
        }

        /* SelectBox 등 컴포넌트 */
        [data-baseweb="select"] > div {
            border-radius: 12px !important;
        }

        /* Expander 스타일링 */
        .streamlit-expanderHeader {
            background-color: white !important;
            border-radius: 12px !important;
            box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1) !important;
            font-weight: 600 !important;
            color: #0f172a !important;
        }
        .streamlit-expanderContent {
            background-color: white !important;
            border-bottom-left-radius: 12px !important;
            border-bottom-right-radius: 12px !important;
            border: 1px solid #e2e8f0;
            border-top: none;
        }

        /* Dataframe */
        [data-testid="stDataFrame"] {
            border-radius: 12px !important;
            overflow: hidden !important;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1) !important;
        }

        /* 사이드바 스타일링 (Slate) */
        [data-testid="stSidebar"] {
            background-color: #ffffff !important;
            border-right: 1px solid #e2e8f0 !important;
        }
        </style>
    """, unsafe_allow_html=True)
