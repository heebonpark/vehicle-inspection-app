import streamlit as st
import json
import hashlib
import hmac
from datetime import datetime

from app.core.db import init_db, get_db
from app.core.handlers import CHECKLIST_SECTIONS, generate_integrated_excel, generate_integrated_pdf, compress_image_bytes
from app.ui.styles import apply_custom_styles

# --- 1. 페이지 설정 및 폰트/스타일 세팅 ---
st.set_page_config(page_title="이륜/업무용 차량 안전관리 종합시스템", layout="wide", initial_sidebar_state="expanded")
apply_custom_styles()

# DB 초기화
init_db()

# --- 2. 로그인 및 권한 관리 ---
# 비밀번호는 평문으로 저장하지 않고 PBKDF2-SHA256(salt+hash)로 저장한다.
USERS = {
    "hq_admin": {
        "salt": "158c58196d1a8f03011df3ce8969954b",
        "pw_hash": "0678fc7d58594e0598b3ef0d577c4711c638e1ae123f0fbfc58d02cc181a27bf",
        "name": "본부 총괄관리자", "role": "본부", "branch": "전체",
    },
    "jungang_mgr": {
        "salt": "cd6df512e7e9c0c8bd9879e1adb1da89",
        "pw_hash": "508bbd55196dddc5b143b30138a0b1046ce8473761b60e9b8f3a47bec68e93b3",
        "name": "중앙지사 관리자", "role": "지사", "branch": "중앙지사",
    },
    "gangbuk_mgr": {
        "salt": "eecba65f5c92937d350e16ff69e7350d",
        "pw_hash": "b46d21fcbfee363e6067e00266588d7d142121fd838d45ce4493915a9f9267b7",
        "name": "강북지사 관리자", "role": "지사", "branch": "강북지사",
    },
    "seodaemun_mgr": {
        "salt": "f28f9d2db8c8731f56e21afb2aaa8a46",
        "pw_hash": "a16c4ca30c6bd802ff6234ac12c4e2a25bd73d973a6a33a15fd46a7fc79469fd",
        "name": "서대문지사 관리자", "role": "지사", "branch": "서대문지사",
    },
}


def verify_password(user_id: str, password: str) -> bool:
    user = USERS.get(user_id)
    if not user:
        return False
    salt = bytes.fromhex(user["salt"])
    computed = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 100_000).hex()
    return hmac.compare_digest(computed, user["pw_hash"])

if "auth_user" not in st.session_state:
    st.session_state.auth_user = None

is_upload_mode = st.query_params.get("mode") == "upload"

st.sidebar.title("이륜/업무용 차량 점검")

if st.session_state.auth_user:
    u = st.session_state.auth_user
    st.sidebar.success(f"{u['name']} 접속 중 ({u['role']})")
    if st.sidebar.button("로그아웃", use_container_width=True):
        st.session_state.auth_user = None
        st.rerun()
else:
    with st.sidebar.expander("관리자 로그인", expanded=not is_upload_mode):
        lid = st.text_input("아이디")
        lpw = st.text_input("비밀번호", type="password")
        if st.button("로그인", use_container_width=True):
            if verify_password(lid, lpw):
                st.session_state.auth_user = USERS[lid]
                st.rerun()
            else:
                st.error("계정 정보를 확인해 주세요.")

menu = ["현장 점검 등록 (체크리스트 + 4면촬영)"]
if st.session_state.auth_user:
    menu.append("관리자 종합 조회/출력")

active_menu = st.sidebar.radio("메뉴", menu, index=1 if (st.session_state.auth_user and not is_upload_mode) else 0)

# --- 3. 화면 1: 현장 등록 (체크리스트 + 4면 사진) ---
if active_menu == "현장 점검 등록 (체크리스트 + 4면촬영)":
    st.title("이륜차량 안전관리 상태 평가 & 4면 사진 등록")
    st.caption("현장에서 점검표 항목을 체크하고, 4면 사진을 촬영하여 전송하세요.")

    c1, c2, c3, c4 = st.columns([1.5, 2, 2, 2.5])
    with c1:
        inspector_in = st.text_input("점검자 성명", placeholder="장부환")
    with c2:
        hq_in = st.text_input("본부명", value="강북/강원본부", disabled=True)
    with c3:
        branch_in = st.selectbox("지사명 선택", ["중앙지사", "강북지사", "서대문지사", "동대문지사", "기타지사"])
    with c4:
        car_in = st.text_input("차량번호", placeholder="경기 안양 아 7027")

    st.divider()

    # [파트 1: 이륜차량 점검표]
    st.subheader("1. 이륜차량 안전관리 점검표")
    st.caption("기본값은 '적정'으로 설정되어 있습니다. 이상이 있는 항목만 '정비필요'로 변경하세요.")

    collected_checks = {}

    for sec in CHECKLIST_SECTIONS:
        with st.expander(f"{sec['category']}", expanded=True):
            for sub_cat, desc, key in sec["items"]:
                if key == "item_km":
                    col_k1, col_k2 = st.columns([3, 1])
                    with col_k1:
                        st.write(f"**{sub_cat}** : {desc}")
                    with col_k2:
                        km_in = st.text_input("누적 km 수 입력", placeholder="예: 24270", key="val_km")
                else:
                    col_t, col_r = st.columns([3.2, 1.2])
                    with col_t:
                        st.write(f"**[{sub_cat}]** {desc}")
                    with col_r:
                        ans = st.radio("상태", ["적정", "정비필요"], horizontal=True, key=f"r_{key}", label_visibility="collapsed")
                        collected_checks[key] = ans

    st.divider()

    # [파트 2: 4면 사진 촬영]
    st.subheader("2. 기술/업무용 차량 4면 사진 등록")
    input_method = st.radio("촬영 방식", ["스마트폰 카메라 촬영", "사진 보관함 업로드"], horizontal=True)

    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown("**● 1. 전면 (Front)**")
        img_f = st.camera_input("전면", key="cam_f") if "카메라" in input_method else st.file_uploader("전면", type=["jpg", "jpeg", "png"], key="up_f")
        st.markdown("**● 3. 우측면 (Right)**")
        img_rt = st.camera_input("우측면", key="cam_rt") if "카메라" in input_method else st.file_uploader("우측면", type=["jpg", "jpeg", "png"], key="up_rt")
    with col_r:
        st.markdown("**● 2. 후면 (Rear)**")
        img_r = st.camera_input("후면", key="cam_r") if "카메라" in input_method else st.file_uploader("후면", type=["jpg", "jpeg", "png"], key="up_r")
        st.markdown("**● 4. 좌측면 (Left)**")
        img_lt = st.camera_input("좌측면", key="cam_lt") if "카메라" in input_method else st.file_uploader("좌측면", type=["jpg", "jpeg", "png"], key="up_lt")

    st.divider()

    if st.button("점검표 및 사진 일괄 전송 완료", type="primary", use_container_width=True):
        km_clean = km_in.replace(",", "").strip() if km_in else ""
        if not inspector_in or not car_in:
            st.error("점검자 성명과 차량번호를 반드시 입력해주세요.")
        elif not (img_f and img_r and img_rt and img_lt):
            st.error("전면, 후면, 우측면, 좌측면 4장의 사진을 모두 등록해주세요.")
        elif not km_clean.isdigit():
            st.error("누적 km 수는 숫자로만 입력해주세요. (예: 24270)")
        else:
            with get_db() as conn:
                conn.execute("""
                    INSERT INTO integrated_inspections (
                        created_at, inspect_date, inspector, hq_name, branch_name, car_no,
                        check_data, accumulated_km, signature_name,
                        img_front, img_rear, img_right, img_left
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    datetime.now().strftime("%y. %m. %d"),
                    inspector_in, hq_in, branch_in, car_in,
                    json.dumps(collected_checks, ensure_ascii=False),
                    km_clean, inspector_in,
                    compress_image_bytes(img_f.getvalue()), compress_image_bytes(img_r.getvalue()),
                    compress_image_bytes(img_rt.getvalue()), compress_image_bytes(img_lt.getvalue())
                ))
            st.success(f"[{car_in}] 차량 점검 데이터와 4면 사진이 정상 등록되었습니다! 관리자 화면에서 즉시 다운로드 가능합니다.")

# --- 4. 화면 2: 관리자 종합 조회 및 2종 세트 출력 ---
elif active_menu == "관리자 종합 조회/출력":
    u = st.session_state.auth_user
    st.title(f"차량 안전관리 상태 종합 대장 ({u['role']}: {u['name']})")

    with get_db() as conn:
        c = conn.cursor()
        if u["role"] == "본부":
            sel_branch = st.selectbox("지사 필터", ["전체", "중앙지사", "강북지사", "서대문지사", "동대문지사", "기타지사"])
            if sel_branch != "전체":
                c.execute("SELECT id, created_at, inspect_date, inspector, hq_name, branch_name, car_no, accumulated_km FROM integrated_inspections WHERE branch_name = ? ORDER BY id DESC", (sel_branch,))
            else:
                c.execute("SELECT id, created_at, inspect_date, inspector, hq_name, branch_name, car_no, accumulated_km FROM integrated_inspections ORDER BY id DESC")
        else:
            sel_branch = u["branch"]
            st.info(f"소속 지사: **[{sel_branch}]** 등록 건만 조회됩니다.")
            c.execute("SELECT id, created_at, inspect_date, inspector, hq_name, branch_name, car_no, accumulated_km FROM integrated_inspections WHERE branch_name = ? ORDER BY id DESC", (sel_branch,))
        rows = c.fetchall()

    if not rows:
        st.warning("등록된 점검 내역이 없습니다.")
    else:
        st.dataframe(
            [{"번호": r[0], "등록일시": r[1], "점검일자": r[2], "점검자": r[3], "지사": r[5], "차량번호": r[6], "누적km": r[7]} for r in rows],
            use_container_width=True
        )

        st.divider()
        st.subheader("2세트 원본 양식 보고서 다운로드")

        target_id = st.selectbox(
            "출력할 차량 선택",
            options=[r[0] for r in rows],
            format_func=lambda x: f"[{next(r[5] for r in rows if r[0] == x)}] {next(r[6] for r in rows if r[0] == x)} (점검자: {next(r[3] for r in rows if r[0] == x)})"
        )

        with get_db() as conn:
            c = conn.cursor()
            c.execute("SELECT * FROM integrated_inspections WHERE id = ?", (target_id,))
            rec = c.fetchone()

        if rec:
            rec_data = {
                "id": rec[0], "created_at": rec[1], "inspect_date": rec[2], "inspector": rec[3],
                "hq_name": rec[4], "branch_name": rec[5], "car_no": rec[6],
                "check_data": rec[7], "accumulated_km": rec[8], "signature_name": rec[9],
                "img_front": rec[10], "img_rear": rec[11], "img_right": rec[12], "img_left": rec[13]
            }

            d_col1, d_col2 = st.columns(2)
            with d_col1:
                excel_bytes = generate_integrated_excel(rec_data)
                st.download_button(
                    label="📊 2세트 통합 엑셀 다운로드 (2개 탭 분리)",
                    data=excel_bytes,
                    file_name=f"차량안전점검2세트_{rec_data['car_no'].replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with d_col2:
                pdf_bytes = generate_integrated_pdf(rec_data)
                st.download_button(
                    label="📄 2세트 통합 PDF 다운로드 (1P점검표 + 2P사진)",
                    data=pdf_bytes,
                    file_name=f"차량안전점검2세트_{rec_data['car_no'].replace(' ', '_')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )

            # 웹 화면 미리보기
            with st.expander("현장 등록 사진 미리보기", expanded=True):
                p_c1, p_c2 = st.columns(2)
                with p_c1:
                    st.caption("● 전면")
                    if rec_data["img_front"]: st.image(rec_data["img_front"], use_container_width=True)
                    st.caption("● 우측면")
                    if rec_data["img_right"]: st.image(rec_data["img_right"], use_container_width=True)
                with p_c2:
                    st.caption("● 후면")
                    if rec_data["img_rear"]: st.image(rec_data["img_rear"], use_container_width=True)
                    st.caption("● 좌측면")
                    if rec_data["img_left"]: st.image(rec_data["img_left"], use_container_width=True)
