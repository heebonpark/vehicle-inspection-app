import streamlit as st
import json
import hashlib
import hmac
import os
import io
import sqlite3
from datetime import datetime
from PIL import Image
from streamlit_drawable_canvas import st_canvas

from app.core.db import init_db, get_db
from app.core.handlers import (
    CHECKLIST_SECTIONS, generate_integrated_excel, generate_integrated_pdf,
    generate_batch_excel, generate_batch_pdf, compress_image_bytes,
)
from app.ui.styles import apply_custom_styles

INSPECTION_COLUMNS = [
    "id", "created_at", "inspect_date", "inspector", "hq_name", "branch_name", "car_no",
    "check_data", "accumulated_km", "signature_name", "signature_image",
    "img_front", "img_rear", "img_right", "img_left",
]


def row_to_inspection_dict(rec):
    """SQLite는 ALTER TABLE로 추가된 컬럼을 항상 테이블 끝에 붙이기 때문에,
    SELECT * 로는 DB 생성 시점에 따라 컬럼 순서가 달라질 수 있다. 그래서 항상
    컬럼명을 명시해서 조회하고, 그 목록(INSPECTION_COLUMNS)과 짝지어 dict로
    변환한다."""
    return dict(zip(INSPECTION_COLUMNS, rec))


def _has_signature(canvas_result, min_dark_pixels=30):
    if canvas_result is None or canvas_result.image_data is None:
        return False
    arr = canvas_result.image_data
    return int((arr[:, :, 0] < 200).sum()) > min_dark_pixels


def _signature_png_bytes(canvas_result):
    arr = canvas_result.image_data.astype("uint8")
    img = Image.fromarray(arr, "RGBA")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

HQ_NAME = "강북/강원본부"
BRANCHES = ["중앙지사", "강북지사", "서대문지사", "고양지사", "의정부지사", "남양주지사", "강릉지사", "원주지사", "춘천고객지원팀"]
DIRECT_INPUT_LABEL = "-- 직접 입력 --"

# 블루투스이어폰 관련 2개 항목은 현재 지급 대상이 아니라 현장에서 입력받지 않고
# 항상 '해당없음'으로 고정한다 (보고서에는 회색 음영 + 대각선으로 계속 표시됨).
FORCED_NA_KEYS = {"item_etc_3", "item_etc_4"}


def force_rear_camera():
    """st.camera_input은 후면/전면 카메라를 지정하는 기능이 없어(전면이 기본으로
    열리는 문제) 대신 파일 업로더의 실제 <input type=file> 요소에 capture="environment"를
    주입한다. 이렇게 하면 모바일에서 사진 등록 버튼을 누를 때 스마트폰 기본 카메라 앱이
    후면 카메라로 바로 열린다."""
    st.iframe("""
        <script>
        (function() {
            function patchInputs() {
                try {
                    var doc = window.parent.document;
                    doc.querySelectorAll('input[type="file"]').forEach(function(el) {
                        if (el.getAttribute('capture') !== 'environment') {
                            el.setAttribute('capture', 'environment');
                        }
                    });
                } catch (e) {}
            }
            patchInputs();
            try {
                new MutationObserver(patchInputs).observe(window.parent.document.body, {childList: true, subtree: true});
            } catch (e) {}
        })();
        </script>
    """, height=1)

# 최초 1회(DB에 계정이 하나도 없을 때)만 자동 생성되는 기본 계정. 이후 비밀번호는
# DB(users 테이블)에 저장되며 로그인 후 화면에서 변경 가능하다.
DEFAULT_PASSWORD = "admin1234"
DEFAULT_ACCOUNTS = [
    {"id": "hq_admin", "name": "본부 총괄관리자", "role": "본부", "branch": "전체"},
    {"id": "jungang_mgr", "name": "중앙지사 관리자", "role": "지사", "branch": "중앙지사"},
    {"id": "gangbuk_mgr", "name": "강북지사 관리자", "role": "지사", "branch": "강북지사"},
    {"id": "seodaemun_mgr", "name": "서대문지사 관리자", "role": "지사", "branch": "서대문지사"},
    {"id": "goyang_mgr", "name": "고양지사 관리자", "role": "지사", "branch": "고양지사"},
    {"id": "uijeongbu_mgr", "name": "의정부지사 관리자", "role": "지사", "branch": "의정부지사"},
    {"id": "namyangju_mgr", "name": "남양주지사 관리자", "role": "지사", "branch": "남양주지사"},
    {"id": "gangneung_mgr", "name": "강릉지사 관리자", "role": "지사", "branch": "강릉지사"},
    {"id": "wonju_mgr", "name": "원주지사 관리자", "role": "지사", "branch": "원주지사"},
    {"id": "chuncheon_mgr", "name": "춘천고객지원팀 관리자", "role": "지사", "branch": "춘천고객지원팀"},
]

# --- 1. 페이지 설정 및 폰트/스타일 세팅 ---
st.set_page_config(page_title="이륜/업무용 차량 안전관리 종합시스템", layout="wide", initial_sidebar_state="expanded")
apply_custom_styles()

# DB 초기화
init_db()


def _hash_password(password: str, salt: bytes) -> str:
    return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 100_000).hex()


def seed_default_users():
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM users")
        if c.fetchone()[0] == 0:
            for acc in DEFAULT_ACCOUNTS:
                salt = os.urandom(16)
                conn.execute(
                    "INSERT INTO users (id, name, role, branch, salt, pw_hash) VALUES (?, ?, ?, ?, ?, ?)",
                    (acc["id"], acc["name"], acc["role"], acc["branch"], salt.hex(), _hash_password(DEFAULT_PASSWORD, salt))
                )


seed_default_users()


def list_users():
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id, name, role, branch FROM users ORDER BY (role != '본부'), branch")
        return c.fetchall()


def get_user(user_id: str):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id, name, role, branch, salt, pw_hash FROM users WHERE id = ?", (user_id,))
        row = c.fetchone()
    if not row:
        return None
    return {"id": row[0], "name": row[1], "role": row[2], "branch": row[3], "salt": row[4], "pw_hash": row[5]}


def verify_password(user_id: str, password: str) -> bool:
    user = get_user(user_id)
    if not user:
        return False
    computed = _hash_password(password, bytes.fromhex(user["salt"]))
    return hmac.compare_digest(computed, user["pw_hash"])


def set_password(user_id: str, new_password: str):
    salt = os.urandom(16)
    with get_db() as conn:
        conn.execute(
            "UPDATE users SET salt = ?, pw_hash = ? WHERE id = ?",
            (salt.hex(), _hash_password(new_password, salt), user_id)
        )


# --- 2. 로그인 및 권한 관리 ---
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
        login_accounts = list_users()
        login_id = st.selectbox(
            "계정 선택", options=[a[0] for a in login_accounts],
            format_func=lambda uid: next(f"{a[1]} ({a[3]})" for a in login_accounts if a[0] == uid),
            key="login_id_select"
        )
        lpw = st.text_input("비밀번호", type="password", key="login_pw_input")
        if st.button("로그인", use_container_width=True):
            if verify_password(login_id, lpw):
                user = get_user(login_id)
                user.pop("salt", None)
                user.pop("pw_hash", None)
                st.session_state.auth_user = user
                st.rerun()
            else:
                st.error("비밀번호가 일치하지 않습니다.")

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
        hq_in = st.text_input("본부명", value=HQ_NAME, disabled=True)
    with c3:
        branch_in = st.selectbox("지사명 선택", BRANCHES)
    with c4:
        with get_db() as conn:
            c = conn.cursor()
            c.execute("SELECT car_no FROM vehicles WHERE branch_name = ? ORDER BY car_no", (branch_in,))
            registered_cars = [r[0] for r in c.fetchall()]

        if registered_cars:
            car_choice = st.selectbox("차량번호 선택", registered_cars + [DIRECT_INPUT_LABEL], key="car_choice_select")
        else:
            car_choice = DIRECT_INPUT_LABEL

        if car_choice == DIRECT_INPUT_LABEL:
            car_in = st.text_input("차량번호 직접 입력", placeholder="경기 안양 아 7027", key="car_in_manual")
        else:
            car_in = car_choice

    # 누적 km는 차량번호 바로 아래(같은 열)에 배치해 가독성을 높인다.
    _km_c1, _km_c2, _km_c3, _km_c4 = st.columns([1.5, 2, 2, 2.5])
    with _km_c4:
        km_in = st.text_input("🛣️ 누적 km 수", placeholder="예: 24270", key="val_km")

    st.divider()

    # [파트 1: 이륜차량 점검표]
    st.subheader("1. 이륜차량 안전관리 점검표")
    st.caption("기본값은 '적정'으로 설정되어 있습니다. 이상이 있는 항목만 '정비필요'로, 해당 차량에 없는 항목(예: 블루투스이어폰 미지급)은 '해당없음'으로 변경하세요.")

    collected_checks = {}

    for sec in CHECKLIST_SECTIONS:
        with st.expander(f"{sec['category']}", expanded=True):
            for sub_cat, desc, key in sec["items"]:
                if key == "item_km":
                    col_k1, col_k2 = st.columns([3, 1])
                    with col_k1:
                        st.write(f"**{sub_cat}** : {desc}")
                    with col_k2:
                        st.write(f"km : {km_in}" if km_in else "km : (상단에 입력)")
                elif key in FORCED_NA_KEYS:
                    col_t, col_r = st.columns([3.2, 1.2])
                    with col_t:
                        st.write(f"**[{sub_cat}]** {desc}")
                    with col_r:
                        st.caption("해당없음 (등록 불필요)")
                    collected_checks[key] = "해당없음"
                else:
                    col_t, col_r = st.columns([3.2, 1.2])
                    with col_t:
                        st.write(f"**[{sub_cat}]** {desc}")
                    with col_r:
                        ans = st.radio("상태", ["적정", "정비필요", "해당없음"], horizontal=True, key=f"r_{key}", label_visibility="collapsed")
                        collected_checks[key] = ans

    st.divider()

    # [파트 2: 4면 사진 촬영]
    st.subheader("2. 기술/업무용 차량 4면 사진 등록")
    st.caption("사진 등록을 누르면 스마트폰 후면 카메라가 바로 열립니다. 기존 사진을 올리려면 '파일 선택'에서 갤러리를 선택하세요.")
    force_rear_camera()

    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown("**● 1. 전면 (Front)**")
        img_f = st.file_uploader("전면", type=["jpg", "jpeg", "png"], key="up_f")
        st.markdown("**● 3. 우측면 (Right)**")
        img_rt = st.file_uploader("우측면", type=["jpg", "jpeg", "png"], key="up_rt")
    with col_r:
        st.markdown("**● 2. 후면 (Rear)**")
        img_r = st.file_uploader("후면", type=["jpg", "jpeg", "png"], key="up_r")
        st.markdown("**● 4. 좌측면 (Left)**")
        img_lt = st.file_uploader("좌측면", type=["jpg", "jpeg", "png"], key="up_lt")

    st.divider()

    # [파트 3: 점검자 서명]
    st.subheader("3. 점검자 서명")
    st.caption("아래 칸에 손가락(모바일) 또는 마우스로 직접 서명해주세요. 서명은 보고서에 그대로 표시됩니다.")
    sig_col, _ = st.columns([1, 1])
    with sig_col:
        canvas_result = st_canvas(
            fill_color="rgba(255,255,255,0)",
            stroke_width=3,
            stroke_color="#000000",
            background_color="#FFFFFF",
            height=140,
            width=350,
            drawing_mode="freedraw",
            key="signature_pad",
        )

    st.divider()

    if st.button("점검표 및 사진 일괄 전송 완료", type="primary", use_container_width=True):
        km_clean = km_in.replace(",", "").strip() if km_in else ""
        if not inspector_in or not car_in:
            st.error("점검자 성명과 차량번호를 반드시 입력해주세요.")
        elif not (img_f and img_r and img_rt and img_lt):
            st.error("전면, 후면, 우측면, 좌측면 4장의 사진을 모두 등록해주세요.")
        elif not km_clean.isdigit():
            st.error("누적 km 수는 숫자로만 입력해주세요. (예: 24270)")
        elif not _has_signature(canvas_result):
            st.error("점검자 서명란에 서명을 해주세요.")
        else:
            sig_bytes = _signature_png_bytes(canvas_result)
            with get_db() as conn:
                conn.execute("""
                    INSERT INTO integrated_inspections (
                        created_at, inspect_date, inspector, hq_name, branch_name, car_no,
                        check_data, accumulated_km, signature_name, signature_image,
                        img_front, img_rear, img_right, img_left
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    datetime.now().strftime("%y. %m. %d"),
                    inspector_in, hq_in, branch_in, car_in,
                    json.dumps(collected_checks, ensure_ascii=False),
                    km_clean, inspector_in, sig_bytes,
                    compress_image_bytes(img_f.getvalue()), compress_image_bytes(img_r.getvalue()),
                    compress_image_bytes(img_rt.getvalue()), compress_image_bytes(img_lt.getvalue())
                ))
            st.success(f"[{car_in}] 차량 점검 데이터와 4면 사진이 정상 등록되었습니다! 관리자 화면에서 즉시 다운로드 가능합니다.")

# --- 4. 화면 2: 관리자 종합 조회 및 2종 세트 출력 ---
elif active_menu == "관리자 종합 조회/출력":
    u = st.session_state.auth_user
    st.title(f"차량 안전관리 상태 종합 대장 ({u['role']}: {u['name']})")

    with st.expander("🔑 비밀번호 변경", expanded=False):
        st.caption(f"현재 계정: {u['name']} ({u['id']})")
        pc1, pc2, pc3 = st.columns(3)
        with pc1:
            pw_cur = st.text_input("현재 비밀번호", type="password", key="pw_cur")
        with pc2:
            pw_new = st.text_input("새 비밀번호", type="password", key="pw_new")
        with pc3:
            pw_new2 = st.text_input("새 비밀번호 확인", type="password", key="pw_new2")
        if st.button("비밀번호 변경", key="pw_change_btn"):
            if not verify_password(u["id"], pw_cur):
                st.error("현재 비밀번호가 일치하지 않습니다.")
            elif len(pw_new) < 4:
                st.error("새 비밀번호는 4자 이상이어야 합니다.")
            elif pw_new != pw_new2:
                st.error("새 비밀번호가 서로 일치하지 않습니다.")
            else:
                set_password(u["id"], pw_new)
                st.success("비밀번호가 변경되었습니다. 다음 로그인부터 새 비밀번호를 사용하세요.")

        if u["role"] == "본부":
            st.divider()
            st.caption("본부 관리자는 다른 계정의 비밀번호를 초기화할 수 있습니다.")
            all_accounts = list_users()
            reset_targets = [a for a in all_accounts if a[0] != u["id"]]
            rc1, rc2, rc3 = st.columns(3)
            with rc1:
                reset_target_id = st.selectbox(
                    "대상 계정", options=[a[0] for a in reset_targets],
                    format_func=lambda uid: next(f"{a[1]} ({a[3]})" for a in reset_targets if a[0] == uid),
                    key="pw_reset_target"
                )
            with rc2:
                pw_reset_new = st.text_input("새 비밀번호 지정", type="password", key="pw_reset_new")
            with rc3:
                st.write("")
                st.write("")
                if st.button("초기화", key="pw_reset_btn", use_container_width=True):
                    if len(pw_reset_new) < 4:
                        st.error("새 비밀번호는 4자 이상이어야 합니다.")
                    else:
                        set_password(reset_target_id, pw_reset_new)
                        st.success("초기화 완료.")

    with st.expander("🚗 차량 등록 관리 (현장 등록 화면 드롭다운 목록)", expanded=False):
        veh_branch_options = BRANCHES if u["role"] == "본부" else [u["branch"]]

        vc1, vc2, vc3 = st.columns([2, 3, 1])
        with vc1:
            veh_branch_in = st.selectbox(
                "지사", veh_branch_options, key="veh_branch_in",
                disabled=(len(veh_branch_options) == 1)
            )
        with vc2:
            veh_car_no_in = st.text_input("등록할 차량번호", placeholder="경기 안양 아 7027", key="veh_car_no_in")
        with vc3:
            st.write("")
            st.write("")
            if st.button("등록", use_container_width=True, key="veh_register_btn"):
                if not veh_car_no_in.strip():
                    st.error("차량번호를 입력해주세요.")
                else:
                    try:
                        with get_db() as conn:
                            conn.execute(
                                "INSERT INTO vehicles (created_at, hq_name, branch_name, car_no) VALUES (?, ?, ?, ?)",
                                (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), HQ_NAME, veh_branch_in, veh_car_no_in.strip())
                            )
                        st.success(f"[{veh_branch_in}] {veh_car_no_in.strip()} 등록 완료.")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("이미 등록된 차량번호입니다.")

        with get_db() as conn:
            c = conn.cursor()
            if u["role"] == "본부":
                c.execute("SELECT id, branch_name, car_no FROM vehicles ORDER BY branch_name, car_no")
            else:
                c.execute("SELECT id, branch_name, car_no FROM vehicles WHERE branch_name = ? ORDER BY car_no", (u["branch"],))
            veh_rows = c.fetchall()

        if veh_rows:
            st.caption(f"등록된 차량 {len(veh_rows)}건 (현장 등록 화면에서 지사 선택 시 드롭다운으로 표시됩니다)")
            for vid, vbranch, vcarno in veh_rows:
                dcol1, dcol2 = st.columns([5, 1])
                with dcol1:
                    st.write(f"[{vbranch}] {vcarno}")
                with dcol2:
                    if st.button("삭제", key=f"veh_del_{vid}", use_container_width=True):
                        with get_db() as conn:
                            conn.execute("DELETE FROM vehicles WHERE id = ?", (vid,))
                        st.rerun()
        else:
            st.caption("등록된 차량이 없습니다. 등록 전까지 현장 등록 화면은 차량번호 직접 입력으로 동작합니다.")

    with get_db() as conn:
        c = conn.cursor()
        if u["role"] == "본부":
            sel_branch = st.selectbox("지사 필터", ["전체"] + BRANCHES)
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
        st.caption("행을 클릭해 여러 건을 선택하면 아래에 일괄 다운로드 옵션이 나타납니다.")
        select_event = st.dataframe(
            [{"번호": r[0], "등록일시": r[1], "점검일자": r[2], "점검자": r[3], "지사": r[5], "차량번호": r[6], "누적km": r[7]} for r in rows],
            use_container_width=True,
            on_select="rerun",
            selection_mode="multi-row",
            key="admin_inspection_table",
        )

        selected_positions = list(select_event.selection.rows) if select_event and select_event.selection else []
        selected_ids = [rows[i][0] for i in selected_positions]

        if selected_ids:
            with get_db() as conn:
                c = conn.cursor()
                placeholders = ",".join("?" * len(selected_ids))
                c.execute(
                    f"SELECT {', '.join(INSPECTION_COLUMNS)} FROM integrated_inspections WHERE id IN ({placeholders})",
                    selected_ids
                )
                batch_by_id = {r[0]: row_to_inspection_dict(r) for r in c.fetchall()}
            # 화면에 표시된 순서(선택한 순서가 아니라 목록 순서)대로 정렬
            batch_recs = [batch_by_id[i] for i in selected_ids if i in batch_by_id]

            st.success(f"✅ {len(batch_recs)}건 선택됨 — 아래에서 한 번에 내려받을 수 있습니다.")
            bb1, bb2 = st.columns(2)
            with bb1:
                batch_xlsx = generate_batch_excel(batch_recs)
                st.download_button(
                    label=f"📊 선택 {len(batch_recs)}건 일괄 엑셀 다운로드 (차량별 시트 분리)",
                    data=batch_xlsx,
                    file_name=f"일괄점검_{len(batch_recs)}건.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="batch_xlsx_dl",
                )
            with bb2:
                batch_pdf = generate_batch_pdf(batch_recs)
                st.download_button(
                    label=f"📄 선택 {len(batch_recs)}건 일괄 PDF 다운로드 (차량별로 이어붙임)",
                    data=batch_pdf,
                    file_name=f"일괄점검_{len(batch_recs)}건.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    key="batch_pdf_dl",
                )

        st.divider()
        st.subheader("2세트 원본 양식 보고서 다운로드 (건별)")

        target_id = st.selectbox(
            "출력할 차량 선택",
            options=[r[0] for r in rows],
            format_func=lambda x: f"[{next(r[5] for r in rows if r[0] == x)}] {next(r[6] for r in rows if r[0] == x)} (점검자: {next(r[3] for r in rows if r[0] == x)})"
        )

        with get_db() as conn:
            c = conn.cursor()
            c.execute(f"SELECT {', '.join(INSPECTION_COLUMNS)} FROM integrated_inspections WHERE id = ?", (target_id,))
            rec = c.fetchone()

        if rec:
            rec_data = row_to_inspection_dict(rec)

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
