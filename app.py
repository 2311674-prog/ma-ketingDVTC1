import streamlit as st
import sqlite3
import pandas as pd
from io import BytesIO
from datetime import datetime


# ============================================================
# CẤU HÌNH ỨNG DỤNG
# ============================================================

st.set_page_config(
    page_title="Quản lý khách hàng",
    page_icon="👥",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# CSS GIAO DIỆN
# ============================================================

st.markdown(
    """
    <style>

    /* Toàn bộ trang */
    .main {
        padding-top: 1rem;
    }

    /* Tiêu đề */
    .main-title {
        font-size: 36px;
        font-weight: 700;
        color: #1f2937;
        margin-bottom: 5px;
    }

    .sub-title {
        font-size: 16px;
        color: #6b7280;
        margin-bottom: 25px;
    }

    /* Card thống kê */
    .stat-card {
        background: white;
        padding: 20px;
        border-radius: 12px;
        border: 1px solid #e5e7eb;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        text-align: center;
    }

    .stat-number {
        font-size: 28px;
        font-weight: bold;
        color: #2563eb;
    }

    .stat-label {
        font-size: 14px;
        color: #6b7280;
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background-color: #f8fafc;
    }

    /* Button */
    .stButton button {
        border-radius: 8px;
        font-weight: 600;
    }

    /* Dataframe */
    [data-testid="stDataFrame"] {
        border-radius: 10px;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# DATABASE
# ============================================================

DB_FILE = "customers.db"


def get_connection():
    """
    Tạo kết nối tới SQLite.
    """
    conn = sqlite3.connect(DB_FILE)
    return conn


def init_database():
    """
    Tạo bảng customers nếu chưa tồn tại.
    """

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS customers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            phone TEXT NOT NULL,
            name TEXT NOT NULL,
            area TEXT,
            note TEXT,
            created_at TEXT NOT NULL
        )
        """
    )

    conn.commit()
    conn.close()


# Khởi tạo database
init_database()


# ============================================================
# CÁC HÀM XỬ LÝ DATABASE
# ============================================================

def add_customer(phone, name, area, note):
    """
    Thêm khách hàng mới.
    """

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        """
        INSERT INTO customers
        (phone, name, area, note, created_at)
        VALUES (?, ?, ?, ?, ?)
        """,
        (
            phone,
            name,
            area,
            note,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        )
    )

    conn.commit()
    conn.close()


def get_all_customers():
    """
    Lấy toàn bộ khách hàng.
    """

    conn = get_connection()

    query = """
        SELECT
            id,
            phone,
            name,
            area,
            note,
            created_at
        FROM customers
        ORDER BY id DESC
    """

    df = pd.read_sql_query(query, conn)

    conn.close()

    return df


def update_customer(customer_id, phone, name, area, note):
    """
    Cập nhật thông tin khách hàng.
    """

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        """
        UPDATE customers
        SET
            phone = ?,
            name = ?,
            area = ?,
            note = ?
        WHERE id = ?
        """,
        (
            phone,
            name,
            area,
            note,
            customer_id
        )
    )

    conn.commit()
    conn.close()


def delete_customer(customer_id):
    """
    Xóa khách hàng.
    """

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        """
        DELETE FROM customers
        WHERE id = ?
        """,
        (customer_id,)
    )

    conn.commit()
    conn.close()


def delete_all_customers():
    """
    Xóa toàn bộ khách hàng.
    """

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("DELETE FROM customers")

    conn.commit()
    conn.close()


def get_customer_by_id(customer_id):
    """
    Lấy một khách hàng theo ID.
    """

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        """
        SELECT
            id,
            phone,
            name,
            area,
            note,
            created_at
        FROM customers
        WHERE id = ?
        """,
        (customer_id,)
    )

    customer = cursor.fetchone()

    conn.close()

    return customer


# ============================================================
# HÀM KIỂM TRA DỮ LIỆU
# ============================================================

def validate_phone(phone):
    """
    Kiểm tra số điện thoại.
    """

    if not phone:
        return False

    phone = phone.strip()

    # Cho phép số điện thoại có khoảng trắng, +, -
    allowed = "0123456789 +-()."

    for char in phone:
        if char not in allowed:
            return False

    return True


def phone_exists(phone, exclude_id=None):
    """
    Kiểm tra số điện thoại đã tồn tại hay chưa.
    """

    conn = get_connection()
    cursor = conn.cursor()

    if exclude_id is None:

        cursor.execute(
            """
            SELECT id
            FROM customers
            WHERE phone = ?
            """,
            (phone,)
        )

    else:

        cursor.execute(
            """
            SELECT id
            FROM customers
            WHERE phone = ?
            AND id != ?
            """,
            (
                phone,
                exclude_id
            )
        )

    result = cursor.fetchone()

    conn.close()

    return result is not None


# ============================================================
# XUẤT EXCEL
# ============================================================

def dataframe_to_excel(df):
    """
    Chuyển DataFrame thành file Excel trong bộ nhớ.
    """

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df_export = df.copy()

        # Đổi tên cột sang tiếng Việt
        df_export = df_export.rename(
            columns={
                "id": "ID",
                "phone": "Số điện thoại",
                "name": "Tên khách hàng",
                "area": "Khu vực",
                "note": "Ghi chú",
                "created_at": "Ngày tạo"
            }
        )

        df_export.to_excel(
            writer,
            index=False,
            sheet_name="Khách hàng"
        )

    return output.getvalue()


# ============================================================
# LẤY DỮ LIỆU
# ============================================================

df = get_all_customers()


# ============================================================
# HEADER
# ============================================================

st.markdown(
    '<div class="main-title">👥 Quản lý khách hàng</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="sub-title">'
    'Quản lý thông tin khách hàng nhanh chóng và thuận tiện'
    '</div>',
    unsafe_allow_html=True
)


# ============================================================
# THỐNG KÊ
# ============================================================

total_customers = len(df)

if total_customers > 0:

    total_areas = df["area"].fillna("").replace("", pd.NA).nunique()

    total_phones = df["phone"].nunique()

else:

    total_areas = 0
    total_phones = 0


col1, col2, col3 = st.columns(3)


with col1:

    st.markdown(
        f"""
        <div class="stat-card">
            <div class="stat-number">{total_customers}</div>
            <div class="stat-label">Tổng khách hàng</div>
        </div>
        """,
        unsafe_allow_html=True
    )


with col2:

    st.markdown(
        f"""
        <div class="stat-card">
            <div class="stat-number">{total_phones}</div>
            <div class="stat-label">Số điện thoại</div>
        </div>
        """,
        unsafe_allow_html=True
    )


with col3:

    st.markdown(
        f"""
        <div class="stat-card">
            <div class="stat-number">{total_areas}</div>
            <div class="stat-label">Khu vực</div>
        </div>
        """,
        unsafe_allow_html=True
    )


st.write("")


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header("⚙️ Chức năng")

    menu = st.radio(
        "Chọn chức năng",
        [
            "➕ Thêm khách hàng",
            "📋 Danh sách khách hàng",
            "✏️ Chỉnh sửa khách hàng",
            "🗑️ Xóa khách hàng"
        ]
    )

    st.divider()

    st.info(
        "💡 Dữ liệu được lưu trực tiếp "
        "trong cơ sở dữ liệu SQLite."
    )


# ============================================================
# THÊM KHÁCH HÀNG
# ============================================================

if menu == "➕ Thêm khách hàng":

    st.subheader("➕ Thêm khách hàng mới")

    st.write(
        "Nhập thông tin khách hàng vào biểu mẫu bên dưới."
    )

    with st.form("add_customer_form", clear_on_submit=True):

        col1, col2 = st.columns(2)

        with col1:

            phone = st.text_input(
                "Số điện thoại *",
                placeholder="Ví dụ: 0901234567"
            )

        with col2:

            name = st.text_input(
                "Tên khách hàng *",
                placeholder="Nhập tên khách hàng"
            )

        col3, col4 = st.columns(2)

        with col3:

            area = st.text_input(
                "Khu vực",
                placeholder="Ví dụ: Quận 1, TP.HCM"
            )

        with col4:

            note = st.text_input(
                "Ghi chú",
                placeholder="Nhập ghi chú"
            )

        st.write("")

        submitted = st.form_submit_button(
            "💾 Lưu khách hàng",
            use_container_width=True
        )

        if submitted:

            phone = phone.strip()
            name = name.strip()
            area = area.strip()
            note = note.strip()

            # Kiểm tra tên
            if not name:

                st.error(
                    "❌ Vui lòng nhập tên khách hàng."
                )

            # Kiểm tra số điện thoại
            elif not phone:

                st.error(
                    "❌ Vui lòng nhập số điện thoại."
                )

            elif not validate_phone(phone):

                st.error(
                    "❌ Số điện thoại không hợp lệ."
                )

            elif phone_exists(phone):

                st.warning(
                    "⚠️ Số điện thoại này đã tồn tại."
                )

            else:

                add_customer(
                    phone,
                    name,
                    area,
                    note
                )

                st.success(
                    "✅ Đã thêm khách hàng thành công!"
                )

                st.rerun()


# ============================================================
# DANH SÁCH KHÁCH HÀNG
# ============================================================

elif menu == "📋 Danh sách khách hàng":

    st.subheader("📋 Danh sách khách hàng")

    df = get_all_customers()

    if df.empty:

        st.info(
            "📭 Chưa có khách hàng nào trong hệ thống."
        )

    else:

        # ----------------------------------------------------
        # TÌM KIẾM
        # ----------------------------------------------------

        st.markdown("### 🔎 Tìm kiếm")

        search_col1, search_col2 = st.columns(2)

        with search_col1:

            search = st.text_input(
                "Tìm theo tên hoặc số điện thoại",
                placeholder="Nhập từ khóa..."
            )

        with search_col2:

            areas = [
                x for x in
                df["area"].dropna().unique().tolist()
                if str(x).strip()
            ]

            area_filter = st.selectbox(
                "Lọc theo khu vực",
                ["Tất cả"] + sorted(areas)
            )

        filtered_df = df.copy()

        # Tìm kiếm
        if search:

            search = search.lower().strip()

            filtered_df = filtered_df[
                filtered_df["name"]
                .astype(str)
                .str.lower()
                .str.contains(
                    search,
                    na=False
                )
                |
                filtered_df["phone"]
                .astype(str)
                .str.lower()
                .str.contains(
                    search,
                    na=False
                )
            ]

        # Lọc khu vực
        if area_filter != "Tất cả":

            filtered_df = filtered_df[
                filtered_df["area"]
                .fillna("")
                .eq(area_filter)
            ]

        st.write("")

        st.write(
            f"**Tìm thấy {len(filtered_df)} khách hàng**"
        )

        # ----------------------------------------------------
        # ĐỔI TÊN CỘT
        # ----------------------------------------------------

        display_df = filtered_df.copy()

        display_df = display_df.rename(
            columns={
                "id": "ID",
                "phone": "Số điện thoại",
                "name": "Tên khách hàng",
                "area": "Khu vực",
                "note": "Ghi chú",
                "created_at": "Ngày tạo"
            }
        )

        # ----------------------------------------------------
        # HIỂN THỊ BẢNG
        # ----------------------------------------------------

        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "ID": st.column_config.NumberColumn(
                    "ID",
                    width="small"
                ),

                "Số điện thoại": st.column_config.TextColumn(
                    "Số điện thoại",
                    width="medium"
                ),

                "Tên khách hàng": st.column_config.TextColumn(
                    "Tên khách hàng",
                    width="medium"
                ),

                "Khu vực": st.column_config.TextColumn(
                    "Khu vực",
                    width="medium"
                ),

                "Ghi chú": st.column_config.TextColumn(
                    "Ghi chú",
                    width="large"
                ),

                "Ngày tạo": st.column_config.TextColumn(
                    "Ngày tạo",
                    width="medium"
                )
            }
        )

        st.write("")

        # ----------------------------------------------------
        # XUẤT DỮ LIỆU
        # ----------------------------------------------------

        st.markdown("### 📥 Xuất dữ liệu")

        col1, col2 = st.columns(2)

        with col1:

            csv_data = display_df.to_csv(
                index=False
            ).encode("utf-8-sig")

            st.download_button(
                label="📄 Tải CSV",
                data=csv_data,
                file_name="danh_sach_khach_hang.csv",
                mime="text/csv",
                use_container_width=True
            )

        with col2:

            excel_data = dataframe_to_excel(
                filtered_df
            )

            st.download_button(
                label="📊 Tải Excel",
                data=excel_data,
                file_name="danh_sach_khach_hang.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True
            )


# ============================================================
# CHỈNH SỬA KHÁCH HÀNG
# ============================================================

elif menu == "✏️ Chỉnh sửa khách hàng":

    st.subheader("✏️ Chỉnh sửa thông tin khách hàng")

    df = get_all_customers()

    if df.empty:

        st.info(
            "📭 Chưa có khách hàng để chỉnh sửa."
        )

    else:

        # Tạo danh sách hiển thị
        customer_options = {}

        for _, row in df.iterrows():

            label = (
                f"{row['id']} - "
                f"{row['name']} - "
                f"{row['phone']}"
            )

            customer_options[label] = int(
                row["id"]
            )

        selected_label = st.selectbox(
            "Chọn khách hàng",
            list(customer_options.keys())
        )

        selected_id = customer_options[
            selected_label
        ]

        customer = get_customer_by_id(
            selected_id
        )

        if customer:

            customer_id = customer[0]
            old_phone = customer[1]
            old_name = customer[2]
            old_area = customer[3] or ""
            old_note = customer[4] or ""

            st.write("")

            with st.form("edit_customer_form"):

                col1, col2 = st.columns(2)

                with col1:

                    new_phone = st.text_input(
                        "Số điện thoại *",
                        value=old_phone
                    )

                with col2:

                    new_name = st.text_input(
                        "Tên khách hàng *",
                        value=old_name
                    )

                col3, col4 = st.columns(2)

                with col3:

                    new_area = st.text_input(
                        "Khu vực",
                        value=old_area
                    )

                with col4:

                    new_note = st.text_input(
                        "Ghi chú",
                        value=old_note
                    )

                st.write("")

                update_button = st.form_submit_button(
                    "💾 Cập nhật",
                    use_container_width=True
                )

                if update_button:

                    new_phone = new_phone.strip()
                    new_name = new_name.strip()
                    new_area = new_area.strip()
                    new_note = new_note.strip()

                    if not new_name:

                        st.error(
                            "❌ Vui lòng nhập tên khách hàng."
                        )

                    elif not new_phone:

                        st.error(
                            "❌ Vui lòng nhập số điện thoại."
                        )

                    elif not validate_phone(new_phone):

                        st.error(
                            "❌ Số điện thoại không hợp lệ."
                        )

                    elif phone_exists(
                        new_phone,
                        exclude_id=customer_id
                    ):

                        st.warning(
                            "⚠️ Số điện thoại này "
                            "đã được sử dụng bởi khách hàng khác."
                        )

                    else:

                        update_customer(
                            customer_id,
                            new_phone,
                            new_name,
                            new_area,
                            new_note
                        )

                        st.success(
                            "✅ Đã cập nhật khách hàng!"
                        )

                        st.rerun()


# ============================================================
# XÓA KHÁCH HÀNG
# ============================================================

elif menu == "🗑️ Xóa khách hàng":

    st.subheader("🗑️ Xóa khách hàng")

    df = get_all_customers()

    if df.empty:

        st.info(
            "📭 Chưa có khách hàng để xóa."
        )

    else:

        customer_options = {}

        for _, row in df.iterrows():

            label = (
                f"{row['id']} - "
                f"{row['name']} - "
                f"{row['phone']}"
            )

            customer_options[label] = int(
                row["id"]
            )

        selected_label = st.selectbox(
            "Chọn khách hàng cần xóa",
            list(customer_options.keys())
        )

        selected_id = customer_options[
            selected_label
        ]

        customer = get_customer_by_id(
            selected_id
        )

        if customer:

            st.warning(
                "⚠️ Bạn đang chuẩn bị xóa khách hàng:"
            )

            info_col1, info_col2 = st.columns(2)

            with info_col1:

                st.write(
                    f"**Tên:** {customer[2]}"
                )

                st.write(
                    f"**Số điện thoại:** {customer[1]}"
                )

            with info_col2:

                st.write(
                    f"**Khu vực:** {customer[3] or '---'}"
                )

                st.write(
                    f"**Ghi chú:** {customer[4] or '---'}"
                )

            st.write("")

            confirm = st.checkbox(
                "Tôi xác nhận muốn xóa khách hàng này."
            )

            if confirm:

                if st.button(
                    "🗑️ Xóa khách hàng",
                    type="primary",
                    use_container_width=True
                ):

                    delete_customer(
                        selected_id
                    )

                    st.success(
                        "✅ Đã xóa khách hàng."
                    )

                    st.rerun()


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "👥 Ứng dụng quản lý khách hàng | "
    "Streamlit + SQLite"
)
