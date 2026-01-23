import streamlit as st
import pandas as pd
import io
import plotly.graph_objects as go
from datetime import datetime

# 1. CẤU HÌNH TRANG
st.set_page_config(page_title="Truck & Vol Analysis Pro", layout="wide", page_icon="🚚")

st.title("🚚 Phân tích Truck & Volume")
st.markdown("---")

# ======================
# 2. INPUT & XỬ LÝ DỮ LIỆU
# ======================
st.sidebar.header("📥 Control Panel")
uploaded_file = st.sidebar.file_uploader("Nạp dữ liệu hệ thống (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        # Đọc dữ liệu
        df = pd.read_excel(uploaded_file, sheet_name="Truck-Vol")
        df.columns = df.columns.str.strip()

        # Mapping cột
        df = df.rename(columns={
            df.columns[0]: "hour", df.columns[1]: "date",
            df.columns[5]: "truck", df.columns[6]: "volume"
        })

        # Tiền xử lý
        df["date"] = pd.to_datetime(df["date"], errors='coerce')
        df = df.dropna(subset=["date"])
        df["hour"] = pd.to_numeric(df["hour"], errors='coerce').fillna(0).astype(int)
        df["truck"] = pd.to_numeric(df["truck"], errors='coerce').fillna(0)
        df["volume"] = pd.to_numeric(df["volume"], errors='coerce').fillna(0)
        df["weekday_num"] = df["date"].dt.dayofweek
        df["slot"] = df["hour"].apply(lambda h: f"{h:02d}:00-{(h+1):02d}:00")

        # ----------------------
        # BỘ LỌC THỨ
        # ----------------------
        st.sidebar.subheader("📅 Bộ lọc theo nhóm thứ")
        filter_type = st.sidebar.selectbox(
            "Chọn kiểu lọc ngày:",
            ["Tất cả (All)", "Thứ 2", "Thứ 3", "Thứ 4 đến Thứ 7", "Chủ nhật"]
        )

        if filter_type == "Thứ 2":
            df_filtered = df[df["weekday_num"] == 0]
        elif filter_type == "Thứ 3":
            df_filtered = df[df["weekday_num"] == 1]
        elif filter_type == "Thứ 4 đến Thứ 7":
            df_filtered = df[df["weekday_num"].isin([2, 3, 4, 5])]
        elif filter_type == "Chủ nhật":
            df_filtered = df[df["weekday_num"] == 6]
        else:
            df_filtered = df

        # ----------------------
        # CHỌN GIAI ĐOẠN
        # ----------------------
        st.subheader(f"📊 Đang phân tích: {filter_type}")
        c1, c2 = st.columns(2)
        with c1:
            range_1 = st.date_input("Giai đoạn 1 (Gốc)", [df["date"].min(), df["date"].median()], key="r1")
        with c2:
            range_2 = st.date_input("Giai đoạn 2 (So sánh)", [df["date"].median(), df["date"].max()], key="r2")

        if len(range_1) == 2 and len(range_2) == 2:
            df_g1 = df_filtered[(df_filtered["date"] >= pd.to_datetime(range_1[0])) & (df_filtered["date"] <= pd.to_datetime(range_1[1]))]
            df_g2 = df_filtered[(df_filtered["date"] >= pd.to_datetime(range_2[0])) & (df_filtered["date"] <= pd.to_datetime(range_2[1]))]

            def get_metrics(data):
                if data.empty: return pd.DataFrame()
                daily = data.groupby(["date", "slot", "hour"])[["truck", "volume"]].agg(["min", "max"])
                daily.columns = ['_'.join(col) for col in daily.columns]
                return daily.groupby(["slot", "hour"]).mean().round(0).reset_index()

            res1 = get_metrics(df_g1)
            res2 = get_metrics(df_g2)

            combined = res1.merge(res2, on=["slot", "hour"], how="outer", suffixes=('_g1', '_g2')).fillna(0)
            combined = combined.sort_values("hour")

            # ======================
            # 3. LOGIC REMARK (CẬP NHẬT TĂNG/GIẢM ĐA CẤP)
            # ======================
            def process_logic(v1, v2):
                diff_val = int(v2 - v1)
                pct = round((diff_val / v1 * 100), 0) if v1 != 0 else 0
                
                # Logic Giảm
                if pct < -15:
                    remark = "🚨 Giảm sâu"
                elif -15 <= pct < -5:
                    remark = "📉 Giảm mạnh"
                elif -5 <= pct < 0:
                    remark = "✅ Bình thường"
                # Logic Tăng
                elif 0 <= pct < 5:
                    remark = "✅ Bình thường"
                elif 5 <= pct < 10:
                    remark = "🔼 Tăng nhẹ"
                elif 10 <= pct < 15:
                    remark = "🚀 Tăng mạnh"
                else: # >= 15%
                    remark = "🚨 CẢNH BÁO"
                
                return f"{diff_val} | {int(pct)}%", remark

            truck_eval = combined.apply(lambda r: process_logic(r['truck_max_g1'], r['truck_max_g2']), axis=1)
            vol_eval = combined.apply(lambda r: process_logic(r['volume_max_g1'], r['volume_max_g2']), axis=1)

            # Bảng hiển thị
            display_df = pd.DataFrame({
                "Khung giờ": combined["slot"],
                "Xe GĐ1 (Avg)": combined["truck_min_g1"].astype(int).astype(str) + "–" + combined["truck_max_g1"].astype(int).astype(str),
                "Xe GĐ2 (Avg)": combined["truck_min_g2"].astype(int).astype(str) + "–" + combined["truck_max_g2"].astype(int).astype(str),
                "Diff Xe (±|%)": [x[0] for x in truck_eval],
                "Truck Remark": [x[1] for x in truck_eval],
                "Vol GĐ1 (Avg)": combined["volume_min_g1"].astype(int).astype(str) + "–" + combined["volume_max_g1"].astype(int).astype(str),
                "Vol GĐ2 (Avg)": combined["volume_min_g2"].astype(int).astype(str) + "–" + combined["volume_max_g2"].astype(int).astype(str),
                "Diff Vol (±|%)": [x[0] for x in vol_eval],
                "Vol Remark": [x[1] for x in vol_eval]
            })

            # Hiển thị bảng màu (Cập nhật màu cho nhóm Giảm)
            def style_remark(val):
                # Nhóm Tăng & Cảnh báo
                if val == "🚨 CẢNH BÁO": return 'background-color: #ff4b4b; color: white'
                if val == "🚀 Tăng mạnh": return 'background-color: #ffa500; color: white'
                if val == "🔼 Tăng nhẹ": return 'background-color: #1e90ff; color: white'
                # Nhóm Giảm
                if val == "🚨 Giảm sâu": return 'background-color: #8b0000; color: white' # Đỏ đậm
                if val == "📉 Giảm mạnh": return 'background-color: #2e8b57; color: white' # Xanh lá đậm
                if val == "📉 Giảm nhẹ": return 'background-color: #90ee90; color: black' # Xanh lá nhạt
                return ''

            st.dataframe(display_df.style.applymap(style_remark, subset=['Truck Remark', 'Vol Remark']), use_container_width=True, hide_index=True)

            # ======================
            # 4. BIỂU ĐỒ COMBO (FIX LỖI TITLE_FONT)
            # ======================
            st.divider()
            st.subheader(f"📈 Biểu đồ Combo - {filter_type} (Giai đoạn 2)")

            fig = go.Figure()

            # TRUCK (BAR)
            fig.add_trace(go.Bar(
                x=combined["slot"],
                y=combined["truck_max_g2"],
                name='Truck',
                marker_color='#1f77b4',
                opacity=0.7,
                yaxis='y1',
                text=combined["truck_max_g2"].astype(int),
                textposition='auto',
            ))

            # VOLUME (LINE)
            fig.add_trace(go.Scatter(
                x=combined["slot"],
                y=combined["volume_max_g2"],
                name='Volume',
                mode='lines+markers+text',
                line=dict(color='#ff7f0e', width=3),
                marker=dict(size=8),
                text=combined["volume_max_g2"].astype(int),
                textposition="top center",
                yaxis='y2'
            ))

            fig.update_layout(
                title=f'Tương quan Truck và Volume ({filter_type})',
                xaxis=dict(title='Khung giờ'),
                yaxis=dict(
                    title='Số lượng Xe (Truck)',
                    title_font=dict(color='#1f77b4'), # Đã sửa thành title_font
                    tickfont=dict(color='#1f77b4'),
                    range=[0, combined["truck_max_g2"].max() * 1.3]
                ),
                yaxis2=dict(
                    title='Tải trọng (Volume)',
                    title_font=dict(color='#ff7f0e'), # Đã sửa thành title_font
                    tickfont=dict(color='#ff7f0e'),
                    overlaying='y',
                    side='right',
                    range=[0, combined["volume_max_g2"].max() * 1.3]
                ),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                height=600,
                margin=dict(t=100)
            )

            st.plotly_chart(fig, use_container_width=True)

            # Nút tải Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                display_df.to_excel(writer, index=False)
            st.download_button("📥 Tải báo cáo (.xlsx)", buffer.getvalue(), f"Bao_cao_{filter_type}.xlsx")

    except Exception as e:
        st.error(f"Lỗi hệ thống: {e}")
else:
    st.info("👋 Hãy nạp file Excel vào thanh công cụ bên trái để bắt đầu.")

