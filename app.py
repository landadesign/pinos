from streamlit import set_page_config
set_page_config(layout="wide")

import streamlit as st
import pandas as pd
from datetime import datetime
import io
from PIL import Image, ImageDraw, ImageFont
import re

# 定数
RATE_PER_KM = 15
DAILY_ALLOWANCE = 200

def create_expense_table_image(df, name):
    # 画像サイズとフォント設定
    width = 1200
    row_height = 40
    header_height = 60
    padding = 30
    title_height = 50
    
    # 全行数を計算（タイトル + ヘッダー + データ行 + 合計行 + 注釈）
    total_rows = len(df) + 3
    height = title_height + header_height + (total_rows * row_height) + padding * 2
    
    # 画像作成
    img = Image.new('RGB', (width, height), 'white')
    draw = ImageDraw.Draw(img)
    
    # フォント設定
    font = ImageFont.load_default()
    
    # タイトル描画
    title = f"{name}様 2024年12月25日～2025年1月 社内通貨（交通費）清算額"
    draw.text((padding, padding), title, fill='black', font=font)
    
    # ヘッダー
    headers = ['日付', '経路', '合計距離(km)', '交通費（距離×15P）(円)', '運転手当(円)', '合計(円)']
    x_positions = [padding, padding + 80, padding + 600, padding + 750, padding + 900, padding + 1050]
    
    y = padding + title_height
    for header, x in zip(headers, x_positions):
        draw.text((x, y), header, fill='black', font=font)
    
    # 罫線
    line_y = y + header_height - 5
    draw.line([(padding, line_y), (width - padding, line_y)], fill='black', width=1)
    
    # データ行
    y = padding + title_height + header_height
    for _, row in df.iterrows():
        for route in row['routes']:
            # 日付
            draw.text((x_positions[0], y), str(row['date']), fill='black', font=font)
            
            # 経路
            draw.text((x_positions[1], y), route['route'], fill='black', font=font)
            
            # 最初のルートの行にのみ数値を表示
            if route == row['routes'][0]:
                # 距離
                draw.text((x_positions[2], y), f"{row['total_distance']:.1f}", fill='black', font=font)
                
                # 交通費
                draw.text((x_positions[3], y), f"{int(row['transportation_fee']):,}", fill='black', font=font)
                
                # 運転手当
                draw.text((x_positions[4], y), f"{int(row['allowance']):,}", fill='black', font=font)
                
                # 合計
                draw.text((x_positions[5], y), f"{int(row['total']):,}", fill='black', font=font)
            
            y += row_height
    
    # 合計行の罫線
    line_y = y - 5
    draw.line([(padding, line_y), (width - padding, line_y)], fill='black', width=1)
    
    # 合計行
    y += 10
    draw.text((x_positions[0], y), "合計", fill='black', font=font)
    draw.text((x_positions[5], y), f"{int(df['total'].sum()):,}", fill='black', font=font)
    
    # 注釈
    y += row_height * 2
    draw.text((padding, y), "※2025年1月分給与にて清算しました。", fill='black', font=font)
    
    # 計算日時
    draw.text((padding, y + row_height), f"計算日時: {datetime.now().strftime('%Y/%m/%d')}", fill='black', font=font)
    
    # 画像をバイト列に変換
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='PNG')
    img_byte_arr = img_byte_arr.getvalue()
    
    return img_byte_arr

def parse_expense_data(text):
    try:
        # テキストの前処理
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        data = []
        daily_routes = {}
        
        # 各行を解析
        for line in lines:
            # 【ピノ】形式のデータを解析
            if '【ピノ】' in line:
                # パターン: 【ピノ】名前 日付(曜日) 経路 距離
                pino_match = re.match(r'【ピノ】\s*([^\s]+)\s+(\d+/\d+)\s*\(.\)\s*(.+?)(?:\s+(\d+\.?\d*)(?:km|㎞|ｋｍ|kｍ))?$', line)
                if pino_match:
                    name = pino_match.group(1).replace('様', '')
                    date = pino_match.group(2)
                    route = pino_match.group(3).strip()
                    distance_str = pino_match.group(4)
                    
                    # 距離の取得
                    if distance_str:
                        distance = float(distance_str)
                    else:
                        # 経路からポイント数を計算（デフォルトの場合）
                        route_points = route.split('→')
                        distance = (len(route_points) - 1) * 5.0
                    
                    if name not in daily_routes:
                        daily_routes[name] = {}
                    if date not in daily_routes[name]:
                        daily_routes[name][date] = []
                    
                    # 重複チェック
                    route_exists = False
                    for existing_route in daily_routes[name][date]:
                        if existing_route['route'] == route:
                            route_exists = True
                            break
                    
                    if not route_exists:
                        daily_routes[name][date].append({
                            'route': route,
                            'distance': distance
                        })
        
        # 日付ごとのデータを集計
        for name, dates in daily_routes.items():
            for date, routes in sorted(dates.items(), key=lambda x: tuple(map(int, x[0].split('/')))):
                # 同じ日の距離を合算
                total_distance = sum(route['distance'] for route in routes)
                transportation_fee = int(total_distance * RATE_PER_KM)  # 切り捨て
                
                data.append({
                    'name': name,
                    'date': date,
                    'routes': routes,
                    'total_distance': total_distance,
                    'transportation_fee': transportation_fee,
                    'allowance': DAILY_ALLOWANCE,  # 1日1回のみ
                    'total': transportation_fee + DAILY_ALLOWANCE
                })
        
        if data:
            # データをDataFrameに変換
            df = pd.DataFrame(data)
            
            # 日付でソート
            df['date_sort'] = df['date'].apply(lambda x: tuple(map(int, x.split('/'))))
            df = df.sort_values(['name', 'date_sort'])
            df = df.drop('date_sort', axis=1)
            
            return df
        
        st.error("データが見つかりませんでした。正しい形式で入力してください。")
        return None
        
    except Exception as e:
        st.error(f"エラーが発生しました: {str(e)}")
        return None

def main():
    st.title("PINO精算アプリケーション")
    
    # テキストエリアの表示
    input_text = st.text_area("精算データを貼り付けてください", height=200)
    
    # 解析ボタンとクリアボタン
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("データを解析"):
            if input_text:
                # データを解析
                df = parse_expense_data(input_text)
                st.session_state['df'] = df
                st.success("データを解析しました！")
    with col2:
        if st.button("クリア"):
            st.session_state['df'] = None
            st.session_state['show_expense_report'] = False
            st.experimental_rerun()
    
    # データ一覧の表示
    if 'df' in st.session_state and st.session_state['df'] is not None:
        df = st.session_state['df']
        if not df.empty:
            st.markdown("""
            <h2 style='text-align: center; padding: 20px 0;'>
                交通費データ一覧
            </h2>
            """, unsafe_allow_html=True)
            
            # 表示用のデータを作成（入力順を維持）
            display_rows = []
            for _, row in df.iterrows():
                display_rows.append({
                    '日付': row['date'],
                    '担当者': row['name'],
                    '経路': row['route'],
                    '距離(km)': row['distance']
                })
            
            # データフレームを表示
            display_df = pd.DataFrame(display_rows)
            st.dataframe(
                display_df,
                column_config={
                    '日付': st.column_config.TextColumn('日付', width=100),
                    '担当者': st.column_config.TextColumn('担当者', width=120),
                    '経路': st.column_config.TextColumn('経路', width=500),
                    '距離(km)': st.column_config.NumberColumn('距離(km)', format="%.1f", width=100)
                },
                hide_index=True,
                height=400
            )
            
            # 合計距離の表示
            total_distance = display_df['距離(km)'].sum()
            st.markdown(f"""
            <div style='text-align: right; padding: 10px; margin-top: 10px;'>
                <h3>合計距離 {total_distance:.1f} km</h3>
            </div>
            """, unsafe_allow_html=True)
            
            # 精算書を表示するボタン
            col1, col2 = st.columns([1, 5])
            with col1:
                if st.button("精算書を表示"):
                    st.session_state['show_expense_report'] = True
                    st.experimental_rerun()

    # 精算書の表示
    if st.session_state.get('show_expense_report', False) and 'df' in st.session_state:
        df = st.session_state['df']
        
        # 担当者ごとのタブを作成
        unique_names = sorted(df['name'].unique())
        tabs = st.tabs([f"{name}様" for name in unique_names])
        
        # 担当者ごとの精算書表示
        for i, name in enumerate(unique_names):
            with tabs[i]:
                st.markdown(f"### {name}様 12月25日～1月 社内通貨（交通費）清算額")
                
                # 担当者のデータを抽出
                person_data = df[df['name'] == name].copy()
                
                # 精算書データの作成
                expense_data = create_expense_report(person_data)
                
                # 精算書の表示
                st.dataframe(
                    expense_data,
                    column_config={
                        '日付': st.column_config.TextColumn('日付', width=100),
                        '経路': st.column_config.TextColumn('経路', width=400),
                        '合計距離(km)': st.column_config.NumberColumn('合計距離(km)', format="%.1f", width=120),
                        '交通費（距離×15P）(円)': st.column_config.NumberColumn('交通費（距離×15P）(円)', format="%d", width=150),
                        '運転手当(円)': st.column_config.NumberColumn('運転手当(円)', format="%d", width=120),
                        '合計(円)': st.column_config.NumberColumn('合計(円)', format="%d", width=120)
                    },
                    hide_index=True
                )
                
                # 注釈表示
                st.markdown("""
                    <div style='margin-top: 15px; color: #666;'>
                        ※2025年1月分給与にて清算しました。
                    </div>
                """, unsafe_allow_html=True)
        
        # 一括印刷ボタン
        st.markdown("---")
        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("表示中の精算書を印刷"):
                st.markdown("""
                <script>
                    window.print();
                </script>
                """, unsafe_allow_html=True)
        with col2:
            if st.button("全ての精算書を一括印刷"):
                st.markdown("""
                <script>
                    window.print();
                </script>
                """, unsafe_allow_html=True)

def create_expense_report(person_data):
    # 精算書データの作成ロジック
    # ... (以前の精算書作成ロジックを実装)
    pass

if __name__ == "__main__":
    main()
