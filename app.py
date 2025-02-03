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
    # 前処理：改行とスペースの正規化
    lines = text.replace('\r\n', '\n').split('\n')
    entries = []
    entry_id = 0
    current_entry = None
    
    for line in lines:
        if '【ピノ】' in line:
            # 前のエントリーがあれば保存
            if current_entry:
                entry_id += 1
                current_entry['id'] = entry_id
                entries.append(current_entry)
            
            # 新しいエントリーの開始
            current_entry = {
                'content': line,
                'distance_found': False
            }
        elif current_entry and not current_entry['distance_found']:
            # 距離が見つかるまで内容を追加
            current_entry['content'] += ' ' + line
            
            # 距離のパターンをチェック
            distance_match = re.search(r'(\d+\.?\d*)(?:km|㎞|ｋｍ|kｍ)', line)
            if distance_match:
                current_entry['distance_found'] = True
    
    # 最後のエントリーを保存
    if current_entry:
        entry_id += 1
        current_entry['id'] = entry_id
        entries.append(current_entry)
    
    # エントリーの解析
    parsed_entries = []
    for entry in entries:
        try:
            content = entry['content']
            parts = content.split('】')[1].strip().split()
            name = parts[0]
            date = parts[1]
            
            # 経路と距離の抽出
            distance_match = re.search(r'(\d+\.?\d*)(?:km|㎞|ｋｍ|kｍ)', content)
            if distance_match:
                distance = float(distance_match.group(1))
                route_end = content.find(distance_match.group(0))
                route_start = content.find(date) + len(date)
                route = content[route_start:route_end].strip()
                
                parsed_entries.append({
                    'id': entry['id'],
                    'name': name,
                    'date': date,
                    'route': route,
                    'distance': distance
                })
        except Exception as e:
            print(f"Error parsing entry: {content}")
            print(f"Error: {str(e)}")
            continue
    
    return pd.DataFrame(parsed_entries)

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
            
            # データフレームを表示（入力順を維持）
            st.dataframe(
                df,
                column_config={
                    'id': st.column_config.NumberColumn('No.', width=70),
                    'date': st.column_config.TextColumn('日付', width=100),
                    'name': st.column_config.TextColumn('担当者', width=120),
                    'route': st.column_config.TextColumn('経路', width=500),
                    'distance': st.column_config.NumberColumn('距離(km)', format="%.1f", width=100)
                },
                hide_index=True,
                height=400
            )
            
            # 合計距離の表示
            total_distance = df['distance'].sum()
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
