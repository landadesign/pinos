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
    col1, col2 = st.columns([1, 5])
    with col1:
        if st.button("データを解析"):
            if input_text:
                # 元のデータ順を保持するために入力テキストを行ごとに分割
                input_lines = [line.strip() for line in input_text.split('\n') if line.strip()]
                
                # データを解析
                df = parse_expense_data(input_text)
                st.session_state['df'] = df
                st.success("データを解析しました！")
                
                # データ一覧の表示
                if not df.empty:
                    st.markdown("""
                    <h2 style='text-align: center; color: #1f77b4; padding: 20px 0;'>
                        交通費データ一覧
                    </h2>
                    """, unsafe_allow_html=True)
                    
                    # 表示用のデータを作成
                    display_rows = []
                    i = 0
                    while i < len(input_lines):
                        line = input_lines[i]
                        if line.startswith('【ピノ】'):
                            # 距離の抽出と変換
                            distance_str = line.split()[-1]
                            distance_str = distance_str.replace('ｋｍ', '').replace('km', '')
                            try:
                                distance = float(distance_str)
                            except ValueError:
                                # 次の行に距離がある可能性をチェック
                                if i + 1 < len(input_lines):
                                    next_line = input_lines[i + 1]
                                    if not next_line.startswith('【ピノ】'):
                                        try:
                                            distance = float(next_line.replace('ｋｍ', '').replace('km', ''))
                                            i += 1  # 次の行を処理済みとしてスキップ
                                        except ValueError:
                                            distance = 0.0
                                else:
                                    distance = 0.0
                                
                            # 経路の抽出
                            route_parts = line.split('】')[1].split()
                            route = ' '.join(route_parts[2:-1]) if distance > 0 else ' '.join(route_parts[2:])
                            
                            display_rows.append({
                                '入力データ': line,
                                '担当者': line.split()[1],
                                '日付': line.split()[2],
                                '経路': route,
                                '距離(km)': distance,
                            })
                        i += 1
                    
                    # 表示用のDataFrame作成
                    display_df = pd.DataFrame(display_rows)
                    
                    # データフレームを大きく表示
                    st.dataframe(
                        display_df[['入力データ', '担当者', '日付', '経路', '距離(km)']],
                        column_config={
                            '入力データ': st.column_config.TextColumn(
                                '入力データ',
                                width=900,
                                help="元の入力データ"
                            ),
                            '担当者': st.column_config.TextColumn(
                                '担当者',
                                width=200,
                                help="担当者名"
                            ),
                            '日付': st.column_config.TextColumn(
                                '日付',
                                width=150,
                                help="実施日"
                            ),
                            '経路': st.column_config.TextColumn(
                                '経路',
                                width=800,
                                help="移動経路"
                            ),
                            '距離(km)': st.column_config.NumberColumn(
                                '距離(km)',
                                format="%.1f km",
                                width=150,
                                help="移動距離"
                            )
                        },
                        hide_index=True,
                        height=800,
                        use_container_width=True
                    )
                    
                    # 合計距離の表示
                    total_distance = display_df['距離(km)'].sum()
                    st.markdown(f"""
                    <div style='text-align: right; padding: 20px; background-color: #f0f2f6; border-radius: 5px; margin-top: 20px;'>
                        <h3>合計距離: {total_distance:.1f} km</h3>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # 精算書を表示するボタン
                    st.markdown("<div style='padding: 20px 0;'>", unsafe_allow_html=True)
                    if st.button("精算書を表示", type="primary"):
                        st.session_state['show_expense_report'] = True
                        st.experimental_rerun()
                    st.markdown("</div>", unsafe_allow_html=True)

    # 精算書の表示
    if st.session_state.get('show_expense_report', False) and 'df' in st.session_state:
        df = st.session_state['df']
        st.write("### 精算書")
        
        # タブの作成
        unique_names = sorted(df['name'].unique())
        tabs = st.tabs([f"{name}様" for name in unique_names])
        
        # 担当者ごとの精算書表示
        for i, name in enumerate(unique_names):
            with tabs[i]:
                # タイトル表示
                st.markdown(f"### {name}様 2024年12月25日～2025年1月　社内通貨（交通費）清算額")
                
                # データの準備
                person_data = df[df['name'] == name].copy()
                
                # 日付ごとのデータをグループ化
                daily_data = {}
                
                for _, row in person_data.iterrows():
                    date = row.get('date', '')
                    if not date:
                        continue
                    
                    if date not in daily_data:
                        daily_data[date] = {
                            'routes': [],
                            'total_distance': 0
                        }
                    
                    for route in row['routes']:
                        route_text = route.get('route', '').replace('\n', ' ').strip()
                        distance = route.get('distance', 0)
                        daily_data[date]['routes'].append({
                            'route': route_text or '',
                            'distance': distance or 0
                        })
                        daily_data[date]['total_distance'] += distance or 0
                
                # 表示用データの作成
                display_rows = []
                
                # 日付でソートしてデータを処理
                for date in sorted(daily_data.keys(), key=lambda x: tuple(map(int, x.split('/')))):
                    day_data = daily_data[date]
                    
                    # 同日の経路を別々の行に表示
                    for route in day_data['routes']:
                        row_data = {
                            '日付': date or '',
                            '経路': route['route'] or '',
                            '合計\n距離\n(km)': route['distance'] or '',
                            '交通費\n(距離×15P)\n(円)': '',
                            '運転\n手当\n(円)': '',
                            '合計\n(円)': ''
                        }
                        display_rows.append(row_data)
                    
                    # 日ごとの合計行を更新
                    total_transport = int((day_data['total_distance'] or 0) * 15)
                    daily_total = total_transport + 200
                    
                    # 最後の行を更新
                    if display_rows:
                        display_rows[-1].update({
                            '交通費\n(距離×15P)\n(円)': f"{total_transport:,}" if total_transport else '',
                            '運転\n手当\n(円)': "200",
                            '合計\n(円)': f"{daily_total:,}" if daily_total else ''
                        })
                
                # 総合計の計算
                if display_rows:
                    total_all = sum(int(row['合計\n(円)'].replace(',', '')) 
                                  for row in display_rows if row['合計\n(円)'])
                    
                    # 合計行の追加
                    display_rows.append({
                        '日付': '合計',
                        '経路': '',
                        '合計\n距離\n(km)': '',
                        '交通費\n(距離×15P)\n(円)': '',
                        '運転\n手当\n(円)': '',
                        '合計\n(円)': f"{total_all:,}" if total_all else ''
                    })
                
                # DataFrameの表示
                display_df = pd.DataFrame(display_rows)
                st.dataframe(
                    display_df,
                    column_config={
                        '日付': st.column_config.TextColumn('日付', width=100),
                        '経路': st.column_config.TextColumn('経路', width=400),
                        '合計\n距離\n(km)': st.column_config.NumberColumn('合計\n距離\n(km)', format="%.1f", width=100),
                        '交通費\n(距離×15P)\n(円)': st.column_config.TextColumn('交通費\n(距離×15P)\n(円)', width=150),
                        '運転\n手当\n(円)': st.column_config.TextColumn('運転\n手当\n(円)', width=100),
                        '合計\n(円)': st.column_config.TextColumn('合計\n(円)', width=100)
                    },
                    use_container_width=False,
                    hide_index=True
                )
                
                # 注釈表示
                st.markdown("""
                    <div style='margin-top: 15px; color: #666;'>
                        ※2025年1月分給与にて清算しました。
                    </div>
                """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
