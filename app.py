"""
在留資格管理システム - Flask Web版
セル単位で背景色を設定できるHTMLベースのGUI
"""

from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for
import pandas as pd
from datetime import datetime, timedelta
from residence_manager import ResidenceStatusManager
import os
import json
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# アップロードフォルダを作成
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# グローバル変数でデータ管理
current_manager = None
current_file = None
last_export_info = {'timestamp': None, 'download_name': None}


def get_deadline_status(deadline_date):
    """期限日の状態を取得"""
    if pd.isna(deadline_date):
        return None
    
    today = datetime.now().date()
    if isinstance(deadline_date, (datetime, pd.Timestamp)):
        deadline_date = deadline_date.date()
    
    if deadline_date < today:
        days_overdue = (today - deadline_date).days
        return {'status': 'overdue', 'days': days_overdue}
    else:
        days_remaining = (deadline_date - today).days
        return {'status': 'ok', 'days': days_remaining}


def get_expiration_status(days):
    """満了日数から状態を取得"""
    if pd.isna(days):
        return 'unknown'
    elif days < 0:
        return 'expired'
    elif days <= 7:
        return 'urgent'
    elif days <= 30:
        return 'warning'
    elif days <= 90:
        return 'caution'
    else:
        return 'safe'


def format_date(date_val):
    """日付をフォーマット"""
    if pd.isna(date_val):
        return '-'
    if isinstance(date_val, (datetime, pd.Timestamp)):
        return date_val.strftime('%Y/%m/%d')
    if isinstance(date_val, str):
        # yyyy-mm-dd を yyyy/mm/dd に変換
        return date_val.replace('-', '/')
    return str(date_val)


def parse_threshold_value(value, default=None):
    """設定期限の値を整数に変換（無効値はデフォルトにフォールバック）"""
    if pd.isna(value):
        return default
    try:
        return int(value)
    except (ValueError, TypeError):
        try:
            return int(float(value))
        except (ValueError, TypeError):
            return default


def resolve_thresholds(setting1=None, setting2=None, setting3=None):
    """設定期限に基づきレベル別の残日数上限を算出"""
    defaults = (90, 60, 30)
    sanitized_values = []
    for value, default in zip((setting1, setting2, setting3), defaults):
        if value is None:
            sanitized_values.append(default)
        else:
            try:
                sanitized_values.append(max(int(value), 0))
            except (ValueError, TypeError):
                try:
                    sanitized_values.append(max(int(float(value)), 0))
                except (ValueError, TypeError):
                    sanitized_values.append(default)

    # level1/level2/level3 should directly reflect 設定期限1-3 respectively
    return {
        'level1_max': sanitized_values[0],
        'level2_max': sanitized_values[1],
        'level3_max': sanitized_values[2],
    }


def determine_deadline_level(days_to_expiration, thresholds):
    """残日数に応じた期限レベルを判定"""
    if days_to_expiration is None or days_to_expiration < 0:
        return None

    level_thresholds = []
    for level in ('level1', 'level2', 'level3'):
        max_val = thresholds.get(f'{level}_max')
        if max_val is None:
            continue
        level_thresholds.append((max_val, level))

    level_thresholds.sort(key=lambda item: item[0])

    for max_val, level in level_thresholds:
        if days_to_expiration <= max_val:
            return level

    return None


def load_default_file(filepath):
    """デフォルトファイルを読み込み"""
    global current_manager, current_file, last_export_info

    try:
        current_manager = ResidenceStatusManager(filepath)
        if current_manager.load_excel() and current_manager.process_data():
            current_file = filepath
            last_export_info = {'timestamp': None, 'download_name': None}
            return True
    except:
        pass
    return False


@app.route('/')
def index():
    """メインページ"""
    return render_template('index.html')


@app.route('/calendar')
def calendar():
    """カレンダーページ"""
    return render_template('calendar.html')


@app.route('/api/data')
def get_data():
    """データを取得するAPI"""
    global current_manager, current_file
    
    if current_manager is None or current_manager.df is None:
        return jsonify({'error': 'データが読み込まれていません'}), 400
    
    df = current_manager.df
    today = datetime.now().date()
    
    # データを整形
    data_list = []
    for idx, row in df.iterrows():
        # 満了日数（特定技能1号の累積日数）
        days = row.get('満了日数', None)
        
        # 既満了日数を取得
        ki_manryo_days = row.get('既満了日数', None)
        
        # 満了日数の表示を決定
        manryo_days_display = '-'
        manryo_days_value = None
        if not pd.isna(days):
            # 満了日数が計算されている場合は表示
            manryo_days_display = f"{int(days)}日"
            manryo_days_value = int(days)
        
        # 満了年月日までの残り日数を計算
        expiration_date = row.get('満了年月日')
        days_to_expiration = None
        if not pd.isna(expiration_date):
            if isinstance(expiration_date, (datetime, pd.Timestamp)):
                exp_date = expiration_date.date()
            else:
                exp_date = pd.to_datetime(expiration_date).date()
            days_to_expiration = (exp_date - today).days
        
        expiration_status = get_expiration_status(days_to_expiration)
        
        # 各期限日の状態をチェック
        deadline1_status = get_deadline_status(row.get('期限日1'))
        deadline2_status = get_deadline_status(row.get('期限日2'))
        deadline3_status = get_deadline_status(row.get('期限日3'))

        setting1_val = parse_threshold_value(row.get('設定期限1'))
        setting2_val = parse_threshold_value(row.get('設定期限2'))
        setting3_val = parse_threshold_value(row.get('設定期限3'))
        thresholds = resolve_thresholds(setting1_val, setting2_val, setting3_val)
        deadline_level = determine_deadline_level(days_to_expiration, thresholds)
        
        data_item = {
            'index': int(idx),
            '担当者コード': str(row.get('担当者コード', '-')),
            '氏名１': str(row.get('氏名１', '-')),
            '氏名２': str(row.get('氏名２', '-')),
            '在留資格': str(row.get('在留資格', '-')),
            '国籍': str(row.get('国籍', '-')),
            '在留カード番号': str(row.get('在留カード番号', '-')),
            '生年月日': format_date(row.get('生年月日')),
            '期生': str(row.get('期生', '-')),
            '許可年月日': format_date(row.get('許可年月日')),
            '満了年月日': format_date(row.get('満了年月日')),
            '満了日数': manryo_days_display,
            '満了日数_値': manryo_days_value,
            '満了年月日までの日数': days_to_expiration,
            '期限日1': format_date(row.get('期限日1')),
            '期限日1_状態': deadline1_status,
            '期限日2': format_date(row.get('期限日2')),
            '期限日2_状態': deadline2_status,
            '期限日3': format_date(row.get('期限日3')),
            '期限日3_状態': deadline3_status,
            '状態': expiration_status,
            '設定期限1': setting1_val,
            '設定期限2': setting2_val,
            '設定期限3': setting3_val,
            '期限レベル1_上限': thresholds['level1_max'],
            '期限レベル2_上限': thresholds['level2_max'],
            '期限レベル3_上限': thresholds['level3_max'],
            '期限レベル分類': deadline_level,
        }
        data_list.append(data_item)
    
    return jsonify({
        'data': data_list,
        'filename': os.path.basename(current_file) if current_file else '未選択',
        'total': len(data_list)
    })


@app.route('/api/summary')
def get_summary():
    """サマリー情報を取得するAPI"""
    global current_manager, current_file, last_export_info
    
    if current_manager is None or current_manager.df is None:
        return jsonify({'error': 'データが読み込まれていません'}), 400
    
    df = current_manager.df
    today = datetime.now().date()
    
    # 期限日超過を計算
    deadline_passed_count = 0
    for idx, row in df.iterrows():
        for deadline_col in ['期限日1', '期限日2', '期限日3']:
            deadline_val = row.get(deadline_col)
            if not pd.isna(deadline_val):
                if isinstance(deadline_val, (datetime, pd.Timestamp)):
                    deadline_date = deadline_val.date()
                    if deadline_date < today:
                        deadline_passed_count += 1
                        break
    
    # 満了年月日までの日数を計算して期限状況を集計
    days_30_count = 0
    days_60_count = 0
    days_90_count = 0
    expired_count = 0
    skill1_limit_count = 0
    
    for idx, row in df.iterrows():
        # 満了年月日までの残り日数を計算
        expiration_date = row.get('満了年月日')
        days_to_expiration = None
        if not pd.isna(expiration_date):
            if isinstance(expiration_date, (datetime, pd.Timestamp)):
                exp_date = expiration_date.date()
            else:
                exp_date = pd.to_datetime(expiration_date).date()
            days_to_expiration = (exp_date - today).days
        
        if days_to_expiration is not None:
            if days_to_expiration < 0:
                expired_count += 1
            else:
                setting1_val = parse_threshold_value(row.get('設定期限1'))
                setting2_val = parse_threshold_value(row.get('設定期限2'))
                setting3_val = parse_threshold_value(row.get('設定期限3'))
                thresholds = resolve_thresholds(setting1_val, setting2_val, setting3_val)
                level = determine_deadline_level(days_to_expiration, thresholds)

                if level == 'level1':
                    days_30_count += 1
                elif level == 'level2':
                    days_60_count += 1
                elif level == 'level3':
                    days_90_count += 1
        
        # 特定技能1号期限超過を計算
        manryo_days = row.get('満了日数')
        if not pd.isna(manryo_days) and (manryo_days + 184) > 1825:
            skill1_limit_count += 1
    
    # 期限状況を計算
    summary = {
        'filename': os.path.basename(current_file) if current_file else '未選択',
        'total': len(df),
        'deadline_passed': deadline_passed_count,
        'expired': expired_count,
        'days_30_count': days_30_count,
        'days_60_count': days_60_count,
        'days_90_count': days_90_count,
        'skill1_limit_count': skill1_limit_count,
    }

    summary['last_exported_at'] = last_export_info['timestamp'] if last_export_info else None
    summary['last_export_filename'] = last_export_info['download_name'] if last_export_info else None

    return jsonify(summary)


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """ファイルをアップロード"""
    global current_manager, current_file, last_export_info
    
    if 'file' not in request.files:
        return jsonify({'error': 'ファイルが選択されていません'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'ファイルが選択されていません'}), 400
    
    if file and file.filename.endswith('.xlsx'):
        filename = secure_filename(file.filename)
        # ファイル名が空になった場合のフォールバック
        if not filename or not filename.endswith('.xlsx'):
            filename = 'uploaded_file.xlsx'
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        print(f"[DEBUG] ファイル名: {filename}, パス: {filepath}")
        
        # ファイルを読み込み
        try:
            print(f"[DEBUG] ファイルをアップロード: {filepath}")
            current_manager = ResidenceStatusManager(filepath)
            
            print("[DEBUG] Excelファイルを読み込み中...")
            if not current_manager.load_excel():
                print("[ERROR] Excelファイルの読み込みに失敗")
                return jsonify({'error': 'ファイルの読み込みに失敗しました'}), 400
            
            print("[DEBUG] データを処理中...")
            if not current_manager.process_data():
                print("[ERROR] データの処理に失敗")
                return jsonify({'error': 'データの処理に失敗しました'}), 400
            
            current_file = filepath
            last_export_info = {'timestamp': None, 'download_name': None}
            print(f"[DEBUG] ファイル読み込み成功: {len(current_manager.df)}件")
            return jsonify({'success': True, 'filename': filename})
        except Exception as e:
            import traceback
            print(f"[ERROR] ファイルアップロードエラー: {str(e)}")
            print(traceback.format_exc())
            return jsonify({'error': f'エラー: {str(e)}'}), 400
    else:
        return jsonify({'error': 'Excelファイル(.xlsx)を選択してください'}), 400


@app.route('/api/export/alert')
def export_alert_list():
    """アラートリストをエクスポート"""
    global current_manager
    
    if current_manager is None or current_manager.df is None:
        return jsonify({'error': 'データが読み込まれていません'}), 400
    
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'アラートリスト.xlsx')
    
    if current_manager.export_alert_list(output_path, days_threshold=30):
        return send_file(output_path, as_attachment=True, download_name='アラートリスト.xlsx')
    else:
        return jsonify({'error': '期限日を超過しているデータはありません'}), 400


@app.route('/api/export/processed')
def export_processed_data():
    """処理済みデータをエクスポート"""
    global current_manager, current_file, last_export_info

    if current_manager is None or current_manager.df is None:
        return jsonify({'error': 'データが読み込まれていません'}), 400

    base_name = '在留資格管理'
    if current_file:
        base_name = os.path.splitext(os.path.basename(current_file))[0] or base_name

    if base_name.endswith('_processed'):
        output_filename = f"{base_name}.xlsx"
    else:
        output_filename = f"{base_name}_processed.xlsx"

    output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)

    if current_manager.save_processed_data(output_path):
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M')
        last_export_info = {
            'timestamp': timestamp,
            'download_name': output_filename
        }
        return send_file(output_path, as_attachment=True, download_name=output_filename)
    else:
        return jsonify({'error': 'エクスポートに失敗しました'}), 400


@app.route('/api/data/add', methods=['POST'])
def add_data():
    """データを追加"""
    global current_manager
    
    if current_manager is None or current_manager.df is None:
        return jsonify({'error': 'データが読み込まれていません'}), 400
    
    try:
        data = request.json
        
        # 日付を変換（時間情報を削除）
        kyoka_date = pd.to_datetime(data.get('許可年月日')).date() if data.get('許可年月日') else None
        manryo_date = pd.to_datetime(data.get('満了年月日')).date() if data.get('満了年月日') else None
        birth_date = pd.to_datetime(data.get('生年月日')).date() if data.get('生年月日') else None
        
        # 在留資格の正規化とカテゴリ判定
        zairyu_shikaku = data.get('在留資格', '')
        zairyu_shikaku = zairyu_shikaku.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
        is_skill1 = ('特定技能' in zairyu_shikaku) and ('1号' in zairyu_shikaku)
        is_skill2 = ('特定技能' in zairyu_shikaku) and ('2号' in zairyu_shikaku)
        is_gino = zairyu_shikaku.startswith('技能実習')

        # 既満了日数の入力検証・整形
        ki_manryo_input = data.get('既満了日数', '')
        if isinstance(ki_manryo_input, str):
            ki_manryo_input = ki_manryo_input.strip()
        ki_manryo_days = None
        if is_skill1:
            # 必須かつ0以上
            if ki_manryo_input == '' or ki_manryo_input is None:
                return jsonify({'error': '特定技能1号では既満了日数は必須です（0以上の数値）。'}), 400
            try:
                ki_manryo_days = int(ki_manryo_input)
            except:
                return jsonify({'error': '既満了日数は整数で入力してください。'}), 400
            if ki_manryo_days < 0:
                return jsonify({'error': '既満了日数は0以上で入力してください。'}), 400
        elif is_gino or is_skill2:
            # 必ず空白
            if ki_manryo_input not in ('', None):
                return jsonify({'error': '技能実習*・特定技能2号では既満了日数は空白にしてください。'}), 400
            ki_manryo_days = None
        else:
            # 任意（空白可）。数値が入っていれば取り込む
            if ki_manryo_input in ('', None):
                ki_manryo_days = None
            else:
                try:
                    ki_manryo_days = int(ki_manryo_input)
                except:
                    return jsonify({'error': '既満了日数は整数で入力してください。'}), 400
        
        # 設定期限を取得
        setting1 = int(data.get('設定期限1', 90))
        setting2 = int(data.get('設定期限2', 60))
        setting3 = int(data.get('設定期限3', 30))
        
        # 期限日1/2/3の計算式
        deadline1_formula = f'={manryo_date}-{setting1}' if manryo_date else None
        deadline2_formula = f'={manryo_date}-{setting2}' if manryo_date else None
        deadline3_formula = f'={manryo_date}-{setting3}' if manryo_date else None
        
        # 既存のDataFrameの列を取得
        existing_columns = current_manager.df.columns.tolist()
        
        # 新しい行を作成（既存の列に合わせる）
        new_row = {}
        
        # 基本情報
        if '担当者コード' in existing_columns:
            new_row['担当者コード'] = str(data.get('担当者コード', ''))
        if '氏名１' in existing_columns:
            new_row['氏名１'] = data.get('氏名１', '')
        if '氏名２' in existing_columns:
            new_row['氏名２'] = data.get('氏名２', '')
        if '在留資格' in existing_columns:
            new_row['在留資格'] = zairyu_shikaku
        if '国籍' in existing_columns:
            new_row['国籍'] = data.get('国籍', '')
        if '在留カード番号' in existing_columns:
            new_row['在留カード番号'] = data.get('在留カード番号', '')
        if '生年月日' in existing_columns:
            new_row['生年月日'] = birth_date
        if '期生' in existing_columns:
            # 期生はそのまま保存（「期」を追加しない）
            new_row['期生'] = data.get('期生', '')
        
        # 日付情報
        if '許可年月日' in existing_columns:
            new_row['許可年月日'] = kyoka_date
        if '満了年月日' in existing_columns:
            new_row['満了年月日'] = manryo_date
        if '既満了日数' in existing_columns:
            new_row['既満了日数'] = ki_manryo_days
        # 満了日数はバックエンドで一括再計算するためここでは設定しない
        
        # 設定期限
        if '設定期限1' in existing_columns:
            new_row['設定期限1'] = setting1
        if '設定期限2' in existing_columns:
            new_row['設定期限2'] = setting2
        if '設定期限3' in existing_columns:
            new_row['設定期限3'] = setting3
        
        # 期限日
        if '期限日1' in existing_columns:
            new_row['期限日1'] = deadline1_formula
        if '期限日2' in existing_columns:
            new_row['期限日2'] = deadline2_formula
        if '期限日3' in existing_columns:
            new_row['期限日3'] = deadline3_formula
        
        # 存在しない列にはNoneを設定
        for col in existing_columns:
            if col not in new_row:
                new_row[col] = None
        
        # DataFrameに追加
        new_df = pd.DataFrame([new_row])
        current_manager.df = pd.concat([current_manager.df, new_df], ignore_index=True)
        
        # 追加した行のインデックスを取得
        new_index = len(current_manager.df) - 1

        # 期限日1/2/3を実際の値に変換
        if manryo_date:
            from datetime import timedelta
            if '期限日1' in existing_columns:
                current_manager.df.at[new_index, '期限日1'] = manryo_date - timedelta(days=setting1)
            if '期限日2' in existing_columns:
                current_manager.df.at[new_index, '期限日2'] = manryo_date - timedelta(days=setting2)
            if '期限日3' in existing_columns:
                current_manager.df.at[new_index, '期限日3'] = manryo_date - timedelta(days=setting3)
        
        # データを再処理
        current_manager.process_data()
        
        print(f"[DEBUG] データ追加成功: {len(current_manager.df)}件")
        return jsonify({'success': True, 'message': 'データを追加しました'})
    except Exception as e:
        import traceback
        print(f"[ERROR] データ追加エラー: {str(e)}")
        print(traceback.format_exc())
        return jsonify({'error': f'追加エラー: {str(e)}'}), 400


@app.route('/api/data/update/<int:index>', methods=['PUT'])
def update_data(index):
    """データを更新"""
    global current_manager
    
    if current_manager is None or current_manager.df is None:
        return jsonify({'error': 'データが読み込まれていません'}), 400
    
    try:
        data = request.json
        
        # インデックスが有効か確認
        if index < 0 or index >= len(current_manager.df):
            return jsonify({'error': '無効なインデックスです'}), 400
        
        # データを更新
        if '担当者コード' in data:
            current_manager.df.at[index, '担当者コード'] = str(data['担当者コード'])  # 文字列として保存
        if '氏名１' in data:
            current_manager.df.at[index, '氏名１'] = data['氏名１']
        if '氏名２' in data:
            current_manager.df.at[index, '氏名２'] = data['氏名２']
        if '在留資格' in data:
            # 全角数字を半角に変換
            zairyu_shikaku = data['在留資格']
            zairyu_shikaku = zairyu_shikaku.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
            current_manager.df.at[index, '在留資格'] = zairyu_shikaku
        if '国籍' in data:
            current_manager.df.at[index, '国籍'] = data['国籍']
        if '在留カード番号' in data:
            current_manager.df.at[index, '在留カード番号'] = data['在留カード番号']
        if '生年月日' in data and data['生年月日']:
            try:
                birth_date = pd.to_datetime(data['生年月日'])
                # 1900年以降の日付のみ許可
                if birth_date.year >= 1900:
                    current_manager.df.at[index, '生年月日'] = birth_date.date()
            except:
                pass  # 無効な日付は無視
        if '期生' in data and data['期生']:
            # 期生はそのまま保存（「期」を追加しない）
            current_manager.df.at[index, '期生'] = data['期生']
        if '許可年月日' in data and data['許可年月日']:
            try:
                kyoka_date = pd.to_datetime(data['許可年月日'])
                if kyoka_date.year >= 1900:
                    current_manager.df.at[index, '許可年月日'] = kyoka_date.date()
            except:
                pass
        if '満了年月日' in data and data['満了年月日']:
            try:
                manryo_date = pd.to_datetime(data['満了年月日'])
                if manryo_date.year >= 1900:
                    current_manager.df.at[index, '満了年月日'] = manryo_date.date()
            except:
                pass
        if '既満了日数' in data:
            ki_manryo_input = data['既満了日数']
            if isinstance(ki_manryo_input, str):
                ki_manryo_input = ki_manryo_input.strip()

            # 在留資格に基づく検証
            zairyu_shikaku_now = str(current_manager.df.at[index, '在留資格'])
            zairyu_shikaku_now = zairyu_shikaku_now.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
            is_skill1_now = ('特定技能' in zairyu_shikaku_now) and ('1号' in zairyu_shikaku_now)
            is_skill2_now = ('特定技能' in zairyu_shikaku_now) and ('2号' in zairyu_shikaku_now)
            is_gino_now = zairyu_shikaku_now.startswith('技能実習')

            if is_skill1_now:
                if ki_manryo_input == '' or ki_manryo_input is None:
                    return jsonify({'error': '特定技能1号では既満了日数は必須です（0以上の数値）。'}), 400
                try:
                    val = int(ki_manryo_input)
                except:
                    return jsonify({'error': '既満了日数は整数で入力してください。'}), 400
                if val < 0:
                    return jsonify({'error': '既満了日数は0以上で入力してください。'}), 400
                current_manager.df.at[index, '既満了日数'] = val
            elif is_gino_now or is_skill2_now:
                if ki_manryo_input not in ('', None):
                    return jsonify({'error': '技能実習*・特定技能2号では既満了日数は空白にしてください。'}), 400
                current_manager.df.at[index, '既満了日数'] = None
            else:
                if ki_manryo_input in ('', None):
                    current_manager.df.at[index, '既満了日数'] = None
                else:
                    try:
                        current_manager.df.at[index, '既満了日数'] = int(ki_manryo_input)
                    except:
                        return jsonify({'error': '既満了日数は整数で入力してください。'}), 400
        if '設定期限1' in data:
            current_manager.df.at[index, '設定期限1'] = int(data['設定期限1'])
        if '設定期限2' in data:
            current_manager.df.at[index, '設定期限2'] = int(data['設定期限2'])
        if '設定期限3' in data:
            current_manager.df.at[index, '設定期限3'] = int(data['設定期限3'])
        
        # 満了日数はバックエンドで再計算
        
        # 期限日1/2/3を実際の値で設定
        manryo_date_updated = current_manager.df.at[index, '満了年月日']
        if not pd.isna(manryo_date_updated):
            from datetime import timedelta
            manryo_date_dt = pd.to_datetime(manryo_date_updated)
            
            setting1 = current_manager.df.at[index, '設定期限1']
            setting2 = current_manager.df.at[index, '設定期限2']
            setting3 = current_manager.df.at[index, '設定期限3']
            
            if not pd.isna(setting1):
                current_manager.df.at[index, '期限日1'] = (manryo_date_dt - timedelta(days=int(setting1))).date()
            if not pd.isna(setting2):
                current_manager.df.at[index, '期限日2'] = (manryo_date_dt - timedelta(days=int(setting2))).date()
            if not pd.isna(setting3):
                current_manager.df.at[index, '期限日3'] = (manryo_date_dt - timedelta(days=int(setting3))).date()
        
        # データを再処理
        current_manager.process_data()
        
        return jsonify({'success': True, 'message': 'データを更新しました'})
    except Exception as e:
        return jsonify({'error': f'更新エラー: {str(e)}'}), 400


@app.route('/api/data/delete/<int:index>', methods=['DELETE'])
def delete_data(index):
    """データを削除"""
    global current_manager, current_file, last_export_info

    if current_manager is None or current_manager.df is None:
        return jsonify({'error': 'データが読み込まれていません'}), 400

    try:
        if index < 0 or index >= len(current_manager.df):
            return jsonify({'error': '無効なインデックスです'}), 400

        # 行を削除
        current_manager.df = current_manager.df.drop(index).reset_index(drop=True)

        # 変更を元ファイルに保存
        if current_file:
            current_manager.save_processed_data(current_file)
            last_export_info = {'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M'), 'download_name': os.path.basename(current_file)}

        return jsonify({'success': True, 'message': 'データを削除しました'})
    except Exception as e:
        return jsonify({'error': f'削除エラー: {str(e)}'}), 400


@app.route('/api/calendar')
def get_calendar_data():
    """カレンダーデータを取得するAPI"""
    global current_manager, current_file
    
    if current_manager is None or current_manager.df is None:
        return jsonify({'error': 'データが読み込まれていません'}), 400
    
    year = int(request.args.get('year'))
    month = int(request.args.get('month'))
    
    df = current_manager.df
    calendar_data = {}
    
    for idx, row in df.iterrows():
        for i in range(1, 4):
            deadline_col = f'期限日{i}'
            if deadline_col in df.columns and not pd.isna(row[deadline_col]):
                deadline_date = row[deadline_col]
                if isinstance(deadline_date, pd.Timestamp):
                    deadline_date = deadline_date.date()
                elif isinstance(deadline_date, str):
                    deadline_date = pd.to_datetime(deadline_date).date()
                
                if deadline_date.year == year and deadline_date.month == month:
                    date_str = deadline_date.strftime('%Y-%m-%d')
                    if date_str not in calendar_data:
                        calendar_data[date_str] = {'deadline1': [], 'deadline2': [], 'deadline3': []}
                    
                    person = {
                        '担当者コード': str(row.get('担当者コード', '-')),
                        '氏名２': str(row.get('氏名２', '-')),
                        '在留資格': str(row.get('在留資格', '-'))
                    }
                    calendar_data[date_str][f'deadline{i}'].append(person)
    
    return jsonify({'calendar_data': calendar_data})


if __name__ == '__main__':
    print("\n" + "=" * 80)
    print("在留資格管理システム - Web版")
    print("=" * 80)
    print("\nブラウザで以下のURLにアクセスしてください:")
    print("http://localhost:5000")
    print("\n終了するには Ctrl+C を押してください")
    print("=" * 80 + "\n")
    
    app.run(debug=True, host='localhost', port=5000)
