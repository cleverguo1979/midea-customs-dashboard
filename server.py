#!/usr/bin/env python3
"""
美的业务关务看板 - 后端 API 服务器
支持本地局域网访问和云部署（Render/Railway/Fly.io 等）
"""
import os
import sys
import json
import socket
import tempfile
from datetime import datetime, timedelta

try:
    from flask import Flask, request, jsonify, send_from_directory
except ImportError:
    print("请先安装 Flask: pip3 install flask")
    sys.exit(1)

import openpyxl

app = Flask(__name__, static_folder=None)

# CORS — allow access from GitHub Pages and any origin
@app.after_request
def add_cors_headers(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    return response

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, 'data.json')
PORT = int(os.environ.get('PORT', 8888))


def excel_serial_to_date(serial):
    base = datetime(1899, 12, 30)
    return base + timedelta(days=int(serial))


def parse_workbook(filepath):
    """Parse Excel file, return structured data dict by year."""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    sheet_names = wb.sheetnames

    daily_sheet_names = [n for n in sheet_names if '日报关情况' in n]
    abnormal_sheet_names = [n for n in sheet_names if '异常跟踪表' in n]
    mapping_sheet_name = next((n for n in sheet_names if n == 'Sheet3'), None)

    if not daily_sheet_names:
        raise ValueError('未找到"日报关情况"工作表')

    all_years = {}

    # Parse daily sheets
    for sname in daily_sheet_names:
        ws = wb[sname]
        result = []
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
            if not row[0] and not row[1]:
                continue
            try:
                serial = int(row[0])
            except (ValueError, TypeError):
                continue
            if serial < 40000:
                continue
            declared = int(row[1] or 0)
            released = int(row[2] or 0)
            if declared == 0 and released == 0:
                continue
            date = excel_serial_to_date(serial)
            result.append({
                'date': date.strftime('%Y-%m-%d'),
                'year': str(date.year),
                'month': date.month,
                'day': date.day,
                'totalDeclared': declared,
                'totalReleased': released,
                'reviewCompleted': int(row[3] or 0),
                'unclearedReason': str(row[4] or '').strip()
            })
        if result:
            year = str(result[0]['year'])
            if year not in all_years:
                all_years[year] = {'dailyData': [], 'abnormalRecords': [], 'mapping': []}
            all_years[year]['dailyData'].extend(result)

    # Parse abnormal sheets
    for sname in abnormal_sheet_names:
        ws = wb[sname]
        result = []
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
            if not row[0]:
                continue
            try:
                serial = int(row[1])
                date_str = excel_serial_to_date(serial).strftime('%Y-%m-%d') if serial > 40000 else str(row[1] or '').strip()
            except (ValueError, TypeError):
                date_str = str(row[1] or '').strip()
            result.append({
                'seq': row[0],
                'date': date_str,
                'category': str(row[2] or '').strip(),
                'bizUnit': str(row[3] or '').strip(),
                'company': str(row[4] or '').strip(),
                'importExport': str(row[5] or '').strip(),
                'customsNo': str(row[6] or '').strip(),
                'bolNo': str(row[7] or '').strip(),
                'containerNo': str(row[8] or '').strip(),
                'description': str(row[9] or '').strip(),
                'responsible': str(row[10] or '').strip(),
                'fee': str(row[11] or '').strip(),
                'progress': str(row[12] or '').strip(),
                'status': str(row[13] or '').strip(),
                'agent': str(row[14] or '').strip()
            })
        if result:
            import re
            ym = re.search(r'(\d{4})', sname)
            year = ym.group(1) if ym else '2026'
            if year not in all_years:
                all_years[year] = {'dailyData': [], 'abnormalRecords': [], 'mapping': []}
            all_years[year]['abnormalRecords'].extend(result)

    # Parse mapping
    mapping = []
    if mapping_sheet_name:
        ws = wb[mapping_sheet_name]
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
            if row[0]:
                mapping.append({
                    'company': str(row[0]).strip(),
                    'bizUnit': str(row[1]).strip(),
                    'rep': str(row[2]).strip()
                })
    for year in all_years:
        all_years[year]['mapping'] = mapping

    return all_years


def load_data():
    """Load stored data from JSON file."""
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {'2025': None, '2026': None}


def save_data(data):
    """Save data to JSON file."""
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# === Routes ===

@app.route('/')
def index():
    return send_from_directory(BASE_DIR, 'index.html')


@app.route('/<path:filename>')
def static_files(filename):
    if filename.startswith('api/'):
        return jsonify({'error': 'Not found'}), 404
    return send_from_directory(BASE_DIR, filename)


@app.route('/api/data', methods=['GET'])
def get_data():
    """Return the current stored data."""
    data = load_data()
    return jsonify(data)


@app.route('/api/upload', methods=['POST'])
def upload():
    """Accept Excel file upload, parse, and store."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Empty filename'}), 400

    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Invalid file format, expected .xlsx or .xls'}), 400

    try:
        # Save to temp file and parse
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        file.save(tmp.name)
        tmp.close()

        parsed = parse_workbook(tmp.name)
        os.unlink(tmp.name)

        # Full replace: merge with existing but new data overwrites
        current = load_data()
        for year in ['2025', '2026']:
            if year in parsed:
                current[year] = parsed[year]
            elif year not in current:
                current[year] = None

        save_data(current)

        total_days = sum(len(v['dailyData']) for v in current.values() if v and v.get('dailyData'))
        total_ab = sum(len(v['abnormalRecords']) for v in current.values() if v and v.get('abnormalRecords'))
        years_found = [y for y in ['2025', '2026'] if current.get(y) and current[y].get('dailyData')]

        return jsonify({
            'success': True,
            'years': years_found,
            'totalDays': total_days,
            'totalAbnormal': total_ab,
            'message': f'数据上传成功！{",".join(years_found)} 年，{total_days} 天记录，{total_ab} 条异常'
        })

    except Exception as e:
        return jsonify({'error': f'文件解析失败: {str(e)}'}), 500


if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 8888))
    print(f"Starting server on port {port}")
    app.run(host='0.0.0.0', port=port, debug=False)
