# -*- coding: utf-8 -*-
"""
GPA 成绩分析 Web 应用
拖拽Excel文件自动生成HTML报告
"""

import os
import uuid
import glob as glob_module
from flask import Flask, render_template, request, send_file, jsonify, make_response
from werkzeug.utils import secure_filename

# 获取项目根目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.path.join(BASE_DIR, 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.secret_key = 'gpa-report-secret-key'

# 静态文件目录
STATIC_REPORTS_DIR = os.path.join(BASE_DIR, 'static', 'reports')

# 确保目录存在
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(STATIC_REPORTS_DIR, exist_ok=True)

# 导入分析模块
from utils.analyzer import GPAAanalyzer

# 允许的文件扩展名
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    """首页"""
    response = make_response(render_template('index.html'))
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

@app.route('/upload', methods=['POST'])
def upload_file():
    """处理文件上传并生成报告"""
    if 'file' not in request.files:
        return jsonify({'error': '没有文件'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '未选择文件'}), 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        if '批量学期成绩下载' not in filename:
            filename = f"批量学期成绩下载_{filename}"
        
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        try:
            analyzer = GPAAanalyzer()
            result = analyzer.process_file(filepath)
            
            return jsonify({
                'success': True,
                'report': result['html'],
                'stats': result['stats']
            })
        except Exception as e:
            import traceback
            return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500
        finally:
            if os.path.exists(filepath):
                os.remove(filepath)
    
    return jsonify({'error': '不支持的文件类型'}), 400

@app.route('/batch', methods=['POST'])
def batch_upload():
    """批量处理多个文件"""
    if 'files' not in request.files:
        return jsonify({'error': '没有文件'}), 400
    
    files = request.files.getlist('files')
    if not files or all(f.filename == '' for f in files):
        return jsonify({'error': '未选择文件'}), 400
    
    # 获取语言设置
    lang = request.form.get('lang', 'zh')
    
    temp_dir = os.path.join(app.config['UPLOAD_FOLDER'], str(uuid.uuid4()))
    os.makedirs(temp_dir, exist_ok=True)
    
    saved_files = []
    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            if '批量学期成绩下载' not in filename:
                filename = f"批量学期成绩下载_{filename}"
            filepath = os.path.join(temp_dir, filename)
            file.save(filepath)
            saved_files.append(filepath)
    
    try:
        analyzer = GPAAanalyzer()
        result = analyzer.process_multiple_files(saved_files, lang=lang)
        
        import time
        report_filename = f"report_{int(time.time())}.html"
        report_path = os.path.join(STATIC_REPORTS_DIR, report_filename)
        
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(result['html'])
        
        return jsonify({
            'success': True,
            'report_url': '/static/reports/' + report_filename,
            'stats': result['stats']
        })
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500
    finally:
        import shutil
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

@app.route('/export-pdf', methods=['POST'])
def export_pdf():
    """导出PDF报告（全部数据）"""
    if 'files' not in request.files:
        return jsonify({'error': '没有文件'}), 400
    
    files = request.files.getlist('files')
    if not files or all(f.filename == '' for f in files):
        return jsonify({'error': '未选择文件'}), 400
    
    lang = request.form.get('lang', 'zh')
    
    temp_dir = os.path.join(app.config['UPLOAD_FOLDER'], str(uuid.uuid4()))
    os.makedirs(temp_dir, exist_ok=True)
    
    saved_files = []
    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            if '批量学期成绩下载' not in filename:
                filename = f"批量学期成绩下载_{filename}"
            filepath = os.path.join(temp_dir, filename)
            file.save(filepath)
            saved_files.append(filepath)
    
    try:
        import time
        analyzer = GPAAanalyzer()
        
        # 生成临时PDF到项目目录
        temp_pdf = analyzer.generate_pdf(saved_files, lang=lang)
        
        # 移动到static/reports目录
        pdf_filename = f"report_{int(time.time())}.pdf"
        pdf_path = os.path.join(STATIC_REPORTS_DIR, pdf_filename)
        os.rename(temp_pdf, pdf_path)
        
        return jsonify({
            'success': True,
            'download_url': '/static/reports/' + pdf_filename
        })
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500
    finally:
        import shutil
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

@app.route('/export-pdf-filtered', methods=['POST'])
def export_pdf_filtered():
    """导出筛选后的PDF报告"""
    data = request.json
    if not data:
        return jsonify({'error': '没有数据'}), 400
    
    lang = data.get('lang', 'zh')
    rows = data.get('rows', [])
    dates = data.get('dates', [])
    classes = data.get('classes', [])
    
    if not rows or not dates:
        return jsonify({'error': '没有有效数据'}), 400
    
    try:
        import time
        analyzer = GPAAanalyzer()
        
        # 直接从传入的数据生成PDF
        pdf_path = analyzer.generate_pdf_from_data(rows, classes, dates, lang=lang)
        
        # 移动到static/reports目录
        pdf_filename = f"report_{int(time.time())}.pdf"
        final_path = os.path.join(STATIC_REPORTS_DIR, pdf_filename)
        os.rename(pdf_path, final_path)
        
        return jsonify({
            'success': True,
            'download_url': '/static/reports/' + pdf_filename
        })
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

if __name__ == '__main__':
    print(f"""
🚀 GPA 成绩分析系统启动中...
   
   访问: http://localhost:5000
   项目目录: {BASE_DIR}
   报告目录: {STATIC_REPORTS_DIR}
""")
    app.run(debug=True, host='0.0.0.0', port=5000)
