# app.py
import os
import threading
import uuid
import time
import shutil
from flask import Flask, render_template, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename

import pdfplumber
import PyPDF2
import xlsxwriter

# ----------------- 绝对路径根 -----------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__)
app.config['UPLOAD_ROOT']   = os.path.join(BASE_DIR, 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(BASE_DIR, 'outputs')
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200MB

os.makedirs(app.config['UPLOAD_ROOT'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

tasks = {}
tasks_lock = threading.Lock()

# ----------------- PDF 修复 CropBox（保留你的逻辑） -----------------
def fix_cropbox(pdf_path, output_path):
    with open(pdf_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        writer = PyPDF2.PdfWriter()
        for page in reader.pages:
            if "/CropBox" not in page:
                page.cropbox = page.mediabox
            writer.add_page(page)
    with open(output_path, "wb") as f:
        writer.write(f)

# ----------------- 原版 PDF 解析逻辑（保留你的实现） -----------------
def get_raw_info(filesName):
    i = 0
    rawInfo = []
    print("开始提取数据")
    while i < len(filesName):
        fixedFileFullName = os.path.join("fixed", filesName[i])  # 这里按你的原路径约定
        with pdfplumber.open(fixedFileFullName) as pdf:
            oneRawInfo = []
            table_settings = {"vertical_strategy": "lines", "horizontal_strategy": "text"}
            for page in pdf.pages:
                tables = page.extract_tables(table_settings)
                for table in tables:
                    for row in table:
                        if row != ['', '', '', '', '', '', '', '', '', '', '', '']:
                            if row != ['', '', '', '', '', '', '', '', '', '', '']:
                                if row != ['', '', '', '', '', '', '', '', '', '']:
                                    if row != ['', '', '', '', '', '', '', '', '']:
                                        if row != ['', '', '', '', '', '', '', '']:
                                            if row != ['', '', '', '', '', '', '', '']:
                                                if row != ['', '', '', '', '', '', '']:
                                                    if row != ['', '', '', '', '', '']:
                                                        if row != ['', '', '', '', '']:
                                                            if row != ['', '', '', '']:
                                                                if row != ['', '', '']:
                                                                    if row != ['', '']:
                                                                        if row != ['']:
                                                                            infoInLine = [row]
                                                                            oneRawInfo += infoInLine
        rawInfo += [oneRawInfo]
        i += 1
        print(f'\r已提取 {i}/{len(filesName)}', end='', flush=True)
    print('\n提取完毕')
    return rawInfo

def Preprocessing_Data(rawData, filesName):
    print("开始处理数据")
    Processed_Data = []
    i = 0
    while i < len(rawData):
        one_PDF_data = rawData[i]
        Form_Data = []

        if one_PDF_data[0][0] == 'To:':
            POSN = one_PDF_data.index(['No','Service No', '', 'Date', 'Qty', '', '', ''])
            PO_No = one_PDF_data[0][1][12:]
            PO_Date = one_PDF_data[1][1][10:]
            pn = one_PDF_data[2][1].index(':') + 2
            Contract_No = one_PDF_data[2][1][pn:]
            Payment_Terms = one_PDF_data[3][1][16:]
            Project_Cost_Center = one_PDF_data[5][1][22:] + " "
            if one_PDF_data[6][0] == 'MALAYSIA':
                Project_Cost_Center += one_PDF_data[6][1]
            else:
                Project_Cost_Center += one_PDF_data[6][0]
            Tracking_No = one_PDF_data[7][1][14:]
        else:
            POSN = one_PDF_data.index(['Service No', '', 'Date', 'Qty', '', '', ''])
            PO_No = one_PDF_data[0][0][12:]
            PO_Date = one_PDF_data[1][0][10:]
            pn = one_PDF_data[2][0].index(':') + 2
            Contract_No = one_PDF_data[2][0][pn:]
            Payment_Terms = one_PDF_data[3][0][16:]
            Project_Cost_Center = one_PDF_data[5][0][22:] + " "
            if one_PDF_data[6][0] == 'MALAYSIA':
                Project_Cost_Center += one_PDF_data[6][1]
            else:
                Project_Cost_Center += one_PDF_data[6][0]
            Tracking_No = one_PDF_data[7][0][14:]

        Service_No, Description, Delivery_Date, Order_Qty, UoM, Unit_Price, Total_Price = [], [], [], [], [], [], []
        while POSN < len(one_PDF_data):
            if len(one_PDF_data[POSN]) > 1:
                if one_PDF_data[POSN][0][:3] == '000':
                    Service_No.append(one_PDF_data[POSN][0])
                    pod = one_PDF_data[POSN][1]
                    a = 1
                    while one_PDF_data[POSN + a + 1][0] != '':
                        if len(one_PDF_data[POSN + a]) > 1:
                            if one_PDF_data[POSN + a][1] != 'Non-SST Registered Supplier Purchases 0%' and one_PDF_data[POSN + a][1] != '':
                                pod += one_PDF_data[POSN + a][1]
                        a += 1
                    Description.append(pod)
                    Delivery_Date.append(one_PDF_data[POSN][2])
                    Order_Qty.append(one_PDF_data[POSN][3])
                    UoM.append(one_PDF_data[POSN][4])
                    Unit_Price.append(one_PDF_data[POSN][5])
                    Total_Price.append(one_PDF_data[POSN + 1][-1])
            POSN += 1

        Form_Data += [Service_No, Description, Delivery_Date, Order_Qty, UoM, Unit_Price, Total_Price]
        infoInLine = [PO_No, PO_Date, Contract_No, Payment_Terms, Project_Cost_Center, Tracking_No, Form_Data]
        Processed_Data.append(infoInLine)
        i += 1
    print("处理完毕")
    return Processed_Data

# ----------------- 任务状态更新 -----------------
def _update_task(task_id, percent=None, message=None, finished=False, files=None, error=False):
    with tasks_lock:
        t = tasks.get(task_id, {})
        if percent is not None:
            t['percent'] = int(max(0, min(100, percent)))
        if message is not None:
            t['message'] = message
        if finished:
            t['finished'] = True
        if files is not None:
            t['files'] = files
        if error:
            t['error'] = True
        tasks[task_id] = t

# ----------------- 工作线程（核心流程） -----------------
def worker_task(task_id, pdf_byte_list, filename_list):
    try:
        _update_task(task_id, percent=0, message="准备处理中 / Preparing...")
        # 统一在 BASE_DIR 内操作，避免工作目录差异
        input_dir = os.path.join(BASE_DIR, 'input')
        fixed_dir = os.path.join(BASE_DIR, 'fixed')
        os.makedirs(input_dir, exist_ok=True)
        os.makedirs(fixed_dir, exist_ok=True)

        # 清空残留
        for f in os.listdir(input_dir):
            try: os.remove(os.path.join(input_dir, f))
            except: pass
        for f in os.listdir(fixed_dir):
            try: os.remove(os.path.join(fixed_dir, f))
            except: pass

        # 保存上传的 PDF
        for i, (bts, fname) in enumerate(zip(pdf_byte_list, filename_list), start=1):
            path = os.path.join(input_dir, fname)
            with open(path, "wb") as f:
                f.write(bts)
            _update_task(task_id, percent=int(5 + 5*i/max(1,len(pdf_byte_list))),
                         message=f'已保存 {i}/{len(pdf_byte_list)} / Saved {i}/{len(pdf_byte_list)}')

        # 修复 PDF 到 fixed/
        filesName_in = sorted(os.listdir(input_dir))
        for i, fname in enumerate(filesName_in, start=1):
            src = os.path.join(input_dir, fname)
            dst = os.path.join(fixed_dir, f"fixed_{fname}")
            fix_cropbox(src, dst)
            _update_task(task_id, percent=int(10 + 20*i/max(1,len(filesName_in))),
                         message=f'修复文件中: {i}/{len(filesName_in)} / Fixing {i}/{len(filesName_in)}')

        # 调用你的原版读取与处理
        os.chdir(BASE_DIR)  # 确保相对路径 "fixed/xxx" 能命中
        filesName = sorted(os.listdir('fixed'))
        print("fixed 下文件：", filesName)

        raw_data = get_raw_info(filesName)
        print("raw_data 条数：", len(raw_data))
        if not raw_data:
            raise RuntimeError("raw_data 为空，未提取到任何表格")

        all_data = Preprocessing_Data(raw_data, filesName)
        print("all_data 条数：", len(all_data))
        if not all_data:
            raise RuntimeError("all_data 为空，未形成任何记录")

        # 生成 Excel（完全保留你的写法）
        _update_task(task_id, percent=85, message="生成Excel文件中 / Generating Excel...")
        xlsxName = f"PO_Results_{time.strftime('%Y%m%d%H%M%S')}_{task_id[:8]}.xlsx"
        xlsxName = xlsxName.replace(" ", "_")
        out_path = os.path.join(app.config['OUTPUT_FOLDER'], xlsxName)
        print("准备写入 Excel：", out_path)

        workbook = xlsxwriter.Workbook(out_path)
        worksheet = workbook.add_worksheet('PO Info')
        worksheet.set_column('A:A', 12)
        worksheet.set_column('B:B', 12)
        worksheet.set_column('C:C', 12)
        worksheet.set_column('D:D', 25)
        worksheet.set_column('E:E', 55)
        worksheet.set_column('F:F', 12)
        worksheet.set_column('G:G', 12)
        worksheet.set_column('H:H', 60)
        worksheet.set_column('I:I', 15)
        worksheet.set_column('J:J', 10)
        worksheet.set_column('K:K', 4)
        worksheet.set_column('L:L', 12)
        worksheet.set_column('M:M', 12)

        # 表头
        worksheet.write(0,0,"PO Number")
        worksheet.write(0,1,"PO Date")
        worksheet.write(0,2,"Contract No")
        worksheet.write(0,3,"Payment Terms")
        worksheet.write(0,4,"Project/Cost Center")
        worksheet.write(0,5,"Tracking No")
        worksheet.write(0,6,"Service No")
        worksheet.write(0,7,"Description")
        worksheet.write(0,8,"Delivery Date")
        worksheet.write(0,9,"Order Qty")
        worksheet.write(0,10,"UoM")
        worksheet.write(0,11,"Unit Price")
        worksheet.write(0,12,"Total Price")

        # 数据写入（保持你的原逻辑）
        i = 1
        a = 0
        while a < len(all_data):
            worksheet.write(i,0,all_data[a][0])
            worksheet.write(i,1,all_data[a][1])
            worksheet.write(i,2,all_data[a][2])
            worksheet.write(i,3,all_data[a][3])
            worksheet.write(i,4,all_data[a][4])
            worksheet.write(i,5,all_data[a][5])
            q = 0
            c = i
            while q < len(all_data[a][6][0]):
                worksheet.write(c,6,all_data[a][6][0][q])
                worksheet.write(c,7,all_data[a][6][1][q])
                worksheet.write(c,8,all_data[a][6][2][q])
                worksheet.write(c,9,all_data[a][6][3][q])
                worksheet.write(c,10,all_data[a][6][4][q])
                worksheet.write(c,11,all_data[a][6][5][q])
                worksheet.write(c,12,all_data[a][6][6][q])
                q += 1
                c += 1
            i = c
            a += 1

        workbook.close()
        print("Excel 已生成：", out_path, "存在？", os.path.exists(out_path))

        _update_task(task_id, percent=100, message="处理完成! / Done", finished=True, files=[os.path.basename(out_path)])

        # 清理 fixed
        try:
            shutil.rmtree(fixed_dir)
        except Exception as _:
            pass

    except Exception as e:
        print("worker_task 异常：", repr(e))
        _update_task(task_id, percent=100, message=f"处理 PDF 出错: {str(e)}", finished=True, error=True)

# ----------------- Flask 路由 -----------------
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/start_process", methods=["POST"])
def start_process():
    if 'pdfs' not in request.files:
        return jsonify({'error': '请上传 PDF 文件'}), 400
    pdf_files = request.files.getlist('pdfs')
    pdf_byte_list, filename_list = [], []
    for f in pdf_files:
        fname = secure_filename(f.filename)
        if fname.lower().endswith('.pdf'):
            pdf_byte_list.append(f.read())
            filename_list.append(fname)
    if not pdf_byte_list:
        return jsonify({'error': '未提供合法的 PDF 文件'}), 400

    task_id = str(uuid.uuid4())
    with tasks_lock:
        tasks[task_id] = {'percent':0,'message':'任务已创建 / Task created','finished':False,'files':[],'error':False}

    threading.Thread(target=worker_task, args=(task_id, pdf_byte_list, filename_list), daemon=True).start()
    return jsonify({'task_id': task_id})

@app.route("/progress")
def progress():
    task_id = request.args.get('task_id')
    if not task_id:
        return jsonify({'error':'缺少任务ID'}),400
    with tasks_lock:
        data = tasks.get(task_id)
        if not data:
            return jsonify({'error':'任务不存在或已过期'}),404
        return jsonify(data)

# 重要：用 <path:filename> 并且用绝对输出目录
@app.route("/download_file/<path:filename>")
def download_file(filename):
    filepath = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    print("尝试下载文件：", filepath, "存在？", os.path.exists(filepath))
    if not os.path.exists(filepath):
        return jsonify({'error':'文件不存在'}),404
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)

if __name__ == "__main__":
    # 运行时确保当前工作目录是 BASE_DIR
    os.chdir(BASE_DIR)
    print("BASE_DIR:", BASE_DIR)
    print("OUTPUT_FOLDER:", app.config['OUTPUT_FOLDER'])
    app.run(debug=True, threaded=True)
