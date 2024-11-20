from flask import Flask, request, send_file, render_template, url_for
import os
from werkzeug.utils import secure_filename
import pandas as pd
import openpyxl
import datetime
from typing import Dict, Any
from concurrent.futures import ThreadPoolExecutor
import queue
import threading
import uuid

class ExcelProcessor:
    def __init__(self, fba_path: str, product_info_path: str):
        self.fba_shipment = None
        self.product_summary = None
        self.packing_list = None
        self.shipment_info = {}
        self.msku_map = {}
        self.product_info_map = {}
        self.packing_info_map = {}
        
        # 读取Excel文件
        self.read_excel_files(fba_path, product_info_path)
        
    def read_excel_files(self, fba_path: str, product_info_path: str):
        """读取Excel文件"""
        self.fba_shipment = pd.read_excel(fba_path)
        excel_file = pd.ExcelFile(product_info_path)
        self.product_summary = pd.read_excel(excel_file, sheet_name="品号汇总")
        self.packing_list = pd.read_excel(excel_file, sheet_name="装箱清单")
        
    def process_fba_shipment(self):
        """处理FBA货件信息"""
        # 获取基本信息
        first_row = self.fba_shipment.iloc[0]
        self.shipment_info = {
            '货件单号': first_row['货件单号'],
            '店铺': first_row['店铺'],
            '国家': first_row['国家'],
            '创建日期': first_row['创建时间'],
            '物流中心编码': first_row['物流中心编码']
        }
        
        # 处理MSKU信息
        for _, row in self.fba_shipment.iterrows():
            msku = row['MSKU']
            product_name = row['品名']
            
            # 解析品名信息
            parts = product_name.split('*')
            model = parts[0]
            product_info = parts[2].split('/')
            
            self.msku_map[msku] = {
                '型号': model,
                '颜色': product_info[3],
                '规格': product_info[1],
                '建单数量': row['申报量'],
                'FNSKU': row['FNSKU']
            }
            
    def process_product_summary(self):
        """处理品号汇总信息"""
        for _, row in self.product_summary.iterrows():
            self.product_info_map[row['乌托邦新品号']] = {
                '客户型号': row['客户型号'],
                '颜色': row['颜色'],
                '描述': row['描述'],
                '品牌': row['品牌']
            }
            
    def process_packing_list(self):
        """处理装箱清单信息"""
        for _, row in self.packing_list.iterrows():
            self.packing_info_map[row['乌托邦新品号']] = {
                '普通装箱数': row['普通箱箱数(PCS)'],
                '是否危险品': True if row['危险品'] == '危险品' else False
            }
            self.packing_info_map[row['客户型号']] = {
                '普通装箱数': row['普通箱箱数(PCS)'],
                '是否危险品': True if row['危险品'] == '危险品' else False
            }
            
    def generate_result(self) -> pd.DataFrame:
        """生成最终结果"""
        result_data = []
        
        for msku, msku_info in self.msku_map.items():
            models = msku_info['型号'].split('/')
            # 欧洲地区：'德国', '法国', '意大利', '西班牙', '英国', '荷兰', '比利时', '瑞典', '波兰'
            
            # 根据店铺名称确定品牌
            if 'charmast'.lower() in self.shipment_info['店铺'].lower():
                target_brand = "超麦"
            elif 'chenying'.lower() in self.shipment_info['店铺'].lower():
                target_brand = "晨樱" 
            elif 'veger'.lower() in self.shipment_info['店铺'].lower():
                target_brand = "艾美柯"
            elif 'vrurc'.lower() in self.shipment_info['店铺'].lower():
                target_brand = "创立嘉城"
            elif 'GH'.lower() in self.shipment_info['店铺'].lower():
                target_brand = "谷和"

            for model in models:
                # target_brand = "超麦" if self.shipment_info['国家'] == "美国" else "晨樱"
                
                # 在product_info_map中查找对应的产品信息
                product_info = None
                for _, info in self.product_info_map.items():
                    if model == _ and info['品牌'].find(target_brand) != -1:
                    # if model == _ :
                        product_info = info
                        break
                
                if product_info is None:
                    continue
                
                # 获取装箱信息
                packing_info = self.packing_info_map.get(model, {})
                if packing_info == {}:
                    packing_info = self.packing_info_map.get(product_info.get('客户型号'), {})
                if packing_info == {}:
                    print(f"未找到对应的装箱信息：{product_info['客户型号']}, 写入默认值")
                    packing_info = {
                        '普通装箱数': 40,
                        "危险品": False
                    }
                boxes_count = int((msku_info['建单数量'] / packing_info.get('普通装箱数', 1) + 0.99999))
                
                result_row = {
                    '账号': target_brand,
                    '货件日期': self.shipment_info['创建日期'],
                    '国家': self.shipment_info['国家'],
                    '货件编码': self.shipment_info['货件单号'],
                    '纸箱编号': '',
                    '产品型号': f"{product_info['客户型号']}{product_info['颜色']}",
                    '品号': model,
                    '产品规格': msku_info['规格'],
                    '建单数量': msku_info['建单数量'],
                    '库存': '',
                    '待生产': '',
                    '件数/箱': boxes_count,
                    '单票合计/箱': '',
                    '箱规': '',
                    '装箱规格个/箱': packing_info.get('普通装箱数', 0),
                    '物流渠道': '',
                    '货件特殊说明': '',
                    '物流中心编码': self.shipment_info['物流中心编码'],
                    '报关单价': '',
                    '平台售价': '',
                    '备注': '',
                    '透明计划标签（MSKU）': msku,
                    '标签(FNSKU)': msku_info['FNSKU'],
                    '外箱标签': '',
                    '班级': ''
                }
                result_data.append(result_row)
        
        # 计算单票合计/箱
        total_boxes = sum(row['件数/箱'] for row in result_data)
        for row in result_data:
            row['单票合计/箱'] = total_boxes

        # 计算每个MSKU的纸箱编号
        start_box_num = 1
        for row in result_data:
            boxes = int(row['件数/箱'])
            if boxes > 0:
                end_box_num = start_box_num + boxes - 1
                row['纸箱编号'] = f"{start_box_num}-{end_box_num}"
                start_box_num = end_box_num + 1
        
        
        return pd.DataFrame(result_data)

class TaskManager:
    def __init__(self, max_workers=3):
        self.task_queue = queue.Queue()
        self.results = {}
        self.executor = ThreadPoolExecutor(max_workers=max_workers)
        self.lock = threading.Lock()
        
    def add_task(self, fba_path):
        task_id = str(uuid.uuid4())
        self.results[task_id] = {'status': 'pending', 'result_file': None}
        self.executor.submit(self.process_task, task_id, fba_path)
        return task_id
        
    def process_task(self, task_id, fba_path):
        try:
            processor = ExcelProcessor(fba_path, './product_info.xlsx')
            processor.process_fba_shipment()
            processor.process_product_summary()
            processor.process_packing_list()
            
            result_df = processor.generate_result()
            timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            result_filename = f'result_{timestamp}.xlsx'
            result_path = os.path.join(app.config['RESULT_FOLDER'], result_filename)
            
            # 创建ExcelWriter对象以设置单元格格式
            writer = pd.ExcelWriter(result_path, engine='openpyxl')
            result_df.to_excel(writer, index=False)
            
            # 获取工作表
            worksheet = writer.sheets['Sheet1']
            
            # 设置所有列的宽度为12
            for col in worksheet.columns:
                worksheet.column_dimensions[col[0].column_letter].width = 20
                
            # 设置所有行的高度为12
            for row in worksheet.rows:
                worksheet.row_dimensions[row[0].row].height = 35
                for cell in row:
                    cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
      
            writer.close()
            
            with self.lock:
                self.results[task_id] = {
                    'status': 'completed',
                    'result_file': result_filename
                }
            
            os.remove(fba_path)
            
        except Exception as e:
            with self.lock:
                self.results[task_id] = {
                    'status': 'error',
                    'error': str(e)
                }
    
    def get_task_status(self, task_id):
        return self.results.get(task_id)

app = Flask(__name__)

# 配置上传文件存储路径和允许的文件类型
UPLOAD_FOLDER = './uploads'
RESULT_FOLDER = './results'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['RESULT_FOLDER'] = RESULT_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# 确保上传和结果文件夹存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

# 创建全局任务管理器实例
task_manager = TaskManager(max_workers=3)

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'fba_file' not in request.files:
        return '没有上传FBA文件', 400
    
    fba_file = request.files['fba_file']
    if fba_file.filename == '':
        return '未选择文件', 400
    
    if fba_file and allowed_file(fba_file.filename):
        fba_filename = secure_filename(fba_file.filename)
        fba_path = os.path.join(app.config['UPLOAD_FOLDER'], fba_filename)
        fba_file.save(fba_path)
        
        # 添加到任务队列
        task_id = task_manager.add_task(fba_path)
        return {'status': 'accepted', 'task_id': task_id}
    
    return '不支持的文件类型', 400

@app.route('/task/<task_id>', methods=['GET'])
def check_task(task_id):
    task_status = task_manager.get_task_status(task_id)
    if not task_status:
        return {'status': 'not_found'}, 404
        
    if task_status['status'] == 'completed':
        download_url = url_for('download_file', filename=task_status['result_file'])
        return {
            'status': 'completed',
            'download_url': download_url
        }
    elif task_status['status'] == 'error':
        return {
            'status': 'error',
            'message': task_status['error']
        }
    else:
        return {'status': 'pending'}

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(
        os.path.join(app.config['RESULT_FOLDER'], filename),
        as_attachment=True,
        download_name=filename
    )

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0',port=80)
