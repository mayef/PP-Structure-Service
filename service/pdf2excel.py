
from io import BytesIO
import os
from tempfile import NamedTemporaryFile
import cv2
from paddleocr import PPStructure, draw_structure_result, save_structure_res
import pandas as pd
from tablepyxl import tablepyxl

ROOTDIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

table_engine = PPStructure(
    lang = 'ch',
    show_log = False, 
    # image_orientation = True,
    # det_model_dir = os.path.join(ROOTDIR, 'model\\ch_ppocr_server_v2.0_det_infer'), 
    # rec_model_dir = os.path.join(ROOTDIR, 'model\\ch_ppocr_server_v2.0_rec_infer'), 
    # table_model_dir = os.path.join(ROOTDIR, 'model\\cn_ppstructure_mobile_v2.0_SLANet_infer'), 
    # layout_model_dir = os.path.join(ROOTDIR, 'model\\picodet_lcnet_x1_0_fgd_layout_cdla_infer')
)

save_folder = './output'
img_path = os.path.join(ROOTDIR, 'img/08.jpg')
img = cv2.imread(img_path)
result0 = table_engine(img)
# save_structure_res(result, save_folder, os.path.basename(img_path).split('.')[0])

img_path = os.path.join(ROOTDIR, 'img/09.jpg')
img = cv2.imread(img_path)
result1 = table_engine(img)

def merge_array(*arr):
    result = []
    for a in arr:
        result.extend(a)
    return result

def result_to_table_bytes(result): 
    tables = []
    for region in result:
        if region['type'].lower() == 'table' and len(region['res']) > 0 and 'html' in region['res']:
            wb = tablepyxl.document_to_workbook(region['res']['html'])
            with NamedTemporaryFile(dir=os.path.join(ROOTDIR, 'tmp'), delete=True) as tmp:
                wb.save(tmp.name)
                tmp.seek(0)
                bytes = tmp.read()
                tables.append(bytes)
    return tables

def combine_tables(tables):
    result = BytesIO()
    writer = pd.ExcelWriter(result, engine='openpyxl') #xlsxwriter
    for i, table in enumerate(tables):
        df = pd.read_excel(BytesIO(table))
        df.to_excel(writer, index=False, startrow=len(df) * i + i * 2, header=False)
        writer.sheets['Sheet1'].row_dimensions[len(df) * i + i * 2].height = 40
    writer.save()
    return result.getvalue()

b0 = result_to_table_bytes(result0)
b1 = result_to_table_bytes(result1)
b = merge_array(b0, b1)
excel = combine_tables(b)
# write to save_folder
with open(os.path.join(ROOTDIR, save_folder, 'result.xlsx'), 'wb') as f:
    f.write(excel)


# for line in result:
#     line.pop('img')
#     print(line)

# from PIL import Image

# font_path = os.path.join(ROOTDIR, 'font/simsun.ttf')
# image = Image.open(img_path).convert('RGB')
# im_show = draw_structure_result(image, result, font_path = font_path)
# im_show = Image.fromarray(im_show)
# im_show.save('result.jpg')

# to_excel(result, 'result.xlsx')

# PDF转图片
# 图片转Excel
# Excel合并
