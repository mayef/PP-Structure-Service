
from io import BytesIO
import os
import subprocess
from tempfile import TemporaryDirectory
import openpyxl
import pandas as pd
import fitz

ROOTDIR = os.path.dirname(os.path.abspath(__file__))

def f_remove_spaces(s):
    r = ''
    try:
        r = s.replace(' ', '')
    except:
        r = s
    return r

def combine_tables(tables, remove_spaces=False):
    result = BytesIO()
    writer = pd.ExcelWriter(result, engine='openpyxl') #xlsxwriter
    len_pre = 1
    for i, table in enumerate(tables):
        df = pd.read_excel(BytesIO(table))
        if remove_spaces:
            df = df.applymap(f_remove_spaces)
        df.to_excel(writer, index=False, startrow=len_pre + i * 2)
        len_pre += len(df)
        writer.sheets['Sheet1'].row_dimensions[len_pre + 2 + i * 2].fill = openpyxl.styles.PatternFill(fill_type='solid', start_color='21B2AA', end_color='21B2AA')
    writer.close()
    return result.getvalue()

def convert_from_bytes(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype='pdf')
    images = []
    for page in doc:
        pix = page.get_pixmap()
        image = fitz.Pixmap(fitz.csRGB, pix)
        images.append(image)
    doc.close()
    return images

def pdf_to_excel(pdf_bytes, remove_spaces=False):
    images = convert_from_bytes(pdf_bytes)
    with TemporaryDirectory(dir=os.path.join(ROOTDIR, 'tmp')) as tmp_dir:
        for idx, image in enumerate(images):
            name = os.path.join(tmp_dir, str(idx) + '.png')
            image.save(name)

        result = []
        # remove comments and change the path to your paddleocr path if you want to use virtual environment
        os.chdir("C:/Users/zzz/anaconda3/envs/paddle_env/Scripts")
        subprocess.call('paddleocr --image_dir=' + tmp_dir + ' --type=structure --layout=false --output=' + tmp_dir, shell=True)
        paths = os.listdir(tmp_dir)
        paths.sort(key=lambda x:int(x.split('.')[0]))
        for dir in paths:
            if not os.path.isdir(os.path.join(tmp_dir, dir)):
                continue
            files = os.listdir(os.path.join(tmp_dir, dir))
            for file in files:
                if file.endswith('.xlsx'):
                    with open(os.path.join(tmp_dir, dir, file), 'rb') as f:
                        result.append(f.read())
        xlsx = combine_tables(result, remove_spaces=remove_spaces)
        return xlsx

def img_to_excel(img_bytes, remove_spaces=False):
    with TemporaryDirectory(dir=os.path.join(ROOTDIR, 'tmp')) as tmp_dir:
        # write image to tmp dir
        name = os.path.join(tmp_dir, '0.jpg')
        with open(name, 'wb') as f:
            f.write(img_bytes)

        result = []
        # remove comments and change the path to your paddleocr path if you want to use virtual environment
        os.chdir("C:/Users/zzz/anaconda3/envs/paddle_env/Scripts")
        subprocess.call('paddleocr --image_dir=' + tmp_dir + '/0.jpg' + ' --type=structure --layout=false --output=' + tmp_dir, shell=True)
        for file in os.listdir(os.path.join(tmp_dir, '0')):
            if file.endswith('.xlsx'):
                with open(os.path.join(tmp_dir, '0', file), 'rb') as f:
                    result.append(f.read())
        xlsx = combine_tables(result, remove_spaces=remove_spaces)
        return xlsx
