FROM paddlepaddle/paddle:2.4.1

# RUN pip install paddleocr>=2.6.0.3 -i https://mirror.baidu.com/pypi/simple
# RUN pip uninstall -y pymupdf numpy
# RUN pip install pymupdf==1.19.0 numpy==1.21.6 -i https://mirror.baidu.com/pypi/simple

# install requirements
COPY requirements.txt /requirements.txt
RUN pip install -r /requirements.txt -i https://mirror.baidu.com/pypi/simple

# copy code
COPY . /app
WORKDIR /app

# run
CMD ["python", "api.py"]

# docker build -t zzz_ocr:1.0 .
# docker run -d -p 10086:10086 --name fuck zzz_ocr:1.0

