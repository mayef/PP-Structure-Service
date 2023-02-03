FROM paddlepaddle/paddle:2.4.1

RUN pip install paddleocr>=2.6.0.3 -i https://mirror.baidu.com/pypi/simple
RUN pip uninstall -y pymupdf numpy
RUN pip install pymupdf==1.19.0 numpy==1.21.6 -i https://mirror.baidu.com/pypi/simple
