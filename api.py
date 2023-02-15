import base64
from sanic import Sanic
from sanic.response import raw

from pdf2excel import img_to_excel, pdf_to_excel

app = Sanic("paddle_service")

@app.route("/pdf-to-excel", methods=["POST"])
async def pdf_to_xlsx(request):
    remove_spaces = request.json.get('remove_spaces', False)
    pdf = request.json.get('pdf', None)
    is_pdf = True
    if pdf is None:
        is_pdf = False
        pdf = request.json['img']
    # decode base64
    pdf_bytes = base64.b64decode(pdf)
    excel_bytes = None
    if not is_pdf:
        excel_bytes = img_to_excel(pdf_bytes, remove_spaces=remove_spaces)
    else:
        excel_bytes = pdf_to_excel(pdf_bytes, remove_spaces=remove_spaces)
    headers = {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename=result.xlsx'
    }
    return raw(excel_bytes, headers=headers)

if __name__ == "__main__":
  app.run(host="0.0.0.0", port=10086)
