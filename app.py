from flask import Flask, request, send_file, render_template_string
import pandas as pd
from io import BytesIO
import openpyxl

app = Flask(__name__)

HTML_TEMPLATE = """
<!doctype html>
<html lang=\"ko\">
  <head>
    <meta charset=\"utf-8\">
    <title>운송장 번호 병합기</title>
  </head>
  <body style=\"font-family: sans-serif; text-align: center; margin-top: 50px;\">
    <h1>📦 운송장번호 자동 병합 도구</h1>
    <form action=\"/merge\" method=\"post\" enctype=\"multipart/form-data\">
      <label>📄 배송리스트 파일 (.xlsx):</label><br>
      <input type=\"file\" name=\"delivery_file\" accept=\".xlsx\" required><br><br>
      <label>📄 발주서_크기순 파일 (.xlsx):</label><br>
      <input type=\"file\" name=\"tracking_file\" accept=\".xlsx\" required><br><br>
      <button type=\"submit\" style=\"font-size: 16px;\">운송장 병합하기</button>
    </form>
  </body>
</html>
"""

@app.route("/", methods=["GET"])
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route("/merge", methods=["POST"])
def merge():
    delivery_file = request.files['delivery_file']
    tracking_file = request.files['tracking_file']

    # 데이터프레임 로드
    delivery_df = pd.read_excel(delivery_file)
    tracking_wb = openpyxl.load_workbook(tracking_file, data_only=True)
    tracking_ws = tracking_wb["발주서_크기순"]

    # 발주서_크기순 시트에서 주문번호와 운송장번호 추출
    header = [cell.value for cell in tracking_ws[1]]
    주문번호_idx = header.index("주문번호")
    운송장_idx = header.index("운송장번호")

    tracking_dict = {}
    for row in tracking_ws.iter_rows(min_row=2, values_only=True):
        tracking_dict[str(row[주문번호_idx]).strip()] = str(row[운송장_idx]).strip()

    # 주문번호 기준으로 기존 운송장번호 열 덮어쓰기
    delivery_df["운송장번호"] = delivery_df["주문번호"].astype(str).map(tracking_dict)

    # 결과 저장
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        delivery_df.to_excel(writer, index=False, sheet_name="배송리스트")

    output.seek(0)
    return send_file(output, as_attachment=True,
                     download_name="배송리스트_운송장포함.xlsx",
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == "__main__":
    app.run(debug=True)
