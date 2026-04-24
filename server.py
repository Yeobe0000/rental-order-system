#!/usr/bin/env python3
from flask import Flask, request, send_file, jsonify
import os, io
from openpyxl import load_workbook
from datetime import datetime

app = Flask(__name__, static_folder='.', static_url_path='')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, '오더지_양식.xlsx')

BASE = {
    'no': 1, 'booth': 2, 'company': 3, 'payment': 4,
    'subtotal': 6, 'memo': 71, 'bt': 72, 'bu': 73, 'bv': 74
}

ITEM_COLS = {
    'I':9,'J':10,'K':11,'L':12,'M':13,'N':14,'O':15,'P':16,'Q':17,
    'R':18,'S':19,'T':20,'U':21,'V':22,'W':23,
    'Z':26,'AA':27,'AB':28,'AC':29,
    'AF':32,'AG':33,'AH':34,
    'AK':37,'AL':38,'AM':39,'AN':40,'AO':41,'AP':42,'AQ':43,
    'AT':46,'AU':47,'AV':48,'AW':49,'AX':50,'AY':51,'AZ':52,
    'BC':55,'BD':56,'BE':57,'BF':58,'BG':59,
    'BJ':62,'BK':63,'BL':64,'BM':65,'BN':66,'BO':67,
    'BR':70,
}

ROW_PRE_START  = 13
ROW_PRE_END    = 63
ROW_POST_START = 67
ROW_POST_END   = 71

@app.route('/')
def index():
    return app.send_static_file('index.html')

@app.route('/ping')
def ping():
    return jsonify({'status': 'ok', 'template': os.path.exists(TEMPLATE_PATH)})

@app.route('/export', methods=['POST'])
def export():
    try:
        data = request.get_json()
        orders = data.get('orders', [])

        if not orders:
            return jsonify({'error': '저장된 발주가 없습니다.'}), 400
        if not os.path.exists(TEMPLATE_PATH):
            return jsonify({'error': '오더지 양식 파일을 찾을 수 없습니다.'}), 400

        with open(TEMPLATE_PATH, 'rb') as f:
            buf = io.BytesIO(f.read())

        wb = load_workbook(buf)
        ws = wb.active

        period = (orders[0].get('period') or '').strip()
        if period:
            ws.cell(3, 1).value = '행사기간 : ' + period

        pre_orders  = [o for o in orders if o.get('type') != 'post']
        post_orders = [o for o in orders if o.get('type') == 'post']

        def write_order(order, row):
            ws.cell(row, BASE['no']).value      = order.get('no') or ''
            ws.cell(row, BASE['booth']).value   = order.get('booth') or ''
            ws.cell(row, BASE['company']).value = order.get('company') or ''
            ws.cell(row, BASE['payment']).value = order.get('payment') or ''
            if order.get('subtotal'):
                ws.cell(row, BASE['subtotal']).value = int(order['subtotal'])
            if order.get('memo'):
                ws.cell(row, BASE['memo']).value = order['memo']
            if order.get('bt'):
                ws.cell(row, BASE['bt']).value = order['bt']
            if order.get('bu'):
                ws.cell(row, BASE['bu']).value = order['bu']
            if order.get('bv'):
                ws.cell(row, BASE['bv']).value = order['bv']
            for key, qty in (order.get('items') or {}).items():
                col = ITEM_COLS.get(key)
                if col and qty and int(qty) > 0:
                    ws.cell(row, col).value = int(qty)

        for i, order in enumerate(pre_orders):
            row = ROW_PRE_START + i
            if row > ROW_PRE_END: break
            write_order(order, row)

        for i, order in enumerate(post_orders):
            row = ROW_POST_START + i
            if row > ROW_POST_END: break
            write_order(order, row)

        out_buf = io.BytesIO()
        wb.save(out_buf)
        out_buf.seek(0)

        filename = f"오더지_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(out_buf, as_attachment=True, download_name=filename,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f"서버 시작: http://localhost:{port}")
    app.run(host='0.0.0.0', port=port, debug=False)
