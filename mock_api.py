import random
from datetime import datetime, timedelta
from flask import Flask, request, jsonify

app = Flask(__name__)

VALID_API_KEY = "1234"

@app.route('/123/log', methods=['GET'])
def mock_api():
    try:
        client_key = request.headers.get('X-API-KEY')   

        # 1. ตรวจสอบ API Key (เปลี่ยนชื่อเป็น api_key ให้ตรงกับหน้าบ้าน)
        if client_key != VALID_API_KEY:
            print(f"❌ Unauthorized access attempt with key: {client_key}")
            return jsonify({"success": False, "message": "Unauthorized"}), 401

        # 2. รับ Parameter และเตรียมตัวแปร
        requested_keys = request.args.get('keys', '').split(',')
        start_t = request.args.get('start_t')
        end_t = request.args.get('end_t')
        
        logs = []
        now = datetime.now()

        # 3. เริ่มลูปสร้างข้อมูล (ตัวแปร i เริ่มทำงานตรงนี้)
        for i in range(24 * 3): 
            log_time = now - timedelta(hours=i)
            ts = int(log_time.timestamp())

            # กรองช่วงเวลา (ถ้ามีการส่งมา)
            if start_t and ts < int(start_t): continue
            if end_t and ts > int(end_t): continue

            # --- คำนวณค่ามิเตอร์ (ต้องอยู่ในลูป for เท่านั้น) ---
            full_values = {
                "kwh": round(5000 + (100 - i * 1.5), 2),
                "v": round(220 + random.uniform(-5, 5), 1),
                "i": round(2 + random.uniform(-1, 2), 2)
            }

            # ถ้าขอ keys=i,kwh ตัวแปร filtered_value จะมีแค่ 2 ค่านั้น
            if requested_keys and requested_keys[0] != '':
                filtered_value = {k: full_values[k] for k in requested_keys if k in full_values}
            else:
                filtered_value = full_values

            item = {
                "ts": ts,
                "value": filtered_value
            }
            logs.append(item)

        return jsonify({
            "success": True,
            "count": len(logs),
            "data": logs
        })
    

    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.route('/234/log', methods=['GET'])
def mock_water_api():
    try:
        client_key = request.headers.get('X-API-KEY')   

        # 1. ตรวจสอบ API Key (เปลี่ยนชื่อเป็น api_key ให้ตรงกับหน้าบ้าน)
        if client_key != VALID_API_KEY:
            print(f"❌ Unauthorized access attempt with key: {client_key}")
            return jsonify({"success": False, "message": "Unauthorized"}), 401

        # 2. รับ Parameter และเตรียมตัวแปร
        requested_keys = request.args.get('keys', '').split(',')
        start_t = request.args.get('start_t')
        end_t = request.args.get('end_t')
        
        logs = []
        now = datetime.now()

        # 3. เริ่มลูปสร้างข้อมูล (ตัวแปร i เริ่มทำงานตรงนี้)
        for i in range(24 * 3): 
            log_time = now - timedelta(hours=i)
            ts = int(log_time.timestamp())

            # กรองช่วงเวลา (ถ้ามีการส่งมา)
            if start_t and ts < int(start_t): continue
            if end_t and ts > int(end_t): continue

            # --- คำนวณค่ามิเตอร์ (ต้องอยู่ในลูป for เท่านั้น) ---
            full_values = {
                "m3": round(5000 + (100 - i * 1.5), 2),
            }

            # ถ้าขอ keys=i,kwh ตัวแปร filtered_value จะมีแค่ 2 ค่านั้น
            if requested_keys and requested_keys[0] != '':
                filtered_value = {k: full_values[k] for k in requested_keys if k in full_values}
            else:
                filtered_value = full_values

            item = {
                "ts": ts,
                "value": filtered_value
            }
            logs.append(item)

        return jsonify({
            "success": True,
            "count": len(logs),
            "data": logs
        })
    

    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == '__main__':
    # รันที่ port 8000
    app.run(host='0.0.0.0', port=8000, debug=True)