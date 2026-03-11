import json,time,re,struct
from flask import jsonify
from pymodbus.client.sync import ModbusSerialClient, ModbusTcpClient
from pymodbus.payload import BinaryPayloadDecoder , BinaryPayloadBuilder
from pymodbus.constants import Endian
from pymodbus.pdu import ExceptionResponse
from pymodbus.exceptions import ModbusException
from threading import Lock
import os
import requests

modbus_lock = Lock()

def convert_to_registers(val, data_type=None, count=None, gain=None, word_swap=False):
    """แปลงค่า real-world -> register list (16-bit)"""
    # ---- Safety: ป้องกัน None ----
    try:
        gain = float(gain)
    except:
        gain = 1

    try:
        val = float(val)
    except:
        raise ValueError(f"Invalid value: {val}")

    # ---- Fix: data_type None ----
    if not data_type:
        data_type = "int16"

    t = data_type.lower()

    val_scaled = val * gain

    if t in ["int16", "uint16"]:
        return [int(val_scaled) & 0xFFFF]

    elif t in ["int32", "uint32"]:
        high = (int(val_scaled) >> 16) & 0xFFFF
        low  = int(val_scaled) & 0xFFFF
        return [low, high] if word_swap else [high, low]

    elif t in ["float32", "single"]:
        b = struct.pack('>f', float(val_scaled))
        reg1 = int.from_bytes(b[0:2], 'big')
        reg2 = int.from_bytes(b[2:4], 'big')
        return [reg2, reg1] if word_swap else [reg1, reg2]

    elif t == "bcd":
        # 1. จัดการทศนิยมและ Gain ก่อน (ถ้ามี)
        val_int = int(round(val_scaled))
        str_val = str(val_int)

        # 2. คำนวณหาจำนวน Register ที่ต้องใช้จริง
        # ถ้าไม่ได้ระบุ count มา ให้คำนวณจากความยาวตัวเลข (4 หลักต่อ 1 Reg)
        needed_regs = count if (count and count > 0) else (len(str_val) + 3) // 4
        
        # 3. เติม 0 ด้านหน้าให้ครบ (zfill) เพื่อให้แบ่งกลุ่มละ 4 ได้พอดี
        # เช่น '111000631' (9 หลัก) -> '000111000631' (12 หลัก = 3 Regs)
        padded_str = str_val.zfill(needed_regs * 4)
        
        regs = []
        # 4. หั่น string ทีละ 4 หลัก แล้วแปลงเป็น Hex (ฐาน 16)
        for i in range(0, len(padded_str), 4):
            chunk = padded_str[i:i+4]
            # ตรวจสอบว่าแต่ละหลักไม่เกิน 9 (เผื่อกรณี input ผิด)
            for char in chunk:
                if int(char) > 9:
                    raise ValueError(f"Invalid BCD digit: {char}")
            
            # แปลง "0631" (str) เป็น 0x0631 (int)
            regs.append(int(chunk, 16))
            
        return regs

    else:
        raise ValueError(f"Unsupported data_type {data_type}")

def bcd_words_to_decimal(words):
    """
    words = list ของค่า BCD เช่น [0x12, 0x34]
    คืนค่า decimal เช่น 1234
    """
    digits = ""

    for w in words:
        high = (w >> 4) & 0x0F
        low = w & 0x0F

        if high > 9 or low > 9:
            raise ValueError(f"Invalid BCD nibble in word: {hex(w)}")

        digits += f"{high}{low}"

    return int(digits)

def bcd_words_to_string(words): 
    """
    แปลง list ของ words เป็นตัวเลข string 
    รองรับทั้ง BCD แท้ (0-9) และ Integer ที่ถูกเขียนลงไปแบบผิด format
    """
    result = ""
    for w in words:
        # 1. แปลงเป็นค่า Integer ก่อน
        val_int = w if isinstance(w, int) else int(str(w), 16)
        
        # 2. ตรวจสอบว่ามีตัวอักษร A-F (Non-BCD) หรือไม่
        hex_str = f"{val_int:04X}"
        is_pure_bcd = all(char.isdigit() for char in hex_str)

        if is_pure_bcd:
            # ถ้าเป็น BCD แท้ (เช่น 1107) ให้ใช้ค่านั้นเลย
            result += hex_str
        else:
            # ถ้าไม่ใช่ BCD (เช่น 044C) ให้แปลงเป็นฐาน 10 ก่อน (จะได้ 1100)
            # แล้วเติม 0 ให้ครบ 4 หลัก
            result += f"{val_int:04d}"

    # ตัดเลข 0 ข้างหน้าออกเพื่อให้สวยงาม
    return result.lstrip('0') if result.lstrip('0') else "0"

def read_meter_tool(model_name=None, register_key=None, port=None, ip=None, 
                   serial_ports=None, address=None, unit_id=None,
                   api_base_url=None, api_token=None, 
                   baudrate=None, parameters=None, pc_port=None, 
                   function_code=None, word_swap=None, connType=None,
                   data_type=None, count=None, gain=None,start_t=None,end_t=None,requested_keys=None):
    
    value = None
    combined_meters = {}

    # 1. โหลด config จากทั้ง model.json และ model_water.json
    config_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config_meter')
    target_files = ['model.json', 'model_water.json']

    for filename in target_files:
        file_path = os.path.join(config_dir, filename)
        if os.path.exists(file_path):
            with open(file_path, 'r', encoding='utf-8') as f:
                try:
                    combined_meters.update(json.load(f))
                except json.JSONDecodeError:
                    print(f"❌ Error decoding {filename}")

    client = None

    if model_name == "custom":
        c_type = connType
        c_address = int(address)
        c_unit_id = int(unit_id)
        c_function_code = int(function_code)
        c_data_type = data_type.lower()
        c_count = int(count)
        c_gain = float(gain)
        c_word_swap = (word_swap == True or word_swap == 'True')
        
        c_port = serial_ports
        c_baudrate = int(baudrate) if baudrate else 9600
        c_params = parameters if parameters else "8N1"

        if c_type == "tcp":
            client = ModbusTcpClient(ip_addr, port=tcp_port)
        elif c_type == "serial":
            parity = c_params[1]
            stopbits = int(c_params[2])
            bytesize = int(c_params[0])
            client = ModbusSerialClient(
                method="rtu", port=c_port, baudrate=c_baudrate,
                parity=parity, stopbits=stopbits, bytesize=bytesize, timeout=2
            )
        
        if client is None:
            return "❌ Error: ไม่สามารถสร้าง Client ได้ (ตรวจสอบ Connection Type)"

        # 4. เริ่มต้นอ่านค่าด้วย Lock
        with modbus_lock:
            try:
                if not client.connect():
                    print(f"❌ Cannot connect to {c_port or ip_addr}")
                    return None

                time.sleep(0.5)
                
                # เลือก Function Code ในการอ่าน
                if c_function_code == 1:
                    rr = client.read_coils(c_address, c_count, unit=c_unit_id)
                elif c_function_code == 2:
                    rr = client.read_discrete_inputs(c_address, c_count, unit=c_unit_id)
                elif c_function_code == 3:
                    rr = client.read_holding_registers(c_address, c_count, unit=c_unit_id)
                elif c_function_code == 4:
                    rr = client.read_input_registers(c_address, c_count, unit=c_unit_id)
                else:
                    return None

                # ตรวจสอบ Error
                error_check = decode_modbus_error(rr)
                if not error_check["success"]:
                    print(f"❌ Modbus Error: {error_check['message']}")
                    return None

                # แปลงผลลัพธ์
                paddedHex = [hex(r).replace("0x", "").zfill(4) for r in rr.registers]
                
                if c_word_swap:
                    if c_count == 2:
                        paddedHex = [paddedHex[1], paddedHex[0]]
                    elif c_count == 4:
                        paddedHex = [paddedHex[1], paddedHex[0], paddedHex[3], paddedHex[2]]

                # Decode ตาม Data Type
                if c_data_type in ["int16", "uint16", "int32", "uint32", "float32", "float", "float64", "double"]:
                    decoder = BinaryPayloadDecoder.fromRegisters(rr.registers, byteorder=Endian.Big, wordorder=Endian.Big)
                    if c_data_type == "int16": value = decoder.decode_16bit_int()
                    elif c_data_type == "uint16": value = decoder.decode_16bit_uint()
                    elif c_data_type == "int32": value = decoder.decode_32bit_int()
                    elif c_data_type == "uint32": value = decoder.decode_32bit_uint()
                    elif c_data_type in ["float32", "float"]: value = decoder.decode_32bit_float()
                    elif c_data_type in ["float64", "double"]: value = decoder.decode_64bit_float()
                elif c_data_type == "bcd":
                    words = [int(h, 16) for h in paddedHex]
                    value = bcd_words_to_decimal(words) if c_count == 1 else bcd_words_to_string(words)

                if value is not None:
                    final_value = round(float(value) * c_gain, 3)
                    print(f"✅ Successful Read: {final_value}")
                    return final_value
                
                return None

            except Exception as e:
                print(f"❌ Modbus Exception: {e}")
                return None
            finally:
                if client:
                    time.sleep(0.5)
                    client.close()
                    print(f"🔌 Connection closed.")
        

    else:
        # ดึงค่าจาก ใน JSON
        module_cfg = combined_meters.get(model_name)
        if not module_cfg:
            raise ValueError(f"ไม่พบโมเดล {model_name} ในไฟล์ตั้งค่า")
        
        c_type = connType or module_cfg.get("type")

        if c_type == "api":
            try:
                base_url = (api_base_url)
                api_path = module_cfg.get("path", "").lstrip('/')
                url = f"{base_url}/{api_path}"
                print(f"🔗 Requesting API: {url}")

                secret = api_token 
                headers = {}
                if secret: # ถ้ามี Token ค่อยใส่ Header
                    headers["X-API-KEY"] = secret

                if not requested_keys: return "กรุณาเลือกคีย์ที่ต้องการอ่าน"

                target_key_names = [module_cfg.get("read_register", {}).get(k, {}).get("key_name") 
                                   for k in requested_keys if "key_name" in module_cfg.get("read_register", {}).get(k, {})]

                params = {
                    "keys": ",".join(target_key_names),
                    "start_t": start_t,
                    "end_t": end_t
                }

                response = requests.get(url, params=params, headers=headers, timeout=5)
                
                if response.status_code == 200:
                    res_data = response.json()
                    logs = res_data.get("data", [])
                    print(f"📦 Total Logs Received: {len(logs)} rows")
                    if not logs: return "No data found"
                    
                    has_time = bool(start_t or end_t)

                    if not has_time:
                        if not logs:
                            return "ไม่พบข้อมูลในระบบ API"
                        latest_entry = logs[0]
                        raw_api_values = latest_entry.get("value", {})
                        final_results = {}
                        for k in requested_keys:
                            reg_info = module_cfg.get("read_register", {}).get(k)
                            if reg_info:
                                api_key = reg_info.get("key_name")
                                val = raw_api_values.get(api_key)
                                gain = reg_info.get("gain", 1)
                                final_results[k] = str(round(float(val) * gain, 2)) if val is not None else "N/A"
                        
                        print(f"✅ API Read Successful (Single): {final_results}")
                        return final_results

                    else:
                        all_data_has_time = []
                        for entry in logs:
                            ts = entry.get("ts")
                            vals = entry.get("value", {})
                            row = {"ts": ts}
                            for k in requested_keys:
                                reg_info = module_cfg.get("read_register", {}).get(k)
                                if reg_info:
                                    api_key = reg_info.get("key_name")
                                    val = vals.get(api_key)
                                    gain = reg_info.get("gain", 1)
                                    row[k] = round(float(val) * gain, 2) if val is not None else "N/A"
                            all_data_has_time.append(row)
                        
                        print(f"✅ API Read Successful (Many): {len(all_data_has_time)} records")
                        return all_data_has_time
                
                return f"❌ API Error: {response.status_code}"

            except Exception as e:
                return f"❌ API Exception: {str(e)}"
    
        else:
            reg_cfg = module_cfg.get("read_register", {}).get(register_key) or \
                    module_cfg.get("config_register", {}).get(register_key)
            
            if not reg_cfg:
                raise ValueError(f"ไม่พบ register '{register_key}' ในโมเดล {model_name}")

            c_type = connType or module_cfg.get("type")
            c_address = int(address or reg_cfg.get("address"))
            c_unit_id = int(unit_id)
            c_function_code = int(function_code or 3) 
            c_data_type = (data_type or reg_cfg.get("data_type")).lower()
            c_count = int(reg_cfg.get("count", 2))
            c_gain = float(reg_cfg.get("gain", 1))
            c_word_swap = reg_cfg.get("word_swap", True)
            # สำหรับ Serial/TCP
            c_port = serial_ports or pc_port
            c_baudrate = int(baudrate or module_cfg.get("baudrate", 9600))
            c_params = parameters or module_cfg.get("parameters", "8N1")
            ip_addr = ip or "192.168.137.20"
            tcp_port = int(port or module_cfg.get("port", 502))

        # 3. สร้าง Modbus Client
        if c_type == "tcp":
            client = ModbusTcpClient(ip_addr, port=tcp_port)
        elif c_type == "serial":
            parity = c_params[1]
            stopbits = int(c_params[2])
            bytesize = int(c_params[0])
            client = ModbusSerialClient(
                method="rtu", port=c_port, baudrate=c_baudrate,
                parity=parity, stopbits=stopbits, bytesize=bytesize, timeout=2
            )
        
        if client is None:
            return "❌ Error: ไม่สามารถสร้าง Client ได้ (ตรวจสอบ Connection Type)"

        # 4. เริ่มต้นอ่านค่าด้วย Lock
        with modbus_lock:
            try:
                if not client.connect():
                    print(f"❌ Cannot connect to {c_port or ip_addr}")
                    return None

                time.sleep(0.5)
                
                # เลือก Function Code ในการอ่าน
                if c_function_code == 1:
                    rr = client.read_coils(c_address, c_count, unit=c_unit_id)
                elif c_function_code == 2:
                    rr = client.read_discrete_inputs(c_address, c_count, unit=c_unit_id)
                elif c_function_code == 3:
                    rr = client.read_holding_registers(c_address, c_count, unit=c_unit_id)
                elif c_function_code == 4:
                    rr = client.read_input_registers(c_address, c_count, unit=c_unit_id)
                else:
                    return None

                # ตรวจสอบ Error
                error_check = decode_modbus_error(rr)
                if not error_check["success"]:
                    print(f"❌ Modbus Error: {error_check['message']}")
                    return None

                # แปลงผลลัพธ์
                paddedHex = [hex(r).replace("0x", "").zfill(4) for r in rr.registers]
                
                if c_word_swap:
                    if c_count == 2:
                        paddedHex = [paddedHex[1], paddedHex[0]]
                    elif c_count == 4:
                        paddedHex = [paddedHex[1], paddedHex[0], paddedHex[3], paddedHex[2]]

                # Decode ตาม Data Type
                if c_data_type in ["int16", "uint16", "int32", "uint32", "float32", "float", "float64", "double"]:
                    decoder = BinaryPayloadDecoder.fromRegisters(rr.registers, byteorder=Endian.Big, wordorder=Endian.Big)
                    if c_data_type == "int16": value = decoder.decode_16bit_int()
                    elif c_data_type == "uint16": value = decoder.decode_16bit_uint()
                    elif c_data_type == "int32": value = decoder.decode_32bit_int()
                    elif c_data_type == "uint32": value = decoder.decode_32bit_uint()
                    elif c_data_type in ["float32", "float"]: value = decoder.decode_32bit_float()
                    elif c_data_type in ["float64", "double"]: value = decoder.decode_64bit_float()
                elif c_data_type == "bcd":
                    words = [int(h, 16) for h in paddedHex]
                    value = bcd_words_to_decimal(words) if c_count == 1 else bcd_words_to_string(words)

                if value is not None:
                    final_value = round(float(value) * c_gain, 3)
                    print(f"✅ Successful Read: {final_value}")
                    return final_value
                
                return None

            except Exception as e:
                print(f"❌ Modbus Exception: {e}")
                return None
            finally:
                if client:
                    time.sleep(0.5)
                    client.close()
                    print(f"🔌 Connection closed.")

def write_meter_tool(model_name=None, register_key=None, port=None, ip=None,
                     serial_ports=None, address=None, unit_id=None,
                     baudrate=None, parameters=None, pc_port=None,
                     function_code=None, word_swap=None, connType=None,
                     data_type=None, count=None, gain=None, data_dec=None):

    if not data_dec:
        print("❌ No value to write")
        return "กรุณาระบุค่าที่ต้องการเขียน"

    combined_meters = {}
    # 1. โหลด config จากทั้ง model.json และ model_water.json
    base_dir = os.path.dirname(os.path.abspath(__file__))
    config_dir = os.path.join(base_dir, 'config_meter')
    target_files = ['model.json', 'model_water.json']

    for filename in target_files:
        file_path = os.path.join(config_dir, filename)
        if os.path.exists(file_path):
            with open(file_path, 'r', encoding='utf-8') as f:
                try:
                    combined_meters.update(json.load(f))
                except Exception as e:
                    print(f"⚠️ Error loading {filename}: {e}")

    client = None
    words = []

    # 2. เตรียม Parameter (แยกหมวดหมู่ Custom / Profile)
    if model_name == "custom":
        c_type = connType
        c_address = int(address)
        c_unit_id = int(unit_id)
        c_baudrate = int(baudrate) if baudrate else 9600
        c_function_code = int(function_code) if function_code else 16
        c_params = parameters if parameters else "8N1"
        c_ip = ip or "192.168.137.20"
        c_port = int(port or 502)
        words = [int(v) for v in data_dec] 

        if c_type == "tcp":
            client = ModbusTcpClient(c_address, port=c_port)
        elif c_type == "serial":
            parity = c_params[1]
            stopbits = int(c_params[2])
            bytesize = int(c_params[0])
            client = ModbusSerialClient(
                method="rtu", port=c_port, baudrate=c_baudrate,
                parity=parity, stopbits=stopbits, bytesize=bytesize, timeout=2
            )
        
        if client is None:
            return "❌ Error: ไม่สามารถสร้าง Client ได้ (ตรวจสอบ Connection Type)"

        # 4. เริ่มต้นอ่านค่าด้วย Lock
        with modbus_lock:
            try:
                if not client.connect():
                    print(f"❌ Cannot connect to {c_port or c_address}")
                    return None

                time.sleep(0.5)
                
                # เลือก Function Code ในการเขียน (Write Operations)
                if c_function_code == 5:
                    # เขียน Single Coil (ส่งค่าเดียว)
                    rr = client.write_coil(c_address, words[0], unit=c_unit_id)
                elif c_function_code == 6:
                    # เขียน Single Holding Register
                    rr = client.write_register(c_address, words[0], unit=c_unit_id)
                elif c_function_code == 15:
                    # เขียน Multiple Coils (ส่งเป็น list ของ boolean/bit)
                    rr = client.write_coils(c_address, words, unit=c_unit_id)
                elif c_function_code == 16:
                    # เขียน Multiple Holding Registers (ส่งเป็น list ของ registers)
                    rr = client.write_registers(c_address, words, unit=c_unit_id)
                else:
                    return "Function Code สำหรับการเขียนไม่ถูกต้อง"

                error_check = decode_modbus_error(rr)
                if not error_check["success"]:
                    print(f"❌ Modbus Error: {error_check['message']}")
                    return None

                paddedHex = [hex(r).replace("0x", "").zfill(4) for r in rr.registers]
                
                if c_word_swap:
                    if c_count == 2:
                        paddedHex = [paddedHex[1], paddedHex[0]]
                    elif c_count == 4:
                        paddedHex = [paddedHex[1], paddedHex[0], paddedHex[3], paddedHex[2]]

                if c_data_type in ["int16", "uint16", "int32", "uint32", "float32", "float", "float64", "double"]:
                    decoder = BinaryPayloadDecoder.fromRegisters(rr.registers, byteorder=Endian.Big, wordorder=Endian.Big)
                    if c_data_type == "int16": value = decoder.decode_16bit_int()
                    elif c_data_type == "uint16": value = decoder.decode_16bit_uint()
                    elif c_data_type == "int32": value = decoder.decode_32bit_int()
                    elif c_data_type == "uint32": value = decoder.decode_32bit_uint()
                    elif c_data_type in ["float32", "float"]: value = decoder.decode_32bit_float()
                    elif c_data_type in ["float64", "double"]: value = decoder.decode_64bit_float()
                elif c_data_type == "bcd":
                    words = [int(h, 16) for h in paddedHex]
                    value = bcd_words_to_decimal(words) if c_count == 1 else bcd_words_to_string(words)

                if value is not None:
                    final_value = round(float(value) * c_gain, 3)
                    print(f"✅ Successful Read: {final_value}")
                    return final_value
                
                return None

            except Exception as e:
                print(f"❌ Modbus Exception: {e}")
                return None
            finally:
                if client:
                    time.sleep(0.5)
                    client.close()
                    print(f"🔌 Connection closed.")
    else:
        # กรณีเลือกจาก Profile มิเตอร์
        module_cfg = combined_meters.get(model_name)
        if not module_cfg:
            return f"ไม่พบโมเดล {model_name} ในระบบ"

        cfg = module_cfg.get("config_register", {}).get(register_key)
        if not cfg:
            return f"ไม่พบ register '{register_key}' ในโมเดล {model_name}"

        c_type = module_cfg.get("type")
        c_address = int(address or cfg.get("address"))
        c_unit_id = int(unit_id)
        c_baudrate = int(baudrate or module_cfg.get("baudrate", 9600))
        c_params = parameters or module_cfg.get("parameters", "8N1")
        c_ip = ip or "192.168.137.20"
        c_port = int(port or cfg.get("port", 502))
        
        c_gain = float(gain or cfg.get("gain", 1))
        c_count = int(count or cfg.get("count", 1))
        c_data_type = data_type or cfg.get("data_type", "int16")
        c_word_swap = word_swap if word_swap is not None else cfg.get("word_swap", False)

        # แปลงค่าตัวเลข (Decimal) ให้เป็น Modbus Registers (Words)
        try:
            if len(data_dec) == 1:
                val = data_dec[0]
                words = convert_to_registers(val, data_type=c_data_type, count=c_count, gain=c_gain, word_swap=c_word_swap)
            else:
                for val in data_dec:
                    regs = convert_to_registers(val, data_type=c_data_type, count=1, gain=c_gain, word_swap=c_word_swap)
                    words.extend(regs)
            
            if c_count and len(words) > c_count:
                words = words[:c_count]
        except Exception as e:
            return f"Error converting data: {str(e)}"

    # 3. สร้าง Client ตามประเภทการเชื่อมต่อ
    if c_type == "tcp":
        client = ModbusTcpClient(c_ip, port=c_port)
    else:
        parity = c_params[1]
        stopbits = int(c_params[2])
        bytesize = int(c_params[0])
        client = ModbusSerialClient(
            method="rtu",
            port=serial_ports or pc_port,
            baudrate=c_baudrate,
            parity=parity,
            stopbits=stopbits,
            bytesize=bytesize,
            timeout=2
        )

    # 4. เขียนข้อมูลลงอุปกรณ์ด้วย Modbus Lock
    with modbus_lock:
        try:
            if not client.connect():
                return "ไม่สามารถเชื่อมต่อกับอุปกรณ์ได้"

            # เขียนข้อมูล (ใช้ FC16 - Write Multiple Registers)
            rr = client.write_registers(c_address, words, unit=c_unit_id)
            print(f"➡️ Write Command: Model={model_name}, Addr={c_address}, Value={words}")

            check = decode_modbus_error(rr)
            if not check["success"]:
                return f"เขียนข้อมูลล้มเหลว: {check['message']}"

            return "เขียนข้อมูลสำเร็จ"

        except Exception as e:
            return f"เกิดข้อผิดพลาด: {str(e)}"
        finally:
            if client:
                client.close()

def read_meter_unit_read(model_name=None, register_key=None, ip=None,
                        port=None, serial_ports=None, slave_id=None, 
                        is_water=False, api_base_url=None, api_token=None):
    client = None
    # 1. โหลดคอนฟิก JSON ตามประเภทน้ำ/ไฟ
    json_path = 'config_meter/model_water.json' if is_water else 'config_meter/model.json'
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            meters = json.load(f)
    except FileNotFoundError:
        return f"Error: ไม่พบไฟล์คอนฟิก {json_path}"

    module_cfg = meters.get(model_name)
    if not module_cfg: 
        return f"Error: ไม่พบโมเดล {model_name} ในไฟล์คอนฟิก"

    reg_info = module_cfg.get("read_register", {}).get(register_key)
    if not reg_info: 
        return f"Error: ไม่พบคีย์ {register_key} ในโมเดล {model_name}"

    # ดึงค่าพื้นฐานจาก JSON
    conn_type = module_cfg.get("type")
    gain = float(reg_info.get("gain", 1))
    data_type = reg_info.get("data_type", "int16").lower()
    word_swap = reg_info.get("word_swap", False)
    byte_swap = reg_info.get("byte_swap", False) # เพิ่มเผื่อไว้

    # --- CASE 1: API (ดึงข้อมูลผ่าน Web Service) ---
    if conn_type == "api":
        try:
            if not api_base_url: return "Error: API URL is missing"
            base = api_base_url.rstrip('/')
            api_path = module_cfg.get("path", "").lstrip('/')
            url = f"{base}/{api_path}"
            
            target_api_key = reg_info.get("key_name")
            if not target_api_key: return f"Error: ไม่ระบุ key_name สำหรับ API ในคีย์ {register_key}"
            
            headers = {"X-API-KEY": api_token} if api_token else {}
            
            response = requests.get(url, params={"keys": target_api_key}, headers=headers, timeout=5)
            if response.status_code == 200:
                res_json = response.json()
                data_list = res_json.get("data", [])
                if not data_list: return "Error: API returned no data"
                
                # ดึงค่าจากแถวแรก (ล่าสุด)
                raw_val = data_list[0].get("value", {}).get(target_api_key)
                if raw_val is None: return f"Error: Key '{target_api_key}' not found in API response"

                final_val = round(float(raw_val) * gain, 2)

                return final_val
            return None
        except Exception as e:
            return f"API Exception: {str(e)}"

    # --- CASE 2: Modbus TCP ---
    elif conn_type == "tcp":
        target_ip = ip or module_cfg.get("ip")
        target_port = int(port or module_cfg.get("port", 502))
        if not target_ip: return "Error: Missing IP Address"
        client = ModbusTcpClient(host=target_ip, port=target_port, timeout=3)

    # --- CASE 3: Modbus Serial (RTU) ---
    elif conn_type == "serial":
        if not serial_ports: return "Error: Missing COM Port"
        baud = int(module_cfg.get("baudrate", 9600))
        params = module_cfg.get("parameters", "8N1") # เช่น 8N1, 7E1
        
        client = ModbusSerialClient(
            method="rtu",
            port=serial_ports,
            baudrate=baud,
            parity=params[1].upper(),
            stopbits=int(params[2]),
            bytesize=int(params[0]),
            timeout=2
        )
    else:
        return f"Error: ไม่รองรับการเชื่อมต่อแบบ {conn_type}"

    # --- ส่วนการอ่านค่า Modbus (TCP/Serial) ---
    with modbus_lock: # ป้องกันการเรียกซ้อนกันใน Port เดียวกัน
        try:
            if not client or not client.connect():
                return f"Error: ไม่สามารถเชื่อมต่อกับอุปกรณ์ ({conn_type})"
            
            address = int(reg_info.get("address"))
            count = int(reg_info.get("count", 2))
            rr = client.read_holding_registers(address, count, unit=slave_id)

            if rr.isError():
                return f"Modbus Error: {rr}"

            # เตรียมการ Decode
            paddedHex = [hex(r).replace("0x", "").zfill(4) for r in rr.registers]
            
            if word_swap:
                    if count == 2:
                        paddedHex = [paddedHex[1], paddedHex[0]]
                    elif count == 4:
                        paddedHex = [paddedHex[1], paddedHex[0], paddedHex[3], paddedHex[2]]

            decoder = BinaryPayloadDecoder.fromRegisters(rr.registers, byteorder=Endian.Big, wordorder=Endian.Big)
            
            # Decode ตาม Data Type
            if data_type == "int16":
                value = decoder.decode_16bit_int()
            elif data_type == "uint16":
                value = decoder.decode_16bit_uint()
            elif data_type == "int32":
                value = decoder.decode_32bit_int()
            elif data_type == "uint32":
                value = decoder.decode_32bit_uint()
            elif data_type in ["float32", "float"]:
                value = decoder.decode_32bit_float()
            elif data_type in ["float64", "double"]:
                value = decoder.decode_64bit_float()
            elif data_type == "bcd":
                words = [int(h, 16) for h in paddedHex]
                value = bcd_words_to_decimal(words) if count == 1 else bcd_words_to_string(words)
            else:
                return f"Error: ไม่รองรับ Data Type {data_type}"

            if value is not None:
                final_value = round(float(value) * gain, 3)
                print(f"✅ Successful Read: {final_value}")
                return final_value
            
            return None
        
        except Exception as e:
            print(f"❌ Modbus Exception: {e}")
            return None
        finally:
            if client:
                time.sleep(0.5)
                client.close()
                print(f"🔌 Connection closed.")
            

def decode_modbus_error(rr):
    """
    ตรวจสอบ Modbus response และแปลงเป็นข้อความอ่านง่าย
    rr: response object จาก pymodbus
    คืนค่า: dict {success, message}
    """
    if rr is None:
        return {"success": False, "message": "No response from device (timeout or wrong unit/address)"}

    if hasattr(rr, "isError") and rr.isError():
        if isinstance(rr, ExceptionResponse):
            exc_code = rr.exception_code
            fc = rr.function_code
            exc_msg = {
                1: "Illegal Function → อุปกรณ์ไม่รองรับฟังก์ชันนี้",
                2: "Illegal Data Address → ตรวจสอบ address/offset",
                3: "Illegal Data Value → ตรวจสอบจำนวน register หรือค่า",
                4: "Slave Device Failure → อุปกรณ์ล้มเหลว",
                5: "Acknowledge → อุปกรณ์กำลังประมวลผล ลองใหม่",
                6: "Slave Device Busy → อุปกรณ์กำลังงาน ลองช้าๆ",
                8: "Memory Parity Error → ตรวจสอบอุปกรณ์",
                10: "Gateway Path Unavailable → ไม่สามารถติดต่อ Slave ผ่าน gateway",
                11: "Gateway Target Device Failed → Slave ไม่ตอบผ่าน gateway",
            }.get(exc_code, f"Unknown exception code {exc_code}")

            return {"success": False, "message": f"Exception Response (FC {fc}): {exc_msg}"}
        else:
            return {"success": False, "message": f"Modbus Error: {rr}"}

    return {"success": True, "message": "OK"}
