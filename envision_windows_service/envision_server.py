# envision_server.py

import logging
from flask import Flask, jsonify, request
from envision import Envision, ENUM_CURRENT_STATE_DESC
import pythoncom
import os
import time

# --- Basic Setup ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(module)s - %(message)s')

# --- COM Initialization ---
try:
    pythoncom.CoInitialize()
    com_initialized = True
except pythoncom.com_error:
    logging.info("COM was already initialized in this thread.")
    com_initialized = True

app = Flask(__name__)

# --- Global Envision Controller Instance ---
if com_initialized:
    try:
        logging.info("Initializing Envision controller...")
        envision_controller = Envision()
        if not envision_controller.connect():
            raise ConnectionError("Failed to connect to EnVision instrument on server startup.")
        logging.info("Envision controller connected and ready.")
    except Exception as e:
        logging.error(f"CRITICAL: Could not initialize Envision controller: {e}", exc_info=True)
        envision_controller = None
else:
    envision_controller = None

# --- Helper Functions ---

def extract_protocol_details(p_obj):
    """Safely extracts details from a COM protocol object for JSON serialization."""
    try:
        return {
            "name": getattr(p_obj, 'ProtName', 'Unknown Name'),
            "id": getattr(p_obj, 'AssayProtocolID', 'N/A'),
            "version": getattr(p_obj, 'ProtVersion', 'N/A')
        }
    except pythoncom.com_error:
        return {"name": "Error reading protocol details", "id": "N/A", "version": "N/A"}

def handle_api_error(e, endpoint_name):
    """A centralized error handler to log and format error responses."""
    error_msg = f"Error in endpoint '/{endpoint_name}': {e}"
    logging.error(error_msg, exc_info=True)
    return jsonify({"error": str(e)}), 500

def _wait_for_ready_for_plate_state(timeout=60):
    """Blocks until the instrument state is 10 (Waiting for Plate) or 2 (IDLE)."""
    logging.info(f"Blocking until instrument is ready for plate (State 10 or 2)...")
    start_time = time.time()
    while True:
        pythoncom.PumpWaitingMessages()
        current_state = envision_controller.get_current_state()
        if current_state in [2, 10]: # 2=IDLE, 10=Waiting for plate
            logging.info(f"Instrument is ready. Final state: {ENUM_CURRENT_STATE_DESC.get(current_state)}")
            return True
        if time.time() - start_time > timeout:
            logging.error(f"Timeout: Instrument did not become ready for a plate within {timeout}s.")
            return False
        time.sleep(0.2)

# --- API Endpoints ---

@app.route('/status', methods=['GET'])
def get_status():
    if not envision_controller: return jsonify({"error": "Controller is offline"}), 503
    try:
        pythoncom.PumpWaitingMessages()
        code = envision_controller.get_current_state()
        desc = ENUM_CURRENT_STATE_DESC.get(code, "Unknown State")
        return jsonify({"status_code": code, "status_description": desc})
    except Exception as e: return handle_api_error(e, "status")

@app.route('/properties', methods=['GET'])
def get_properties():
    if not envision_controller: return jsonify({"error": "Controller is offline"}), 503
    try:
        return jsonify(envision_controller.get_instrument_properties())
    except Exception as e: return handle_api_error(e, "properties")

@app.route('/protocols/valid', methods=['GET'])
def get_valid_protocols():
    if not envision_controller: return jsonify({"error": "Controller is offline"}), 503
    try:
        protocols = envision_controller.get_valid_protocols()
        return jsonify([extract_protocol_details(p) for p in protocols])
    except Exception as e: return handle_api_error(e, "protocols/valid")

@app.route('/protocols/all', methods=['GET'])
def get_all_protocols():
    if not envision_controller: return jsonify({"error": "Controller is offline"}), 503
    try:
        protocol_list = []
        raw_collection = envision_controller.get_all_protocols_raw_collection()
        if raw_collection:
            count = raw_collection.Count if hasattr(raw_collection, 'Count') else raw_collection.Count()
            for i in range(1, count + 1):
                protocol_list.append(extract_protocol_details(raw_collection.Item(i)))
        return jsonify(protocol_list)
    except Exception as e: return handle_api_error(e, "protocols/all")

@app.route('/protocols/default', methods=['GET'])
def get_default_protocol():
    if not envision_controller: return jsonify({"error": "Controller is offline"}), 503
    try:
        protocol_obj = envision_controller.get_default_protocol_obj()
        return jsonify(extract_protocol_details(protocol_obj) if protocol_obj else {"name": "N/A"})
    except Exception as e: return handle_api_error(e, "protocols/default")

@app.route('/protocols/default', methods=['POST'])
def set_default_protocol():
    if not envision_controller: return jsonify({"error": "Controller is offline"}), 503
    try:
        index = request.get_json().get('index')
        envision_controller.set_default_protocol_index(index)
        return jsonify({"status": "success", "message": "Default protocol updated."})
    except Exception as e: return handle_api_error(e, "protocols/default")

@app.route('/assay/start_default', methods=['POST'])
def start_default_assay():
    if not envision_controller: return jsonify({"error": "Controller is offline"}), 503
    try:
        if not envision_controller.start_default_assay():
            raise RuntimeError("start_default_assay COM call returned False.")
        if not _wait_for_ready_for_plate_state():
            raise RuntimeError("Instrument failed to enter 'Waiting for Plate' state after command.")
        return jsonify({"status": "success", "message": "Default assay initiated.", "assay_id": envision_controller.active_assay_id})
    except Exception as e: return handle_api_error(e, "assay/start_default")

@app.route('/assay/start_by_index', methods=['POST'])
def start_assay_by_index():
    if not envision_controller: return jsonify({"error": "Controller is offline"}), 503
    try:
        index = request.get_json().get('index')
        valid_protocols = envision_controller.get_valid_protocols()
        if not (isinstance(index, int) and 0 <= index < len(valid_protocols)):
            raise ValueError("Protocol index is out of bounds.")
        protocol_to_run = valid_protocols[index]
        if not envision_controller.start_assay(protocol_to_run):
            raise RuntimeError("start_assay COM call returned False.")
        if not _wait_for_ready_for_plate_state():
            raise RuntimeError("Instrument failed to enter 'Waiting for Plate' state after command.")
        return jsonify({"status": "success", "message": "Assay initiated.", "assay_id": envision_controller.active_assay_id})
    except Exception as e: return handle_api_error(e, "assay/start_by_index")

@app.route('/assay/save_results', methods=['POST'])
def save_results():
    if not envision_controller: return jsonify({"error": "Controller is offline"}), 503
    try:
        data = request.get_json()
        assay_id = data.get('assay_id')
        filename = data.get('filename')
        if not (isinstance(assay_id, int) and assay_id > 0):
            return jsonify({"error": "A valid, positive assay_id is required"}), 400
        if envision_controller.save_result(assay_id, filename):
            return jsonify({"status": "success", "message": "Results saved."})
        raise RuntimeError(f"save_result for Assay ID {assay_id} returned False.")
    except Exception as e: return handle_api_error(e, "assay/save_results")

@app.route('/assay/save_calculated_results', methods=['POST'])
def save_calculated_results():
    if not envision_controller: return jsonify({"error": "Controller is offline"}), 503
    try:
        data = request.get_json()
        assay_id = data.get('assay_id')
        prefix = data.get('filename_prefix')
        if not (isinstance(assay_id, int) and assay_id > 0):
            return jsonify({"error": "A valid, positive assay_id is required"}), 400
        if envision_controller.save_calculated_results(assay_id, prefix):
            return jsonify({"status": "success", "message": "Calculated results saved."})
        raise RuntimeError(f"save_calculated_results for Assay ID {assay_id} returned False.")
    except Exception as e: return handle_api_error(e, "assay/save_calculated_results")

@app.route('/instrument/load_plate', methods=['POST'])
def load_plate():
    if not envision_controller: return jsonify({"error": "Controller is offline"}), 503
    try:
        barcode = request.get_json().get('barcode')
        if not envision_controller.load_plate(barcode):
            raise RuntimeError("LoadPlate COM call failed.")
        if not envision_controller.getstatus_blocking(target_state=2, timeout_seconds=600): # 10 min timeout
            raise RuntimeError("Timed out waiting for instrument to return to IDLE after load.")
        read_barcode = envision_controller.last_plate_barcode
        return jsonify({"status": "success", "message": f"Operation complete. Read barcode: '{read_barcode}'."})
    except Exception as e: return handle_api_error(e, "instrument/load_plate")

@app.route('/instrument/unload_plate', methods=['POST'])
def unload_plate():
    if not envision_controller: return jsonify({"error": "Controller is offline"}), 503
    try:
        if not bool(getattr(envision_controller.system_raw_com_object, 'IsPlateLoaded', False)):
            raise RuntimeError("Cannot Unload: No plate is currently loaded.")
        if not envision_controller.unload_plate():
            raise RuntimeError("UnloadPlate COM call returned False.")
        if not envision_controller.getstatus_blocking(target_state=2, timeout_seconds=60):
            raise RuntimeError("Timed out waiting for instrument to return to IDLE after unload.")
        return jsonify({"status": "success", "message": "Plate unload complete."})
    except Exception as e: return handle_api_error(e, "instrument/unload_plate")

@app.route('/instrument/state_bitmask', methods=['GET'])
def get_state_bitmask():
    if not envision_controller: return jsonify({"error": "Controller is offline"}), 503
    try:
        mask = envision_controller.get_instrument_state_bitmask()
        return jsonify({"bitmask": mask, "decoded": envision_controller.decode_instrument_state(mask)})
    except Exception as e: return handle_api_error(e, "instrument/state_bitmask")

@app.route('/instrument/configuration_bitmask', methods=['GET'])
def get_config_bitmask():
    if not envision_controller: return jsonify({"error": "Controller is offline"}), 503
    try:
        mask = envision_controller.get_instrument_configuration_bitmask()
        return jsonify({"bitmask": mask, "decoded": envision_controller.decode_instrument_configuration(mask)})
    except Exception as e: return handle_api_error(e, "instrument/configuration_bitmask")

# --- Main Execution ---
if __name__ == '__main__':
    if envision_controller:
        app.run(host='0.0.0.0', port=5000, debug=False, threaded=False)
        print("FATAL: Could not start server because Envision controller failed to initialize.")