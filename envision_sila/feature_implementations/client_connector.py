# envision_sila_driver/feature_implementations/client_connector.py
import requests
import json
import sys
import time

# --- Configuration ---
# Private IP for direct link.
WINDOWS_PC_IP = "192.168.99.30"

class EnvisionClient:
    """A client to interact with the Envision Flask server with built-in retry logic."""
    def __init__(self, server_ip=WINDOWS_PC_IP, port=5000):
        self.base_url = f"http://{server_ip}:{port}"
        print(f"Attempting to connect to Envision API server at {self.base_url}...")
        
        max_wait_time = 60
        retry_interval = 5
        start_time = time.time()
        connected = False

        while not connected and (time.time() - start_time) < max_wait_time:
            try:
                self.get_status()
                print("Successfully connected to Envision API server.")
                connected = True
            except ConnectionError:
                remaining_time = int(max_wait_time - (time.time() - start_time))
                if remaining_time > 0:
                    print(f"  > API server not found or refused connection. Retrying in {retry_interval}s... ({remaining_time}s left)")
                    time.sleep(retry_interval)
        
        if not connected:
            print(f"\nFATAL: Could not connect to the Envision API server after {max_wait_time} seconds.")
            print("Please ensure the envision_server.py script is running on the Windows PC.")
            sys.exit(1)

    def _make_request(self, method, endpoint, **kwargs):
        """Helper method to make requests and handle common errors."""
        try:
            kwargs.setdefault('timeout', 180)
            response = requests.request(method, f"{self.base_url}{endpoint}", **kwargs)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.HTTPError as e:
            try:
                server_error = e.response.json().get('error', 'Unknown server error')
            except json.JSONDecodeError:
                server_error = e.response.text
            raise ConnectionError(f"API Server Error: {server_error}")
        except requests.exceptions.RequestException as e:
            # This will catch ConnectionRefusedError and others
            raise ConnectionError(f"Network error communicating with API server: {e}")
        return None

    # --- Client Methods ---
    def get_status(self): return self._make_request('get', '/status')
    def get_properties(self): return self._make_request('get', '/properties')
    def get_valid_protocols(self): return self._make_request('get', '/protocols/valid')
    def initiate_default_assay(self): return self._make_request('post', '/assay/start_default')
    def initiate_assay_by_index(self, index): return self._make_request('post', '/assay/start_by_index', json={'index': index})
    def load_plate_and_wait_for_completion(self, barcode=None): return self._make_request('post', '/instrument/load_plate', json={'barcode': barcode}, timeout=900)
    def unload_plate(self): return self._make_request('post', '/instrument/unload_plate', timeout=60)
    def save_results(self, assay_id, filename=None): return self._make_request('post', '/assay/save_results', json={'assay_id': assay_id, 'filename': filename})
    def set_default_protocol(self, index): return self._make_request('post', '/protocols/default', json={'index': index})
    def get_state_bitmask(self): return self._make_request('get', '/instrument/state_bitmask')
    def get_config_bitmask(self): return self._make_request('get', '/instrument/configuration_bitmask')