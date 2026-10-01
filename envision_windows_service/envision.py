import win32com.client
import pythoncom  # For com_error and the Missing object
import time
import os
import logging
from typing import Callable, Optional, List, Dict, Any, Tuple

# Setup logging - Ensure this is configured by the application using this module.
# For standalone testing, uncomment the line below.
# logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(module)s - %(message)s')

# --- Constants (INSTRUMENT_STATES_DESC, INSTRUMENT_CONFIG_DESC, ENUM_CURRENT_STATE_DESC) ---
INSTRUMENT_STATES_DESC = {
    1: "Connected to device", 2: "Initialisation is running", 4: "Initialised (Idle)",
    8: "Shutting down", 16: "Loading the plate", 32: "Unloading the plate",
    64: "Measuring is in process", 128: "Measuring the reference is in process",
    256: "Ejecting filters is in process", 512: "Scanning filters is in process",
    1024: "Cover lid is open", 2048: "Connecting to device is in process",
    4096: "Side lid is open", 8192: "Device is waiting for a plate",
    16384: "Restack the stacker is in process", 32768: "Dispenser maintenance operation is in process",
    65536: "Reset Stacker is in process"
}
INSTRUMENT_CONFIG_DESC = {
    1: "Barcode reading from long side of plate is set", 2: "Barcode reading from short side of plate is set",
    4: "Barcode reading from front side of plate is set", 8: "Ultra Sensitive Luminescence option is set",
    16: "Laser excitation for AlphaScreen is set", 32: "Second detector option is set",
    128: "Enhanced Luminescence option is set", 256: "Stacker for automatic plate loading is set",
    512: "Dispenser pump unit 1 is set", 1024: "Detector temperature controller for AlphaScreen",
    2048: "Plate height sensor for Enhanced Luminescence is set", 4096: "IR-link for height adjustment plate (service tool) is set",
    8192: "Plate cooler has heater also", 16384: "HTS AlphaScreen option is set",
    32768: "Dispenser pump unit 2 is set", 65536: "X-direction cold plate is set",
    131072: "Dispenser hot plate is set", 262144: "Dispenser waste pump is set",
    2**19: "Plate Temperature Sensor is set", 2**20: "Excitation Monochromator is loaded",
    2**21: "Emission Monochromator is loaded", 2**22: "Light source switch is loaded",
    2**23: "Second light source is loaded", 2**24: "Excitation Monochromator dual is loaded"
}
ENUM_CURRENT_STATE_DESC = {
    0: "Connecting to Instrument is in process", 1: "Initialisation is in process",
    2: "Instrument is in IDLE state", 3: "Instrument is loading a plate",
    4: "Instrument is unloading the plate", 5: "Instrument is measuring reference",
    6: "Instrument is in measuring process", 7: "Instrument is shaking the plate",
    8: "Instrument is in waiting modus", 9: "Instrument in pause modus",
    10: "Instrument is waiting for a plate", 11: "Instrument is saving",
    12: "Cover lid is open", 13: "Side lid is open",
    14: "Instrument is scanning the filters", 15: "Instrument is shutting down",
    16: "Instrument does a restack", 17: "Instrument was stopped",
    18: "Instrument is dispensing", 19: "Instrument does a dispenser measuring",
    20: "Instrument is measuring the background", 21: "Instrument does a dispenser maintenance procedure",
    22: "Instrument is waiting for temperature", 23: "Instrument resets the stacker",
    65535: "Undefined state"
}
# --- End Constants ---

class EnvisionEvents:
    def __init__(self, parent_handler: Optional[Any] = None):
        self.parent_handler = parent_handler

    def _handle_event(self, event_name: str, *args):
        if self.parent_handler and hasattr(self.parent_handler, 'handle_instrument_event'):
            try:
                self.parent_handler.handle_instrument_event(event_name, *args)
            except Exception as e:
                logging.error(f"Error in parent_handler.handle_instrument_event for {event_name}: {e}", exc_info=True)
        else:
            logging.debug(f"Event '{event_name}' received by EnvisionEvents, but no parent_handler.handle_instrument_event method: Args: {args}")

    def OnCurrentStateChanged(self, newState: int):
        logging.info(f"Event: EnVision CurrentState Changed to {newState} - {ENUM_CURRENT_STATE_DESC.get(newState, 'Unknown State')}")
        self._handle_event("CurrentStateChanged", newState)

    def OnError(self, Source: str, Number: int, Description: str, Choices: int, Action: Any):
        logging.error(f"Event: EnVision Error - Source: {Source}, Number: {Number}, Description: {Description}, Choices: {Choices}, Action: {Action}")
        self._handle_event("Error", Source, Number, Description, Choices, Action)

    def OnPlateLoaded(self, Barcode: str, Height: int, MaxHeight: int):
        logging.info(f"Event: Plate Loaded - Barcode: '{Barcode}', Height: {Height}, MaxHeight: {MaxHeight}")
        self._handle_event("PlateLoaded", Barcode, Height, MaxHeight)

    def OnPlateBarcodeRead(self, Barcode: str, Acknowledge: bool):
        logging.info(f"Event: Plate Barcode Read - Barcode: '{Barcode}', Acknowledge: {Acknowledge}")
        self._handle_event("PlateBarcodeRead", Barcode, Acknowledge)

    def OnAssayStarted(self, Assay: Any):
        assay_id, protocol_name = "Unknown", "Unknown"
        try:
            if Assay:
                assay_id = getattr(Assay, 'AssayID', "Unknown")
                if hasattr(Assay, 'Protocol') and Assay.Protocol:
                    protocol_name = getattr(Assay.Protocol, 'ProtName', "Unknown")
        except pythoncom.com_error: pass
        logging.info(f"Event: Assay Started - AssayID: {assay_id}, Protocol: {protocol_name}")
        self._handle_event("AssayStarted", Assay)

    def OnAssaySaved(self, Assay: Any):
        assay_id = "Unknown"
        try:
            if Assay: assay_id = getattr(Assay, 'AssayID', "Unknown")
        except pythoncom.com_error: pass
        logging.info(f"Event: Assay Saved - AssayID: {assay_id}")
        self._handle_event("AssaySaved", Assay)

    def OnAssayContinued(self, Assay: Any):
        assay_id = "Unknown"
        try:
            if Assay: assay_id = getattr(Assay, 'AssayID', "Unknown")
        except pythoncom.com_error: pass
        logging.info(f"Event: Assay Continued - AssayID: {assay_id}")
        self._handle_event("AssayContinued", Assay)

    def OnInstrumentStateChanged(self, NewState: int):
        decoded_states = [desc for bit, desc in INSTRUMENT_STATES_DESC.items() if NewState & bit]
        logging.info(f"Event: InstrumentState Changed - Bitmask: {NewState}, Decoded: {decoded_states if decoded_states else 'None'}")
        self._handle_event("InstrumentStateChanged", NewState, decoded_states)

class Envision:
    def __init__(self,
                 system_prog_id: str = "Envision112.WMLR4SystemServer",
                 data_prog_id: str = "Envision112.WMLR4SysDataAdv",
                 event_callback: Optional[Callable[[str, Any], None]] = None):
        self.system_raw_com_object: Optional[win32com.client.CDispatch] = None
        self.system_event_sink: Optional[win32com.client.DispatchWithEvents] = None # For event handling wrapper
        self.data: Optional[win32com.client.CDispatch] = None
        self.assay_object: Optional[Any] = None
        self.user_event_callback = event_callback

        self.default_protocol_index: int = 0
        self.results_subdir: str = "envision_results"
        self.output_dir_path: str = ""

        self.last_known_current_state: int = 65535
        self.last_known_instrument_state_bitmask: int = 0
        self.last_plate_barcode: Optional[str] = None
        self.is_plate_loaded_event_flag: bool = False
        self.active_assay_id: Optional[int] = None
        self.is_assay_running_flag: bool = False
        self.last_error_details: Optional[Dict[str, Any]] = None

        try:
            pythoncom.CoInitialize()
            logging.info("COM Initialized for the current thread.")

            self.system_raw_com_object = win32com.client.Dispatch(system_prog_id)
            self.data = win32com.client.Dispatch(data_prog_id)
            logging.info(f"Successfully dispatched COM objects: {system_prog_id}, {data_prog_id}")

            event_handler_instance = EnvisionEvents(parent_handler=self)
            self.system_event_sink = win32com.client.WithEvents(self.system_raw_com_object, event_handler_instance.__class__)
            logging.info("COM Event handling enabled for SystemServer.")

        except pythoncom.com_error as e:
            logging.error(f"Failed to dispatch COM objects or set up events: {e}", exc_info=True)
            try: pythoncom.CoUninitialize()
            except: pass
            raise ConnectionError(f"Failed to initialize EnVision COM components: {e}")
        except Exception as e_init:
            logging.error(f"Generic error during Envision __init__: {e_init}", exc_info=True)
            try: pythoncom.CoUninitialize()
            except: pass
            raise

        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.output_dir_path = os.path.join(script_dir, self.results_subdir)
        if not os.path.exists(self.output_dir_path):
            try:
                os.makedirs(self.output_dir_path)
                logging.info(f"Created results directory: {self.output_dir_path}")
            except OSError as e:
                logging.error(f"Could not create results directory {self.output_dir_path}: {e}")
                self.output_dir_path = script_dir

    def handle_instrument_event(self, event_name: str, *args: Any):
        logging.debug(f"Envision class received event via internal callback: {event_name} with args: {args}")
        if event_name == "CurrentStateChanged":
            self.last_known_current_state = args[0]
            if args[0] in [2, 17]: self.is_assay_running_flag = False
        elif event_name == "InstrumentStateChanged": self.last_known_instrument_state_bitmask = args[0]
        elif event_name == "PlateLoaded": 
            self.last_plate_barcode = args[0]
            self.is_plate_loaded_event_flag = True
        elif event_name == "PlateBarcodeRead": self.last_plate_barcode = args[0]
        elif event_name == "AssayStarted":
            self.is_assay_running_flag = True
            try:
                if args[0] and hasattr(args[0], 'AssayID'): self.active_assay_id = args[0].AssayID
            except pythoncom.com_error: logging.warning("Could not get AssayID from AssayStarted event.")
        elif event_name == "AssaySaved": self.is_assay_running_flag = False
        elif event_name == "Error":
            self.last_error_details = {"source": args[0], "number": args[1], "description": args[2]}
            self.is_assay_running_flag = False

        if self.user_event_callback:
            try: self.user_event_callback(event_name, *args)
            except Exception as e: logging.error(f"Error in user_event_callback for {event_name}: {e}", exc_info=True)

    def connect(self, user: str = "MLR4Admin", password: str = "MLR4Admin") -> bool:
        if not self.system_raw_com_object or not self.data:
            logging.error("Raw COM objects not initialized. Cannot connect.")
            return False
        try:
            self.data.Logon(user, password)
            self.data.Init(self.system_raw_com_object)
            logging.info("EnVision system connected and initialized with user: %s", user)
            _ = self.system_raw_com_object.CurrentState # Test with the raw COM object
            return True
        except pythoncom.com_error as e:
            logging.error(f"Connection/Login/Init failed: {e}", exc_info=True)
            return False

    def get_current_state(self, server_format: bool = False) -> Any:
        if not self.system_raw_com_object: return "Error: System not initialized" if server_format else -1
        try:
            pythoncom.PumpWaitingMessages()
            state_code = self.system_raw_com_object.CurrentState
            self.last_known_current_state = state_code
            if server_format:
                return ENUM_CURRENT_STATE_DESC.get(state_code, f"Unknown state code: {state_code}")
            return state_code
        except pythoncom.com_error as e:
            logging.error(f"Error getting current state: {e}", exc_info=True)
            return "Error: COM Exception" if server_format else -1

    def getstatus_blocking(self, target_state: int = 2, timeout_seconds: int = 300) -> bool:
        if not self.system_raw_com_object: # Check raw object
            logging.error("System not initialized. Cannot block for status.")
            return False
        logging.info(f"Blocking until system reaches state {target_state} ({ENUM_CURRENT_STATE_DESC.get(target_state, 'Unknown')})...")
        start_time = time.time()
        while True:
            pythoncom.PumpWaitingMessages()
            current_state_val = self.get_current_state()
            if current_state_val == target_state:
                logging.info(f"System reached target state {target_state}.")
                return True
            if time.time() - start_time > timeout_seconds:
                logging.warning(f"Timeout waiting for state {target_state}. Current state: {current_state_val} ({ENUM_CURRENT_STATE_DESC.get(current_state_val)})")
                return False
            time.sleep(0.1)
    
    def wait_for_plate_load_event(self, timeout_seconds: int = 30) -> bool:
        """Blocks until the OnPlateLoaded event flag is set to True."""
        logging.info("Blocking until OnPlateLoaded event is received...")
        start_time = time.time()
        while not self.is_plate_loaded_event_flag:
            pythoncom.PumpWaitingMessages()
            if time.time() - start_time > timeout_seconds:
                logging.warning("Timeout waiting for OnPlateLoaded event.")
                return False
            time.sleep(0.1)
        logging.info("OnPlateLoaded event received.")
        return True

    def get_valid_protocols(self) -> List[Any]:
        if not self.data: return []
        protocols_list = []
        try:
            protocols_com_collection = self.data.GetValidProtocols()
            if protocols_com_collection:
                count = 0
                try: count = protocols_com_collection.Count()
                except TypeError: count = protocols_com_collection.Count
                except AttributeError: logging.error("GetValidProtocols has no Count."); return []
                for i in range(1, count + 1):
                    try: protocols_list.append(protocols_com_collection.Item(i))
                    except pythoncom.com_error as e: logging.warning(f"Error getting valid protocol item {i}: {e}")
            logging.info(f"Retrieved {len(protocols_list)} valid protocols.")
            return protocols_list
        except (pythoncom.com_error, AttributeError) as e:
            logging.error(f"Error getting valid protocols: {e}", exc_info=True); return []

    def get_all_protocols(self) -> Optional[List[Any]]:
        if not self.data: return None
        protocols_list = []
        try:
            all_protocols_com_collection = self.data.AssayProtocols
            if all_protocols_com_collection:
                count = 0
                try: count = all_protocols_com_collection.Count()
                except TypeError: count = all_protocols_com_collection.Count
                except AttributeError: logging.error("AssayProtocols has no Count."); return None
                for i in range(1, count + 1):
                    try: protocols_list.append(all_protocols_com_collection.Item(i))
                    except pythoncom.com_error as e: logging.warning(f"Error getting all_protocols item {i}: {e}")
            logging.info(f"Retrieved {len(protocols_list)} total protocols.")
            return protocols_list
        except (pythoncom.com_error, AttributeError) as e:
            logging.error(f"Error getting all protocols: {e}", exc_info=True); return None

    def get_all_protocols_raw_collection(self) -> Optional[Any]:
        """Returns the raw COM collection object for all protocols."""
        if not self.data:
            logging.error("Data object not initialized. Cannot get all protocols raw collection.")
            return None
        try:
            # The AssayProtocols property returns the COM collection
            return self.data.AssayProtocols
        except (pythoncom.com_error, AttributeError) as e:
            logging.error(f"Error getting raw protocol collection (AssayProtocols): {e}", exc_info=True)
            return None

    def check_protocol_validity(self, protocol_com_object: Any, protocol_name_str: str = "Unknown Protocol") -> Tuple[int, str]:
        if not self.data or not protocol_com_object:
            return 1, "Invalid input for protocol check."
        try:
            assay_prot_id = protocol_com_object.AssayProtocolID
            prot_version = protocol_com_object.ProtVersion
            
            validity_code = self.data.CheckValidityOfProtocol(assay_prot_id, prot_version, pythoncom.Missing, pythoncom.Missing)
            validity_code = int(validity_code)
            
            message = ""
            if validity_code != 0:
                message = self.get_protocol_error_string(validity_code)

            logging.info(f"Protocol '{protocol_name_str}' (ID: {assay_prot_id}, Ver: {prot_version}) validity: Code {validity_code}, Msg: '{message}'")
            return validity_code, message
        except (pythoncom.com_error, AttributeError, ValueError) as e:
            logging.error(f"Error checking protocol validity for '{protocol_name_str}': {e}", exc_info=True)
            return 1, str(e)

    def get_protocol_error_string(self, validity_code_int: int) -> str:
        if not self.data:
            return "Error: Data object not initialized."
        try:
            # Pass pythoncom.Missing for the optional second [Data] parameter.
            return self.data.GetProtocolErrorString(validity_code_int, pythoncom.Missing) or ""
        except pythoncom.com_error as e:
            logging.error(f"COM error in GetProtocolErrorString for code {validity_code_int}: {e}")
            return f"(COM Error: {e})"

    def set_default_protocol_index(self, index: int):
        valid_protocols = self.get_valid_protocols()
        if 0 <= index < len(valid_protocols): self.default_protocol_index = index; logging.info(f"Default protocol index set to: {index}")
        else: logging.warning(f"Cannot set default protocol index to {index}, invalid for {len(valid_protocols)} protocols.")

    def get_default_protocol_obj(self) -> Optional[Any]:
        protocols = self.get_valid_protocols()
        if protocols and 0 <= self.default_protocol_index < len(protocols):
            try:
                default_protocol = protocols[self.default_protocol_index]
                logging.info(f"Default protocol is '{getattr(default_protocol, 'ProtName', 'Unknown')}' at index {self.default_protocol_index}")
                return default_protocol
            except Exception as e: logging.error(f"Error accessing default protocol: {e}", exc_info=True)
        return None

    def start_assay(self, protocol_com_object: Any) -> bool:
        """
        Starts an assay run. This has been simplified to align with the EnVision workflow.
        It creates and starts the assay, which typically unloads any existing plate and
        puts the instrument into a "waiting for plate" state (code 10).
        This function is non-blocking; it returns after sending the start command.
        The calling application is responsible for subsequently issuing a LoadPlate command
        to begin the measurement.
        """
        if not self.system_raw_com_object or not protocol_com_object:
            logging.error("System not initialized or protocol object is None. Cannot start assay.")
            return False

        protocol_name = getattr(protocol_com_object, 'ProtName', 'Unknown Protocol')
        logging.info(f"Attempting to start assay for protocol: '{protocol_name}'")

        try:
            # Per documentation and observed behavior, we just create the assay and start it.
            # This action prepares the instrument for a run and makes it wait for a plate.
            if self.assay_object:
                self.assay_object = None # Clear previous assay object
                
            self.assay_object = self.system_raw_com_object.CreateAssay(protocol_com_object)
            if not self.assay_object:
                logging.error(f"Failed to create assay object for '{protocol_name}'.")
                return False
                
            self.assay_object.Start()
            logging.info(f"Assay.Start() command sent for protocol '{protocol_name}'. The instrument should now be waiting for a plate.")
            self.is_assay_running_flag = True
            try:
                self.active_assay_id = self.assay_object.AssayID
            except pythoncom.com_error:
                self.active_assay_id = None
                logging.warning("Could not retrieve AssayID immediately after starting.")

            # This function is non-blocking. It returns immediately.
            return True

        except (pythoncom.com_error, Exception) as e:
            logging.error(f"An exception occurred while trying to run assay '{protocol_name}': {e}", exc_info=True)
            self.is_assay_running_flag = False
            return False

    def start_default_assay(self) -> bool:
        default_protocol = self.get_default_protocol_obj()
        if not default_protocol:
            logging.error("No default protocol set or found. Cannot start assay.")
            return False
        return self.start_assay(default_protocol)

    def choose_protocol_from_list(self, protocol_list: List[Any]) -> Optional[Any]:
        """Helper function to display a list of protocols and get user choice."""
        if not protocol_list:
            print("No protocols available to choose from.")
            return None
        
        print("\nPlease select a protocol to run:")
        for i, p_obj in enumerate(protocol_list):
            try:
                print(f"{i+1}. {p_obj.ProtName} (ID: {p_obj.AssayProtocolID})")
            except (AttributeError, pythoncom.com_error):
                print(f"{i+1}. <Error reading protocol details>")
        
        try:
            choice = int(input("Enter choice number: "))
            if 1 <= choice <= len(protocol_list):
                return protocol_list[choice - 1]
            else:
                print("Invalid choice.")
        except (ValueError, IndexError):
            print("Invalid input.")
        return None
    
    def choose_protocol_from_details_list(self, protocol_details_list: List[Dict[str, Any]]) -> Tuple[Optional[Dict[str, Any]], None]:
        """Helper function to display a list of protocols from dicts and get user choice."""
        if not protocol_details_list:
            print("No protocols available to choose from.")
            return None, None

        print("\nPlease select a protocol:")
        for i, p_details in enumerate(protocol_details_list):
            try:
                print(f"{i+1}. {p_details.get('name', 'N/A')} (ID: {p_details.get('id', 'N/A')}, Ver: {p_details.get('version', 'N/A')})")
            except Exception as e:
                print(f"{i+1}. <Error reading protocol details from dictionary: {e}>")

        try:
            choice = int(input("Enter choice number: "))
            if 1 <= choice <= len(protocol_details_list):
                return protocol_details_list[choice - 1], None
            else:
                print("Invalid choice.")
        except (ValueError, IndexError):
            print("Invalid input.")
        
        return None, None
        
    def start_main_interface_assay(self) -> bool:
        """Gets valid protocols, asks user to choose, and starts the assay."""
        print("Fetching valid protocols for selection...")
        valid_protocols = self.get_valid_protocols()
        if not valid_protocols:
            print("Could not find any valid protocols to run.")
            return False

        chosen_protocol = self.choose_protocol_from_list(valid_protocols)
        
        if chosen_protocol:
            return self.start_assay(chosen_protocol)
        else:
            print("No protocol selected. Assay run cancelled.")
            return False

    def change_default_protocol_interface(self):
        """Gets valid protocols, asks user to choose, and sets the new default index."""
        print("Fetching valid protocols to set a new default...")
        valid_protocols = self.get_valid_protocols()
        if not valid_protocols:
            print("Could not find any valid protocols.")
            return

        chosen_protocol = self.choose_protocol_from_list(valid_protocols)
        
        if chosen_protocol:
            try:
                chosen_id = chosen_protocol.AssayProtocolID
                chosen_version = chosen_protocol.ProtVersion
                
                new_index = -1
                for i, p in enumerate(valid_protocols):
                    if p.AssayProtocolID == chosen_id and p.ProtVersion == chosen_version:
                        new_index = i
                        break
                
                if new_index != -1:
                    self.set_default_protocol_index(new_index)
                    print(f"\nDefault protocol changed to: '{chosen_protocol.ProtName}'")
                else:
                    print("Could not find the chosen protocol in the list to set the index.")

            except (pythoncom.com_error, AttributeError) as e:
                logging.error(f"Error accessing properties of chosen protocol: {e}")
                print("An error occurred while setting the new default protocol.")
        else:
            print("No protocol selected. Default remains unchanged.")

    def stop_current_assay(self) -> bool:
        if self.assay_object:
            try: self.assay_object.Assay_Stop(); logging.info("Stop sent."); return True
            except pythoncom.com_error as e: logging.error(f"COM error stopping: {e}")
        else: logging.warning("No active assay to stop.")
        return False

    def pause_current_assay(self) -> bool:
        if self.assay_object and self.is_assay_running_flag:
            try: self.assay_object.Pause(); logging.info("Pause sent."); return True
            except pythoncom.com_error as e: logging.error(f"COM error pausing: {e}")
        else: logging.warning("No running assay to pause.")
        return False

    def continue_current_assay(self) -> bool:
        if self.assay_object and self.last_known_current_state == 9: # Is Paused
            try: self.assay_object.Continue(); logging.info("Continue sent."); return True
            except pythoncom.com_error as e: logging.error(f"COM error continuing: {e}")
        else: logging.warning(f"Cannot continue, not paused or no assay (State: {self.last_known_current_state}).")
        return False

    def _prepare_output_path(self, filename: str) -> str:
        return os.path.join(self.output_dir_path, filename)

    def save_result(self, assay_id_to_save: Optional[int] = None, filename: Optional[str] = None) -> bool:
        if not self.data: logging.error("Data object not initialized."); return False
        try:
            aid = assay_id_to_save or self.active_assay_id
            if not aid and hasattr(self.data, 'AssayResults') and hasattr(self.data.AssayResults, 'GetMaxAssayResultID'):
                try:
                    aid = self.data.AssayResults.GetMaxAssayResultID()
                except pythoncom.com_error as e:
                    logging.warning(f"Could not get max assay result ID: {e}")
                    aid = 0

            if not aid or aid == 0: logging.error("No valid Assay ID for saving results."); return False
            
            fname = filename or f"envision_output_assay_{aid}.txt"
            fpath = self._prepare_output_path(fname)
            self.data.AssayResults.InitCollection(f"AssayResultID={aid}")
            
            lines = [f"Results for AssayID: {aid}\n"]
            num_plates = self.data.AssayResults.CountPlateResult(str(aid))
            if num_plates == 0: lines.append(f"No plate data found for AssayID {aid}.\n")

            for p_idx in range(num_plates):
                plate = self.data.AssayResults.ItemPlateResult(str(aid), p_idx)
                lines.append(f"Plate Number: {plate.PlateNumber}, Plate ID: {plate.PlateID}\n")
                num_wells = self.data.AssayResults.CountWellResult(str(aid), p_idx)
                for w_idx in range(num_wells):
                    well = self.data.AssayResults.ItemWellResult(str(aid), p_idx, w_idx)
                    num_res = self.data.AssayResults.CountResult(str(aid), p_idx, w_idx)
                    for r_idx in range(num_res):
                        res = self.data.AssayResults.ItemResult(str(aid), p_idx, w_idx, r_idx)
                        lines.append(f"  Well {well.Well}: MeasA={res.MeasA}\tMeasB={getattr(res, 'MeasB', 'N/A')}\tCalcMeas={res.CalcMeas}\n")
            with open(fpath, "w") as f: f.writelines(lines)
            logging.info(f"Results for AssayID {aid} saved to: {fpath}")
            return True
        except (pythoncom.com_error, AttributeError, OSError) as e:
            logging.error(f"Error saving results for AssayID '{assay_id_to_save}': {e}", exc_info=True)
            return False


    def get_assay_results_collection(self, assay_id: int) -> Optional[List[Dict[str, Any]]]:
        if not self.data: logging.error("Data obj not init."); return None
        try:
            self.data.AssayResults.InitCollection(f"AssayResultID={assay_id}")
            res_data_list = []
            num_plates = self.data.AssayResults.CountPlateResult(str(assay_id))
            for p_idx in range(num_plates):
                plate_obj = self.data.AssayResults.ItemPlateResult(str(assay_id), p_idx)
                p_data: Dict[str, Any] = {"PlateNumber": plate_obj.PlateNumber, "PlateID": plate_obj.PlateID, "Wells": []}
                num_wells = self.data.AssayResults.CountWellResult(str(assay_id), p_idx)
                for w_idx in range(num_wells):
                    well_obj = self.data.AssayResults.ItemWellResult(str(assay_id), p_idx, w_idx)
                    w_data: Dict[str, Any] = {"Well": well_obj.Well, "Measurements": []}
                    num_meas = self.data.AssayResults.CountResult(str(assay_id), p_idx, w_idx)
                    for m_idx in range(num_meas):
                        meas_obj = self.data.AssayResults.ItemResult(str(assay_id), p_idx, w_idx, m_idx)
                        w_data["Measurements"].append({"MeasA": meas_obj.MeasA, "MeasB": getattr(meas_obj, 'MeasB', 'N/A'), "CalcMeas": meas_obj.CalcMeas})
                    p_data["Wells"].append(w_data)
                res_data_list.append(p_data)
            logging.info(f"Retrieved results collection for AssayID {assay_id}")
            return res_data_list
        except (pythoncom.com_error, AttributeError) as e:
            logging.error(f"Error getting results collection for AssayID {assay_id}: {e}", exc_info=True); return None

    def get_calculated_results(self, assay_id: int, plate_id_filter: Optional[int] = None) -> Optional[List[Dict[str, Any]]]:
        if not self.data: logging.error("Data obj not init."); return None
        try:
            p_ids = []
            if plate_id_filter: p_ids.append(plate_id_filter)
            else:
                self.data.AssayResults.InitCollection(f"AssayResultID={assay_id}")
                for p_idx in range(self.data.AssayResults.CountPlateResult(str(assay_id))):
                    p_ids.append(self.data.AssayResults.ItemPlateResult(str(assay_id), p_idx).PlateID)
            if not p_ids: logging.warning(f"No plates for AssayID {assay_id}"); return []
            
            all_calc_res_list = []
            for p_id_val in p_ids:
                self.data.CalcResults.InitCollection(f"PlateID={p_id_val}", "CalcID")
                p_calc_res: Dict[str, Any] = {"PlateID": p_id_val, "Calculations": []}
                for c_idx in range(self.data.CalcResults.Count): 
                    calc_obj = self.data.CalcResults.Item(c_idx)
                    p_calc_res["Calculations"].append({
                        "CalcID": calc_obj.CalcID, "WellID": getattr(calc_obj, 'WellID', 'N/A'),
                        "LabelIndex": getattr(calc_obj, 'LabelIndex', 'N/A'), "CalcMeas": calc_obj.CalcMeas,
                        "Formula": getattr(calc_obj, 'Formula', 'N/A')
                    })
                all_calc_res_list.append(p_calc_res)
            logging.info(f"Retrieved calculated results for AssayID {assay_id}.")
            return all_calc_res_list
        except (pythoncom.com_error, AttributeError) as e:
            logging.error(f"Error getting calculated results for AssayID {assay_id}: {e}", exc_info=True); return None

    def save_calculated_results(self, assay_id: int, filename_prefix: str = "calculated_results", plate_id_filter: Optional[int] = None) -> bool:
        calc_data = self.get_calculated_results(assay_id, plate_id_filter)
        if calc_data is None: return False
        if not calc_data: logging.info(f"No calc results for AssayID {assay_id}"); return True
        fpath = self._prepare_output_path(f"{filename_prefix}_assay_{assay_id}.txt")
        try:
            with open(fpath, "w") as f:
                for p_res in calc_data:
                    f.write(f"--- Plate ID: {p_res['PlateID']} ---\n")
                    if not p_res["Calculations"]: f.write("  No calculations for this plate.\n")
                    for calc in p_res["Calculations"]:
                        f.write(f"  CalcID: {calc.get('CalcID', 'N/A')}, WellID: {calc.get('WellID', 'N/A')}\n"
                                f"  Value: {calc.get('CalcMeas', 'N/A')}, Formula: {calc.get('Formula', 'N/A')}\n\n") # Simplified output
            logging.info(f"Calculated results saved to {fpath}")
            return True
        except OSError as e: logging.error(f"Error saving calculated results: {e}"); return False

    def load_plate(self, barcode_string: Optional[str] = None) -> bool:
        if not self.system_raw_com_object: logging.error("System not init."); return False
        try:
            arg = barcode_string or ""
            success = self.system_raw_com_object.LoadPlate(arg)
            logging.info(f"LoadPlate call with barcode argument '{arg}' returned: {success}")
            return success
        except pythoncom.com_error as e: logging.error(f"COM error LoadPlate: {e}"); return False

    def unload_plate(self) -> bool:
        if not self.system_raw_com_object: logging.error("System not init."); return False
        try:
            success = self.system_raw_com_object.UnloadPlate()
            logging.info(f"UnloadPlate call success: {success}")
            return success
        except pythoncom.com_error as e: logging.error(f"COM error UnloadPlate: {e}"); return False

    def get_instrument_properties(self) -> Dict[str, Any]:
        if not self.system_raw_com_object: return {"Error": "System not initialized"}
        props: Dict[str, Any] = {}
        try:
            pythoncom.PumpWaitingMessages()
            props["CurrentStateCode"] = self.last_known_current_state
            props["CurrentStateDesc"] = ENUM_CURRENT_STATE_DESC.get(self.last_known_current_state, "Unknown")

            props["IsInstrumentServerConnected"] = bool(getattr(self.system_raw_com_object, 'IsInstrumentServerConnected', False))
            is_loaded = bool(getattr(self.system_raw_com_object, 'IsPlateLoaded', False))
            props["IsPlateLoaded_Property"] = is_loaded
            props["IsPlateLoaded_EventDriven"] = self.is_plate_loaded_event_flag
            props["PlateBarcode_EventDriven"] = self.last_plate_barcode
            props["PlateBarcode_Property"] = getattr(self.system_raw_com_object, 'PlateBarcode', 'N/A') if is_loaded else "N/A"
            props["SerialNumber"] = getattr(self.system_raw_com_object, 'SerialNumber', 'N/A')
            props["SimulationMode"] = getattr(self.system_raw_com_object, 'SimulationMode', 'N/A')
            props["InsideTemp_K_x100"] = getattr(self.system_raw_com_object, 'InsideTemp', None)
            props["OutsideTemp_K_x100"] = getattr(self.system_raw_com_object, 'OutsideTemp', None)
            props["Humidity_Percent"] = getattr(self.system_raw_com_object, 'Humidity', None)
            props["AllowUnloadPlate_Setting"] = bool(getattr(self.system_raw_com_object, 'AllowUnloadPlate', False))

            if props.get("InsideTemp_K_x100") not in [None, "N/A"]: props["InsideTemp_C"] = (props["InsideTemp_K_x100"]/100.0)-273.15
            if props.get("OutsideTemp_K_x100") not in [None, "N/A"]: props["OutsideTemp_C"] = (props["OutsideTemp_K_x100"]/100.0)-273.15

            props["LastActiveAssayID"] = self.active_assay_id
            props["IsAssayConsideredRunningScript"] = self.is_assay_running_flag
            props["LastErrorDetails"] = self.last_error_details
            logging.info("Retrieved instrument properties.")
            return props
        except pythoncom.com_error as e:
            logging.error(f"COM Error getting instrument properties: {e}", exc_info=True)
            props["Error_COM"] = str(e)
            return props
        except AttributeError as e:
             logging.error(f"Attribute Error getting instrument properties: {e}", exc_info=True)
             props["Error_Attribute"] = str(e)
             return props
        except Exception as e:
            logging.error(f"Unexpected Error getting instrument properties: {e}", exc_info=True)
            props["Error_Unexpected"] = str(e)
            return props


    def get_instrument_state_bitmask(self) -> Optional[int]:
        if not self.system_raw_com_object: return None
        try: pythoncom.PumpWaitingMessages(); return self.system_raw_com_object.InstrumentState
        except pythoncom.com_error as e: logging.error(f"COM error get_instrument_state_bitmask: {e}"); return None

    def decode_instrument_state(self, state_bitmask: Optional[int]) -> List[str]:
        if state_bitmask is None: return ["State bitmask N/A."]
        active = [desc for bit, desc in INSTRUMENT_STATES_DESC.items() if state_bitmask & bit]
        return active or ["No specific state active or unknown state"]

    def get_instrument_configuration_bitmask(self) -> Optional[int]:
        if not self.system_raw_com_object: return None
        try: return self.system_raw_com_object.InstrumentConfiguration
        except pythoncom.com_error as e: logging.error(f"COM error get_instrument_configuration_bitmask: {e}"); return None

    def decode_instrument_configuration(self, config_bitmask: Optional[int]) -> List[str]:
        if config_bitmask is None: return ["Config bitmask N/A."]
        active = [desc for bit, desc in INSTRUMENT_CONFIG_DESC.items() if config_bitmask & bit]
        return active or ["No specific configuration active or unknown configuration"]

    def shutdown_envision_manager(self, clients_only: bool = False) -> bool:
        if not self.system_raw_com_object: logging.error("System not init."); return False
        try: self.system_raw_com_object.ShutDown(clients_only); logging.info(f"ShutDown({clients_only}) sent."); return True
        except pythoncom.com_error as e: logging.error(f"COM error ShutDown: {e}"); return False

    def __del__(self):
        logging.info("Envision object deleting. Releasing COM objects, uninit COM for thread.")
        self.assay_object = None
        self.data = None
        if self.system_event_sink and hasattr(self.system_event_sink, 'close'):
            try: self.system_event_sink.close(); logging.info("Event sink (system_event_sink) closed.")
            except Exception as e: logging.error(f"Error closing event sink: {e}", exc_info=True)
        self.system_event_sink = None
        self.system_raw_com_object = None
        try: pythoncom.CoUninitialize(); logging.info("COM Uninitialized for the current thread.")
        except Exception as e: logging.error(f"Error during CoUninitialize: {e}", exc_info=True)

if __name__ == '__main__':
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(module)s - %(funcName)s - %(message)s')
    logging.info("Starting Envision direct test...")
    envision_controller: Optional[Envision] = None

    def my_event_logger(event_name: str, *args): # Define the callback
        logging.info(f"APP_CALLBACK - Event: {event_name}, Args: {args}")

    try:
        envision_controller = Envision(event_callback=my_event_logger) # Pass callback
        if envision_controller.connect():
            print("Successfully connected to EnVision.")
            time.sleep(0.5) # Give a moment for any initial events
            pythoncom.PumpWaitingMessages()

            props = envision_controller.get_instrument_properties()
            print("\n--- Initial Properties ---")
            for k, v in props.items(): print(f"  {k}: {v}")
            print("------------------------\n")
        else:
            print("Failed to connect to EnVision.")
    except ConnectionError as e:
        print(f"Connection Error: {e}")
        logging.exception("ConnectionError in main test block.")
    except Exception as e:
        print(f"An unexpected error occurred in main: {e}")
        logging.exception("Unexpected error in main test block.")
    finally:
        if envision_controller:
            print("Cleaning up Envision object in finally...")
            del envision_controller
            envision_controller = None
        try: pythoncom.PumpWaitingMessages()
        except: pass # Ignore errors on final pump during shutdown
        logging.info("Envision direct test finished.")
        print("Exiting test script.")