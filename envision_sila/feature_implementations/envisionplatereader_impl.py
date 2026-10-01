# envision_sila_driver/feature_implementations/envisionplatereader_impl.py
from __future__ import annotations
import json
from typing import TYPE_CHECKING, List

from sila2.server import MetadataDict, ObservableCommandInstance
from sila2.framework.errors.defined_execution_error import DefinedExecutionError

from ..generated.envisionplatereader import (
    EnvisionPlateReaderBase,
    LoadPlate_Responses,
    ProtocolInfo,
    InstrumentDetailsType,
    RunAssayByIndex_Responses,
    RunDefaultAssay_Responses,
    SaveResults_Responses,
    SetDefaultProtocol_Responses,
    StateType,
    UnloadPlate_Responses,
    OperationFailed
)
from .client_connector import EnvisionClient

if TYPE_CHECKING:
    from ..server import Server

class EnvisionPlateReaderImpl(EnvisionPlateReaderBase):
    def __init__(self, parent_server: Server) -> None:
        super().__init__(parent_server=parent_server)
        self.update_CurrentState(StateType(StatusCode=-1, StatusDescription="Initializing..."))
        self.update_InstrumentDetails(
            InstrumentDetailsType(PropertiesJSON="{}", DecodedState=["Initializing..."], DecodedConfiguration=["Initializing..."])
        )
        self.client = EnvisionClient()
        self.run_periodically(self._update_observable_properties, delay_seconds=2)

    def _update_observable_properties(self):
        try:
            status = self.client.get_status()
            if status: self.update_CurrentState(StateType(status['status_code'], status['status_description']))
            details = self._get_instrument_details_helper()
            if details: self.update_InstrumentDetails(details)
        except ConnectionError:
            self.update_CurrentState(StateType(-1, "Error: API Server Unreachable"))

    def _get_instrument_details_helper(self) -> InstrumentDetailsType | None:
        try:
            props = self.client.get_properties()
            state_mask = self.client.get_state_bitmask()
            config_mask = self.client.get_config_bitmask()
            if not all([props, state_mask, config_mask]): return None
            return InstrumentDetailsType(json.dumps(props, indent=2), state_mask.get('decoded', []), config_mask.get('decoded', []))
        except ConnectionError:
            return None

    def get_InstrumentDetails(self, *, metadata: MetadataDict) -> InstrumentDetailsType:
        details = self._get_instrument_details_helper()
        if details is None: raise OperationFailed("Could not fetch instrument details.")
        return details

    def get_ValidProtocols(self, *, metadata: MetadataDict) -> List[ProtocolInfo]:
        protocols_data = self.client.get_valid_protocols()
        if protocols_data is None: raise OperationFailed("Could not fetch valid protocols.")
        return [ProtocolInfo(Index=i, Name=p['name'], ProtocolID=str(p['id']), Version=str(p['version'])) for i, p in enumerate(protocols_data)]

    def SaveResults(self, AssayID: int, FileName: str, *, metadata: MetadataDict) -> SaveResults_Responses:
        response = self.client.save_results(AssayID, FileName or None)
        if response and response.get('status') == 'success':
            return SaveResults_Responses(StatusMessage=response.get('message', 'Success'))
        else:
            raise OperationFailed(f"Failed to save results: {response}")

    def SetDefaultProtocol(self, ProtocolIndex: int, *, metadata: MetadataDict) -> SetDefaultProtocol_Responses:
        response = self.client.set_default_protocol(ProtocolIndex)
        if response and response.get('status') == 'success':
            return SetDefaultProtocol_Responses(StatusMessage=response.get('message', 'Success'))
        else:
            raise OperationFailed(f"Failed to set default protocol: {response}")

    def RunDefaultAssay(self, *, metadata: MetadataDict, instance: ObservableCommandInstance) -> RunDefaultAssay_Responses:
        instance.begin_execution() 
        try:
            response = self.client.initiate_default_assay()
            if response and response.get('status') == 'success':
                assay_id = response.get('assay_id', 0)
                return RunDefaultAssay_Responses(AssayID=assay_id)
            else:
                raise OperationFailed(f"API call failed: {response.get('error', 'Unknown error')}")
        except Exception as e:
            raise OperationFailed(str(e))

    def RunAssayByIndex(self, ProtocolIndex: int, *, metadata: MetadataDict, instance: ObservableCommandInstance) -> RunAssayByIndex_Responses:
        instance.begin_execution()
        try:
            response = self.client.initiate_assay_by_index(ProtocolIndex)
            if response and response.get('status') == 'success':
                assay_id = response.get('assay_id', 0)
                return RunAssayByIndex_Responses(AssayID=assay_id)
            else:
                raise OperationFailed(f"API call failed: {response.get('error', 'Unknown error')}")
        except Exception as e:
            raise OperationFailed(str(e))

    def LoadPlate(self, Barcode: str, *, metadata: MetadataDict, instance: ObservableCommandInstance) -> LoadPlate_Responses:
        instance.begin_execution()
        try:
            response = self.client.load_plate_and_wait_for_completion(Barcode)
            if response and response.get('status') == 'success':
                return LoadPlate_Responses()
            else:
                raise OperationFailed(f"API call failed: {response.get('error', 'Unknown error')}")
        except Exception as e:
            raise OperationFailed(str(e))

    def UnloadPlate(self, *, metadata: MetadataDict, instance: ObservableCommandInstance) -> UnloadPlate_Responses:
        instance.begin_execution()
        try:
            response = self.client.unload_plate()
            if response and response.get('status') == 'success':
                return UnloadPlate_Responses()
            else:
                raise OperationFailed(f"API call failed: {response.get('error', 'Unknown error')}")
        except Exception as e:
            raise OperationFailed(str(e))