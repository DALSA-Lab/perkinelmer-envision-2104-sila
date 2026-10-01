# envision_sila_driver/server.py
from typing import Optional
from uuid import UUID, uuid4

from sila2.server import SilaServer

from .feature_implementations.envisionplatereader_impl import EnvisionPlateReaderImpl
from .generated.envisionplatereader import EnvisionPlateReaderFeature


class Server(SilaServer):
    def __init__(
        self,
        server_uuid: Optional[UUID] = None,
        name: Optional[str] = None,
        description: Optional[str] = None,
    ):
        if name is None:
            name = "EnvisionPlateReaderServer"
        
        if description is None:
            description = "A SiLA 2 server for the PerkinElmer Envision Plate Reader"

        super().__init__(
            server_name=name,
            server_description=description,
            server_type="PlateReaderController",
            server_version="1.0",
            server_vendor_url="https://gitlab.com/SiLA2/sila_python",
            server_uuid=server_uuid if server_uuid is not None else uuid4(),
        )

        self.envisionplatereader = EnvisionPlateReaderImpl(self)
        self.set_feature_implementation(EnvisionPlateReaderFeature, self.envisionplatereader)