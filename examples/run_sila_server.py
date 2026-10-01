import sys
import signal
import contextlib
from pathlib import Path

# Allow running this file directly from the repo root
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from envision_sila.server import Server

def main() -> None:
    host = "127.0.0.1"
    port = 50052

    print("Initializing EnVision SiLA Server...")
    server = Server(
        name="EnVision Local Server",
        description="Local test server for the PerkinElmer EnVision"
    )

    server.start_insecure(host, port, enable_discovery=False)
    print(f"SiLA server is running on {host}:{port} (insecure)")
    print("Press Ctrl+C to stop.")

    signal.signal(signal.SIGTERM, lambda *args: server.grpc_server.stop())

    with contextlib.suppress(KeyboardInterrupt):
        server.grpc_server.wait_for_termination()

    server.stop()
    print("SiLA server stopped")

if __name__ == "__main__":
    main()