"""
Connection mixin for AutoCAD adapter.

Handles connection, disconnection, and validation methods.
"""

import logging
import sys
from typing import TYPE_CHECKING, Any

if sys.platform == "win32":
    import win32com.client
    import pythoncom
    import pywintypes
else:
    raise ImportError("AutoCAD adapter requires Windows OS with COM support")

from core import CADConnectionError

if TYPE_CHECKING:
    from core.config import CADConfig

logger = logging.getLogger(__name__)


class ConnectionMixin:
    """Mixin for connection management."""

    if TYPE_CHECKING:
        # Tell type checker this mixin is used with CADAdapterProtocol
        cad_type: str
        config: "CADConfig"
        application: Any
        document: Any

        def _wait_for(
            self, condition: Any, timeout: float = 20.0, interval: float = 0.1
        ) -> bool: ...

    def _find_cad_executable(self, prog_id: str) -> Any:
        """Attempt to locate the standalone CAD executable path from Windows Registry."""
        try:
            import winreg
            from pathlib import Path

            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{prog_id}\\CLSID") as key:
                clsid = winreg.QueryValue(key, "")
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\LocalServer32") as key:
                cmd = winreg.QueryValue(key, "")

            if cmd:
                lower_cmd = cmd.lower()
                if ".exe" in lower_cmd:
                    idx = lower_cmd.index(".exe") + 4
                    exe_cand = cmd[:idx].strip().strip('"')
                    if Path(exe_cand).is_file():
                        return exe_cand
        except Exception as e:
            logger.debug(f"Could not locate CAD executable in registry for {prog_id}: {e}")
        return None

    def connect(self, only_if_running: bool = False) -> bool:
        """Connect to the CAD application via COM, initializing COM for this thread.

        Tries to attach to an already-running instance first. If none is found
        and ``only_if_running`` is False, launches a new instance.

        Args:
            only_if_running: When True, return False instead of launching a new
                CAD instance if none is currently running.

        Returns:
            True if the connection was established successfully.

        Raises:
            CADConnectionError: If COM initialization fails or the ProgID is invalid.
        """
        try:
            logger.info(f"Connecting to {self.cad_type}...")

            # Initialize COM for this thread
            # CoInitialize() may raise if already initialized, which is fine
            try:
                pythoncom.CoInitialize()
            except Exception as e:
                logger.debug(
                    f"CoInitialize: {e} (may already be initialized for thread)"
                )

            # Try to get existing instance
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    self.application = win32com.client.GetActiveObject(
                        self.config.prog_id
                    )
                    logger.info(
                        f"{self.cad_type} instance found (active via GetActiveObject)"
                    )
                    break
                except Exception as e:
                    if attempt == max_retries - 1:
                        logger.debug(
                            f"GetActiveObject for {self.config.prog_id} failed after {max_retries} attempts: {e}"
                        )
                    else:
                        pythoncom.CoInitialize()  # Re-init just in case
                        import time

                        time.sleep(0.5)

            # If not found via GetActiveObject
            if self.application is None:
                if only_if_running:
                    logger.debug(
                        f"{self.cad_type} not running and only_if_running=True. Skipping launch."
                    )
                    return False

                # 1. Prefer launching as a standalone GUI process so CAD does NOT run as a slave /Automation process
                exe_path = self._find_cad_executable(self.config.prog_id)
                if exe_path:
                    try:
                        import subprocess
                        import time

                        logger.info(
                            f"Launching {self.cad_type} as standalone GUI process: {exe_path}"
                        )
                        subprocess.Popen([exe_path])
                        start_time = time.time()
                        wait_limit = min(float(self.config.startup_wait_time), 15.0)
                        while time.time() - start_time < wait_limit:
                            time.sleep(1.0)
                            try:
                                pythoncom.CoInitialize()
                                self.application = win32com.client.GetActiveObject(self.config.prog_id)
                                if self.application is not None:
                                    logger.info(
                                        f"{self.cad_type} connected via GetActiveObject after standalone launch"
                                    )
                                    break
                            except Exception:
                                continue
                    except Exception as e:
                        logger.debug(f"Failed standalone launch for {self.cad_type}: {e}")

                # 2. Fallback to Dispatch if still not connected
                if self.application is None:
                    logger.info(f"{self.cad_type} not active via standalone launch, starting/connecting via Dispatch...")
                    try:
                        self.application = win32com.client.Dispatch(self.config.prog_id)
                    except pywintypes.com_error as com_err:
                        error_code = com_err.args[0] if com_err.args else None
                        if error_code == -2147221005:
                            error_msg = (
                                f"Invalid ProgID '{self.config.prog_id}'. "
                                f"Either {self.cad_type.upper()} is not installed or the ProgID is incorrect. "
                                f"Check config.json and ensure the application is installed."
                            )
                        else:
                            error_msg = str(com_err)
                        logger.error(
                            f"Failed to create {self.cad_type} instance: {error_msg}"
                        )
                        raise CADConnectionError(self.cad_type, error_msg)

                    self._wait_for(
                        lambda: self.application is not None,
                        timeout=self.config.startup_wait_time,
                    )
                    logger.info(
                        f"{self.cad_type} instance connected/started via Dispatch"
                    )

            if self.application is not None:
                # Try to make application visible (not all CAD types support this)
                try:
                    self.application.Visible = True
                except (pywintypes.com_error, AttributeError) as e:
                    logger.debug(
                        f"{self.cad_type} doesn't support Visible property or it's read-only: {e}"
                    )
                # Ensure UserControl is True so CAD remains open after automation client disconnects
                try:
                    if hasattr(self.application, "UserControl"):
                        self.application.UserControl = True
                except Exception as e:
                    logger.debug(f"UserControl property could not be set: {e}")

            # Get active document or create new
            if self.application is not None:
                # Standard AutoCAD/ZWCAD/BricsCAD/GstarCAD handling
                if self.application.Documents.Count > 0:
                    self.document = self.application.ActiveDocument
                    logger.info("Using existing active document")
                else:
                    self.document = self.application.Documents.Add()
                    logger.info("Created new document")

            # Validate connection
            if not self._validate_document():
                raise CADConnectionError(self.cad_type, "Document validation failed")

            logger.info(f"✓ Successfully connected to {self.cad_type}")

            return True

        except pywintypes.com_error as e:
            error_msg = f"COM error: {str(e)}"
            logger.error(f"Failed to connect to {self.cad_type}: {error_msg}")
            raise CADConnectionError(self.cad_type, error_msg)
        except Exception as e:
            logger.error(f"Failed to connect to {self.cad_type}: {e}")
            raise CADConnectionError(self.cad_type, str(e))

    def disconnect(self) -> bool:
        """Disconnect from CAD application with COM cleanup."""
        try:
            if self.application:
                try:
                    if hasattr(self.application, "UserControl"):
                        self.application.UserControl = True
                except Exception:
                    pass
                self.application = None
                self.document = None
            pythoncom.CoUninitialize()
            logger.info(f"Disconnected from {self.cad_type}")
            return True
        except Exception as e:
            logger.error(f"Error disconnecting: {e}")
            return False

    def __enter__(self):
        """Enter context: connect to CAD application.

        Allows using the adapter as a context manager:
            with AutoCADAdapter("autocad") as adapter:
                adapter.draw_line((0,0), (10,10))
                # Auto-disconnect on exit

        Returns:
            Self (the adapter instance)

        Raises:
            CADConnectionError: If connection fails
        """
        if not self.connect():
            raise CADConnectionError(
                self.cad_type, "Connection failed during context manager initialization"
            )
        return self

    def __exit__(self, _exc_type, _exc_val, _exc_tb) -> None:
        """Exit context: disconnect from CAD application.

        Args:
            _exc_type: Exception type if raised
            _exc_val: Exception value if raised
            _exc_tb: Exception traceback if raised
        """
        self.disconnect()

    def is_connected(self) -> bool:
        """Check if connected to CAD application."""
        try:
            return (
                self.application is not None
                and self.document is not None
                and self._validate_document()
            )
        except Exception:
            return False

    def _validate_document(self) -> bool:
        """Validate that document is accessible."""
        try:
            if self.document is None:
                return False
            _ = self.document.Name
            return True
        except Exception:
            return False

    def check_document_change(self) -> bool:
        """
        Check if the active document in the CAD application has changed.
        If it has, update self.document to the new active document.

        Returns:
            bool: True if the document changed, False otherwise.
        """
        try:
            if not self.application:
                return False

            # If no documents are open, there's nothing to check
            if self.application.Documents.Count == 0:
                if self.document is not None:
                    self.document = None
                    return True
                return False

            active_doc = self.application.ActiveDocument

            # If we didn't have a document before, but now we do
            if self.document is None:
                self.document = active_doc
                return True

            # Compare names
            if self.document.Name != active_doc.Name:
                self.document = active_doc
                logger.info(f"Active document changed to: {active_doc.Name}")
                return True

            return False
        except Exception as e:
            # Silently catch COM errors during polling
            logger.debug(f"Error checking document change: {e}")
            return False
