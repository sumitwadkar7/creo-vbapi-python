"""
Minimal example: connect to a running Creo Parametric session from Python
via the Creo VB API (COM automation).

Tested on Creo 12.4.1.0 with Python 3.x and pywin32.

Before running:
  1. Register the COM classes once, from an elevated PowerShell:
         cd "C:\\PTC\\Creo 12.4.1.0\\Common Files\\x86e_win64\\obj"
         .\\pfclscom.exe /RegServer
  2. Start Creo and open a model.
  3. Run:  python connect_creo.py
"""

import sys
import win32com.client

# CCpfcAsyncConnection has no usable ProgID, so it must be created by CLSID.
# This CLSID is for Creo 12.4.1.0 -- it may differ on other Creo versions.
# See the README for how to find the CLSID for your build.
ASYNC_CONNECTION_CLSID = "{456E0110-2031-3907-AFE5-9201C97A915E}"


def get_async_connection():
    """Create the async-connection object by CLSID."""
    try:
        return win32com.client.Dispatch(ASYNC_CONNECTION_CLSID)
    except Exception as exc:
        print("Could not create the async connection object.")
        print("Did you run 'pfclscom.exe /RegServer' as Administrator?")
        print(f"Details: {exc}")
        sys.exit(1)


def main():
    async_conn = get_async_connection()

    # Attach to a running Creo session.
    # GetActiveConnection() returns None if Creo was started independently of
    # the API (i.e. you double-clicked it) rather than in async mode.
    connection = async_conn.GetActiveConnection()
    if connection is None:
        print("No active Creo connection found.")
        print("GetActiveConnection() returns None when Creo was started")
        print("independently of the API. See the README for how to start Creo")
        print("from the script instead.")
        sys.exit(1)

    session = connection.Session
    model = session.CurrentModel
    if model is None:
        print("Connected to Creo, but no model is currently open.")
        sys.exit(0)

    try:
        name = model.FileName
    except Exception:
        name = str(model)

    print(f"Connected. Active model: {name}")


if __name__ == "__main__":
    main()
