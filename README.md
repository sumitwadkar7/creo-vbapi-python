# creo-vbapi-python

Connect to **PTC Creo Parametric** from **Python** using the Creo VB API (COM automation).


## Why this exists

Creo ships a VB API that is COM-based, but connecting to it from an *external*
Python process is badly documented. The async-connection class you need is not
exposed under a normal ProgID, so the usual `CreateObject("pfcls...")` call
simply fails with no useful error. This repo documents the exact steps that
actually work — registering the API, the CLSID workaround, and a minimal
connection example.

Tested on **Creo 12.4.1.0**, Windows, Python 3.x.

## What you need

- Creo Parametric with the **VB API installed**. It lives under
  `<CreoInstall>\Common Files\vbapi` (you should see `vbug.pdf` and an
  `examples` folder there).
- Python 3.x on Windows.
- `pywin32` → `pip install pywin32`
- Creo's async messaging executable present: `pfclscom.exe` under
  `<CreoInstall>\Common Files\x86e_win64\obj`, and the environment variable
  `PRO_COMM_MSG_EXE` set (Creo's installer normally sets this for you).

## Step 1 — Register the COM classes

The `pfcls.*` COM classes are often not registered with Windows out of the box.
Register them **once**, from an **elevated (Administrator) PowerShell**:

```powershell
cd "C:\PTC\Creo 12.4.1.0\Common Files\x86e_win64\obj"
.\pfclscom.exe /RegServer
```

Adjust the path to match your Creo version. A successful run registers the
`pfcls.*` ProgIDs. (To undo it later: `.\pfclscom.exe /UnRegServer`.)

## Step 2 — The CLSID gotcha

Even after registration, the async-connection class `CCpfcAsyncConnection` has
**no usable ProgID**, so `CreateObject("pfcls.CCpfcAsyncConnection")` will not
work. You have to create it by its **CLSID**.

On Creo 12.4.1.0 the CLSID is:

```
{456E0110-2031-3907-AFE5-9201C97A915E}
```

> ⚠️ **This CLSID is tied to the Creo version.** On a different Creo build it
> may be different. To find yours, open the pfcls type library in the
> **OLE/COM Object Viewer** (`oleview.exe`, ships with the Windows SDK) and look
> for the `CCpfcAsyncConnection` coclass, or dump the type library and search
> for it.

## Step 3 — Connect

Start Creo, open a model, then run:

```powershell
python connect_creo.py
```

> **Note:** `GetActiveConnection()` only attaches to a Creo session that was
> started in async mode / started *by the API*. If you launch Creo normally and
> independently, it may return `None`. In that case, start Creo from the script
> using the async `Start(...)` method instead — see `vbug.pdf` in your VB API
> folder for the exact call for your version.

## Troubleshooting

- **`Dispatch` fails on the CLSID** → make sure Step 1 ran as Administrator, and
  that you're running Python as the same Windows user. As a fallback you can
  create the object with `pythoncom.CoCreateInstance(...)`.
- **`GetActiveConnection()` returns None** → Creo was started independently of
  the API (see the note above).
- **Wrong CLSID** → you're likely on a different Creo version; re-check Step 2.

## License

MIT — see [LICENSE](LICENSE).

## Contributing

Found the CLSID for a different Creo version, or got the "attach to a
running session" path working cleanly? PRs and issues welcome — the goal is a
version-by-version reference that saves the next person the guesswork.

"Not affiliated with or endorsed by PTC. Creo and PTC are trademarks of their respective owners."