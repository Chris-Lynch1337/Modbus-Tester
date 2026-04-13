# Modbus TCP Output Tester  v2.2.0

Production-ready desktop utility for testing Modbus TCP holding registers.
Designed for AutomationDirect Productivity Suite (6x / 400001+ addressing).

## Project Structure

```
modbus_tester_app/
  main.py                  <- Entry point
  requirements.txt
  modbus_tester/
    __init__.py
    constants.py           <- Colours, stylesheet, DTYPE_OPTIONS, app metadata
    datatypes.py           <- pack_value, decode_words, dataclasses, validation
    workers.py             <- ConnectionWorker, CommandProcessor
    main_window.py         <- ModbusTester QMainWindow (~950 lines)
    ui/
      __init__.py
      dialogs.py           <- AboutDialog, ColorButton
      register_tab.py      <- Holding Registers tab
      batch_tab.py         <- Ramp / Batch tab
      sweep_tab.py         <- Tag Sweep tab
      settings_tab.py      <- Settings tab
```

## Setup

```bash
pip install -r requirements.txt
python main.py
```

## Build Standalone EXE

```bash
pip install pyinstaller
python -m PyInstaller --onefile --windowed --icon=tester.ico --add-data "tester.ico;." main.py
```

Rename dist/main.exe to ModbusTester.exe.

### Automated Release Script (Windows PowerShell)

For a one-command build that installs dependencies, runs unit tests, and invokes PyInstaller using `main.spec`, run:

```powershell
./build_release.ps1
```

The script recreates the `build/` and `dist/` folders and leaves the final executable in `dist/ModbusTester.exe`.

## Addressing

All fields use Productivity Suite 6x notation (400001+).
The app subtracts 400001 to get zero-based Modbus address.

| Productivity Suite | Modbus |
|--------------------|--------|
| 400001             | 0      |
| 400003             | 2      |
| 401301             | 1300   |

## Word Order for Productivity Suite

INT32 / DINT tags:  UINT32 Lo/Hi (CD AB)
REAL tags:          FLOAT32 Lo/Hi (CD AB)
