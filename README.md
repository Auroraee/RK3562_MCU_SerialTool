# RK3562 MCU UART Validation Tool

A Windows desktop GUI tool that simulates the RK3562 SoC side of a UART serial protocol for testing MCU firmware. Built with Python 3.10 and tkinter.

## Features

- Serial port connection with configurable baud rate
- 16+ command types (device queries, status reports, LED/motor control, heartbeat, etc.)
- Structured payload input via modal dialogs
- Real-time hex log with send/receive filtering
- Automatic CRC-8/SAE-J1850 checksum calculation
- Optional heartbeat keep-alive (1s interval)
- Log export to file
- Bilingual UI (Chinese / English)

## Requirements

- Python 3.10+
- `pyserial`

```bash
pip install pyserial
```

## Usage

```bash
python rk3562_uart_tester.py
```

1. Select the COM port and baud rate from the toolbar
2. Click **Connect** to open the serial port
3. Choose a command from the left panel and click **Send**
4. View sent/received frames in the right log panel

## Build Executable

Requires `pyinstaller` and `pillow`:

```bash
pip install pyinstaller pillow
```

Then build with the release script (Windows PowerShell):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build-release.ps1 -Version X.Y.Z
```

The output executable will be in the `release/` directory.

## Protocol

Frame format:

```
Header(0x5A) | FrameType(1) | SerialCount(2) | CMD(2) | Length(2) | Payload(N) | CRC8(1) | End(0xA5)
```

- **Frame types:** No ACK (0x01), Need ACK (0x02), ACK OK (0x03), ACK Error (0x04)
- **CRC-8:** SAE-J1850 (Poly=0x1D, Init=0xFF, XorOut=0xFF)

## License

[MIT](LICENSE)
