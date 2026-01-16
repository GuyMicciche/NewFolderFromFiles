# New Folder From Files

A Windows Explorer shell extension that creates a new folder from selected files.

**Select files → Right-click → "New folder with selection"** → Files move into a new folder with inline rename.

![Windows 10/11](https://img.shields.io/badge/Windows-10%2F11-blue) ![64-bit](https://img.shields.io/badge/64--bit-only-lightgrey) ![License](https://img.shields.io/badge/license-MIT-green)

## Features

- Works with any files and folders
- Smart folder naming based on common filename prefix
- Single undo (Ctrl+Z) reverts the entire operation
- Keyboard shortcut: **Ctrl+Alt+N**
- Clean Windows-native integration — no bloat
- System tray helper with toggle option

<picture>
  <img src="https://github.com/user-attachments/assets/cd17a0a7-92de-4383-961d-65234b26e147" alt="Screenshot 2" width="450">
</picture>

<picture>
  <img src="https://github.com/user-attachments/assets/b9ffec23-a049-436e-b4a9-8a97d9a2f1c0" alt="Screenshot 3" height="450">
</picture>

<picture>
  <img src="https://github.com/user-attachments/assets/f480a60b-9811-4449-a475-5fe3cd52130b" alt="Screenshot 1" width="450">
</picture>

<picture>
  <img src="https://github.com/user-attachments/assets/c585bf4d-5205-4902-91b5-1b58042744d4" alt="Screenshot 4" width="200">
</picture>

## Download

Get the latest installer from [**Releases**](https://github.com/GuyMicciche/NewFolderFromFiles/releases).

Run `NewFolderFromFiles-Setup.exe` and follow the prompts.

## Installation (Windows SmartScreen Warning)

1. Download `NewFolderFromFiles-Setup.exe`
2. **Right-click → Properties** → Verify "Digitally signed"
3. Double‑click → **"More info" → "Run anyway"**
4. Done! Self‑signed for development.

## Usage

**Context Menu:**
1. Select files/folders in Explorer
2. Right-click → **New folder with selection**
3. Type a name (or keep the suggested one)

**Keyboard Shortcut:**
1. Select files/folders in Explorer
2. Press **Ctrl+Alt+N**

**Hotkey Helper (system tray):**
- Right-click tray icon to enable/disable hotkey
- Starts automatically with Windows (optional)

## Build from Source

### Requirements

- Windows 10/11 x64
- Visual Studio 2019+ with C++ and ATL workloads
- CMake 3.16+
- Inno Setup 6 (for installer)

### Build

```batch
git clone https://github.com/GuyMicciche/NewFolderFromFiles.git
cd NewFolderFromFiles
mkdir build
cd build
cmake .. -G "Visual Studio 17 2022" -A x64
cmake --build . --config Release
```

Output:
- `build/bin/Release/NewFolderFromFiles.dll`
- `build/bin/Release/NewFolderFromFilesHotkey.exe`

### Create Installer

1. Install [Inno Setup 6](https://jrsoftware.org/isdl.php)
2. Open `installer/setup.iss`
3. Press Ctrl+F9
4. Installer: `build/NewFolderFromFiles-Setup.exe`

### Manual Registration (for testing)

```batch
:: Register (run as admin)
regsvr32 "build\bin\Release\NewFolderFromFiles.dll"

:: Unregister
regsvr32 /u "build\bin\Release\NewFolderFromFiles.dll"
```

## Project Structure

```
NewFolderFromFiles/
├── .github/workflows/     # CI/CD
│   └── build.yml
├── src/
│   ├── dllmain.cpp                           # DLL entry + registration
│   ├── NewFolderFromFilesClassFactory.cpp    # COM class factory
│   ├── NewFolderFromFilesContextMenuHandler.cpp  # Context menu logic
│   ├── HotkeyHelper.cpp                      # Tray app for Ctrl+Alt+N
│   └── *.h
├── installer/
│   └── setup.iss                             # Inno Setup script
├── CMakeLists.txt
├── LICENSE
└── README.md
```

## Uninstall

**Settings** → **Apps** → **New Folder From Files** → **Uninstall**

Or via Control Panel → Programs → Uninstall a program.

## License

MIT
