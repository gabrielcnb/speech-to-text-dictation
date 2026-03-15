# Speech to Text Dictation

Windows desktop app for continuous speech dictation. Press and hold a configurable hotkey to record, release to transcribe — text is typed automatically at the cursor position.

## Features

- Configurable hotkey (default: Left Ctrl)
- Auto-paste transcribed text at cursor
- Runs in the system tray — always available
- Statistics and usage history with charts
- PyQt6 GUI for settings

## Stack

- Python 3.10+
- PyQt6 (GUI + system tray)
- SpeechRecognition + PyAudio (transcription)
- pynput (hotkey detection)
- matplotlib (usage charts)
- win32api (Windows clipboard)

## Requirements

```bash
pip install PyQt6 SpeechRecognition pyaudio pynput matplotlib pywin32
```

## Usage

```bash
python dictation_assistant.py
```

The app starts minimized in the system tray. Right-click the tray icon to configure the hotkey or exit.
