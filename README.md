# speech-to-text-dictation

PyQt6 desktop application that transcribes microphone audio in real time and types the result directly into any focused window.

## Screenshot

![Dictation Assistant UI](docs/screenshot.png)

*No screenshot available yet. Place a screenshot at `docs/screenshot.png`.*

## Features

- Hold-to-record via configurable hotkey (default: mouse side button 5) or keyboard shortcut
- Real-time audio spectrum visualizer using matplotlib embedded in the Qt window
- Automatic text injection into the active window via `pyautogui` and `win32clipboard`
- Multilingual recognition (default: `pt-BR`); any language tag supported by Google Speech API can be set
- Continuous recognition mode for hands-free transcription
- Configurable audio quality (`high` default), sample rate (48000 Hz), and sensitivity (0-100)
- System tray icon with quick access to start/stop and settings
- Persistent JSON configuration stored at `~/dictation_assistant_config.json`
- Dark theme UI

## Stack

| Component | Library | Purpose |
|---|---|---|
| UI framework | PyQt6 | Main window, dialogs, system tray |
| Audio capture | PyAudio | Microphone input stream |
| Speech recognition | SpeechRecognition | Google Speech API backend |
| Audio visualization | matplotlib (QtAgg backend) | Live waveform display |
| Hotkey / mouse events | keyboard, pynput | Global hotkey capture |
| Text injection | pyautogui, win32clipboard | Type into active window |
| Win32 integration | pywin32 (win32gui, win32con) | Window focus and clipboard operations |

## Setup / Installation

**Requirements:** Windows, Python 3.9+

```bash
git clone https://github.com/gabrielcnb/speech-to-text-dictation.git
cd speech-to-text-dictation
python -m venv .venv
.venv\Scripts\activate
pip install PyQt6 pyaudio SpeechRecognition numpy matplotlib keyboard pynput pyautogui pywin32
```

An active internet connection is required for the Google Speech API backend.

## Usage

```bash
python dictation_assistant.py
```

The application starts minimized to the system tray.

**Workflow:**
1. Focus any text input (browser, editor, chat application).
2. Hold the configured hotkey (default: `mouse5`).
3. Speak. Release the hotkey to transcribe and inject the text.

**Changing the hotkey:**
Open the settings dialog from the tray icon. Set any key or mouse button under "Hotkey".

**Switching language:**
In settings, change the "Language" field. Examples: `en-US`, `es-ES`, `fr-FR`.

**Config file location:**
```
C:\Users\<user>\dictation_assistant_config.json
```

Default configuration:
```json
{
    "hotkey": "mouse5",
    "language": "pt-BR",
    "sample_rate": 48000,
    "chunk_size": 1024,
    "auto_start": true,
    "sensitivity": 70,
    "theme": "dark",
    "continuous_recognition": false,
    "show_realtime_text": true,
    "audio_quality": "high"
}
```

## File Structure

```
speech-to-text-dictation/
├── dictation_assistant.py   # Main application: UI, audio capture, recognition, text injection
└── README.md
```

## Known Limitations

- Windows only (uses `win32gui`, `win32clipboard`, `ctypes.windll`).
- Requires an internet connection for Google Speech API.
- Global hotkeys may require running as Administrator on some Windows configurations.
