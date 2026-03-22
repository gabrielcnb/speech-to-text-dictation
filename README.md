# Speech-to-Text Dictation

![Python](https://img.shields.io/badge/Python-3.9+-3776AB?logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)

Desktop dictation app for Windows. Hold a hotkey, speak, and the transcribed text is typed directly into any focused window.

Built with PyQt6 and Google Speech Recognition API.

## Features

- Hold-to-record via configurable hotkey (default: mouse side button 5) or keyboard shortcut
- Real-time audio spectrum visualizer (matplotlib embedded in Qt)
- Automatic text injection into the active window via pyautogui + win32clipboard
- Multilingual recognition (default: pt-BR, supports any Google Speech API language tag)
- Continuous recognition mode for hands-free transcription
- Configurable audio quality, sample rate (48 kHz default), and sensitivity (0-100)
- System tray icon with quick access to start/stop and settings
- Persistent JSON configuration at `~/dictation_assistant_config.json`
- Dark theme UI

## Stack

| Component | Library |
|---|---|
| UI | PyQt6 |
| Audio capture | PyAudio |
| Speech recognition | SpeechRecognition (Google API) |
| Visualization | matplotlib (QtAgg backend) |
| Hotkeys | keyboard, pynput |
| Text injection | pyautogui, win32clipboard |
| Win32 integration | pywin32 (win32gui, win32con) |

## Getting Started

**Requirements:** Windows, Python 3.9+, internet connection

```bash
git clone https://github.com/gabrielcnb/speech-to-text-dictation.git
cd speech-to-text-dictation
pip install -r requirements.txt
python dictation_assistant.py
```

The app starts minimized to the system tray.

### Usage

1. Focus any text input (browser, editor, chat).
2. Hold the configured hotkey (default: mouse5).
3. Speak. Release to transcribe and inject text.

Settings (hotkey, language, audio quality) are accessible from the tray icon.

## License

[MIT](LICENSE)
