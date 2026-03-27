# Live Dictate

![Python](https://img.shields.io/badge/Python-3.9+-3776AB?logo=python&logoColor=white)
![Windows](https://img.shields.io/badge/Windows-0078D6?logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)

Windows desktop app for speech-to-text dictation. Hold a hotkey, speak, and the transcribed text is typed into any focused window.

## Features

- **Hold-to-record** via configurable hotkey (default: mouse side button) or keyboard shortcut
- **Auto-punctuation** for pt-BR — adds `.` `?` `!` based on context
- **Stutter correction** — removes repeated words from recognition artifacts
- **Real-time preview** — see partial transcription as you speak
- **Audio spectrum visualizer** embedded in the UI
- **Text injection** into any active window via clipboard
- **Accept/Reject workflow** — review text before pasting (Mouse5/Mouse6)
- **Single instance guard** — prevents duplicate processes
- **Multilingual** — default pt-BR, supports any Google Speech API language
- **System tray** with quick access to settings
- **File logging** for troubleshooting (`live_dictate.log`)
- **Windows autostart** option via registry
- **Persistent config** at `~/dictation_assistant_config.json`

## Stack

| Component | Library |
|---|---|
| UI | PyQt6 |
| Audio capture | PyAudio |
| Speech recognition | SpeechRecognition (Google API) |
| Visualization | matplotlib (QtAgg backend) |
| Hotkeys | keyboard, pynput |
| Text injection | pyautogui, pywin32 |

## Getting Started

**Requirements:** Windows 10/11, Python 3.9+, internet connection

```bash
git clone https://github.com/gabrielcnb/live-dictate.git
cd live-dictate
pip install -r requirements.txt
python dictation_assistant.py
```

### Usage

1. Focus any text input (browser, editor, chat)
2. Hold the configured hotkey (default: mouse5)
3. Speak — release to transcribe and inject text
4. Mouse5 to accept, Mouse6 to reject and retry

Settings (hotkey, language, audio quality, theme) are accessible from the tray icon or the main window.

## License

[MIT](LICENSE)
