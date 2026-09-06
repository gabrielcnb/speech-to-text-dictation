import sys
import os
import logging
import numpy as np
import pyaudio
import threading
import time
import keyboard
import pyautogui
import speech_recognition as sr
from PyQt6.QtWidgets import (QApplication, QMainWindow, QSystemTrayIcon, QMenu, 
                            QLabel, QPushButton, QVBoxLayout, QHBoxLayout, QWidget,
                            QSlider, QComboBox, QLineEdit, QCheckBox, QSpinBox,
                            QDialog, QTabWidget, QGridLayout, QGroupBox, QMessageBox,
                            QProgressBar, QFrame)
from PyQt6.QtGui import QIcon, QAction, QColor, QPainter, QPen, QFont, QKeySequence, QPalette, QPixmap
from PyQt6.QtCore import Qt, QTimer, QSize, pyqtSignal, QThread, QRect, QPropertyAnimation, QEasingCurve
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg
import matplotlib.animation as animation
from matplotlib.figure import Figure
import win32gui
import win32con
import win32clipboard
import ctypes
import json
from pynput import mouse

# ============== LOGGING ==============
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(SCRIPT_DIR, "live_dictate.log")

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
log = logging.getLogger("LiveDictate")


# ============== SINGLE INSTANCE ==============
def ensure_single_instance():
    """Prevents multiple instances via Windows mutex."""
    mutex_name = "LiveDictate_SingleInstance_Mutex"
    _mutex = ctypes.windll.kernel32.CreateMutexW(None, False, mutex_name)
    if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
        ctypes.windll.kernel32.CloseHandle(_mutex)
        from tkinter import messagebox
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        messagebox.showwarning("Live Dictate",
            "The app is already running.\nCheck the system tray icon.")
        root.destroy()
        sys.exit(0)
    return _mutex


# ============== TEXT POST-PROCESSING ==============
def fix_stutter(text):
    """Removes consecutive repeated words (speech recognition artifact)."""
    if not text:
        return text
    words = text.split()
    result = []
    i = 0
    while i < len(words):
        current = words[i].lower().strip('.,!?')
        j = i + 1
        while j < len(words) and words[j].lower().strip('.,!?') == current:
            j += 1
        result.append(words[i])
        i = j
    return ' '.join(result)


def auto_punctuate(text):
    """Adds punctuation and capitalizes first letter for pt-BR text."""
    if not text:
        return text
    text = text.strip()
    text = text[0].upper() + text[1:] if len(text) > 1 else text.upper()
    if text[-1] in '.!?':
        return text

    lower = text.lower()
    words = lower.split()

    question_starters = [
        'o que', 'oque', 'como', 'quando', 'onde', 'por que', 'porque',
        'qual', 'quais', 'quem', 'quanto', 'quantos', 'quantas',
        'sera', 'seria', 'pode', 'posso', 'podemos',
    ]
    for q in question_starters:
        if lower.startswith(q):
            return text + '?'

    question_enders = [
        'ne', 'certo', 'sim', 'nao', 'hein', 'ein',
        'mesmo', 'sabe', 'entende', 'entendeu', 'ta', 'ok',
    ]
    if words and words[-1].strip('.,!?') in question_enders:
        return text + '?'

    exclamations = [
        'nossa', 'caramba', 'uau', 'wow', 'legal', 'incrivel',
        'demais', 'otimo', 'perfeito', 'excelente', 'maravilhoso',
        'obrigado', 'obrigada', 'valeu', 'parabens',
    ]
    for e in exclamations:
        if e in lower:
            return text + '!'

    return text + '.'


# Application configuration
CONFIG_FILE = os.path.join(os.path.expanduser('~'), 'dictation_assistant_config.json')
DEFAULT_CONFIG = {
    'hotkey': 'mouse5',
    'language': 'pt-BR',
    'sample_rate': 48000,  # Aumentado para melhor qualidade
    'chunk_size': 1024,
    'auto_start': True,
    'sensitivity': 70,  # Raised so fast speech is picked up better
    'theme': 'dark',
    'continuous_recognition': False,
    'show_realtime_text': True,
    'audio_quality': 'high',  # High by default, for better recognition
}

class Config:
    def __init__(self):
        self.data = DEFAULT_CONFIG.copy()
        self.load()
    
    def load(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    saved_config = json.load(f)
                    self.data.update(saved_config)
            except Exception as e:
                log.error(f"Failed to load settings: {e}")
    
    def save(self):
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, indent=4, ensure_ascii=False)
        except Exception as e:
            log.error(f"Failed to save settings: {e}")
    
    def get(self, key):
        return self.data.get(key, DEFAULT_CONFIG.get(key))
    
    def set(self, key, value):
        self.data[key] = value
        self.save()

# Audio spectrum visualiser
class AudioSpectrumCanvas(FigureCanvasQTAgg):
    def __init__(self, parent=None, width=5, height=2, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = self.fig.add_subplot(111)
        super(AudioSpectrumCanvas, self).__init__(self.fig)
        self.setParent(parent)
        
        # Chart styling
        self.fig.patch.set_facecolor('#2E2E2E')
        self.axes.set_facecolor('#2E2E2E')
        self.axes.spines['top'].set_visible(False)
        self.axes.spines['right'].set_visible(False)
        self.axes.spines['bottom'].set_visible(False)
        self.axes.spines['left'].set_visible(False)
        self.axes.tick_params(axis='both', colors='#CCCCCC')
        
        # Initial setup
        self.x = np.arange(0, 100)
        self.y = np.zeros(100)
        self.line, = self.axes.plot(self.x, self.y, '-', lw=2, color='#00AAFF')
        
        # Axis limits
        self.axes.set_ylim(-0.5, 0.5)
        self.axes.set_xlim(0, 100)
        self.axes.set_xticks([])
        self.axes.set_yticks([])
        
        self.fig.tight_layout(pad=0)
        self.setMinimumHeight(120)
    
    def update_plot(self, data):
        # Refresh the view with new data
        if len(data) > 0:
            # Normaliza os dados
            normalized_data = np.frombuffer(data, dtype=np.int16).astype(np.float32)
            normalized_data = normalized_data / 32768.0  # Normalise to -1.0..1.0
            
            # Only refresh the most recent data
            data_len = min(len(normalized_data), 100)
            self.y = np.roll(self.y, -data_len)
            self.y[-data_len:] = normalized_data[:data_len]
            
            # Update the chart line
            self.line.set_ydata(self.y)
            self.draw()

# Audio processing on its own thread
class AudioProcessor(QThread):
    audio_data = pyqtSignal(bytes)
    text_ready = pyqtSignal(str)
    partial_text = pyqtSignal(str)
    progress_update = pyqtSignal(int)
    
    def __init__(self, config):
        super().__init__()
        self.config = config
        self.recording = False
        self.recorded_data = []
        
        # Recogniser set up with tuned settings
        self.recognizer = sr.Recognizer()
        # Aggressive initial settings for voice detection
        self.recognizer.energy_threshold = 300  # Valor baixo para captar fala mais suave
        self.recognizer.dynamic_energy_threshold = True
        self.recognizer.dynamic_energy_adjustment_damping = 0.1  # Faster response (default is 0.15)
        self.recognizer.dynamic_energy_ratio = 1.1  # More sensitive to change (default is 1.5)
        self.recognizer.pause_threshold = 0.3  # Tempo menor entre frases
        self.recognizer.phrase_threshold = 0.1  # Mais agressivo para detectar frases
        self.recognizer.non_speaking_duration = 0.2  # Requires less silence
        
        self.continuous_mode = config.get('continuous_recognition')
        self.show_realtime = config.get('show_realtime_text')
        self.last_partial_text = ""
        self.update_counter = 0
        
        # Audio quality tuned for fast speech
        self.sample_rate = config.get('sample_rate')
        # Pinned to 16kHz, the rate Google Speech API recommends
        self.sample_rate = 16000  # Taxa ideal para reconhecimento de fala
    
    def run(self):
        # Inicializa PyAudio
        p = pyaudio.PyAudio()
        
        # List available devices, for debugging
        print("Available audio devices:")
        for i in range(p.get_device_count()):
            dev = p.get_device_info_by_index(i)
            if dev['maxInputChannels'] > 0:  # Only show devices with an input
                log.debug(f"[{i}] {dev['name']}")
                
        # Try the default input device
        try:
            default_device_info = p.get_default_input_device_info()
            device_index = int(default_device_info['index'])
            log.info(f"Using default audio device: {default_device_info['name']}")
        except Exception as e:
            log.error(f"Failed to get the default device: {e}")
            # Busca algum dispositivo de entrada
            device_index = None
            for i in range(p.get_device_count()):
                dev = p.get_device_info_by_index(i)
                if dev['maxInputChannels'] > 0:
                    device_index = i
                    break
        
        # Set up the input stream
        try:
            if device_index is not None:
                stream = p.open(
                    format=pyaudio.paInt16,
                    channels=1,
                    rate=self.sample_rate,
                    input=True,
                    frames_per_buffer=self.config.get('chunk_size'),
                    input_device_index=device_index
                )
                log.info(f"Stream aberto com sucesso usando dispositivo {device_index}")
            else:
                stream = p.open(
                    format=pyaudio.paInt16,
                    channels=1,
                    rate=self.sample_rate,
                    input=True,
                    frames_per_buffer=self.config.get('chunk_size')
                )
                log.info("Stream opened on the default device")
        except Exception as e:
            log.error(f"Erro ao abrir stream: {e}")
            self.progress_update.emit(100)
            self.text_ready.emit("")
            self.partial_text.emit("Could not open the microphone. Check your settings.")
            return
            
        # Ready to record
        self.recording = True
        self.recorded_data = []
        
        # Record the audio
        self.partial_text.emit("Recording...")
        self.progress_update.emit(10)
        
        while self.recording:
            try:
                data = stream.read(self.config.get('chunk_size'), exception_on_overflow=False)
                self.audio_data.emit(data)
                
                if self.recording:
                    self.recorded_data.append(data)
                    
                    # Melhorado: Feedback em tempo real mais detalhado
                    if self.show_realtime:
                        self.update_counter += 1
                        # Refresh the partial view more often (every 10 chunks instead of 15)
                        if self.update_counter % 10 == 0:
                            try:
                                # Use the last 3 seconds, which catches fast phrases better
                                recent_data = self.recorded_data[-45:]
                                if recent_data:
                                    audio_data = b''.join(recent_data)
                                    audio = sr.AudioData(audio_data, self.sample_rate, 2)
                                    
                                    # Slightly longer timeout, which improves partial recognition
                                    try:
                                        # Settings specific to partial recognition of fast speech
                                        self.recognizer.operation_timeout = 1.5
                                        temp_recognizer = sr.Recognizer()
                                        temp_recognizer.energy_threshold = self.recognizer.energy_threshold
                                        temp_recognizer.pause_threshold = 0.4  # Shorter still, for fast speech
                                        
                                        partial_text = temp_recognizer.recognize_google(
                                            audio, 
                                            language=self.config.get('language'),
                                            show_all=False
                                        )
                                        if partial_text:
                                            self.last_partial_text = partial_text
                                            self.partial_text.emit(f"Ouvindo: {partial_text}")
                                    except:
                                        # On failure, keep showing the last good partial transcript
                                        if self.last_partial_text:
                                            self.partial_text.emit(f"Ouvindo: {self.last_partial_text}...")
                                        else:
                                            self.partial_text.emit("Ouvindo...")
                            except Exception as partial_e:
                                log.debug(f"Partial transcription failed: {partial_e}")
            except Exception as e:
                log.error(f"Erro ao ler do stream: {e}")
                break
        
        # Recording finished
        try:
            stream.stop_stream()
            stream.close()
        except:
            pass
        
        p.terminate()
        
        if len(self.recorded_data) > 0:
            self.process_audio()
    
    def stop(self):
        self.recording = False
    
    def process_audio(self):
        try:
            self.partial_text.emit("Processing audio...")
            self.progress_update.emit(50)
            
            # Convert the recorded audio to text
            audio_data = b''.join(self.recorded_data)
            audio = sr.AudioData(audio_data, self.sample_rate, 2)  # 2 bytes por sample (16 bits)
            
            # Tuned for recognising fast speech in Portuguese
            # Aggressive settings to improve recognition
            self.recognizer.pause_threshold = 0.3  # Lowered further to cope with very fast speech
            self.recognizer.operation_timeout = 30  # Aumentado para dar mais tempo ao processamento completo
            self.recognizer.energy_threshold = 300  # Valor baixo para detectar fala mais suave
            
            try:
                # Recognition through the Google Speech API, with specific tweaks
                self.partial_text.emit("Tentando reconhecer texto...")
                self.progress_update.emit(70)
                
                # Several recognition attempts with different settings
                try:
                    # First attempt: normal settings
                    text = self.recognizer.recognize_google(
                        audio, 
                        language=self.config.get('language'),
                        show_all=False,
                    )
                except sr.UnknownValueError:
                    # Second attempt: alternative settings
                    # Brief pause so the API can restart
                    time.sleep(0.5)
                    
                    # Build a new recogniser for the second attempt
                    backup_recognizer = sr.Recognizer()
                    backup_recognizer.pause_threshold = 0.2
                    backup_recognizer.operation_timeout = 20
                    
                    text = backup_recognizer.recognize_google(
                        audio, 
                        language=self.config.get('language'),
                        show_all=False,
                    )
                
                self.progress_update.emit(100)
                if text:
                    # Apply corrections for common Portuguese words
                    text = self._correct_common_portuguese_errors(text)
                    text = fix_stutter(text)
                    text = auto_punctuate(text)
                    self.text_ready.emit(text)
                else:
                    self.text_ready.emit("")
                    self.partial_text.emit("Nenhum texto reconhecido")
            except sr.UnknownValueError:
                self.progress_update.emit(100)
                self.text_ready.emit("")
                self.partial_text.emit("Could not understand the audio. Try again, speaking more slowly and clearly.")
            except sr.RequestError as e:
                self.progress_update.emit(100)
                log.error(f"Recognition service request failed: {e}")
                self.text_ready.emit("")
                self.partial_text.emit("ERROR: recognition service unavailable. Check your connection.")
        except Exception as e:
            self.progress_update.emit(100)
            log.error(f"Failed to process audio: {e}")
            self.text_ready.emit("")
            self.partial_text.emit(f"ERRO: {str(e)}")
    
    def _correct_common_portuguese_errors(self, text):
        """Fix common Portuguese recognition errors.

        The keys and values stay in Portuguese on purpose: this is pt-BR
        language data, not interface text.
        """
        # Common correction pairs
        corrections = {
            'hum': 'um',
            'nao': 'não',
            'nau': 'não',
            'e ': 'é ',
            'eh ': 'é ',
            'voce': 'você',
            'vc': 'você',
            'pra': 'para',
            'pro': 'para o',
            'ta ': 'está ',
            'ta?': 'está?',
            'entao': 'então',
            'entaum': 'então',
            'tambem': 'também',
            'td': 'tudo'
        }
        
        # Apply the corrections
        result = text
        for wrong, correct in corrections.items():
            result = result.replace(wrong, correct)
            
        return result

# Checks whether the focused window accepts text input
class TextInputChecker:
    @staticmethod
    def is_text_input_focused():
        # Get the focused window
        hwnd = win32gui.GetForegroundWindow()
        
        # Get the window's class name
        try:
            class_name = win32gui.GetClassName(hwnd).lower()
        except:
            return False
        
        # Window classes that usually accept text input
        text_input_classes = [
            'edit', 'richedit', 'textbox', 'tbfind', 'scintilla',
            'thundrebirdwindowclass', 'mozillamaintreeclasswindow',
            'chromiumwidget', 'chrome_widget', 'webkit', 'atl:edit',
            'atom', 'vscode', 'notepad', 'wordpad', 'txview'
        ]
        
        # Check the class against the list
        for text_class in text_input_classes:
            if text_class in class_name:
                return True
                
        # Names of known text applications
        app_names = [
            'notepad', 'word', 'wordpad', 'excel', 'powerpnt', 'onenote',
            'outlook', 'write', 'textedit', 'ultraedit', 'sublime_text',
            'atom', 'code', 'chrome', 'firefox', 'iexplore', 'opera',
            'brave', 'discord', 'slack', 'whatsapp', 'teams'
        ]
        
        # Get the window title
        try:
            window_title = win32gui.GetWindowText(hwnd).lower()
            for app in app_names:
                if app in window_title or app in class_name:
                    return True
        except:
            pass
                
        # Inspect the control type for the specific cases
        try:
            # Read information about the control
            control_info = ctypes.create_string_buffer(1024)
            result = ctypes.windll.user32.GetClassInfoA(0, class_name.encode(), control_info)
            if result:
                # Look for edit-control styles
                style = ctypes.c_long.from_buffer(control_info, 16).value
                if style & 0x00800000:  # ES_MULTILINE
                    return True
        except Exception as e:
            log.debug(f"Erro ao verificar controle: {e}")
        
        return False

# Microphone test dialog
class MicTestDialog(QDialog):
    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self.setWindowTitle("Teste de Microfone")
        self.setMinimumSize(400, 300)
        
        # Set up the layout
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # Title
        title_label = QLabel("Teste de Microfone")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 18px;
            font-weight: 600;
            color: #4285f4;
            margin-bottom: 10px;
        """)
        layout.addWidget(title_label)
        
        # Card principal
        main_card = QFrame()
        main_card.setObjectName("mainCard")
        main_card.setStyleSheet("""
            #mainCard {
                background-color: #f5f5f5;
                border: 1px solid #dddddd;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        card_layout = QVBoxLayout(main_card)
        card_layout.setSpacing(15)
        
        # Label de status
        self.status_label = QLabel("Clique em Iniciar para testar seu microfone")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 14px;
            color: #333333;
            padding: 8px;
            background-color: #ffffff;
            border: 1px solid #dddddd;
            border-radius: 4px;
        """)
        card_layout.addWidget(self.status_label)
        
        # Audio level meter
        level_group = QGroupBox("Audio Level")
        level_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #dddddd;
                border-radius: 4px;
                margin-top: 15px;
                padding-top: 15px;
                color: #4285f4;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 5px;
            }
        """)
        level_layout = QVBoxLayout(level_group)
        
        self.level_bar = QProgressBar()
        self.level_bar.setMinimum(0)
        self.level_bar.setMaximum(100)
        self.level_bar.setValue(0)
        self.level_bar.setStyleSheet("""
            QProgressBar {
                border: none;
                border-radius: 3px;
                background-color: #e0e0e0;
                height: 20px;
                text-align: center;
                color: #333333;
            }
            QProgressBar::chunk {
                background-color: #4285f4;
                border-radius: 3px;
            }
        """)
        level_layout.addWidget(self.level_bar)
        
        card_layout.addWidget(level_group)
        
        # Texto reconhecido
        text_group = QGroupBox("Texto Reconhecido")
        text_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #dddddd;
                border-radius: 4px;
                margin-top: 15px;
                padding-top: 15px;
                color: #4285f4;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 5px;
            }
        """)
        text_layout = QVBoxLayout(text_group)
        
        self.text_output = QLabel("No text recognised yet")
        self.text_output.setWordWrap(True)
        self.text_output.setStyleSheet("""
            background-color: #ffffff;
            padding: 10px;
            border-radius: 4px;
            color: #333333;
            border: 1px solid #dddddd;
            border-left: 3px solid #4285f4;
        """)
        self.text_output.setMinimumHeight(80)
        self.text_output.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        text_layout.addWidget(self.text_output)
        
        card_layout.addWidget(text_group)
        
        # Buttons
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        
        self.start_button = QPushButton("Iniciar Teste")
        self.start_button.setStyleSheet("""
            QPushButton {
                background-color: #4285f4;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 15px;
                font-weight: 600;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #3b78e7;
            }
            QPushButton:pressed {
                background-color: #3367d6;
            }
            QPushButton:disabled {
                background-color: #e0e0e0;
                color: #9e9e9e;
            }
        """)
        self.start_button.clicked.connect(self.start_test)
        
        self.stop_button = QPushButton("Parar")
        self.stop_button.setStyleSheet("""
            QPushButton {
                background-color: #db4437;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 15px;
                font-weight: 600;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #c53929;
            }
            QPushButton:pressed {
                background-color: #b31412;
            }
            QPushButton:disabled {
                background-color: #e0e0e0;
                color: #9e9e9e;
            }
        """)
        self.stop_button.clicked.connect(self.stop_test)
        self.stop_button.setEnabled(False)
        
        close_button = QPushButton("Fechar")
        close_button.setStyleSheet("""
            QPushButton {
                background-color: #f5f5f5;
                color: #555555;
                border: 1px solid #dddddd;
                border-radius: 4px;
                padding: 8px 15px;
                font-weight: 600;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
            QPushButton:pressed {
                background-color: #d5d5d5;
            }
        """)
        close_button.clicked.connect(self.close)
        
        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.stop_button)
        button_layout.addWidget(close_button)
        
        card_layout.addLayout(button_layout)
        layout.addWidget(main_card)
        
        # Dica
        tip_label = QLabel("Dica: Fale normalmente para testar a qualidade do reconhecimento de voz.")
        tip_label.setWordWrap(True)
        tip_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        tip_label.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 12px;
            color: #777777;
            margin-top: 10px;
        """)
        layout.addWidget(tip_label)
        
        # Timer that refreshes the audio level
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_level)
        
        # Audio
        self.p = None
        self.stream = None
        self.audio_processor = None
        
        # Aplica estilo
        self.apply_theme()
    
    def apply_theme(self):
        """Apply the theme to the interface."""
        self.setStyleSheet("""
            QDialog {
                background-color: #ffffff;
            }
            QLabel {
                color: #333333;
            }
            QWidget {
                font-family: 'Segoe UI', sans-serif;
            }
        """)
        
        # Ajustar o tamanho da janela
        self.setFixedSize(450, 550)
    
    def start_test(self):
        self.status_label.setText("Gravando... Fale algo")
        self.status_label.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 14px;
            color: #db4437;
            padding: 8px;
            background-color: #ffffff;
            border: 1px solid #dddddd;
            border-radius: 4px;
            font-weight: 600;
        """)
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        
        # Start audio processing
        self.audio_processor = AudioProcessor(self.config)
        self.audio_processor.text_ready.connect(self.update_text)
        self.audio_processor.progress_update.connect(self.level_bar.setValue)
        self.audio_processor.start()
        
        # Start the audio level timer
        self.timer.start(100)
    
    def stop_test(self):
        self.status_label.setText("Processing audio...")
        self.status_label.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 14px;
            color: #f5a623;
            padding: 8px;
            background-color: #ffffff;
            border: 1px solid #dddddd;
            border-radius: 4px;
            font-weight: 600;
        """)
        self.stop_button.setEnabled(False)
        
        if self.audio_processor and self.audio_processor.isRunning():
            self.audio_processor.stop()
            self.timer.stop()
    
    def update_level(self):
        if self.audio_processor and self.audio_processor.isRunning():
            # Fake an audio level, purely for visual feedback
            import random
            self.level_bar.setValue(random.randint(30, 80))
    
    def update_text(self, text):
        if text:
            self.text_output.setText(text)
        else:
            self.text_output.setText("No text could be recognised")
        
        self.status_label.setText("Test complete")
        self.status_label.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 14px;
            color: #0f9d58;
            padding: 8px;
            background-color: #ffffff;
            border: 1px solid #dddddd;
            border-radius: 4px;
            font-weight: 600;
        """)
        self.start_button.setEnabled(True)
    
    def closeEvent(self, event):
        # Make sure audio processing stops on close
        if self.audio_processor and self.audio_processor.isRunning():
            self.audio_processor.stop()
            self.timer.stop()
        event.accept()

# Settings window
class SettingsDialog(QDialog):
    config_changed = pyqtSignal()
    
    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self.original_hotkey = config.get('hotkey')
        self.setWindowTitle("Settings")
        self.setMinimumSize(500, 400)
        
        # Aplicar estilo
        self.apply_theme()
        
        # Layout principal
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # Window title
        title_label = QLabel("Dictation Assistant Settings")
        title_label.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 18px;
            font-weight: 600;
            color: #4285f4;
            margin-bottom: 10px;
        """)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)
        
        # Criar abas
        tab_widget = QTabWidget()
        tab_widget.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #dddddd;
                border-radius: 4px;
                background-color: #ffffff;
            }
            QTabBar::tab {
                background-color: #f5f5f5;
                color: #555555;
                padding: 8px 15px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                border: 1px solid #dddddd;
                border-bottom: none;
            }
            QTabBar::tab:selected {
                background-color: #ffffff;
                color: #4285f4;
                border-bottom: 2px solid #4285f4;
            }
            QTabBar::tab:hover:!selected {
                background-color: #e0e0e0;
            }
        """)
        
        # Aba Geral
        general_tab = QWidget()
        general_layout = QVBoxLayout(general_tab)
        general_layout.setSpacing(15)
        
        # Autostart group
        autostart_group = QGroupBox("Startup")
        autostart_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #dddddd;
                border-radius: 4px;
                margin-top: 15px;
                padding-top: 15px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 5px;
                color: #4285f4;
            }
        """)
        autostart_layout = QVBoxLayout(autostart_group)
        
        self.autostart_check = QCheckBox("Iniciar com o Windows")
        self.autostart_check.setStyleSheet("""
            QCheckBox {
                color: #333333;
                spacing: 5px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
            QCheckBox::indicator:unchecked {
                border: 1px solid #dddddd;
                border-radius: 3px;
                background-color: #ffffff;
            }
            QCheckBox::indicator:checked {
                border: 1px solid #4285f4;
                border-radius: 3px;
                background-color: #4285f4;
            }
        """)
        self.autostart_check.setChecked(self.config.get('auto_start'))
        autostart_layout.addWidget(self.autostart_check)
        
        theme_layout = QHBoxLayout()
        theme_label = QLabel("Tema:")
        theme_label.setStyleSheet("color: #333333;")
        theme_layout.addWidget(theme_label)
        
        self.theme_combo = QComboBox()
        self.theme_combo.setStyleSheet("""
            QComboBox {
                border: 1px solid #dddddd;
                border-radius: 3px;
                padding: 5px;
                background-color: #ffffff;
                color: #333333;
                min-width: 150px;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left: 1px solid #dddddd;
            }
        """)
        self.theme_combo.addItem("Tema Claro", "light")
        self.theme_combo.addItem("Tema Escuro", "dark")
        theme_layout.addWidget(self.theme_combo)
        autostart_layout.addLayout(theme_layout)
        
        idx = self.theme_combo.findData(self.config.get('theme'))
        if idx >= 0:
            self.theme_combo.setCurrentIndex(idx)
        
        general_layout.addWidget(autostart_group)
        
        # Grupo de Atalho
        hotkey_group = QGroupBox("Atalho de Teclado")
        hotkey_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #dddddd;
                border-radius: 4px;
                margin-top: 15px;
                padding-top: 15px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 5px;
                color: #4285f4;
            }
        """)
        hotkey_layout = QVBoxLayout(hotkey_group)
        
        hotkey_label = QLabel("Shortcut to start/stop recording:")
        hotkey_label.setStyleSheet("color: #333333;")
        hotkey_layout.addWidget(hotkey_label)
        
        self.hotkey_edit = QLineEdit(self.config.get('hotkey'))
        self.hotkey_edit.setPlaceholderText("Press a key combination")
        self.hotkey_edit.setStyleSheet("""
            QLineEdit {
                border: 1px solid #dddddd;
                border-radius: 3px;
                padding: 8px;
                background-color: #ffffff;
                color: #333333;
            }
            QLineEdit:focus {
                border: 1px solid #4285f4;
            }
        """)
        hotkey_layout.addWidget(self.hotkey_edit)
        
        hotkey_note = QLabel("Exemplo: ctrl+alt+d, shift+f12, mouse5, etc.")
        hotkey_note.setStyleSheet("color: #777777; font-size: 11px;")
        hotkey_layout.addWidget(hotkey_note)
        
        general_layout.addWidget(hotkey_group)
        
        # Grupo de Comportamento
        behavior_group = QGroupBox("Comportamento")
        behavior_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #dddddd;
                border-radius: 4px;
                margin-top: 15px;
                padding-top: 15px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 5px;
                color: #4285f4;
            }
        """)
        behavior_layout = QVBoxLayout(behavior_group)
        
        self.continuous_check = QCheckBox("Continuous recognition (experimental)")
        self.continuous_check.setStyleSheet("""
            QCheckBox {
                color: #333333;
                spacing: 5px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
            QCheckBox::indicator:unchecked {
                border: 1px solid #dddddd;
                border-radius: 3px;
                background-color: #ffffff;
            }
            QCheckBox::indicator:checked {
                border: 1px solid #4285f4;
                border-radius: 3px;
                background-color: #4285f4;
            }
        """)
        self.continuous_check.setChecked(self.config.get('continuous_recognition'))
        behavior_layout.addWidget(self.continuous_check)
        
        self.realtime_check = QCheckBox("Show text in real time while recording")
        self.realtime_check.setStyleSheet("""
            QCheckBox {
                color: #333333;
                spacing: 5px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
            QCheckBox::indicator:unchecked {
                border: 1px solid #dddddd;
                border-radius: 3px;
                background-color: #ffffff;
            }
            QCheckBox::indicator:checked {
                border: 1px solid #4285f4;
                border-radius: 3px;
                background-color: #4285f4;
            }
        """)
        self.realtime_check.setChecked(self.config.get('show_realtime_text'))
        behavior_layout.addWidget(self.realtime_check)
        
        general_layout.addWidget(behavior_group)
        
        # Audio tab
        audio_tab = QWidget()
        audio_layout = QVBoxLayout(audio_tab)
        audio_layout.setSpacing(15)
        
        # Audio quality group
        quality_group = QGroupBox("Audio Quality")
        quality_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #dddddd;
                border-radius: 4px;
                margin-top: 15px;
                padding-top: 15px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 5px;
                color: #4285f4;
            }
        """)
        quality_layout = QVBoxLayout(quality_group)
        
        quality_label = QLabel("Qualidade de captura:")
        quality_label.setStyleSheet("color: #333333;")
        quality_layout.addWidget(quality_label)
        
        self.quality_combo = QComboBox()
        self.quality_combo.setStyleSheet("""
            QComboBox {
                border: 1px solid #dddddd;
                border-radius: 3px;
                padding: 5px;
                background-color: #ffffff;
                color: #333333;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left: 1px solid #dddddd;
            }
        """)
        self.quality_combo.addItem("Alta (melhor qualidade)", "high")
        self.quality_combo.addItem("Medium (balanced)", "medium")
        self.quality_combo.addItem("Low (faster)", "low")
        
        quality_idx = 0
        for i in range(self.quality_combo.count()):
            if self.quality_combo.itemData(i) == self.config.get('audio_quality'):
                quality_idx = i
                break
        self.quality_combo.setCurrentIndex(quality_idx)
        
        quality_layout.addWidget(self.quality_combo)
        
        # Microphone test button
        mic_test_button = QPushButton("Testar Microfone")
        mic_test_button.setStyleSheet("""
            QPushButton {
                background-color: #4285f4;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #3b78e7;
            }
            QPushButton:pressed {
                background-color: #3367d6;
            }
        """)
        mic_test_button.clicked.connect(self.open_mic_test)
        quality_layout.addWidget(mic_test_button)
        
        audio_layout.addWidget(quality_group)
        
        # Grupo de Idioma
        language_group = QGroupBox("Idioma")
        language_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #dddddd;
                border-radius: 4px;
                margin-top: 15px;
                padding-top: 15px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 5px;
                color: #4285f4;
            }
        """)
        language_layout = QVBoxLayout(language_group)
        
        language_label = QLabel("Idioma para reconhecimento:")
        language_label.setStyleSheet("color: #333333;")
        language_layout.addWidget(language_label)
        
        self.language_combo = QComboBox()
        self.language_combo.setStyleSheet("""
            QComboBox {
                border: 1px solid #dddddd;
                border-radius: 3px;
                padding: 5px;
                background-color: #ffffff;
                color: #333333;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left: 1px solid #dddddd;
            }
        """)
        self.language_combo.addItem("Portuguese (Brazil)", "pt-BR")
        self.language_combo.addItem("English (US)", "en-US")
        self.language_combo.addItem("Spanish", "es-ES")
        self.language_combo.addItem("French", "fr-FR")
        self.language_combo.addItem("Italian", "it-IT")
        self.language_combo.addItem("German", "de-DE")
        
        idx = 0
        for i in range(self.language_combo.count()):
            if self.language_combo.itemData(i) == self.config.get('language'):
                idx = i
                break
        self.language_combo.setCurrentIndex(idx)
        
        language_layout.addWidget(self.language_combo)
        
        audio_layout.addWidget(language_group)
        
        # Grupo de Sensibilidade
        sensitivity_group = QGroupBox("Sensibilidade")
        sensitivity_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #dddddd;
                border-radius: 4px;
                margin-top: 15px;
                padding-top: 15px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 5px;
                color: #4285f4;
            }
        """)
        sensitivity_layout = QVBoxLayout(sensitivity_group)
        
        sensitivity_label = QLabel("Sensibilidade do microfone:")
        sensitivity_label.setStyleSheet("color: #333333;")
        sensitivity_layout.addWidget(sensitivity_label)
        
        self.sensitivity_slider = QSlider(Qt.Orientation.Horizontal)
        self.sensitivity_slider.setMinimum(0)
        self.sensitivity_slider.setMaximum(100)
        self.sensitivity_slider.setValue(self.config.get('sensitivity'))
        self.sensitivity_slider.setTickPosition(QSlider.TickPosition.TicksBelow)
        self.sensitivity_slider.setTickInterval(10)
        self.sensitivity_slider.setStyleSheet("""
            QSlider::groove:horizontal {
                border: 1px solid #dddddd;
                height: 8px;
                background: #f5f5f5;
                margin: 2px 0;
                border-radius: 4px;
            }
            QSlider::handle:horizontal {
                background: #4285f4;
                border: 1px solid #4285f4;
                width: 18px;
                height: 18px;
                margin: -5px 0;
                border-radius: 9px;
            }
        """)
        
        sensitivity_layout.addWidget(self.sensitivity_slider)
        
        self.sensitivity_label = QLabel(f"Valor: {self.config.get('sensitivity')}%")
        self.sensitivity_label.setStyleSheet("color: #333333;")
        sensitivity_layout.addWidget(self.sensitivity_label)
        
        self.sensitivity_slider.valueChanged.connect(self.update_sensitivity_label)
        
        audio_layout.addWidget(sensitivity_group)
        
        # Adicionar todas as abas
        tab_widget.addTab(general_tab, "Geral")
        tab_widget.addTab(audio_tab, "Audio")
        
        # OK and Cancel buttons
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        
        save_button = QPushButton("Salvar")
        save_button.setStyleSheet("""
            QPushButton {
                background-color: #0f9d58;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 15px;
                font-weight: 600;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #0b8043;
            }
            QPushButton:pressed {
                background-color: #0a753a;
            }
        """)
        save_button.clicked.connect(self.save_settings)
        
        cancel_button = QPushButton("Cancelar")
        cancel_button.setStyleSheet("""
            QPushButton {
                background-color: #db4437;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 15px;
                font-weight: 600;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #c53929;
            }
            QPushButton:pressed {
                background-color: #b31412;
            }
        """)
        cancel_button.clicked.connect(self.reject)
        
        button_layout.addStretch()
        button_layout.addWidget(save_button)
        button_layout.addWidget(cancel_button)
        
        # Montar layout final
        layout.addWidget(tab_widget)
        layout.addLayout(button_layout)
    
    def open_mic_test(self):
        mic_test = MicTestDialog(self.config, self)
        mic_test.exec()
    
    def apply_theme(self):
        """Apply the theme to the interface."""
        self.setStyleSheet("""
            QDialog {
                background-color: #ffffff;
            }
            QLabel {
                color: #333333;
            }
            QWidget {
                font-family: 'Segoe UI', sans-serif;
            }
        """)
        
        # Window settings
        self.setFixedSize(500, 600)
    
    def update_sensitivity_label(self, value):
        self.sensitivity_label.setText(f"Valor: {value}%")
    
    def save_settings(self):
        # Save settings
        self.config.set('auto_start', self.autostart_check.isChecked())
        self.config.set('theme', self.theme_combo.currentData())
        self.config.set('hotkey', self.hotkey_edit.text())
        self.config.set('language', self.language_combo.currentData())
        self.config.set('sensitivity', self.sensitivity_slider.value())
        self.config.set('continuous_recognition', self.continuous_check.isChecked())
        self.config.set('show_realtime_text', self.realtime_check.isChecked())
        self.config.set('audio_quality', self.quality_combo.currentData())
        
        # Configure autostart
        setup_autostart(self.autostart_check.isChecked())
        
        # Signal that the settings changed
        self.config_changed.emit()
        
        # Close the dialog
        self.accept()

# Main application window
class MainWindow(QMainWindow):
    def __init__(self, config):
        super().__init__()
        self.config = config
        self.audio_processor = AudioProcessor(config)
        self.pending_text = ""
        self.original_hotkey = config.get('hotkey')
        self.recording_active = False
        self.text_collected = False
        self.text_rejected = False  # Novo campo para controlar se o texto foi rejeitado
        self.mouse_listener = None  # Para armazenar o listener do mouse
        
        self.setWindowTitle("Assistente de Ditado")
        self.setMinimumSize(500, 400)
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint)
        
        # Set the application icon
        self.setWindowIcon(QIcon("mic_icon.png"))
        
        # Configurar bandeja do sistema
        self.setup_tray()
        
        # Configurar interface
        self.setup_ui()
        
        # Conectar sinais
        self.connect_signals()
        
        # Registrar atalho global
        self.register_hotkey()
        
        # Periodic check on the pending-text state
        self.pending_text_timer = QTimer(self)
        self.pending_text_timer.timeout.connect(self.check_pending_text)
        self.pending_text_timer.start(1000)  # Verifica a cada segundo
    
    def setup_tray(self):
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon("mic_icon.png"))
        
        # Menu da bandeja
        tray_menu = QMenu()
        
        # Menu actions
        show_action = QAction("Mostrar", self)
        show_action.triggered.connect(self.show)
        
        settings_action = QAction("Settings", self)
        settings_action.triggered.connect(self.show_settings)
        
        quit_action = QAction("Sair", self)
        quit_action.triggered.connect(self.quit_app)
        
        # Add the actions to the menu
        tray_menu.addAction(show_action)
        tray_menu.addAction(settings_action)
        tray_menu.addSeparator()
        tray_menu.addAction(quit_action)
        
        # Attach the menu to the tray icon
        self.tray_icon.setContextMenu(tray_menu)
        
        # Show the tray icon
        self.tray_icon.show()
        
        # Wire up clicks on the tray icon
        self.tray_icon.activated.connect(self.tray_icon_activated)
    
    def setup_ui(self):
        # Main widget with a clean layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout with sensible margins
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        # Simple header
        header = QLabel("Assistente de Ditado")
        header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 22px;
            font-weight: 600;
            color: #333333;
        """)
        main_layout.addWidget(header)
        
        # === STATUS AND CONTROL SECTION ===
        status_card = QFrame()
        status_card.setObjectName("statusCard")
        status_card.setStyleSheet("""
            #statusCard {
                background-color: #f5f5f5;
                border: 1px solid #dddddd;
                border-radius: 8px;
            }
        """)
        status_layout = QVBoxLayout(status_card)
        status_layout.setSpacing(10)
        
        # Current status, with an icon
        status_container = QWidget()
        status_container_layout = QHBoxLayout(status_container)
        status_container_layout.setContentsMargins(10, 10, 10, 10)
        
        # Microphone icon
        self.mic_icon = QLabel()
        self.mic_icon.setFixedSize(32, 32)
        
        # Use the icon file when present, otherwise draw a basic one
        if os.path.exists("mic_icon.png"):
            self.mic_icon.setPixmap(QPixmap("mic_icon.png").scaled(32, 32, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        else:
            # Draw a basic microphone icon
            pixmap = QPixmap(32, 32)
            pixmap.fill(Qt.GlobalColor.transparent)
            painter = QPainter(pixmap)
            painter.setPen(QPen(QColor("#555555"), 2))
            painter.setBrush(QColor("#555555"))
            painter.drawEllipse(8, 8, 16, 16)
            painter.drawRect(14, 18, 4, 10)
            painter.end()
            self.mic_icon.setPixmap(pixmap)
        
        status_container_layout.addWidget(self.mic_icon, 0, Qt.AlignmentFlag.AlignCenter)
        
        # Status atual - Texto claro
        self.status_label = QLabel("Pronto para ditar")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 16px;
            font-weight: 600;
            color: #333333;
        """)
        status_container_layout.addWidget(self.status_label)
        status_layout.addWidget(status_container)
        
        # Main record button
        self.record_button = QPushButton("Iniciar")
        self.record_button.setFixedHeight(40)
        
        # Build the button icon when it is missing
        if not os.path.exists("mic_icon.png"):
            pixmap = QPixmap(24, 24)
            pixmap.fill(Qt.GlobalColor.transparent)
            painter = QPainter(pixmap)
            painter.setPen(QPen(QColor("#FFFFFF"), 2))
            painter.setBrush(QColor("#FFFFFF"))
            painter.drawEllipse(6, 6, 12, 12)
            painter.drawRect(10, 14, 4, 8)
            painter.end()
            self.record_button.setIcon(QIcon(pixmap))
        else:
            self.record_button.setIcon(QIcon("mic_icon.png"))
        
        self.record_button.setIconSize(QSize(20, 20))
        self.record_button.setStyleSheet("""
            QPushButton {
                background-color: #4285f4;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 5px 15px;
                font-family: 'Segoe UI', sans-serif;
                font-size: 14px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #3b78e7;
            }
            QPushButton:pressed {
                background-color: #3367d6;
            }
        """)
        self.record_button.clicked.connect(self.toggle_recording)
        status_layout.addWidget(self.record_button)
        
        # Barra de progresso
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(6)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: none;
                border-radius: 3px;
                background-color: #e0e0e0;
            }
            QProgressBar::chunk {
                background-color: #4285f4;
                border-radius: 3px;
            }
        """)
        status_layout.addWidget(self.progress_bar)
        
        main_layout.addWidget(status_card)
        
        # === REAL-TIME TEXT SECTION ===
        realtime_card = QFrame()
        realtime_card.setObjectName("realtimeCard")
        realtime_card.setStyleSheet("""
            #realtimeCard {
                background-color: #f5f5f5;
                border: 1px solid #dddddd;
                border-radius: 8px;
                margin-top: 10px;
            }
        """)
        realtime_layout = QVBoxLayout(realtime_card)
        realtime_layout.setContentsMargins(15, 15, 15, 15)
        
        # Label for the real-time text
        realtime_header = QLabel("Texto em tempo real:")
        realtime_header.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 14px;
            font-weight: 600;
            color: #555555;
        """)
        realtime_layout.addWidget(realtime_header)
        
        # Texto em tempo real
        self.realtime_label = QLabel("Aguardando sua voz...")
        self.realtime_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.realtime_label.setWordWrap(True)
        self.realtime_label.setMinimumHeight(40)
        self.realtime_label.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 14px;
            font-style: italic;
            color: #777777;
            padding: 10px;
            background-color: #ffffff;
            border: 1px solid #e0e0e0;
            border-radius: 4px;
        """)
        realtime_layout.addWidget(self.realtime_label)
        main_layout.addWidget(realtime_card)
        
        # === RESULT SECTION ===
        result_card = QFrame()
        result_card.setObjectName("resultCard")
        result_card.setStyleSheet("""
            #resultCard {
                background-color: #f5f5f5;
                border: 1px solid #dddddd;
                border-radius: 8px;
                margin-top: 10px;
            }
        """)
        result_layout = QVBoxLayout(result_card)
        result_layout.setContentsMargins(15, 15, 15, 15)
        
        # Result header
        result_header = QLabel("TEXTO RECONHECIDO")
        result_header.setAlignment(Qt.AlignmentFlag.AlignLeft)
        result_header.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 14px;
            font-weight: 600;
            color: #555555;
        """)
        result_layout.addWidget(result_header)
        
        # Texto reconhecido
        self.text_output = QLabel("")
        self.text_output.setWordWrap(True)
        self.text_output.setMinimumHeight(120)
        self.text_output.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        self.text_output.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 14px;
            color: #333333;
            padding: 10px;
            background-color: #ffffff;
            border: 1px solid #e0e0e0;
            border-radius: 4px;
            border-left: 3px solid #4285f4;
        """)
        result_layout.addWidget(self.text_output)
        
        # Action area for the recognised text
        actions_layout = QHBoxLayout()
        actions_layout.setSpacing(10)
        
        # Accept button
        accept_button = QPushButton("ACEITAR (MOUSE5)")
        accept_button.setStyleSheet("""
            QPushButton {
                background-color: #0f9d58;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px;
                font-family: 'Segoe UI', sans-serif;
                font-size: 12px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #0b8043;
            }
        """)
        accept_button.clicked.connect(self.paste_collected_text)
        actions_layout.addWidget(accept_button)
        
        # Reject button
        reject_button = QPushButton("REJEITAR (MOUSE6)")
        reject_button.setStyleSheet("""
            QPushButton {
                background-color: #db4437;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px;
                font-family: 'Segoe UI', sans-serif;
                font-size: 12px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #c53929;
            }
        """)
        reject_button.clicked.connect(self.reject_text)
        actions_layout.addWidget(reject_button)
        
        result_layout.addLayout(actions_layout)
        main_layout.addWidget(result_card)
        
        # === QUICK SETTINGS SECTION ===
        settings_card = QFrame()
        settings_card.setObjectName("settingsCard")
        settings_card.setStyleSheet("""
            #settingsCard {
                background-color: #f5f5f5;
                border: 1px solid #dddddd;
                border-radius: 8px;
                margin-top: 10px;
            }
        """)
        settings_layout = QHBoxLayout(settings_card)
        settings_layout.setContentsMargins(15, 10, 15, 10)
        
        # Settings button
        settings_button = QPushButton("Settings")
        
        # Build the button icon
        pixmap = QPixmap(16, 16)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        painter.setPen(QPen(QColor("#555555"), 1))
        painter.drawEllipse(4, 4, 8, 8)
        painter.drawLine(8, 8, 8, 12)
        painter.end()
        settings_button.setIcon(QIcon(pixmap))
        
        settings_button.setStyleSheet("""
            QPushButton {
                background-color: #f5f5f5;
                color: #555555;
                border: 1px solid #dddddd;
                border-radius: 4px;
                padding: 8px;
                font-family: 'Segoe UI', sans-serif;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
        """)
        settings_button.clicked.connect(self.show_settings)
        settings_layout.addWidget(settings_button)
        
        # Microphone test button
        mic_test_button = QPushButton("Testar Microfone")
        
        # Build the button icon
        pixmap = QPixmap(16, 16)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        painter.setPen(QPen(QColor("#555555"), 1))
        painter.drawEllipse(4, 4, 8, 8)
        painter.drawRect(7, 8, 2, 6)
        painter.end()
        mic_test_button.setIcon(QIcon(pixmap))
        
        mic_test_button.setStyleSheet("""
            QPushButton {
                background-color: #f5f5f5;
                color: #555555;
                border: 1px solid #dddddd;
                border-radius: 4px;
                padding: 8px;
                font-family: 'Segoe UI', sans-serif;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
        """)
        mic_test_button.clicked.connect(lambda: MicTestDialog(self.config, self).exec())
        settings_layout.addWidget(mic_test_button)
        
        main_layout.addWidget(settings_card)
        
        # Footer hint
        footer = QLabel("Tip: use the side mouse buttons (MOUSE5 and MOUSE6) to control the app.")
        footer.setWordWrap(True)
        footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        footer.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 12px;
            color: #777777;
            margin-top: 10px;
        """)
        main_layout.addWidget(footer)
        
        # Aplicar tema
        self.apply_theme()
    
    def connect_signals(self):
        # Connect the audio processor's signals
        self.audio_processor.text_ready.connect(self.update_text)
        self.audio_processor.partial_text.connect(self.update_realtime_text)
        self.audio_processor.progress_update.connect(self.progress_bar.setValue)
    
    def register_hotkey(self):
        try:
            # Remover atalho anterior se existir
            if self.mouse_listener:
                self.mouse_listener.stop()
                self.mouse_listener = None
            
            self.original_hotkey = self.config.get('hotkey')
            
            # When it is a mouse button
            if 'mouse' in self.original_hotkey.lower():
                # Set up the mouse listener
                self.mouse_listener = mouse.Listener(on_click=self.on_mouse_click)
                self.mouse_listener.start()
                log.info(f"Registrado listener do mouse para: {self.original_hotkey}")
            else:
                # For a keyboard key, use the keyboard library
                try:
                    keyboard.remove_hotkey(self.original_hotkey)
                except:
                    pass
                keyboard.add_hotkey(self.original_hotkey, self.toggle_recording)
            
            # Update the button text
            self.record_button.setText("INICIAR")
        except Exception as e:
            log.error(f"Erro ao registrar atalho: {e}")
            QMessageBox.warning(self, "Error", f"Could not register the shortcut: {e}")
    
    def on_mouse_click(self, x, y, button, pressed):
        """Callback for mouse click events."""
        # Stringify the button name and drop the "Button." prefix
        button_name = str(button).replace('Button.', '')
        target_button = self.original_hotkey.lower().replace('mouse', '')
        
        # Check for the mouse5 button (x2)
        if button_name == 'x2' and self.original_hotkey.lower() == 'mouse5' and pressed:
            # Toggle recording when it is the right button
            QTimer.singleShot(0, self.toggle_recording)
            
        # Check for the mouse6 button, which rejects the text
        elif button_name == 'x1' and pressed and self.text_collected:
            # Rejeita o texto reconhecido
            QTimer.singleShot(0, self.reject_text)
    
    def toggle_recording(self):
        """Start or stop audio recording."""
        
        if self.recording_active:
            # Stop recording
            self.audio_processor.stop()
            self.status_label.setText("Processando...")
            self.status_label.setStyleSheet("""
                font-family: 'Segoe UI', sans-serif;
                font-size: 16px;
                font-weight: 600;
                color: #f5a623;
            """)
            self.record_button.setText("Iniciar")
            self.record_button.setStyleSheet("""
                QPushButton {
                    background-color: #4285f4;
                    color: white;
                    border: none;
                    border-radius: 4px;
                    padding: 5px 15px;
                    font-family: 'Segoe UI', sans-serif;
                    font-size: 14px;
                    font-weight: 600;
                }
                QPushButton:hover {
                    background-color: #3b78e7;
                }
                QPushButton:pressed {
                    background-color: #3367d6;
                }
            """)
            self.recording_active = False
            
            # Stop the microphone animation
            if hasattr(self, 'mic_animation') and self.mic_animation is not None:
                self.mic_animation.stop()
            
        elif self.text_collected:
            # Efeito visual de sucesso
            self.paste_collected_text()
            self.text_collected = False
            self.status_label.setText("Texto colado!")
            self.status_label.setStyleSheet("""
                font-family: 'Segoe UI', sans-serif;
                font-size: 16px;
                font-weight: 600;
                color: #0f9d58;
            """)
            
        else:
            # Start recording
            self.status_label.setText("Gravando...")
            self.status_label.setStyleSheet("""
                font-family: 'Segoe UI', sans-serif;
                font-size: 16px;
                font-weight: 600;
                color: #db4437;
            """)
            self.realtime_label.setText("Ouvindo sua voz...")
            self.record_button.setText("Parar")
            self.record_button.setStyleSheet("""
                QPushButton {
                    background-color: #db4437;
                    color: white;
                    border: none;
                    border-radius: 4px;
                    padding: 5px 15px;
                    font-family: 'Segoe UI', sans-serif;
                    font-size: 14px;
                    font-weight: 600;
                }
                QPushButton:hover {
                    background-color: #c53929;
                }
                QPushButton:pressed {
                    background-color: #b31412;
                }
            """)
            self.recording_active = True
            self.text_collected = False
            self.text_rejected = False
            self.progress_bar.setValue(0)
            self.text_output.setText("")
            
            # Start the pulsing microphone animation
            self.mic_animation = QPropertyAnimation(self.mic_icon, b"geometry")
            self.mic_animation.setDuration(1000)
            self.mic_animation.setLoopCount(-1)  # Loop infinito
            mic_geometry = self.mic_icon.geometry()
            self.mic_animation.setStartValue(mic_geometry)
            expanded_geometry = QRect(mic_geometry.x()-2, mic_geometry.y()-2, 
                                     mic_geometry.width()+4, mic_geometry.height()+4)
            self.mic_animation.setEndValue(expanded_geometry)
            self.mic_animation.setEasingCurve(QEasingCurve.Type.InOutQuad)
            self.mic_animation.start()
            
            self.audio_processor.start()
    
    def update_text(self, text):
        if not text:
            self.status_label.setText("Nenhum texto reconhecido.")
            return
            
        self.text_output.setText(text)
        self.status_label.setText("Texto reconhecido! Mouse5: colar texto | Mouse6: rejeitar texto")
        self.realtime_label.setText("")
        self.pending_text = text
        self.text_collected = True
        self.text_rejected = False
        
        # Show a notification
        self.tray_icon.showMessage(
            "Assistente de Ditado",
            "Texto reconhecido. Mouse5: colar texto | Mouse6: rejeitar texto",
            QSystemTrayIcon.MessageIcon.Information,
            3000
        )
    
    def reject_text(self):
        """Reject the recognised text."""
        if self.text_collected:
            self.text_rejected = True
            self.text_collected = False
            self.pending_text = ""
            self.text_output.setText("")
            self.status_label.setText("TEXTO REJEITADO! REINICIANDO...")
            
            # Som de feedback (beep) - som de erro
            try:
                ctypes.windll.user32.MessageBeep(0x10)  # MB_ICONHAND
            except:
                pass
            
            # Show a notification
            self.tray_icon.showMessage(
                "Assistente de Ditado",
                "TEXTO REJEITADO! Reiniciando para nova tentativa...",
                QSystemTrayIcon.MessageIcon.Warning,
                2000
            )
            
            # Restart recording automatically after a short delay
            if not self.recording_active:
                QTimer.singleShot(1000, self.toggle_recording)
    
    def paste_collected_text(self):
        # Check whether a text area has focus
        if TextInputChecker.is_text_input_focused():
            # Go through the clipboard: copy, then paste
            clipboard_backup = None
            try:
                # Back up the current clipboard
                win32clipboard.OpenClipboard()
                if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                    clipboard_backup = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                win32clipboard.CloseClipboard()
                
                # Copy the recognised text to the clipboard
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardText(self.pending_text, win32clipboard.CF_UNICODETEXT)
                win32clipboard.CloseClipboard()
                
                # Send Ctrl+V to paste
                pyautogui.hotkey('ctrl', 'v')
                
                # Restore the original clipboard after a short delay
                if clipboard_backup is not None:
                    time.sleep(0.1)  # Pequena pausa para garantir que o texto foi colado
                    win32clipboard.OpenClipboard()
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardText(clipboard_backup, win32clipboard.CF_UNICODETEXT)
                    win32clipboard.CloseClipboard()
                
                # Limpar texto pendente
                self.pending_text = ""
            except Exception as e:
                log.error(f"Clipboard handling failed: {e}")
        else:
            # Report that no text area has focus
            self.status_label.setText("No editable area selected")
            self.tray_icon.showMessage(
                "Assistente de Ditado",
                "No editable area is selected. Click a text field and try again.",
                QSystemTrayIcon.MessageIcon.Warning,
                3000
            )
    
    def update_realtime_text(self, text):
        self.realtime_label.setText(text)
    
    def apply_theme(self):
        """Apply the theme to the interface."""
        
        # Estilo global
        style = """
            QMainWindow {
                background-color: #ffffff;
            }
            QDialog {
                background-color: #ffffff;
            }
            QLabel {
                color: #333333;
            }
            QWidget {
                font-family: 'Segoe UI', sans-serif;
            }
            QScrollBar:vertical {
                border: none;
                background: #f0f0f0;
                width: 8px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical {
                background: #c0c0c0;
                min-height: 20px;
                border-radius: 4px;
            }
            QScrollBar::add-line, QScrollBar::sub-line {
                height: 0px;
            }
            QToolTip {
                background-color: #f5f5f5;
                color: #333333;
                border: 1px solid #dddddd;
                padding: 5px;
                border-radius: 4px;
            }
        """
        self.setStyleSheet(style)
        
        # Window settings
        self.setFixedSize(500, 650)
    
    def show(self):
        self.showNormal()
    
    def show_settings(self):
        settings_dialog = SettingsDialog(self.config, self)
        settings_dialog.config_changed.connect(self.apply_settings_changes)
        settings_dialog.exec()
    
    def apply_settings_changes(self):
        self.audio_processor.config.set('auto_start', self.config.get('auto_start'))
        self.audio_processor.config.set('theme', self.config.get('theme'))
        self.audio_processor.config.set('hotkey', self.config.get('hotkey'))
        self.audio_processor.config.set('language', self.config.get('language'))
        self.audio_processor.config.set('sensitivity', self.config.get('sensitivity'))
        self.audio_processor.config.set('continuous_recognition', self.config.get('continuous_recognition'))
        self.audio_processor.config.set('show_realtime_text', self.config.get('show_realtime_text'))
        self.audio_processor.config.set('audio_quality', self.config.get('audio_quality'))
        self.audio_processor.config.save()
    
    def quit_app(self):
        self.close()
    
    def tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            self.show()
    
    def check_pending_text(self):
        # Kept empty so existing references do not break
        pass
        
    def closeEvent(self, event):
        # Para o listener do mouse antes de fechar
        if self.mouse_listener:
            self.mouse_listener.stop()
        # Make sure any running recording is stopped
        if self.audio_processor.isRunning():
            self.audio_processor.stop()
            self.audio_processor.wait()
        event.accept()

# Draws a basic microphone icon
def create_mic_icon():
    try:
        # Draw a simple microphone image
        from PIL import Image, ImageDraw
        
        img = Image.new('RGBA', (128, 128), color=(0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        
        # Desenhar um microfone simples
        draw.rectangle([48, 32, 80, 80], fill=(0, 170, 255))
        draw.ellipse([44, 24, 84, 40], fill=(0, 170, 255))
        draw.rectangle([60, 80, 68, 96], fill=(0, 170, 255))
        draw.ellipse([40, 96, 88, 112], fill=(0, 170, 255))
        
        img.save("mic_icon.png")
    except Exception as e:
        log.error(f"Failed to create the icon: {e}")

# Configures autostart with Windows
def setup_autostart(enable=True):
    import winreg
    
    app_name = "AssistenteDitado"
    app_path = os.path.abspath(sys.argv[0])
    
    try:
        # Abrir registro do Windows
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0,
            winreg.KEY_SET_VALUE | winreg.KEY_QUERY_VALUE
        )
        
        if enable:
            # Adicionar ao iniciar
            winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, f'"{app_path}"')
        else:
            # Remove from startup
            try:
                winreg.DeleteValue(key, app_name)
            except FileNotFoundError:
                pass
        
        winreg.CloseKey(key)
        return True
    except Exception as e:
        log.error(f"Failed to configure autostart: {e}")
        return False

if __name__ == "__main__":
    _mutex = ensure_single_instance()
    log.info("Live Dictate starting...")
    app = QApplication(sys.argv)
    config = Config()
    window = MainWindow(config)
    window.show()
    sys.exit(app.exec()) 