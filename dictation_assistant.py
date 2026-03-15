import sys
import os
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
from pynput import mouse  # Adiciona suporte para eventos do mouse

# Configuração da aplicação
CONFIG_FILE = os.path.join(os.path.expanduser('~'), 'dictation_assistant_config.json')
DEFAULT_CONFIG = {
    'hotkey': 'mouse5',
    'language': 'pt-BR',
    'sample_rate': 48000,  # Aumentado para melhor qualidade
    'chunk_size': 1024,
    'auto_start': True,
    'sensitivity': 70,  # Sensibilidade aumentada para melhor captar fala rápida
    'theme': 'dark',
    'continuous_recognition': False,
    'show_realtime_text': True,
    'audio_quality': 'high',  # Definido como alto por padrão para melhor reconhecimento
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
                print(f"Erro ao carregar configurações: {e}")
    
    def save(self):
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"Erro ao salvar configurações: {e}")
    
    def get(self, key):
        return self.data.get(key, DEFAULT_CONFIG.get(key))
    
    def set(self, key, value):
        self.data[key] = value
        self.save()

# Classe para visualização do espectro de áudio
class AudioSpectrumCanvas(FigureCanvasQTAgg):
    def __init__(self, parent=None, width=5, height=2, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = self.fig.add_subplot(111)
        super(AudioSpectrumCanvas, self).__init__(self.fig)
        self.setParent(parent)
        
        # Definir estilo do gráfico
        self.fig.patch.set_facecolor('#2E2E2E')
        self.axes.set_facecolor('#2E2E2E')
        self.axes.spines['top'].set_visible(False)
        self.axes.spines['right'].set_visible(False)
        self.axes.spines['bottom'].set_visible(False)
        self.axes.spines['left'].set_visible(False)
        self.axes.tick_params(axis='both', colors='#CCCCCC')
        
        # Configuração inicial
        self.x = np.arange(0, 100)
        self.y = np.zeros(100)
        self.line, = self.axes.plot(self.x, self.y, '-', lw=2, color='#00AAFF')
        
        # Configuração dos limites
        self.axes.set_ylim(-0.5, 0.5)
        self.axes.set_xlim(0, 100)
        self.axes.set_xticks([])
        self.axes.set_yticks([])
        
        self.fig.tight_layout(pad=0)
        self.setMinimumHeight(120)
    
    def update_plot(self, data):
        # Atualiza a visualização com novos dados
        if len(data) > 0:
            # Normaliza os dados
            normalized_data = np.frombuffer(data, dtype=np.int16).astype(np.float32)
            normalized_data = normalized_data / 32768.0  # Normaliza para -1.0 até 1.0
            
            # Atualiza somente os dados mais recentes
            data_len = min(len(normalized_data), 100)
            self.y = np.roll(self.y, -data_len)
            self.y[-data_len:] = normalized_data[:data_len]
            
            # Atualiza a linha do gráfico
            self.line.set_ydata(self.y)
            self.draw()

# Classe para processamento de áudio em thread separada
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
        
        # Inicialização do reconhecedor com configurações otimizadas
        self.recognizer = sr.Recognizer()
        # Configurações iniciais agressivas para detecção de voz
        self.recognizer.energy_threshold = 300  # Valor baixo para captar fala mais suave
        self.recognizer.dynamic_energy_threshold = True
        self.recognizer.dynamic_energy_adjustment_damping = 0.1  # Resposta mais rápida (padrão é 0.15)
        self.recognizer.dynamic_energy_ratio = 1.1  # Mais sensível a mudanças (padrão é 1.5)
        self.recognizer.pause_threshold = 0.3  # Tempo menor entre frases
        self.recognizer.phrase_threshold = 0.1  # Mais agressivo para detectar frases
        self.recognizer.non_speaking_duration = 0.2  # Menos tempo de silêncio necessário
        
        self.continuous_mode = config.get('continuous_recognition')
        self.show_realtime = config.get('show_realtime_text')
        self.last_partial_text = ""
        self.update_counter = 0
        
        # Configurações otimizadas de qualidade de áudio para fala rápida
        self.sample_rate = config.get('sample_rate')
        # Fixar em 16kHz que é o recomendado para o Google Speech API
        self.sample_rate = 16000  # Taxa ideal para reconhecimento de fala
    
    def run(self):
        # Inicializa PyAudio
        p = pyaudio.PyAudio()
        
        # Lista os dispositivos disponíveis para debug
        print("Dispositivos de áudio disponíveis:")
        for i in range(p.get_device_count()):
            dev = p.get_device_info_by_index(i)
            if dev['maxInputChannels'] > 0:  # Só mostra dispositivos com entrada
                print(f"[{i}] {dev['name']}")
                
        # Tenta usar o dispositivo padrão de entrada
        try:
            default_device_info = p.get_default_input_device_info()
            device_index = int(default_device_info['index'])
            print(f"Usando dispositivo de áudio padrão: {default_device_info['name']}")
        except Exception as e:
            print(f"Erro ao obter dispositivo padrão: {e}")
            # Busca algum dispositivo de entrada
            device_index = None
            for i in range(p.get_device_count()):
                dev = p.get_device_info_by_index(i)
                if dev['maxInputChannels'] > 0:
                    device_index = i
                    break
        
        # Configura o stream de entrada
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
                print(f"Stream aberto com sucesso usando dispositivo {device_index}")
            else:
                stream = p.open(
                    format=pyaudio.paInt16,
                    channels=1,
                    rate=self.sample_rate,
                    input=True,
                    frames_per_buffer=self.config.get('chunk_size')
                )
                print("Stream aberto com dispositivo padrão")
        except Exception as e:
            print(f"Erro ao abrir stream: {e}")
            self.progress_update.emit(100)
            self.text_ready.emit("")
            self.partial_text.emit("Erro ao abrir microfone. Verifique as configurações.")
            return
            
        # Pronto para gravar
        self.recording = True
        self.recorded_data = []
        
        # Grava o áudio
        self.partial_text.emit("Gravando áudio...")
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
                        # Atualiza a visualização parcial com maior frequência (a cada 10 chunks em vez de 15)
                        if self.update_counter % 10 == 0:
                            try:
                                # Usa os últimos 3 segundos para capturar melhor frases rápidas
                                recent_data = self.recorded_data[-45:]
                                if recent_data:
                                    audio_data = b''.join(recent_data)
                                    audio = sr.AudioData(audio_data, self.sample_rate, 2)
                                    
                                    # Usa um timeout um pouco maior para melhorar reconhecimento parcial
                                    try:
                                        # Configurações específicas para reconhecimento parcial de fala rápida
                                        self.recognizer.operation_timeout = 1.5
                                        temp_recognizer = sr.Recognizer()
                                        temp_recognizer.energy_threshold = self.recognizer.energy_threshold
                                        temp_recognizer.pause_threshold = 0.4  # Ainda mais curto para fala rápida
                                        
                                        partial_text = temp_recognizer.recognize_google(
                                            audio, 
                                            language=self.config.get('language'),
                                            show_all=False
                                        )
                                        if partial_text:
                                            self.last_partial_text = partial_text
                                            self.partial_text.emit(f"Ouvindo: {partial_text}")
                                    except:
                                        # Se falhar, mostra a última transcrição parcial bem-sucedida
                                        if self.last_partial_text:
                                            self.partial_text.emit(f"Ouvindo: {self.last_partial_text}...")
                                        else:
                                            self.partial_text.emit("Ouvindo...")
                            except Exception as partial_e:
                                print(f"Erro na transcrição parcial: {partial_e}")
            except Exception as e:
                print(f"Erro ao ler do stream: {e}")
                break
        
        # Fim da gravação
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
            self.partial_text.emit("Processando áudio...")
            self.progress_update.emit(50)
            
            # Convertendo áudio gravado para texto
            audio_data = b''.join(self.recorded_data)
            audio = sr.AudioData(audio_data, self.sample_rate, 2)  # 2 bytes por sample (16 bits)
            
            # Ajusta configurações para MELHOR RECONHECIMENTO DE FALA RÁPIDA EM PORTUGUÊS
            # Configurações agressivas para melhorar reconhecimento
            self.recognizer.pause_threshold = 0.3  # Reduzido ainda mais para melhor lidar com fala muito rápida
            self.recognizer.operation_timeout = 30  # Aumentado para dar mais tempo ao processamento completo
            self.recognizer.energy_threshold = 300  # Valor baixo para detectar fala mais suave
            
            try:
                # Reconhecimento usando Google Speech API com ajustes específicos
                self.partial_text.emit("Tentando reconhecer texto...")
                self.progress_update.emit(70)
                
                # Tentativas múltiplas de reconhecimento com diferentes configurações
                try:
                    # Primeira tentativa: configurações normais
                    text = self.recognizer.recognize_google(
                        audio, 
                        language=self.config.get('language'),
                        show_all=False,
                    )
                except sr.UnknownValueError:
                    # Segunda tentativa: com configurações alternativas
                    # Pequena pausa para reiniciar a API
                    time.sleep(0.5)
                    
                    # Criar um novo reconhecedor para segunda tentativa
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
                    # Adicionar algumas correções para palavras portuguesas comuns
                    text = self._correct_common_portuguese_errors(text)
                    self.text_ready.emit(text)
                else:
                    self.text_ready.emit("")
                    self.partial_text.emit("Nenhum texto reconhecido")
            except sr.UnknownValueError:
                self.progress_update.emit(100)
                self.text_ready.emit("")
                self.partial_text.emit("Não foi possível entender o áudio. Tente novamente falando mais devagar e claramente.")
            except sr.RequestError as e:
                self.progress_update.emit(100)
                print(f"Erro na requisição ao serviço de reconhecimento: {e}")
                self.text_ready.emit("")
                self.partial_text.emit("ERRO: Serviço de reconhecimento indisponível. Verifique sua conexão.")
        except Exception as e:
            self.progress_update.emit(100)
            print(f"Erro ao processar áudio: {e}")
            self.text_ready.emit("")
            self.partial_text.emit(f"ERRO: {str(e)}")
    
    def _correct_common_portuguese_errors(self, text):
        """Corrige erros comuns de reconhecimento em português"""
        # Dicionário de correções comuns
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
        
        # Aplicar correções
        result = text
        for wrong, correct in corrections.items():
            result = result.replace(wrong, correct)
            
        return result

# Classe para verificar se a janela atual aceita entrada de texto
class TextInputChecker:
    @staticmethod
    def is_text_input_focused():
        # Obtém a janela em foco
        hwnd = win32gui.GetForegroundWindow()
        
        # Obtém o nome da classe da janela
        try:
            class_name = win32gui.GetClassName(hwnd).lower()
        except:
            return False
        
        # Lista de classes que geralmente aceitam entrada de texto
        text_input_classes = [
            'edit', 'richedit', 'textbox', 'tbfind', 'scintilla',
            'thundrebirdwindowclass', 'mozillamaintreeclasswindow',
            'chromiumwidget', 'chrome_widget', 'webkit', 'atl:edit',
            'atom', 'vscode', 'notepad', 'wordpad', 'txview'
        ]
        
        # Verifica se a classe está na lista
        for text_class in text_input_classes:
            if text_class in class_name:
                return True
                
        # Nomes específicos de aplicações de texto
        app_names = [
            'notepad', 'word', 'wordpad', 'excel', 'powerpnt', 'onenote',
            'outlook', 'write', 'textedit', 'ultraedit', 'sublime_text',
            'atom', 'code', 'chrome', 'firefox', 'iexplore', 'opera',
            'brave', 'discord', 'slack', 'whatsapp', 'teams'
        ]
        
        # Obtém título da janela
        try:
            window_title = win32gui.GetWindowText(hwnd).lower()
            for app in app_names:
                if app in window_title or app in class_name:
                    return True
        except:
            pass
                
        # Verifica o tipo de controle para casos específicos
        try:
            # Obtém informações sobre o controle
            control_info = ctypes.create_string_buffer(1024)
            result = ctypes.windll.user32.GetClassInfoA(0, class_name.encode(), control_info)
            if result:
                # Verifica estilos específicos de controles de edição
                style = ctypes.c_long.from_buffer(control_info, 16).value
                if style & 0x00800000:  # ES_MULTILINE
                    return True
        except Exception as e:
            print(f"Erro ao verificar controle: {e}")
        
        return False

# Classe para o diálogo de teste de microfone
class MicTestDialog(QDialog):
    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self.setWindowTitle("Teste de Microfone")
        self.setMinimumSize(400, 300)
        
        # Configura layout
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # Título
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
        
        # Visualizador de nível de áudio
        level_group = QGroupBox("Nível de Áudio")
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
        
        self.text_output = QLabel("Ainda não foi reconhecido nenhum texto")
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
        
        # Botões
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
        
        # Timer para atualizar nível de áudio
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_level)
        
        # Áudio
        self.p = None
        self.stream = None
        self.audio_processor = None
        
        # Aplica estilo
        self.apply_theme()
    
    def apply_theme(self):
        """Aplica o tema à interface"""
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
        
        # Inicia processamento de áudio
        self.audio_processor = AudioProcessor(self.config)
        self.audio_processor.text_ready.connect(self.update_text)
        self.audio_processor.progress_update.connect(self.level_bar.setValue)
        self.audio_processor.start()
        
        # Inicia timer para atualizar nível de áudio
        self.timer.start(100)
    
    def stop_test(self):
        self.status_label.setText("Processando áudio...")
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
            # Simula nível de áudio para feedback visual
            import random
            self.level_bar.setValue(random.randint(30, 80))
    
    def update_text(self, text):
        if text:
            self.text_output.setText(text)
        else:
            self.text_output.setText("Não foi possível reconhecer nenhum texto")
        
        self.status_label.setText("Teste concluído")
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
        # Garantir que o processamento de áudio seja interrompido ao fechar
        if self.audio_processor and self.audio_processor.isRunning():
            self.audio_processor.stop()
            self.timer.stop()
        event.accept()

# Classe para a janela de configurações
class SettingsDialog(QDialog):
    config_changed = pyqtSignal()
    
    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self.original_hotkey = config.get('hotkey')
        self.setWindowTitle("Configurações")
        self.setMinimumSize(500, 400)
        
        # Aplicar estilo
        self.apply_theme()
        
        # Layout principal
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # Título da janela
        title_label = QLabel("Configurações do Assistente de Ditado")
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
        
        # Grupo de Auto-inicialização
        autostart_group = QGroupBox("Inicialização")
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
        
        hotkey_label = QLabel("Atalho para iniciar/parar gravação:")
        hotkey_label.setStyleSheet("color: #333333;")
        hotkey_layout.addWidget(hotkey_label)
        
        self.hotkey_edit = QLineEdit(self.config.get('hotkey'))
        self.hotkey_edit.setPlaceholderText("Pressione uma combinação de teclas")
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
        
        self.continuous_check = QCheckBox("Reconhecimento contínuo (experimental)")
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
        
        self.realtime_check = QCheckBox("Mostrar texto em tempo real durante gravação")
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
        
        # Aba de Áudio
        audio_tab = QWidget()
        audio_layout = QVBoxLayout(audio_tab)
        audio_layout.setSpacing(15)
        
        # Grupo de Qualidade de Áudio
        quality_group = QGroupBox("Qualidade de Áudio")
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
        self.quality_combo.addItem("Média (equilibrado)", "medium")
        self.quality_combo.addItem("Baixa (mais rápido)", "low")
        
        quality_idx = 0
        for i in range(self.quality_combo.count()):
            if self.quality_combo.itemData(i) == self.config.get('audio_quality'):
                quality_idx = i
                break
        self.quality_combo.setCurrentIndex(quality_idx)
        
        quality_layout.addWidget(self.quality_combo)
        
        # Botão de teste de microfone
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
        self.language_combo.addItem("Português (Brasil)", "pt-BR")
        self.language_combo.addItem("Inglês (EUA)", "en-US")
        self.language_combo.addItem("Espanhol", "es-ES")
        self.language_combo.addItem("Francês", "fr-FR")
        self.language_combo.addItem("Italiano", "it-IT")
        self.language_combo.addItem("Alemão", "de-DE")
        
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
        tab_widget.addTab(audio_tab, "Áudio")
        
        # Botões de OK e Cancelar
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
        """Aplica o tema à interface"""
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
        
        # Configurações de janela
        self.setFixedSize(500, 600)
    
    def update_sensitivity_label(self, value):
        self.sensitivity_label.setText(f"Valor: {value}%")
    
    def save_settings(self):
        # Salvar configurações
        self.config.set('auto_start', self.autostart_check.isChecked())
        self.config.set('theme', self.theme_combo.currentData())
        self.config.set('hotkey', self.hotkey_edit.text())
        self.config.set('language', self.language_combo.currentData())
        self.config.set('sensitivity', self.sensitivity_slider.value())
        self.config.set('continuous_recognition', self.continuous_check.isChecked())
        self.config.set('show_realtime_text', self.realtime_check.isChecked())
        self.config.set('audio_quality', self.quality_combo.currentData())
        
        # Configure inicialização automática
        setup_autostart(self.autostart_check.isChecked())
        
        # Emitir sinal de que as configurações foram alteradas
        self.config_changed.emit()
        
        # Fechar diálogo
        self.accept()

# Janela principal da aplicação
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
        
        # Configurar ícone da aplicação
        self.setWindowIcon(QIcon("mic_icon.png"))
        
        # Configurar bandeja do sistema
        self.setup_tray()
        
        # Configurar interface
        self.setup_ui()
        
        # Conectar sinais
        self.connect_signals()
        
        # Registrar atalho global
        self.register_hotkey()
        
        # Temporização para verificação periódica do estado do texto pendente
        self.pending_text_timer = QTimer(self)
        self.pending_text_timer.timeout.connect(self.check_pending_text)
        self.pending_text_timer.start(1000)  # Verifica a cada segundo
    
    def setup_tray(self):
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon("mic_icon.png"))
        
        # Menu da bandeja
        tray_menu = QMenu()
        
        # Ações do menu
        show_action = QAction("Mostrar", self)
        show_action.triggered.connect(self.show)
        
        settings_action = QAction("Configurações", self)
        settings_action.triggered.connect(self.show_settings)
        
        quit_action = QAction("Sair", self)
        quit_action.triggered.connect(self.quit_app)
        
        # Adicionar ações ao menu
        tray_menu.addAction(show_action)
        tray_menu.addAction(settings_action)
        tray_menu.addSeparator()
        tray_menu.addAction(quit_action)
        
        # Definir menu para o ícone da bandeja
        self.tray_icon.setContextMenu(tray_menu)
        
        # Mostrar ícone na bandeja
        self.tray_icon.show()
        
        # Conectar clique no ícone da bandeja
        self.tray_icon.activated.connect(self.tray_icon_activated)
    
    def setup_ui(self):
        # Widget principal com layout limpo
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principal com margens adequadas
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        # Cabeçalho simples
        header = QLabel("Assistente de Ditado")
        header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 22px;
            font-weight: 600;
            color: #333333;
        """)
        main_layout.addWidget(header)
        
        # === SEÇÃO DE STATUS E CONTROLE ===
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
        
        # Status atual com ícone
        status_container = QWidget()
        status_container_layout = QHBoxLayout(status_container)
        status_container_layout.setContentsMargins(10, 10, 10, 10)
        
        # Ícone de microfone
        self.mic_icon = QLabel()
        self.mic_icon.setFixedSize(32, 32)
        
        # Verificar se o arquivo de ícone existe, caso contrário criar um ícone básico
        if os.path.exists("mic_icon.png"):
            self.mic_icon.setPixmap(QPixmap("mic_icon.png").scaled(32, 32, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        else:
            # Criar um ícone básico de microfone
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
        
        # Botão de gravação principal
        self.record_button = QPushButton("Iniciar")
        self.record_button.setFixedHeight(40)
        
        # Criar ícone para o botão se não existir
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
        
        # === SEÇÃO DE TEXTO EM TEMPO REAL ===
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
        
        # Rótulo para texto em tempo real
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
        
        # === SEÇÃO DE RESULTADO ===
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
        
        # Cabeçalho de resultado
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
        
        # Área de ações para o texto reconhecido
        actions_layout = QHBoxLayout()
        actions_layout.setSpacing(10)
        
        # Botão de aceitar
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
        
        # Botão de rejeitar
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
        
        # === SEÇÃO DE CONFIGURAÇÕES RÁPIDAS ===
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
        
        # Botão de configurações
        settings_button = QPushButton("Configurações")
        
        # Criar ícone para o botão
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
        
        # Botão para testar microfone
        mic_test_button = QPushButton("Testar Microfone")
        
        # Criar ícone para o botão
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
        
        # Rodapé com dica
        footer = QLabel("Dica: Use os botões laterais do mouse (MOUSE5 e MOUSE6) para controlar o aplicativo.")
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
        # Conectar sinais do processador de áudio
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
            
            # Se for um botão do mouse
            if 'mouse' in self.original_hotkey.lower():
                # Configura o listener do mouse
                self.mouse_listener = mouse.Listener(on_click=self.on_mouse_click)
                self.mouse_listener.start()
                print(f"Registrado listener do mouse para: {self.original_hotkey}")
            else:
                # Se for uma tecla do teclado, usa a biblioteca keyboard
                try:
                    keyboard.remove_hotkey(self.original_hotkey)
                except:
                    pass
                keyboard.add_hotkey(self.original_hotkey, self.toggle_recording)
            
            # Atualizar texto do botão
            self.record_button.setText("INICIAR")
        except Exception as e:
            print(f"Erro ao registrar atalho: {e}")
            QMessageBox.warning(self, "Erro", f"Não foi possível registrar o atalho: {e}")
    
    def on_mouse_click(self, x, y, button, pressed):
        """Função de callback para eventos de clique do mouse"""
        # Converte o nome do botão para string e remove o prefixo "Button."
        button_name = str(button).replace('Button.', '')
        target_button = self.original_hotkey.lower().replace('mouse', '')
        
        # Verifica se é o botão mouse5 (x2)
        if button_name == 'x2' and self.original_hotkey.lower() == 'mouse5' and pressed:
            # Chama toggle_recording se for o botão correto
            QTimer.singleShot(0, self.toggle_recording)
            
        # Verifica se é o botão mouse6 (usado para rejeitar texto)
        elif button_name == 'x1' and pressed and self.text_collected:
            # Rejeita o texto reconhecido
            QTimer.singleShot(0, self.reject_text)
    
    def toggle_recording(self):
        """Inicia ou para a gravação de áudio"""
        
        if self.recording_active:
            # Parar gravação
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
            
            # Parar animação do microfone
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
            # Iniciar gravação
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
            
            # Iniciar animação do microfone (efeito pulsante)
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
        
        # Mostrar notificação
        self.tray_icon.showMessage(
            "Assistente de Ditado",
            "Texto reconhecido. Mouse5: colar texto | Mouse6: rejeitar texto",
            QSystemTrayIcon.MessageIcon.Information,
            3000
        )
    
    def reject_text(self):
        """Função para rejeitar o texto reconhecido"""
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
            
            # Mostrar notificação
            self.tray_icon.showMessage(
                "Assistente de Ditado",
                "TEXTO REJEITADO! Reiniciando para nova tentativa...",
                QSystemTrayIcon.MessageIcon.Warning,
                2000
            )
            
            # Reinicia a gravação automaticamente após um pequeno atraso
            if not self.recording_active:
                QTimer.singleShot(1000, self.toggle_recording)
    
    def paste_collected_text(self):
        # Verifica se há uma área de texto em foco
        if TextInputChecker.is_text_input_focused():
            # Usar a abordagem de copiar para a área de transferência e colar
            clipboard_backup = None
            try:
                # Fazer backup da área de transferência atual
                win32clipboard.OpenClipboard()
                if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                    clipboard_backup = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                win32clipboard.CloseClipboard()
                
                # Copiar o texto reconhecido para a área de transferência
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardText(self.pending_text, win32clipboard.CF_UNICODETEXT)
                win32clipboard.CloseClipboard()
                
                # Simular Ctrl+V para colar
                pyautogui.hotkey('ctrl', 'v')
                
                # Restaurar a área de transferência original após um pequeno delay
                if clipboard_backup is not None:
                    time.sleep(0.1)  # Pequena pausa para garantir que o texto foi colado
                    win32clipboard.OpenClipboard()
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardText(clipboard_backup, win32clipboard.CF_UNICODETEXT)
                    win32clipboard.CloseClipboard()
                
                # Limpar texto pendente
                self.pending_text = ""
            except Exception as e:
                print(f"Erro ao manipular área de transferência: {e}")
        else:
            # Informar que não há área de texto em foco
            self.status_label.setText("Nenhuma área editável selecionada!")
            self.tray_icon.showMessage(
                "Assistente de Ditado",
                "Não há área editável selecionada. Clique em um campo de texto e tente novamente.",
                QSystemTrayIcon.MessageIcon.Warning,
                3000
            )
    
    def update_realtime_text(self, text):
        self.realtime_label.setText(text)
    
    def apply_theme(self):
        """Aplica o tema à interface"""
        
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
        
        # Configurações de janela
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
        # Função mantida vazia só para não quebrar referências
        pass
        
    def closeEvent(self, event):
        # Para o listener do mouse antes de fechar
        if self.mouse_listener:
            self.mouse_listener.stop()
        # Certifique-se de parar qualquer gravação em andamento
        if self.audio_processor.isRunning():
            self.audio_processor.stop()
            self.audio_processor.wait()
        event.accept()

# Função para gerar um ícone básico de microfone
def create_mic_icon():
    try:
        # Criar uma imagem simples de microfone
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
        print(f"Erro ao criar ícone: {e}")

# Função para configurar a inicialização automática com o Windows
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
            # Remover da inicialização
            try:
                winreg.DeleteValue(key, app_name)
            except FileNotFoundError:
                pass
        
        winreg.CloseKey(key)
        return True
    except Exception as e:
        print(f"Erro ao configurar inicialização automática: {e}")
        return False

if __name__ == "__main__":
    app = QApplication(sys.argv)
    config = Config()
    window = MainWindow(config)
    window.show()
    sys.exit(app.exec()) 