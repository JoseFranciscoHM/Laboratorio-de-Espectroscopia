import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import threading
import time
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import json
import os
import serial
import serial.tools.list_ports
import csv
import numpy as np
import logging
import platform
from typing import List, Tuple, Any, Optional, Dict

# Configuración de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('optical_system.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Manejo multiplataforma para funciones de Windows
try:
    if platform.system() == 'Windows':
        import ctypes
        kernel32 = ctypes.windll.kernel32
        user32 = ctypes.windll.user32
        QS_ALLINPUT = 0x04FF
        OUT_PORT = 0x378
    else:
        # Implementación para Linux/Mac
        kernel32 = None
        user32 = None
        QS_ALLINPUT = 0
        OUT_PORT = 0
except Exception as e:
    logger.warning(f"No se pudieron cargar librerías específicas de Windows: {e}")
    kernel32 = None
    user32 = None
    QS_ALLINPUT = 0
    OUT_PORT = 0

class OpticalSystem:
    def __init__(self):
        self.Out_TTL = 0
        self.In_Port = 0
        self.Out_Port = OUT_PORT
        
        # Variables de estado
        self.n = 0
        self.numarchivo = 0
        self.A = 0
        self.K = 0
        self.sec = 0
        self.cont2 = 0
        self.nuevalon = 0
        self.op = 0
        self.desmagbob = 0
        self.dest = 0
        self.dato = 0
        self.Resol = 0.0
        self.angulo = 0.0
        self.angulor = 0.0
        
        # Controladores de dispositivos
        self.ser_monochromator = None
        self.ser_multimeter = None
        self.ser_lockin = None
        
        # Configuración de puertos
        self.config_file = "optical_system_config.json"
        self.default_ports = {
            "mono_port": "COM1",
            "multimeter_port": "COM4", 
            "lockin_port": "COM2",
            "start_wl": "400",
            "end_wl": "700",
            "step": "10",
            "readings": "5",
            "measurement_interval": "1.0"
        }
        
    def load_config(self) -> Dict:
        """Carga la configuración desde archivo JSON"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    logger.info("Configuración cargada exitosamente")
                    return config
            else:
                logger.info("Archivo de configuración no encontrado, usando valores por defecto")
                return self.default_ports.copy()
        except Exception as e:
            logger.error(f"Error cargando configuración: {e}")
            return self.default_ports.copy()
    
    def save_config(self, config_data: Dict):
        """Guarda la configuración en archivo JSON"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=4, ensure_ascii=False)
            logger.info("Configuración guardada exitosamente")
        except Exception as e:
            logger.error(f"Error guardando configuración: {e}")

    def setup_serial_ports(self, mono_port='COM1', multimeter_port='COM4', lockin_port='COM2') -> bool:
        """Configura todos los puertos seriales necesarios con manejo robusto de errores"""
        devices_config = {
            'monochromator': (mono_port, 9600),
            'multimeter': (multimeter_port, 9600),
            'lockin': (lockin_port, 9600)
        }
        
        successful_connections = []
        
        try:
            for device_name, (port, baudrate) in devices_config.items():
                try:
                    logger.info(f"Conectando {device_name} en {port}...")
                    
                    ser = serial.Serial(
                        port=port,
                        baudrate=baudrate,
                        bytesize=serial.EIGHTBITS,
                        parity=serial.PARITY_NONE,
                        stopbits=serial.STOPBITS_ONE,
                        timeout=2,
                        write_timeout=2
                    )
                    
                    # Test de comunicación básica
                    if device_name == 'monochromator':
                        ser.write(b'\r\n')  # Comando de eco
                        time.sleep(0.5)
                        response = ser.read_all()
                        logger.info(f"Respuesta prueba {device_name}: {response}")
                    
                    setattr(self, f"ser_{device_name}", ser)
                    successful_connections.append(device_name)
                    logger.info(f"✅ {device_name} conectado exitosamente en {port}")
                    
                except serial.SerialException as e:
                    logger.error(f"❌ Error conectando {device_name} en {port}: {e}")
                    # Continuar con otros dispositivos en lugar de fallar completamente
                    continue
            
            # Verificar que al menos el monocromador esté conectado
            if 'monochromator' in successful_connections:
                logger.info(f"Dispositivos conectados: {successful_connections}")
                return True
            else:
                logger.error("No se pudo conectar el monocromador, esencial para el sistema")
                self.close()
                return False
                
        except Exception as e:
            logger.error(f"Error inesperado en setup_serial_ports: {e}")
            self.close()
            return False

    def initialize_monochromator(self, status_callback=None) -> Optional[str]:
        """Inicializa el monocromador con manejo robusto de errores"""
        if not self.ser_monochromator or not self.ser_monochromator.is_open:
            if status_callback:
                status_callback("Error: Monocromador no conectado")
            return None
            
        try:
            # Limpiar buffer
            self.ser_monochromator.reset_input_buffer()
            self.ser_monochromator.reset_output_buffer()
            
            # Enviar comando de inicialización
            self.ser_monochromator.write(b'\r\n')
            time.sleep(1)
            
            # Leer respuesta inicial
            initial_response = self.ser_monochromator.read_all().decode('ascii', errors='ignore')
            logger.info(f"Respuesta inicial monocromador: {initial_response}")
            
            # Configurar archivo de longitud de onda actual
            config_dir = os.path.dirname(self.config_file)
            loa_file = os.path.join(config_dir, "LOA.txt")
            
            # Verificar longitud de onda actual
            try:
                if os.path.exists(loa_file):
                    with open(loa_file, "r", encoding='utf-8') as f:
                        current_wl = f.read().strip()
                else:
                    current_wl = "800"  # Valor por defecto
                    with open(loa_file, "w", encoding='utf-8') as f:
                        f.write(current_wl)
            except Exception as e:
                logger.warning(f"Error manejando archivo LOA: {e}")
                current_wl = "800"
            
            if status_callback:
                status_callback("Monocromador inicializado correctamente")
                
            return current_wl
            
        except Exception as e:
            error_msg = f"Error inicializando monocromador: {e}"
            logger.error(error_msg)
            if status_callback:
                status_callback(error_msg)
            return None

    def validate_wavelength(self, wavelength: float) -> Tuple[bool, str]:
        """Valida que la longitud de onda esté en rango permitido"""
        try:
            wl = float(wavelength)
            if wl < 200:
                return False, "Longitud de onda muy baja (mínimo 200 nm)"
            elif wl > 1300:
                return False, "Longitud de onda muy alta (máximo 1300 nm)"
            else:
                return True, f"Longitud de onda válida: {wl} nm"
        except (ValueError, TypeError):
            return False, "Longitud de onda debe ser un número válido"

    def move_monochromator(self, wavelength: float, status_callback=None) -> bool:
        """Mueve el monocromador a longitud de onda específica con validación robusta"""
        # Validar entrada
        is_valid, validation_msg = self.validate_wavelength(wavelength)
        if not is_valid:
            if status_callback:
                status_callback(f"Error: {validation_msg}")
            return False
            
        if not self.ser_monochromator or not self.ser_monochromator.is_open:
            error_msg = "Monocromador no conectado"
            if status_callback:
                status_callback(error_msg)
            logger.error(error_msg)
            return False
            
        try:
            # Obtener longitud de onda actual
            config_dir = os.path.dirname(self.config_file)
            loa_file = os.path.join(config_dir, "LOA.txt")
            current_wl = 800.0  # Valor por defecto
            
            try:
                if os.path.exists(loa_file):
                    with open(loa_file, "r", encoding='utf-8') as f:
                        current_wl = float(f.read().strip())
            except:
                pass
            
            # Limpiar buffers
            self.ser_monochromator.reset_input_buffer()
            self.ser_monochromator.reset_output_buffer()
            
            # Enviar comando GOTO
            command = f"{wavelength} GOTO\r\n"
            logger.info(f"Enviando comando al monocromador: {command.strip()}")
            
            if status_callback:
                status_callback(f"Moviendo a {wavelength} nm...")
            
            self.ser_monochromator.write(command.encode('ascii'))
            
            # Esperar y leer respuesta
            time.sleep(0.5)
            response_buffer = ""
            start_time = time.time()
            
            while time.time() - start_time < 10:  # Timeout de 10 segundos
                if self.ser_monochromator.in_waiting:
                    chunk = self.ser_monochromator.read(self.ser_monochromator.in_waiting)
                    response_buffer += chunk.decode('ascii', errors='ignore')
                    
                    # Verificar si recibimos confirmación
                    if 'ok' in response_buffer.lower() or 'done' in response_buffer.lower():
                        break
                    if 'error' in response_buffer.lower():
                        error_msg = f"Error del monocromador: {response_buffer}"
                        if status_callback:
                            status_callback(error_msg)
                        logger.error(error_msg)
                        return False
                
                time.sleep(0.1)
            
            # Actualizar archivo de longitud de onda actual
            try:
                with open(loa_file, "w", encoding='utf-8') as f:
                    f.write(str(wavelength))
            except Exception as e:
                logger.warning(f"Error actualizando archivo LOA: {e}")
            
            success_msg = f"Monocromador movido exitosamente a {wavelength} nm"
            if status_callback:
                status_callback(success_msg)
            logger.info(success_msg)
            
            return True
            
        except Exception as e:
            error_msg = f"Error moviendo monocromador: {e}"
            logger.error(error_msg)
            if status_callback:
                status_callback(error_msg)
            return False

    def read_voltage_dc(self) -> Optional[float]:
        """Lee el voltaje DC del multímetro con manejo robusto de errores"""
        if not self.ser_multimeter or not self.ser_multimeter.is_open:
            logger.error("Multímetro no conectado")
            return None
            
        try:
            self.ser_multimeter.reset_input_buffer()
            self.ser_multimeter.write(b":FETCH?\r")
            time.sleep(0.2)

            strBuffer1 = ""
            start_time = time.time()
            
            while time.time() - start_time < 3:  # Timeout de 3 segundos
                bytes_to_read = self.ser_multimeter.in_waiting
                if bytes_to_read:
                    chunk = self.ser_multimeter.read(bytes_to_read)
                    strBuffer1 += chunk.decode('ascii', errors='ignore')
                    
                    if '\r' in strBuffer1:
                        break
                
                time.sleep(0.05)

            if strBuffer1.strip():
                try:
                    # Limpiar y convertir respuesta
                    cleaned_response = strBuffer1.strip().replace('\r', '').replace('\n', '')
                    Vreg = float(cleaned_response)
                    logger.debug(f"Voltaje leído: {Vreg} V")
                    return Vreg
                except ValueError as e:
                    logger.error(f"Error convirtiendo voltaje: '{strBuffer1}' - {e}")
                    return None
            else:
                logger.warning("No se recibió respuesta del multímetro")
                return None
                
        except Exception as e:
            logger.error(f"Error leyendo voltaje: {e}")
            return None

    def read_lockin_data(self) -> Tuple[float, float]:
        """Lee datos del lock-in (canales Q1 y Q2) con manejo robusto"""
        if not self.ser_lockin or not self.ser_lockin.is_open:
            logger.error("Lock-In no conectado")
            return 0.0, 0.0
            
        try:
            n1, n2 = 0.0, 0.0
            
            # Leer canal Q1
            self.ser_lockin.reset_input_buffer()
            self.ser_lockin.write(b"Q1\r")
            time.sleep(0.3)
            response1 = self.ser_lockin.read_until(b'\r').decode('ascii', errors='ignore').strip()
            if response1:
                try:
                    n1 = float(response1)
                except ValueError:
                    logger.warning(f"Respuesta Q1 inválida: '{response1}'")
            
            # Leer canal Q2
            self.ser_lockin.reset_input_buffer()
            self.ser_lockin.write(b"Q2\r")
            time.sleep(0.3)
            response2 = self.ser_lockin.read_until(b'\r').decode('ascii', errors='ignore').strip()
            if response2:
                try:
                    n2 = float(response2)
                except ValueError:
                    logger.warning(f"Respuesta Q2 inválida: '{response2}'")
            
            logger.debug(f"Lock-In Q1: {n1}, Q2: {n2}")
            return n1, n2
            
        except Exception as e:
            logger.error(f"Error leyendo lock-in: {e}")
            return 0.0, 0.0

    def msg_wait(self, ms: int):
        """Implementa retardo multiplataforma"""
        if kernel32:
            # Implementación original para Windows
            start_time = kernel32.GetTickCount()
            elapsed = 0
            
            while elapsed < ms:
                time.sleep(0.001)
                current_time = kernel32.GetTickCount()
                elapsed = current_time - start_time
                
                if elapsed < 0:
                    start_time = current_time
                    elapsed = 0
        else:
            # Implementación para otras plataformas
            time.sleep(ms / 1000.0)

    def close(self):
        """Cierra todas las conexiones de forma segura"""
        for device_name in ['monochromator', 'multimeter', 'lockin']:
            ser = getattr(self, f"ser_{device_name}", None)
            if ser and ser.is_open:
                try:
                    ser.close()
                    logger.info(f"✅ {device_name} desconectado")
                except Exception as e:
                    logger.error(f"Error cerrando {device_name}: {e}")

class MonochromatorController:
    def __init__(self, optical_system, status_callback=None):
        self.optical_system = optical_system
        self.status_callback = status_callback
        self.current_wavelength = 0
        self.is_initialized = False
        
    def enviar_comando(self, wavelength):
        """Envía comando al monocromador usando el sistema óptico mejorado"""
        try:
            # Validar longitud de onda primero
            is_valid, message = self.optical_system.validate_wavelength(wavelength)
            if not is_valid:
                if self.status_callback:
                    self.status_callback(f"Error: {message}")
                return False
            
            # Usar el método mejorado del optical_system
            success = self.optical_system.move_monochromator(wavelength, self.status_callback)
            
            if success:
                self.current_wavelength = wavelength
                if self.status_callback:
                    self.status_callback(f"✅ Monocromador en {wavelength} nm")
            else:
                if self.status_callback:
                    self.status_callback("❌ Error moviendo monocromador")
            
            return success
            
        except Exception as e:
            error_msg = f"Error en enviar_comando: {str(e)}"
            logger.error(error_msg)
            if self.status_callback:
                self.status_callback(error_msg)
            return False
    
    def goto_wavelength(self, wavelength):
        """Mueve el monocromador a una longitud de onda específica"""
        return self.enviar_comando(wavelength)
    
    def move_monochromator(self, wavelength):
        """Función alternativa para mover el monocromador"""
        return self.enviar_comando(wavelength)
    
    def get_current_wavelength(self):
        """Obtiene la longitud de onda actual"""
        return self.current_wavelength
    
    def test_communication(self):
        """Prueba la comunicación con el monocromador"""
        if self.status_callback:
            self.status_callback("Probando comunicación con monocromador...")
        
        # Intentar mover a una posición conocida y volver
        test_wavelength = 800
        success = self.enviar_comando(test_wavelength)
        
        if success:
            if self.status_callback:
                self.status_callback("✅ Comunicación con monocromador OK")
        else:
            if self.status_callback:
                self.status_callback("❌ Error en comunicación con monocromador")
        
        return success
    
    def close(self):
        """Cierra la conexión"""
        pass

class VoltmeterController:
    def __init__(self, optical_system, status_callback=None):
        self.optical_system = optical_system
        self.status_callback = status_callback
        self.is_connected = False
        self.is_measuring_continuous = False
        self.continuous_thread = None
        self.continuous_callback = None
        self.measurement_data = []
        
    def connect(self):
        """Conecta el voltímetro"""
        self.is_connected = True
        if self.status_callback:
            self.status_callback("Voltímetro conectado")
        return True
    
    def read_voltage_dc(self):
        """Lee voltaje DC usando el sistema óptico"""
        return self.optical_system.read_voltage_dc()
    
    def read_voltage_multiple(self, num_readings=5, delay=0.1):
        """Lee múltiples voltajes y promedia"""
        readings = []
        for i in range(num_readings):
            voltage = self.read_voltage_dc()
            if voltage is not None:
                readings.append(voltage)
            time.sleep(delay)
        
        if readings:
            return sum(readings) / len(readings)
        else:
            return None

    def start_continuous_measurement(self, callback, interval=0.5):
        """Inicia medición continua"""
        if not self.is_connected:
            if self.status_callback:
                self.status_callback("Voltímetro no conectado")
            return False
            
        if self.is_measuring_continuous:
            if self.status_callback:
                self.status_callback("Medición continua ya activa")
            return False
            
        self.is_measuring_continuous = True
        self.continuous_callback = callback
        self.measurement_data = []
        
        self.continuous_thread = threading.Thread(
            target=self._continuous_measurement_worker,
            args=(interval,),
            daemon=True
        )
        self.continuous_thread.start()
        
        if self.status_callback:
            self.status_callback(f"Medición continua iniciada (intervalo: {interval}s)")
        return True
    
    def _continuous_measurement_worker(self, interval):
        """Worker para medición continua"""
        while self.is_measuring_continuous and self.is_connected:
            try:
                voltage = self.optical_system.read_voltage_dc()
                if voltage is not None:
                    timestamp = datetime.now()
                    self.measurement_data.append((timestamp, voltage))
                    
                    if self.continuous_callback:
                        self.continuous_callback(voltage, timestamp)
                        
                    if len(self.measurement_data) % 10 == 0 and self.status_callback:
                        self.status_callback(f"Medición #{len(self.measurement_data)}: {voltage:.6f} V")

            except Exception as e:
                logger.error(f"Error en medición continua: {e}")
                if self.status_callback:
                    self.status_callback(f"Error en medición continua: {str(e)}")
            
            time.sleep(interval)
    
    def stop_continuous_measurement(self):
        """Detiene la medición continua"""
        if not self.is_measuring_continuous:
            return
            
        self.is_measuring_continuous = False
        
        if self.continuous_thread:
            self.continuous_thread.join(timeout=2.0)
            self.continuous_thread = None
            
        self.continuous_callback = None
        
        if self.status_callback:
            self.status_callback(f"Medición continua detenida. Total de lecturas: {len(self.measurement_data)}")
    
    def get_measurement_data(self):
        """Obtiene los datos de medición"""
        return self.measurement_data.copy()
    
    def clear_measurement_data(self):
        """Limpia los datos de medición"""
        self.measurement_data = []
    
    def close(self):
        """Cierra el controlador"""
        self.stop_continuous_measurement()

class Reflectancia(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Sistema Completo de Caracterización Óptica - Mejorado")
        self.geometry("1200x900")
        self.configure(bg="#f0f0f0")

        # Sistema óptico principal
        self.optical_system = OpticalSystem()
        
        # Cargar configuración
        self.config_data = self.optical_system.load_config()
        
        # Controladores de dispositivos
        self.monochromator = None
        self.voltmeter = None
        
        # Datos del experimento
        self.measurement_data = []
        self.rds_data = []
        self.is_measuring = False
        self.measurement_thread = None
        self.continuous_measurement_active = False
        
        # Configuración de estilo
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.estilo()
        
        # Variables de estado
        self.status_var = tk.StringVar(value="Sistema listo")
        
        # Crear la interfaz
        self.crear_herramientas()
        
        # Configurar cierre seguro
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        """Maneja el cierre seguro de la aplicación"""
        try:
            # Guardar configuración actual
            self.optical_system.save_config(self.config_data)
            
            # Detener mediciones
            self.stop_measurement()
            
            # Cerrar conexiones
            if self.optical_system:
                self.optical_system.close()
                
            # Cerrar aplicación
            self.destroy()
            
        except Exception as e:
            logger.error(f"Error durante el cierre: {e}")
            self.destroy()

    def estilo(self):
        self.style.configure('TFrame', background="#f0f0f0")
        self.style.configure('TLabel', background="#f0f0f0", font=('Arial', 9))
        self.style.configure('TButton', font=('Arial', 9), background="#4CAF50", foreground="white")
        self.style.configure('TEntry', font=('Arial', 9))
        self.style.configure('TCombobox', font=('Arial', 9))
        self.style.configure('TLabelframe', background="#f0f0f0", font=('Arial', 10, 'bold'))
        self.style.configure('TLabelframe.Label', background="#f0f0f0", font=('Arial', 10, 'bold'))
        
        self.style.map('TButton', 
                      background=[('active', '#45a049'), ('disabled', '#cccccc')],
                      foreground=[('active', 'white'), ('disabled', '#666666')])

    def crear_herramientas(self):
        # Notebook para diferentes experimentos
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Pestaña de Configuración General
        config_frame = ttk.Frame(notebook)
        notebook.add(config_frame, text="Configuración")
        
        # Pestaña de Reflectancia
        reflectancia_frame = ttk.Frame(notebook)
        notebook.add(reflectancia_frame, text="Reflectancia")
        
        # Pestaña de Reflectancia Diferencial (RDS)
        rds_frame = ttk.Frame(notebook)
        notebook.add(rds_frame, text="RDS")
        
        # Configurar cada pestaña
        self.crear_pestana_configuracion(config_frame)
        self.crear_pestana_reflectancia(reflectancia_frame)
        self.crear_pestana_rds(rds_frame)
        
        # Barra de estado
        status_frame = ttk.Frame(self)
        status_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(status_frame, textvariable=self.status_var, 
                 foreground="#2c3e50", font=('Arial', 9)).pack(side=tk.LEFT)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(status_frame, variable=self.progress_var, mode='determinate')
        self.progress_bar.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(10, 0))

    def crear_pestana_configuracion(self, parent):
        """Crea la interfaz para configuración de puertos"""
        frame_principal = ttk.Frame(parent, padding="15")
        frame_principal.pack(fill=tk.BOTH, expand=True)

        # Título
        title_label = ttk.Label(frame_principal, text="Configuración de Puertos Seriales - Mejorado", 
                               font=("Arial", 16, "bold"), foreground="#2c3e50")
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # Puerto Monocromador
        ttk.Label(frame_principal, text="Puerto Monocromador:", 
                 font=('Arial', 10, 'bold')).grid(row=1, column=0, padx=5, pady=10, sticky="e")
        
        mono_frame = ttk.Frame(frame_principal)
        mono_frame.grid(row=1, column=1, padx=5, pady=10, sticky="w")
        
        self.puerto_mono = ttk.Combobox(mono_frame, width=12, values=self.get_serial_ports())
        self.puerto_mono.set(self.config_data.get("mono_port", "COM1"))
        self.puerto_mono.pack(side=tk.LEFT, padx=2)

        # Botón de test para monocromador
        ttk.Button(mono_frame, text="Test", 
                  command=self.test_monochromator, width=6).pack(side=tk.LEFT, padx=5)

        # Puerto Voltímetro
        ttk.Label(frame_principal, text="Puerto Voltímetro:", 
                 font=('Arial', 10, 'bold')).grid(row=2, column=0, padx=5, pady=10, sticky="e")
        
        volt_frame = ttk.Frame(frame_principal)
        volt_frame.grid(row=2, column=1, padx=5, pady=10, sticky="w")
        
        self.combobox_voltimetro = ttk.Combobox(volt_frame, width=12, values=self.get_serial_ports())
        self.combobox_voltimetro.set(self.config_data.get("voltimetro_port", "COM4"))
        self.combobox_voltimetro.pack(side=tk.LEFT, padx=2)

        # Puerto Lock-In
        ttk.Label(frame_principal, text="Puerto Lock-In:", 
                 font=('Arial', 10, 'bold')).grid(row=3, column=0, padx=5, pady=10, sticky="e")
        
        lockin_frame = ttk.Frame(frame_principal)
        lockin_frame.grid(row=3, column=1, padx=5, pady=10, sticky="w")
        
        self.puerto_lockin = ttk.Combobox(lockin_frame, width=12, values=self.get_serial_ports())
        self.puerto_lockin.set(self.config_data.get("lockin_port", "COM2"))
        self.puerto_lockin.pack(side=tk.LEFT, padx=2)

        # Botones de control
        button_frame = ttk.Frame(frame_principal)
        button_frame.grid(row=4, column=0, columnspan=3, pady=30)

        ttk.Button(button_frame, text="🔄 Refrescar Puertos", 
                  command=self.refresh_ports, width=18).pack(side=tk.LEFT, padx=8)

        ttk.Button(button_frame, text="🔌 Conectar Todo", 
                  command=self.conectar_sistema_completo, width=18).pack(side=tk.LEFT, padx=8)

        ttk.Button(button_frame, text="🔒 Desconectar Todo", 
                  command=self.desconectar_sistema, width=18).pack(side=tk.LEFT, padx=8)

        # Estado de conexión
        self.connection_status = ttk.Label(frame_principal, text="🔴 Sistema Desconectado", 
                                          foreground="red", font=("Arial", 12, "bold"))
        self.connection_status.grid(row=5, column=0, columnspan=3, pady=15)

        # Información de dispositivos
        info_frame = ttk.LabelFrame(frame_principal, text="Estado de Dispositivos")
        info_frame.grid(row=6, column=0, columnspan=3, sticky="ew", padx=5, pady=15)

        self.mono_status = ttk.Label(info_frame, text="🔴 Monocromador: Desconectado", 
                                    foreground="red", font=('Arial', 9))
        self.mono_status.pack(anchor="w", padx=10, pady=5)

        self.volt_status = ttk.Label(info_frame, text="🔴 Voltímetro: Desconectado", 
                                   foreground="red", font=('Arial', 9))
        self.volt_status.pack(anchor="w", padx=10, pady=5)

        self.lockin_status = ttk.Label(info_frame, text="🔴 Lock-In: Desconectado", 
                                     foreground="red", font=('Arial', 9))
        self.lockin_status.pack(anchor="w", padx=10, pady=5)

        # Frame para control manual
        manual_frame = ttk.LabelFrame(frame_principal, text="Control Manual del Monocromador")
        manual_frame.grid(row=7, column=0, columnspan=3, sticky="ew", padx=5, pady=15)

        ttk.Label(manual_frame, text="Longitud de onda (nm):").pack(side=tk.LEFT, padx=5, pady=5)
        self.manual_wavelength = ttk.Entry(manual_frame, width=10)
        self.manual_wavelength.insert(0, "500")
        self.manual_wavelength.pack(side=tk.LEFT, padx=5, pady=5)
        
        ttk.Button(manual_frame, text="Mover", 
                  command=self.mover_a_longitud, width=10).pack(side=tk.LEFT, padx=5, pady=5)

        self.manual_status = ttk.Label(manual_frame, text="", foreground="blue")
        self.manual_status.pack(side=tk.LEFT, padx=10, pady=5)

        # Información del sistema
        info_text = f"Sistema operativo: {platform.system()} {platform.release()}"
        system_info = ttk.Label(frame_principal, text=info_text, font=('Arial', 8), foreground="gray")
        system_info.grid(row=8, column=0, columnspan=3, pady=10)

    def crear_pestana_reflectancia(self, parent):
        """Crea la interfaz para medición de reflectancia normal"""
        frame_principal = ttk.Frame(parent, padding="10")
        frame_principal.pack(fill=tk.BOTH, expand=True)

        # Título
        title_label = ttk.Label(frame_principal, text="Medición de Reflectancia Normal", 
                               font=("Arial", 14, "bold"), foreground="#2c3e50")
        title_label.grid(row=0, column=0, columnspan=4, pady=(0, 15))

        # Frame de parámetros
        params_frame = ttk.LabelFrame(frame_principal, text="Parámetros de Medición")
        params_frame.grid(row=1, column=0, columnspan=4, sticky="ew", padx=5, pady=5)

        # Fila 1 de parámetros
        ttk.Label(params_frame, text="Longitud inicial (nm):").grid(row=0, column=0, padx=5, pady=8, sticky="e")   
        self.iniciar_longitud = ttk.Entry(params_frame, width=12)
        self.iniciar_longitud.insert(0, self.config_data.get("start_wl", "400"))
        self.iniciar_longitud.grid(row=0, column=1, padx=5, pady=8, sticky="w")

        ttk.Label(params_frame, text="Longitud final (nm):").grid(row=0, column=2, padx=5, pady=8, sticky="e")
        self.fin_longitud = ttk.Entry(params_frame, width=12)
        self.fin_longitud.insert(0, self.config_data.get("end_wl", "700"))
        self.fin_longitud.grid(row=0, column=3, padx=5, pady=8, sticky="w")

        # Fila 2 de parámetros
        ttk.Label(params_frame, text="Paso (nm):").grid(row=1, column=0, padx=5, pady=8, sticky="e")
        self.paso = ttk.Entry(params_frame, width=12)
        self.paso.insert(0, self.config_data.get("step", "10"))
        self.paso.grid(row=1, column=1, padx=5, pady=8, sticky="w")

        ttk.Label(params_frame, text="Lecturas por punto:").grid(row=1, column=2, padx=5, pady=8, sticky="e")
        self.lecturas = ttk.Entry(params_frame, width=12)
        self.lecturas.insert(0, self.config_data.get("readings", "5")) 
        self.lecturas.grid(row=1, column=3, padx=5, pady=8, sticky="w")

        # Botones de control
        button_frame = ttk.Frame(frame_principal)
        button_frame.grid(row=2, column=0, columnspan=4, padx=5, pady=15)
        
        self.start_reflectance_button = ttk.Button(button_frame, text="▶ Iniciar Reflectancia", 
                                                  command=self.start_reflectance, width=20)
        self.start_reflectance_button.pack(side=tk.LEFT, padx=5)

        self.stop_reflectance_button = ttk.Button(button_frame, text="⏹ Detener", 
                                                 command=self.stop_measurement, state=tk.DISABLED, width=15)
        self.stop_reflectance_button.pack(side=tk.LEFT, padx=5)

        ttk.Button(button_frame, text="💾 Guardar Datos", 
                  command=self.save_reflectance_data, width=15).pack(side=tk.LEFT, padx=5)

        # Display en tiempo real
        display_frame = ttk.LabelFrame(frame_principal, text="Medición en Tiempo Real")
        display_frame.grid(row=3, column=0, columnspan=4, sticky="ew", padx=5, pady=10)

        ttk.Label(display_frame, text="Voltaje actual:").pack(side=tk.LEFT, padx=10, pady=5)
        self.voltage_display = ttk.Label(display_frame, text="--- V", font=("Arial", 11, "bold"), 
                                        foreground="blue")
        self.voltage_display.pack(side=tk.LEFT, padx=5, pady=5)

        ttk.Label(display_frame, text="Longitud actual:").pack(side=tk.LEFT, padx=10, pady=5)
        self.current_wl_display = ttk.Label(display_frame, text="--- nm", font=("Arial", 11, "bold"), 
                                          foreground="green")
        self.current_wl_display.pack(side=tk.LEFT, padx=5, pady=5)

        # Gráfica
        graph_frame = ttk.LabelFrame(frame_principal, text="Gráfica de Reflectancia", padding="10")
        graph_frame.grid(row=4, column=0, columnspan=4, sticky="nsew", padx=5, pady=5)
        frame_principal.rowconfigure(4, weight=1)
        frame_principal.columnconfigure(0, weight=1)

        self.figure_reflectance, self.ax_reflectance = plt.subplots(figsize=(8, 5))
        self.ax_reflectance.set_xlabel("Longitud de onda (nm)", fontsize=10)
        self.ax_reflectance.set_ylabel("Voltaje (V)", fontsize=10)
        self.ax_reflectance.grid(True, alpha=0.3)
        self.canvas_reflectance = FigureCanvasTkAgg(self.figure_reflectance, master=graph_frame)
        self.canvas_reflectance.get_tk_widget().pack(fill="both", expand=True)

        # Barra de herramientas para gráfica
        toolbar_frame = ttk.Frame(graph_frame)
        toolbar_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(toolbar_frame, text="Zoom +", command=self.zoom_in_reflectance, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar_frame, text="Zoom -", command=self.zoom_out_reflectance, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar_frame, text="Autoajustar", command=self.autoscale_reflectance, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar_frame, text="Limpiar", command=self.clear_reflectance_plot, width=8).pack(side=tk.LEFT, padx=2)

    def crear_pestana_rds(self, parent):
        """Crea la interfaz para Reflectancia Diferencial (RDS) - Mejorado"""
        frame_principal = ttk.Frame(parent, padding="10")
        frame_principal.pack(fill=tk.BOTH, expand=True)

        # Título
        title_label = ttk.Label(frame_principal, text="Reflectancia Diferencial (RDS)", 
                               font=("Arial", 14, "bold"), foreground="#2c3e50")
        title_label.grid(row=0, column=0, columnspan=4, pady=(0, 15))

        # Frame de parámetros RDS
        params_frame = ttk.LabelFrame(frame_principal, text="Parámetros RDS")
        params_frame.grid(row=1, column=0, columnspan=4, sticky="ew", padx=5, pady=5)

        # Fila 1 de parámetros
        ttk.Label(params_frame, text="Longitud inicial (nm):").grid(row=0, column=0, padx=5, pady=8, sticky="e")   
        self.rds_start_wl = ttk.Entry(params_frame, width=12)
        self.rds_start_wl.insert(0, "400")
        self.rds_start_wl.grid(row=0, column=1, padx=5, pady=8, sticky="w")

        ttk.Label(params_frame, text="Longitud final (nm):").grid(row=0, column=2, padx=5, pady=8, sticky="e")
        self.rds_end_wl = ttk.Entry(params_frame, width=12)
        self.rds_end_wl.insert(0, "700")
        self.rds_end_wl.grid(row=0, column=3, padx=5, pady=8, sticky="w")

        # Fila 2 de parámetros
        ttk.Label(params_frame, text="Paso (nm):").grid(row=1, column=0, padx=5, pady=8, sticky="e")
        self.rds_step = ttk.Entry(params_frame, width=12)
        self.rds_step.insert(0, "5")
        self.rds_step.grid(row=1, column=1, padx=5, pady=8, sticky="w")

        ttk.Label(params_frame, text="Lecturas Lock-In:").grid(row=1, column=2, padx=5, pady=8, sticky="e")
        self.rds_lockin_readings = ttk.Entry(params_frame, width=12)
        self.rds_lockin_readings.insert(0, "10")
        self.rds_lockin_readings.grid(row=1, column=3, padx=5, pady=8, sticky="w")

        # Fila 3 de parámetros
        ttk.Label(params_frame, text="Lecturas Voltímetro:").grid(row=2, column=0, padx=5, pady=8, sticky="e")
        self.rds_volt_readings = ttk.Entry(params_frame, width=12)
        self.rds_volt_readings.insert(0, "5")
        self.rds_volt_readings.grid(row=2, column=1, padx=5, pady=8, sticky="w")

        ttk.Label(params_frame, text="Tiempo estabilización (s):").grid(row=2, column=2, padx=5, pady=8, sticky="e")
        self.rds_stabilization = ttk.Entry(params_frame, width=12)
        self.rds_stabilization.insert(0, "3.0")
        self.rds_stabilization.grid(row=2, column=3, padx=5, pady=8, sticky="w")

        # Botones de control RDS
        button_frame = ttk.Frame(frame_principal)
        button_frame.grid(row=2, column=0, columnspan=4, padx=5, pady=15)

        self.start_rds_button = ttk.Button(button_frame, text="▶ Iniciar RDS", 
                                          command=self.start_rds_measurement, width=18)
        self.start_rds_button.pack(side=tk.LEFT, padx=5)

        ttk.Button(button_frame, text="🔍 Test Lock-In", 
                  command=self.test_lockin, width=15).pack(side=tk.LEFT, padx=5)

        self.stop_rds_button = ttk.Button(button_frame, text="⏹ Parar", 
                                         command=self.stop_measurement, state=tk.DISABLED, width=12)
        self.stop_rds_button.pack(side=tk.LEFT, padx=5)

        ttk.Button(button_frame, text="💾 Guardar RDS", 
                  command=self.save_rds_data, width=15).pack(side=tk.LEFT, padx=5)

        # Display de datos en tiempo real MEJORADO
        display_frame = ttk.LabelFrame(frame_principal, text="Datos en Tiempo Real - RDS")
        display_frame.grid(row=3, column=0, columnspan=4, sticky="ew", padx=5, pady=10)

        # Crear un frame organizado para los displays
        display_grid = ttk.Frame(display_frame)
        display_grid.pack(padx=10, pady=8, fill=tk.X)
        
        # Fila 1: Señales principales
        row1 = ttk.Frame(display_grid)
        row1.pack(fill=tk.X, pady=2)
        
        ttk.Label(row1, text="R (Voltaje DC):", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=8)
        self.rds_voltage_display = ttk.Label(row1, text="--- V", foreground="blue", font=("Arial", 9))
        self.rds_voltage_display.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row1, text="ΔR (Lock-In Q1):", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=8)
        self.lockin_q1_display = ttk.Label(row1, text="---", foreground="red", font=("Arial", 9))
        self.lockin_q1_display.pack(side=tk.LEFT, padx=5)
        
        # Fila 2: Señales secundarias y resultado
        row2 = ttk.Frame(display_grid)
        row2.pack(fill=tk.X, pady=2)
        
        ttk.Label(row2, text="Q2:", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=8)
        self.lockin_q2_display = ttk.Label(row2, text="---", foreground="green", font=("Arial", 9))
        self.lockin_q2_display.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row2, text="ΔR/R:", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=8)
        self.delta_r_display = ttk.Label(row2, text="---", foreground="purple", font=("Arial", 10, "bold"))
        self.delta_r_display.pack(side=tk.LEFT, padx=5)

        # Gráfica RDS
        graph_frame = ttk.LabelFrame(frame_principal, text="Gráfica RDS - ΔR/R vs Longitud de Onda", padding="10")
        graph_frame.grid(row=4, column=0, columnspan=4, sticky="nsew", padx=5, pady=5)
        frame_principal.rowconfigure(4, weight=1)
        frame_principal.columnconfigure(0, weight=1)

        self.figure_rds, self.ax_rds = plt.subplots(figsize=(8, 5))
        self.ax_rds.set_xlabel("Longitud de onda (nm)", fontsize=10)
        self.ax_rds.set_ylabel("ΔR/R", fontsize=10)
        self.ax_rds.grid(True, alpha=0.3)
        self.canvas_rds = FigureCanvasTkAgg(self.figure_rds, master=graph_frame)
        self.canvas_rds.get_tk_widget().pack(fill="both", expand=True)

        # Barra de herramientas para gráfica RDS
        toolbar_frame = ttk.Frame(graph_frame)
        toolbar_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(toolbar_frame, text="Zoom +", command=self.zoom_in_rds, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar_frame, text="Zoom -", command=self.zoom_out_rds, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar_frame, text="Autoajustar", command=self.autoscale_rds, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar_frame, text="Limpiar", command=self.clear_rds_plot, width=8).pack(side=tk.LEFT, padx=2)

        # Barra de progreso y estado RDS
        self.rds_progress = ttk.Progressbar(frame_principal, mode='determinate')
        self.rds_progress.grid(row=5, column=0, columnspan=4, sticky="ew", padx=5, pady=5)

        self.rds_status = ttk.Label(frame_principal, text="Listo para medición RDS", 
                                   foreground="blue", font=('Arial', 9))
        self.rds_status.grid(row=6, column=0, columnspan=4, pady=5)

    # ===== FUNCIONES DE CONEXIÓN =====

    def conectar_sistema_completo(self):
        """Conecta todos los dispositivos del sistema óptico"""
        try:
            # Actualizar configuración
            self.config_data.update({
                "mono_port": self.puerto_mono.get(),
                "voltimetro_port": self.combobox_voltimetro.get(),
                "lockin_port": self.puerto_lockin.get()
            })
            
            success = self.optical_system.setup_serial_ports(
                mono_port=self.config_data["mono_port"],
                multimeter_port=self.config_data["voltimetro_port"],
                lockin_port=self.config_data["lockin_port"]
            )
            
            if success:
                # Inicializar controladores
                self.monochromator = MonochromatorController(self.optical_system, self.actualizar_estado)
                self.voltmeter = VoltmeterController(self.optical_system, self.actualizar_estado)
                self.voltmeter.connect()
                
                # Inicializar monocromador
                current_wl = self.optical_system.initialize_monochromator(self.actualizar_estado)
                
                # Actualizar estados
                self.connection_status.config(text="🟢 Sistema Conectado", foreground="green")
                self.mono_status.config(text="🟢 Monocromador: Conectado", foreground="green")
                self.volt_status.config(text="🟢 Voltímetro: Conectado", foreground="green")
                self.lockin_status.config(text="🟢 Lock-In: Conectado", foreground="green")
                
                self.status_var.set(f"Sistema conectado correctamente. Longitud actual: {current_wl} nm")
                
                # Guardar configuración
                self.optical_system.save_config(self.config_data)
                
                # Probar dispositivos
                self.probar_dispositivos()
                
            else:
                messagebox.showerror("Error", "No se pudo conectar el sistema completo")
                
        except Exception as e:
            logger.error(f"Error conectando sistema: {e}")
            messagebox.showerror("Error", f"Error conectando sistema: {str(e)}")

    def desconectar_sistema(self):
        """Desconecta todos los dispositivos"""
        if self.optical_system:
            self.optical_system.close()
        
        self.connection_status.config(text="🔴 Sistema Desconectado", foreground="red")
        self.mono_status.config(text="🔴 Monocromador: Desconectado", foreground="red")
        self.volt_status.config(text="🔴 Voltímetro: Desconectado", foreground="red")
        self.lockin_status.config(text="🔴 Lock-In: Desconectado", foreground="red")
        
        self.status_var.set("Sistema desconectado")
        messagebox.showinfo("Desconexión", "Sistema desconectado")

    def test_monochromator(self):
        """Prueba la comunicación con el monocromador"""
        if not self.monochromator:
            messagebox.showwarning("Advertencia", "Conecte el sistema primero")
            return
        
        def do_test():
            success = self.monochromator.test_communication()
            self.after(0, lambda: messagebox.showinfo(
                "Test Monocromador", 
                "Comunicación exitosa" if success else "Error en comunicación"
            ))
        
        threading.Thread(target=do_test, daemon=True).start()

    def probar_dispositivos(self):
        """Prueba todos los dispositivos después de conectar"""
        def prueba():
            # Probar voltímetro
            voltage = self.optical_system.read_voltage_dc()
            if voltage is not None:
                self.after(0, lambda: self.voltage_display.config(text=f"{voltage:.6f} V"))
                self.after(0, lambda: self.rds_voltage_display.config(text=f"{voltage:.6f} V"))
            
            # Probar lock-in
            q1, q2 = self.optical_system.read_lockin_data()
            self.after(0, lambda: self.lockin_q1_display.config(text=f"{q1:.4f}"))
            self.after(0, lambda: self.lockin_q2_display.config(text=f"{q2:.4f}"))
            
            # Calcular ΔR/R inicial si hay datos
            if voltage is not None and q1 != 0:
                delta_r = q1 / voltage
                self.after(0, lambda: self.delta_r_display.config(text=f"{delta_r:.6f}"))
        
        threading.Thread(target=prueba, daemon=True).start()

    def test_lockin(self):
        """Prueba rápida del Lock-In"""
        if not self.optical_system.ser_lockin:
            messagebox.showwarning("Advertencia", "Lock-In no conectado")
            return
        
        def test():
            for i in range(3):  # Tomar 3 lecturas
                q1, q2 = self.optical_system.read_lockin_data()
                self.after(0, lambda q1=q1, q2=q2: self._update_lockin_display(q1, q2))
                time.sleep(0.5)
            
            self.after(0, lambda: messagebox.showinfo("Test Lock-In", 
                     f"Lecturas estables:\nQ1 (ΔR): {q1:.4f}\nQ2: {q2:.4f}"))
        
        threading.Thread(target=test, daemon=True).start()

    def mover_a_longitud(self):
        """Mueve el monocromador a longitud manual"""
        if not self.monochromator:
            messagebox.showwarning("Advertencia", "Conecte el sistema primero")
            return
            
        try:
            wavelength = float(self.manual_wavelength.get())
            if wavelength <= 0 or wavelength > 2000:
                messagebox.showerror("Error", "Longitud de onda debe estar entre 1 y 2000 nm")
                return
                
            self.manual_status.config(text="Moviendo...", foreground="orange")
            threading.Thread(target=self._mover_mono_thread, args=(wavelength,), daemon=True).start()
            
        except ValueError:
            messagebox.showerror("Error", "Longitud de onda inválida")

    def _mover_mono_thread(self, wavelength):
        """Hilo para mover monocromador"""
        try:
            success = self.monochromator.enviar_comando(wavelength)
            if success:
                self.after(0, lambda: self.manual_status.config(
                    text=f"✅ En {wavelength} nm", foreground="green"))
                self.after(0, lambda: self.current_wl_display.config(
                    text=f"{wavelength} nm"))
            else:
                self.after(0, lambda: self.manual_status.config(
                    text="❌ Error moviendo", foreground="red"))
        except Exception as e:
            self.after(0, lambda: self.manual_status.config(
                text=f"❌ Error: {str(e)}", foreground="red"))

    # ===== FUNCIONES PARA REFLECTANCIA NORMAL =====

    def start_reflectance(self):
        """Inicia medición de reflectancia normal"""
        if not self.monochromator:
            messagebox.showwarning("Advertencia", "Conecte el sistema primero")
            return
            
        if not self.voltmeter or not self.voltmeter.is_connected:
            messagebox.showwarning("Advertencia", "Voltímetro no conectado")
            return
            
        try:
            start_wl = float(self.iniciar_longitud.get())
            end_wl = float(self.fin_longitud.get())
            step = float(self.paso.get())
            readings_per_point = int(self.lecturas.get())
            
            if start_wl >= end_wl:
                messagebox.showerror("Error", "La longitud inicial debe ser menor que la final")
                return
                
            if step <= 0:
                messagebox.showerror("Error", "El paso debe ser mayor que 0")
                return
                
            self.is_measuring = True
            self.start_reflectance_button.config(state=tk.DISABLED)
            self.stop_reflectance_button.config(state=tk.NORMAL)
            self.measurement_data = []
            
            self.measurement_thread = threading.Thread(
                target=self._reflectance_measurement_worker,
                args=(start_wl, end_wl, step, readings_per_point),
                daemon=True
            )
            self.measurement_thread.start()
            
        except ValueError as e:
            messagebox.showerror("Error", "Valores de entrada inválidos")

    def _reflectance_measurement_worker(self, start_wl, end_wl, step, readings_per_point):
        """Hilo de trabajo para medición de reflectancia normal"""
        try:
            self.measurement_data = []
            self.after(0, self.clear_reflectance_plot)
            
            total_points = int((end_wl - start_wl) / step) + 1
            current_point = 0
            
            # Función para calcular tiempo de estabilización
            def get_stabilization_time(step_size):
                if step_size <= 0.2:
                    return 3.0
                elif step_size <= 0.5:
                    return 2.0
                elif step_size <= 1.0:
                    return 1.5
                else:
                    return 1.0
            
            wavelength = start_wl
            while wavelength <= end_wl and self.is_measuring:
                # Mover monocromador
                success = self.monochromator.enviar_comando(wavelength)
                if not success:
                    self.after(0, lambda: self.status_var.set(f"Error moviendo a {wavelength} nm"))
                    break
                
                # Esperar estabilización
                stabilization_time = get_stabilization_time(step)
                self.after(0, lambda: self.status_var.set(
                    f"Estabilizando en {wavelength} nm ({stabilization_time}s)..."
                ))
                time.sleep(stabilization_time)
                
                if not self.is_measuring:
                    break
                    
                # Tomar lecturas del voltímetro
                voltages = []
                for i in range(readings_per_point):
                    if not self.is_measuring:
                        break
                        
                    voltage = self.optical_system.read_voltage_dc()
                    if voltage is not None:
                        voltages.append(voltage)
                        
                        # Actualizar display en tiempo real
                        self.after(0, lambda v=voltage: self.voltage_display.config(
                            text=f"{v:.6f} V"
                        ))
                        self.after(0, lambda wl=wavelength: self.current_wl_display.config(
                            text=f"{wl} nm"
                        ))
                    
                    time.sleep(0.3)
                
                if voltages:
                    avg_voltage = sum(voltages) / len(voltages)
                    
                    # Guardar dato
                    data_point = {
                        'wavelength': wavelength,
                        'voltage': avg_voltage,
                        'timestamp': datetime.now(),
                        'readings': voltages
                    }
                    self.measurement_data.append(data_point)
                    
                    # Actualizar gráfica
                    self.after(0, lambda wl=wavelength, v=avg_voltage: 
                              self._update_reflectance_plot(wl, v))
                    
                    self.after(0, lambda: self.status_var.set(
                        f"Medido: {wavelength} nm, {avg_voltage:.6f} V"
                    ))
                
                # Actualizar progreso
                current_point += 1
                progress = (current_point / total_points) * 100
                self.after(0, lambda p=progress: self.progress_var.set(p))
                
                wavelength += step
            
            # Medición completada
            self.after(0, self._reflectance_measurement_completed)
            
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Error", f"Error en medición: {str(e)}"))
            self.after(0, self._reflectance_measurement_completed)

    def _update_reflectance_plot(self, wavelength, voltage):
        """Actualiza la gráfica de reflectancia normal"""
        wavelengths = [point['wavelength'] for point in self.measurement_data]
        voltages = [point['voltage'] for point in self.measurement_data]
        
        self.ax_reflectance.clear()
        self.ax_reflectance.plot(wavelengths, voltages, 'b-', linewidth=1.5, marker='o', markersize=4, label='Reflectancia')
        self.ax_reflectance.set_xlabel("Longitud de onda (nm)", fontsize=10)
        self.ax_reflectance.set_ylabel("Voltaje (V)", fontsize=10)
        self.ax_reflectance.grid(True, alpha=0.3)
        self.ax_reflectance.set_title("Curva de Reflectancia")
        self.ax_reflectance.legend()
        
        self.canvas_reflectance.draw()

    def _reflectance_measurement_completed(self):
        """Limpieza después de completar medición de reflectancia"""
        self.is_measuring = False
        self.start_reflectance_button.config(state=tk.NORMAL)
        self.stop_reflectance_button.config(state=tk.DISABLED)
        self.status_var.set("Medición de reflectancia completada")
        self.progress_var.set(0)
        
        if self.measurement_data:
            messagebox.showinfo("Completado", 
                              f"Medición completada. {len(self.measurement_data)} puntos medidos.")

    def save_reflectance_data(self):
        """Guarda los datos de reflectancia en CSV"""
        if not self.measurement_data:
            messagebox.showwarning("Advertencia", "No hay datos de reflectancia para guardar")
            return
            
        try:
            filename = f"reflectancia_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
                fieldnames = ['wavelength_nm', 'voltage_V', 'timestamp']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                
                writer.writeheader()
                for point in self.measurement_data:
                    writer.writerow({
                        'wavelength_nm': point['wavelength'],
                        'voltage_V': point['voltage'],
                        'timestamp': point['timestamp'].strftime('%Y-%m-%d %H:%M:%S')
                    })
            
            messagebox.showinfo("Guardado", f"Datos guardados en {filename}")
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo: {str(e)}")

    # ===== FUNCIONES PARA RDS =====

    def start_rds_measurement(self):
        """Inicia medición de Reflectancia Diferencial"""
        if not self.monochromator:
            messagebox.showwarning("Advertencia", "Conecte el sistema primero")
            return
            
        try:
            start_wl = float(self.rds_start_wl.get())
            end_wl = float(self.rds_end_wl.get())
            step = float(self.rds_step.get())
            lockin_readings = int(self.rds_lockin_readings.get())
            volt_readings = int(self.rds_volt_readings.get())
            stabilization = float(self.rds_stabilization.get())
            
            if start_wl >= end_wl:
                messagebox.showerror("Error", "La longitud inicial debe ser menor que la final")
                return
                
            if step <= 0:
                messagebox.showerror("Error", "El paso debe ser mayor que 0")
                return
            
            self.is_measuring = True
            self.start_rds_button.config(state=tk.DISABLED)
            self.stop_rds_button.config(state=tk.NORMAL)
            self.rds_data = []
            
            # Iniciar medición en hilo separado
            self.measurement_thread = threading.Thread(
                target=self._rds_measurement_worker,
                args=(start_wl, end_wl, step, lockin_readings, volt_readings, stabilization),
                daemon=True
            )
            self.measurement_thread.start()
            
        except ValueError as e:
            messagebox.showerror("Error", "Valores de entrada inválidos")

    def _rds_measurement_worker(self, start_wl, end_wl, step, lockin_readings, volt_readings, stabilization):
        """Hilo de trabajo para medición RDS - Mejorado para calcular R, ΔR y ΔR/R"""
        try:
            self.rds_data = []
            total_points = int((end_wl - start_wl) / step) + 1
            current_point = 0
            
            wavelength = start_wl
            while wavelength <= end_wl and self.is_measuring:
                # Mover monocromador
                success = self.monochromator.enviar_comando(wavelength)
                if not success:
                    break
                
                # Esperar estabilización
                self.after(0, lambda: self.rds_status.config(
                    text=f"Estabilizando en {wavelength} nm ({stabilization}s)...",
                    foreground="orange"
                ))
                time.sleep(stabilization)
                
                if not self.is_measuring:
                    break
                
                # Tomar lecturas del Lock-In (señal diferencial ΔR)
                q1_readings = []
                q2_readings = []
                for i in range(lockin_readings):
                    if not self.is_measuring:
                        break
                    q1, q2 = self.optical_system.read_lockin_data()
                    q1_readings.append(q1)
                    q2_readings.append(q2)
                    
                    # Actualizar display en tiempo real
                    self.after(0, lambda q1=q1, q2=q2: self._update_lockin_display(q1, q2))
                    time.sleep(0.2)
                
                # Tomar lecturas del voltímetro (señal DC R)
                R_readings = []
                for i in range(volt_readings):
                    if not self.is_measuring:
                        break
                    voltage = self.optical_system.read_voltage_dc()
                    if voltage is not None:
                        R_readings.append(voltage)
                        self.after(0, lambda v=voltage: self.rds_voltage_display.config(
                            text=f"{v:.6f} V"
                        ))
                    time.sleep(0.1)
                
                if q1_readings and R_readings:
                    # Calcular promedios
                    avg_delta_R = sum(q1_readings) / len(q1_readings)  # ΔR = Q1
                    avg_q2 = sum(q2_readings) / len(q2_readings)
                    avg_R = sum(R_readings) / len(R_readings)          # R = Voltaje DC
                    
                    # Calcular ΔR/R
                    if avg_R != 0:
                        delta_R_over_R = avg_delta_R / avg_R
                    else:
                        delta_R_over_R = 0
                    
                    # Guardar datos completos
                    data_point = {
                        'wavelength': wavelength,
                        'energy': 1239.4 / wavelength if wavelength != 0 else 0,
                        'q1': avg_delta_R,        # Esto es ΔR
                        'q2': avg_q2,             # Canal Q2 adicional
                        'voltage': avg_R,         # Esto es R (señal DC)
                        'delta_r_over_r': delta_R_over_R,  # ΔR/R
                        'timestamp': datetime.now()
                    }
                    self.rds_data.append(data_point)
                    
                    # Actualizar gráfica
                    self.after(0, lambda wl=wavelength, dr=delta_R_over_R: 
                              self._update_rds_plot(wl, dr))
                    
                    # Actualizar display con todos los valores
                    self.after(0, lambda: self.delta_r_display.config(
                        text=f"{delta_R_over_R:.6f}"
                    ))
                    
                    # Mostrar todos los valores en el status
                    self.after(0, lambda: self.rds_status.config(
                        text=f"Medido: {wavelength} nm | R: {avg_R:.4f}V | ΔR: {avg_delta_R:.4f} | ΔR/R: {delta_R_over_R:.6f}",
                        foreground="green"
                    ))
                
                # Actualizar progreso
                current_point += 1
                progress = (current_point / total_points) * 100
                self.after(0, lambda p=progress: self.rds_progress.config(value=p))
                
                wavelength += step
            
            # Medición completada
            self.after(0, self._rds_measurement_completed)
            
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Error", f"Error en medición RDS: {str(e)}"))
            self.after(0, self._rds_measurement_completed)

    def _update_lockin_display(self, q1, q2):
        """Actualiza el display del Lock-In con nombres más descriptivos"""
        self.lockin_q1_display.config(text=f"{q1:.4f}")  # Esto es ΔR
        self.lockin_q2_display.config(text=f"{q2:.4f}")  # Q2 adicional

    def _update_rds_plot(self, wavelength, delta_r_over_r):
        """Actualiza la gráfica RDS"""
        wavelengths = [point['wavelength'] for point in self.rds_data]
        delta_r_values = [point['delta_r_over_r'] for point in self.rds_data]
        
        self.ax_rds.clear()
        self.ax_rds.plot(wavelengths, delta_r_values, 'r-', linewidth=1.5, marker='o', markersize=4, label='ΔR/R')
        self.ax_rds.set_xlabel("Longitud de onda (nm)", fontsize=10)
        self.ax_rds.set_ylabel("ΔR/R", fontsize=10)
        self.ax_rds.grid(True, alpha=0.3)
        self.ax_rds.set_title("Reflectancia Diferencial (RDS)")
        self.ax_rds.legend()
        
        # Ajustar escala Y automáticamente
        if delta_r_values:
            y_max = max(abs(min(delta_r_values)), abs(max(delta_r_values)))
            if y_max > 0:
                self.ax_rds.set_ylim(-y_max * 1.1, y_max * 1.1)
        
        self.canvas_rds.draw()

    def _rds_measurement_completed(self):
        """Limpieza después de completar medición RDS"""
        self.is_measuring = False
        self.start_rds_button.config(state=tk.NORMAL)
        self.stop_rds_button.config(state=tk.DISABLED)
        self.rds_status.config(text="Medición RDS completada", foreground="green")
        self.rds_progress.config(value=0)
        
        if self.rds_data:
            messagebox.showinfo("Completado", 
                              f"Medición RDS completada. {len(self.rds_data)} puntos medidos.")
            # Guardar datos automáticamente
            self.save_rds_data()

    def save_rds_data(self):
        """Guarda los datos RDS en CSV con columnas: wavelength, energy, R, ΔR, ΔR/R"""
        if not self.rds_data:
            messagebox.showwarning("Advertencia", "No hay datos RDS para guardar")
            return
            
        try:
            filename = f"rds_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
                # Columnas completas para análisis RDS
                fieldnames = [
                    'wavelength_nm', 
                    'energy_eV', 
                    'R_voltage',      # Señal DC (Reflectancia)
                    'delta_R',        # Señal AC del Lock-In (ΔR = Q1)
                    'delta_R_over_R', # ΔR/R calculado
                    'Q2',             # Canal Q2 adicional
                    'timestamp'
                ]
                
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                
                for point in self.rds_data:
                    writer.writerow({
                        'wavelength_nm': point['wavelength'],
                        'energy_eV': point['energy'],
                        'R_voltage': point['voltage'],           # R
                        'delta_R': point['q1'],                  # ΔR
                        'delta_R_over_R': point['delta_r_over_r'], # ΔR/R
                        'Q2': point['q2'],
                        'timestamp': point['timestamp'].strftime('%Y-%m-%d %H:%M:%S')
                    })
            
            # Mensaje informativo sobre las columnas guardadas
            messagebox.showinfo("Guardado", 
                f"Datos RDS guardados en {filename}\n\n"
                "Columnas incluidas:\n"
                "• Longitud de onda (nm)\n"
                "• Energía (eV)\n" 
                "• R (Voltaje DC - Reflectancia)\n"
                "• ΔR (Señal Lock-In Q1)\n"
                "• ΔR/R (Relación calculada)\n"
                "• Q2 (Canal adicional Lock-In)\n"
                "• Timestamp")
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar archivo RDS: {str(e)}")

    # ===== FUNCIONES UTILITARIAS =====

    def stop_measurement(self):
        """Detiene cualquier medición en curso"""
        self.is_measuring = False
        self.status_var.set("Medición detenida por el usuario")
        self.rds_status.config(text="Medición detenida", foreground="orange")

    def get_serial_ports(self):
        """Obtiene lista de puertos seriales disponibles"""
        ports = serial.tools.list_ports.comports()
        return [port.device for port in ports]

    def refresh_ports(self):
        """Actualiza la lista de puertos disponibles"""
        ports = self.get_serial_ports()
        self.puerto_mono['values'] = ports
        self.combobox_voltimetro['values'] = ports
        self.puerto_lockin['values'] = ports
        self.status_var.set("Lista de puertos actualizada")

    def actualizar_estado(self, mensaje):
        """Actualiza el estado del sistema"""
        def actualizar():
            self.status_var.set(mensaje)
        self.after(0, actualizar)

    # ===== FUNCIONES PARA GESTIÓN DE GRÁFICAS =====

    def zoom_in_reflectance(self):
        """Zoom in en gráfica de reflectancia"""
        xlim = self.ax_reflectance.get_xlim()
        ylim = self.ax_reflectance.get_ylim()
        self.ax_reflectance.set_xlim(xlim[0] * 0.8, xlim[1] * 0.8)
        self.ax_reflectance.set_ylim(ylim[0] * 0.8, ylim[1] * 0.8)
        self.canvas_reflectance.draw()

    def zoom_out_reflectance(self):
        """Zoom out en gráfica de reflectancia"""
        xlim = self.ax_reflectance.get_xlim()
        ylim = self.ax_reflectance.get_ylim()
        self.ax_reflectance.set_xlim(xlim[0] * 1.2, xlim[1] * 1.2)
        self.ax_reflectance.set_ylim(ylim[0] * 1.2, ylim[1] * 1.2)
        self.canvas_reflectance.draw()

    def autoscale_reflectance(self):
        """Autoajusta gráfica de reflectancia"""
        self.ax_reflectance.relim()
        self.ax_reflectance.autoscale_view()
        self.canvas_reflectance.draw()

    def clear_reflectance_plot(self):
        """Limpia gráfica de reflectancia"""
        self.ax_reflectance.clear()
        self.ax_reflectance.set_xlabel("Longitud de onda (nm)", fontsize=10)
        self.ax_reflectance.set_ylabel("Voltaje (V)", fontsize=10)
        self.ax_reflectance.grid(True, alpha=0.3)
        self.canvas_reflectance.draw()

    def zoom_in_rds(self):
        """Zoom in en gráfica RDS"""
        xlim = self.ax_rds.get_xlim()
        ylim = self.ax_rds.get_ylim()
        self.ax_rds.set_xlim(xlim[0] * 0.8, xlim[1] * 0.8)
        self.ax_rds.set_ylim(ylim[0] * 0.8, ylim[1] * 0.8)
        self.canvas_rds.draw()

    def zoom_out_rds(self):
        """Zoom out en gráfica RDS"""
        xlim = self.ax_rds.get_xlim()
        ylim = self.ax_rds.get_ylim()
        self.ax_rds.set_xlim(xlim[0] * 1.2, xlim[1] * 1.2)
        self.ax_rds.set_ylim(ylim[0] * 1.2, ylim[1] * 1.2)
        self.canvas_rds.draw()

    def autoscale_rds(self):
        """Autoajusta gráfica RDS"""
        self.ax_rds.relim()
        self.ax_rds.autoscale_view()
        self.canvas_rds.draw()

    def clear_rds_plot(self):
        """Limpia gráfica RDS"""
        self.ax_rds.clear()
        self.ax_rds.set_xlabel("Longitud de onda (nm)", fontsize=10)
        self.ax_rds.set_ylabel("ΔR/R", fontsize=10)
        self.ax_rds.grid(True, alpha=0.3)
        self.canvas_rds.draw()
        self.rds_data = []

    def __del__(self):
        if self.optical_system:
            self.optical_system.close()

# Ejecutar la aplicación
if __name__ == "__main__":
    try:
        logger.info("Iniciando aplicación de caracterización óptica")
        app = Reflectancia()
        app.mainloop()
    except Exception as e:
        logger.critical(f"Error crítico en la aplicación: {e}")
        messagebox.showerror("Error Crítico", f"La aplicación encontró un error: {e}")
