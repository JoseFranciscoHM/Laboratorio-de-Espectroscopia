import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import pyvisa
import threading
import time
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import json
import os

class MultimetroGPIBApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Control de Multímetro GPIB")
        self.geometry("900x700")
        self.configure(bg="#f0f0f0")
        
        # Variables de estado
        self.rm = None
        self.multimetro = None
        self.connected = False
        self.measuring = False
        self.measurement_data = []
        
        # Configurar estilo
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.configure_styles()
        
        # Cargar configuración
        self.load_config()
        
        self.create_widgets()
        
        # Intentar inicializar pyvisa
        try:
            self.rm = pyvisa.ResourceManager()
            self.update_devices_list()
        except Exception as e:
            self.log_message(f"Error al inicializar PyVISA: {str(e)}")
        
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def configure_styles(self):
        self.style.configure('TFrame', background="#f0f0f0")
        self.style.configure('TLabel', background="#f0f0f0", foreground="#333333", font=('Arial', 9))
        self.style.configure('TButton', font=('Arial', 9))
        self.style.configure('TEntry', font=('Arial', 9))
        self.style.configure('TCombobox', font=('Arial', 9))
        self.style.configure('TLabelframe', background="#f0f0f0", foreground="#2c3e50")
        
        # Estilos para botones de estado
        self.style.configure('Connected.TButton', background="#4CAF50", foreground="white")
        self.style.configure('Disconnected.TButton', background="#F44336", foreground="white")
        self.style.configure('Measuring.TButton', background="#2196F3", foreground="white")

    def load_config(self):
        self.config_data = {
            "gpib_address": "GPIB0::26::INSTR",
            "command": "MEAS:VOLT:DC?",
            "timeout": "5",
            "readings_count": "10",
            "delay": "1.0"
        }
        
        try:
            if os.path.exists("config_gpib.json"):
                with open("config_gpib.json", "r") as f:
                    self.config_data = json.load(f)
        except:
            pass

    def save_config(self):
        try:
            with open("config_gpib.json", "w") as f:
                json.dump(self.config_data, f)
        except:
            pass

    def create_widgets(self):
        # Frame de configuración
        config_frame = ttk.LabelFrame(self, text="Configuración del Multímetro", padding="10")
        config_frame.pack(fill="x", padx=10, pady=5)

        # Dirección GPIB
        ttk.Label(config_frame, text="Dirección GPIB:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.gpib_address = ttk.Entry(config_frame, width=20)
        self.gpib_address.insert(0, self.config_data["gpib_address"])
        self.gpib_address.grid(row=0, column=1, padx=5, pady=5)
        
        # Lista de dispositivos
        ttk.Button(config_frame, text="Actualizar dispositivos", command=self.update_devices_list).grid(row=0, column=2, padx=5, pady=5)
        
        self.devices_list = ttk.Combobox(config_frame, width=30, state="readonly")
        self.devices_list.grid(row=0, column=3, padx=5, pady=5)
        self.devices_list.bind('<<ComboboxSelected>>', self.on_device_selected)

        # Comando
        ttk.Label(config_frame, text="Comando:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.command_var = tk.StringVar(value=self.config_data["command"])
        self.command = ttk.Combobox(config_frame, width=20, textvariable=self.command_var)
        self.command['values'] = (
            "MEAS:VOLT:DC?", "MEAS:VOLT:AC?", "MEAS:CURR:DC?", "MEAS:CURR:AC?",
            "MEAS:RES?", "MEAS:FRES?", "MEAS:FREQ?", "MEAS:PER?", 
            "MEAS:CONT?", "MEAS:DIO?", "MEAS:TEMP?"
        )
        self.command.grid(row=1, column=1, padx=5, pady=5)

        # Timeout
        ttk.Label(config_frame, text="Timeout (s):").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.timeout = ttk.Entry(config_frame, width=10)
        self.timeout.insert(0, self.config_data["timeout"])
        self.timeout.grid(row=1, column=3, padx=5, pady=5)

        # Configuración de medición continua
        ttk.Label(config_frame, text="Nº de lecturas:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.readings_count = ttk.Entry(config_frame, width=10)
        self.readings_count.insert(0, self.config_data["readings_count"])
        self.readings_count.grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(config_frame, text="Intervalo (s):").grid(row=2, column=2, padx=5, pady=5, sticky="e")
        self.delay = ttk.Entry(config_frame, width=10)
        self.delay.insert(0, self.config_data["delay"])
        self.delay.grid(row=2, column=3, padx=5, pady=5)

        # Botones de control
        button_frame = ttk.Frame(config_frame)
        button_frame.grid(row=3, column=0, columnspan=4, pady=10)

        self.connect_button = ttk.Button(button_frame, text="Conectar", command=self.toggle_connection)
        self.connect_button.pack(side=tk.LEFT, padx=5)

        self.single_measure_button = ttk.Button(button_frame, text="Medición Única", 
                                               command=self.single_measurement, state=tk.DISABLED)
        self.single_measure_button.pack(side=tk.LEFT, padx=5)

        self.continuous_button = ttk.Button(button_frame, text="Medición Continua", 
                                           command=self.toggle_continuous_measurement, state=tk.DISABLED)
        self.continuous_button.pack(side=tk.LEFT, padx=5)

        self.trigger_button = ttk.Button(button_frame, text="Forzar Trigger", 
                                        command=self.force_trigger, state=tk.DISABLED)
        self.trigger_button.pack(side=tk.LEFT, padx=5)

        self.clear_button = ttk.Button(button_frame, text="Clear Device", 
                                      command=self.clear_device, state=tk.DISABLED)
        self.clear_button.pack(side=tk.LEFT, padx=5)

        self.status_button = ttk.Button(button_frame, text="Leer Estado", 
                                       command=self.read_status, state=tk.DISABLED)
        self.status_button.pack(side=tk.LEFT, padx=5)

        # Frame para visualización de datos
        data_frame = ttk.LabelFrame(self, text="Datos de Medición", padding="10")
        data_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Gráfica
        graph_frame = ttk.Frame(data_frame)
        graph_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        self.figure, self.ax = plt.subplots(figsize=(8, 4))
        self.ax.set_xlabel("Tiempo")
        self.ax.set_ylabel("Valor")
        self.ax.grid(True, alpha=0.3)
        self.canvas = FigureCanvasTkAgg(self.figure, master=graph_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        # Consola de mensajes
        console_frame = ttk.Frame(data_frame)
        console_frame.pack(fill="both", expand=True)
        
        self.console = scrolledtext.ScrolledText(console_frame, height=8, state=tk.DISABLED)
        self.console.pack(fill="both", expand=True)

        # Barra de estado
        status_frame = ttk.Frame(self)
        status_frame.pack(fill="x", padx=10, pady=5)
        
        self.status_var = tk.StringVar(value="Desconectado")
        status_label = ttk.Label(status_frame, textvariable=self.status_var, foreground="#2c3e50")
        status_label.pack(side=tk.LEFT)

        self.connection_status = ttk.Label(status_frame, text="Desconectado", foreground="red")
        self.connection_status.pack(side=tk.RIGHT)

    def update_devices_list(self):
        try:
            devices = self.rm.list_resources()
            self.devices_list['values'] = devices
            if devices:
                self.devices_list.set(devices[0])
            self.log_message(f"Dispositivos encontrados: {devices}")
        except Exception as e:
            self.log_message(f"Error al listar dispositivos: {str(e)}")

    def on_device_selected(self, event):
        self.gpib_address.delete(0, tk.END)
        self.gpib_address.insert(0, self.devices_list.get())

    def toggle_connection(self):
        if self.connected:
            self.disconnect()
        else:
            self.connect()

    def connect(self):
        try:
            address = self.gpib_address.get().strip()
            timeout = int(self.timeout.get())
            
            self.multimetro = self.rm.open_resource(address)
            self.multimetro.timeout = timeout * 1000  # Convertir a ms
            
            # Guardar configuración
            self.config_data = {
                "gpib_address": address,
                "command": self.command_var.get(),
                "timeout": self.timeout.get(),
                "readings_count": self.readings_count.get(),
                "delay": self.delay.get()
            }
            self.save_config()
            
            self.connected = True
            self.update_ui_connection_state()
            self.log_message(f"Conectado a {address}")
            
        except Exception as e:
            self.log_message(f"Error de conexión: {str(e)}")
            messagebox.showerror("Error de conexión", f"No se pudo conectar al dispositivo:\n{e}")

    def disconnect(self):
        try:
            if self.multimetro:
                self.multimetro.close()
            self.connected = False
            self.measuring = False
            self.update_ui_connection_state()
            self.log_message("Desconectado")
        except Exception as e:
            self.log_message(f"Error al desconectar: {str(e)}")

    def update_ui_connection_state(self):
        if self.connected:
            self.connect_button.config(text="Desconectar", style='Connected.TButton')
            self.connection_status.config(text="Conectado", foreground="green")
            self.status_var.set("Conectado")
            
            # Habilitar botones
            self.single_measure_button.config(state=tk.NORMAL)
            self.continuous_button.config(state=tk.NORMAL)
            self.trigger_button.config(state=tk.NORMAL)
            self.clear_button.config(state=tk.NORMAL)
            self.status_button.config(state=tk.NORMAL)
        else:
            self.connect_button.config(text="Conectar", style='TButton')
            self.connection_status.config(text="Desconectado", foreground="red")
            self.status_var.set("Desconectado")
            
            # Deshabilitar botones
            self.single_measure_button.config(state=tk.DISABLED)
            self.continuous_button.config(state=tk.DISABLED)
            self.trigger_button.config(state=tk.DISABLED)
            self.clear_button.config(state=tk.DISABLED)
            self.status_button.config(state=tk.DISABLED)
            
            # Detener medición continua si está activa
            if self.measuring:
                self.continuous_button.config(text="Medición Continua", style='TButton')

    def single_measurement(self):
        try:
            command = self.command_var.get()
            self.multimetro.write(command)
            response = self.multimetro.read()
            
            timestamp = datetime.now().strftime("%H:%M:%S")
            self.log_message(f"[{timestamp}] {response.strip()}")
            
            # Intentar convertir a número para la gráfica
            try:
                value = float(response)
                self.measurement_data.append((time.time(), value))
                self.update_plot()
            except ValueError:
                pass
                
        except Exception as e:
            self.log_message(f"Error en medición: {str(e)}")

    def toggle_continuous_measurement(self):
        if self.measuring:
            self.measuring = False
            self.continuous_button.config(text="Medición Continua", style='TButton')
            self.log_message("Medición continua detenida")
        else:
            try:
                readings = int(self.readings_count.get())
                delay = float(self.delay.get())
                
                self.measuring = True
                self.continuous_button.config(text="Detener Medición", style='Measuring.TButton')
                
                # Iniciar medición continua en un hilo separado
                threading.Thread(
                    target=self.continuous_measurement, 
                    args=(readings, delay), 
                    daemon=True
                ).start()
                
            except ValueError:
                messagebox.showerror("Error", "Por favor, ingrese valores numéricos válidos")

    def continuous_measurement(self, readings, delay):
        count = 0
        command = self.command_var.get()
        
        while self.measuring and (readings == 0 or count < readings):
            try:
                self.multimetro.write(command)
                response = self.multimetro.read()
                
                timestamp = datetime.now().strftime("%H:%M:%S")
                self.log_message(f"[{timestamp}] {response.strip()}")
                
                # Intentar convertir a número para la gráfica
                try:
                    value = float(response)
                    self.measurement_data.append((time.time(), value))
                    self.update_plot()
                except ValueError:
                    pass
                
                count += 1
                time.sleep(delay)
                
            except Exception as e:
                self.log_message(f"Error en medición continua: {str(e)}")
                break
        
        self.measuring = False
        self.after(0, lambda: self.continuous_button.config(text="Medición Continua", style='TButton'))

    def force_trigger(self):
        try:
            self.multimetro.assert_trigger()
            self.log_message("Trigger forzado")
        except Exception as e:
            self.log_message(f"Error al forzar trigger: {str(e)}")

    def clear_device(self):
        try:
            self.multimetro.clear()
            self.log_message("Dispositivo limpiado")
        except Exception as e:
            self.log_message(f"Error al limpiar dispositivo: {str(e)}")

    def read_status(self):
        try:
            status = self.multimetro.stb
            self.log_message(f"Estado del dispositivo: {status} (0x{status:02X})")
        except Exception as e:
            self.log_message(f"Error al leer estado: {str(e)}")

    def update_plot(self):
        if not self.measurement_data:
            return
            
        self.ax.clear()
        
        # Extraer tiempos y valores
        times = [t for t, v in self.measurement_data]
        values = [v for t, v in self.measurement_data]
        
        # Convertir tiempos a segundos relativos
        start_time = times[0]
        rel_times = [t - start_time for t in times]
        
        self.ax.plot(rel_times, values, 'b-', marker='o', markersize=3)
        self.ax.set_xlabel("Tiempo (s)")
        self.ax.set_ylabel("Valor")
        self.ax.grid(True, alpha=0.3)
        
        # Ajustar automáticamente los límites
        if len(values) > 1:
            value_range = max(values) - min(values)
            if value_range > 0:
                self.ax.set_ylim(min(values) - 0.1 * value_range, max(values) + 0.1 * value_range)
        
        self.canvas.draw()

    def log_message(self, message):
        self.console.config(state=tk.NORMAL)
        self.console.insert(tk.END, message + "\n")
        self.console.see(tk.END)
        self.console.config(state=tk.DISABLED)

    def on_closing(self):
        if self.measuring:
            self.measuring = False
            time.sleep(0.5)  # Esperar a que el hilo se detenga
        
        self.disconnect()
        self.destroy()

if __name__ == "__main__":
    app = MultimetroGPIBApp()
    app.mainloop()
