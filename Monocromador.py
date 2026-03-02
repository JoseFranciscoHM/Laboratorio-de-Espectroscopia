import tkinter as tk
from tkinter import ttk, messagebox
import serial
import serial.tools.list_ports

class FormMonocromador(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Control de Monocromador")
        self.geometry("500x350")
        self.configure(bg="#f0f0f0")  # Fondo claro y neutro
        
        # Configurar estilo
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.configure_styles()
        
        self.crear_widgets()
        
    def configure_styles(self):
        # Configurar estilos para los widgets con colores claros
        self.style.configure('TFrame', background="#f0f0f0")
        self.style.configure('TLabel', background="#f0f0f0", foreground="#333333", font=('Arial', 10))
        self.style.configure('TButton', font=('Arial', 10), background="#4CAF50", foreground="white")
        self.style.configure('TEntry', font=('Arial', 10), fieldbackground="white", foreground="#333333")
        self.style.configure('TCombobox', font=('Arial', 10), fieldbackground="white", foreground="#333333")
        self.style.map('TButton', 
                      background=[('active', '#45a049')],
                      foreground=[('active', 'white')])
        
    def crear_widgets(self):
        # Marco principal
        main_frame = ttk.Frame(self, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        title_label = ttk.Label(main_frame, text="Control de Monocromador", 
                               font=('Arial', 16, 'bold'), foreground="#2c3e50")
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Etiqueta y combobox para el puerto COM
        ttk.Label(main_frame, text="Puerto COM:").grid(row=1, column=0, padx=5, pady=10, sticky="e")
        
        # Frame para el combobox y botón de actualización
        port_frame = ttk.Frame(main_frame)
        port_frame.grid(row=1, column=1, padx=5, pady=10, sticky="ew")
        
        self.combo_port = ttk.Combobox(port_frame, width=20)
        self.combo_port.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        refresh_btn = ttk.Button(port_frame, text="↻", width=3, command=self.actualizar_puertos)
        refresh_btn.pack(side=tk.RIGHT, padx=(5, 0))
        
        # Etiqueta y entrada para la longitud de onda
        ttk.Label(main_frame, text="Longitud de onda (nm):").grid(row=2, column=0, padx=5, pady=10, sticky="e")
        self.entry_wavelength = ttk.Entry(main_frame, width=25)
        self.entry_wavelength.grid(row=2, column=1, padx=5, pady=10, sticky="w")
        
        # Botón para mover el monocromador
        self.btn_mover = ttk.Button(main_frame, text="Mover Monocromador", command=self.mover_monocromador)
        self.btn_mover.grid(row=3, column=0, columnspan=2, pady=20)
        
        # Área de estado
        ttk.Label(main_frame, text="Estado:").grid(row=4, column=0, padx=5, pady=5, sticky="e")
        self.status_var = tk.StringVar(value="Desconectado")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, foreground="#e74c3c")
        status_label.grid(row=4, column=1, padx=5, pady=5, sticky="w")
        
        # Consola de mensajes
        ttk.Label(main_frame, text="Mensajes:").grid(row=5, column=0, padx=5, pady=(20, 5), sticky="ne")
        
        self.text_console = tk.Text(main_frame, height=8, width=50, 
                                   bg="white", fg="#333333",  # Fondo blanco, texto gris oscuro
                                   font=('Consolas', 9), 
                                   relief=tk.SUNKEN, bd=1)
        self.text_console.grid(row=5, column=1, padx=5, pady=(20, 5), sticky="nsew")
        
        # Barra de desplazamiento para la consola
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.text_console.yview)
        scrollbar.grid(row=5, column=2, padx=(0, 5), pady=(20, 5), sticky="ns")
        self.text_console.configure(yscrollcommand=scrollbar.set)
        
        # Configurar expansión de filas y columnas
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)
        
        # Inicializar lista de puertos
        self.actualizar_puertos()
        
        # Establecer el foco en la entrada de longitud de onda
        self.entry_wavelength.focus()
        
        # Vincular la tecla Enter al botón de mover
        self.bind('<Return>', lambda event: self.mover_monocromador())
        
    def actualizar_puertos(self):
        """Actualiza la lista de puertos COM disponibles"""
        ports = [port.device for port in serial.tools.list_ports.comports()]
        self.combo_port['values'] = ports
        if ports:
            self.combo_port.set(ports[0])
        else:
            self.combo_port.set('')
            self.mostrar_mensaje("No se encontraron puertos COM disponibles")
            
    def mover_monocromador(self):
        """Envía el comando para mover el monocromador a la longitud de onda especificada"""
        port = self.combo_port.get().strip()
        if not port:
            messagebox.showwarning("Puerto no seleccionado", "Por favor, seleccione un puerto COM.")
            return
            
        try:
            wavelength = float(self.entry_wavelength.get())
            if wavelength < 0:
                raise ValueError("La longitud de onda no puede ser negativa")
                
            self.mostrar_mensaje(f"Enviando comando: mover a {wavelength} nm")
            self.btn_mover.config(state=tk.DISABLED)
            self.status_var.set("Conectando...")
            self.update()
            
            self.enviar_comando(port, wavelength)
            
        except ValueError as e:
            messagebox.showwarning("Entrada inválida", f"Por favor, ingrese un número válido para la longitud de onda.\n{str(e)}")
            self.btn_mover.config(state=tk.NORMAL)
            
    def enviar_comando(self, port, wavelength):
        """Envía el comando al monocromador a través del puerto serial"""
        try:
            with serial.Serial(port, baudrate=9600, timeout=2) as ser:
                self.status_var.set("Conectado")
                self.mostrar_mensaje(f"Conectado al puerto {port}")
                
                command = f'{wavelength} GOTO\r\n'
                ser.write(command.encode('utf-8'))
                self.mostrar_mensaje(f"Comando enviado: {command.strip()}")
                
                response = ser.readline().decode('utf-8').strip()
                self.mostrar_mensaje(f"Respuesta: {response}")
                
                messagebox.showinfo("Éxito", f"Monocromador movido a {wavelength} nm.")
                self.status_var.set("Comando ejecutado")
                
        except serial.SerialException as e:
            error_msg = f"No se pudo conectar al puerto {port}:\n{str(e)}"
            self.mostrar_mensaje(f"Error: {error_msg}")
            messagebox.showerror("Error de conexión", error_msg)
            self.status_var.set("Error de conexión")
            
        except Exception as e:
            error_msg = f"Ocurrió un error inesperado:\n{str(e)}"
            self.mostrar_mensaje(f"Error: {error_msg}")
            messagebox.showerror("Error", error_msg)
            self.status_var.set("Error")
            
        finally:
            self.btn_mover.config(state=tk.NORMAL)
            
    def mostrar_mensaje(self, mensaje):
        """Añade un mensaje a la consola de texto"""
        self.text_console.insert(tk.END, f"> {mensaje}\n")
        self.text_console.see(tk.END)  # Auto-desplazamiento al final

if __name__ == "__main__":
    app = FormMonocromador()
    app.mainloop()
