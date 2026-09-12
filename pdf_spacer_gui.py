import sys
import customtkinter as ctk
import config
import subprocess
import os
from customtkinter import filedialog
from PIL import Image

NOMBRE_MOTOR = "pdf_spacer_engine.exe"

# 1. Configuración del tema global (Dark, Light o System)
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")  # Opciones: "blue", "green", "dark-blue"

# 2. Creamos la clase de nuestra App heredando de ctk.CTk
class PDFSpacerGUI(ctk.CTk):
    def __init__(self, config_file):
        super().__init__()

        # Configurar la ventana
        self.title("PDF Spacer+")
        self.geometry("500x635")
        self.resizable(False, False)  # Evita cambiar el tamaño si no queremos

        # --- CREACIÓN DE WIDGETS ---
        
        img_logo = ctk.CTkImage(
        light_image=Image.open("icono.png"), # Imagen a usar en modo claro
        dark_image=Image.open("icono.png"),  # Imagen a usar en modo oscuro (o la misma)
        size=(200, 200)                        # Ancho y alto en píxeles
        )

        self.label_logo = ctk.CTkLabel(self, image=img_logo, text="")
        self.label_logo.pack(pady=(40,25), side="top")

        """
        # Un botón interactivo (Button)
        self.boton = ctk.CTkButton(
            self, 
            text="Procesar PDF", 
            command=self.accion_boton  # Función que se ejecuta al hacer clic
        )
        self.boton.pack(pady=10)
        """

        
        self.label_input = ctk.CTkLabel(self, text="Input folder:", font=("Microsoft Sans Seriff", 14, "bold"))
        self.label_input.pack(anchor="w", padx=20, pady=(0, 1))

        # Contenedor horizontal para Input
        self.frame_input = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_input.pack(fill="x", padx=20, pady=5)

        # Campo de texto (Entry)
        self.entry_input = ctk.CTkEntry(self.frame_input, placeholder_text=config_file["input_folder"], corner_radius=0)
        self.entry_input.pack(side="left", fill="x", expand=True, padx=(0, 10))

        # Botón Examinar
        self.btn_browse_input = ctk.CTkButton(
            self.frame_input, 
            text="Browse...", 
            width=100, 
            command=self.buscar_carpeta_input,
            corner_radius=0,
        )
        self.btn_browse_input.pack(side="right")


        # --- SECCIÓN CARPETA OUTPUT ---
        self.label_output = ctk.CTkLabel(self, text="Output folder:", font=("Microsoft Sans Seriff", 14, "bold"))
        self.label_output.pack(anchor="w", padx=20, pady=(15, 1))

        # Contenedor horizontal para Output
        self.frame_output = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_output.pack(fill="x", padx=20, pady=5)

        # Campo de texto (Entry)
        self.entry_output = ctk.CTkEntry(self.frame_output, placeholder_text=config_file["output_folder"], corner_radius=0)
        self.entry_output.pack(side="left", fill="x", expand=True, padx=(0, 10))

        # Botón Examinar
        self.btn_browse_output = ctk.CTkButton(
            self.frame_output, 
            text="Browse...", 
            width=100, 
            command=self.buscar_carpeta_output,
            corner_radius=0
        )
        self.btn_browse_output.pack(side="right")


        frame_horizontal_continous = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0) # transparent para que no se vea el cuadro
        frame_horizontal_continous.pack(pady=20, padx=20, fill="x") # fill="x" hace que ocupe todo el ancho disponible
        
        # Un interruptor ON/OFF (Switch)
        self.switch_escaneo = ctk.CTkSwitch(
            frame_horizontal_continous, 
            text="Continous mode",
            command=self.continous_mode_switch
        )

        value = self.switch_escaneo.get()
        if (value != config_file["continous_mode"]) and config_file["continous_mode"] == 1:
            self.switch_escaneo.select()
        self.switch_escaneo.pack(side="left", padx=10, expand=True)

        btn_stop = ctk.CTkButton(frame_horizontal_continous, text="Stop", corner_radius=0, command=self.stop_process)
        btn_stop.pack(side="left", padx=10, expand=True)

        # Un Slider (Deslizador para el margen)
        self.slider = ctk.CTkSlider(
            self, 
            from_=1.5, 
            to=3.0, 
            number_of_steps=15,
            command=self.accion_slider,
            corner_radius=0
        )
        self.slider.set(config_file.get("multiplier", 2.0))  # Valor inicial por defecto
        self.slider.pack(pady=10)

        # Etiqueta para mostrar el valor del Slider
        slider_text = f"Multiplier: {self.slider.get()}x"
        self.label_slider_val = ctk.CTkLabel(self, text=slider_text)
        self.label_slider_val.pack(pady=5)

        
        frame_horizontal = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0) # transparent para que no se vea el cuadro
        frame_horizontal.pack(pady=20, padx=20, fill="x") # fill="x" hace que ocupe todo el ancho disponible

        # 2. Elemento 1 (Izquierda)
        launch_btn = ctk.CTkButton(frame_horizontal, text="Launch", corner_radius=0, command=self.launch_main)
        launch_btn.pack(side="left", padx=10, expand=True)

        # 3. Elemento 2 (Derecha / Al mismo nivel)
        save_btn = ctk.CTkButton(frame_horizontal, text="Save", fg_color="red", corner_radius=0, command=self.save_config_ui)
        save_btn.pack(side="left", padx=10, expand=True)


    # --- FUNCIONES DE BÚSQUEDA ---

    def buscar_carpeta_input(self):
        # Abre la ventana nativa del SO para elegir directorio
        carpeta_seleccionada = filedialog.askdirectory(title="Seleccionar Carpeta de Entrada")
        
        # Si el usuario eligió una carpeta y no canceló la ventana
        if carpeta_seleccionada:
            self.entry_input.delete(0, "end")  # Limpiamos la caja
            self.entry_input.insert(0, carpeta_seleccionada)  # Escribimos la ruta
            

    def buscar_carpeta_output(self):
        carpeta_seleccionada = filedialog.askdirectory(title="Seleccionar Carpeta de Salida")
        if carpeta_seleccionada:
            self.entry_output.delete(0, "end")
            self.entry_output.insert(0, carpeta_seleccionada)
            

    # --- FUNCIONES / EVENTOS DE LOS WIDGETS ---


    def continous_mode_switch(self):
        estado = "ON" if self.switch_escaneo.get() == 1 else "OFF"
        print(f" El escáner ahora está: {estado}")

        
    def accion_slider(self, value):
        # Actualizamos la etiqueta con el valor flotante del slider formateado
        self.label_slider_val.configure(text=f"Multiplicador: {value:.2f}x")


    def save_config_ui(self):

        cfg = config.load_config()

        new_input_folder = self.entry_input.get()
        if new_input_folder != '':
            cfg["input_folder"] = new_input_folder

        new_output_folder = self.entry_output.get()
        if new_output_folder != '':
            cfg["output_folder"] = new_output_folder

        cfg["continous_mode"] = self.switch_escaneo.get()
        cfg["multiplier"] = self.slider.get()
        
        config.save_config(cfg)
    

    def es_motor_ejecutandose(self) -> bool:
        """Verifica si motor.exe está corriendo en los procesos de Windows."""
        try:
            salida = subprocess.check_output(
                ["tasklist", "/FI", f"IMAGENAME eq {NOMBRE_MOTOR}"],
                creationflags=subprocess.CREATE_NO_WINDOW
            ).decode("utf-8", errors="ignore")
            
            return NOMBRE_MOTOR.lower() in salida.lower()
        except Exception:
            return False

    def stop_process(self):
        # Forzamos el cierre del proceso en Windows por su nombre de ejecutable
        try:
            subprocess.run(
                ["taskkill", "/F", "/IM", NOMBRE_MOTOR], 
                creationflags=subprocess.CREATE_NO_WINDOW,
                stdout=subprocess.DEVNULL, 
                stderr=subprocess.DEVNULL
            )
            print(f"proceso {NOMBRE_MOTOR} eliminado")
        except Exception as e:
            print(f"Error al intentar detener el motor: {e}")

    def launch_main(self):

        # Comprobamos si el motor ya está corriendo en el sistema
        if self.es_motor_ejecutandose():
            print("El motor ya está en ejecución.")
            return

        # Si el ejecutable existe (cuando compiles a .exe)
        if os.path.exists(NOMBRE_MOTOR):
            # creationflags=subprocess.CREATE_NO_WINDOW evita que abra una consola negra fea
            self.proceso_motor = subprocess.Popen(
                [NOMBRE_MOTOR], 
                creationflags=subprocess.CREATE_NO_WINDOW
            )
        else:
            # Modo desarrollo (mientras estás en VS Code antes de compilar)
            self.proceso_motor = subprocess.Popen([sys.executable, "pdf_spacer_engine.py"])

# 3. Punto de entrada para ejecutar la ventana
if __name__ == "__main__":
    user_cfg = config.load_config()
    app = PDFSpacerGUI(user_cfg)
    app.mainloop()  # Este bucle mantiene la ventana abierta escuchando clics