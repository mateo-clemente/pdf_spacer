import pymupdf
import time
import os
import threading
import config
import pythoncom
import win32com.client
import tempfile
import logging
from logging.handlers import RotatingFileHandler

from queue import Queue
from watchdog.events import FileSystemEvent, FileSystemEventHandler
from watchdog.observers import Observer
from pathlib import Path

# ==========================================
# CONFIGURACIÓN DEL SISTEMA DE LOGS
# ==========================================
def setup_logging(log_dir="logs"):
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, "backend_app.log")

    # Crear el logger principal
    logger = logging.getLogger()
    logger.setLevel(logging.INFO) # Cambiar a logging.DEBUG si necesitas ver más detalle

    # Formato del log: Fecha | Nivel | Hilo | Mensaje
    formato = logging.Formatter(
        '%(asctime)s | %(levelname)-8s | %(threadName)-15s | %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    # 1. Archivo siempre activo
    file_handler = RotatingFileHandler(log_file, maxBytes=5*1024*1024, backupCount=3, encoding='utf-8')
    file_handler.setFormatter(formato)
    logger.addHandler(file_handler)

    # 2. Consola solo si existe una ventana de comandos disponible
    if sys.stdout is not None:
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formato)
        logger.addHandler(console_handler)

# ==========================================
# FUNCIONES DEL PROGRAMA
# ==========================================

def wait_for_file_ready(filepath, timeout=60):
    """
    Espera a que el archivo deje de estar bloqueado Y su tamaño sea estable
    por al menos 3 comprobaciones consecutivas (útil para descargas lentas).
    """
    start_time = time.time()
    last_size = -1
    stable_seconds = 0

    while True:
        try:
            if not os.path.exists(filepath):
                return False

            # 1. Forzar un bloqueo de lectura/escritura. 
            # Si el navegador está descargando, esto fallará inmediatamente.
            with open(filepath, 'r+b') as f:
                pass 

            # 2. Comprobar que el tamaño ha dejado de crecer
            current_size = os.path.getsize(filepath)
            
            if current_size == last_size and current_size > 0:
                stable_seconds += 1
            else:
                stable_seconds = 0  # Reiniciamos el contador si el tamaño cambió
            
            last_size = current_size

            # Si el archivo lleva 3 segundos exactos sin cambiar ni estar bloqueado, está listo
            if stable_seconds >= 3:
                return True

        except (OSError, IOError, PermissionError):
            # El archivo sigue bloqueado por el navegador/sistema
            stable_seconds = 0 

        if time.time() - start_time > timeout:
            logging.warning(f"⏳ Timeout de {timeout}s alcanzado esperando a: {os.path.basename(filepath)}")
            return False
        
        # Esperamos 1 segundo antes de la siguiente comprobación
        time.sleep(1)




def convert_ppt_to_pdf(ppt_path: str) -> str:
    """
    Convierte un archivo .ppt o .pptx a .pdf.
    """
    ppt_path = os.path.abspath(ppt_path)
    filename = os.path.basename(ppt_path)
    base_name = os.path.splitext(filename)[0]
    
    temp_dir = tempfile.gettempdir()
    pdf_path = os.path.join(temp_dir, f"{base_name}_temp_conv.pdf")
    
    pythoncom.CoInitialize()
    powerpoint = None
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        presentation = powerpoint.Presentations.Open(
            ppt_path, 
            WithWindow=False, 
            ReadOnly=True
        )
        presentation.SaveAs(pdf_path, 32)
        presentation.Close()
        logging.info(f"PPT convertido con éxito a PDF temporal: {pdf_path}")
        return pdf_path
    except Exception as e:
        # exc_info=True guarda el stacktrace completo en el log
        logging.error(f"Error al convertir PPT {filename}: {e}", exc_info=True)
        raise e
    finally:
        if powerpoint:
            powerpoint.Quit()
        pythoncom.CoUninitialize()

def folder_sweep(input_folder: str, queue: Queue):
    logging.info(f"Iniciando barrido de carpeta: {input_folder}")
    try:
        path_folder = Path(input_folder)
        valid_extensions = ("*.pdf", "*.ppt", "*.pptx")
        for ext in valid_extensions:
            for file in path_folder.glob(ext):
                ruta_completa = str(file.resolve())
                if is_valid_file(ruta_completa):
                    queue.put(ruta_completa)
    except Exception as e:
        logging.error(f"Error durante el barrido de carpeta: {e}", exc_info=True)

def pdf_blank_spacer(input_pdf, output_pdf, multiplier=2.0):
    for page in range(len(input_pdf)):
        original_page = input_pdf[page]
        rotation = original_page.rotation
        rect = original_page.rect
        original_width = rect.width
        original_height = rect.height

        new_page = output_pdf.new_page(width=original_width*multiplier, height=original_height)
        new_page.set_rotation(rotation)
        rect_destination = pymupdf.Rect(0,0, original_width, original_height)
        new_page.show_pdf_page(rect_destination, input_pdf, page)
    return

def pdf_processor(input_pdf_path, output_pdf_path, pdf_multiplier):
    input_pdf = pymupdf.open(input_pdf_path)
    
    # RED DE SEGURIDAD: Comprobar que el PDF tiene contenido real
    if len(input_pdf) == 0:
        input_pdf.close()
        raise ValueError("El documento PDF tiene 0 páginas o está corrupto. Abortando.")

    output_pdf = pymupdf.open()
    pdf_blank_spacer(input_pdf=input_pdf, output_pdf=output_pdf, multiplier=pdf_multiplier)

    output_pdf.save(output_pdf_path)
    output_pdf.close()
    input_pdf.close()
    
    # Nota: He quitado el os.remove() de aquí porque ya lo haces en el queue_worker. 
    # Es mejor tener la responsabilidad de borrar centralizada en un solo sitio.
    return

def queue_worker(queue: Queue, output, pdf_multiplier):
    while True:
        item = queue.get() 

        if not os.path.exists(item):
            queue.task_done()
            continue
        
        while not wait_for_file_ready(item):
            logging.info(f"El archivo {os.path.basename(item)} no está listo para ser procesado.")


        filename = os.path.basename(item)
        file_ext = os.path.splitext(filename)[1].lower()
        
        base_name = os.path.splitext(filename)[0]
        output_pdf_path = os.path.join(output, f"extended_{base_name}.pdf")
        temp_pdf_to_clean = None
        
        try:
            logging.info(f"Procesando: {filename}")
            
            if file_ext in ['.ppt', '.pptx']:
                temp_pdf_to_clean = convert_ppt_to_pdf(item)
                pdf_to_process = temp_pdf_to_clean
            elif file_ext == '.pdf':
                pdf_to_process = item
            else:
                logging.warning(f"Archivo ignorado (formato no soportado): {filename}")
                continue

            pdf_processor(
                input_pdf_path=pdf_to_process, 
                output_pdf_path=output_pdf_path, 
                pdf_multiplier=pdf_multiplier
            )
            
            if os.path.exists(item):
                logging.info(f"Procesado y eliminado original: {filename}")
                os.remove(item)
                
            

        except Exception as e:
            logging.error(f"Error al procesar {filename}: {e}", exc_info=True)
            
        finally:
            if temp_pdf_to_clean and os.path.exists(temp_pdf_to_clean):
                try:
                    os.remove(temp_pdf_to_clean)
                except Exception:
                    pass
            queue.task_done()

def is_valid_file(file_path: str) -> bool:
    filename = os.path.basename(file_path)
    # Ignorar archivos temporales explícitos
    if filename.startswith("~$") or "_temp_conv.pdf" in filename:
        return False
    if filename.endswith(('.crdownload', '.part', '.tmp')):
        return False
        
    return filename.lower().endswith(('.pdf', '.ppt', '.pptx'))

class Enqueuer(FileSystemEventHandler):
    def __init__(self, queue: Queue):
        self.paths_queue = queue

    def process_event(self, event: FileSystemEvent):
        if not event.is_directory and is_valid_file(event.src_path):
            file_path = os.path.abspath(event.src_path)
            self.paths_queue.put(file_path)

    def on_created(self, event: FileSystemEvent) -> None:
        logging.info(f"Enqueuing {os.path.abspath(event.src_path)} from on_created")
        self.process_event(event)

    def on_modified(self, event: FileSystemEvent) -> None:
        logging.info(f"Enqueuing {os.path.abspath(event.src_path)} from on_modified")
        self.process_event(event)

    def on_moved(self, event: FileSystemEvent) -> None:
        if not event.is_directory and is_valid_file(event.dest_path):
            logging.info(f"Enqueuing {os.path.abspath(event.dest_path)} from on_moved")
            file_path = os.path.abspath(event.dest_path)
            self.paths_queue.put(file_path)

if __name__ == "__main__":
    
    # 0. Inicializar Logs antes que nada
    setup_logging()
    logging.info("Iniciando aplicación...")

    cfg = config.load_config()

    DIR_INPUT = os.path.abspath(cfg.get("input_folder", "input"))
    DIR_OUTPUT = os.path.abspath(cfg.get("output_folder", "output"))
    INTERVALO_SCANNER_SEGUNDOS = cfg.get("scan_interval_seconds", 60)
    CONTINOUS_MODE = cfg.get("continous_mode", 0)
    MULTIPLIER = cfg.get("width_multiplier", 2.0) 

    if DIR_INPUT == DIR_OUTPUT:
        DIR_OUTPUT = "output"   

    os.makedirs(DIR_INPUT, exist_ok=True)
    os.makedirs(DIR_OUTPUT, exist_ok=True)

    paths_pdf_queue = Queue()
    
    # 1. Arrancar el Hilo Consumidor (Worker)
    worker_thread = threading.Thread(
        target=queue_worker, 
        args=(paths_pdf_queue, DIR_OUTPUT, MULTIPLIER), 
        daemon=True,
        name="WorkerThread" # Damos nombre al hilo para identificarlo en el log
    )
    worker_thread.start()
    logging.info("Hilo de procesamiento (Worker) iniciado.")

    if CONTINOUS_MODE:
        # 2. Arrancar Watchdog
        pdf_enqueuer = Enqueuer(paths_pdf_queue)
        observer = Observer()
        observer.schedule(pdf_enqueuer, DIR_INPUT, recursive=False)
        observer.start()
        logging.info(f"Watchdog observando directorio: {DIR_INPUT}")

        # 3. PRIMER BARRIDO INICIAL
        folder_sweep(DIR_INPUT, paths_pdf_queue)

        # 4. BUCLE PRINCIPAL
        contador_tiempo = 0
        try:
            while True:
                time.sleep(1)
                contador_tiempo += 1
                
                if contador_tiempo >= INTERVALO_SCANNER_SEGUNDOS:
                    folder_sweep(DIR_INPUT, paths_pdf_queue)
                    contador_tiempo = 0 
        except KeyboardInterrupt:
            logging.info("Detención solicitada por el usuario (KeyboardInterrupt).")
        finally:
            observer.stop()
            observer.join()
            logging.info("Aplicación cerrada correctamente.")
        
    else:
        logging.info("Modo de ejecución única (No continuo). Realizando barrido...")
        folder_sweep(DIR_INPUT, paths_pdf_queue)
        paths_pdf_queue.join()
        logging.info("Procesamiento finalizado.")