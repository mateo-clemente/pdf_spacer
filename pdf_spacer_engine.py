import pymupdf
import time
import os
import threading
import config
import pythoncom
import win32com.client
import tempfile


from queue import Queue
from watchdog.events import FileSystemEvent, FileSystemEventHandler
from watchdog.observers import Observer
from pathlib import Path


def convert_ppt_to_pdf(ppt_path: str) -> str:
    """
    Convierte un archivo .ppt o .pptx a .pdf.
    Guarda el PDF temporal fuera de la carpeta monitored (en %TEMP%).
    """
    ppt_path = os.path.abspath(ppt_path)
    filename = os.path.basename(ppt_path)
    base_name = os.path.splitext(filename)[0]
    
    # 🛑 Guardar en la carpeta TEMP del sistema operativo para NO activar Watchdog
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
        print(f"📄 PPT convertido con éxito a PDF temporal: {pdf_path}")
        return pdf_path
    except Exception as e:
        print(f"❌ Error al convertir PPT {filename}: {e}")
        raise e
    finally:
        if powerpoint:
            powerpoint.Quit()
        pythoncom.CoUninitialize()

def folder_sweep(input_folder: str, queue: Queue):
    try:
        path_folder = Path(input_folder)
        valid_extensions = ("*.pdf", "*.ppt", "*.pptx")
        for ext in valid_extensions:
            for file in path_folder.glob(ext):
                ruta_completa = str(file.resolve())
                if is_valid_file(ruta_completa):
                    queue.put(ruta_completa)
    except Exception as e:
        print(f"❌ Error durante el barrido de carpeta: {e}")

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
    output_pdf = pymupdf.open()

    pdf_blank_spacer(input_pdf=input_pdf, output_pdf=output_pdf, multiplier=pdf_multiplier)

    output_pdf.save(output_pdf_path)
    output_pdf.close()
    input_pdf.close()
    os.remove(input_pdf_path)
    return

def queue_worker(queue:Queue, output, pdf_multiplier):
    while True:
        item = queue.get() 

        if not os.path.exists(item):
            queue.task_done()
            continue
        
        # Construir la ruta de salida basada en el nombre del archivo de entrada
        filename = os.path.basename(item)
        file_ext = os.path.splitext(filename)[1].lower()
        
        # Generar nombre único para el PDF final de salida
        base_name = os.path.splitext(filename)[0]
        output_pdf_path = os.path.join(output, f"extended_{base_name}.pdf")

        temp_pdf_to_clean = None
        
        try:
            print(f"⚙️ Procesando: {filename}")
            
            # CASO A: Si es un archivo de PowerPoint (.ppt o .pptx)
            if file_ext in ['.ppt', '.pptx']:
                temp_pdf_to_clean = convert_ppt_to_pdf(item)
                pdf_to_process = temp_pdf_to_clean
            # CASO B: Si ya es un PDF directamente
            elif file_ext == '.pdf':
                pdf_to_process = item
            else:
                print(f"⚠️ Archivo ignorado (formato no soportado): {filename}")
                continue

            # Procesamiento de agrandado (PyMuPDF)
            pdf_processor(
                input_pdf_path=pdf_to_process, 
                output_pdf_path=output_pdf_path, 
                pdf_multiplier=pdf_multiplier
            )
            
            # Borrar el archivo original procesado (.ppt/.pptx o .pdf)
            if os.path.exists(item):
                os.remove(item)
                
            print(f"✅ Procesado y eliminado original: {filename}")

        except Exception as e:
            print(f"❌ Error al procesar {filename}: {e}")
            
        finally:
            # Limpiar el PDF temporal intermedio si existió la conversión de PPT
            if temp_pdf_to_clean and os.path.exists(temp_pdf_to_clean):
                try:
                    os.remove(temp_pdf_to_clean)
                except Exception:
                    pass
            queue.task_done()

def is_valid_file(file_path: str) -> bool:
    """Filtra archivos temporales de Office y temporales de conversión."""
    filename = os.path.basename(file_path)
    # Ignorar archivos temporales de Office (~$archivo.pptx) y PDFs temporales
    if filename.startswith("~$") or "_temp_conv.pdf" in filename:
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
        print(f"enqueuing {os.path.abspath(event.src_path)} from on_created")
        self.process_event(event)

    def on_modified(self, event: FileSystemEvent) -> None:
        print(f"enqueuing {os.path.abspath(event.src_path)} from on_modified")
        self.process_event(event)

    def on_moved(self, event: FileSystemEvent) -> None:
        if not event.is_directory and is_valid_file(event.dest_path):
            print(f"enqueuing {os.path.abspath(event.src_path)} from on_moved")
            file_path = os.path.abspath(event.dest_path)
            self.paths_queue.put(file_path)

if __name__ == "__main__":

    cfg = config.load_config()

    DIR_INPUT = os.path.abspath(cfg.get("input_folder", "input"))
    DIR_OUTPUT = os.path.abspath(cfg.get("output_folder", "output"))
    INTERVALO_SCANNER_SEGUNDOS =  cfg.get("scan_interval_seconds", 60) # Frecuencia de respaldo
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
        daemon=True
    )
    worker_thread.start()

    if CONTINOUS_MODE:
        # 2. Arrancar Watchdog para eventos en tiempo real
        pdf_enqueuer = Enqueuer(paths_pdf_queue)
        observer = Observer()
        observer.schedule(pdf_enqueuer, DIR_INPUT, recursive=False)
        observer.start()
        

        # 3. PRIMER BARRIDO INICIAL: Para capturar lo que ya estaba antes de abrir el programa
        folder_sweep(DIR_INPUT, paths_pdf_queue)

        # 4. BUCLE PRINCIPAL CON BARRIDO DE RESPALDO (FALLBACK)
        contador_tiempo = 0
        try:
            while True:
                time.sleep(1)
                contador_tiempo += 1
                
                # Cada 60 segundos hace una comprobación de seguridad
                if contador_tiempo >= INTERVALO_SCANNER_SEGUNDOS:
                    folder_sweep(DIR_INPUT, paths_pdf_queue)
                    contador_tiempo = 0 # Reiniciar contador
        except KeyboardInterrupt:
            pass
        finally:
            observer.stop()
            observer.join()
        
    else:
        folder_sweep(DIR_INPUT, paths_pdf_queue)
        paths_pdf_queue.join()

    


