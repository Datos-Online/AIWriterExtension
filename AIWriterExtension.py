import sys
import os 
import uno
import json
import unohelper
import officehelper # type: ignore
import urllib.request
import urllib.parse
import platform

from com.sun.star.task import XJobExecutor # type: ignore
from com.sun.star.awt import XActionListener # type: ignore
from com.sun.star.frame import XDispatchProvider, XDispatch # type: ignore
from com.sun.star.beans import PropertyValue # type: ignore
from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS # type: ignore
from com.sun.star.awt import MessageBoxResults as MSG_RESULTS # type: ignore
try:
    from com.sun.star.util import PathSubstitution  # type: ignore
    HAS_PATH_SUBSTITUTION = True
except ImportError:
    HAS_PATH_SUBSTITUTION = False

API_URL = "https://api.openai.com/v1/chat/completions"
SEETINGS_FILE = "aiwriter.json"
DEFAULT_LANG = "es"
EXTENSION_NAME = "AIWriterExtension.oxt"

class AIWriterExtension(unohelper.Base, XJobExecutor):
    """
    Clase principal para la extensión AI Writer de LibreOffice.

    Esta clase implementa la interfaz XJobExecutor para actuar como el punto de
    entrada para los comandos ejecutados desde la interfaz de usuario de LibreOffice
    (menús, barras de herramientas, atajos de teclado). También hereda de unohelper.Base
    para la integración con el entorno de UNO.

    Se encarga de:
    - Inicializar la extensión y cargar los recursos de idioma.
    - Gestionar los comandos del usuario como 'settings', 'translate', 'complete', etc.
    - Interactuar con el documento de Writer para obtener texto e insertar resultados.
    - Administrar la configuración de la extensión (API key, modelo, etc.).
    - Comunicarse con la API de OpenAI para procesar el texto.
    - Mostrar diálogos de configuración y notificación al usuario.
    """

    def __init__(self, ctx):
        """
        Inicializa la instancia de AIWriterExtension.
        
        Este constructor configura el entorno de la extensión, cargando las cadenas de
        idioma e inicializando los servicios esenciales de LibreOffice. Maneja diferentes
        contextos de ejecución (desde dentro de LibreOffice o externamente).

        :param ctx: El contexto del componente UNO, que proporciona acceso a los servicios de LibreOffice.
        """
        self.lang = json.loads(self.get_language())
        
        self.ctx = ctx
        try:
            self.sm = ctx.getServiceManager()
            self.desktop = XSCRIPTCONTEXT.getDesktop() # type: ignore
        except NameError:
            self.sm = ctx.ServiceManager
            self.desktop = self.ctx.getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)

    def trigger(self, command):
        """
        Ejecuta un comando específico dentro de LibreOffice Writer.

        Esta función actúa como punto de entrada para diversas acciones de la extensión,
        tales como abrir el diálogo de configuración, traducir texto, o procesar texto
        con la IA (completar, resumir, mejorar, expandir).

        :param command: La acción a ejecutar. Los valores posibles son 'settings', 'translate', 'complete', 'summarize', 'improve', 'expand'.

        :raises Exception: Si ocurre un error durante la ejecución de un comando, se muestra un mensaje de error.
        """

        desktop = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
        model = desktop.getCurrentComponent()
        if not hasattr(model, "Text"):
            model = self.desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())
        selection = model.CurrentController.getSelection()

        if command == "hello":
            return
        
        elif command == "settings":
            try:
                result = self.settings_box(self.lang['settings'])
                if not result:
                    return
                else:
                    openai_api_key = result['openai_api_key']
                    self.set_config("openai_api_key", openai_api_key)
                    openai_model = result["model"]
                    self.set_config("model", openai_model)
                    max_tokens = result["max_tokens"]
                    self.set_config("max_tokens", max_tokens)
                    temperature = result["temperature"]
                    self.set_config("temperature", temperature)
                return
            except Exception as e:
                text_range = selection.getByIndex(0)
                text_range.setString(text_range.getString() + ":error: " + str(e))
        
        elif command == "translate":
            doc = self.get_document()
            if not doc:
                return
            
            selection = doc.getCurrentController().getViewCursor().getString()            
            if not selection.strip():
                self.show_dialog(self.lang["error"], self.lang["no_text_selected"])
                return
            
            openai_api_key = self.get_config("openai_api_key", "")
            if openai_api_key == "":
                self.show_dialog(self.lang["error"], self.lang["no_api_key"])
                return
            
            try:
                result = self.translation_box(self.lang['translate'])
                if not result:
                    return
                else:
                    language = result['language']
                    self.set_config("language", language)
                    if language != "":                        
                        ai_result = self.process_text(selection, command, language)
                        if ai_result:
                            self.insert_text(doc, ai_result, selection, command)

            except Exception as e:
                text_range = selection.getByIndex(0)
                text_range.setString(text_range.getString() + ":error: " + str(e))
        
        else:
            doc = self.get_document()
            if not doc:
                return
            
            selection = doc.getCurrentController().getViewCursor().getString()            
            if not selection.strip():
                self.show_dialog(self.lang["error"], self.lang["no_text_selected"])
                return

            openai_api_key = self.get_config("openai_api_key", "")
            if openai_api_key == "":
                self.show_dialog(self.lang["error"], self.lang["no_api_key"])
                return
        
            ai_result = self.process_text(selection, command)
            if ai_result:
                self.insert_text(doc, ai_result, selection, command)
        

    def get_config(self,key,default):
        """
        Obtiene un valor de configuración del archivo de ajustes.

        Lee el archivo JSON de configuración y devuelve el valor asociado a una clave.
        Si el archivo o la clave no existen, devuelve un valor por defecto.

        :param key: La clave de configuración a buscar.
        :param default: El valor a devolver si la clave no se encuentra.
        :return: El valor de la configuración o el valor por defecto.
        """
        path_settings = self.sm.createInstanceWithContext('com.sun.star.util.PathSettings', self.ctx)
        user_config_path = getattr(path_settings, "UserConfig")
        if user_config_path.startswith('file://'):
            user_config_path = str(uno.fileUrlToSystemPath(user_config_path))

        # Asegurar que la ruta termine con el nombre del archivo
        config_file_path = os.path.join(user_config_path, SEETINGS_FILE)

        # Comprobar si el archivo existe
        if not os.path.exists(config_file_path):
            return default

        # Intentar cargar el contenido JSON del archivo
        try:
            with open(config_file_path, 'r') as file:
                config_data = json.load(file)
        except (IOError, json.JSONDecodeError):
            return default

        # Devolver el valor correspondiente a la clave, o el valor por defecto si la clave no se encuentra
        return config_data.get(key, default)

    def set_config(self, key, value):
        """
        Establece un valor de configuración en el archivo de ajustes.

        Escribe un par clave-valor en el archivo de configuración JSON.
        Si el archivo no existe, lo crea.

        :param key: La clave de configuración a establecer.
        :param value: El valor a guardar.
        """
        path_settings = self.sm.createInstanceWithContext('com.sun.star.util.PathSettings', self.ctx)
        user_config_path = getattr(path_settings, "UserConfig")

        if user_config_path.startswith('file://'):
            user_config_path = str(uno.fileUrlToSystemPath(user_config_path))

        # Asegurar que la ruta termine con el nombre del archivo
        config_file_path = os.path.join(user_config_path, SEETINGS_FILE)

        # Cargar la configuración existente si el archivo existe
        if os.path.exists(config_file_path):
            try:
                with open(config_file_path, 'r') as file:
                    config_data = json.load(file)
            except (IOError, json.JSONDecodeError):
                config_data = {}
        else:
            config_data = {}

        # Actualizar la configuración con el nuevo par clave-valor
        config_data[key] = value

        # Escribir la configuración actualizada de nuevo en el archivo
        try:
            with open(config_file_path, 'w') as file:
                json.dump(config_data, file, indent=4)
        except IOError as e:
            # Manejar posibles errores de E/S (opcional)
            print(f"Error writing to {config_file_path}: {e}")

    def get_document(self):
        """
        Obtiene el componente de documento actual en LibreOffice.

        :return: El objeto del componente del documento activo.
        """
        desktop = self.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx
        )
        return desktop.getCurrentComponent()
    
    def process_text(self, text, command, lang=""):
        """
        Procesa el texto seleccionado enviándolo a la API de OpenAI.

        Construye el prompt adecuado según el comando, y envía la solicitud a la API.

        :param text: El texto a procesar.
        :param command: El comando de IA a ejecutar (ej. 'complete', 'summarize').
        :param lang: El idioma de destino para la traducción (opcional).
        :return: El texto procesado por la IA, o None si el comando no es válido.
        """
        prompt_map = {
            "complete": f"{self.lang['prompt_complete']}: {text}",
            "summarize": f"{self.lang['prompt_summarize']}: {text}",
            "improve": f"{self.lang['prompt_improve']}: {text}",
            "expand": f"{self.lang['prompt_expand']}: {text}",
            "translate": f"{self.lang['prompt_translate']} {lang}: {text}",
        }        
        if command not in prompt_map:
            return None
        
        openai_api_key = self.get_config("openai_api_key", "")
        headers = {
            "Authorization": f"Bearer {openai_api_key}",
            "Content-Type": "application/json"
        }
        
        payload = {
            "model": self.get_config("model", "gpt-4o-mini"),
            "messages": [
                {
                    "role":"assistant", 
                    "content": [
                        { 
                            "type": "text",
                            "text": f"{self.lang['prompt_assistant']}"
                        }
                    ]
                },{
                    "role": "user", 
                    "content": [
                        {
                            "type": "text",
                            "text": prompt_map[command]
                        }
                    ]
                }
            ],
            "max_completion_tokens": int(self.get_config("max_tokens", "1000")),
            "temperature": float(self.get_config("temperature", "0.5"))
        }
        try:
            data = json.dumps(payload).encode("utf-8")  # Convertir payload a JSON y codificarlo
            req = urllib.request.Request(API_URL, data=data, headers=headers, method="POST")
            
            with urllib.request.urlopen(req) as response:
                if response.status == 200:
                    response_data = json.loads(response.read().decode("utf-8"))
                    return response_data["choices"][0]["message"]["content"].strip()
                else:
                    self.show_dialog(self.lang["error"], f"{response.read().decode('utf-8')}")
        except Exception as e:
            self.show_dialog(self.lang["error"], str(e))

    def insert_text(self, doc, new_text, selection, command):
        """
        Inserta el texto generado por la IA en el documento.

        Reemplaza la selección actual con el texto original seguido del resultado de la IA,
        envuelto en bloques de inicio y fin para mayor claridad.

        :param doc: El documento activo.
        :param new_text: El texto generado por la IA para insertar.
        :param selection: El texto original que fue seleccionado.
        :param command: El comando de IA que se ejecutó.
        """
        language = ""
        if(command == "translate"):
            language = self.get_config("language", "")
        
        command = self.lang['command_' + command]

        cursor = doc.getCurrentController().getViewCursor()
        cursor.setString(selection + f"\n\n[---{self.lang['block_start']} {command.upper()} {language}---]\n{new_text}\n[/---{self.lang['block_end']} {command.upper()} {language}---]\n\n")

    def show_dialog(self, title, message, type = "ERRORBOX"):
        """
        Muestra un cuadro de diálogo modal con un título y un mensaje.

        Utiliza el toolkit de la AWT de UNO para crear y mostrar un cuadro de diálogo.

        :param title: El título del cuadro de diálogo.
        :param message: El mensaje a mostrar en el cuadro de diálogo.
        :param type: El tipo de cuadro de diálogo (ej. "ERRORBOX", "INFOBOX").
        """
        localContext = uno.getComponentContext()
    
        # Obtener el servicio de ventanas
        smgr = localContext.ServiceManager
        toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", localContext)
        
        # Crear un cuadro de diálogo de tipo información
        msgbox = toolkit.createMessageBox(
            None,  # Ventana padre
            type,  # Tipo de cuadro de diálogo (INFOBOX, WARNINGBOX, ERRORBOX, etc.)
            MSG_BUTTONS.BUTTONS_OK,  # Botones a mostrar (OK, YES_NO, etc.)
            title,  # Título del diálogo
            message  # Mensaje del diálogo
        )
        
        # Mostrar el cuadro de diálogo y obtener la respuesta
        result = msgbox.execute()
        
        """Verificar la respuesta del usuario
        if result == MSG_RESULTS.OK:
            print("El usuario hizo clic en OK")
        """

    def settings_box(self,title="", x=None, y=None):
        """
        Crea y muestra el cuadro de diálogo de configuración.

        Permite al usuario introducir la clave de la API de OpenAI, el modelo,
        la temperatura y el máximo de tokens.

        :param title: El título para el cuadro de diálogo.
        :param x: Coordenada X opcional para la posición del diálogo.
        :param y: Coordenada Y opcional para la posición del diálogo.
        :return: Un diccionario con los valores de configuración guardados.
        """
        WIDTH = 600
        HORI_MARGIN = VERT_MARGIN = 8
        BUTTON_WIDTH = 100
        BUTTON_HEIGHT = 26
        LABEL_HEIGHT = BUTTON_HEIGHT  + 5
        EDIT_HEIGHT = 24
        HEIGHT = 400
        import uno
        from com.sun.star.awt.PosSize import POS, SIZE, POSSIZE # type: ignore
        from com.sun.star.awt.PushButtonType import OK, CANCEL # type: ignore
        from com.sun.star.util.MeasureUnit import TWIP # type: ignore
        ctx = uno.getComponentContext()
        def create(name):
            return ctx.getServiceManager().createInstanceWithContext(name, ctx)
        dialog = create("com.sun.star.awt.UnoControlDialog")
        dialog_model = create("com.sun.star.awt.UnoControlDialogModel")
        dialog.setModel(dialog_model)
        dialog.setVisible(False)
        dialog.setTitle(title)
        dialog.setPosSize(0, 0, WIDTH, HEIGHT, SIZE)
        def add(name, type, x_, y_, width_, height_, props):
            model = dialog_model.createInstance("com.sun.star.awt.UnoControl" + type + "Model")
            dialog_model.insertByName(name, model)
            control = dialog.getControl(name)
            control.setPosSize(x_, y_, width_, height_, POSSIZE)
            for key, value in props.items():
                setattr(model, key, value)
        label_width = 120
        control_width = 450
        add("label_openai_api_key", "FixedText", HORI_MARGIN, VERT_MARGIN+4, label_width, LABEL_HEIGHT, {"Label": f"{self.lang['openai_api_key']}:", "NoLabel": True})
        add("edit_openai_api_key", "Edit", HORI_MARGIN + label_width, VERT_MARGIN, control_width, EDIT_HEIGHT, {"Text": str(self.get_config("openai_api_key","")), "MultiLine":False})
        
        add("label_model", "FixedText", HORI_MARGIN, 62, label_width, LABEL_HEIGHT, {"Label": f"{self.lang['openai_model']}:", "NoLabel": True})
        add("edit_model", "Edit", HORI_MARGIN + label_width, 58, control_width, EDIT_HEIGHT, {"Text": str(self.get_config("model","gpt-4o-mini"))})
        
        add("label_max_tokens", "FixedText", HORI_MARGIN, 112, label_width, LABEL_HEIGHT, {"Label": f"{self.lang['max_tokens']}:", "NoLabel": True})
        add("edit_max_tokens", "Edit", HORI_MARGIN + label_width, 108, control_width, EDIT_HEIGHT, {"Text": str(self.get_config("max_tokens","1000"))})

        add("label_temperature", "FixedText", HORI_MARGIN, 162, label_width, LABEL_HEIGHT, {"Label": f"{self.lang['temperature']}:", "NoLabel": True})
        add("edit_temperature", "Edit", HORI_MARGIN+ label_width, 158, control_width, EDIT_HEIGHT, {"Text": str(self.get_config("temperature","0.5"))})

        add("btn_ok", "Button", WIDTH - 120, HEIGHT - 50, BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": OK, "DefaultButton": True})

        frame = create("com.sun.star.frame.Desktop").getCurrentFrame()
        window = frame.getContainerWindow() if frame else None
        dialog.createPeer(create("com.sun.star.awt.Toolkit"), window)
        if not x is None and not y is None:
            ps = dialog.convertSizeToPixel(uno.createUnoStruct("com.sun.star.awt.Size", x, y), TWIP)
            _x, _y = ps.Width, ps.Height
        elif window:
            ps = window.getPosSize()
            _x = ps.Width / 2 - WIDTH / 2
            _y = ps.Height / 2 - HEIGHT / 2
        dialog.setPosSize(_x, _y, 0, 0, POS)
        
        edit_openai_api_key = dialog.getControl("edit_openai_api_key")
        edit_openai_api_key.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(self.get_config("openai_api_key","")))))
        
        edit_model = dialog.getControl("edit_model")
        edit_model.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(self.get_config("model","gpt-4o-mini")))))

        edit_max_tokens = dialog.getControl("edit_max_tokens")
        edit_max_tokens.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(self.get_config("max_tokens","1000")))))

        edit_temperature = dialog.getControl("edit_temperature")
        edit_temperature.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(self.get_config("temperature","0.5")))))
        
        btn_ok = dialog.getControl("btn_ok")
        btn_ok.setFocus()

        if dialog.execute():
            result = {
                "openai_api_key":edit_openai_api_key.getModel().Text,
                "model": edit_model.getModel().Text,
                "temperature": float(edit_temperature.getModel().Text)
            }
            if edit_max_tokens.getModel().Text.isdigit():
                result["max_tokens"] = int(edit_max_tokens.getModel().Text)
        else:
            result = {
                "openai_api_key": str(self.get_config("openai_api_key","")),
                "model": str(self.get_config("model","gpt-4o-mini")),
                "max_tokens": str(self.get_config("max_tokens","1000")),
                "temperature": float(self.get_config("temperature", "0.5"))
            }

        dialog.dispose()
        return result
    
    def translation_box(self,title="", x=None, y=None):
        """
        Crea y muestra un cuadro de diálogo para introducir el idioma de traducción.

        Permite al usuario especificar el idioma al que desea traducir el texto
        seleccionado.

        :param title: El título para el cuadro de diálogo.
        :param x: Coordenada X opcional para la posición del diálogo.
        :param y: Coordenada Y opcional para la posición del diálogo.
        :return: Un diccionario con el idioma de destino.
        """
        WIDTH = 600
        HORI_MARGIN = VERT_MARGIN = 8
        BUTTON_WIDTH = 100
        BUTTON_HEIGHT = 26
        HORI_SEP = 8
        VERT_SEP = 4
        LABEL_HEIGHT = BUTTON_HEIGHT  + 5
        EDIT_HEIGHT = 24
        HEIGHT = 200
        import uno
        from com.sun.star.awt.PosSize import POS, SIZE, POSSIZE # type: ignore
        from com.sun.star.awt.PushButtonType import OK, CANCEL # type: ignore
        from com.sun.star.util.MeasureUnit import TWIP # type: ignore
        ctx = uno.getComponentContext()
        def create(name):
            return ctx.getServiceManager().createInstanceWithContext(name, ctx)
        dialog = create("com.sun.star.awt.UnoControlDialog")
        dialog_model = create("com.sun.star.awt.UnoControlDialogModel")
        dialog.setModel(dialog_model)
        dialog.setVisible(False)
        dialog.setTitle(title)
        dialog.setPosSize(0, 0, WIDTH, HEIGHT, SIZE)
        def add(name, type, x_, y_, width_, height_, props):
            model = dialog_model.createInstance("com.sun.star.awt.UnoControl" + type + "Model")
            dialog_model.insertByName(name, model)
            control = dialog.getControl(name)
            control.setPosSize(x_, y_, width_, height_, POSSIZE)
            for key, value in props.items():
                setattr(model, key, value)
        label_width = 120
        control_width = 450
        add("label_language", "FixedText", HORI_MARGIN, VERT_MARGIN+4, label_width, LABEL_HEIGHT, {"Label": f"{self.lang['translate_to']}:", "NoLabel": True})
        add("edit_language", "Edit", HORI_MARGIN + label_width, VERT_MARGIN, control_width, EDIT_HEIGHT, {"Text": str(self.get_config("language",""))})
        
        add("btn_ok", "Button", WIDTH - 120, HEIGHT - 50, BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": OK, "DefaultButton": True})

        frame = create("com.sun.star.frame.Desktop").getCurrentFrame()
        window = frame.getContainerWindow() if frame else None
        dialog.createPeer(create("com.sun.star.awt.Toolkit"), window)
        if not x is None and not y is None:
            ps = dialog.convertSizeToPixel(uno.createUnoStruct("com.sun.star.awt.Size", x, y), TWIP)
            _x, _y = ps.Width, ps.Height
        elif window:
            ps = window.getPosSize()
            _x = ps.Width / 2 - WIDTH / 2
            _y = ps.Height / 2 - HEIGHT / 2
        dialog.setPosSize(_x, _y, 0, 0, POS)
        
        edit_language = dialog.getControl("edit_language")
        edit_language.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(self.get_config("language","")))))
        
        edit_language.setFocus()

        if dialog.execute():
            result = {
                "language":edit_language.getModel().Text
            }
        else:
            result = {
                "language": str(self.get_config("language",""))
            }

        dialog.dispose()
        return result
    
    def find_extension_path(self):
        """
        Encuentra la ruta de instalación de la extensión.

        Esta función determina la ruta completa al archivo de la extensión (.oxt)
        manejando las diferencias entre sistemas operativos (Windows y Linux) y
        adaptándose a distintas versiones de LibreOffice, incluyendo aquellas
        que no disponen del servicio `PathSubstitution`.

        El método construye una lista de rutas potenciales donde la extensión podría
        estar instalada. Primero, intenta usar `PathSubstitution` para resolver
        variables de ruta de LibreOffice. Si este servicio no está disponible
        (en versiones más antiguas), recurre a métodos manuales como consultar el
        registro de Windows o construir rutas basadas en el directorio de inicio del usuario.

        :return: La ruta completa al archivo de la extensión si se encuentra.
                 Devuelve `None` si no se puede localizar la extensión, tras mostrar
                 un cuadro de diálogo de error.
        """
        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager

        system = platform.system()
        potential_paths = []

        if HAS_PATH_SUBSTITUTION:
            ps = smgr.createInstanceWithContext(PathSubstitution, ctx)
            if system == "Windows":
                # Ruta de Windows: más robusta, tiene en cuenta las diferentes versiones de LibreOffice
                potential_paths = [
                    ps.substituteVariables("${user}/AppData/Roaming/LibreOffice/4/user/uno_packages/cache/uno_packages/", True),
                    ps.substituteVariables("${user}/AppData/Roaming/LibreOffice/user/uno_packages/cache/uno_packages/", True),
                    ps.substituteVariables("${ProgramFiles}/LibreOffice/program/uno_packages/cache/uno_packages/", True)  # para la ubicación de instalación en Program Files.
                ]
            elif system == "Linux":
                potential_paths = [
                    ps.substituteVariables("${user}/uno_packages/cache/uno_packages/", True),
                    ps.substituteVariables("${user}/.config/libreoffice/4/user/uno_packages/cache/uno_packages/", True)  # Para instalaciones flatpak
                ]
            else:
                self.show_dialog(self.lang['error'], f"{self.lang['unsuported_os']} {system}.")
                return None
        else:
            # Para versiones antiguas de LibreOffice sin PathSubstitution:
            if system == "Windows":
                # Alternativa a la ruta directa, menos robusta pero debería funcionar en la mayoría de los casos
                try:
                    import winreg
                    key_path = r"SOFTWARE\LibreOffice\UNO\InstallPath"
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                        install_dir = winreg.QueryValue(key, None)

                    potential_paths.append(os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "LibreOffice", "4", "user", "uno_packages", "cache", "uno_packages"))
                    potential_paths.append(os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "LibreOffice", "user", "uno_packages", "cache", "uno_packages"))
                    potential_paths.append(os.path.join(install_dir, "program", "uno_packages", "cache", "uno_packages"))

                except (ImportError, OSError, FileNotFoundError) as e:
                    self.show_dialog(self.lang['error'], f"{self.lang['lo_path_error']}: {e}")
                    return None
            elif system == "Linux":
                potential_paths.append(os.path.join(os.path.expanduser("~"), ".config", "libreoffice", "4", "user", "uno_packages", "cache", "uno_packages"))
                potential_paths.append(os.path.join(os.path.expanduser("~"), "uno_packages", "cache", "uno_packages"))
            else:
                self.show_dialog(self.lang['error'], f"{self.lang['unsuported_os']} {system}.")
                return None

        for potential_path in potential_paths:
            if HAS_PATH_SUBSTITUTION:
                base_path = uno.fileUrlToSystemPath(potential_path)
            else:
                base_path = potential_path

            if os.path.exists(base_path):
                for folder in os.listdir(base_path):
                    extension_path = os.path.join(base_path, folder)
                    if os.path.isdir(extension_path) and EXTENSION_NAME in os.listdir(extension_path):
                        return os.path.join(extension_path, EXTENSION_NAME)

        self.show_dialog(self.lang['error'], f"ERROR: Could not find extension '{EXTENSION_NAME}'. Checked potential paths: {potential_paths}")
        return None


    def get_ui_language(self):
        """Obtiene el idioma de la interfaz de usuario de LibreOffice."""
        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        config_provider = smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)

        # Definir el parámetro de acceso a la configuración
        param = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        param.Name = "nodepath"
        param.Value = "/org.openoffice.Setup/L10N"

        # Acceder al nodo de configuración
        access = config_provider.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (param,))
        
        # Obtener el idioma de la UI
        return access.getPropertyValue("ooLocale")

    def get_language(self):
        """
        Carga el archivo de idioma correcto basado en la configuración de la UI de LibreOffice.

        Busca un archivo .json que coincida con el código de idioma de la UI. Si no lo encuentra,
        recurre al idioma por defecto (español).
        """
        base_path = self.find_extension_path()
        lang = self.get_ui_language()[0:2]
        filename = os.path.join(base_path, "lang", f"{lang}.json")

        if not os.path.exists(filename):
            filename = os.path.join(base_path, "lang", f"{DEFAULT_LANG}.json")

        with open(filename, "r", encoding="utf-8") as lang_file:
            return lang_file.read()

# Iniciando desde un IDE de Python
def main():
    """
    Punto de entrada principal para ejecutar la extensión desde un IDE de Python o al arrancar LibreOffice.
    Inicializa el contexto y crea una instancia de la clase de la extensión.
    """
    try:
        ctx = XSCRIPTCONTEXT # type: ignore
    except NameError:
        ctx = officehelper.bootstrap()
        if ctx is None:
            print("ERROR: Could not bootstrap default Office.")
            sys.exit(1)
    job = AIWriterExtension(ctx)
    job.trigger("hello")
# Iniciando desde la línea de comandos
if __name__ == "__main__":
    main()
# Registrar la implementación de la extensión en LibreOffice
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    AIWriterExtension,
    "com.datosonline.AIWriterExtension.do",
    ("com.sun.star.task.Job",),
)
