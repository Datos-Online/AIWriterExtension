# CLAUDE.md

Este archivo da instrucciones a Claude Code (claude.ai/code) para trabajar con el código de este repositorio.

## Qué es este proyecto

Una extensión (`.oxt`) para LibreOffice Writer escrita en Python (UNO) que envía el texto seleccionado a la API Chat Completions de OpenAI para ampliarlo, completarlo, mejorarlo, resumirlo o traducirlo. Compatible con Linux y Windows 10/11, y LibreOffice 24.2 o superior. macOS no está soportado (ver "Carga de idiomas" más abajo).

No hay linter ni dependencias de terceros. El código se ejecuta dentro del Python incluido en LibreOffice, así que solo están disponibles la librería estándar y `uno`/`unohelper` (las peticiones HTTP se hacen con `urllib`). La lógica que no depende de UNO tiene pruebas `unittest` (ver "Cómo probar").

## Compilar e instalar

```bash
./create_extension.sh          # Linux: comprime los archivos en dist/AIWriterExtension.oxt
create_extension.bat           # Windows: usa tar -a (formato zip) y renombra .zip -> .oxt
unopkg add -f dist/AIWriterExtension.oxt   # instala/reemplaza en LibreOffice; luego reiniciar LibreOffice
```

La lista de archivos empaquetados está escrita a mano en **ambos** scripts (`FILES_TO_INCLUDE`). Un archivo nuevo que deba ir en la extensión se añade a los dos scripts, y los componentes UNO y archivos de configuración `.xcu` también deben registrarse en `META-INF/manifest.xml`. Los módulos nuevos dentro de `pythonpath/` ya van incluidos (se empaqueta la carpeta entera, sin `__pycache__`) y no se registran en `manifest.xml`. La carpeta `tests/` no se empaqueta. Para cada versión publicada, subir el número en `description.xml`.

## Arquitectura

El código se divide en dos partes:

- `AIWriterExtension.py`: la clase `AIWriterExtension` (`XJobExecutor`), registrada al final del archivo como servicio UNO `com.datosonline.AIWriterExtension.do`. Contiene solo lo que depende de UNO: `trigger`, los diálogos, `insert_text`, la ruta `UserConfig` y la carga de idiomas.
- `pythonpath/aiwriter/`: lógica pura, sin `uno`, probada con `unittest`. LibreOffice añade `pythonpath/` a `sys.path` al cargar la extensión (al ejecutar el `.py` desde un IDE, el propio archivo lo añade). Estos módulos deben ser compatibles con Python 3.9, que es el que trae LibreOffice 24.2 en Windows.
  - `models.py`: tabla de capacidades por prefijo de modelo (`MODEL_CAPABILITIES`, `supports`).
  - `api.py`: prompt (`PROMPT_KEYS`, `build_prompt`), payload (`build_payload`), envío (`send`, con `API_URL` y `API_TIMEOUT`) e interpretación de la respuesta. Los errores se lanzan como `ApiError` con una clave de `lang/*.json`; la clase UNO los muestra en un diálogo con `ApiError.message(self.lang)`.
  - `config.py`: lectura y escritura de `aiwriter.json` y `validate_settings`, que devuelve una clave de error de `lang/*.json`.

**Despacho de comandos:** las entradas de la interfaz invocan la URL `service:com.datosonline.AIWriterExtension.do?<comando>`, y LibreOffice llama a `trigger(<comando>)`. Estas URL se conectan en:
- `Addons.xcu`: el submenú "AI Writer" insertado después del menú Herramientas, el `AddonMenu` (configuración) y el menú contextual de texto (`OfficePopupMenuMerging`). Los elementos del submenú están duplicados entre la barra de menús y el menú contextual.
- `Accelerators.xcu`: atajos Ctrl+Shift+A/C/M/R/T/G para Writer.

Para añadir un comando de IA nuevo hay que tocar: los dos bloques de menú de `Addons.xcu`, opcionalmente `Accelerators.xcu`, `PROMPT_KEYS` en `pythonpath/aiwriter/api.py`, y las claves `command_<cmd>` / `prompt_<cmd>` en **ambos** `lang/en.json` y `lang/es.json`.

**Flujo de una petición:** `trigger` → lee el texto seleccionado mediante el cursor de la vista → `process_text` construye el prompt y el payload con `aiwriter.api` y lo envía a OpenAI (POST, timeout de `API_TIMEOUT` = 90 s) → `insert_text` reemplaza la selección con el texto original seguido del resultado envuelto en marcadores `[---START <COMANDO>---] ... [/---END <COMANDO>---]` (texto plano; no se conserva el formato). La llamada es síncrona, por lo que la interfaz se bloquea mientras espera.

**Particularidades del payload de OpenAI:** usa `max_completion_tokens` (no `max_tokens`), el prompt de instrucciones se envía con rol `assistant`, y se elimina `temperature` en los modelos que la rechazan según `MODEL_CAPABILITIES` (hoy, los que empiezan por `gpt-5` y `gpt-6`). Para un modelo nuevo con otras restricciones, se añade una entrada a esa tabla con su prueba en `tests/test_models.py`.

**Configuración:** la clave API, el modelo (por defecto `gpt-4o-mini`), la temperatura, los tokens máximos y el último idioma de traducción se guardan como JSON en `aiwriter.json`, dentro del directorio `UserConfig` de LibreOffice (`get_config` / `set_config` resuelven la ruta y delegan en `aiwriter.config`). Los diálogos (`settings_box`, `translation_box`) se construyen por código con `UnoControlDialog`, no con archivos `.xdl`.

**Carga de idiomas:** los textos de la interfaz y los prompts vienen de `lang/<idioma-ui>.json` (con `es` como respaldo). `__init__` llama a `get_language` → `find_extension_path`, que recorre los directorios `uno_packages/cache/uno_packages/*/` de LibreOffice buscando una carpeta que contenga `AIWriterExtension.oxt`. Por eso el nombre del archivo compilado debe ser exactamente `EXTENSION_NAME`, y los nuevos sistemas o formas de instalación (macOS, otras rutas de flatpak/snap) necesitan entradas en `find_extension_path` o la extensión falla al construirse.

## Cómo probar

**Pruebas automáticas** (sin LibreOffice, con cualquier Python 3 del sistema), desde la raíz del repositorio:

```bash
python3 -m unittest -v
```

Cubren `pythonpath/aiwriter/` (payload, tabla de capacidades, respuestas y errores de la API, configuración y validación) y comprueban que las claves de idioma usadas existen en `es.json` y `en.json`. Cualquier cambio en esos módulos debe venir con su prueba.

**Pruebas manuales:** tras `unopkg add -f` y reiniciar LibreOffice, verificar en Writer:
- Cada comando (ampliar, completar, mejorar, resumir, traducir) con texto seleccionado, desde el menú, el menú contextual y el atajo de teclado.
- Diálogo de configuración: guardar, cerrar y reabrir; la clave API, el modelo, la temperatura y los tokens deben conservarse.
- Casos de error: clave inválida, sin conexión y respuesta vacía. Debe mostrarse un mensaje claro, sin colgar LibreOffice.
- Interfaz en español y en inglés.
- Si el cambio afecta a rutas o al empaquetado, probar también en Windows y en Linux.

## Reglas de trabajo

- No añadir dependencias externas: solo librería estándar y `uno`/`unohelper`.
- Nunca subir claves API ni el archivo `aiwriter.json` al repositorio.
- Mantener la compatibilidad con Windows y Linux; no romper `find_extension_path`.
- Antes de cambios grandes (estructura del código, `manifest.xml`, menús, flujo de red), proponer un plan y esperar confirmación.
- Cambios pequeños: un commit por tarea, con mensaje en español.
- Trabajar en la rama `V2`; no modificar `main` directamente.

## No cambiar sin avisar

- El nombre del servicio UNO `com.datosonline.AIWriterExtension.do`: de él dependen los menús, los atajos y el registro.
- El nombre del archivo compilado (`EXTENSION_NAME`): de él depende la carga de idiomas.
- La estructura de claves de `lang/en.json` y `lang/es.json`: deben mantenerse sincronizadas.

## Deuda técnica conocida (candidatos a mejora)

Los números entre corchetes son los puntos de la auditoría en `docs/plan-V2.md`. Ese archivo es local, está en `.gitignore` y puede no existir en otra copia del repositorio.

**Robustez y API**
- [1.7] Si `find_extension_path` no encuentra la extensión, `__init__` falla sin mostrar ningún mensaje útil: `show_dialog` usa `self.lang` antes de que exista, y `os.path.join(None, …)` lanza un error. Además, `PathSubstitution` nunca se importa, así que solo se prueban rutas fijas. Sustituirlo por `PackageInformationProvider.getPackageLocation("com.datosonline.AIWriterExtension")`, que funciona en todos los sistemas, incluido macOS.
- [1.8] El prompt de instrucciones se envía con rol `assistant`; evaluar `system` o `developer` (en `api.build_payload` y su prueba).
- [1.9] Faltan los modelos `o1`/`o3`/`o4-mini` en `MODEL_CAPABILITIES`; rechazan `temperature`.
- [1.10] La llamada a OpenAI es síncrona y bloquea la interfaz. Si se hace en un hilo, el resultado debe insertarse en el hilo principal (por ejemplo, con `com.sun.star.awt.AsyncCallback`).
- [1.11] `_x/_y` quedan sin definir en `settings_box`/`translation_box` si no hay ventana. Si no hay un documento Writer activo, `trigger` crea uno nuevo.

**Seguridad y configuración**
- [2.2] La clave API se guarda en JSON de texto plano con los permisos del `umask`. Mejora rápida: `chmod 600` y la variable de entorno `OPENAI_API_KEY` como respaldo. Mejora completa: `com.sun.star.task.PasswordContainer`.
- [2.3] `config.set_value` no escribe de forma atómica; usar un archivo temporal y `os.replace`.
- [2.4] `API_URL` está fija en el código; falta una "URL base" configurable (Azure, Ollama).
- [2.5] `Accelerators.xcu` usa `oor:op="replace"` y puede pisar atajos nativos de Writer (Ctrl+Shift+C es "Registrar cambios").
- [2.6] `description.xml` declara la versión mínima 4.0 de LibreOffice, pero se soporta 24.2 o superior.
- [2.7] Falta un aviso de privacidad: el texto seleccionado se envía a OpenAI.

**Calidad del código**
- [3.2] La clase UNO sigue concentrando la interfaz: diálogos duplicados en `settings_box` y `translation_box`, validaciones repetidas en `trigger`, y la configuración se relee del disco en cada `get_config`.
- [3.3] Código muerto: el comando `"hello"`, imports sin uso (`XActionListener`, `XDispatchProvider`, `XDispatch`, `PropertyValue`, `MSG_RESULTS`, `urllib.parse`), `HORI_SEP`/`VERT_SEP`, `import uno` repetido y la rama `HAS_PATH_SUBSTITUTION`.
- [3.5] Hay textos sin localizar: el error de `find_extension_path` está en inglés fijo. La clave `unsuported_os` tiene una errata; si se renombra, hay que hacerlo en los dos archivos de idioma a la vez.
- [3.6] La lista de archivos del build está duplicada en `.sh` y `.bat`; un único `build.py` con `zipfile` serviría para los dos sistemas.

## Convenciones

Los comentarios del código, los docstrings, los mensajes de commit y el README principal están en español (`docs/README.en.md` es el README en inglés); mantener ese estilo.
