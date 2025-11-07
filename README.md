[Read this in English](./docs/README.en.md)

## Extensión IA para LibreOffice Writer

Esta extensión para LibreOffice Writer integra capacidades de inteligencia artificial directamente en el editor de texto, permitiéndo mejorar, resumir, ampliar, completar y traducir textos de manera eficiente utilizando la API de OpenAI. 

## Características

*   **Ampliar texto:** Añade más detalles y contenido al texto existente.
*   **Completar texto:** Genera continuaciones para el texto seleccionado.
*   **Mejorar redacción:** Optimiza la calidad y el estilo del texto.
*   **Resumir texto:** Crea resúmenes concisos de pasajes largos.
*   **Traducir texto:** Traduce el texto seleccionado a diferentes idiomas.
*   **Configuración personalizable:** Permite configurar la clave API de OpenAI, el modelo de IA, la temperatura y el número máximo de tokens.

![Demostración de AI Writer Extension](./docs/media/ui-example.gif)

## Instalación

Para instalar la extensión en LibreOffice Writer, sigue estos pasos:

1.  **Descarga el archivo de extensión:** Obtén el archivo `AIWriterExtension.oxt` desde la sección de lanzamientos (releases) del repositorio o directamente desde la fuente.
2.  **Abre LibreOffice Writer:** Inicia tu aplicación LibreOffice Writer.
3.  **Accede al Gestor de Extensiones:** Ve a `Herramientas > Gestor de Extensiones...` en el menú principal.
4.  **Añade la extensión:** En el Gestor de Extensiones, haz clic en el botón `Añadir...` y selecciona el archivo `AIWriterExtension.oxt` que descargaste.
5.  **Completa la instalación:** Sigue las instrucciones en pantalla para finalizar la instalación. Es posible que se te pida aceptar una licencia.
6.  **Reinicia LibreOffice:** Para que los cambios surtan efecto, cierra y vuelve a abrir LibreOffice Writer.

## Uso

Una vez instalada la extensión, podrás acceder a sus funcionalidades de la siguiente manera:

1.  **Selecciona el texto:** En tu documento de Writer, selecciona el fragmento de texto que deseas procesar con la IA.
2.  **Ejecuta un comando:** Accede a las opciones de la extensión (estas suelen aparecer en un nuevo menú o barra de herramientas dentro de LibreOffice Writer).
3.  **Elige la acción:** Selecciona la acción deseada: "Completar", "Ampliar", "Mejorar redacción", "Resumir" o "Traducir".
4.  **Para la traducción:** Si eliges "Traducir", se te presentará un cuadro de diálogo para especificar el idioma de destino.

El texto procesado por la IA se insertará en tu documento.

## Configuración

Antes de usar la extensión, es necesario configurar tu clave API de OpenAI y otros parámetros:

1.  **Accede a la configuración:** Busca la opción de "Ajustes" de la extensión (generalmente en el mismo menú donde se encuentran las acciones de IA).
2.  **Introduce tu clave API:** En el cuadro de diálogo de configuración, introduce tu clave API de OpenAI.
3.  **Ajusta los parámetros:** Puedes modificar el modelo de IA (por defecto `gpt-4o-mini`), la temperatura (que controla la creatividad de la respuesta) y el número máximo de tokens (longitud máxima de la respuesta).

## Problemas Conocidos

1.  La interfaz de usuario puede dejar de responder temporalmente mientras se espera la respuesta de la API.
2.  El formato del texto original (como negrita, cursiva o colores) no se conserva en la respuesta generada, que se inserta como texto plano.

## Licencia

Este proyecto está bajo la licencia Apache 2.0.
