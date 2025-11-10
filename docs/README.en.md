## AI Extension for LibreOffice Writer

This extension for LibreOffice Writer integrates artificial intelligence capabilities directly into the text editor, allowing you to efficiently improve, summarize, expand, complete, and translate texts using the OpenAI API.

### Compatibility

The extension has been tested and is compatible with the following environments:
*   **Operating Systems:** Ubuntu 22.04 (and derived distributions), Windows 10 and 11.
*   **LibreOffice:** Version 24.2 or higher.

**Note on macOS:** Currently, the extension is not compatible with macOS. The main difficulty lies in the automatic detection of the installation path in this operating system, a necessary step to load the interface language files.

## Features

*   **Expand text:** Adds more detail and content to the existing text.
*   **Complete text:** Generates continuations for the selected text.
*   **Improve wording:** Optimizes the quality and style of the text.
*   **Summarize text:** Creates concise summaries of long passages.
*   **Translate text:** Translates the selected text into different languages.
*   **Customizable settings:** Allows you to configure your OpenAI API key, AI model, temperature, and maximum number of tokens.

## Installation

To install the extension in LibreOffice Writer, follow these steps:

1.  **Download the extension file:** Get the `AIWriterExtension.oxt` file from the repository's releases section or directly from the source.
2.  **Open LibreOffice Writer:** Start your LibreOffice Writer application.
3.  **Access the Extension Manager:** Go to `Tools > Extension Manager...` in the main menu.
4.  **Add the extension:** In the Extension Manager, click the `Add...` button and select the `AIWriterExtension.oxt` file you downloaded.
5.  **Complete the installation:** Follow the on-screen instructions to finish the installation. You may be asked to accept a license.
6.  **Restart LibreOffice:** For the changes to take effect, close and reopen LibreOffice Writer.

## Usage

Once the extension is installed, you can access its features as follows:

1.  **Select the text:** In your Writer document, select the piece of text you want to process with the AI.
2.  **Execute a command:** Access the extension's options (these usually appear in a new menu or toolbar within LibreOffice Writer).
3.  **Choose the action:** Select the desired action: "Complete", "Expand", "Improve wording", "Summarize", or "Translate".
4.  **For translation:** If you choose "Translate", a dialog box will appear for you to specify the target language.

The AI-processed text will be inserted into your document.

## Configuration

Before using the extension, you need to configure your OpenAI API key and other parameters:

1.  **Access the settings:** Look for the extension's "Settings" option (usually in the same menu where the AI actions are located).
2.  **Enter your API key:** In the settings dialog, enter your OpenAI API key.
3.  **Adjust the parameters:** You can modify the AI model (default is `gpt-4o-mini`), the temperature (which controls the creativity of the response), and the maximum number of tokens (maximum length of the response).

## Known Issues

1.  The user interface may become temporarily unresponsive while waiting for the API response.
2.  The original text formatting (such as bold, italics, or colors) is not preserved in the generated response, which is inserted as plain text.

## License

This project is licensed under the Apache 2.0 License.

