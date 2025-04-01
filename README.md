# Microsoft Document Translator

Microsoft Document Translator is a Python application designed to translate Microsoft Office documents — specifically Word (.docx), Excel (.xlsx), and PowerPoint (.pptx) files — while preserving their original formatting. The tool leverages the Azure AI Translator service to perform translations.<br>

<p align="center">
  <img src="https://github.com/user-attachments/assets/37e67d3e-4065-4fd9-9c46-4d047f12e55a" />
</p>

## Features

- **Supported Formats**: Translates Word, Excel, and PowerPoint documents without altering their formatting.
- **User-Friendly Interface**: Offers a graphical user interface (GUI) for easy interaction.
- **Cross-Platform Compatibility**: Runs on Windows, macOS, and Linux systems.

## Usage

Run the compiled `.exe` from the `dist/` folder.

Or run the `.py` file:

```bash
python gui.py
```

## Requirements (for development)

- **Python**: Ensure Python 3.10 is installed.
- **Dependencies**: Install required Python packages using the provided `requirements.txt` file.
- **Azure Subscription**: An active Azure subscription is necessary to access the Azure AI Translator service.


## Installation (for development)

1. Clone the repository:

    ```bash
    git clone https://github.com/cyclonicalperson/microsoft-document-translator.git
    cd microsoft-document-translator
    ```

2. Install dependencies:

    ```bash
    pip install -r requirements.txt
    ```

3. Build the executable (optional):

    ```bash
    pyinstaller --onefile --noconsole --icon=icon.ico --name="Microsoft Document Translator" gui.py
    ```

## Azure AI Translator Setup

1. **Create an Azure Account**: If you don't have one, sign up at the [Azure portal](https://portal.azure.com/).

2. **Create a Translator Resource**:

   - Navigate to the Azure portal.
   - Search for "Translator" and create a new Translator resource.
   - Obtain the subscription key and endpoint URL from the resource's "Keys and Endpoint" section.

3. **Configure the Application**:

   - Replace the placeholder values in the application's configuration with your Azure subscription key and endpoint URL.

## License

This project is licensed under the GNU General Public License v3.0. For more details, refer to the [LICENSE](https://github.com/cyclonicalperson/microsoft-document-translator/blob/main/LICENSE) file.
