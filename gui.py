import sys
import os
from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QComboBox, QTextEdit, QLabel, \
    QProgressBar, QFileDialog
from word_translator import WordTranslationApp
from excel_translator import ExcelTranslationApp
from powerpoint_translator import PowerPointTranslationApp

LANGUAGE_CODES = {
            "Afrikaans": "af", "Albanian": "sq", "Amharic": "am", "Arabic": "ar", "Armenian": "hy", "Azerbaijani": "az",
            "Basque": "eu", "Belarusian": "be", "Bengali": "bn", "Bosnian": "bs", "Bulgarian": "bg", "Catalan": "ca",
            "Cebuano": "ceb", "Chichewa": "ny", "Chinese": "zh-CN", "Croatian": "hr", "Czech": "cs", "Danish": "da",
            "Dutch": "nl", "English": "en", "Esperanto": "eo", "Estonian": "et", "Filipino": "tl", "Finnish": "fi",
            "French": "fr", "Georgian": "ka", "German": "de", "Greek": "el", "Gujarati": "gu", "Haitian Creole": "ht",
            "Hebrew": "iw", "Hindi": "hi", "Hungarian": "hu", "Icelandic": "is", "Indonesian": "id", "Irish": "ga",
            "Italian": "it", "Japanese": "ja", "Javanese": "jw", "Kazakh": "kk", "Khmer": "km", "Korean": "ko",
            "Kurdish": "ku", "Lao": "lo", "Latin": "la", "Latvian": "lv", "Lithuanian": "lt", "Macedonian": "mk",
            "Malagasy": "mg", "Malay": "ms", "Malayalam": "ml", "Marathi": "mr", "Mongolian": "mn", "Myanmar": "my",
            "Nepali": "ne", "Norwegian": "no", "Pashto": "ps", "Persian": "fa", "Polish": "pl", "Portuguese": "pt",
            "Punjabi": "pa", "Romanian": "ro", "Russian": "ru", "Serbian (Cyrillic)": "sr", "Serbian (Latin)": "sr_Latn",
            "Sindhi": "sd", "Sinhala": "si", "Slovak": "sk", "Slovenian": "sl", "Somali": "so", "Spanish": "es",
            "Sundanese": "su", "Swahili": "sw", "Swedish": "sv", "Tajik": "tg", "Tamil": "ta", "Telugu": "te",
            "Thai": "th", "Turkish": "tr", "Ukrainian": "uk", "Urdu": "ur", "Uzbek": "uz", "Vietnamese": "vi",
            "Welsh": "cy", "Xhosa": "xh", "Yiddish": "yi", "Yoruba": "yo", "Zulu": "zu",
        }


class TranslationThread(QThread):
    # Signals to communicate with the main thread
    progress_updated = pyqtSignal(int)  # For progress bar updates (0 to 100)
    message_updated = pyqtSignal(str)  # For displaying messages (success, error)

    def __init__(self, input_path, output_path, target_lang, document_type):
        super().__init__()
        self.input_path = input_path
        self.output_path = output_path
        self.target_lang = target_lang
        self.document_type = document_type

    def run(self):
        """Handle the translation process in the background thread."""
        try:
            # Initialize the progress callback function
            def progress_callback(progress):
                self.progress_updated.emit(progress)

            # Perform translation
            if self.document_type == "docx":
                WordTranslationApp(input_path=self.input_path, output_path=self.output_path,
                                   target_lang=self.target_lang, progress_callback=progress_callback)

            elif self.document_type == "xlsx":
                ExcelTranslationApp(input_path=self.input_path, output_path=self.output_path,
                                    target_lang=self.target_lang, progress_callback=progress_callback)

            elif self.document_type == "pptx":
                PowerPointTranslationApp(input_path=self.input_path, output_path=self.output_path,
                                         target_lang=self.target_lang, progress_callback=progress_callback)

            # Emit final success message
            self.message_updated.emit(f"Document translated successfully to {self.output_path}")
            self.progress_updated.emit(100)

        except Exception as e:
            # Emit error message in case of failure
            self.message_updated.emit(f"Error: {str(e)}")
            self.progress_updated.emit(0)


class TranslationApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Document Translator')
        self.setGeometry(100, 100, 400, 500)

        # Set the default values
        self.translation_thread = None
        self.document_type = ""
        self.document_path = ""  # Path to the document being translated
        self.folder_path = ""  # Path to where the translated document will be saved

        """Initialize the UI components."""
        layout = QVBoxLayout()

        # Create language selection dropdowns
        layout.addWidget(QLabel("From Language"))
        self.from_language_dropdown = QComboBox()
        self.from_language_dropdown.addItems(sorted(LANGUAGE_CODES.keys()))  # Add languages dynamically
        self.from_language_dropdown.setCurrentText("English")  # Set default choice to English
        layout.addWidget(self.from_language_dropdown)

        layout.addWidget(QLabel("To Language"))
        self.to_language_dropdown = QComboBox()
        self.to_language_dropdown.addItems(sorted(LANGUAGE_CODES.keys()))  # Add languages dynamically
        self.to_language_dropdown.setCurrentText("Serbian (Latin)")  # Set default choice to Serbian (Latin)
        layout.addWidget(self.to_language_dropdown)

        # Button to select the input document
        self.select_document_button = QPushButton('Select Document')
        self.select_document_button.clicked.connect(self.select_document)
        layout.addWidget(self.select_document_button)

        # Button to select the target folder
        self.select_folder_button = QPushButton('Select Target Destination')
        self.select_folder_button.clicked.connect(self.select_target_destination)
        layout.addWidget(self.select_folder_button)

        # Translate button
        self.translate_button = QPushButton('Translate')
        self.translate_button.clicked.connect(self.translate_document)
        layout.addWidget(self.translate_button)

        # Text area for input (Document Textbox)
        self.text_box = QTextEdit(self)
        self.text_box.setPlaceholderText("The translation progress will appear here...")
        self.text_box.setReadOnly(True)
        layout.addWidget(self.text_box)

        # Progress bar for translation
        self.progress_bar = QProgressBar(self)
        layout.addWidget(self.progress_bar)
        self.progress_bar.setMinimum(0)  # Minimum value
        self.progress_bar.setMaximum(100)  # Maximum value
        self.progress_bar.setValue(0)  # Initial progress (0%)
        self.show()

        self.setLayout(layout)

    def select_document(self):
        """Open a file dialog to select the input document and store its path."""
        file_dialog = QFileDialog(self)
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
        file_dialog.setNameFilter("Documents (*.docx *.xlsx *.pptx)")  # Filter for document types

        if file_dialog.exec():
            selected_file = file_dialog.selectedFiles()[0]  # Returns list of selected files, takes the first file
            self.document_type = selected_file.split('.')[-1].lower()  # Update the type of selected document
            if selected_file:
                self.document_path = selected_file  # Get the file if one is selected
                self.text_box.append(f"Selected document: {self.document_path}")  # Show the path in the text box
            else:
                self.text_box.append("No document selected.")  # Handle case where no file is selected

    def select_target_destination(self):
        """Open a dialog to select the target destination (file)."""
        file_dialog = QFileDialog(self)
        file_dialog.setFileMode(QFileDialog.FileMode.AnyFile)
        if self.document_path:
            base_name = os.path.splitext(os.path.basename(self.document_path))[0]
            default_file_name = f"{base_name}-translated.{self.document_type}"
        else:
            default_file_name = f"translated.{self.document_type}"

        file_dialog.selectFile(default_file_name)

        if file_dialog.exec():
            selected_file = file_dialog.selectedFiles()[0]  # Get the selected file path
            self.folder_path = selected_file  # Store the selected destination path
            self.text_box.append(f"Selected output path: {self.folder_path}")  # Show the path in the text box

    def translate_document(self):
        """Start translation in a separate thread."""
        selected_to_lang = self.to_language_dropdown.currentText()

        # Map selected language names to language codes
        to_lang_code = LANGUAGE_CODES.get(selected_to_lang, "en")

        # Ensure document path and target folder are set
        if self.document_path and self.folder_path:
            self.progress_bar.setValue(0)  # Reset progress bar
            # Create a new thread for translation to avoid blocking the GUI
            self.translation_thread = TranslationThread(self.document_path, self.folder_path, to_lang_code,
                                                        self.document_type)
            self.translation_thread.progress_updated.connect(self.update_progress)
            self.translation_thread.message_updated.connect(self.update_message)
            self.translation_thread.start()
        else:
            # Notify the user to select the input and output paths first
            self.text_box.append(f"The input and output must first be selected.")

    def update_progress(self, progress):
        """Update the progress bar based on worker's progress."""
        self.progress_bar.setValue(progress)

    def update_message(self, message):
        """Update the message area based on worker's output."""
        self.text_box.append(message)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TranslationApp()
    window.show()
    sys.exit(app.exec())
