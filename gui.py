import sys
import os
import threading
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


class TranslationApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Document Translator')
        self.setGeometry(100, 100, 800, 600)

        # Set the default values
        self.document_type = ""
        self.document_path = ""  # Path to the document being translated
        self.folder_path = ""  # Path to where the translated document will be saved

        """Initialize the UI components."""
        main_layout = QVBoxLayout()

        # Left layout (For language selection and document type)
        left_layout = QVBoxLayout()

        # Create language selection dropdowns
        left_layout.addWidget(QLabel("From Language"))
        self.from_language_dropdown = QComboBox()
        self.from_language_dropdown.addItems(sorted(LANGUAGE_CODES.keys()))  # Add languages dynamically
        self.from_language_dropdown.setCurrentText("English")  # Set default choice to English
        left_layout.addWidget(self.from_language_dropdown)

        left_layout.addWidget(QLabel("To Language"))
        self.to_language_dropdown = QComboBox()
        self.to_language_dropdown.addItems(sorted(LANGUAGE_CODES.keys()))  # Add languages dynamically
        self.to_language_dropdown.setCurrentText("Serbian (Latin)")  # Set default choice to Serbian (Latin)
        left_layout.addWidget(self.to_language_dropdown)

        # Button to select the input document
        self.select_document_button = QPushButton('Select Document')
        self.select_document_button.clicked.connect(self.select_document)
        left_layout.addWidget(self.select_document_button)

        # Button to select the target folder
        self.select_folder_button = QPushButton('Select Target Destination')
        self.select_folder_button.clicked.connect(self.select_target_destination)
        left_layout.addWidget(self.select_folder_button)

        # Translate button
        self.translate_button = QPushButton('Translate')
        self.translate_button.clicked.connect(self.translate_document)
        left_layout.addWidget(self.translate_button)

        # Text area for input (Document Textbox)
        self.text_box = QTextEdit(self)
        self.text_box.setPlaceholderText("The translation progress will appear here...")
        self.text_box.setReadOnly(True)
        left_layout.addWidget(self.text_box)

        # Progress bar for translation
        self.progress_bar = QProgressBar(self)
        left_layout.addWidget(self.progress_bar)

        # Right layout (For translated document display)
        right_layout = QVBoxLayout()
        self.translated_text_area = QTextEdit(self)
        self.translated_text_area.setPlaceholderText("Your translated document will appear here...")
        self.translated_text_area.setReadOnly(True)
        right_layout.addWidget(self.translated_text_area)

        # Combine the left and right layout horizontally
        h_layout = QHBoxLayout()
        h_layout.addLayout(left_layout)
        h_layout.addLayout(right_layout)

        main_layout.addLayout(h_layout)
        self.setLayout(main_layout)

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
            # Create a new thread for translation to avoid blocking the GUI
            translation_thread = threading.Thread(target=self.run_translation, args=(self.document_path, self.folder_path, to_lang_code))
            translation_thread.start()
        else:
            # Notify the user to select the input and output paths first
            self.text_box.append(f"The input and output must first be selected.")

    def run_translation(self, input_path, output_path, target_lang):
        """Handle the translation process in a separate thread."""
        try:
            if self.document_type == "docx":
                WordTranslationApp(input_path=input_path, output_path=output_path, target_lang=target_lang)

            elif self.document_type == "xlsx":
                ExcelTranslationApp(input_path=input_path, output_path=output_path, target_lang=target_lang)

            elif self.document_type == "pptx":
                PowerPointTranslationApp(input_path=input_path, output_path=output_path, target_lang=target_lang)

            else:
                raise ValueError("Unsupported file type. Only .docx, .xlsx, and .pptx files are supported.")

            # Show success message after translation is complete
            self.text_box.append(f"Document translated successfully to {output_path}")

        except Exception as e:
            self.text_box.append(f"Error: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TranslationApp()
    window.show()
    sys.exit(app.exec())
