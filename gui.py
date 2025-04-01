import sys
import os
from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QComboBox, QTextEdit, QLabel, \
    QProgressBar, QFileDialog
from word_translator import WordTranslationApp
from excel_translator import ExcelTranslationApp
from powerpoint_translator import PowerPointTranslationApp

LANGUAGE_CODES = {
    # Major Languages
    "Arabic": "ar", "Chinese (Simplified)": "zh-Hans", "Chinese (Traditional)": "zh-Hant", "English": "en",
    "French": "fr", "German": "de", "Hindi": "hi", "Italian": "it", "Japanese": "ja", "Korean": "ko",
    "Portuguese (Brazil)": "pt", "Russian": "ru", "Spanish": "es",

    # Complete List (A-Z)
    "Afrikaans": "af", "Albanian": "sq", "Amharic": "am", "Armenian": "hy", "Assamese": "as", "Azerbaijani": "az",
    "Bangla": "bn", "Bashkir": "ba", "Basque": "eu", "Bosnian": "bs", "Bulgarian": "bg",
    "Cantonese (Traditional)": "yue", "Catalan": "ca", "Croatian": "hr", "Czech": "cs", "Danish": "da", "Dari": "prs",
    "Dutch": "nl", "Estonian": "et", "Fijian": "fj", "Filipino": "fil", "Finnish": "fi", "Galician": "gl",
    "Georgian": "ka", "Greek": "el", "Gujarati": "gu", "Haitian Creole": "ht", "Hebrew": "he", "Hungarian": "hu",
    "Icelandic": "is", "Indonesian": "id", "Irish": "ga", "Kannada": "kn", "Kazakh": "kk", "Khmer": "km",
    "Klingon": "tlh-Latn", "Lao": "lo", "Latvian": "lv", "Lithuanian": "lt", "Macedonian": "mk", "Malagasy": "mg",
    "Malay": "ms", "Malayalam": "ml", "Maltese": "mt", "Maori": "mi", "Marathi": "mr",
    "Mongolian (Cyrillic)": "mn-Cyrl", "Myanmar": "my", "Nepali": "ne", "Norwegian": "nb", "Odia": "or", "Pashto": "ps",
    "Persian": "fa", "Polish": "pl", "Punjabi": "pa", "Romanian": "ro", "Serbian (Cyrillic)": "sr-Cyrl",
    "Serbian (Latin)": "sr-Latn", "Slovak": "sk", "Slovenian": "sl", "Somali": "so", "Swahili": "sw", "Swedish": "sv",
    "Tamil": "ta", "Telugu": "te", "Thai": "th", "Turkish": "tr", "Ukrainian": "uk", "Urdu": "ur", "Uzbek": "uz",
    "Vietnamese": "vi"
}


class TranslationThread(QThread):
    progress_updated = pyqtSignal(int)
    message_updated = pyqtSignal(str)

    def __init__(self, input_path, output_path, target_lang, document_type):
        super().__init__()
        self.input_path = input_path
        self.output_path = output_path
        self.target_lang = target_lang
        self.document_type = document_type

    def run(self):
        try:
            def progress_callback(progress):
                self.progress_updated.emit(progress)

            # Perform translation based on document type
            if self.document_type == "docx":
                WordTranslationApp(input_path=self.input_path, output_path=self.output_path,
                                   target_lang=self.target_lang, progress_callback=progress_callback)
            elif self.document_type == "xlsx":
                ExcelTranslationApp(input_path=self.input_path, output_path=self.output_path,
                                    target_lang=self.target_lang, progress_callback=progress_callback)
            elif self.document_type == "pptx":
                PowerPointTranslationApp(input_path=self.input_path, output_path=self.output_path,
                                         target_lang=self.target_lang, progress_callback=progress_callback)

            self.message_updated.emit(f"Document translated successfully to {self.output_path}")
            self.progress_updated.emit(100)
        except Exception as e:
            self.message_updated.emit(f"Error: {str(e)}")
            self.progress_updated.emit(0)


class TranslationApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Document Translator')
        self.setGeometry(100, 100, 400, 500)
        self.setWindowIcon(QIcon("icon.ico"))

        self.translation_thread = None
        self.document_type = ""
        self.document_path = ""
        self.folder_path = ""

        layout = QVBoxLayout()

        # Initialize language dropdowns
        layout.addWidget(QLabel("From Language"))
        self.from_language_dropdown = QComboBox()
        self.from_language_dropdown.addItems(sorted(LANGUAGE_CODES.keys()))
        self.from_language_dropdown.setCurrentText("English")
        layout.addWidget(self.from_language_dropdown)

        layout.addWidget(QLabel("To Language"))
        self.to_language_dropdown = QComboBox()
        self.to_language_dropdown.addItems(sorted(LANGUAGE_CODES.keys()))
        self.to_language_dropdown.setCurrentText("Serbian (Latin)")
        layout.addWidget(self.to_language_dropdown)

        # Document selection buttons
        self.select_document_button = QPushButton('Select Document')
        self.select_document_button.clicked.connect(self.select_document)
        layout.addWidget(self.select_document_button)

        self.select_folder_button = QPushButton('Select Target Destination')
        self.select_folder_button.clicked.connect(self.select_target_destination)
        layout.addWidget(self.select_folder_button)

        self.translate_button = QPushButton('Translate')
        self.translate_button.clicked.connect(self.translate_document)
        layout.addWidget(self.translate_button)

        # Text area for translation progress messages
        self.text_box = QTextEdit(self)
        self.text_box.setPlaceholderText("The translation progress will appear here...")
        self.text_box.setReadOnly(True)
        layout.addWidget(self.text_box)

        # Progress bar for translation progress
        self.progress_bar = QProgressBar(self)
        layout.addWidget(self.progress_bar)
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.show()

        self.setLayout(layout)

    def select_document(self):
        """Open a file dialog to select the input document and store its path."""
        file_dialog = QFileDialog(self)
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
        file_dialog.setNameFilter("Documents (*.docx *.xlsx *.pptx)")

        if file_dialog.exec():
            selected_file = file_dialog.selectedFiles()[0]
            self.document_type = selected_file.split('.')[-1].lower()
            if selected_file:
                self.document_path = selected_file
                self.text_box.append(f"Selected document: {self.document_path}")
            else:
                self.text_box.append("No document selected.")

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
            selected_file = file_dialog.selectedFiles()[0]
            self.folder_path = selected_file
            self.text_box.append(f"Selected output path: {self.folder_path}")

    def translate_document(self):
        """Start translation in a separate thread."""
        selected_to_lang = self.to_language_dropdown.currentText()
        to_lang_code = LANGUAGE_CODES.get(selected_to_lang, "en")

        if self.document_path and self.folder_path:
            self.progress_bar.setValue(0)  # Reset progress bar
            self.translation_thread = TranslationThread(self.document_path, self.folder_path, to_lang_code,
                                                        self.document_type)
            self.translation_thread.progress_updated.connect(self.update_progress)
            self.translation_thread.message_updated.connect(self.update_message)
            self.translation_thread.start()
        else:
            self.text_box.append("<font color='red'>Please select both input and output paths.")

    def update_progress(self, progress):
        """Update the progress bar."""
        self.progress_bar.setValue(progress)

    def update_message(self, message):
        """Update the message area with progress updates."""
        self.text_box.append(message)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TranslationApp()
    window.show()
    sys.exit(app.exec())
