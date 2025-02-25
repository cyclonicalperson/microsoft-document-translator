import sys
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QComboBox, QTextEdit, QLabel, \
    QProgressBar, QFileDialog


class TranslationApp(QWidget):
    def __init__(self):
        super().__init__()

        # Set the window properties
        self.setWindowTitle('Document Translator')
        self.setGeometry(100, 100, 800, 600)

        # Main layout
        main_layout = QVBoxLayout()

        # Top navigation bar (File, Edit, Settings)
        self.create_top_nav_bar()

        # Left layout (For language selection)
        left_layout = QVBoxLayout()
        left_layout.addWidget(QLabel("From Language"))
        self.from_language_dropdown = QComboBox()
        self.from_language_dropdown.addItems(["English", "Spanish", "French", "German", "Chinese"])
        left_layout.addWidget(self.from_language_dropdown)

        left_layout.addWidget(QLabel("To Language"))
        self.to_language_dropdown = QComboBox()
        self.to_language_dropdown.addItems(["English", "Spanish", "French", "German", "Chinese"])
        left_layout.addWidget(self.to_language_dropdown)

        # Text area for input (Document Textbox)
        self.text_input = QTextEdit(self)
        self.text_input.setPlaceholderText("Paste or upload your document here...")
        left_layout.addWidget(self.text_input)

        # Button to select the input document
        self.select_document_button = QPushButton('Select Document')
        self.select_document_button.clicked.connect(self.select_document)
        left_layout.addWidget(self.select_document_button)

        # Button to select the target folder
        self.select_folder_button = QPushButton('Select Target Folder')
        self.select_folder_button.clicked.connect(self.select_target_folder)
        left_layout.addWidget(self.select_folder_button)

        # Translate button
        self.translate_button = QPushButton('Translate')
        self.translate_button.clicked.connect(self.translate_document)
        left_layout.addWidget(self.translate_button)

        # Progress bar for translation
        self.progress_bar = QProgressBar(self)
        left_layout.addWidget(self.progress_bar)

        # Right layout (For translated document display)
        right_layout = QVBoxLayout()
        self.translated_text_area = QTextEdit(self)
        self.translated_text_area.setPlaceholderText("Your translated document will appear here...")
        self.translated_text_area.setReadOnly(True)
        right_layout.addWidget(self.translated_text_area)

        # Buttons for copy, save, and share
        self.copy_button = QPushButton('Copy')
        self.save_button = QPushButton('Save')
        self.share_button = QPushButton('Share')

        right_layout.addWidget(self.copy_button)
        right_layout.addWidget(self.save_button)
        right_layout.addWidget(self.share_button)

        # Combine the left and right layout horizontally
        h_layout = QHBoxLayout()
        h_layout.addLayout(left_layout)
        h_layout.addLayout(right_layout)

        # Add top navigation bar and horizontal layout to the main layout
        main_layout.addLayout(self.top_nav_bar)
        main_layout.addLayout(h_layout)

        # Set the main layout
        self.setLayout(main_layout)

    def create_top_nav_bar(self):
        """Create a simple top navigation bar with File, Edit, and Settings options."""
        self.top_nav_bar = QHBoxLayout()

        self.file_button = QPushButton('File')
        self.edit_button = QPushButton('Edit')
        self.settings_button = QPushButton('Settings')

        self.top_nav_bar.addWidget(self.file_button)
        self.top_nav_bar.addWidget(self.edit_button)
        self.top_nav_bar.addWidget(self.settings_button)

    def select_document(self):
        """Open a file dialog to select the input document."""
        file_dialog = QFileDialog(self)
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFiles)
        file_dialog.setNameFilter("Text files (*.txt);;All files (*.*)")
        if file_dialog.exec():
            file_paths = file_dialog.selectedFiles()
            if file_paths:
                self.text_input.setText(f"Selected document: {file_paths[0]}")

    def select_target_folder(self):
        """Open a folder dialog to select the target folder."""
        folder_dialog = QFileDialog(self)
        folder_dialog.setFileMode(QFileDialog.FileMode.Directory)
        folder_dialog.setOption(QFileDialog.Option.ShowDirsOnly, True)
        if folder_dialog.exec():
            folder_path = folder_dialog.selectedFiles()
            if folder_path:
                self.translated_text_area.setText(f"Selected folder: {folder_path[0]}")

    def translate_document(self):
        """Simulate document translation and show progress."""
        self.progress_bar.setValue(0)  # Reset the progress bar
        self.translated_text_area.clear()

        # Simulate a translation progress
        for i in range(101):
            self.progress_bar.setValue(i)
            QApplication.processEvents()  # Update UI to show progress
            if i == 100:
                self.translated_text_area.setText("Translated text will appear here...")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TranslationApp()
    window.show()
    sys.exit(app.exec())
