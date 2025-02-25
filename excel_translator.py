import asyncio
from googletrans import Translator
import openpyxl
from serbian_text_converter import SerbianTextConverter


class ExcelTranslationApp:
    def __init__(self, input_path, output_path, target_lang="en"):
        self.input_path = input_path
        self.output_path = output_path
        self.target_language_code = target_lang  # Default to "en" (English)

        # Start the translation process
        self.translate_excel_file()

    async def translate_cell(self, cell, translator, target_lang='en'):
        if cell.value and isinstance(cell.value, str):
            original_text = cell.value
            try:
                # Translate the text using Google Translate
                translated_text = await translator.translate(original_text, dest=target_lang)
                print(f"Original: {original_text} -> Translated: {translated_text.text}")

                # If translating to Serbian Latin, convert Cyrillic to Latin
                if target_lang == "sr_Latn":
                    translated_text = SerbianTextConverter.to_latin(translated_text.text)

                # Update the cell value with translated text
                cell.value = translated_text
            except Exception as e:
                print(f"Error translating cell value: {original_text}, Error: {e}")

    async def translate_sheet(self, sheet, translator, target_lang='en'):
        # Iterate through all rows and columns in the sheet
        for row in sheet.iter_rows():
            for cell in row:
                await self.translate_cell(cell, translator, target_lang)

    def translate_excel_file(self):
        if not self.input_path or not self.output_path or not self.target_language_code:
            print("Error: Please provide all required paths and target language.")
            return

        try:
            # Get the current event loop
            loop = asyncio.get_event_loop()

            # Run the async method with the event loop
            loop.run_until_complete(
                self.translate_excel_file_async(self.input_path, self.output_path, self.target_language_code))
        except Exception as e:
            print(f"Error: An error occurred: {e}")

    async def translate_excel_file_async(self, input_path, output_path, target_lang):
        try:
            # Load the Excel file
            workbook = openpyxl.load_workbook(input_path)
            translator = Translator()

            # Process each sheet in the workbook
            for sheet in workbook.sheetnames:
                current_sheet = workbook[sheet]
                print(f"Translating sheet: {sheet}")
                await self.translate_sheet(current_sheet, translator, target_lang)

            # Save the translated workbook
            workbook.save(output_path)
            print(f"Excel file has been translated and saved as {output_path}.")
        except Exception as e:
            print(f"Error: An error occurred: {e}")


# Create and run the application manually (Initialization with paths)
#input_path = "path_to_your_input_excel_file.xlsx"  # Replace with actual input file path
#output_path = "path_to_output_translated_excel_file.xlsx"  # Replace with desired output path
#target_lang = "sr_Latn"  # Replace with the desired language code (e.g., "es" for Spanish)

#app = ExcelTranslationApp(input_path, output_path, target_lang)
