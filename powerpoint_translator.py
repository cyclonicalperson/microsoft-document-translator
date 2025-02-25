import asyncio
from googletrans import Translator
from pptx import Presentation
from serbian_text_converter import SerbianTextConverter


class PowerPointTranslationApp:
    def __init__(self, input_path, output_path, target_lang="en"):
        self.input_path = input_path
        self.output_path = output_path
        self.target_language_code = target_lang  # Default to "en" (English)

        # Start the translation process
        self.translate_pptx_file()

    async def translate_text(self, text, translator, target_lang='en'):
        if text and isinstance(text, str):
            try:
                # Translate the text using Google Translate
                translated = await translator.translate(text, dest=target_lang)
                translated_text = translated.text  # Access the .text attribute of the Translated object
                print(f"Original: {text} -> Translated: {translated_text}")

                # If translating to Serbian Latin, convert Cyrillic to Latin
                if target_lang == "sr_Latn":
                    translated_text = SerbianTextConverter.to_latin(translated_text)

                return translated_text
            except Exception as e:
                print(f"Error translating text: {text}, Error: {e}")
                return text
        return text

    async def translate_slide(self, slide, translator, target_lang='en'):
        # Iterate through all shapes in the slide
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text:  # If the shape has text
                original_text = shape.text
                # Save the original font properties
                original_font_size = self.get_font_size(shape)

                # Translate the text
                translated_text = await self.translate_text(original_text, translator, target_lang)

                # Reapply the original font size and formatting
                self.set_text_and_formatting(shape, translated_text, original_font_size)

    def get_font_size(self, shape):
        """Extracts and returns the font size of the text in a shape."""
        if hasattr(shape, "text_frame") and shape.text_frame is not None:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    return run.font.size  # Assuming all runs in the shape have the same font size
        return None  # Default if no font size found

    def set_text_and_formatting(self, shape, translated_text, original_font_size):
        """Sets the translated text and applies the original font size."""
        if hasattr(shape, "text_frame") and shape.text_frame is not None:
            # Clear any existing text (needed to reapply formatted text)
            shape.text_frame.clear()

            # Create a new paragraph and set the translated text
            p = shape.text_frame.add_paragraph()
            p.text = translated_text

            # Reapply the original font size if available
            if original_font_size is not None:
                for run in p.runs:
                    run.font.size = original_font_size

    def translate_pptx_file(self):
        if not self.input_path or not self.output_path or not self.target_language_code:
            print("Error: Please provide all required paths and target language.")
            return

        try:
            # Get the current event loop
            loop = asyncio.get_event_loop()

            # Run the async method with the event loop
            loop.run_until_complete(
                self.translate_pptx_file_async(self.input_path, self.output_path, self.target_language_code))
        except Exception as e:
            print(f"Error: An error occurred: {e}")

    async def translate_pptx_file_async(self, input_path, output_path, target_lang):
        try:
            # Load the PowerPoint presentation
            presentation = Presentation(input_path)
            translator = Translator()

            # Process each slide in the presentation
            for slide_number, slide in enumerate(presentation.slides, start=1):
                print(f"Translating slide: {slide_number}")
                await self.translate_slide(slide, translator, target_lang)

            # Save the translated presentation
            presentation.save(output_path)
            print(f"PowerPoint file has been translated and saved as {output_path}.")
        except Exception as e:
            print(f"Error: An error occurred: {e}")


# Create and run the application manually (Initialization with paths)
#input_path = "path_to_your_input_pptx_file.pptx"  # Replace with actual input file path
#output_path = "path_to_output_translated_pptx_file.pptx"  # Replace with desired output path
#target_lang = "sr_Latn"  # Replace with the desired language code (e.g., "es" for Spanish)

#app = PowerPointTranslationApp(input_path, output_path, target_lang)
