import argparse
from pptx import Presentation
from googletrans import Translator
import os

def translate_text(text, target_lang):
    translator = Translator()
    try:
        result = translator.translate(text, dest=target_lang)
        return result.text
    except:
        return text

def translate_presentation(input_file, target_lang):
    # Open presentation
    prs = Presentation(input_file)
    
    # Translate each slide's text content
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                shape.text = translate_text(shape.text, target_lang)
    
    # Generate output filename
    file_name, file_ext = os.path.splitext(input_file)
    output_file = f"{file_name}-{target_lang}{file_ext}"
    
    # Save translated presentation
    prs.save(output_file)
    return output_file

def main():
    parser = argparse.ArgumentParser(description='Translate PowerPoint content to specified language')
    parser.add_argument('input_file', help='Input PowerPoint file')
    parser.add_argument('language', choices=['ru', 'fi', 'et'], 
                        help='Target language (ru=Russian, fi=Finnish, et=Estonian)')
    
    args = parser.parse_args()
    
    output_file = translate_presentation(args.input_file, args.language)
    print(f"Translated presentation saved as: {output_file}")

if __name__ == "__main__":
    main()
