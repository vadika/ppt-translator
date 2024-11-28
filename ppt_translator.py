import argparse
from pptx import Presentation
import os
import anthropic

def translate_text(text, target_lang, shape_type=""):
    if not text.strip():  # Skip empty or whitespace-only text
        return text
        
    print(f"\nTranslating {shape_type}:")
    print(f"Original: '{text[:50]}{'...' if len(text) > 50 else ''}'")
    
    # Get API key from environment
    api_key = os.getenv('ANTHROPIC_API_KEY')
    if not api_key:
        raise ValueError("ANTHROPIC_API_KEY environment variable not set")
    
    # Map language codes to full names
    lang_map = {
        'ru': 'Russian',
        'fi': 'Finnish',
        'et': 'Estonian'
    }
    
    client = anthropic.Anthropic(api_key=api_key)
    try:
        prompt = f"Translate the following text to {lang_map[target_lang]}. Only return the translation, no explanations:\n\n{text}"
        message = client.messages.create(
            model="claude-3-opus-20240229",
            max_tokens=1000,
            temperature=0,
            messages=[{"role": "user", "content": prompt}]
        )
        # Extract the text content and ensure it's a string
        translated = message.content[0].text if isinstance(message.content, list) else str(message.content)
        print(f"Translated: '{translated[:50]}{'...' if len(translated) > 50 else ''}'")
        return translated
    except Exception as e:
        print(f"Translation error: {e}")
        return text

def translate_table(table, target_lang):
    """Translate text in table cells"""
    for row in table.rows:
        for cell in row.cells:
            if cell.text.strip():
                cell.text = translate_text(cell.text, target_lang, "Table cell")

def translate_smartart(shape, target_lang):
    """Recursively translate text in SmartArt graphics"""
    if hasattr(shape, 'text'):
        shape_type = "SmartArt text"
        shape.text = translate_text(shape.text, target_lang, shape_type)
    
    # Process child shapes if they exist
    if hasattr(shape, 'shapes'):
        for child_shape in shape.shapes:
            translate_smartart(child_shape, target_lang)

def translate_presentation(input_file, target_lang):
    # Open presentation
    print(f"\nOpening presentation: {input_file}")
    prs = Presentation(input_file)
    
    # Translate each slide's text content
    total_slides = len(prs.slides)
    for slide_num, slide in enumerate(prs.slides, 1):
        print(f"\nProcessing slide {slide_num}/{total_slides}")
        for shape in slide.shapes:
            if shape.shape_type == 6:  # MSO_SHAPE_TYPE.GROUP (SmartArt)
                translate_smartart(shape, target_lang)
            elif shape.shape_type == 19:  # MSO_SHAPE_TYPE.TABLE
                translate_table(shape.table, target_lang)
            elif hasattr(shape, "text"):
                shape_type = type(shape).__name__
                shape.text = translate_text(shape.text, target_lang, shape_type)
    
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
    
    print(f"\nStarting translation to {args.language.upper()}...")
    output_file = translate_presentation(args.input_file, args.language)
    print(f"\nTranslation complete!")
    print(f"Translated presentation saved as: {output_file}")

if __name__ == "__main__":
    main()
