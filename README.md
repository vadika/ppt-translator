# PowerPoint Translator

A Python tool that uses AI to translate PowerPoint presentations while preserving their original formatting and structure. Powered by Claude 3 Opus via OpenRouter API.

## Features

- Translates all text content while preserving PowerPoint formatting
- Handles complex elements including:
  - Tables
  - SmartArt graphics
  - Placeholders
  - Nested shapes
- Supports translation to:
  - Russian (ru)
  - Finnish (fi)
  - Estonian (et)
  - Swedish (sv)
  - English (en)
  - Spanish (es)
  - German (de)
  - Latvian (lv)

## Prerequisites

- Python 3.x
- OpenRouter API key (set as OPENROUTER_API_KEY environment variable)

## Installation

1. Clone this repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

```bash
python ppt_translator.py INPUT_FILE LANGUAGE [-v]
```

Arguments:
- `INPUT_FILE`: Path to your PowerPoint file
- `LANGUAGE`: Target language code (ru, fi, et, sv, en, es, de, lv)
- `-v, --verbose`: Optional flag to enable debug output

The translated presentation will be saved as `{original_name}-{language_code}.pptx`

## Example

```bash
python ppt_translator.py presentation.pptx es
```

This will translate presentation.pptx to Spanish and save it as presentation-es.pptx

## Dependencies

- python-pptx: PowerPoint file handling
- anthropic: Claude AI integration

## Environment Variables

- `OPENROUTER_API_KEY`: Your OpenRouter API key (required)
