# Key Translations - Translation Memory Dictionary

A browser-based translation memory tool for professional EN-ES translators. Built for [Key Translations](https://keytranslations.cl), a Chilean translation company specializing in legal and technical English-Spanish translation.

## What it does

- **Dictionary search**: search across previous translations in English or Spanish with instant results, highlighting, and filters (language direction, content type, source document)
- **Web lookup**: one-click links to DeepL, Google Translate, Linguee, Reverso Context, and WordReference for cross-referencing
- **Import/Export**: load translation memory from JSON files and export filtered results
- **Fully offline**: single self-contained HTML file, no server or internet required to run

## How to use

1. Open `KeyTranslations_Dictionary.html` in any browser
2. Click **Import Translations** to load your translation memory JSON file
3. Search, filter, and browse your translations

## Translation data format

The app expects a JSON file with an array of translation pairs:

```json
[
  {
    "en": "English text",
    "es": "Spanish translation",
    "source": "document_name",
    "type": "paragraph"
  }
]
```

| Field    | Required | Description                                      |
|----------|----------|--------------------------------------------------|
| `en`     | yes      | English text                                     |
| `es`     | yes      | Spanish text                                     |
| `source` | no       | Source document identifier (defaults to filename) |
| `type`   | no       | Content type: `paragraph`, `table_cell`, etc.    |

## Extracting translations from DOCX pairs

Use `extract_translations.py` to extract parallel translations from paired Word documents (original + translation). The script:

1. Reads both DOCX files using `python-docx`
2. Uses section headings as structural anchors to align content
3. Extracts and pairs paragraphs and table cells
4. Outputs a JSON file in the format above

### Requirements

```
pip install python-docx
```

### Usage

Edit the file paths at the bottom of `extract_translations.py` to point to your English original and Spanish translation DOCX files, then run:

```
python extract_translations.py
```

## Privacy

Translation data files (`.docx`, `.json`) are excluded from this repository via `.gitignore`. All client content remains local.

## Technical details

- **Languages**: English (EN) and Spanish (ES)
- **Domains**: legal, technical, engineering, mining, energy, geophysics
- **Search**: client-side fuzzy matching with relevance scoring
- **Pagination**: 50 results per page
- **Data**: embedded or imported via JSON — no database or backend needed
