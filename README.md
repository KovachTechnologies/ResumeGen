# ResumeGen
**ResumeGen** is a Python tool that converts JSON-formatted resumes into professional Word documents (.docx). It supports hyperlinks, customizable styling, and both local and remote JSON sources, making it easy to generate polished resumes programmatically.

## Features

- **JSON to Word Conversion**: Transform structured JSON resumes into formatted `.docx` files.
- **Hyperlink Support**: Automatically converts HTML `<a>` tags in bullet points to clickable Word hyperlinks.
- **Customizable Styling**: Configure fonts, margins, and other document properties via a configuration dictionary.
- **Flexible Input**: Load resume data from local JSON files or remote URLs.
- **Robust Error Handling**: Includes logging and validation for reliable operation.
- **Command-Line Interface**: Simple CLI with options for file/URL input and output path.
- **Tested Codebase**: Comprehensive unit tests ensure functionality and stability.

## Installation

1. **Clone the Repository**:

```bash
git clone https://github.com/<your-username>/ResumeGen.git
cd ResumeGen
```

2. **Install Dependencies**:
Ensure Python 3.6+ is installed, then install required packages:

```bash
pip install -r requirements.txt
```

## Usage
Generate a Word resume from a JSON file or URL using the command-line interface.

### Example Commands
* From a Local JSON Fil

``` bash
python resumegen.py --file data/resume.json --output resume.docx
```

* From a Remote URL:

``` bash
python resumegen.py --url https://example.com/resume.json --output resume.docx
```

* Custom Output Name:
The `--output` option supports a `{datetime}` placeholder:

``` bash
python resumegen.py --file data/resume.json --output "{datetime}_resume.docx"
```

This generates a file prefixed by the date (YYYY-MM) e.g. `2025-04_resume.docx`.

```bash
{
  "header": {
    "name": "Nicholas Borbaki",
    "title": "Senior Software Engineer"
  },
  "contents": [
    {
      "title": "Experience",
      "id": 5,
      "content": [
        {
          "id": 1,
          "position": "Senior Software Engineer, Example",
          "date": "2021-Present",
          "items": ["Led development with <a href='https://example.com'>Example</a>."]
        }
      ]
    }
  ]
}
```

See `data/resume.json` for a complete example.

## Configuration
Customize document styling by editing the config dictionary in resumegen.py:

```bash
config = {
    "font": {"name": "Arial", "size": 8},
    "margins": {"top": 0.5, "bottom": 0.5, "left": 1.0, "right": 1.0},
}
```

Future versions may support external configuration files.

## Testing
ResumeGen includes a comprehensive test suite in `tests/test_resumegen.py`. To run tests:

```bash
python3 -m unittest tests/test_resumegen.py -v
```

The test suite covers:
* JSON loading and fetching
* Hyperlink processing
* Document generation
* Error handling

## License
This project is licensed under the MIT License. See the LICENSE file for details.
