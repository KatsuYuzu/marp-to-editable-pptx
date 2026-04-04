# Marp Editable PPTX

A VS Code extension that exports [Marp](https://marp.app/) Markdown presentations to editable PowerPoint (.pptx) files.

## Usage

1. Open a Marp Markdown file (`.md`) in VS Code
2. Press `F1` and run **Marp: Export to Editable PPTX**
3. Choose a save location in the dialog
4. The editable `.pptx` file is generated

## Requirements

- A Chromium-based browser (Google Chrome or Microsoft Edge) must be installed

## How it works

1. Converts the Markdown to HTML using [@marp-team/marp-cli](https://github.com/marp-team/marp-cli)
2. Launches a headless browser to render the slides and extract layout information
3. Builds an editable `.pptx` where each shape (text, image, etc.) is individually placed

## License

MIT
