# PPTX to Obsidian Notes Converter

This Python script extracts text and images from a PowerPoint presentation (`.pptx` file), summarizes the text using a Hugging Face transformer model (`facebook/bart-large-cnn`), and exports the summarized content to a markdown file compatible with [Obsidian](https://obsidian.md/) note-taking software.

## Features

- **Text Extraction**: Extracts and cleans text from each slide in the `.pptx` file.
- **Image Extraction**: Saves images from the slides into a specified directory.
- **Text Summarization**: Summarizes the text content of each slide using a pre-trained model from Hugging Face.
- **Markdown Export**: Converts the slide data into a markdown file, formatting titles, summaries, and embedding images with HTML for resizing in Obsidian.

## Requirements

Ensure you have the following libraries installed before running the script:

```bash
pip install python-pptx transformers
