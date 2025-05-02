# 📄 PDF Section-wise Chunking for LLM Applications

This repository is an automated Python-based utility to convert PDF documents into clean, section-wise chunks using their headings and paragraph content. It is particularly useful for LLM ingestion pipelines such as RAG (Retrieval-Augmented Generation) systems, document Q&A systems, or search-based applications. 

### Format for section-wise chunking of pdfs:
Section-wise chunking creates a single meaningful unit by combining each heading with its associated paragraph, preserving the document’s natural structure. This approach has shown to outperform methods like recursive or fixed-size chunking, which often break context. Automated heading-paragraph detection achieves around 80% retrieval accuracy, making it highly effective for LLM applications.
```
filename|Heading1|Heading2|...|Paragraph

```

## 1. 📝 Word-Based Section-Wise Chunking (Preferred)
Best automation, minimal manual effort.

✅ Converts PDF to DOCX using Microsoft Word
✅ Automatically detects:

- Standard Word headings (Heading 1 to Heading 11)
- Custom heading styles:
- Bold/italic/underlined phrases at the start of a paragraph
- Outputs a CSV with each line as a full section path and content.

### Setup & Dependencies
```
pip install python-docx
pip install pywin32

```

## 2. 📝 HTML-Based Section-Wise Chunking
- Converts PDF files to HTML using PyMuPDF (fitz) for precise layout extraction.
- Uses BeautifulSoup to parse HTML and extract structured content.
- Automatically removes footers and headers based on user configuration.
- Detects headings based on <b> tags and associates them with their corresponding paragraph content.
- Outputs chunked content into a CSV file for further processing.
- Supports flexible formats, including:
- - With or without <div> tags.
- - Customizable heading/footer lines to skip.
- - Clean UTF-8 ASCII-only output.


### Setup & Dependencies
For PDF to HTML Section-wise chunking:
```
pip install pymupdf beautifulsoup4 html5lib

```

### PDF to HTML Section wise chunking:
Following are the arguments used for the function CreateSectionWiseChunks(...)

| Parameter       | Description                                                      |      |
| --------------- | ---------------------------------------------------------------- | ---- |
| `filename`      | Name of the file (used as prefix in chunks).                     |      |
| `extractedText` | HTML content obtained from `ConvertPdfToHTML`.                   |      |
| `divTag`        | Whether to chunk using `<div>` blocks per page.                  |      |
| `headerTag`     | HTML tag used for heading detection (ignored if `divTag=True`).  |      |
| `headcount`     | Number of lines to skip from the top of each page (headers).     |      |
| `footercount`   | Number of lines to skip from the bottom of each page (footers).  |      |
| `tagToRemove`   | Currently unused. Placeholder for future tag exclusions.         |      |
| `total_pages`   | Total number of pages in the document.                           |      |
| `separator`     | Separator to use between heading and content fields (default: \` | \`). |


## Notes
- The heading is detected based on <b> tags. You may modify logic inside CreateSectionWiseChunks to support different heading formats.
- Ensure your PDF documents are structured well (e.g., tagged PDFs) for the most accurate HTML conversion.
- Modify headcount and footercount depending on the presence of repeated headers/footers in the PDF.