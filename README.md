# some_doc_viewers

This project contains a TypeScript-based document viewing solution that supports multiple file formats, including PPTX, PDF, XLSX, and DOCX. The main functionality is implemented in the `Viewer` component and individual preview components for each file type.

## How it works

The document viewing process can be summarized in the following steps:

1. Load the document file
2. Determine the file type
3. Render the appropriate preview component
4. Convert and display the document content

## Key components

1. **Viewer**: The main component that orchestrates the preview process based on file type.
2. **PPTXPreview**: Handles PPTX file preview using the PPTX class and pptxjs library.
3. **PDFPreview**: Renders PDF files using react-pdf library.
4. **XLSXPreview**: Displays Excel spreadsheets using xlsx library.
5. **DOCXPreview**: Converts and shows Word documents using mammoth library.
6. **PPTX class**: Handles PPTX to HTML conversion, based on the pptxjs library.

## PPTX Conversion and pptxjs

The PPTX to HTML conversion process is implemented in the `PPTX` class within the `src/helpers/pptx/PPTX.ts` file. This implementation is a rewrite of the original pptxjs library, using modern TypeScript.

## Usage

To use the document viewer:

1. Import the `Viewer` component.
2. Provide the file URL and file type as props.

Example:

```typescript
import Viewer from './Viewer';

function App() {
  return (
    <Viewer fileUrl="path/to/document.pptx" fileType="pptx" />
  );
}
```

## Integration with divs2slides

For PPTX files, the generated HTML is further processed using the `divs2slides.js` library located in `src/assets/pptxjs/js/divs2slides.js`. This library converts the HTML content into a slide-based presentation format, enabling features like navigation between slides and fullscreen mode.

The `usePPTX` hook in `src/helpers/pptx/usePPTX.ts` handles the integration of divs2slides with the converted PPTX content.

For more detailed information on the implementation, please refer to the individual component files in the `src/components/` directory and the `src/helpers/pptx/PPTX.ts` file.
