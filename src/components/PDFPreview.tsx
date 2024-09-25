import React, { useState } from 'react';
import { Document, Page, pdfjs } from 'react-pdf';
import 'react-pdf/dist/esm/Page/TextLayer.css'
import '../styles/pdf.less';

interface PDFPreviewProps {
  fileUrl: string;
}

pdfjs.GlobalWorkerOptions.workerSrc = `//unpkg.com/pdfjs-dist@${pdfjs.version}/build/pdf.worker.min.mjs`;

const PDFPreview: React.FC<PDFPreviewProps> = ({ fileUrl }) => {
    const [pageNumber, setPageNumber] = useState<number>(1);

    function onDocumentLoadSuccess({ numPages }: { numPages: number }): void {
        setPageNumber(numPages);
    }
  return (
    <div className="pdf-preview">
      <Document file={fileUrl} onLoadSuccess={onDocumentLoadSuccess}>
        {Array.from({ length: pageNumber }, (_, index) => (
          <Page key={index} pageNumber={index + 1} renderTextLayer={false} scale={1.5} />
        ))}
      </Document>
    </div>
  );
};

export default PDFPreview;