import React from 'react';
import { Document, Page } from 'react-pdf';
import '../styles/pdf.less';

interface PDFPreviewProps {
  fileUrl: string;
}

const PDFPreview: React.FC<PDFPreviewProps> = ({ fileUrl }) => {
  return (
    <div className="pdf-preview">
      <Document file={fileUrl}>
        <Page pageNumber={1} />
      </Document>
    </div>
  );
};

export default PDFPreview;