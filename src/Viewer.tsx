import React from 'react';
import PDFPreview from './components/PDFPreview';
import XLSXPreview from './components/XLSXPreview';
import DOCXPreview from './components/DOCXPreview';
import PPTXPreview from './components/PPTXPreview';

interface FilePreviewProps {
  fileUrl: string;
  fileType: 'pdf' | 'xlsx' | 'docx' | 'pptx';
}

const FilePreview: React.FC<FilePreviewProps> = ({ fileUrl, fileType }) => {
  const renderPreview = () => {
    switch (fileType) {
      case 'pdf':
        return <PDFPreview fileUrl={fileUrl} />;
      case 'xlsx':
        return <XLSXPreview fileUrl={fileUrl} />;
      case 'docx':
        return <DOCXPreview fileUrl={fileUrl} />;
      case 'pptx':
        return <PPTXPreview fileUrl={fileUrl} />;
      default:
        return <div>Unsupported file type</div>;
    }
  };

  return (
    <div className="file-preview-container">
      {renderPreview()}
    </div>
  );
};

export default FilePreview;
