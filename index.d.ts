import React from 'react';

declare module 'some_doc_viewers' {
  export interface FilePreviewProps {
    fileUrl: string;
    fileType: 'pdf' | 'xlsx' | 'docx' | 'pptx';
  }

  const FilePreview: React.FC<FilePreviewProps>;

  export default FilePreview;
}