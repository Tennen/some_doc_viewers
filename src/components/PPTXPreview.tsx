import React from 'react';
import { usePPTX } from '../helpers/pptx/usePPTX';

interface PPTXPreviewProps {
  fileUrl: string;
}

const PPTXPreview: React.FC<PPTXPreviewProps> = ({ fileUrl }) => {
  const pptxComponent = usePPTX({ url: fileUrl });
  return pptxComponent;
};

export default PPTXPreview;