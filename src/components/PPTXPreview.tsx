import React, { useRef } from 'react';
import { usePPTX } from '../helpers/pptx/usePPTX';

interface PPTXPreviewProps {
  fileUrl: string;
}

const PPTXPreview: React.FC<PPTXPreviewProps> = ({ fileUrl }) => {
  const containerRef = useRef<HTMLDivElement>(null);
  usePPTX({ url: fileUrl, containerRef: containerRef });
  return <div ref={containerRef}></div>;
};

export default PPTXPreview;