import React, { useRef } from 'react';
import { usePPTX } from '../helpers/pptx/usePPTX';

interface PPTXPreviewProps {
  url?: string;
  file?: ArrayBuffer;
  processFullTheme?: boolean | string;
  mediaProcess?: boolean;
  fontSizeFactor?: number;
  slideFactor?: number;
}

const PPTXPreview: React.FC<PPTXPreviewProps> = ({ url }) => {
  const containerRef = useRef<HTMLDivElement>(null);
  usePPTX({ url, containerRef: containerRef });
  return <div ref={containerRef}></div>;
};

export default PPTXPreview;