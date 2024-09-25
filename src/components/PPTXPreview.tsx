import React, { useEffect, useRef } from 'react';
import '../assets/pptxjs';
import '../styles/pptx.less';

interface PPTXPreviewProps {
  fileUrl: string;
}

const PPTXPreview: React.FC<PPTXPreviewProps> = ({ fileUrl }) => {
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (containerRef.current) {
      const options = {
        pptxFileUrl: fileUrl,
        slidesScale: '50%',
        slideMode: true,
        keyBoardShortCut: false,
      };

      // @ts-ignore
      $("#" + containerRef.current.id).pptxToHtml(options);
    }
  }, [fileUrl]);

  return <div ref={containerRef} id="pptx-preview" className="pptx-preview" />;
};

export default PPTXPreview;