import React, { useEffect, useState } from 'react';
import mammoth from 'mammoth';

interface DOCXPreviewProps {
  fileUrl: string;
}

const DOCXPreview: React.FC<DOCXPreviewProps> = ({ fileUrl }) => {
  const [content, setContent] = useState('');

  useEffect(() => {
    const fetchData = async () => {
      const response = await fetch(fileUrl);
      const arrayBuffer = await response.arrayBuffer();
      const result = await mammoth.convertToHtml({ arrayBuffer });
      setContent(result.value);
    };

    fetchData();
  }, [fileUrl]);

  return (
    <div className="docx-preview" dangerouslySetInnerHTML={{ __html: content }} />
  );
};

export default DOCXPreview;