import React, { useEffect, useState } from 'react';
import mammoth from 'mammoth';

interface DOCXPreviewProps {
  url: string;
}

const DOCXPreview: React.FC<DOCXPreviewProps> = ({ url }) => {
  const [content, setContent] = useState('');

  useEffect(() => {
    const fetchData = async () => {
      const response = await fetch(url);
      const arrayBuffer = await response.arrayBuffer();
      const result = await mammoth.convertToHtml({ arrayBuffer });
      setContent(result.value);
    };

    fetchData();
  }, [url]);

  return (
    <div className="docx-preview" dangerouslySetInnerHTML={{ __html: content }} />
  );
};

export default DOCXPreview;