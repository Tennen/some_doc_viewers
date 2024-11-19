import React, { useEffect, useState } from 'react';
import * as XLSX from 'xlsx';
import '../styles/xlsx.less';

interface XLSXPreviewProps {
  url: string;
}

const XLSXPreview: React.FC<XLSXPreviewProps> = ({ url }) => {
  const [data, setData] = useState<any[][]>([]);

  useEffect(() => {
    const fetchData = async () => {
      const response = await fetch(url);
      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      setData(jsonData as any[][]);
    };

    fetchData();
  }, [url]);

  return (
    <div className="xlsx-preview">
      <table>
        <tbody>
          {data.map((row, rowIndex) => (
            <tr key={rowIndex}>
              {row.map((cell, cellIndex) => (
                <td key={cellIndex}>{cell}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export default XLSXPreview;