import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import './app.css';

function UploadExcel() {
  const [excelData, setExcelData] = useState(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [duplicates, setDuplicates] = useState([]);
  const [filteredDuplicates, setFilteredDuplicates] = useState([]);

  useEffect(() => {
    if (excelData && excelData.length >= 2) {
      const products = {};

      excelData.slice(1).forEach((row, rowIndex) => {
        const productName = row[9] || '';
        const price = parseFloat(row[13]) || 0;

        if (!products[productName]) {
          products[productName] = {
            indices: [rowIndex],
            maxPrice: price,
          };
        } else {
          products[productName].indices.push(rowIndex);
          if (price > products[productName].maxPrice) {
            products[productName].maxPrice = price;
          }
        }
      });

      const duplicatesFound = Object.values(products)
        .filter((product) => product.indices.length > 1)
        .flatMap((product) => {
          return product.indices.map((index) => excelData[index]);
        });

      setDuplicates(duplicatesFound);
    }
  }, [excelData]);

  useEffect(() => {
    const filtered = duplicates.filter((row) =>
      row.some((cell) =>
        String(cell).toLowerCase().includes(searchTerm.toLowerCase())
      )
    );
    setFilteredDuplicates(filtered);
  }, [duplicates, searchTerm]);

  function handleFileUpload(event) {
    const file = event.target.files[0];

    if (!file) {
      return;
    }

    const reader = new FileReader();

    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const dataParsed = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      setExcelData(dataParsed);
    };

    reader.readAsArrayBuffer(file);
  }

  const downloadFilteredExcel = () => {
    const filteredData = excelData.filter(
      (row, index) => index === 0 || filteredDuplicates.includes(row)
    );
    const wb = XLSX.utils.book_new();
    const wsData = [excelData[0], ...filteredData];
    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // Formatar colunas de F a N como moeda (exceto na primeira linha)
    for (let i = 5; i <= 13; i++) {
      const columnLetter = XLSX.utils.encode_col(i);
      for (let j = 2; j <= filteredData.length; j++) {
        const cell = ws[`${columnLetter}${j}`];
        if (cell && cell.t === 'n') {
          const cellValue = cell.v;
          if (cellValue && typeof cellValue === 'number') {
            ws[`${columnLetter}${j}`].z = '#,##0.00';
          }
        }
      }
    }

    XLSX.utils.book_append_sheet(wb, ws, 'Filtered Data');

    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });

    const s2ab = (s) => {
      const buf = new ArrayBuffer(s.length);
      const view = new Uint8Array(buf);
      for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xff;
      return buf;
    };

    const fileData = s2ab(wbout);
    const blob = new Blob([fileData], { type: 'application/octet-stream' });
    const fileName = 'Filtered_Data.xlsx';

    if (window.navigator && window.navigator.msSaveOrOpenBlob) {
      window.navigator.msSaveOrOpenBlob(blob, fileName);
    } else {
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', fileName);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  };

  const showDuplicates = duplicates.length > 0 && (
    <div>
      <h3>Registros com Produto e Fornecedor Iguais</h3>
      <h4>Total de Registros: {duplicates.length}</h4>
      <table style={{ borderCollapse: 'collapse', width: '100%' }}>
        {/* ... (seu código de cabeçalho da tabela) */}
        <tbody>
          {filteredDuplicates.map((row, rowIndex) => {
            const product = row[9];
            let maxPrice = 0;
            let maxIndex = -1;
  
            filteredDuplicates.forEach((r, i) => {
              const rProduct = r[9];
              const price = parseFloat(r[13]) || 0;
  
              if (rProduct === product && price > maxPrice) {
                maxPrice = price;
                maxIndex = i;
              }
            });
  
            return (
              <tr
                key={rowIndex}
                style={{
                  backgroundColor: rowIndex === maxIndex ? 'lightgreen' : '',
                }}
              >
                <td
                  style={{
                    border: '1px solid black',
                    padding: '8px',
                  }}
                >
                  {rowIndex + 1}
                </td>
                {row.map((value, cellIndex) => {
                  const isMaxPriceCell = cellIndex === 13 && rowIndex === maxIndex;
                  return (
                    <td
                      key={cellIndex}
                      style={{
                        border: '1px solid black',
                        padding: '8px',
                      }}
                    >
                      {isMaxPriceCell ? (
                        <>
                          {typeof value === 'number' ? (
                            new Intl.NumberFormat('pt-BR', {
                              style: 'currency',
                              currency: 'BRL',
                            }).format(value)
                          ) : (
                            value
                          )}
                          <br />
                          <span>(valor máximo)</span>
                        </>
                      ) : cellIndex >= row.length - 2 ? (
                        typeof value === 'number' ? (
                          new Intl.NumberFormat('pt-BR', {
                            style: 'currency',
                            currency: 'BRL',
                          }).format(value)
                        ) : (
                          value
                        )
                      ) : (
                        value
                      )}
                    </td>
                  );
                })}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
  

  return (
    <div style={{ textAlign: 'center' }}>
      <h2>Verificação De Diferença</h2>
      <input type="file" onChange={handleFileUpload} style={{ marginBottom: '20px' }} />
      <div style={{ display: 'flex', justifyContent: 'center', marginBottom: '20px' }}>
        <span style={{ marginRight: '10px', fontSize: '30px' }}>Buscar</span>
        <input
          type="text"
          placeholder="Digite para buscar..."
          value={searchTerm}
          onChange={(e) => setSearchTerm(e.target.value)}
          style={{ padding: '10px', fontSize: '16px', flex: '1' }}
        />
      </div>
      <div style={{ display: 'flex', justifyContent: 'center', marginBottom: '20px' }}>
        <button onClick={downloadFilteredExcel} style={{ marginRight: '20px' }}>
          Baixar Planilha
        </button>
      </div>
      {showDuplicates}
    </div>
  );
}

export default UploadExcel;
