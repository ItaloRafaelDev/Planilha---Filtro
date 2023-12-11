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

    reader.onload = function (e) {
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

    const updatedData = filteredData.map((row, rowIndex) => {
      const product = row[9];
      let maxPrice = 0;

      filteredDuplicates.forEach((r) => {
        const rProduct = r[9];
        const price = parseFloat(r[13]) || 0;

        if (rProduct === product && price > maxPrice) {
          maxPrice = price;
        }
      });

      const updatedRow = row.map((value, cellIndex) => {
        if (cellIndex === 13 && value === maxPrice) {
          return typeof value === 'number'
            ? `${new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(value)} (valor máximo)`
            : value;
        } else {
          return value;
        }
      });

      return updatedRow;
    });

    const ws = XLSX.utils.aoa_to_sheet(updatedData);

    // Formatar colunas M (13) e N (14) como moeda, ignorando o cabeçalho
    const range = XLSX.utils.decode_range(ws['!ref']); // Obter o intervalo de células
    for (let i = range.s.r + 1; i <= range.e.r; i++) {
      for (let j = range.s.c; j <= range.e.c; j++) {
        const cellAddress = XLSX.utils.encode_cell({ r: i, c: j });
        if (j === 12 || j === 13) {
          if (typeof ws[cellAddress].v === 'number') {
            ws[cellAddress].z = '#,##0.00';
          }
        }
      }
    }

    const wb = XLSX.utils.book_new();
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
  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
    <h2 style={{ fontSize: '40px', marginRight: '50px', marginTop: '110px' }}>Verificação De Dados</h2>
    <img src="logo.ico" alt="Descrição da imagem" style={{ width: '290px', height: '200px', marginTop: '50px' }} />
  </div>
      <input type="file" onChange={handleFileUpload} style={{ marginBottom: '30px' }} />
      <div style={{ display: 'flex', justifyContent: 'center', marginBottom: '20px' }}>
        <span style={{ marginRight: '10px', fontSize: '30px' }}>Buscar :</span>
        <input
          type="text"
          placeholder="Digite para buscar..."
          value={searchTerm}
          onChange={(e) => setSearchTerm(e.target.value)}
          style={{ padding: '10px', fontSize: '22px', flex: '1' }}
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
