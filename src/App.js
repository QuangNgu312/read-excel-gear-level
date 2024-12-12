import React, { useState } from 'react';
import * as XLSX from 'xlsx';

function App() {
  return <ExcelReader />;
}

const ExcelReader = () => {
  const [fileData, setFileData] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  const [sheetData, setSheetData] = useState(null);
  const [columnNames, setColumnNames] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [currentGroup, setCurrentGroup] = useState(0);
  const [customName, setCustomName] = useState('');

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        const unmergedData = unmergeSheet(worksheet);
        setFileData(file);
        setSheetData(unmergedData);
        setSheetNames(workbook.SheetNames);
      };
      reader.readAsArrayBuffer(file);
    }
  }
  const ReadFile = (file, sheetName) => {
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[sheetName];

        const unmergedData = unmergeSheet(worksheet);
        setFileData(file);
        setSheetData(unmergedData);
        setSheetNames(workbook.SheetNames);
      };
      reader.readAsArrayBuffer(file);
    }
  }
  const unmergeSheet = (worksheet) => {
    // Parse sheet into JSON for easier manipulation
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Check for merged cells
    const merges = worksheet['!merges'] || [];
    merges.forEach((merge) => {
      const startRow = merge.s.r; // Start row
      const endRow = merge.e.r;   // End row
      const startCol = merge.s.c; // Start column
      const endCol = merge.e.c;   // End column

      // Value in the top-left cell of the merge
      const value = jsonData[startRow][startCol];

      // Apply value to all cells in the merged range
      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
          if (!jsonData[row]) jsonData[row] = [];
          jsonData[row][col] = value;
        }
      }
    });

    return jsonData;
  };

  const loadSheetData = (workbook, sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, range: 2 });
    setSheetData(removeLastTwoColumns(jsonData));
  };

  const handleSheetSelect = (event) => {
    const selectedSheetName = event.target.value;
    ReadFile(fileData, selectedSheetName);
    setSelectedSheet(selectedSheetName);
  };

  const removeLastTwoColumns = (data) => data.map(row => row.slice(0, -2));

  const updateHeaderNames = (headers) => {
    const updatedHeaders = [...headers];
    return updatedHeaders;
  };

  const handleNameChange = (e) => setCustomName(e.target.value);

  const handleSaveName = () => {
    const updatedNames = [...columnNames];
    updatedNames[currentGroup] = customName;
    setColumnNames(updatedNames);
    setCustomName('');
    setCurrentGroup(currentGroup + 1);

    if (currentGroup + 1 >= Math.ceil(sheetData[0].length / 6)) {
      setShowModal(false);
    }
  };

  const openModal = () => {
    setShowModal(true);
    setCurrentGroup(0);
  };

  const generateJson = () => {
    // Filter the rows to exclude invalid waves
    let newSheetData = sheetData.filter((item, index) => {
      const wave = item[0];
      // Skip rows where the wave value is invalid (null, "-", or "")
      return wave && wave !== "-" && wave !== "" && index > 0;
    });

    let finalData = {
      "level":[]
    }
    newSheetData.forEach(row => {
      let currentLevel = {
        level:row[0],
        material:row[1],
        gold:row[2],
        xRate:row[3]
      }
      finalData.level.push(currentLevel);
    })

    
    const blob = new Blob([JSON.stringify(finalData, null, 2)], { type: 'application/json' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'output.json';
    link.click();
  };





  return (
    <div style={styles.container}>
      <input type="file" onChange={handleFileUpload} accept=".xlsx, .xls" />

      {sheetNames.length > 0 && (
        <select value={selectedSheet} onChange={handleSheetSelect} style={styles.select}>
          {sheetNames.map((name, index) => (
            <option key={index} value={name}>{name}</option>
          ))}
        </select>
      )}

      {sheetData && (
        <>
          <button onClick={generateJson} style={styles.button}>Generate JSON</button>
          <table style={styles.table}>
            <thead>
              <tr style={styles.headerRow}>
                {updateHeaderNames(sheetData[0]).map((header, index) => (
                  index < sheetData[0].length && (
                    <th key={index} style={styles.headerCell}>{header}</th>
                  )
                ))}
              </tr>
            </thead>
            <tbody>
              {sheetData.map((row, index) => {
                if (index <= 0) {
                  return (<tr></tr>);
                }
                return (
                  <tr key={index} style={index % 2 === 0 ? styles.evenRow : styles.oddRow}>
                    {row.map((cell, i) => (
                      <td key={i} style={styles.cell}>{cell}</td>
                    ))}
                  </tr>
                )
              })}
            </tbody>
          </table>
        </>
      )}

      {showModal && (
        <div style={styles.modal}>
          <div style={styles.modalContent}>
            <h3>Enter a name for columns {currentGroup * 5 + 5} - {Math.min((currentGroup + 1) * 5, sheetData[0].length)}</h3>
            <input
              type="text"
              value={customName}
              onChange={handleNameChange}
              style={styles.input}
            />
            <button onClick={handleSaveName} style={styles.button}>Save Name</button>
            <button onClick={() => setShowModal(false)} style={styles.button}>Close</button>
          </div>
        </div>
      )}
    </div>
  );
};

// Styling object
const styles = {
  container: { padding: '20px', fontFamily: 'Arial, sans-serif' },
  select: { padding: '10px', margin: '10px 0', fontSize: '16px' },
  table: { width: '100%', borderCollapse: 'collapse', marginTop: '20px' },
  headerCell: { padding: '10px', border: '1px solid #ddd', textAlign: 'left', fontWeight: 'bold' },
  evenRow: { backgroundColor: '#f2f2f2' },
  oddRow: { backgroundColor: '#ffffff' },
  cell: { padding: '10px', border: '1px solid #ddd', textAlign: 'left' },
  button: { padding: '10px 20px', backgroundColor: '#4CAF50', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer', marginTop: '20px', marginRight: '10px' },
  modal: { position: 'fixed', top: '0', left: '0', width: '100%', height: '100%', backgroundColor: 'rgba(0, 0, 0, 0.5)', display: 'flex', justifyContent: 'center', alignItems: 'center' },
  modalContent: { backgroundColor: 'white', padding: '20px', borderRadius: '4px', width: '300px', textAlign: 'center' },
  input: { width: '100%', padding: '10px', marginBottom: '10px', border: '1px solid #ccc', borderRadius: '4px' },
  headerRow: { backgroundColor: '#4CAF50', color: 'white' },
};

export default App;
