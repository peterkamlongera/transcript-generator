import React, { useState } from "react";
import ExcelJS from "exceljs";

const DebugWorkbook = () => {
  const [selectedFile, setSelectedFile] = useState(null);

  const handleFileChange = (e) => {
    setSelectedFile(e.target.files[0] || null);
  };

  const handleCheckWorkbook = async () => {
    if (!selectedFile) {
      alert("Please select an Excel file first.");
      return;
    }

    try {
      const workbook = new ExcelJS.Workbook();

      // Must be .xlsx, not .xls or .numbers, etc.
      if (
        !selectedFile.name.toLowerCase().endsWith(".xlsx") &&
        !selectedFile.name.toLowerCase().endsWith(".xls")
      ) {
        alert("Please make sure you are uploading a valid .xlsx or .xls file");
        return;
      }

      const arrayBuffer = await selectedFile.arrayBuffer();
      await workbook.xlsx.load(arrayBuffer);

      console.log("Number of worksheets found:", workbook.worksheets.length);
      workbook.worksheets.forEach((ws, idx) => {
        console.log(`Worksheet #${idx} => name: "${ws.name}"`);
      });

      const firstSheet = workbook.worksheets[0];
      if (!firstSheet) {
        console.error("No worksheet found at index [0].");
        alert("No worksheet found at index [0].");
        return;
      }

      // If we get here, we successfully recognized the first sheet
      console.log("First worksheet name:", firstSheet.name);
      alert(`Success! First worksheet name is "${firstSheet.name}".`);
    } catch (error) {
      console.error("Error reading workbook:", error);
      alert("Error reading workbook: " + error.message);
    }
  };

  return (
    <div style={{ margin: "20px" }}>
      <h3>Debug Workbook</h3>
      <input type="file" accept=".xls,.xlsx" onChange={handleFileChange} />
      <button onClick={handleCheckWorkbook} style={{ marginLeft: "10px" }}>
        Check Workbook
      </button>
    </div>
  );
};

export default DebugWorkbook;
