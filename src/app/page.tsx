"use client";
import React, { useRef, useEffect, useState } from "react";
import Handsontable from "handsontable/base";
import { registerAllModules } from "handsontable/registry";
import "handsontable/styles/handsontable.min.css";
import "handsontable/styles/ht-theme-main.min.css";
import { HotTable } from "@handsontable/react-wrapper";
import { HyperFormula } from "hyperformula";

registerAllModules();

// Define the structure of a cell
interface CellData {
  value: string | number;
  formula: string | null;
}

const MySpreadsheet: React.FC = () => {
  const hotRef = useRef<HotTable>(null);
  const [formula, setFormula] = useState<string>(""); // State for formula bar input
  const [selectedCell, setSelectedCell] = useState<{ row: number; col: number } | null>(null); // Track selected cell

  // Initialize HyperFormula
  const hyperformulaInstance = HyperFormula.buildEmpty({
    licenseKey: "gpl-v3", // Add your license key if required
  });

  // Add a sheet to HyperFormula
  const sheetName = "Sheet1";
  const sheetId = hyperformulaInstance.addSheet(sheetName);

  // Initialize data with the custom cell structure
  const [initialData, setInitialData] = useState<CellData[][]>(
    Array(100)
      .fill(null)
      .map(() =>
        Array(50)
          .fill(null)
          .map(() => ({ value: "", formula: null }))
      )
  );

  useEffect(() => {
    if (typeof window !== "undefined") {
      try {
        const savedData = localStorage.getItem("spreadsheetData");
        if (hotRef.current) {
          const hotInstance = hotRef.current.hotInstance;
          if (hotInstance) {
            const data = savedData ? JSON.parse(savedData) : initialData;
            setInitialData(data); // Load data into state
            hotInstance.loadData(data.map((row: CellData[]) => row.map((cell) => cell.value))); // Load only values into Handsontable
          }
        }
      } catch (error) {
        console.error("Error loading data from localStorage:", error);
      }
    }
  }, []);

  // Handle cell selection
  const handleCellSelect = () => {
    if (hotRef.current) {
      const hotInstance = hotRef.current.hotInstance;
      if (hotInstance) {
        const selected = hotInstance.getSelected()?.[0];
        if (selected) {
          const [startRow, startCol] = selected;
          setSelectedCell({ row: startRow, col: startCol });

          // Get the cell data
          const cellData = initialData[startRow][startCol];

          // Display the formula (if any) or the value in the formula bar
          setFormula(cellData.formula || cellData.value.toString());
        }
      }
    }
  };

  // Handle formula submission
  const handleFormulaSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (selectedCell && hotRef.current) {
      const hotInstance = hotRef.current.hotInstance;
      if (hotInstance) {
        const { row, col } = selectedCell;

        // Update the cell data
        if (formula.startsWith("=")) {
          // If the input starts with "=", treat it as a formula
          const cellAddress = {
            sheet: sheetId,
            row,
            col,
          };

          try {
            // Set the formula in HyperFormula
            hyperformulaInstance.setCellFormula(cellAddress, formula.slice(1));

            // Get the computed value
            const computedValue = hyperformulaInstance.getCellValue(cellAddress);

            // Update the initialData structure
            const newData = [...initialData];
            newData[row][col] = {
              value: computedValue,
              formula: formula,
            };
            setInitialData(newData);

            // Update the grid with the computed value
            hotInstance.setDataAtCell(row, col, computedValue);
          } catch (error) {
            console.error("Error setting cell formula:", error);
          }
        } else {
          // If it's plain text or a number, set it directly
          const newData = [...initialData];
          newData[row][col] = {
            value: formula,
            formula: null, // Clear the formula
          };
          setInitialData(newData);

          // Update the grid with the plain value
          hotInstance.setDataAtCell(row, col, formula);
        }

        setFormula(""); // Clear the formula bar
      }
    }
  };

  // Save data to localStorage on change
  const handleAfterChange = () => {
    if (hotRef.current) {
      try {
        localStorage.setItem("spreadsheetData", JSON.stringify(initialData));
      } catch (error) {
        console.error("Error saving data to localStorage:", error);
      }
    }
  };

  return (
    <div className="ht-theme-main-dark-auto">
      {/* Formula Bar */}
      <form onSubmit={handleFormulaSubmit} style={{ marginBottom: "10px" }}>
        <input
          type="text"
          value={formula}
          onChange={(e) => setFormula(e.target.value)}
          placeholder="Enter formula (e.g., =A1+B1) or plain text"
          style={{ width: "100%", padding: "5px", fontSize: "16px" }}
        />
        <button type="submit" style={{ display: "none" }}>
          Apply
        </button>
      </form>

      {/* Handsontable Grid */}
      <HotTable
        ref={hotRef}
        data={initialData.map((row) => row.map((cell) => cell.value))} // Load only values into Handsontable
        contextMenu={true}
        dropdownMenu={true}
        manualColumnResize={true}
        manualRowResize={true}
        filters={true}
        manualColumnMove={true}
        manualRowMove={true}
        rowHeaders={true}
        colHeaders={true}
        height="auto"
        width="auto"
        autoWrapRow={true}
        autoWrapCol={true}
        minCols={50}
        minRows={100}
        colWidths={100}
        autofill={false}
        licenseKey="non-commercial-and-evaluation"
        formulas={{
          engine: hyperformulaInstance,
        }}
        afterSelection={handleCellSelect} // Track cell selection
        afterChange={handleAfterChange} // Save data to localStorage
        formulaCalculationMode="onEdit"
      />
    </div>
  );
};

export default MySpreadsheet;
