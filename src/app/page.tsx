"use client";
import React, { useRef, useEffect } from "react";
import Handsontable from "handsontable/base";
import { registerAllModules } from "handsontable/registry";
import "handsontable/styles/handsontable.min.css";
import "handsontable/styles/ht-theme-main.min.css";
import { HotTable } from "@handsontable/react-wrapper";

registerAllModules();

const MySpreadsheet: React.FC = () => {
  const hotRef = useRef<HotTable>(null);

  useEffect(() => {
    if (typeof window !== "undefined") {
      const savedData = localStorage.getItem("spreadsheetData");
      if (hotRef.current) {
        hotRef.current.hotInstance.loadData(savedData ? JSON.parse(savedData) : Array(100).fill(Array(50).fill("")));
      }
    }
  }, []);

  const handleBeforeAutofill = (startRange: any, entireRange: any) => {
    if (!hotRef.current || !startRange || !entireRange) return;
    const hotInstance = hotRef.current.hotInstance;
    if (!hotInstance) return;

    const { from: startCell, to: endCell } = startRange;
    const { to: finalCell } = entireRange;

    if (!startCell || !endCell || !finalCell) return;

    const startRow = startCell.row;
    const endRow = endCell.row;
    const finalRow = finalCell.row;
    const col = startCell.col;

    const values: number[] = [];

    // Extract the values from the selection
    for (let row = startRow; row <= endRow; row++) {
      const cellValue = Number(hotInstance.getDataAtCell(row, col));
      if (!isNaN(cellValue)) {
        values.push(cellValue);
      } else {
        return; // Stop autofill if any non-numeric value is found
      }
    }

    // Ensure at least two values to detect a pattern
    if (values.length > 1) {
      const step = values[1] - values[0]; // Detect step size (common difference)
      let lastValue = values[values.length - 1];

      for (let row = endRow + 1; row <= finalRow; row++) {
        lastValue += step;
        hotInstance.setDataAtCell(row, col, lastValue);
      }
    }
  };

  return (
    <div className="ht-theme-main-dark-auto">
      <HotTable
        ref={hotRef}
        contextMenu={true}
        dropdownMenu={true}
        manualColumnResize={true}  // Allows resizing columns
        manualRowResize={true}      // Allows resizing rows
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
        autofill={true}
        licenseKey="non-commercial-and-evaluation"
        beforeAutofill={handleBeforeAutofill} // Custom autofill logic
        afterChange={() => {
          if (hotRef.current) {
            localStorage.setItem("spreadsheetData", JSON.stringify(hotRef.current.hotInstance.getData()));
          }
        }}
      />
    </div>
  );
};

export default MySpreadsheet;
